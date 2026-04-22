"""
Extractor do Relatório de Pontos de Operação — RasCol.

Fluxo:
  1. Login → seleciona empresa/filial → navega ao relatório (uma única vez)
  2. Define Data/Hora da primeira janela
  3. Seleciona DOMICILIAR → aguarda lista de veículos (uma única vez)
  4. Para cada janela de 7 dias:
       a. Atualiza Data/Hora (somente se não for a primeira janela)
       b. Para cada veículo (em ordem):
            - Seleciona veículo
            - Pesquisa
            - Se houver resultado → Exportar → aguarda download
  5. Executa ShapesProcessor se solicitado
"""

import os
import time
from datetime import date, timedelta
from pathlib import Path
from typing import List, Optional, Tuple

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    StaleElementReferenceException,
)

from rascol_automation.core.browser import open_browser
from rascol_automation.core.waits import wait_load_rascol, _js_click
from rascol_automation.core.auth import (
    login_rascol,
    select_company_locar,
    select_filial,
    navigate_to_report,
)
from rascol_automation.config.settings import RASCOL_DOWNLOAD_DIR
from rascol_automation.config.rascol_config import RasColConfig

# IDs dos elementos do formulário de pesquisa
_ID_ROTULOS   = "cad_ctl00_ContentPlaceHolder_Filtro_rotulos_ddlRotulos"
_ID_VEICULOS  = "cad_ctl00_ContentPlaceHolder_Filtro_rotulos_ddlVeiculos"
_ID_DT_INICIO = "cad_ctl00_ContentPlaceHolder_Filtro_txtDataInicio"
_ID_HR_INICIO = "cad_ctl00_ContentPlaceHolder_Filtro_txtHoraInicio"
_ID_DT_FIM    = "cad_ctl00_ContentPlaceHolder_Filtro_txtDataFim"
_ID_HR_FIM    = "cad_ctl00_ContentPlaceHolder_Filtro_txtHoraFim"
_ID_PESQUISAR = "cad_ctl00_ContentPlaceHolder_Filtro_btnPesquisar_btnInterno"
_ID_EXPORTAR  = "cad_ctl00_ContentPlaceHolder_upnCentral_btnExportar_btnInterno"

DOMICILIAR_VALUE = "2368"
MAX_WINDOW_DAYS  = 7


# ---------------------------------------------------------------------------
# Helpers de data
# ---------------------------------------------------------------------------

def _date_windows(dates: List[date]) -> List[Tuple[date, date]]:
    """
    Divide [min(dates), max(dates)] em janelas de no máximo MAX_WINDOW_DAYS dias.
    Retorna lista de (data_inicio, data_fim).
    """
    if not dates:
        return []
    start = min(dates)
    end   = max(dates)
    windows = []
    current = start
    while current <= end:
        window_end = min(current + timedelta(days=MAX_WINDOW_DAYS - 1), end)
        windows.append((current, window_end))
        current = window_end + timedelta(days=1)
    return windows


def _fmt_date(d: date) -> str:
    return d.strftime("%d/%m/%Y")


# ---------------------------------------------------------------------------
# Helpers de interação com o formulário
# ---------------------------------------------------------------------------

def _clear_and_set(driver, element_id: str, value: str):
    el = driver.find_element(By.ID, element_id)
    driver.execute_script("arguments[0].value = '';", el)
    el.click()
    el.send_keys(value)


def _set_date_range(driver, dt_inicio: date, dt_fim: date):
    """
    Preenche data/hora do formulário.
    Janela: dt_inicio 05:00 → (dt_fim + 1 dia) 05:00
    """
    _clear_and_set(driver, _ID_DT_INICIO, _fmt_date(dt_inicio))
    _clear_and_set(driver, _ID_HR_INICIO, "05:00")
    _clear_and_set(driver, _ID_DT_FIM,    _fmt_date(dt_fim + timedelta(days=1)))
    _clear_and_set(driver, _ID_HR_FIM,    "05:00")


def _has_no_results(driver) -> bool:
    try:
        driver.find_element(By.CSS_SELECTOR, "table.emptyData")
        return True
    except NoSuchElementException:
        return False


def _exportar_button_available(driver) -> bool:
    try:
        btn = driver.find_element(By.ID, _ID_EXPORTAR)
        return btn.is_displayed()
    except NoSuchElementException:
        return False


# ---------------------------------------------------------------------------
# Detecção de download de .xls
# ---------------------------------------------------------------------------

def _wait_for_xls_download(download_dir: Path, start_ts: float, timeout=120) -> List[str]:
    """
    Aguarda um arquivo .xls aparecer em download_dir após start_ts.
    Retorna lista de caminhos encontrados.
    """
    end_time = time.time() + timeout

    while time.time() < end_time:
        try:
            entries = os.listdir(download_dir)
        except FileNotFoundError:
            time.sleep(0.5)
            continue

        if any(f.lower().endswith((".crdownload", ".part")) for f in entries):
            time.sleep(0.5)
            continue

        candidates = []
        for fname in entries:
            if not fname.lower().endswith((".xls", ".xlsx")):
                continue
            full = str(download_dir / fname)
            try:
                mtime = os.path.getmtime(full)
            except OSError:
                continue
            if mtime >= start_ts - 1.0:
                candidates.append(full)

        if not candidates:
            time.sleep(0.5)
            continue

        stable = True
        for p in candidates:
            try:
                s1 = os.path.getsize(p)
                time.sleep(0.5)
                s2 = os.path.getsize(p)
                if s1 != s2:
                    stable = False
                    break
            except OSError:
                stable = False
                break

        if stable:
            return candidates

        time.sleep(0.5)

    return []


# ---------------------------------------------------------------------------
# Extractor principal
# ---------------------------------------------------------------------------

class PontosExtractor:
    """
    Extrai o Relatório de Pontos de Operação do RasCol para todos os
    veículos DOMICILIAR, dividindo datas em janelas de até 7 dias.

    Loop externo = janelas de datas; loop interno = veículos.
    Data/Hora é definida uma vez por janela. DOMICILIAR é selecionado uma vez.
    """

    def __init__(
        self,
        dates: List[date],
        config: Optional[RasColConfig] = None,
        run_shapes: bool = False,
        progress_callback=None,
    ):
        from rascol_automation.config.rascol_config import load_rascol_config
        self.dates      = sorted(dates)
        self.config     = config or load_rascol_config()
        self.run_shapes = run_shapes
        self._log       = progress_callback or print
        self.driver     = None

        self.total_vehicles  = 0
        self.total_downloads = 0
        self.total_skipped   = 0
        self.errors: List[str] = []

    # ------------------------------------------------------------------

    def run(self):
        """Ponto de entrada principal."""
        try:
            self._setup()

            self._log("Fazendo login e navegando ao relatório...")
            self._login_and_navigate()

            windows = _date_windows(self.dates)
            if not windows:
                self._log("Nenhuma data selecionada.")
                return

            self._log(f"Janelas de datas: {len(windows)} (máx. {MAX_WINDOW_DAYS} dias cada)")

            # Define Data/Hora da primeira janela ANTES de selecionar DOMICILIAR
            _set_date_range(self.driver, *windows[0])

            self._log("Selecionando rótulo DOMICILIAR...")
            self._select_domiciliar()

            vehicles = self._get_vehicle_options()
            self.total_vehicles = len(vehicles)
            self._log(f"Veículos encontrados: {len(vehicles)}")

            for win_idx, (dt_inicio, dt_fim) in enumerate(windows):
                win_label = _fmt_date(dt_inicio)
                if dt_inicio != dt_fim:
                    win_label += f" → {_fmt_date(dt_fim)}"
                self._log(f"\n— Janela {win_idx + 1}/{len(windows)}: {win_label} —")

                # Para a primeira janela as datas já foram definidas acima
                if win_idx > 0:
                    _set_date_range(self.driver, dt_inicio, dt_fim)

                for v_idx, (value, label) in enumerate(vehicles, 1):
                    self._log(f"  [{v_idx}/{len(vehicles)}] {label}")
                    self._process_vehicle(value, label, dt_inicio, dt_fim)

            self._log(
                f"\nExtração concluída. "
                f"Downloads: {self.total_downloads} | "
                f"Sem registro: {self.total_skipped} | "
                f"Erros: {len(self.errors)}"
            )

            if self.run_shapes:
                self._run_shapes_processor()

        finally:
            self._teardown()

    # ------------------------------------------------------------------
    # Setup / Teardown
    # ------------------------------------------------------------------

    def _setup(self):
        self.driver = open_browser()

    def _teardown(self):
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Login e navegação — executa uma única vez
    # ------------------------------------------------------------------

    def _login_and_navigate(self):
        driver = self.driver
        login_rascol(driver, self.config.username, self.config.password)
        self._log("  login OK")

        select_company_locar(driver)
        self._log("  empresa Locar OK")

        select_filial(driver, self.config.filial)
        self._log(f"  filial '{self.config.filial}' OK")

        navigate_to_report(driver)
        self._log("  relatório de Pontos OK")

    # ------------------------------------------------------------------
    # Seleção de DOMICILIAR — executa uma única vez (ou na recuperação)
    # ------------------------------------------------------------------

    def _select_domiciliar(self):
        """Seleciona DOMICILIAR e aguarda a lista de veículos ser populada."""
        driver = self.driver
        sel = Select(driver.find_element(By.ID, _ID_ROTULOS))
        sel.select_by_value(DOMICILIAR_VALUE)
        wait_load_rascol(driver, timeout=30)
        # Aguarda o dropdown de veículos ter pelo menos uma opção real
        try:
            WebDriverWait(driver, 30).until(
                lambda d: len(Select(d.find_element(By.ID, _ID_VEICULOS)).options) > 1
            )
        except TimeoutException:
            pass

    # ------------------------------------------------------------------
    # Leitura da lista de veículos — executa uma única vez
    # ------------------------------------------------------------------

    def _get_vehicle_options(self) -> List[Tuple[str, str]]:
        """
        Retorna lista de (value, text) excluindo o placeholder.
        Usa JavaScript para evitar StaleElementReferenceException ao iterar
        as options enquanto o AJAX ainda pode estar atualizando o dropdown.
        """
        driver = self.driver
        try:
            data = driver.execute_script(
                """
                var sel = document.getElementById(arguments[0]);
                if (!sel) return [];
                var out = [];
                for (var i = 0; i < sel.options.length; i++) {
                    var v = sel.options[i].value;
                    if (v) out.push([v, sel.options[i].text.trim()]);
                }
                return out;
                """,
                _ID_VEICULOS,
            )
            return [(v, t) for v, t in data] if data else []
        except Exception:
            self._log("Aviso: não foi possível ler a lista de veículos.")
            return []

    # ------------------------------------------------------------------
    # Processamento de um veículo (datas já definidas no formulário)
    # ------------------------------------------------------------------

    def _process_vehicle(
        self,
        vehicle_value: str,
        vehicle_label: str,
        dt_inicio: date,
        dt_fim: date,
    ):
        """
        Seleciona o veículo, pesquisa e exporta se houver resultado.
        Em caso de StaleElementReferenceException faz recovery uma vez:
        re-navega ao relatório, redefine datas e re-seleciona DOMICILIAR.
        """
        for attempt in range(2):
            try:
                # Seleciona o veículo no dropdown (datas já estão preenchidas)
                sel = Select(self.driver.find_element(By.ID, _ID_VEICULOS))
                sel.select_by_value(vehicle_value)

                # Pesquisa — clica e aguarda o botão ficar visível novamente.
                # O onclick chama ativarExibicaoDoComponenteLoading que seta
                # visibility:hidden no botão durante todo o postback; quando
                # os resultados estão renderizados o botão volta a ser visível.
                btn = self.driver.find_element(By.ID, _ID_PESQUISAR)
                _js_click(self.driver, btn)
                try:
                    WebDriverWait(self.driver, 120).until(
                        EC.visibility_of_element_located((By.ID, _ID_PESQUISAR))
                    )
                except TimeoutException:
                    pass

                # Verifica resultado
                if _has_no_results(self.driver):
                    self._log("    Sem registros.")
                    self.total_skipped += 1
                    return

                if not _exportar_button_available(self.driver):
                    self._log("    Botão Exportar indisponível.")
                    self.total_skipped += 1
                    return

                # Exporta e aguarda download
                start_ts = time.time()
                btn_exp = self.driver.find_element(By.ID, _ID_EXPORTAR)
                _js_click(self.driver, btn_exp)
                wait_load_rascol(self.driver, timeout=30)

                files = _wait_for_xls_download(RASCOL_DOWNLOAD_DIR, start_ts, timeout=120)
                if files:
                    self.total_downloads += 1
                    self._log(f"    Download: {os.path.basename(files[0])}")
                else:
                    self._log("    Download não concluído no tempo limite.")
                    self.errors.append(f"{vehicle_label}: timeout download")
                return

            except StaleElementReferenceException:
                if attempt > 0:
                    self.errors.append(f"{vehicle_label}: stale persistente após recovery")
                    return
                self._log("    Elemento stale — renavegando ao relatório...")
                try:
                    navigate_to_report(self.driver)
                    _set_date_range(self.driver, dt_inicio, dt_fim)
                    self._select_domiciliar()
                except Exception as re_e:
                    self.errors.append(f"{vehicle_label}: recuperação falhou: {re_e}")
                    return
                # tenta novamente (attempt 1)

            except Exception as e:
                self._log(f"    Erro: {e}")
                self.errors.append(f"{vehicle_label}: {e}")
                return

    # ------------------------------------------------------------------
    # Pós-processamento
    # ------------------------------------------------------------------

    def _run_shapes_processor(self):
        from rascol_automation.processors.processor_shapes import ShapesProcessor
        self._log("\nIniciando geração de shapefiles...")
        ShapesProcessor(
            excel_dir=RASCOL_DOWNLOAD_DIR,
            max_date=max(self.dates) if self.dates else None,
            log=self._log,
        ).run()
