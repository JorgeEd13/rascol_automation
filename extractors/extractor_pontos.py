"""
Extractor do Relatório de Pontos de Operação — RasCol.

Fluxo por veículo / janela de 7 dias:
  1. Seleciona o veículo no dropdown ddlVeiculos
  2. Preenche DataInicio / DataFim
  3. Clica em Pesquisar e aguarda AJAX
  4. Se "Nenhum registro" → próximo veículo/janela
  5. Se resultados → clica Exportar e aguarda download do .xls
"""

import os
import time
from datetime import date, timedelta
from pathlib import Path
from typing import List, Optional, Tuple

from selenium import webdriver
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
    Divide o intervalo [min(dates), max(dates)] em janelas de no máximo
    MAX_WINDOW_DAYS dias consecutivos.

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
    """Limpa um campo de texto e preenche com o valor."""
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
    """Retorna True se a tabela de resultados mostrar 'Nenhum registro'."""
    try:
        driver.find_element(By.CSS_SELECTOR, "table.emptyData")
        return True
    except NoSuchElementException:
        return False


def _exportar_button_available(driver) -> bool:
    """Retorna True se o botão Exportar existir e estiver visível."""
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

        # Downloads em andamento
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

        # Verifica estabilidade (tamanho não cresce)
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
    """

    def __init__(
        self,
        dates: List[date],
        config: Optional[RasColConfig] = None,
        run_shapes: bool = False,
        progress_callback=None,
    ):
        """
        Args:
            dates:             Lista de datas selecionadas pelo usuário.
            config:            Configuração RasCol (credenciais + filial).
            run_shapes:        Se True, executa ShapesProcessor após os downloads.
            progress_callback: Função(msg) chamada a cada evento para logs na GUI.
        """
        from rascol_automation.config.rascol_config import load_rascol_config
        self.dates      = sorted(dates)
        self.config     = config or load_rascol_config()
        self.run_shapes = run_shapes
        self._log       = progress_callback or print
        self.driver     = None

        # Estatísticas
        self.total_vehicles  = 0
        self.total_downloads = 0
        self.total_skipped   = 0
        self.errors: List[str] = []

    # ------------------------------------------------------------------

    def run(self):
        """Ponto de entrada principal."""
        try:
            self._setup()
            self._login_and_navigate()
            vehicles = self._get_vehicle_options()
            self.total_vehicles = len(vehicles)
            self._log(f"Veículos encontrados: {len(vehicles)}")

            windows = _date_windows(self.dates)
            self._log(
                f"Janelas de datas: {len(windows)} "
                f"(até {MAX_WINDOW_DAYS} dias cada)"
            )

            for idx, (value, label) in enumerate(vehicles, 1):
                self._log(f"\n[{idx}/{len(vehicles)}] {label}")
                self._process_vehicle(value, label, windows)

            self._log("\nTodos os veículos processados.")
            self._log(
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
    # Login e navegação inicial
    # ------------------------------------------------------------------

    def _login_and_navigate(self):
        driver = self.driver
        self._log("Fazendo login...")
        login_rascol(driver, self.config.username, self.config.password)

        self._log("Selecionando empresa Locar...")
        select_company_locar(driver)

        self._log(f"Selecionando filial '{self.config.filial}'...")
        select_filial(driver, self.config.filial)

        self._log("Navegando para Relatório de Pontos de Operação...")
        navigate_to_report(driver)

        self._log("Selecionando rótulo DOMICILIAR...")
        self._select_domiciliar()

    def _select_domiciliar(self):
        """Seleciona DOMICILIAR no dropdown de Rótulos e aguarda a lista de veículos."""
        driver = self.driver
        sel = Select(driver.find_element(By.ID, _ID_ROTULOS))
        sel.select_by_value(DOMICILIAR_VALUE)
        wait_load_rascol(driver, timeout=30)
        # Aguarda o dropdown de veículos ter pelo menos uma opção válida
        try:
            WebDriverWait(driver, 30).until(
                lambda d: len(Select(d.find_element(By.ID, _ID_VEICULOS)).options) > 1
            )
        except TimeoutException:
            pass

    # ------------------------------------------------------------------
    # Leitura da lista de veículos
    # ------------------------------------------------------------------

    def _get_vehicle_options(self) -> List[Tuple[str, str]]:
        """
        Retorna lista de (value, text) de todos os veículos do dropdown,
        excluindo a opção placeholder "[Selecione]".
        """
        driver = self.driver
        try:
            sel = Select(driver.find_element(By.ID, _ID_VEICULOS))
            return [
                (opt.get_attribute("value"), opt.text.strip())
                for opt in sel.options
                if opt.get_attribute("value")
            ]
        except NoSuchElementException:
            self._log("Aviso: dropdown de veículos não encontrado.")
            return []

    # ------------------------------------------------------------------
    # Loop por veículo
    # ------------------------------------------------------------------

    def _process_vehicle(
        self,
        vehicle_value: str,
        vehicle_label: str,
        windows: List[Tuple[date, date]],
    ):
        driver = self.driver

        for dt_inicio, dt_fim in windows:
            window_str = (
                f"{_fmt_date(dt_inicio)}"
                + (f" → {_fmt_date(dt_fim)}" if dt_inicio != dt_fim else "")
            )
            self._log(f"  Janela {window_str}...")

            for attempt in range(2):
                try:
                    self._select_vehicle(vehicle_value)
                    _set_date_range(driver, dt_inicio, dt_fim)
                    self._click_pesquisar()

                    if _has_no_results(driver):
                        self._log("    Sem registros.")
                        self.total_skipped += 1
                        break

                    if not _exportar_button_available(driver):
                        self._log("    Botão Exportar indisponível.")
                        self.total_skipped += 1
                        break

                    downloaded = self._click_exportar_and_wait(vehicle_label, window_str)
                    if downloaded:
                        self.total_downloads += 1
                        self._log(f"    Download: {downloaded}")
                    else:
                        self._log("    Download não concluído no tempo limite.")
                        self.errors.append(f"{vehicle_label} {window_str}: timeout download")
                    break

                except StaleElementReferenceException:
                    self._log("    Elemento stale — renavegando ao relatório...")
                    try:
                        navigate_to_report(driver)
                        self._select_domiciliar()
                    except Exception as re_e:
                        self.errors.append(
                            f"{vehicle_label} {window_str}: recuperação falhou: {re_e}"
                        )
                        break
                    if attempt > 0:
                        self.errors.append(
                            f"{vehicle_label} {window_str}: stale persistente após recovery"
                        )
                        break
                    # attempt == 0: continua para retry automático

                except Exception as e:
                    self._log(f"    Erro: {e}")
                    self.errors.append(f"{vehicle_label} {window_str}: {e}")
                    break

    def _select_vehicle(self, value: str):
        """Seleciona um veículo pelo value no dropdown."""
        driver = self.driver
        sel = Select(driver.find_element(By.ID, _ID_VEICULOS))
        sel.select_by_value(value)

    def _click_pesquisar(self):
        """Clica no botão Pesquisar e aguarda o resultado AJAX."""
        driver = self.driver
        btn = driver.find_element(By.ID, _ID_PESQUISAR)
        _js_click(driver, btn)
        wait_load_rascol(driver, timeout=120)

    def _click_exportar_and_wait(self, label: str, window_str: str) -> Optional[str]:
        """
        Clica em Exportar, aguarda o arquivo .xls na pasta de downloads
        e retorna o nome do arquivo.
        """
        driver = self.driver
        start_ts = time.time()
        btn = driver.find_element(By.ID, _ID_EXPORTAR)
        _js_click(driver, btn)

        # Aguarda o UpdatePanel terminar o postback do botão Exportar
        wait_load_rascol(driver, timeout=30)

        files = _wait_for_xls_download(RASCOL_DOWNLOAD_DIR, start_ts, timeout=120)
        if files:
            return os.path.basename(files[0])
        return None

    # ------------------------------------------------------------------
    # Pós-processamento
    # ------------------------------------------------------------------

    def _run_shapes_processor(self):
        """Converte os .xlsx baixados em shapefiles via ShapesProcessor."""
        from rascol_automation.processors.processor_shapes import ShapesProcessor
        self._log("\nIniciando geração de shapefiles...")
        processor = ShapesProcessor(
            excel_dir=RASCOL_DOWNLOAD_DIR,
            max_date=max(self.dates) if self.dates else None,
            log=self._log,
        )
        processor.run()
