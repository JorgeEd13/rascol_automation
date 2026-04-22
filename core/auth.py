"""
Autenticação e navegação no RasCol.

Cobre as etapas:
  1. Login (página 1)
  2. Seleção de empresa "Locar" (página 2)
  3. Seleção da filial por keyword (página 3)
  4. Navegação ao Relatório de Pontos de Operação (página 4 → 5)
"""

import time
import unicodedata

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

from rascol_automation.core.waits import wait_load_rascol, wait_for_element, _js_click

RELATORIO_URL = (
    "https://rascol.rassystem.com.br/Operacoes/RelatorioPontosVeiculos.aspx"
)


def _normalize(text: str) -> str:
    """Remove acentos e converte para maiúsculas."""
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).upper()


def login_rascol(driver, username: str, password: str):
    """Preenche credenciais e faz login no RasCol CAS."""
    wait_for_element(driver, By.ID, "username", timeout=30)

    user_el = driver.find_element(By.ID, "username")
    pass_el = driver.find_element(By.ID, "password")

    user_el.clear()
    user_el.send_keys(username)
    pass_el.clear()
    pass_el.send_keys(password)

    driver.find_element(By.ID, "submitBtn").click()

    # Aguarda redirecionamento para EmpresaFilialSeleciona.aspx
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located(
            (By.ID, "ctl00_ContentPlaceHolder_cboEmpresa")
        )
    )


def select_company_locar(driver):
    """Seleciona 'Locar' no dropdown de empresa e aguarda a lista de filiais."""
    sel = Select(
        driver.find_element(By.ID, "ctl00_ContentPlaceHolder_cboEmpresa")
    )
    sel.select_by_value("542")  # Locar
    wait_load_rascol(driver)

    # Aguarda o accordion de filiais aparecer
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "tdSetaEscolherFilial"))
    )


def select_filial(driver, filial_keyword: str = "JABOATAO"):
    """
    Encontra a filial cujo nome normalizado contém filial_keyword
    e clica no botão de acesso.

    Como a Locar tem apenas uma filial (Jaboatão), o ID
    'tdSetaEscolherFilial' é único — mas verificamos o nome por segurança.
    """
    keyword = _normalize(filial_keyword)

    # Lê todos os spans de nome de filial e compara normalizando texto
    try:
        spans = driver.find_elements(
            By.XPATH, "//span[contains(@id, 'lblNomeFilial')]"
        )
        matched = False
        for span in spans:
            if keyword in _normalize(span.text):
                # Tenta clicar no tdSetaEscolherFilial dentro do mesmo bloco
                try:
                    arrow = span.find_element(
                        By.XPATH,
                        "ancestor::table//td[contains(@id, 'tdSetaEscolherFilial') "
                        "or @id='tdSetaEscolherFilial']"
                    )
                    _js_click(driver, arrow)
                    matched = True
                    break
                except NoSuchElementException:
                    pass
        if not matched:
            raise NoSuchElementException("Filial não encontrada por keyword")
    except NoSuchElementException:
        # Fallback: clica no único tdSetaEscolherFilial disponível
        arrow = driver.find_element(By.ID, "tdSetaEscolherFilial")
        _js_click(driver, arrow)

    # Aguarda a página principal com o menu de relatórios
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located(
            (By.XPATH, "//div[@role='treeitem']")
        )
    )
    time.sleep(0.5)


def navigate_to_report(driver):
    """
    Expande o menu 'Relatórios' e navega para
    'Relatório de Pontos de Operação'.
    """
    # Clica em "Relatórios" para expandir o submenu
    try:
        relatorios = driver.find_element(
            By.XPATH,
            "//div[@role='treeitem' and contains(., 'Relatórios') "
            "and @aria-level='1']"
        )
        _js_click(driver, relatorios)
        time.sleep(0.5)
    except NoSuchElementException:
        pass

    # Clica em "Relatório de Pontos de Operação"
    try:
        pontos_item = driver.find_element(
            By.XPATH,
            "//div[@role='treeitem' and contains(., 'Relatório de Pontos')]"
        )
        # Lê o href caso exista e usa window.location para navegar
        href = pontos_item.get_attribute("href")
        if href:
            driver.execute_script(f"window.location.href = '{href}';")
        else:
            _js_click(driver, pontos_item)
    except NoSuchElementException:
        # Fallback direto pela URL
        driver.get(RELATORIO_URL)

    # Aguarda o formulário do relatório
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located(
            (By.ID, "cad_ctl00_ContentPlaceHolder_Filtro_rotulos_ddlRotulos")
        )
    )
    wait_load_rascol(driver)
