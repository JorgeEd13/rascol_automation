"""
Funções de espera específicas do RasCol.

Reutiliza os helpers genéricos do inlog_automation e adiciona
wait_load_rascol para os indicadores de carregamento do RasCol.
"""

import time

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Re-exporta helpers genéricos de clique
from inlog_automation.core.waits import _js_click, safe_click  # noqa: F401


def wait_load_rascol(driver, timeout=120):
    """
    Aguarda o carregamento AJAX do RasCol completar.

    Verifica a invisibilidade do elemento #loading, que é exibido
    dentro de #processMessage durante postbacks UpdatePanel.
    """
    time.sleep(0.4)
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.ID, "loading"))
        )
    except (TimeoutException, NoSuchElementException):
        pass


def wait_for_element(driver, by, value, timeout=30):
    """Aguarda um elemento ficar presente no DOM."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
        return driver.find_element(by, value)
    except (TimeoutException, NoSuchElementException):
        return None


def wait_and_click(driver, by, value, timeout=30):
    """Aguarda elemento e clica via JavaScript."""
    el = wait_for_element(driver, by, value, timeout)
    if el:
        _js_click(driver, el)
    return el
