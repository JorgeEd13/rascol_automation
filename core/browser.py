"""
Browser Selenium configurado para RasCol.
"""

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from rascol_automation.config.settings import RASCOL_DOWNLOAD_DIR, SELENIUM_SCALE

RASCOL_LOGIN_URL = (
    "https://id.rassystem.com.br/cas/login"
    "?service=https%3a%2f%2frascol.rassystem.com.br"
    "%2fGeomapas%2fEmpresaFilialSeleciona.aspx"
)


def open_browser():
    """Abre Chrome configurado para o RasCol."""
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument(f"--force-device-scale-factor={SELENIUM_SCALE}")
    options.add_argument("--log-level=3")
    options.add_experimental_option(
        "excludeSwitches", ["enable-automation", "enable-logging"]
    )
    options.add_experimental_option("useAutomationExtension", False)

    prefs = {
        "download.default_directory": str(RASCOL_DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        # Permite downloads múltiplos automáticos sem popup de confirmação
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=options)
    driver.get(RASCOL_LOGIN_URL)
    return driver
