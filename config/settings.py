"""
Configurações do rascol_automation.
Compartilha caminhos base com inlog_automation via DEPS_DIR.
"""

from inlog_automation.config.settings import DEPS_DIR, LOGS_DIR, SELENIUM_SCALE  # noqa: F401

RASCOL_DOWNLOAD_DIR = DEPS_DIR / "rascol_downloads"
RASCOL_DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

SHAPES_DIR = DEPS_DIR / "rascol_shapes"
SHAPES_DIR.mkdir(parents=True, exist_ok=True)
