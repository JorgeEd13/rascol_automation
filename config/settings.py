"""
Configurações do rascol_automation.
Compartilha caminhos base com inlog_automation via DEPS_DIR.
"""

from inlog_automation.config.settings import (  # noqa: F401
    DEPS_DIR,
    LOGS_DIR,
    SELENIUM_SCALE,
    SHAPES_DIR as INLOG_SHAPES_DIR,   # dependencias/shapes  (ZIPs da Inlog)
)

RASCOL_DOWNLOAD_DIR = DEPS_DIR / "rascol_downloads"
RASCOL_DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

RASCOL_SHAPES_DIR = DEPS_DIR / "rascol_shapes"    # destino fallback dos ZIPs RasCol
RASCOL_SHAPES_DIR.mkdir(parents=True, exist_ok=True)
