"""
Leitura da seção [RASCOL] do .env compartilhado.

Formato esperado no .env:

    [RASCOL]
    usuario = 
    senha   = 
    filial  = JABOATAO
"""

from inlog_automation.config.user_config import _find_config_file

DEFAULT_RASCOL_USERNAME = ""
DEFAULT_RASCOL_PASSWORD = ""
DEFAULT_RASCOL_FILIAL   = "JABOATAO"


class RasColConfig:
    def __init__(self):
        self.username: str = DEFAULT_RASCOL_USERNAME
        self.password: str = DEFAULT_RASCOL_PASSWORD
        self.filial:   str = DEFAULT_RASCOL_FILIAL
        self.loaded:   bool = False

    @property
    def has_credentials(self) -> bool:
        return bool(self.username and self.password)

    def __repr__(self):
        return (
            f"RasColConfig(loaded={self.loaded}, "
            f"username='{self.username}', filial='{self.filial}')"
        )


def load_rascol_config() -> RasColConfig:
    """Procura e carrega a seção [RASCOL] do .env."""
    cfg = RasColConfig()
    config_path = _find_config_file()
    if config_path is None:
        return cfg

    try:
        text = config_path.read_text(encoding="utf-8")
    except Exception:
        try:
            text = config_path.read_text(encoding="latin-1")
        except Exception:
            return cfg

    in_rascol = False
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("[") and line.endswith("]"):
            in_rascol = (line[1:-1].strip().upper() == "RASCOL")
            continue
        if not in_rascol or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip().lower()
        value = value.strip()
        if not value:
            continue
        if key in ("usuario", "usuário", "user", "username"):
            cfg.username = value
        elif key in ("senha", "password", "pass"):
            cfg.password = value
        elif key in ("filial", "contrato"):
            cfg.filial = value.upper()

    cfg.loaded = True
    return cfg
