# -*- mode: python ; coding: utf-8 -*-
"""
RasColAutomation.spec — Configuração do PyInstaller
====================================================

Para gerar o executável, execute a partir de Locar/rascol_automation/:
    pyinstaller RasColAutomation.spec

O executável será gerado em dist/RasColAutomation.exe
Copie-o para a pasta Locar/ (ao lado de InlogAutomation.exe).
"""

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

BASE_DIR = Path(SPECPATH)                           # Locar/rascol_automation/
LOCAR_DIR = BASE_DIR.parent                         # Locar/
ICON_PATH = LOCAR_DIR / "dependencias" / "logo_DIE.ico"
icon = str(ICON_PATH) if ICON_PATH.exists() else None

selenium_hidden      = collect_submodules('selenium')
webdriver_mgr_hidden = collect_submodules('webdriver_manager')

datas = []
if ICON_PATH.exists():
    datas.append((str(ICON_PATH), '.'))
try:
    datas += collect_data_files('selenium')
except Exception as e:
    print(f"⚠ selenium data files: {e}")
try:
    datas += collect_data_files('webdriver_manager')
except Exception as e:
    print(f"⚠ webdriver_manager data files: {e}")

a = Analysis(
    ['run.py'],
    pathex=[str(LOCAR_DIR)],    # torna inlog_automation e rascol_automation importáveis
    binaries=[],
    datas=datas,
    hiddenimports=[
        # ── rascol_automation ────────────────────────────────────────────
        'rascol_automation',
        'rascol_automation.config',
        'rascol_automation.config.settings',
        'rascol_automation.config.rascol_config',
        'rascol_automation.core',
        'rascol_automation.core.browser',
        'rascol_automation.core.waits',
        'rascol_automation.core.auth',
        'rascol_automation.extractors',
        'rascol_automation.extractors.extractor_pontos',
        'rascol_automation.gui',
        'rascol_automation.gui.main_gui',
        'rascol_automation.gui.runner',
        'rascol_automation.processors',
        'rascol_automation.processors.processor_shapes',

        # ── inlog_automation (compartilhado) ─────────────────────────────
        'inlog_automation',
        'inlog_automation.config',
        'inlog_automation.config.settings',
        'inlog_automation.config.user_config',
        'inlog_automation.core',
        'inlog_automation.core.waits',
        'inlog_automation.gui',
        'inlog_automation.gui.main_gui',

        # ── Tkinter ──────────────────────────────────────────────────────
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'tkinter.scrolledtext',

        # ── Selenium / WebDriverManager ──────────────────────────────────
        *selenium_hidden,
        *webdriver_mgr_hidden,

        # ── Pandas / OpenPyXL ────────────────────────────────────────────
        'pandas',
        'pandas._libs',
        'pandas._libs.tslibs.timedeltas',
        'openpyxl',
        'openpyxl.cell._writer',

        # ── NumPy ────────────────────────────────────────────────────────
        'numpy',
        'numpy.core',
        'numpy.core.multiarray',

        # ── GeoPandas / Shapely / Fiona / PyProj ─────────────────────────
        'geopandas',
        'shapely',
        'shapely.geometry',
        'fiona',
        'pyproj',

        # ── GUI extras ───────────────────────────────────────────────────
        'sv_ttk',
        'tkcalendar',

        # ── Stdlib extras ────────────────────────────────────────────────
        'dataclasses',
        'zipfile',
        'tempfile',
        'unicodedata',
        'packaging',
        'packaging.version',
        'packaging.specifiers',
        'packaging.requirements',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy.random._examples',
        'pytest',
        'IPython',
        'jupyter',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='RasColAutomation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon,
    version=None,
    manifest=None,
)

print("\n" + "=" * 60)
print("CONFIGURAÇÃO DO BUILD")
print("=" * 60)
print(f"  Nome:   RasColAutomation.exe")
print(f"  Ícone:  {icon if icon else 'Não definido'}")
print(f"  Console: Desabilitado")
print(f"  Modo:   Executável único (onefile)")
print("=" * 60 + "\n")
