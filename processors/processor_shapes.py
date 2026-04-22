"""
Converte os arquivos .xlsx baixados do RasCol em shapefiles de segmentos GPS.

Destino do ZIP por dia:
  - Se já existir um ZIP para esse dia em dependencias/shapes (pasta da Inlog),
    os shapefiles são adicionados a esse ZIP existente.
  - Caso contrário, o ZIP é criado (ou acrescentado) em dependencias/rascol_shapes.

Para cada Excel processado com sucesso o arquivo é apagado.
"""

import zipfile
import unicodedata
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, List, Optional

import pandas as pd
import geopandas as gpd
from shapely.geometry import LineString

from rascol_automation.config.settings import (
    RASCOL_DOWNLOAD_DIR,
    RASCOL_SHAPES_DIR,
    INLOG_SHAPES_DIR,
)

HORA_CORTE    = 5
CONTRATO_SAIDA = "JABOATAO"


def _normalizar_placa(placa: str) -> str:
    return placa.replace("-", "").replace(" ", "").upper()


class ShapesProcessor:
    """
    Processa todos os .xlsx em excel_dir e gera ZIPs de shapefiles.

    Regra de destino por dia:
      - Se dependencias/shapes/<zip> existir  →  acrescenta a esse ZIP
      - Se não existir                        →  cria/acrescenta em rascol_shapes/<zip>

    Args:
        excel_dir:  Pasta com os .xlsx de entrada (padrão: RASCOL_DOWNLOAD_DIR).
        max_date:   Último dia operacional permitido (inclusive); evita dias espúrios.
        hora_corte: Hora de virada do dia operacional (padrão: 5).
        contrato:   Sufixo nos nomes de arquivo (padrão: JABOATAO).
        log:        Função de log — recebe str (padrão: print).
    """

    def __init__(
        self,
        excel_dir: Optional[Path] = None,
        max_date: Optional[date] = None,
        hora_corte: int = HORA_CORTE,
        contrato: str = CONTRATO_SAIDA,
        log: Optional[Callable[[str], None]] = None,
    ):
        self.excel_dir  = excel_dir  or RASCOL_DOWNLOAD_DIR
        self.max_date   = max_date
        self.hora_corte = hora_corte
        self.contrato   = contrato
        self._log       = log or print

        RASCOL_SHAPES_DIR.mkdir(parents=True, exist_ok=True)

        self.total_files  = 0
        self.total_shapes = 0
        self.errors: List[str] = []
        self._ok_files: List[Path] = []

    # ------------------------------------------------------------------

    def run(self):
        """Processa todos os Excel e apaga os que foram convertidos com sucesso."""
        xlsx_files = list(self.excel_dir.glob("*.xlsx"))
        if not xlsx_files:
            self._log("Nenhum arquivo .xlsx encontrado para processar.")
            return

        self._log(f"\nProcessando {len(xlsx_files)} arquivo(s) Excel...")
        for excel_file in xlsx_files:
            try:
                self._process_file(excel_file)
                self.total_files += 1
                self._ok_files.append(excel_file)
            except Exception as e:
                msg = f"{excel_file.name}: {e}"
                self._log(f"  Erro: {msg}")
                self.errors.append(msg)

        self._delete_processed_excels()

        self._log(
            f"\nShapes concluído — arquivos: {self.total_files} | "
            f"shapefiles: {self.total_shapes} | "
            f"erros: {len(self.errors)}"
        )

    # ------------------------------------------------------------------
    # Processamento individual de cada Excel
    # ------------------------------------------------------------------

    def _process_file(self, excel_file: Path):
        self._log(f"\n  {excel_file.name}")

        df_header = pd.read_excel(excel_file, sheet_name=0, header=None)
        placa    = str(df_header.iloc[3, 1]).strip()   # B4
        contrato = str(df_header.iloc[2, 4]).strip()   # E3

        if placa.lower() in ("nan", "", "none"):
            raise ValueError("Placa inválida no cabeçalho")
        if contrato.lower() in ("nan", "", "none"):
            raise ValueError("Contrato inválido no cabeçalho")

        placa_norm = _normalizar_placa(placa)

        df = pd.read_excel(excel_file, sheet_name=0, skiprows=6)
        df.columns = [c.strip().lower() for c in df.columns]

        if not {"data/hora", "latitude", "longitude"}.issubset(df.columns):
            raise ValueError("Colunas esperadas não encontradas (data/hora, latitude, longitude)")

        df["datahora"]   = pd.to_datetime(df["data/hora"], errors="coerce")
        df["latitude"]   = pd.to_numeric(df["latitude"],  errors="coerce")
        df["longitude"]  = pd.to_numeric(df["longitude"], errors="coerce")
        df["velocidade"] = (
            pd.to_numeric(df["velocidade"], errors="coerce")
            if "velocidade" in df.columns
            else None
        )

        df = df.dropna(subset=["datahora", "latitude", "longitude"])
        if df.empty:
            self._log("    Nenhum ponto válido.")
            return

        df["dia_operacional"] = (
            df["datahora"] - pd.Timedelta(hours=self.hora_corte)
        ).dt.date

        for dia in sorted(df["dia_operacional"].unique()):
            if self.max_date and dia > self.max_date:
                continue
            self._generate_and_zip(df, dia, placa_norm)

    # ------------------------------------------------------------------
    # Geração do shapefile e adição ao ZIP correto
    # ------------------------------------------------------------------

    def _resolve_zip_path(self, data_str: str) -> Path:
        """
        Retorna o caminho do ZIP de destino para o dia data_str.

        Prioridade: ZIP existente em dependencias/shapes; fallback: rascol_shapes.
        """
        zip_name  = f"Shapes - {self.contrato.capitalize()} - {data_str}.zip"
        inlog_zip = INLOG_SHAPES_DIR / zip_name
        if inlog_zip.exists():
            return inlog_zip
        return RASCOL_SHAPES_DIR / zip_name

    def _generate_and_zip(self, df: "pd.DataFrame", dia, placa_norm: str):
        inicio = datetime.combine(dia, datetime.min.time()) + timedelta(hours=self.hora_corte)
        fim    = inicio + timedelta(days=1)

        df_dia = (
            df[(df["datahora"] >= inicio) & (df["datahora"] < fim)]
            .sort_values("datahora")
            .reset_index(drop=True)
        )

        if len(df_dia) < 2:
            return

        data_str = inicio.strftime("%d.%m.%Y")
        shp_base = f"{data_str}_{placa_norm}_{self.contrato}"
        zip_path = self._resolve_zip_path(data_str)
        mode     = "a" if zip_path.exists() else "w"

        segmentos = []
        for i in range(len(df_dia) - 1):
            p1, p2 = df_dia.iloc[i], df_dia.iloc[i + 1]
            segmentos.append({
                "DATAHORA":   p1["datahora"].strftime("%d/%m/%Y %H:%M:%S"),
                "VELOCIDADE": p1["velocidade"],
                "geometry":   LineString([
                    (p1["longitude"], p1["latitude"]),
                    (p2["longitude"], p2["latitude"]),
                ]),
            })

        gdf = gpd.GeoDataFrame(segmentos, crs="EPSG:4326")

        with tempfile.TemporaryDirectory() as tmpdir:
            shp_path = Path(tmpdir) / f"{shp_base}.shp"
            gdf.to_file(shp_path)

            with zipfile.ZipFile(zip_path, mode, zipfile.ZIP_DEFLATED) as zipf:
                for f in Path(tmpdir).glob("*"):
                    zipf.write(f, arcname=f.name)

        self.total_shapes += 1
        self._log(f"    ✔ {shp_base}.shp  ({len(segmentos)} seg.) → {zip_path.parent.name}/{zip_path.name}")

    # ------------------------------------------------------------------
    # Limpeza dos Excel processados
    # ------------------------------------------------------------------

    def _delete_processed_excels(self):
        if not self._ok_files:
            return
        self._log("\nApagando Excel processados...")
        for f in self._ok_files:
            try:
                f.unlink()
                self._log(f"  Apagado: {f.name}")
            except Exception as e:
                self._log(f"  Aviso: não foi possível apagar {f.name}: {e}")
