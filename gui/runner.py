"""
RasCol Automation — Runner

Integra a GUI com o PontosExtractor.
"""

from rascol_automation.gui.main_gui import RasColGUI, ProgressWindow, show_result_dialog
from rascol_automation.config.rascol_config import load_rascol_config
from rascol_automation.config.settings import RASCOL_DOWNLOAD_DIR, RASCOL_SHAPES_DIR, INLOG_SHAPES_DIR
from rascol_automation.extractors.extractor_pontos import PontosExtractor


def main():
    # 1. GUI
    gui = RasColGUI()
    result = gui.run()

    if result is None or result.get("cancelled"):
        return

    # 2. Monta configuração a partir do resultado da GUI
    cfg = load_rascol_config()
    cfg.username = result["username"]
    cfg.password = result["password"]
    cfg.filial   = result["filial"]

    # 3. Cria extractor com callback de progresso (para ProgressWindow)
    messages = []

    def log_msg(msg: str):
        messages.append(msg)
        print(msg)

    extractor = PontosExtractor(
        dates=result["dates"],
        config=cfg,
        run_shapes=result.get("post_process", False),
        progress_callback=log_msg,
    )

    # 4. ProgressWindow
    pw = ProgressWindow("Extraindo Pontos de Operação (RasCol)...")

    error_msg = None

    def task():
        nonlocal error_msg
        try:
            extractor.run()
        except Exception as e:
            error_msg = str(e)
            import traceback
            traceback.print_exc()

    _, run_err, _ = pw.run(task)
    if run_err and not error_msg:
        error_msg = str(run_err)

    # 5. Resultado
    if error_msg:
        show_result_dialog(
            success=False,
            message="Erro durante a extração",
            details=error_msg,
        )
    else:
        details = (
            f"Veículos: {extractor.total_vehicles}\n"
            f"Downloads: {extractor.total_downloads}\n"
            f"Sem registros: {extractor.total_skipped}\n"
            f"Erros: {len(extractor.errors)}\n"
            f"Downloads: {RASCOL_DOWNLOAD_DIR}\n"
            + (f"Shapefiles: {INLOG_SHAPES_DIR} / {RASCOL_SHAPES_DIR.name}" if result.get("post_process") else "")
        )
        if extractor.errors:
            details += "\n\nErros:\n" + "\n".join(extractor.errors[:10])
        show_result_dialog(
            success=True,
            message="Extração RasCol concluída!",
            details=details,
        )


if __name__ == "__main__":
    main()
