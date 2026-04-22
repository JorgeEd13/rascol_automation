"""
RasCol Automation — GUI Principal

Segue o mesmo padrão visual do inlog_automation:
  - Calendário de seleção de datas (drag, shift+click, mês inteiro)
  - Campo de filial/contrato
  - Credenciais com toggle padrão/manual
  - Botão único "Extrair Pontos!"
"""

import os
import sys
import time as _time
import calendar
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from datetime import datetime, date
from typing import Optional, Dict, Any, List

from rascol_automation.config.rascol_config import load_rascol_config

try:
    import sv_ttk
    HAS_SV_TTK = True
except ImportError:
    HAS_SV_TTK = False

try:
    from tkcalendar import Calendar
    HAS_CALENDAR = True
except ImportError:
    HAS_CALENDAR = False


def _get_icon_path():
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            p = Path(meipass) / "logo_DIE.ico"
            if p.exists():
                return str(p)
        start = Path(sys.executable).parent
    else:
        start = Path(__file__).resolve().parent

    for directory in [start] + list(start.parents)[:6]:
        for candidate in [
            directory / "dependencias" / "logo_DIE.ico",
            directory / "logo_DIE.ico",
        ]:
            if candidate.exists():
                return str(candidate)
    return None


ICON_PATH = _get_icon_path()


class RasColGUI:
    """GUI da RasCol Automation."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("RasCol Automation")
        self.root.resizable(False, True)

        if ICON_PATH:
            self.root.iconbitmap(default=ICON_PATH)

        if HAS_SV_TTK:
            sv_ttk.set_theme("dark")

        ttk.Style().configure("Round.TCheckbutton", font=("Segoe UI", 11))

        self.selected_dates: set = set()
        self.result: Optional[Dict[str, Any]] = None

        # Drag-select state
        self._drag_start_date: Optional[date] = None
        self._drag_end_date:   Optional[date] = None
        self._drag_preview_dates: set = set()
        self._is_dragging = False
        self._drag_just_ended = False
        self._shift_held = False
        self._last_click_date: Optional[date] = None

        # Carrega configuração padrão
        self._rascol_cfg = load_rascol_config()

        self.main_frame = ttk.Frame(self.root, padding=12)
        self.main_frame.pack(fill="both", expand=True)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)

        self._create_widgets()

        self.root.bind("<KeyPress-Shift_L>",   lambda e: setattr(self, "_shift_held", True))
        self.root.bind("<KeyPress-Shift_R>",   lambda e: setattr(self, "_shift_held", True))
        self.root.bind("<KeyRelease-Shift_L>", lambda e: setattr(self, "_shift_held", False))
        self.root.bind("<KeyRelease-Shift_R>", lambda e: setattr(self, "_shift_held", False))

    # =========================================================================
    # Widgets
    # =========================================================================

    def _create_widgets(self):
        # Título
        ttk.Label(
            self.main_frame,
            text="RasCol — Relatório de Pontos de Operação",
            font=("Segoe UI", 12, "bold"),
        ).grid(row=0, column=0, columnspan=2, pady=(0, 8))

        # ---------- COL 0: Datas ----------
        left = ttk.Frame(self.main_frame, padding=6)
        left.grid(row=1, column=0, sticky="nw")

        ttk.Label(left, text="Selecione As Datas:", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        cal_frame = ttk.Frame(left)
        cal_frame.pack()
        if HAS_CALENDAR:
            self.calendar = Calendar(
                cal_frame,
                selectmode="day",
                locale="pt_BR",
                date_pattern="dd/mm/yyyy",
                firstweekday="sunday",
                showweeknumbers=False,
                tooltipdelay=0,
            )
            self.calendar.pack()
            self.calendar.bind("<<CalendarSelected>>",     self._on_day_click)
            self.calendar.bind("<<CalendarMonthChanged>>", self._on_month_change)
            self._bind_drag_events()

        cal_opts = ttk.Frame(left)
        cal_opts.pack(anchor="w", pady=(4, 0))
        self.full_month_var = tk.IntVar()
        _row = ttk.Frame(cal_opts)
        _row.pack(anchor="w")
        ttk.Checkbutton(
            _row, text="Mês Inteiro",
            style="Round.TCheckbutton",
            variable=self.full_month_var,
            command=self._toggle_full_month,
        ).pack(side="left")
        ttk.Button(_row, text="Limpar", command=self._clear_dates, width=7).pack(side="left", padx=(8, 0))
        ttk.Label(
            cal_opts,
            text="Arraste ou Shift+Click para intervalo",
            font=("Segoe UI", 8), foreground="gray",
        ).pack(anchor="w", pady=(2, 0))

        # ---------- COL 1: Config ----------
        right = ttk.Frame(self.main_frame, padding=6)
        right.grid(row=1, column=1, sticky="nw")

        # Filial
        ttk.Label(right, text="Filial / Contrato:", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.filial_var = tk.StringVar(value=self._rascol_cfg.filial)
        ttk.Entry(right, textvariable=self.filial_var, width=22).pack(anchor="w", pady=(2, 8))

        # Credenciais
        ttk.Separator(right, orient="horizontal").pack(fill="x", pady=4)
        ttk.Label(right, text="Credenciais RasCol:", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        self.use_default_var = tk.IntVar(value=1)
        ttk.Checkbutton(
            right, text="Usar credenciais padrão",
            variable=self.use_default_var,
            command=self._toggle_credentials,
        ).pack(anchor="w", pady=(2, 0))

        ttk.Label(right, text="Usuário:").pack(anchor="w")
        self.user_entry = ttk.Entry(right, width=22)
        self.user_entry.insert(0, self._rascol_cfg.username)
        self.user_entry.pack(anchor="w")

        ttk.Label(right, text="Senha:").pack(anchor="w")
        self.pass_entry = ttk.Entry(right, show="*", width=22)
        self.pass_entry.insert(0, self._rascol_cfg.password)
        self.pass_entry.pack(anchor="w")
        self._toggle_credentials()

        # Pós-processamento
        ttk.Separator(right, orient="horizontal").pack(fill="x", pady=8)
        ttk.Label(right, text="Pós-processamento:", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        self.post_process_var = tk.IntVar(value=1)
        ttk.Checkbutton(
            right, text="Gerar shapefiles após extração",
            variable=self.post_process_var,
            style="Round.TCheckbutton",
        ).pack(anchor="w", pady=(2, 0))

        ttk.Label(
            right,
            text="(executado após todos os downloads)",
            font=("Segoe UI", 8), foreground="gray",
        ).pack(anchor="w")

        # Indicação do .env
        if self._rascol_cfg.loaded:
            ttk.Label(
                right,
                text="📄 .env carregado",
                font=("Segoe UI", 8), foreground="gray",
            ).pack(anchor="w", pady=(8, 0))

        # ---------- Botão de execução ----------
        ttk.Button(
            self.main_frame,
            text="🗺️ Extrair Pontos!",
            command=self._on_run,
        ).grid(row=2, column=0, columnspan=2, pady=10)

        self.root.protocol("WM_DELETE_WINDOW", self._on_cancel)

    # =========================================================================
    # Credenciais
    # =========================================================================

    def _toggle_credentials(self):
        st = tk.DISABLED if self.use_default_var.get() else tk.NORMAL
        self.user_entry.configure(state=st)
        self.pass_entry.configure(state=st)

    def get_credentials(self):
        if self.use_default_var.get():
            return (self._rascol_cfg.username, self._rascol_cfg.password)
        return (self.user_entry.get(), self.pass_entry.get())

    # =========================================================================
    # Calendário (mesma lógica do inlog_automation)
    # =========================================================================

    def _get_displayed_month(self):
        if not HAS_CALENDAR:
            return (datetime.now().year, datetime.now().month)
        a, b = self.calendar.get_displayed_month()
        return (a, b) if a > 12 else (b, a)

    def _get_month_days(self, year, month):
        return [date(year, month, d) for d in range(1, calendar.monthrange(year, month)[1] + 1)]

    def _remove_visual(self, d):
        if HAS_CALENDAR:
            for ev in self.calendar.get_calevents(date=d):
                self.calendar.calevent_remove(ev)

    def _add_visual(self, d, tag="marked"):
        if HAS_CALENDAR:
            self.calendar.calevent_create(d, "sel", tag)

    def _date_range(self, d1, d2):
        start, end = min(d1, d2), max(d1, d2)
        days, current = [], start
        while current <= end:
            days.append(current)
            current = date.fromordinal(current.toordinal() + 1)
        return days

    def _bind_drag_events(self):
        if HAS_CALENDAR:
            self._bind_recursive(self.calendar)

    def _bind_recursive(self, widget):
        widget.bind("<ButtonPress-1>",   self._on_drag_start,  add=True)
        widget.bind("<B1-Motion>",       self._on_drag_motion, add=True)
        widget.bind("<ButtonRelease-1>", self._on_drag_end,    add=True)
        for child in widget.winfo_children():
            self._bind_recursive(child)

    def _get_date_under_cursor(self, event):
        if not HAS_CALENDAR:
            return None
        try:
            root_x = event.widget.winfo_rootx() + event.x
            root_y = event.widget.winfo_rooty() + event.y
            target = self.root.winfo_containing(root_x, root_y)
            if target is None:
                return None
            try:
                text = target.cget("text")
            except (tk.TclError, AttributeError):
                return None
            if not text or not text.strip().isdigit():
                return None
            day_num = int(text.strip())
            year, month = self._get_displayed_month()
            max_day = calendar.monthrange(year, month)[1]
            if 1 <= day_num <= max_day:
                return date(year, month, day_num)
        except Exception:
            pass
        return None

    def _on_drag_start(self, event):
        self._is_dragging = False
        self._drag_preview_dates.clear()
        self._drag_just_ended = False
        d = self._get_date_under_cursor(event)
        if d:
            self._drag_start_date = d

    def _on_drag_motion(self, event):
        if not self._drag_start_date:
            return
        d = self._get_date_under_cursor(event)
        if not d or d == self._drag_start_date:
            if not self._is_dragging:
                return
        if d and d != self._drag_start_date and not self._is_dragging:
            self._is_dragging = True
            self._drag_just_ended = True
        if not d or not self._is_dragging:
            return
        self._drag_end_date = d
        for pd in self._drag_preview_dates:
            if pd not in self.selected_dates:
                self._remove_visual(pd)
        new_range = self._date_range(self._drag_start_date, self._drag_end_date)
        self._drag_preview_dates = set(new_range)
        for rd in new_range:
            if rd not in self.selected_dates:
                self._add_visual(rd, "preview")
        try:
            self.calendar.tag_config("preview", background="#5599cc", foreground="white")
        except Exception:
            pass

    def _on_drag_end(self, event):
        if self._is_dragging and self._drag_start_date and self._drag_end_date:
            drag_range = self._date_range(self._drag_start_date, self._drag_end_date)
            for pd in self._drag_preview_dates:
                if pd not in self.selected_dates:
                    self._remove_visual(pd)
            for d in drag_range:
                if d not in self.selected_dates:
                    self.selected_dates.add(d)
                    self._add_visual(d, "marked")
            y, m = self._get_displayed_month()
            if set(self._get_month_days(y, m)).issubset(self.selected_dates):
                self.full_month_var.set(1)
            else:
                self.full_month_var.set(0)
            self._last_click_date = self._drag_end_date
        self._drag_start_date = None
        self._drag_end_date = None
        self._drag_preview_dates.clear()
        self._is_dragging = False

    def _on_day_click(self, event):
        if not HAS_CALENDAR:
            return
        if self._drag_just_ended:
            self._drag_just_ended = False
            return
        d = self.calendar.selection_get()
        if self._shift_held and self._last_click_date and self._last_click_date != d:
            for rd in self._date_range(self._last_click_date, d):
                if rd not in self.selected_dates:
                    self.selected_dates.add(rd)
                    self._add_visual(rd, "marked")
            self._last_click_date = d
        else:
            self._remove_visual(d)
            if d in self.selected_dates:
                self.selected_dates.remove(d)
            else:
                self.selected_dates.add(d)
                self._add_visual(d, "marked")
            self._last_click_date = d
        y, m = self._get_displayed_month()
        self.full_month_var.set(
            1 if set(self._get_month_days(y, m)).issubset(self.selected_dates) else 0
        )

    def _on_month_change(self, event=None):
        if self.full_month_var.get():
            self._toggle_full_month()

    def _toggle_full_month(self):
        if not HAS_CALENDAR:
            return
        y, m = self._get_displayed_month()
        days = self._get_month_days(y, m)
        if self.full_month_var.get():
            for d in days:
                if d not in self.selected_dates:
                    self.selected_dates.add(d)
                    self._add_visual(d, "marked")
        else:
            for d in days:
                self.selected_dates.discard(d)
                self._remove_visual(d)

    def _clear_dates(self):
        if HAS_CALENDAR:
            for ev in self.calendar.get_calevents():
                self.calendar.calevent_remove(ev)
        self.selected_dates.clear()
        self.full_month_var.set(0)

    def get_selected_dates(self) -> List[date]:
        return sorted(self.selected_dates)

    # =========================================================================
    # Execução
    # =========================================================================

    def _on_run(self):
        if not self.get_selected_dates():
            messagebox.showwarning("Atenção", "Selecione pelo menos uma data.")
            return

        user, password = self.get_credentials()
        if not user or not password:
            messagebox.showwarning("Atenção", "Preencha usuário e senha.")
            return

        dates = self.get_selected_dates()
        msg = [
            "Operação: Extrair Pontos de Operação (RasCol)",
            f"Filial: {self.filial_var.get()}",
            f"Datas: {len(dates)} ({dates[0].strftime('%d/%m/%Y')} → {dates[-1].strftime('%d/%m/%Y')})",
            "\nIniciar?",
        ]
        if messagebox.askyesno("Confirmar", "\n".join(msg)):
            self.result = {
                "cancelled":     False,
                "dates":         dates,
                "username":      user,
                "password":      password,
                "filial":        self.filial_var.get().strip().upper(),
                "post_process":  bool(self.post_process_var.get()),
            }
            self.root.destroy()

    def _on_cancel(self):
        self.result = {"cancelled": True}
        self.root.destroy()

    def run(self) -> Optional[Dict[str, Any]]:
        self.root.mainloop()
        return self.result


# ---------------------------------------------------------------------------
# Janela de progresso (reusa do inlog_automation)
# ---------------------------------------------------------------------------

from inlog_automation.gui.main_gui import (  # noqa: E402
    ProgressWindow,
    show_result_dialog,
)
