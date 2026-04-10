from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from desktop_gui.config import BG, FONT, FONT_BOLD, FONT_SMALL, MUTED, PANEL, PRIMARY, TEXT, WEEKDAY_NAMES
from desktop_gui.utils import format_hours, normalize_short_date, parse_date, parse_hours, progress_rank

try:
    from tkcalendar import Calendar
except Exception:
    Calendar = None


class ScrollFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        self.scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas, bg=BG)
        self.window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scroll.pack(side="right", fill="y")
        self.inner.bind("<Configure>", lambda _: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.window, width=e.width))


class EntryField(tk.Frame):
    def __init__(self, parent, label: str, width: int = 28, on_change=None):
        super().__init__(parent, bg=BG)
        self.on_change = on_change
        tk.Label(self, text=label, bg=BG, fg=TEXT, font=FONT_SMALL, anchor="w").pack(fill="x")
        self.var = tk.StringVar()
        ttk.Entry(self, textvariable=self.var, width=width).pack(fill="x", pady=(2, 0))
        if self.on_change:
            self.var.trace_add("write", lambda *_: self.on_change())

    def get(self) -> str:
        return self.var.get().strip()

    def set(self, value: str) -> None:
        self.var.set(value or "")


class DateField(tk.Frame):
    def __init__(
        self,
        parent,
        label: str,
        include_year: bool = True,
        year_range: tuple[int, int] | None = None,
        on_change=None,
    ):
        super().__init__(parent, bg=BG)
        self.include_year = include_year
        self.on_change = on_change
        self.current_year = datetime.now().year
        start_year, end_year = year_range or (self.current_year - 2, self.current_year + 3)
        self.year_values = [str(i) for i in range(start_year, end_year + 1)]
        self.var = tk.StringVar()
        self.widget = None
        self.popup = None
        self.calendar = None

        tk.Label(self, text=label, bg=BG, fg=TEXT, font=FONT_SMALL, anchor="w").pack(fill="x")
        row = tk.Frame(self, bg=BG)
        row.pack(fill="x", pady=(2, 0))

        if Calendar is not None and include_year:
            self.widget = ttk.Entry(row, textvariable=self.var, width=14, state="readonly")
            self.widget.pack(side="left")
            ttk.Button(row, text="Elegir", command=self.open_picker).pack(side="left", padx=(6, 4))
            ttk.Button(row, text="Limpiar", command=self.clear).pack(side="left")
        else:
            self.day = tk.StringVar()
            self.month = tk.StringVar()
            self.year = tk.StringVar(value=str(self.current_year))
            ttk.Combobox(row, textvariable=self.day, width=4, values=[f"{i:02d}" for i in range(1, 32)], state="readonly").pack(side="left")
            tk.Label(row, text="/", bg=BG, fg=MUTED).pack(side="left", padx=2)
            ttk.Combobox(row, textvariable=self.month, width=4, values=[f"{i:02d}" for i in range(1, 13)], state="readonly").pack(side="left")

            if include_year:
                tk.Label(row, text="/", bg=BG, fg=MUTED).pack(side="left", padx=2)
                ttk.Combobox(row, textvariable=self.year, width=6, values=self.year_values, state="readonly").pack(side="left")

            self.day.trace_add("write", lambda *_: self._notify())
            self.month.trace_add("write", lambda *_: self._notify())
            self.year.trace_add("write", lambda *_: self._notify())

    def _notify(self) -> None:
        if self.on_change:
            self.on_change()

    def open_picker(self) -> None:
        if Calendar is None:
            return
        if self.popup is not None and self.popup.winfo_exists():
            self.popup.lift()
            return

        self.popup = tk.Toplevel(self)
        self.popup.title("Seleccionar fecha")
        self.popup.resizable(False, False)
        self.popup.transient(self.winfo_toplevel())
        self.popup.grab_set()

        selected = self.get() or datetime.now().strftime("%d/%m/%Y")
        try:
            current = parse_date(selected)
        except ValueError:
            current = datetime.now()

        self.calendar = Calendar(
            self.popup,
            selectmode="day",
            year=current.year,
            month=current.month,
            day=current.day,
            date_pattern="dd/mm/yyyy",
            firstweekday="monday",
        )
        self.calendar.pack(padx=10, pady=10)

        buttons = tk.Frame(self.popup, bg=BG)
        buttons.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(buttons, text="Usar fecha", command=self.use_selected_date).pack(side="right")
        ttk.Button(buttons, text="Cancelar", command=self.close_picker).pack(side="right", padx=(0, 8))

    def use_selected_date(self) -> None:
        if self.calendar is None:
            return
        self.var.set(self.calendar.get_date())
        self.close_picker()
        self._notify()

    def close_picker(self) -> None:
        if self.popup is not None and self.popup.winfo_exists():
            self.popup.destroy()
        self.popup = None
        self.calendar = None

    def clear(self) -> None:
        if self.widget is not None:
            self.var.set("")
        else:
            self.day.set("")
            self.month.set("")
            if self.include_year:
                self.year.set(str(self.current_year))
        self._notify()

    def get(self) -> str:
        if self.widget is not None:
            return self.var.get().strip()
        if not self.day.get() or not self.month.get():
            return ""
        if self.include_year:
            if not self.year.get():
                return ""
            return f"{self.day.get()}/{self.month.get()}/{self.year.get()}"
        return f"{self.day.get()}/{self.month.get()}"

    def set(self, value: str) -> None:
        raw = (value or "").strip()
        if not raw:
            self.clear()
            return

        if self.widget is not None:
            self.var.set(raw)
        else:
            parts = raw.split("/")
            if len(parts) >= 2:
                self.day.set(parts[0].zfill(2))
                self.month.set(parts[1].zfill(2))
            if self.include_year and len(parts) >= 3:
                self.year.set(parts[2])
        self._notify()


class SegmentedField(tk.Frame):
    def __init__(self, parent, label: str, segments: list[str], on_change=None):
        super().__init__(parent, bg=BG)
        self.vars = [tk.StringVar() for _ in segments]
        self.on_change = on_change

        tk.Label(self, text=label, bg=BG, fg=TEXT, font=FONT_SMALL, anchor="w").pack(fill="x")
        row = tk.Frame(self, bg=BG)
        row.pack(fill="x", pady=(2, 0))

        for index, (segment, var) in enumerate(zip(segments, self.vars)):
            cell = tk.Frame(row, bg=BG)
            cell.pack(side="left", fill="x", expand=True)
            ttk.Entry(cell, textvariable=var).pack(fill="x")
            tk.Label(cell, text=segment, bg=BG, fg=MUTED, font=FONT_SMALL, anchor="w").pack(fill="x")
            if self.on_change:
                var.trace_add("write", lambda *_: self.on_change())
            if index < len(segments) - 1:
                tk.Label(row, text=" / ", bg=BG, fg=MUTED, font=FONT_BOLD).pack(side="left", padx=4)

    def get(self) -> str:
        parts = [var.get().strip() for var in self.vars if var.get().strip()]
        return " / ".join(parts)

    def get_parts(self) -> list[str]:
        return [var.get().strip() for var in self.vars]

    def set(self, value: str) -> None:
        parts = [part.strip() for part in (value or "").split("/")]
        for index, var in enumerate(self.vars):
            var.set(parts[index] if index < len(parts) else "")
        if self.on_change:
            self.on_change()


class DurationField(tk.Frame):
    def __init__(self, parent, on_change):
        super().__init__(parent, bg=PANEL)
        self.on_change = on_change
        self.hours = tk.StringVar(value="0")
        self.minutes = tk.StringVar(value="00")

        ttk.Spinbox(self, from_=0, to=12, width=4, textvariable=self.hours, command=self._notify).pack(side="left")
        tk.Label(self, text="h", bg=PANEL, fg=MUTED, font=FONT_SMALL).pack(side="left", padx=(2, 6))
        ttk.Combobox(self, width=4, textvariable=self.minutes, values=[f"{i:02d}" for i in range(0, 60, 5)], state="readonly").pack(side="left")
        tk.Label(self, text="m", bg=PANEL, fg=MUTED, font=FONT_SMALL).pack(side="left", padx=(2, 0))

        self.hours.trace_add("write", lambda *_: self._notify())
        self.minutes.trace_add("write", lambda *_: self._notify())

    def _notify(self) -> None:
        self.on_change()

    def get_value(self) -> str:
        hours = self.hours.get().strip() or "0"
        minutes = self.minutes.get().strip() or "00"
        return f"{int(hours)}:{minutes}"

    def get_float(self) -> float:
        try:
            hours = int(self.hours.get().strip() or "0")
        except ValueError:
            hours = 0
        try:
            minutes = int(self.minutes.get().strip() or "0")
        except ValueError:
            minutes = 0
        return max(hours, 0) + max(minutes, 0) / 60.0

    def set_value(self, value: str) -> None:
        raw = (value or "").strip()
        if not raw:
            self.hours.set("0")
            self.minutes.set("00")
            return
        if ":" in raw:
            h, m = raw.split(":", 1)
            try:
                hours = max(int(h or 0), 0)
                minutes = max(int(m or 0), 0)
            except ValueError:
                hours = 0
                minutes = 0
            self.hours.set(str(hours))
            self.minutes.set(f"{minutes:02d}")
            return
        number = parse_hours(raw)
        total_minutes = round(number * 60)
        self.hours.set(str(total_minutes // 60))
        self.minutes.set(f"{total_minutes % 60:02d}")


class TextField(tk.Frame):
    def __init__(self, parent, label: str, height: int = 4, hint: str = ""):
        super().__init__(parent, bg=BG)
        tk.Label(self, text=label, bg=BG, fg=TEXT, font=FONT_SMALL, anchor="w").pack(fill="x")
        if hint:
            tk.Label(self, text=hint, bg=BG, fg=MUTED, font=FONT_SMALL, anchor="w", justify="left").pack(fill="x")
        self.text = tk.Text(self, height=height, wrap="word", relief="solid", bd=1, font=FONT)
        self.text.pack(fill="both", expand=True, pady=(2, 0))

    def get(self) -> str:
        return self.text.get("1.0", "end-1c").strip()

    def set(self, value: str) -> None:
        self.text.delete("1.0", tk.END)
        if value:
            self.text.insert("1.0", value)


class FileField(tk.Frame):
    def __init__(self, parent, label: str, types, on_pick=None):
        super().__init__(parent, bg=BG)
        self.types = types
        self.on_pick = on_pick
        tk.Label(self, text=label, bg=BG, fg=TEXT, font=FONT_SMALL, anchor="w").pack(fill="x")
        row = tk.Frame(self, bg=BG)
        row.pack(fill="x", pady=(2, 0))
        self.var = tk.StringVar()
        ttk.Entry(row, textvariable=self.var).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Buscar", command=self.pick).pack(side="left", padx=(6, 0))

    def pick(self) -> None:
        path = filedialog.askopenfilename(filetypes=self.types)
        if path:
            self.var.set(path)
            if self.on_pick:
                self.on_pick(path)

    def get(self) -> str:
        return self.var.get().strip()

    def set(self, value: str) -> None:
        self.var.set(value or "")


class ActivityRow(tk.Frame):
    def __init__(self, parent, on_change):
        super().__init__(parent, bg=PANEL)
        self.on_change = on_change
        self.desc = tk.StringVar()
        ttk.Entry(self, textvariable=self.desc).pack(side="left", fill="x", expand=True)
        self.duration = DurationField(self, self.on_change)
        self.duration.pack(side="left", padx=6)
        ttk.Button(self, text="Quitar", command=self.remove_row).pack(side="left")
        self.desc.trace_add("write", lambda *_: self.on_change())

    def remove_row(self) -> None:
        self.destroy()
        self.on_change()

    def get(self) -> dict | None:
        desc = self.desc.get().strip()
        hours = self.duration.get_value()
        if not desc and self.duration.get_float() == 0:
            return None
        return {"descripcion": desc, "horas": hours}

    def set(self, data: dict) -> None:
        self.desc.set(data.get("descripcion", ""))
        self.duration.set_value(str(data.get("horas", "")))


class ActivityBox(tk.LabelFrame):
    def __init__(self, parent, title: str, on_change):
        super().__init__(parent, text=title, bg=BG, fg=TEXT, font=FONT_SMALL, padx=8, pady=8)
        self.on_change = on_change
        self.rows_frame = tk.Frame(self, bg=PANEL)
        self.rows_frame.pack(fill="both", expand=True)
        self.rows = []
        footer = tk.Frame(self, bg=BG)
        footer.pack(fill="x", pady=(6, 0))
        ttk.Button(footer, text="Agregar actividad", command=self.add_row).pack(side="left")
        self.total_label = tk.Label(footer, text="Total: 0 horas", bg=BG, fg=MUTED, font=FONT_SMALL)
        self.total_label.pack(side="right")
        self.add_row()

    def add_row(self, data: dict | None = None) -> None:
        row = ActivityRow(self.rows_frame, self._changed)
        row.pack(fill="x", pady=3)
        if data:
            row.set(data)
        self.rows.append(row)
        self._changed()

    def _changed(self) -> None:
        self.rows = [row for row in self.rows if row.winfo_exists()]
        total = self.total_hours()
        self.total_label.config(text=f"Total: {format_hours(total) or '0 horas'}")
        self.on_change()

    def total_hours(self) -> float:
        total = 0.0
        for row in self.rows:
            data = row.get()
            if not data:
                continue
            try:
                total += parse_hours(data["horas"])
            except ValueError:
                continue
        return total

    def get_entries(self) -> list[dict]:
        items = []
        for row in self.rows:
            data = row.get()
            if data and data["descripcion"]:
                items.append(data)
        return items

    def set_entries(self, entries: list[dict]) -> None:
        for row in list(self.rows):
            if row.winfo_exists():
                row.destroy()
        self.rows.clear()
        if entries:
            for entry in entries:
                self.add_row(entry)
        else:
            self.add_row()
        self._changed()


class DayBlock(tk.LabelFrame):
    def __init__(self, parent, day_name: str, company_getter, on_change):
        super().__init__(parent, text=day_name, bg=BG, fg=TEXT, font=FONT_BOLD, padx=10, pady=10)
        self.day_name = day_name
        self.company_getter = company_getter
        self.on_change = on_change
        self._ready = False
        top = tk.Frame(self, bg=BG)
        top.pack(fill="x", pady=(0, 8))
        self.date = DateField(top, "Fecha", include_year=True, on_change=self._changed)
        self.date.pack(side="left", padx=(0, 14))
        self.total_label = tk.Label(top, text="Total del dia: 0 horas", bg=BG, fg=PRIMARY, font=FONT_SMALL)
        self.total_label.pack(side="left")
        grid = tk.Frame(self, bg=BG)
        grid.pack(fill="x")
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)
        self.senati = ActivityBox(grid, "SENATI", self._changed)
        self.senati.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.company = ActivityBox(grid, "Empresa", self._changed)
        self.company.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self._ready = True
        self._changed()

    def _changed(self) -> None:
        if not self._ready:
            return
        total = self.senati.total_hours() + self.company.total_hours()
        self.total_label.config(text=f"Total del dia: {format_hours(total) or '0 horas'}")
        self.on_change()

    def set_data(self, data: dict) -> None:
        self.date.set(data.get("fecha", ""))
        self.senati.set_entries(data.get("senati_entries", []))
        self.company.set_entries(data.get("empresa_entries", []))
        self._changed()

    def gui_data(self) -> dict:
        return {
            "dia": self.day_name,
            "fecha": self.date.get(),
            "senati_entries": self.senati.get_entries(),
            "empresa_entries": self.company.get_entries(),
        }

    def report_data(self) -> dict:
        senati_entries = self.senati.get_entries()
        company_entries = self.company.get_entries()
        company_name = self.company_getter() or "Empresa"
        lines = []
        senati_total = sum(parse_hours(item["horas"]) for item in senati_entries if item.get("horas"))
        company_total = sum(parse_hours(item["horas"]) for item in company_entries if item.get("horas"))
        if senati_entries:
            lines.append(f"Senati: {format_hours(senati_total)}")
            lines.extend(f"- {item['descripcion']}" for item in senati_entries)
        if senati_entries and company_entries:
            lines.append("-----------------------------------------------------------")
        if company_entries:
            lines.append(f"Empresa: {company_name}: {format_hours(company_total)}")
            lines.extend(f"- {item['descripcion']}" for item in company_entries)
        return {
            "dia": self.day_name,
            "fecha": self.date.get(),
            "actividades": lines,
            "horas": format_hours(self.senati.total_hours() + self.company.total_hours()),
        }


class WeekEditor(tk.Frame):
    def __init__(self, parent, company_getter):
        super().__init__(parent, bg=BG)
        self.company_getter = company_getter
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", pady=(0, 8))
        self.week = EntryField(header, "Numero de semana", width=10)
        self.week.pack(side="left", padx=(0, 10))
        self.start = DateField(header, "Inicio", include_year=True, on_change=self.update_total)
        self.start.pack(side="left", padx=(0, 10))
        self.end = DateField(header, "Fin", include_year=True, on_change=self.update_total)
        self.end.pack(side="left", padx=(0, 10))
        ttk.Button(header, text="Autocompletar dias", command=self.fill_dates).pack(side="left")
        self.total_label = tk.Label(header, text="Total semanal: 0 horas", bg=BG, fg=PRIMARY, font=FONT_BOLD)
        self.total_label.pack(side="right")
        self.days = {}
        for name in WEEKDAY_NAMES:
            block = DayBlock(self, name, self.company_getter, self.update_total)
            block.pack(fill="x", pady=6)
            self.days[name] = block

    def fill_dates(self) -> None:
        try:
            current = parse_date(self.start.get())
        except ValueError:
            messagebox.showwarning("Fecha invalida", "Usa el formato dd/mm/yyyy en la fecha inicial.")
            return
        for index, name in enumerate(WEEKDAY_NAMES):
            self.days[name].date.set((current + timedelta(days=index)).strftime("%d/%m/%Y"))
        if not self.end.get():
            self.end.set((current + timedelta(days=5)).strftime("%d/%m/%Y"))

    def update_total(self) -> None:
        total = sum(block.senati.total_hours() + block.company.total_hours() for block in self.days.values())
        self.total_label.config(text=f"Total semanal: {format_hours(total) or '0 horas'}")

    def gui_data(self) -> dict:
        return {
            "numero_semana": self.week.get(),
            "desde": self.start.get(),
            "hasta": self.end.get(),
            "dias": [self.days[name].gui_data() for name in WEEKDAY_NAMES],
        }

    def set_data(self, data: dict) -> None:
        self.week.set(data.get("numero_semana", ""))
        self.start.set(data.get("desde", ""))
        self.end.set(data.get("hasta", ""))
        for block in self.days.values():
            block.set_data({})
        for day in data.get("dias", []):
            name = day.get("dia", "").upper()
            if name in self.days:
                self.days[name].set_data(day)
        self.update_total()

    def report_week(self) -> dict:
        total = sum(block.senati.total_hours() + block.company.total_hours() for block in self.days.values())
        return {
            "numero_semana": self.week.get(),
            "dias": [self.days[name].report_data() for name in WEEKDAY_NAMES],
            "horas_totales": format_hours(total),
        }

    def rotation(self, area_text: str) -> dict | None:
        if not self.week.get() and not self.start.get() and not self.end.get():
            return None
        return {
            "area": area_text,
            "desde": normalize_short_date(self.start.get()),
            "hasta": normalize_short_date(self.end.get()),
            "semana": self.week.get(),
        }


class PeaRow(tk.Frame):
    def __init__(self, parent, item: dict):
        super().__init__(parent, bg=PANEL, relief="solid", bd=1, padx=6, pady=6)
        self.level = tk.IntVar(value=0)
        top = tk.Frame(self, bg=PANEL)
        top.pack(fill="x")
        tk.Label(top, text=item["numero"], width=4, bg=PANEL, fg=PRIMARY, font=FONT_BOLD, anchor="w").pack(side="left")
        tk.Label(top, text=item["descripcion"], bg=PANEL, fg=TEXT, font=FONT_SMALL, justify="left", wraplength=640).pack(side="left", fill="x", expand=True)
        opts = tk.Frame(self, bg=PANEL)
        opts.pack(fill="x", pady=(6, 0))
        ttk.Radiobutton(opts, text="0", value=0, variable=self.level).pack(side="left", padx=(0, 8))
        for value, label in [(1, "1"), (2, "2"), (3, "3"), (4, "4")]:
            ttk.Radiobutton(opts, text=label, value=value, variable=self.level).pack(side="left", padx=(0, 8))

    def progress(self) -> dict | None:
        value = self.level.get()
        if value <= 0:
            return None
        return {f"op{value}": "X"}

    def set_progress(self, data: dict) -> None:
        self.level.set(progress_rank(data))
