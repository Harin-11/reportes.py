import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from desktop_gui.config import (
    BG,
    DATA_DIR,
    DEFAULT_TEMPLATE,
    DRAFT_PATH,
    FONT_BOLD,
    FONT_SMALL,
    FONT_TITLE,
    MUTED,
    PANEL,
    PEA_PATH,
    PEA_PROGRESS_PATH,
    PRIMARY,
    PROFILE_PATH,
    TEXT,
    ensure_dirs,
)
from desktop_gui.importers import import_json_payload
from desktop_gui.services import generate_report_files
from desktop_gui.utils import (
    load_json,
    normalize_short_date,
    progress_rank,
    resolve_template_path,
    save_json,
    split_lines,
)
from desktop_gui.widgets import (
    DateField,
    EntryField,
    FileField,
    PeaRow,
    ScrollFrame,
    SegmentedField,
    TextField,
    WeekEditor,
)
from reporting.shared import get_image_pixel_size, join_text_lines

REPORT_FIELD_BINDINGS = (
    ("tarea_significativa", "task", False),
    ("porque_tarea_significativa", "why", True),
    ("descripcion_tarea_significativa", "process", True),
    ("maquinas_usadas", "maquinas", True),
    ("equipos_usados", "equipos", True),
    ("herramientas_usadas", "herramientas", True),
    ("materiales_usados", "materiales", True),
    ("charlas_seguridad", "security", True),
    ("explicacion_actividades_y_medidas_seguras_usadas", "security_notes", True),
    ("resultado_ejecucion", "results", True),
    ("objetivo_logrado_o_no", "objective", True),
    ("recomendaciones_resultado", "recommendations", True),
    ("recomendaciones_monitor", "monitor", False),
    ("diagrama_path", "diagram", False),
    ("firma_estudiante_path", "student_signature", False),
    ("firma_monitor_path", "monitor_signature", False),
)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        ensure_dirs()
        self.title("Generador de Informes FPE")
        self.geometry("1180x860")
        self.configure(bg=BG)
        self.pea_master = load_json(PEA_PATH, [])
        self.status = tk.StringVar(value="Listo.")
        self.output_name = tk.StringVar(value="informe_fpe")
        self.export_pdf = tk.BooleanVar(value=False)
        self.profile_fields = {}
        self.week_editors = []
        self.pea_rows = {}
        self._autosave_job = None
        self._configure_styles()
        self._build()
        self._load_profile()
        self._load_pea_state()
        self._autoload_draft()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TNotebook", background=BG, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(12, 8), font=FONT_SMALL)
        style.map("TNotebook.Tab", background=[("selected", PANEL)], foreground=[("selected", PRIMARY)])
        style.configure("TButton", padding=(10, 6))
        style.configure("TRadiobutton", background=PANEL, foreground=TEXT)

    def _build(self) -> None:
        bar = tk.Frame(self, bg=PRIMARY, padx=16, pady=12)
        bar.pack(fill="x")
        tk.Label(bar, text="Generador de Informes FPE", bg=PRIMARY, fg="white", font=FONT_TITLE).pack(side="left")
        ttk.Button(bar, text="Guardar borrador", command=self.save_draft).pack(side="right")
        ttk.Button(bar, text="Cargar", command=self.load_draft).pack(side="right", padx=(0, 8))
        ttk.Button(bar, text="Generar informe", command=self.generate).pack(side="right", padx=(0, 8))
        tk.Label(self, textvariable=self.status, bg="#E7EEF8", fg=TEXT, anchor="w", padx=12, pady=6, font=FONT_SMALL).pack(fill="x")

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=12, pady=12)

        self.general_tab = ScrollFrame(notebook)
        self.week_tabs = [ScrollFrame(notebook) for _ in range(3)]
        self.pea_tab = ScrollFrame(notebook)
        self.report_tab = ScrollFrame(notebook)

        notebook.add(self.general_tab, text="Datos generales")
        notebook.add(self.week_tabs[0], text="Semana 1")
        notebook.add(self.week_tabs[1], text="Semana 2")
        notebook.add(self.week_tabs[2], text="Semana 3")
        notebook.add(self.pea_tab, text="PEA")
        notebook.add(self.report_tab, text="Informe")

        self._build_general()
        self._build_weeks()
        self._build_pea()
        self._build_report()

    def _build_general(self) -> None:
        frame = self.general_tab.inner
        tk.Label(frame, text="Perfil y configuracion", bg=BG, fg=PRIMARY, font=FONT_TITLE).pack(anchor="w", pady=(0, 10))
        tk.Label(
            frame,
            text="Estos datos se guardan automaticamente y volveran a cargarse en el siguiente reporte.",
            bg=BG,
            fg=MUTED,
            font=FONT_SMALL,
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(0, 8))

        grid = tk.Frame(frame, bg=BG)
        grid.pack(fill="x")
        fields = [
            ("nombre_estudiante", EntryField, {"label": "Nombre del estudiante", "on_change": self.schedule_profile_save}),
            ("id_estudiante", EntryField, {"label": "ID", "on_change": self.schedule_profile_save}),
            ("bloque", EntryField, {"label": "Bloque", "on_change": self.schedule_profile_save}),
            ("carrera", EntryField, {"label": "Carrera", "on_change": self.schedule_profile_save}),
            ("instructor", EntryField, {"label": "Instructor", "on_change": self.schedule_profile_save}),
            ("semestre", EntryField, {"label": "Semestre", "on_change": self.schedule_profile_save}),
            ("fecha_inicio_semestre", DateField, {"label": "Fecha inicio semestre", "include_year": True, "on_change": self.schedule_profile_save}),
            ("fecha_fin_semestre", DateField, {"label": "Fecha fin semestre", "include_year": True, "on_change": self.schedule_profile_save}),
            ("escuela_segmentada", SegmentedField, {"label": "CFP / UCP / Escuela", "segments": ["CFP", "UCP", "Escuela"], "on_change": self.schedule_profile_save}),
            ("nombre_empresa", EntryField, {"label": "Nombre de la empresa", "on_change": self.schedule_profile_save}),
            ("area_empresa_segmentada", SegmentedField, {"label": "Area / seccion / empresa", "segments": ["Area", "Seccion", "Empresa"], "on_change": self.schedule_profile_save}),
        ]
        for index, (key, widget_cls, kwargs) in enumerate(fields):
            cell = tk.Frame(grid, bg=BG)
            cell.grid(row=index // 2, column=index % 2, sticky="ew", padx=8, pady=8)
            widget = widget_cls(cell, **kwargs)
            widget.pack(fill="x")
            self.profile_fields[key] = widget
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        box = tk.LabelFrame(frame, text="Salida", bg=BG, fg=TEXT, font=FONT_BOLD, padx=10, pady=10)
        box.pack(fill="x", pady=(16, 0))
        row = tk.Frame(box, bg=BG)
        row.pack(fill="x")
        tk.Label(row, text="Nombre del archivo", bg=BG, fg=TEXT, font=FONT_SMALL).pack(side="left")
        ttk.Entry(row, textvariable=self.output_name, width=32).pack(side="left", padx=8)
        ttk.Checkbutton(row, text="Exportar PDF", variable=self.export_pdf).pack(side="left", padx=8)
        self.template_field = FileField(box, "Plantilla DOCX", [("Word", "*.docx")])
        self.template_field.pack(fill="x", pady=(10, 0))
        self.template_field.set(str(resolve_template_path(str(DEFAULT_TEMPLATE))))

    def _company_name_for_display(self) -> str:
        name = self.profile_fields["nombre_empresa"].get().strip()
        if name:
            return name
        parts = self.profile_fields["area_empresa_segmentada"].get_parts()
        return parts[2] if len(parts) >= 3 else ""

    def _build_weeks(self) -> None:
        for index, tab in enumerate(self.week_tabs, start=1):
            tk.Label(tab.inner, text="Cada fila tiene descripcion y horas. Las rotaciones saldran de estas fechas.", bg=BG, fg=MUTED, font=FONT_SMALL).pack(anchor="w", pady=(0, 8))
            editor = WeekEditor(tab.inner, self._company_name_for_display)
            editor.pack(fill="x")
            editor.week.set(str(index))
            self.week_editors.append(editor)

    def _build_pea(self) -> None:
        frame = self.pea_tab.inner
        tk.Label(frame, text="Marca solo las tareas realmente avanzadas", bg=BG, fg=PRIMARY, font=FONT_TITLE).pack(anchor="w", pady=(0, 8))
        tk.Label(frame, text="Las tareas sin marca no se incluiran en el informe final.", bg=BG, fg=MUTED, font=FONT_SMALL).pack(anchor="w", pady=(0, 10))
        action_row = tk.Frame(frame, bg=BG)
        action_row.pack(fill="x", pady=(0, 10))
        tk.Label(action_row, text="Cada tarea admite una sola marca y el avance se recuerda entre reportes.", bg=BG, fg=MUTED, font=FONT_SMALL).pack(side="left")
        ttk.Button(action_row, text="Reiniciar progreso guardado", command=self.reset_pea_state).pack(side="right")
        for item in self.pea_master:
            row = PeaRow(frame, item)
            row.pack(fill="x", pady=4)
            self.pea_rows[item["numero"]] = row

    def _build_report(self) -> None:
        frame = self.report_tab.inner
        tk.Label(frame, text="Secciones del informe", bg=BG, fg=PRIMARY, font=FONT_TITLE).pack(anchor="w", pady=(0, 10))
        self.task = EntryField(frame, "Tarea mas significativa")
        self.task.pack(fill="x", pady=6)
        self.why = TextField(frame, "Por que elegiste esta tarea", 4, "Una idea por linea.")
        self.why.pack(fill="x", pady=6)
        self.process = TextField(frame, "Descripcion del proceso", 7, "Explica la secuencia de pasos. Una idea por linea.")
        self.process.pack(fill="x", pady=6)

        grid = tk.Frame(frame, bg=BG)
        grid.pack(fill="x", pady=6)
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)
        self.maquinas = TextField(grid, "Maquinas", 4, "Una por linea.")
        self.maquinas.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=4)
        self.equipos = TextField(grid, "Equipos", 4, "Uno por linea.")
        self.equipos.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=4)
        self.herramientas = TextField(grid, "Herramientas", 4, "Una por linea.")
        self.herramientas.grid(row=1, column=0, sticky="nsew", padx=(0, 6), pady=4)
        self.materiales = TextField(grid, "Materiales", 4, "Uno por linea.")
        self.materiales.grid(row=1, column=1, sticky="nsew", padx=(6, 0), pady=4)

        self.security = TextField(frame, "Charla de 5 minutos / ATS / temas de seguridad", 4, "Una idea por linea.")
        self.security.pack(fill="x", pady=6)
        self.security_notes = TextField(
            frame,
            "Medidas seguras aplicadas y uso del ATS",
            4,
            "Una idea por linea. Ejemplo: EPP, riesgos, controles y pasos seguros.",
        )
        self.security_notes.pack(fill="x", pady=6)
        self.results = TextField(frame, "Resultados", 4)
        self.results.pack(fill="x", pady=6)
        self.objective = TextField(frame, "Objetivo logrado o no", 3)
        self.objective.pack(fill="x", pady=6)
        self.recommendations = TextField(frame, "Recomendaciones", 4, "Una por linea.")
        self.recommendations.pack(fill="x", pady=6)
        self.monitor = TextField(frame, "Observaciones del monitor", 3)
        self.monitor.pack(fill="x", pady=6)
        tk.Label(
            frame,
            text="Para evitar que se descuadre la tabla, usa de preferencia una imagen horizontal o un recorte del esquema.",
            bg=BG,
            fg=MUTED,
            font=FONT_SMALL,
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(4, 0))
        self.diagram = FileField(
            frame,
            "Imagen del esquema, dibujo o diagrama",
            [("Imagenes", "*.png;*.jpg;*.jpeg")],
            on_pick=self._on_diagram_selected,
        )
        self.diagram.pack(fill="x", pady=6)
        self.student_signature = FileField(frame, "Firma del estudiante", [("Imagenes", "*.png;*.jpg;*.jpeg")])
        self.student_signature.pack(fill="x", pady=6)
        self.monitor_signature = FileField(frame, "Firma del monitor", [("Imagenes", "*.png;*.jpg;*.jpeg")])
        self.monitor_signature.pack(fill="x", pady=6)

    def schedule_profile_save(self) -> None:
        if self._autosave_job is not None:
            self.after_cancel(self._autosave_job)
        self._autosave_job = self.after(400, self._autosave_profile)

    def _autosave_profile(self) -> None:
        self._autosave_job = None
        try:
            self._save_profile()
            self.status.set("Datos generales guardados automaticamente.")
        except Exception:
            pass

    def _profile_data(self) -> dict:
        raw = {key: field.get() for key, field in self.profile_fields.items()}
        school_parts = self.profile_fields["escuela_segmentada"].get_parts()
        area_parts = self.profile_fields["area_empresa_segmentada"].get_parts()
        company_name = raw.get("nombre_empresa", "").strip() or (area_parts[2] if len(area_parts) >= 3 else "")
        area_value = " / ".join(part for part in area_parts if part) or company_name
        raw["nombre_empresa"] = company_name
        raw["area_empresa"] = area_value
        raw["escuela"] = " / ".join(part for part in school_parts if part)
        return raw

    def _load_profile(self) -> None:
        self._apply_profile(load_json(PROFILE_PATH, {}))

    def _save_profile(self) -> None:
        save_json(PROFILE_PATH, self._profile_data())

    def _apply_profile(self, profile: dict) -> None:
        for key, field in self.profile_fields.items():
            field.set(profile.get(key, ""))
        self.profile_fields["escuela_segmentada"].set(profile.get("escuela_segmentada") or profile.get("escuela", ""))
        self.profile_fields["area_empresa_segmentada"].set(profile.get("area_empresa_segmentada") or profile.get("area_empresa", ""))

    def _report_section_data(self, *, split_multiline: bool) -> dict:
        report = {}
        for key, attribute, multiline in REPORT_FIELD_BINDINGS:
            value = getattr(self, attribute).get()
            report[key] = split_lines(value) if split_multiline and multiline else value
        return report

    def _apply_report_section(self, report: dict) -> None:
        for key, attribute, _ in REPORT_FIELD_BINDINGS:
            getattr(self, attribute).set(join_text_lines(report.get(key, "")))

    def _on_diagram_selected(self, path_text: str) -> None:
        size = get_image_pixel_size(Path(path_text))
        if not size:
            return

        width_px, height_px = size
        if width_px <= 0 or height_px <= 0:
            return

        ratio = height_px / width_px
        if ratio > 1.1:
            messagebox.showwarning(
                "Imagen vertical",
                (
                    "La imagen del esquema es bastante vertical.\n\n"
                    "Se ajustará automáticamente en el DOCX, pero conviene usar un recorte más horizontal "
                    "para evitar que la tabla se vea apretada."
                ),
            )

    def _pea_progress(self) -> dict:
        progress = {}
        for number, row in self.pea_rows.items():
            value = row.progress()
            if value:
                progress[number] = value
        return progress

    def _load_pea_state(self) -> None:
        state = load_json(PEA_PROGRESS_PATH, {})
        for number, progress in state.items():
            if number in self.pea_rows:
                self.pea_rows[number].set_progress(progress)

    def _save_pea_state(self) -> None:
        existing = load_json(PEA_PROGRESS_PATH, {})
        current = self._pea_progress()
        merged = {}
        for item in self.pea_master:
            number = item["numero"]
            existing_progress = existing.get(number, {})
            current_progress = current.get(number, {})
            merged[number] = current_progress if progress_rank(current_progress) >= progress_rank(existing_progress) else existing_progress
        save_json(PEA_PROGRESS_PATH, {k: v for k, v in merged.items() if progress_rank(v) > 0})

    def reset_pea_state(self) -> None:
        if not messagebox.askyesno("Reiniciar PEA", "Esto borrara el avance acumulado guardado del PEA. ¿Continuar?"):
            return
        save_json(PEA_PROGRESS_PATH, {})
        for row in self.pea_rows.values():
            row.set_progress({})
        self.status.set("Progreso persistente del PEA reiniciado.")

    def _draft_data(self) -> dict:
        return {
            "profile": self._profile_data(),
            "output_name": self.output_name.get().strip(),
            "export_pdf": self.export_pdf.get(),
            "template_path": str(resolve_template_path(self.template_field.get())),
            "weeks": [editor.gui_data() for editor in self.week_editors],
            "pea_avances": self._pea_progress(),
            "report": self._report_section_data(split_multiline=False),
        }

    def _apply_draft(self, draft: dict) -> None:
        persistent_state = load_json(PEA_PROGRESS_PATH, {})
        self._apply_profile(draft.get("profile", {}))
        self.output_name.set(draft.get("output_name", "informe_fpe"))
        self.export_pdf.set(bool(draft.get("export_pdf", False)))
        self.template_field.set(str(resolve_template_path(draft.get("template_path", str(DEFAULT_TEMPLATE)))))

        weeks = draft.get("weeks", [])
        for index, editor in enumerate(self.week_editors):
            editor.set_data(weeks[index] if index < len(weeks) else {})

        draft_progress = draft.get("pea_avances", {})
        for number, row in self.pea_rows.items():
            persisted = persistent_state.get(number, {})
            draft_value = draft_progress.get(number, {})
            chosen = draft_value if progress_rank(draft_value) >= progress_rank(persisted) else persisted
            row.set_progress(chosen)

        self._apply_report_section(draft.get("report", {}))

    def _autoload_draft(self) -> None:
        if DRAFT_PATH.exists():
            try:
                self._apply_draft(load_json(DRAFT_PATH, {}))
                self.status.set("Borrador cargado automaticamente.")
            except Exception:
                self.status.set("No se pudo cargar el borrador automatico.")

    def save_draft(self) -> None:
        self._save_profile()
        self._save_pea_state()
        save_json(DRAFT_PATH, self._draft_data())
        self.status.set(f"Borrador guardado en {DRAFT_PATH.name}.")
        messagebox.showinfo("Borrador guardado", f"Se guardo el borrador en:\n{DRAFT_PATH}")

    def load_draft(self) -> None:
        path = filedialog.askopenfilename(initialdir=str(DATA_DIR), filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            payload = load_json(Path(path), {})
            imported = import_json_payload(payload, template_path=str(resolve_template_path(self.template_field.get())))
            self._apply_draft(imported.draft)
            self.status.set(f"{imported.message} Archivo: {Path(path).name}.")
        except Exception as exc:
            self.status.set("No se pudo cargar el archivo.")
            messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{exc}")

    def _build_report_json(self) -> dict:
        profile = self._profile_data()
        area_text = profile.get("area_empresa") or profile.get("nombre_empresa") or ""
        weeks = []
        rotations = []
        for editor in self.week_editors:
            week = editor.report_week()
            has_data = any(day["fecha"] or day["actividades"] or day["horas"] for day in week["dias"])
            if not has_data:
                continue
            weeks.append(week)
            rotation = editor.rotation(area_text)
            if rotation:
                rotations.append(rotation)
        return {
            "nombre_estudiante": profile.get("nombre_estudiante", ""),
            "id_estudiante": profile.get("id_estudiante", ""),
            "bloque": profile.get("bloque", ""),
            "carrera": profile.get("carrera", ""),
            "instructor": profile.get("instructor", ""),
            "semestre": profile.get("semestre", ""),
            "fecha_inicio_semestre": normalize_short_date(profile.get("fecha_inicio_semestre", "")),
            "fecha_fin_semestre": normalize_short_date(profile.get("fecha_fin_semestre", "")),
            "escuela": profile.get("escuela", ""),
            "nombre_empresa": profile.get("nombre_empresa", ""),
            "rotaciones": rotations,
            "pea_avances": self._pea_progress(),
            "semanas": weeks,
            **self._report_section_data(split_multiline=True),
        }

    def _validate(self, report_data: dict) -> None:
        template_path = resolve_template_path(self.template_field.get())
        self.template_field.set(str(template_path))
        if not template_path.exists():
            raise ValueError("La plantilla DOCX no existe.")
        if not report_data["nombre_estudiante"]:
            raise ValueError("Completa el nombre del estudiante.")
        if not report_data["semanas"]:
            raise ValueError("Completa al menos una semana con datos.")
        if not report_data["pea_avances"]:
            raise ValueError("Marca al menos una tarea del PEA que hayas avanzado.")

    def generate(self) -> None:
        try:
            self._set_status_busy("Preparando datos...")
            self._save_profile()
            self._save_pea_state()
            save_json(DRAFT_PATH, self._draft_data())
            report_data = self._build_report_json()
            self._validate(report_data)
            output_name = self.output_name.get().strip() or "informe_fpe"
            pdf_path = None

            if self.export_pdf.get():
                try:
                    result = generate_report_files(
                        template_path=resolve_template_path(self.template_field.get()),
                        report_data=report_data,
                        pea_master=self.pea_master,
                        output_name=output_name,
                        export_pdf=True,
                        status_callback=self._set_status_busy,
                    )
                    docx_path = result.docx_path
                    pdf_path = result.pdf_path
                except Exception as exc:
                    result = generate_report_files(
                        template_path=resolve_template_path(self.template_field.get()),
                        report_data=report_data,
                        pea_master=self.pea_master,
                        output_name=output_name,
                        export_pdf=False,
                        status_callback=self._set_status_busy,
                    )
                    docx_path = result.docx_path
                    messagebox.showwarning("PDF no generado", f"Se genero el DOCX, pero el PDF fallo:\n{exc}")
            else:
                result = generate_report_files(
                    template_path=resolve_template_path(self.template_field.get()),
                    report_data=report_data,
                    pea_master=self.pea_master,
                    output_name=output_name,
                    export_pdf=False,
                    status_callback=self._set_status_busy,
                )
                docx_path = result.docx_path

            self.status.set(f"Informe generado: {docx_path.name}")
            text = f"Informe generado correctamente:\n\n{docx_path}"
            if pdf_path and pdf_path.exists():
                text += f"\n{pdf_path}"
            messagebox.showinfo("Proceso completado", text)
            try:
                os.startfile(docx_path)
            except Exception:
                pass
        except Exception as exc:
            self.status.set("Ocurrio un error.")
            messagebox.showerror("Error", str(exc))

    def _set_status_busy(self, text: str) -> None:
        self.status.set(text)
        self.update_idletasks()

    def on_close(self) -> None:
        try:
            self._save_profile()
            self._save_pea_state()
            save_json(DRAFT_PATH, self._draft_data())
        except Exception:
            pass
        self.destroy()


def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()
