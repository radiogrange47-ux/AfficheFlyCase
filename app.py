"""AfficheFlyCase Windows desktop application."""

from __future__ import annotations

import os
import re
import sys
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from afficheflycase.pdf_labels import LabelOverflowError, generate_pdf
from afficheflycase.settings import (
    AppSettings,
    SettingsError,
    load_settings,
    save_settings,
    settings_directory,
    store_logo,
)
from afficheflycase.workbook import FlyCase, WorkbookFormatError, WorkbookImportResult, load_flycases


class AfficheFlyCaseApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("AfficheFlyCase — étiquettes de fly cases")
        self.root.geometry("1040x720")
        self.root.minsize(860, 600)
        self.cases: list[FlyCase] = []
        self.case_by_item: dict[str, FlyCase] = {}
        self.import_result: WorkbookImportResult | None = None
        self.workbook_path: Path | None = None
        self.settings_error: str | None = None
        try:
            self.settings = load_settings()
        except SettingsError as error:
            self.settings = AppSettings()
            self.settings_error = str(error)
        self._settings_save_job: str | None = None
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self.show_var.trace_add("write", self._schedule_settings_save)
        self.date_var.trace_add("write", self._schedule_settings_save)
        self.venue_var.trace_add("write", self._schedule_settings_save)
        self.root.after(100, self._startup)

    @staticmethod
    def _application_directory() -> Path:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parent

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=16)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="AfficheFlyCase", font=("Segoe UI", 20, "bold")).pack(anchor="w")
        ttk.Label(
            main,
            text="Importez le listing matériel et créez une étiquette A4 par fly case.",
        ).pack(anchor="w", pady=(0, 12))

        file_row = ttk.Frame(main)
        file_row.pack(fill="x", pady=(0, 12))
        ttk.Button(file_row, text="Choisir le fichier Excel…", command=self.choose_workbook).pack(side="left")
        self.path_label = ttk.Label(file_row, text="Aucun fichier sélectionné", anchor="w")
        self.path_label.pack(side="left", fill="x", expand=True, padx=10)

        production = ttk.LabelFrame(main, text="Informations de la production", padding=10)
        production.pack(fill="x", pady=(0, 12))
        self.show_var = tk.StringVar(value=self.settings.show)
        self.date_var = tk.StringVar(value=self.settings.dates)
        self.venue_var = tk.StringVar(value=self.settings.venue)
        for column, (label, variable) in enumerate(
            (("Nom du show", self.show_var), ("Dates / période", self.date_var), ("Lieu / salle", self.venue_var))
        ):
            field = ttk.Frame(production)
            field.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
            ttk.Label(field, text=label).pack(anchor="w")
            ttk.Entry(field, textvariable=variable).pack(fill="x", pady=(4, 0))
            production.columnconfigure(column, weight=1)

        logo_row = ttk.Frame(main)
        logo_row.pack(fill="x", pady=(0, 12))
        ttk.Label(logo_row, text="Logo de l’étiquette :").pack(side="left")
        ttk.Button(logo_row, text="Choisir un logo…", command=self.choose_logo).pack(side="left", padx=(8, 0))
        ttk.Button(logo_row, text="Supprimer", command=self.remove_logo).pack(side="left", padx=(6, 0))
        self.logo_path_var = tk.StringVar(value=self.settings.logo_path)
        ttk.Label(logo_row, textvariable=self.logo_path_var, anchor="w").pack(
            side="left", fill="x", expand=True, padx=10
        )
        if self.settings.logo_path and not Path(self.settings.logo_path).is_file():
            self.settings_error = (
                f"Le logo mémorisé est introuvable : {self.settings.logo_path}. "
                "Choisissez un nouveau fichier logo."
            )

        content = ttk.Panedwindow(main, orient="horizontal")
        content.pack(fill="both", expand=True)
        list_frame = ttk.LabelFrame(content, text="Fly cases — tout sélectionner ou Ctrl+clic", padding=8)
        preview_frame = ttk.LabelFrame(content, text="Aperçu des informations", padding=8)
        content.add(list_frame, weight=1)
        content.add(preview_frame, weight=2)

        self.case_tree = ttk.Treeview(
            list_frame,
            columns=("items",),
            show="tree headings",
            selectmode="extended",
        )
        self.case_tree.heading("#0", text="ID")
        self.case_tree.column("#0", width=90, stretch=False)
        self.case_tree.heading("items", text="Éléments")
        self.case_tree.column("items", width=90, anchor="center")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.case_tree.yview)
        self.case_tree.configure(yscrollcommand=scrollbar.set)
        self.case_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.case_tree.bind("<<TreeviewSelect>>", self._show_preview)

        self.preview_tabs = ttk.Notebook(preview_frame)
        self.preview_tabs.pack(fill="both", expand=True)
        case_preview_frame = ttk.Frame(self.preview_tabs)
        errors_frame = ttk.Frame(self.preview_tabs)
        self.preview_tabs.add(case_preview_frame, text="Aperçu fly case")
        self.preview_tabs.add(errors_frame, text="Erreurs d’affectation (0)")
        self.preview = self._create_text_view(case_preview_frame)
        self.errors_preview = self._create_text_view(errors_frame)

        bottom = ttk.Frame(main)
        bottom.pack(fill="x", pady=(12, 0))
        self.status_var = tk.StringVar(value="Sélectionnez le classeur Excel à traiter.")
        ttk.Label(bottom, textvariable=self.status_var, wraplength=700).pack(side="left", fill="x", expand=True)
        self.generate_button = ttk.Button(
            bottom, text="Générer le PDF…", command=self.create_pdf, state="disabled"
        )
        self.generate_button.pack(side="right")

    def _startup(self) -> None:
        if self.settings_error:
            messagebox.showwarning("Préférences", self.settings_error, parent=self.root)
        self.choose_workbook()

    def _schedule_settings_save(self, *_args) -> None:
        if self._settings_save_job:
            self.root.after_cancel(self._settings_save_job)
        self._settings_save_job = self.root.after(600, self._save_settings)

    def _save_settings(self) -> None:
        self._settings_save_job = None
        self.settings.show = self.show_var.get()
        self.settings.dates = self.date_var.get()
        self.settings.venue = self.venue_var.get()
        self.settings.logo_path = self.logo_path_var.get()
        try:
            save_settings(self.settings)
        except SettingsError as error:
            messagebox.showerror("Préférences non enregistrées", str(error), parent=self.root)

    def _close(self) -> None:
        if self._settings_save_job:
            self.root.after_cancel(self._settings_save_job)
            self._settings_save_job = None
        self._save_settings()
        self.root.destroy()

    def choose_logo(self) -> None:
        current_path = Path(self.logo_path_var.get()) if self.logo_path_var.get() else None
        if current_path and current_path.exists():
            initial_directory = current_path.parent
        else:
            preferences_directory = settings_directory()
            initial_directory = (
                preferences_directory if preferences_directory.is_dir() else self._application_directory()
            )
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="Choisir le logo des étiquettes",
            initialdir=str(initial_directory),
            filetypes=[("Images PNG ou JPEG", "*.png *.jpg *.jpeg"), ("Tous les fichiers", "*.*")],
        )
        if not selected:
            return
        try:
            stored_path = store_logo(selected)
            self.logo_path_var.set(str(stored_path))
            self._save_settings()
        except SettingsError as error:
            messagebox.showerror("Logo non enregistré", str(error), parent=self.root)

    def remove_logo(self) -> None:
        logo_path = self.logo_path_var.get()
        if not logo_path:
            return
        self.logo_path_var.set("")
        self._save_settings()
        try:
            path = Path(logo_path)
            if path.parent.resolve() == settings_directory().resolve() and path.name.lower().startswith("logo."):
                path.unlink(missing_ok=True)
        except OSError as error:
            messagebox.showerror("Logo non supprimé", f"Le logo n’a pas pu être supprimé : {error}", parent=self.root)

    @staticmethod
    def _create_text_view(parent: ttk.Frame) -> tk.Text:
        text = tk.Text(parent, wrap="word", height=18, state="disabled", font=("Segoe UI", 10))
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        return text

    def choose_workbook(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="Sélectionner le listing Excel",
            initialdir=str(self._application_directory()),
            filetypes=[("Classeurs Excel", "*.xlsx *.xlsm"), ("Tous les fichiers", "*.*")],
        )
        if not selected:
            return
        self.workbook_path = Path(selected)
        self.path_label.configure(text=str(self.workbook_path))
        try:
            result = load_flycases(self.workbook_path)
        except WorkbookFormatError as error:
            self.cases.clear()
            self.case_by_item.clear()
            self.import_result = None
            self.case_tree.delete(*self.case_tree.get_children())
            self.generate_button.configure(state="disabled")
            self.preview_tabs.tab(1, text="Erreurs d’affectation (0)")
            self._set_text(self.errors_preview, "Aucun matériel importé.")
            self.status_var.set("Import refusé : le fichier ne correspond pas au modèle attendu.")
            self._set_preview(str(error))
            messagebox.showerror("Fichier Excel non conforme", str(error), parent=self.root)
            return
        except Exception as error:
            self.cases.clear()
            self.case_by_item.clear()
            self.import_result = None
            self.case_tree.delete(*self.case_tree.get_children())
            self.generate_button.configure(state="disabled")
            self.preview_tabs.tab(1, text="Erreurs d’affectation (0)")
            self._set_text(self.errors_preview, "Aucun matériel importé.")
            self.status_var.set("Erreur pendant la lecture du fichier Excel.")
            self._set_preview(str(error))
            messagebox.showerror(
                "Erreur de lecture", f"Une erreur inattendue empêche de lire le classeur :\n{error}", parent=self.root
            )
            return

        self.import_result = result
        self.cases = result.cases
        self.case_tree.delete(*self.case_tree.get_children())
        self.case_by_item.clear()
        for case in self.cases:
            item_id = self.case_tree.insert(
                "",
                "end",
                text=case.identifier,
                values=(len(case.materials),),
            )
            self.case_by_item[item_id] = case
        self.case_tree.selection_set(self.case_tree.get_children())
        self.generate_button.configure(state="normal")
        unassigned = result.unassigned_materials
        self.preview_tabs.tab(1, text=f"Erreurs d’affectation ({len(unassigned)})")
        error_lines = [
            f"Ligne {item.material.source_row} — {item.description}\n  {item.reason}"
            for item in unassigned
        ]
        self._set_text(self.errors_preview, "\n\n".join(error_lines) or "Aucun matériel non affecté.")
        status = f"{len(self.cases)} fly case(s) importé(s) depuis les deux premiers onglets."
        if unassigned:
            status += f" {len(unassigned)} matériel(s) non affecté(s) : consulter l’onglet des erreurs."
        self.status_var.set(status)
        self.preview_tabs.select(1 if unassigned else 0)
        self._show_preview()
        if unassigned:
            messagebox.showwarning(
                "Matériel non affecté",
                f"{len(unassigned)} matériel(s) ne sont associé(s) à aucun fly case. "
                "Chaque ligne est détaillée dans l’onglet « Erreurs d’affectation ». "
                "Ce matériel sera ignoré et la génération des étiquettes reste possible.",
                parent=self.root,
            )

    def _show_preview(self, _event=None) -> None:
        selected = self.case_tree.selection()
        if not selected:
            self._set_preview("Sélectionnez un fly case dans la liste.")
            return
        case = self.case_by_item[selected[0]]
        lines = [
            f"FLY CASE {case.identifier}",
            f"Dimensions : {case.dimensions}",
        ]
        if case.footprint:
            lines.append(case.footprint)
        if case.volume:
            lines.append(case.volume)
        lines.extend(("", "CONTENU"))
        for item in case.materials:
            line = f"• {item.element or 'Élément non renseigné'}"
            if item.quantity:
                line += f" — Qté : {item.quantity}"
            if item.spare and item.spare != "-":
                line += f" — Spare : {item.spare}"
            if item.position:
                line += f" — Position : {item.position}"
            lines.append(line)
        if not case.materials:
            lines.append("Aucun élément associé dans le listing matériel.")
        if case.comment:
            lines.extend(("", "COMMENTAIRES", case.comment))
        self._set_preview("\n".join(lines))

    def _set_preview(self, text: str) -> None:
        self._set_text(self.preview, text)

    @staticmethod
    def _set_text(widget: tk.Text, text: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")

    def create_pdf(self) -> None:
        selected_items = self.case_tree.selection()
        selected_cases = [self.case_by_item[item] for item in selected_items]
        if not selected_cases:
            messagebox.showwarning("Aucun fly case", "Sélectionnez au moins un fly case.", parent=self.root)
            return
        show = self.show_var.get().strip()
        dates = self.date_var.get().strip()
        venue = self.venue_var.get().strip()
        missing = [name for name, value in (("nom du show", show), ("dates", dates), ("lieu", venue)) if not value]
        if missing:
            messagebox.showwarning(
                "Informations manquantes",
                "Veuillez renseigner : " + ", ".join(missing) + ".",
                parent=self.root,
            )
            return

        default_name = re.sub(r'[<>:"/\\|?*]+', "_", show).strip(" .") or "Etiquettes_FlyCase"
        destination = filedialog.asksaveasfilename(
            parent=self.root,
            title="Enregistrer les étiquettes PDF",
            initialdir=str(self.workbook_path.parent if self.workbook_path else Path.home()),
            initialfile=f"{default_name}_{date.today():%Y%m%d}.pdf",
            defaultextension=".pdf",
            filetypes=[("Document PDF", "*.pdf")],
            confirmoverwrite=True,
        )
        if not destination:
            return
        try:
            count = generate_pdf(
                destination,
                selected_cases,
                show,
                dates,
                venue,
                logo_path=self.logo_path_var.get() or None,
            )
        except (LabelOverflowError, ValueError, OSError) as error:
            messagebox.showerror("PDF non généré", str(error), parent=self.root)
            self.status_var.set("La génération a échoué ; aucun PDF de réussite n’a été annoncé.")
            return
        except Exception as error:
            messagebox.showerror(
                "Erreur de génération", f"Impossible de créer le PDF :\n{error}", parent=self.root
            )
            self.status_var.set("La génération du PDF a échoué.")
            return
        self.status_var.set(f"{count} page(s) A4 créées : {destination}")
        if messagebox.askyesno(
            "PDF créé", f"{count} étiquette(s) enregistrée(s).\n\nOuvrir le PDF ?", parent=self.root
        ):
            try:
                os.startfile(destination)
            except OSError as error:
                messagebox.showerror(
                    "Ouverture impossible",
                    f"Le PDF a été créé, mais Windows ne peut pas l’ouvrir :\n{error}",
                    parent=self.root,
                )


def main() -> None:
    root = tk.Tk()
    AfficheFlyCaseApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
