"""AfficheFlyCase Windows desktop application."""

from __future__ import annotations

import os
import re
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from afficheflycase.pdf_labels import LabelOverflowError, generate_pdf
from afficheflycase.workbook import FlyCase, WorkbookFormatError, load_flycases


class AfficheFlyCaseApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("AfficheFlyCase — étiquettes de fly cases")
        self.root.geometry("1040x720")
        self.root.minsize(860, 600)
        self.cases: list[FlyCase] = []
        self.case_by_item: dict[str, FlyCase] = {}
        self.workbook_path: Path | None = None
        self._build_ui()

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
        self.show_var = tk.StringVar()
        self.date_var = tk.StringVar()
        self.venue_var = tk.StringVar()
        for column, (label, variable) in enumerate(
            (("Nom du show", self.show_var), ("Dates / période", self.date_var), ("Lieu / salle", self.venue_var))
        ):
            field = ttk.Frame(production)
            field.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
            ttk.Label(field, text=label).pack(anchor="w")
            ttk.Entry(field, textvariable=variable).pack(fill="x", pady=(4, 0))
            production.columnconfigure(column, weight=1)

        content = ttk.Panedwindow(main, orient="horizontal")
        content.pack(fill="both", expand=True)
        list_frame = ttk.LabelFrame(content, text="Fly cases — tout sélectionner ou Ctrl+clic", padding=8)
        preview_frame = ttk.LabelFrame(content, text="Aperçu des informations", padding=8)
        content.add(list_frame, weight=1)
        content.add(preview_frame, weight=2)

        self.case_tree = ttk.Treeview(
            list_frame,
            columns=("type", "tip", "stack", "items"),
            show="tree headings",
            selectmode="extended",
        )
        self.case_tree.heading("#0", text="ID")
        self.case_tree.column("#0", width=90, stretch=False)
        for name, label, width in (
            ("type", "Type", 100),
            ("tip", "Tip", 70),
            ("stack", "Gerbable", 90),
            ("items", "Éléments", 70),
        ):
            self.case_tree.heading(name, text=label)
            self.case_tree.column(name, width=width, anchor="center")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.case_tree.yview)
        self.case_tree.configure(yscrollcommand=scrollbar.set)
        self.case_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.case_tree.bind("<<TreeviewSelect>>", self._show_preview)

        self.preview = tk.Text(preview_frame, wrap="word", height=18, state="disabled", font=("Segoe UI", 10))
        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview.yview)
        self.preview.configure(yscrollcommand=preview_scroll.set)
        self.preview.pack(side="left", fill="both", expand=True)
        preview_scroll.pack(side="right", fill="y")

        bottom = ttk.Frame(main)
        bottom.pack(fill="x", pady=(12, 0))
        self.status_var = tk.StringVar(value="Sélectionnez le classeur Excel à traiter.")
        ttk.Label(bottom, textvariable=self.status_var, wraplength=700).pack(side="left", fill="x", expand=True)
        self.generate_button = ttk.Button(
            bottom, text="Générer le PDF…", command=self.create_pdf, state="disabled"
        )
        self.generate_button.pack(side="right")

    def choose_workbook(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="Sélectionner le listing Excel",
            filetypes=[("Classeurs Excel", "*.xlsx *.xlsm"), ("Tous les fichiers", "*.*")],
        )
        if not selected:
            return
        self.workbook_path = Path(selected)
        self.path_label.configure(text=str(self.workbook_path))
        try:
            loaded = load_flycases(self.workbook_path)
        except WorkbookFormatError as error:
            self.cases.clear()
            self.case_by_item.clear()
            self.case_tree.delete(*self.case_tree.get_children())
            self.generate_button.configure(state="disabled")
            self.status_var.set("Import refusé : le fichier ne correspond pas au modèle attendu.")
            self._set_preview(str(error))
            messagebox.showerror("Fichier Excel non conforme", str(error), parent=self.root)
            return
        except Exception as error:
            self.cases.clear()
            self.case_by_item.clear()
            self.case_tree.delete(*self.case_tree.get_children())
            self.generate_button.configure(state="disabled")
            self.status_var.set("Erreur pendant la lecture du fichier Excel.")
            self._set_preview(str(error))
            messagebox.showerror(
                "Erreur de lecture", f"Une erreur inattendue empêche de lire le classeur :\n{error}", parent=self.root
            )
            return

        self.cases = loaded
        self.case_tree.delete(*self.case_tree.get_children())
        self.case_by_item.clear()
        for case in self.cases:
            warning = " ⚠" if not case.tip or not case.stackable else ""
            item_id = self.case_tree.insert(
                "",
                "end",
                text=case.identifier + warning,
                values=(case.case_type or "—", case.tip_label, case.stackable_label, len(case.materials)),
            )
            self.case_by_item[item_id] = case
        self.case_tree.selection_set(self.case_tree.get_children())
        self.generate_button.configure(state="normal")
        self.status_var.set(
            f"{len(self.cases)} fly case(s) importé(s) depuis les deux premiers onglets. "
            "Les lignes marquées ⚠ ont un statut de manutention à vérifier."
        )
        self._show_preview()

    def _show_preview(self, _event=None) -> None:
        selected = self.case_tree.selection()
        if not selected:
            self._set_preview("Sélectionnez un fly case dans la liste.")
            return
        case = self.case_by_item[selected[0]]
        lines = [
            f"FLY CASE {case.identifier}",
            f"Type : {case.case_type or 'Non renseigné'}",
            f"Couleur : {case.color or 'Non renseignée'}",
            f"Dimensions : {case.dimensions}",
            f"Empattement : {case.footprint or 'Non renseigné'}",
            f"Cubage : {case.volume or 'Non renseigné'}",
            f"Tip : {case.tip_label}",
            f"Gerbage : {case.stackable_label}",
            f"Commentaire : {case.comment or 'Aucun'}",
            "",
            "CONTENU / REMARQUES",
        ]
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
        self._set_preview("\n".join(lines))

    def _set_preview(self, text: str) -> None:
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", text)
        self.preview.configure(state="disabled")

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
            count = generate_pdf(destination, selected_cases, show, dates, venue)
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
