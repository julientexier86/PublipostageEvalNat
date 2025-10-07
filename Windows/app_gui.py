# -*- coding: utf-8 -*-
"""
GUI légère (Tkinter) pour piloter le pipeline EvalNat.
- 5 onglets : Contexte • Découpage • Récupération mails parents • Message aux parents • Publipostage
- Capture le log stdout du pipeline dans l’interface
- Multiplateforme (macOS/Windows/Linux)
"""
import sys, os, subprocess, threading, shlex, shutil, re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- Frozen resources (PyInstaller) ----------------------------------------
def resource_path(relative_path=""):
    """
    Retourne un chemin absolu vers une ressource embarquée.
    - En mode 'frozen' (PyInstaller), les données sont extraites dans sys._MEIPASS.
    - En mode dev, on retourne le chemin relatif depuis le dossier du script.
    """
    if hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS  # type: ignore[attr-defined]
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)


FROZEN_BASE = resource_path("")  # dossier où PyInstaller extrait l'app

# --- Pipeline binary resolver (PyInstaller bundled) -------------------------
def pipeline_binary() -> str | None:
    """
    Retourne le chemin absolu du binaire pipeline embarqué.
    Accepte les deux noms historiques :
      - evalnat-pipeline
      - pipeline_evalnat
    Recherche dans les emplacements usuels d'une app PyInstaller :
      - Contents/MacOS/<nom>
      - Contents/Frameworks/<nom>
      - Ressources extraites (MEIPASS)
    Et en mode dev :
      - ./dist/<nom>
      - ../dist/<nom>
      - dossier du projet (à côté de app_gui.py)
    Fallback : propose une sélection manuelle si introuvable.
    """
    from pathlib import Path

    # Dossier de l'exécutable courant (dans la .app : .../Contents/MacOS)
    exe_dir = Path(sys.executable).resolve().parent
    # Dossier source (dev) où se trouve app_gui.py
    src_dir = Path(__file__).resolve().parent

    names = [
        "evalnat-pipeline",
        "pipeline_evalnat",
        # variantes Windows éventuelles (au cas où)
        "evalnat-pipeline.exe",
        "pipeline_evalnat.exe",
    ]

    candidates: list[Path] = []
    for nm in names:
        # Emplacements typiques dans la .app
        candidates += [
            exe_dir / nm,                                        # Contents/MacOS/<nm>
            (exe_dir / f"../Frameworks/{nm}").resolve(),         # Contents/Frameworks/<nm>
            Path(resource_path(nm)),                             # MEIPASS/<nm>
        ]
        # Emplacements "dev"
        candidates += [
            (src_dir / "dist" / nm),                             # ./dist/<nm>
            (src_dir / ".." / "dist" / nm).resolve(),            # ../dist/<nm>
            (src_dir / nm),                                      # à côté de app_gui.py
        ]

    for c in candidates:
        try:
            if c.exists() and c.is_file():
                try:
                    if os.name != "nt":
                        os.chmod(c, 0o755)
                except Exception:
                    pass
                return str(c)
        except Exception:
            continue

    # Si rien trouvé : proposer sélection manuelle + tracer où on a cherché (verbose)
    try:
        from tkinter import messagebox, filedialog
        messagebox.showwarning(
            "Binaire introuvable",
            "Le binaire 'evalnat-pipeline' ou 'pipeline_evalnat' n'a pas été trouvé dans l’application.\n"
            "Sélectionnez-le manuellement (par ex. dist/evalnat-pipeline ou dist/pipeline_evalnat)."
        )
        fp = filedialog.askopenfilename(
            title="Choisir le binaire (evalnat-pipeline / pipeline_evalnat)",
            initialdir=str((src_dir / "dist").resolve()),
            filetypes=[("Exécutable", "*")],
        )
        if fp:
            return fp
    except Exception:
        pass

    # Petit log en console pour aider au debug si lancé en dev
    try:
        print("[DEBUG] Binaire pipeline introuvable. Chemins testés :")
        for c in candidates:
            print("  -", c)
    except Exception:
        pass

    return None

APP_TITLE = "EvalNat – Publipostage"
DEFAULT_YEAR = "2025-2026"


# --- Helpers UI --------------------------------------------------------------
def choose_file(entry: tk.Entry, title="Choisir un fichier", types=(("Tous", "*.*"),)):
    path = filedialog.askopenfilename(title=title, filetypes=types)
    if path:
        entry.delete(0, tk.END); entry.insert(0, path)

def choose_files(listbox: tk.Listbox, title="Choisir des fichiers", types=(("CSV", "*.csv"), ("Tous", "*.*"))):
    paths = filedialog.askopenfilenames(title=title, filetypes=types)
    for p in paths:
        listbox.insert(tk.END, p)

def choose_dir(entry: tk.Entry, title="Choisir un dossier"):
    path = filedialog.askdirectory(title=title)
    if path:
        entry.delete(0, tk.END); entry.insert(0, path)

def append_log(text: tk.Text, s: str):
    text.configure(state="normal")
    text.insert(tk.END, s)
    text.see(tk.END)
    text.configure(state="disabled")

def run_async(fn):
    th = threading.Thread(target=fn, daemon=True)
    th.start()
    return th

# Cross-platform opener
def open_path(path: str):
    if not path:
        return
    try:
        if sys.platform.startswith("darwin"):
            subprocess.Popen(["open", path])
        elif os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass

# --- Command builder ---------------------------------------------------------
def build_pipeline_cmd(values: dict) -> list[str]:
    """
    Construit la commande en appelant directement le binaire 'evalnat-pipeline'
    embarqué dans l'app (pas d'interpréteur Python externe requis).
    """
    pipebin = pipeline_binary()
    if not pipebin:
        raise FileNotFoundError(
            "Binaire 'evalnat-pipeline' ou 'pipeline_evalnat' introuvable.\n"
            "Solutions :\n"
            " • Copiez l’un des deux (même contenu) dans\n"
            "   'EvalNat-Publipostage.app/Contents/MacOS/' ou '.../Frameworks/'.\n"
            " • Ou relancez et choisissez le binaire manuellement quand la boîte s’affiche."
        )

    args = [pipebin,
            "--classe", values["classe"],
            "--annee", values["annee"],
            "--out-dir", values["out_dir"]]

    # Découpage
    if values.get("no_split"):
        args += ["--no-split"]
    else:
        if not values.get("input_pdf"):
            raise ValueError("PDF d’entrée manquant (onglet Découpage PDF).")
        args += ["--input-pdf", values["input_pdf"]]
        # Forcés par le produit
        args += ["--keep-accents", "--auto-ocr", "--ocr-lang", values.get("ocr_lang") or "fra"]

    # Fusion (toujours via exports SIECLE)
    parents_csvs = values.get("parents_csvs") or []
    if not parents_csvs:
        raise ValueError("Aucun CSV SIECLE fourni (onglet Récupération mails parents).")
    args += ["--parents"] + parents_csvs

    # Message commun aux parents (si présent)
    msg = values.get("message_text")
    if isinstance(msg, str) and msg.strip():
        args += ["--message-text", msg]

    # Build & checks (onglet supprimé → valeurs par défaut)
    args += ["--preflight-threshold", "0.8"]
    # Pas de --strict exposé

    # Publipostage Thunderbird
    if values.get("run_tb"):
        args += ["--run-tb"]
        if values.get("dry_run"):
            args += ["--dry-run"]
        if isinstance(values.get("limit"), int) and values["limit"] > 0:
            args += ["--limit", str(values["limit"])]
        if isinstance(values.get("skip"), int) and values["skip"] > 0:
            args += ["--skip", str(values["skip"])]
        if values.get("sleep"):
            args += ["--sleep", str(values["sleep"])]
        if values.get("csv_tb"):
            args += ["--csv-tb", values["csv_tb"]]
        if values.get("tb_binary"):
            args += ["--tb-binary", values["tb_binary"]]

    return args

# --- Main App ---------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x720")

        # valeurs partagées
        self.vars = {
            "scripts_dir": tk.StringVar(value=FROZEN_BASE),
            "classe": tk.StringVar(value="4D"),
            "annee": tk.StringVar(value=DEFAULT_YEAR),
            "verbose": tk.BooleanVar(value=False),

            # Split
            "input_pdf": tk.StringVar(value=""),
            "out_dir": tk.StringVar(value=""),
            "keep_accents": tk.BooleanVar(value=True),   # forcé, non affiché
            "auto_ocr": tk.BooleanVar(value=True),       # forcé, non affiché
            "ocr_lang": tk.StringVar(value="fra"),
            "no_split": tk.BooleanVar(value=False),

            # Merge
            "parents_csvs": [],  # via Listbox

            # Message commun
            "message_text": tk.StringVar(value=""),

            # TB
            "run_tb": tk.BooleanVar(value=True),
            "dry_run": tk.BooleanVar(value=False),
            "limit": tk.IntVar(value=0),
            "skip": tk.IntVar(value=0),
            "sleep": tk.DoubleVar(value=0.7),
            "csv_tb": tk.StringVar(value=""),
            "tb_binary": tk.StringVar(value=""),
        }

        self.build_ui()
        self._current_step = 1
        self._total_steps = 4

    def build_ui(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        # Onglet 0 — Contexte
        f0 = ttk.Frame(nb); nb.add(f0, text="Contexte")
        self._tab_context(f0)

        # Onglet 1 — Découpage
        f1 = ttk.Frame(nb); nb.add(f1, text="1) Découpage PDF")
        self._tab_split(f1)

        # Onglet 2 — Fusion
        f2 = ttk.Frame(nb); nb.add(f2, text="2) Récupération mails parents")
        self._tab_merge(f2)

        # Onglet 3 — Message aux parents
        f3m = ttk.Frame(nb); nb.add(f3m, text="3) Message aux parents")
        self._tab_message(f3m)

        # Onglet 4 — Publipostage
        f5 = ttk.Frame(nb); nb.add(f5, text="4) Publipostage")
        self._tab_tb(f5)

        # Zone log OU barre de progression + bouton lancer
        bottom = ttk.Frame(self); bottom.pack(fill="both", expand=False, padx=8, pady=(0,8))
        # Widgets de sortie
        self.log = tk.Text(bottom, height=14, wrap="word", state="disabled")
        self.sb = ttk.Scrollbar(bottom, command=self.log.yview)
        self.log.configure(yscrollcommand=self.sb.set)
        self.progress = ttk.Progressbar(bottom, mode="indeterminate")
        self.percent = ttk.Label(bottom, text="0%")

        # Cadre boutons
        buttons_frame = ttk.Frame(self)
        buttons_frame.pack(pady=(0,10))
        ttk.Button(buttons_frame, text="Ouvrir le dossier de sortie",
                   command=lambda: open_path(self.vars["out_dir"].get())).pack(side="left", padx=(0,8))
        self.btn_run = ttk.Button(buttons_frame, text="C'est parti ▶", command=self.on_run)
        self.btn_run.pack(side="left")

        # Affichage initial selon 'verbose'
        self._toggle_verbose()
        # Reagir au changement de la case 'verbose'
        try:
            self.vars["verbose"].trace_add("write", lambda *args: self._toggle_verbose())
        except Exception:
            pass

    # -- Tabs -----------------------------------------------------------------
    def _tab_context(self, f):
        row=0
        ttk.Label(f, text="Classe").grid(row=row, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.vars["classe"], width=10).grid(row=row, column=0, sticky="w", padx=(60,0))
        ttk.Label(f, text="Année scolaire").grid(row=row, column=0, sticky="w", padx=(150,0))
        ttk.Entry(f, textvariable=self.vars["annee"], width=14).grid(row=row, column=0, sticky="w", padx=(270,0))
        row += 1
        ttk.Checkbutton(
            f,
            text="Mode verbose (afficher le log détaillé)",
            variable=self.vars["verbose"]
        ).grid(row=row, column=0, sticky="w")
        f.grid_columnconfigure(0, weight=1)
    def _toggle_verbose(self):
        # Nettoyage du frame 'bottom' actuel
        parent = self.log.master
        for w in (self.log, self.sb, self.progress, self.percent):
            try:
                w.pack_forget()
            except Exception:
                pass
        if self.vars["verbose"].get():
            # Afficher le log + scrollbar
            self.log.pack(fill="both", expand=True, side="left")
            self.sb.pack(side="right", fill="y")
        else:
            # Afficher la barre de progression + pourcentage
            try:
                self.progress.configure(mode="determinate", maximum=100, value=0)
            except Exception:
                pass
            self.progress.pack(fill="x", expand=True, side="left", padx=(0,6))
            self.percent.config(text="0%")
            self.percent.pack(side="left")

    def _set_progress(self, value: float):
        try:
            v = max(0, min(100, float(value)))
            self.progress.configure(value=v)
            self.percent.configure(text=f"{int(v)}%")
        except Exception:
            pass

    def _reset_progress(self):
        try:
            self.progress.configure(value=0)
            self.percent.configure(text="0%")
        except Exception:
            pass

    def _tab_split(self, f):
        row=0
        ttk.Checkbutton(f, text="Sauter le découpage (réutiliser un dossier déjà prêt)", variable=self.vars["no_split"]).grid(row=row, column=0, sticky="w"); row+=1
        ttk.Label(f, text="PDF à découper (export complet)").grid(row=row, column=0, sticky="w"); row+=1
        e_pdf = ttk.Entry(f, textvariable=self.vars["input_pdf"], width=90); e_pdf.grid(row=row, column=0, sticky="we")
        ttk.Button(f, text="Choisir PDF…", command=lambda: choose_file(e_pdf, "Choisir un PDF", (("PDF", "*.pdf"), ("Tous", "*.*")))).grid(row=row, column=1, padx=6)
        row+=1
        ttk.Label(f, text="Dossier de sortie (pièces jointes par élève)").grid(row=row, column=0, sticky="w"); row+=1
        e_out = ttk.Entry(f, textvariable=self.vars["out_dir"], width=90); e_out.grid(row=row, column=0, sticky="we")
        box = ttk.Frame(f); box.grid(row=row, column=1, padx=6, sticky="n")
        ttk.Button(box, text="Choisir…", command=lambda: choose_dir(e_out, "Choisir un dossier")).pack()
        ttk.Button(box, text="Ouvrir", command=lambda: open_path(self.vars["out_dir"].get())).pack(pady=(6,0))
        row+=1
        ttk.Label(f, text="Langue OCR").grid(row=row, column=0, sticky="w", padx=(260,0))
        ttk.Entry(f, textvariable=self.vars["ocr_lang"], width=8).grid(row=row, column=0, sticky="w", padx=(340,0))
        f.grid_columnconfigure(0, weight=1)

    def _tab_merge(self, f):
        row=0
        ttk.Label(f, text="Exports SIECLE à fusionner").grid(row=row, column=0, sticky="w"); row+=1
        self.lb_csvs = tk.Listbox(f, height=6); self.lb_csvs.grid(row=row, column=0, sticky="we")
        ttk.Button(f, text="Ajouter CSV…", command=lambda: choose_files(self.lb_csvs)).grid(row=row, column=1, padx=6, sticky="n")
        ttk.Button(f, text="Supprimer sélection", command=self._remove_selected_csvs).grid(row=row, column=1, padx=6, pady=(40,0), sticky="n")
        row+=1
        lbl = ttk.Label(
            f,
            text="Récupérer le fichier au format CSV dans : Exploitation → Extractions personnalisées → « Adresse mail parents » (par classe : choisissez la classe à exporter).",
            wraplength=760,
            justify="left"
        )
        lbl.grid(row=row, column=0, columnspan=2, sticky="w", pady=(8,0))
        f.grid_columnconfigure(0, weight=1)

    def _progress_from_line(self, line: str):
        """
        Déduit une progression (%) à partir d'une ligne du stdout du pipeline.
        Heuristique :
          - détecte "Étape i/n" → base = (i-1)/n
          - si "Pages traitées: a / b" pendant l'étape, ajoute la fraction locale
          - détecte un pourcentage '[ 75% ]'
        """
        try:
            # Détection directe d'un pourcentage de type "[ 75% ]"
            percent_match = re.search(r"\[\s*(\d+)\s*%\s*\]", line)
            if percent_match:
                pct = int(percent_match.group(1))
                self._set_progress(pct)
                return

            step_match = re.search(r"Étape\s+(\d+)\s*/\s*(\d+)", line)
            if step_match:
                self._current_step = int(step_match.group(1))
                self._total_steps = max(1, int(step_match.group(2)))
                # avance au début de l'étape courante
                base = (self._current_step - 1) / self._total_steps * 100.0
                self._set_progress(base)
                return

            # Progression fine pendant l'étape 1 (découpage) : "Pages traitées: x / y"
            pages_match = re.search(r"Pages traitées:\s*(\d+)\s*/\s*(\d+)", line)
            if pages_match:
                done = int(pages_match.group(1))
                total = max(1, int(pages_match.group(2)))
                step = getattr(self, "_current_step", 1)
                total_steps = getattr(self, "_total_steps", 4)
                base = (step - 1) / total_steps
                local = min(1.0, done / total)
                progress = (base + local / total_steps) * 100.0
                self._set_progress(progress)
                return

            # Si on voit "✅ Pipeline terminé" / "✅ Terminé" → 100 %
            if "✅ Pipeline terminé" in line or "✅ Terminé" in line:
                self._set_progress(100.0)
                return
        except Exception:
            # En cas d'erreur de parsing, on ne casse pas l'UI
            pass

    def _tab_message(self, f):
        row = 0
        ttk.Label(
            f,
            text="Rédigez le message commun aux parents (sera copié dans « CorpsMessage » du mail merge)"
        ).grid(row=row, column=0, sticky="w"); row += 1

        helper = ttk.Label(
            f,
            text="Vous pouvez utiliser des retours à la ligne. Le message sera identique pour tous.",
            wraplength=760, justify="left"
        )
        helper.grid(row=row, column=0, sticky="w", pady=(0,6)); row += 1

        self.txt_message = tk.Text(f, height=12, wrap="word")
        self.txt_message.grid(row=row, column=0, columnspan=2, sticky="nsew"); row += 1

        btns = ttk.Frame(f); btns.grid(row=row, column=0, sticky="w", pady=(6,0))
        def _import_message():
            path = filedialog.askopenfilename(
                title="Importer un message texte",
                filetypes=(("Texte", "*.txt"), ("Tous", "*.*"))
            )
            if path and os.path.exists(path):
                try:
                    with open(path, "r", encoding="utf-8") as fh:
                        content = fh.read()
                    self.txt_message.delete("1.0", "end")
                    self.txt_message.insert("1.0", content)
                except Exception as e:
                    messagebox.showerror("Erreur", f"Impossible de lire le fichier :\n{e}")
        ttk.Button(btns, text="Importer depuis un .txt…", command=_import_message).pack(side="left")

        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(2, weight=1)



    def _tab_tb(self, f):
        row=0
        ttk.Checkbutton(f, text="Ouvrir les brouillons Thunderbird", variable=self.vars["run_tb"]).grid(row=row, column=0, sticky="w"); row+=1
        ttk.Label(f, text="CSV pour TB (optionnel)").grid(row=row, column=0, sticky="w"); row+=1
        e_tb = ttk.Entry(f, textvariable=self.vars["csv_tb"], width=90); e_tb.grid(row=row, column=0, sticky="we")
        ttk.Button(f, text="Choisir CSV…", command=lambda: choose_file(e_tb, "Choisir CSV", (("CSV", "*.csv"), ("Tous", "*.*")))).grid(row=row, column=1, padx=6)
        row+=1
        ttk.Label(f, text="Binaire Thunderbird (optionnel)").grid(row=row, column=0, sticky="w"); row+=1
        e_bin = ttk.Entry(f, textvariable=self.vars["tb_binary"], width=90); e_bin.grid(row=row, column=0, sticky="we")
        ttk.Button(f, text="Choisir exécutable…", command=lambda: choose_file(e_bin, "Choisir Thunderbird")).grid(row=row, column=1, padx=6)
        f.grid_columnconfigure(0, weight=1)

    def vars_entry(self, entry_widget):
        return entry_widget

    def _remove_selected_csvs(self):
        sel = list(self.lb_csvs.curselection())[::-1]
        for i in sel:
            self.lb_csvs.delete(i)

    # -- Run ------------------------------------------------------------------
    def collect_values(self) -> dict:
        v = {k: var.get() if not isinstance(var, list) else var for k,var in self.vars.items()}
        v["parents_csvs"] = list(self.lb_csvs.get(0, tk.END))
        if not v["classe"].strip():
            raise ValueError("Champ 'Classe' vide.")
        if not v["annee"].strip():
            raise ValueError("Champ 'Année' vide.")
        if not v["out_dir"].strip():
            raise ValueError("Dossier de sortie des PDFs manquant (onglet Découpage PDF).")

        # Récupération du message multi-ligne (onglet « Message aux parents »)
        try:
            if hasattr(self, "txt_message"):
                v["message_text"] = self.txt_message.get("1.0", "end-1c")
        except Exception:
            v["message_text"] = v.get("message_text", "")

        return v

    def on_run(self):
        try:
            vals = self.collect_values()
            cmd = build_pipeline_cmd(vals)
        except FileNotFoundError as e:
            messagebox.showerror("Binaire introuvable", str(e))
            return
        except Exception as e:
            messagebox.showerror("Paramètres invalides", str(e))
            return

        self.attributes("-alpha", 0.98)

        # Réinitialiser la barre de progression si non-verbose
        if not self.vars["verbose"].get():
            self._reset_progress()
        try:
            self.btn_run.configure(state="disabled")
        except Exception:
            pass

        def worker(vals_local=vals, cmd_local=cmd):
            rc = None
            try:
                try:
                    if hasattr(sys, "_MEIPASS"):
                        os.chdir(FROZEN_BASE)
                except Exception:
                    pass

                if self.vars["verbose"].get():
                    append_log(self.log, "\n\n" + " ".join(shlex.quote(c) for c in cmd_local) + "\n")
                try:
                    bin_path = cmd_local[0]
                    if os.path.exists(bin_path) and not os.access(bin_path, os.X_OK):
                        os.chmod(bin_path, 0o755)
                    if sys.platform.startswith("darwin"):
                        try:
                            subprocess.run(["xattr", "-d", "com.apple.quarantine", bin_path], check=False)
                        except Exception:
                            pass
                except Exception:
                    pass
                # Subprocess creation with forced UTF-8 and no console on Windows
                creationflags = 0
                startupinfo = None
                if os.name == "nt":
                    try:
                        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
                        startupinfo = subprocess.STARTUPINFO()
                        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    except Exception:
                        pass
                proc = subprocess.Popen(
                    cmd_local,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    startupinfo=startupinfo,
                    creationflags=creationflags,
                )
                for line in proc.stdout:
                    if self.vars["verbose"].get():
                        append_log(self.log, line)
                    # Mise à jour de la progression (même en mode non-verbose)
                    try:
                        self.after(0, self._progress_from_line, line)
                    except Exception:
                        pass
                rc = proc.wait()
                if self.vars["verbose"].get():
                    if rc == 0:
                        append_log(self.log, "\n✅ Terminé sans erreur.\n")
                    else:
                        append_log(self.log, f"\n❌ Erreur (code {rc}).\n")
            except FileNotFoundError as e:
                if self.vars["verbose"].get():
                    append_log(self.log, f"\n❌ Fichier introuvable: {e}\n")
            except Exception as e:
                if self.vars["verbose"].get():
                    append_log(self.log, f"\n❌ Exception: {e}\n")
            finally:
                try:
                    if not self.vars["verbose"].get():
                        # Si succès et pas déjà à 100, on force à 100
                        if rc == 0:
                            self._set_progress(100.0)
                except Exception:
                    pass
                # Arrêter la barre de progression si nécessaire
                try:
                    self.progress.stop()
                except Exception:
                    pass
                try:
                    self.btn_run.configure(state="normal")
                except Exception:
                    pass
                self.attributes("-alpha", 1.0)

        run_async(worker)

# --- main --------------------------------------------------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
