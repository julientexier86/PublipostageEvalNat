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

# --- Debug logging (helps when the app seems to "bounce" and quit on macOS) ---
from datetime import datetime
DEBUG_LOG = os.path.expanduser("~/Library/Logs/PublipostageEVALNAT.log")

def dlog(msg: str):
    try:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {msg}\n"
        # print to console if launched from terminal
        try:
            print(line, end="")
        except Exception:
            pass
        # append to file
        with open(DEBUG_LOG, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass

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
                dlog(f"pipeline_binary: found → {c}")
                return str(c)
        except Exception:
            continue

    # Si rien trouvé : proposer sélection manuelle + tracer où on a cherché (verbose)
    try:
        # Ne tente pas d'ouvrir des boîtes de dialogue si aucun root Tk n'existe encore
        try:
            import tkinter as _tk
            if getattr(_tk, "_default_root", None) is None:
                raise RuntimeError("no_tk_root_yet")
        except Exception:
            raise

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
        # En phase d'initialisation (pas encore de root Tk), on reviendra plus tard.
        return None

    # Petit log en console pour aider au debug si lancé en dev
    try:
        dlog("[DEBUG] Binaire pipeline introuvable. Chemins testés :")
        for c in candidates:
            dlog(f"  - {c}")
    except Exception:
        pass

    return None

APP_TITLE = "PublipostageEVALNAT"
APP_VERSION = "macOS bundle v1.2"

dlog("=== PublipostageEVALNAT starting ===")
dlog(f"Python: {sys.version.split()[0]} | Platform: {sys.platform} | Frozen: {getattr(sys, 'frozen', False)}")

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

def append_log(text: tk.Text | None, s: str):
    if not text:
        return
    try:
        text.configure(state="normal")
        text.insert(tk.END, s)
        text.see(tk.END)
        text.configure(state="disabled")
    except Exception:
        pass

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
    dlog(f"build_pipeline_cmd called with values: classe={values.get('classe')} annee={values.get('annee')}")
    pipebin = pipeline_binary()
    if not pipebin:
        # Laisse l'appelant gérer l'UX (boîte de dialogue déjà prévue côté GUI)
        raise FileNotFoundError("Binaire pipeline introuvable. Merci de le sélectionner depuis l'onglet Contexte, ou rebuild l'app avec le binaire embarqué.")

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

    # Message commun + OBJET
    msg = values.get("message_text")
    if isinstance(msg, str) and msg.strip():
        args += ["--message-text", msg]

    subj = values.get("subject_template")
    if not isinstance(subj, str) or not subj.strip():
        subj = "Evaluations nationales - {NOM} {Prénom} ({Classe})"
    args += ["--subject-template", subj]

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
        self.withdraw()  # avoid a brief flash; show once ready
        dlog("App.__init__ entered")
        self.title(APP_TITLE)
        self.geometry("980x720")
        # Résolution différée du binaire pipeline (après que la fenêtre existe)
        self._pipebin_path = ""

        # valeurs partagées
        self.vars = {
            "classe": tk.StringVar(value="4D"),
            "annee": tk.StringVar(value=DEFAULT_YEAR),
            "verbose": tk.BooleanVar(value=False),

            # Split
            "input_pdf": tk.StringVar(value=""),
            "out_dir": tk.StringVar(value=""),
            "ocr_lang": tk.StringVar(value="fra"),
            "no_split": tk.BooleanVar(value=False),

            # Merge
            "parents_csvs": [],

            # Message commun
            "subject_template": tk.StringVar(value="Evaluations nationales - {NOM} {Prénom} ({Classe})"),
            "message_text": tk.StringVar(value=""),

            # TB
            "run_tb": tk.BooleanVar(value=True),
            "dry_run": tk.BooleanVar(value=False),
            "limit": tk.IntVar(value=0),
            "skip": tk.IntVar(value=0),
            "sleep": tk.DoubleVar(value=0.7),
            "tb_binary": tk.StringVar(value=""),
        }

        self.build_ui()
        self.log_text = None  # created lazily in Context tab when verbose is toggled
        self._current_step = 1
        self._total_steps = 4

        self.update_idletasks()
        self.deiconify()
        dlog("Main window built and deiconified")

        # Résolution tardive du binaire pour éviter tout crash avant mainloop
        self.after(200, self._late_resolve_pipeline)

    def build_ui(self):
        # En-tête (version + binaire pipeline détecté)
        header = ttk.Frame(self)
        header.pack(fill="x", padx=8, pady=(8, 0))

        lbl_title = ttk.Label(header, text=f"{APP_TITLE} — {APP_VERSION}", font=("TkDefaultFont", 11, "bold"))
        lbl_title.pack(side="left")

        pipe_txt = self._pipebin_path if self._pipebin_path else "binaire pipeline non détecté"
        self._lbl_pipe = ttk.Label(header, text=f" • Pipeline: {pipe_txt}", foreground="#555555")
        self._lbl_pipe.pack(side="left", padx=(10, 0))

        # Progression
        prog_frame = ttk.Frame(header)
        prog_frame.pack(side="right", padx=6)
        ttk.Label(prog_frame, text="Progression :").pack(side="left")
        self.progress = ttk.Progressbar(prog_frame, length=220, mode="determinate", maximum=100)
        self.progress.pack(side="left", padx=(4, 0))

        # Notebook (5 onglets)
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        # Onglet 0 — Contexte
        f0 = ttk.Frame(nb); nb.add(f0, text="Contexte")
        App._tab_context(self, f0)

        # Onglet 1 — Découpage PDF
        f1 = ttk.Frame(nb); nb.add(f1, text="1) Découpage PDF")
        App._tab_split(self, f1)

        # Onglet 2 — Récupération mails parents
        f2 = ttk.Frame(nb); nb.add(f2, text="2) Récupération mails parents")
        App._tab_parents(self, f2)

        # Onglet 3 — Message aux parents
        f3 = ttk.Frame(nb); nb.add(f3, text="3) Message aux parents")
        App._tab_message(self, f3)

        # Onglet 4 — Publipostage
        f4 = ttk.Frame(nb); nb.add(f4, text="4) Publipostage")
        App._tab_tb(self, f4)

        dlog("UI constructed (tabs + progress)")

    def _tab_context(self, parent: ttk.Frame):
        frm = parent
        frm.columnconfigure(1, weight=1)

        # Classe
        ttk.Label(frm, text="Classe :").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(frm, textvariable=self.vars["classe"], width=12).grid(row=0, column=1, sticky="w", padx=6, pady=4)

        # Année scolaire
        ttk.Label(frm, text="Année scolaire :").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(frm, textvariable=self.vars["annee"], width=14).grid(row=1, column=1, sticky="w", padx=6, pady=4)

        # Toggle verbose (journal)
        self.var_verbose_ui = tk.BooleanVar(value=False)
        def _toggle_verbose():
            if self.var_verbose_ui.get():
                # create section if needed
                if getattr(self, "log_container", None) is None:
                    self.log_container = ttk.LabelFrame(frm, text="Journal du pipeline (verbose)")
                    self.log_container.grid(row=10, column=0, columnspan=3, sticky="nsew", padx=6, pady=(8,6))
                    frm.rowconfigure(10, weight=1)
                    self.log_text = tk.Text(self.log_container, height=12, wrap="word", state="disabled")
                    self.log_text.pack(fill="both", expand=True, padx=6, pady=6)
                else:
                    self.log_container.grid()
            else:
                if getattr(self, "log_container", None) is not None:
                    self.log_container.grid_remove()
            self.update_idletasks()

        ttk.Checkbutton(frm, text="Mode verbose (afficher le journal)", variable=self.var_verbose_ui, command=_toggle_verbose)\
            .grid(row=2, column=1, sticky="w", padx=6, pady=(8,4))

    def _late_resolve_pipeline(self):
        try:
            path = pipeline_binary()
            if isinstance(path, str):
                self._pipebin_path = path
            else:
                self._pipebin_path = ""
        except Exception:
            self._pipebin_path = ""
        # Mettre à jour l'étiquette d'en-tête si elle existe
        try:
            if hasattr(self, "_lbl_pipe"):
                pipe_txt = self._pipebin_path if self._pipebin_path else "binaire pipeline non détecté"
                self._lbl_pipe.configure(text=f" • Pipeline: {pipe_txt}")
        except Exception:
            pass
        dlog(f"_late_resolve_pipeline: resolved = {self._pipebin_path or 'None'}")

    def _tab_split(self, parent: ttk.Frame):
        parent.columnconfigure(1, weight=1)
        ttk.Label(parent, text="PDF source (si découpage) :").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        e_pdf = ttk.Entry(parent, textvariable=self.vars["input_pdf"])
        e_pdf.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(parent, text="Parcourir…", command=lambda: choose_file(e_pdf, "Choisir le PDF", (("PDF", "*.pdf"), ("Tous", "*.*")))).grid(row=0, column=2, padx=6, pady=4)

        ttk.Label(parent, text="Dossier de sortie (PDF par élève) :").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        e_out = ttk.Entry(parent, textvariable=self.vars["out_dir"])
        e_out.grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(parent, text="Choisir…", command=lambda: choose_dir(e_out, "Choisir le dossier de sortie")).grid(row=1, column=2, padx=6, pady=4)

        chk = ttk.Checkbutton(parent, text="Ne pas découper (réutiliser les PDFs déjà présents)", variable=self.vars["no_split"])
        chk.grid(row=2, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(parent, text="Langue OCR :").grid(row=3, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(parent, textvariable=self.vars["ocr_lang"], width=10).grid(row=3, column=1, sticky="w", padx=6, pady=4)

    def _tab_parents(self, parent: ttk.Frame):
        parent.columnconfigure(1, weight=1)
        ttk.Label(parent, text="Export SIECLE (CSV) :").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        e_csv = ttk.Entry(parent, textvariable=tk.StringVar(), width=40)
        e_csv.grid(row=0, column=1, sticky="ew", padx=6, pady=4)

        def _choose():
            path = filedialog.askopenfilename(title="Choisir le CSV SIECLE", filetypes=[("CSV", "*.csv"), ("Tous", "*.*")])
            if path:
                e_csv.delete(0, tk.END); e_csv.insert(0, path)
                self.vars["parents_csvs"] = [path]

        ttk.Button(parent, text="Parcourir…", command=_choose).grid(row=0, column=2, padx=6, pady=4)

    def _tab_message(self, parent: ttk.Frame):
        parent.columnconfigure(1, weight=1)
        ttk.Label(parent, text="Objet (modèle) :").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(parent, textvariable=self.vars["subject_template"]).grid(row=0, column=1, columnspan=2, sticky="ew", padx=6, pady=4)

        ttk.Label(parent, text="Message aux parents :").grid(row=1, column=0, sticky="nw", padx=6, pady=4)
        self.msg_text = tk.Text(parent, height=12, wrap="word")
        self.msg_text.grid(row=1, column=1, columnspan=2, sticky="nsew", padx=6, pady=4)
        parent.rowconfigure(1, weight=1)

        def _sync():
            self.vars["message_text"].set(self.msg_text.get("1.0", "end-1c"))
        self.msg_text.bind("<FocusOut>", lambda e: _sync())

        # Right-click context menu (paste/copy/cut)
        menu = tk.Menu(self.msg_text, tearoff=0)
        menu.add_command(label="Coller", command=lambda: self.msg_text.event_generate("<<Paste>>"))
        menu.add_command(label="Copier", command=lambda: self.msg_text.event_generate("<<Copy>>"))
        menu.add_command(label="Couper", command=lambda: self.msg_text.event_generate("<<Cut>>"))
        def _popup(e):
            try:
                menu.tk_popup(e.x_root, e.y_root)
            finally:
                menu.grab_release()
        self.msg_text.bind("<Button-2>", _popup)  # some macOS configurations
        self.msg_text.bind("<Button-3>", _popup)

    def _tab_tb(self, parent: ttk.Frame):
        parent.columnconfigure(1, weight=1)

        ttk.Label(parent, text="Ouvrir automatiquement les brouillons Thunderbird").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Checkbutton(parent, variable=self.vars["run_tb"]).grid(row=0, column=1, sticky="w", padx=6, pady=4)

        # Advanced options toggle
        self.var_adv = tk.BooleanVar(value=False)
        def _toggle_adv():
            if self.var_adv.get():
                adv_frame.grid(row=2, column=0, columnspan=3, sticky="ew", padx=6, pady=6)
            else:
                adv_frame.grid_remove()

        ttk.Checkbutton(parent, text="Options avancées", variable=self.var_adv, command=_toggle_adv)\
            .grid(row=1, column=0, sticky="w", padx=6, pady=(4,4))

        # Advanced frame (initially hidden)
        adv_frame = ttk.Frame(parent)
        adv_frame.grid_remove()

        ttk.Label(adv_frame, text="Chemin Thunderbird (optionnel) :").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        e_tb = ttk.Entry(adv_frame, textvariable=self.vars["tb_binary"])
        e_tb.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(adv_frame, text="Parcourir…", command=lambda: choose_file(e_tb, "Choisir Thunderbird", (("Thunderbird", "thunderbird*"), ("Tous", "*.*")))).grid(row=0, column=2, padx=6, pady=4)

        ttk.Label(adv_frame, text="Limit / Skip / Sleep (s) :").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        row2 = ttk.Frame(adv_frame); row2.grid(row=1, column=1, sticky="w", padx=6, pady=4)
        ttk.Entry(row2, width=8, textvariable=self.vars["limit"]).pack(side="left", padx=(0,6))
        ttk.Entry(row2, width=8, textvariable=self.vars["skip"]).pack(side="left", padx=(0,6))
        ttk.Entry(row2, width=8, textvariable=self.vars["sleep"]).pack(side="left", padx=(0,6))

        # Run button
        run_frame = ttk.Frame(parent); run_frame.grid(row=99, column=0, columnspan=3, sticky="ew", padx=6, pady=10)
        run_frame.columnconfigure(0, weight=1)
        ttk.Button(run_frame, text="C'est parti", command=self._on_start).grid(row=0, column=0, sticky="ew")

    def _gather_values(self) -> dict:
        # Assure la synchronisation du Text => StringVar
        try:
            if hasattr(self, "msg_text"):
                self.vars["message_text"].set(self.msg_text.get("1.0", "end-1c"))
        except Exception:
            pass
        # parents_csvs is already maintained when choosing the file
        values = {
            "classe": self.vars["classe"].get(),
            "annee": self.vars["annee"].get(),
            "out_dir": self.vars["out_dir"].get(),
            "no_split": self.vars["no_split"].get(),
            "input_pdf": self.vars["input_pdf"].get(),
            "ocr_lang": self.vars["ocr_lang"].get(),
            "parents_csvs": self.vars.get("parents_csvs", []),
            "subject_template": self.vars["subject_template"].get(),
            "message_text": self.vars["message_text"].get(),
            "run_tb": self.vars["run_tb"].get(),
            "dry_run": self.vars["dry_run"].get(),
            "limit": self.vars["limit"].get(),
            "skip": self.vars["skip"].get(),
            "sleep": self.vars["sleep"].get(),
            "csv_tb": self.vars.get("csv_tb", tk.StringVar(value="")).get() if isinstance(self.vars.get("csv_tb"), tk.Variable) else "",
            "tb_binary": self.vars["tb_binary"].get(),
        }
        return values

    def _on_start(self):
        run_async(self._run_pipeline)

    def _run_pipeline(self):
        # Reset progression
        try:
            self.progress.configure(value=0)
        except Exception:
            pass

        vals = self._gather_values()
        try:
            cmd = build_pipeline_cmd(vals)
        except Exception as e:
            messagebox.showerror("Paramétrage incomplet", str(e))
            return

        append_log(self.log_text, "Commande:\n  " + " ".join(shlex.quote(x) for x in cmd) + "\n\n")
        dlog("Launching pipeline: " + " ".join(cmd))

        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                env=os.environ.copy(),
                text=True,
                bufsize=1,
                universal_newlines=True,
            )
        except FileNotFoundError as e:
            messagebox.showerror("Binaire introuvable", f"{e}")
            return
        except Exception as e:
            messagebox.showerror("Erreur de lancement", f"{e}")
            return

        percent_re = re.compile(r"^\[\s*(\d+)%\]")
        for line in proc.stdout:  # type: ignore[union-attr]
            if not line:
                continue
            append_log(self.log_text, line)
            m = percent_re.match(line.strip())
            if m:
                try:
                    p = int(m.group(1))
                    self.progress.configure(value=p)
                    self.update_idletasks()
                except Exception:
                    pass

        rc = proc.wait()
        if rc == 0:
            self.progress.configure(value=100)
            append_log(self.log_text, "\n✅ Terminé sans erreur.\n")
        else:
            append_log(self.log_text, f"\n❌ Erreur (code {rc}).\n")


if __name__ == "__main__":
    def _excepthook(exc_type, exc, tb):
        try:
            import traceback
            trace = "".join(traceback.format_exception(exc_type, exc, tb))
            dlog("UNHANDLED EXCEPTION:\n" + trace)
        except Exception:
            pass
        # Try to show a dialog if possible
        try:
            import tkinter as tk
            from tkinter import messagebox
            if getattr(tk, "_default_root", None) is None:
                r = tk.Tk(); r.withdraw()
            messagebox.showerror("Erreur fatale", f"L'application a rencontré une erreur au démarrage :\n{exc}")
        except Exception:
            pass
        os._exit(1)

    sys.excepthook = _excepthook

    try:
        dlog("Creating App() ...")
        app = App()
        dlog("Entering mainloop()")
        app.mainloop()
        dlog("Exited mainloop()")
    except Exception as e:
        _excepthook(type(e), e, e.__traceback__)