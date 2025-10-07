#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

"""
Pipeline EvalNat — orchestre tes 4 scripts existants dans le bon ordre :
1) split_4C.py           → génère les PDFs par élève (Fr/Math) dans OUTPUT_DIR
2) merge_parents_4e.py   → fusionne les exports Siècle → parents_4e_merged.csv
3) build_mailmerge_...   → construit mailmerge_4e.csv en scannant OUTPUT_DIR
4) tb_mailmerge_mac.py   → (optionnel) ouvre les brouillons Thunderbird (macOS)


# Build recommandé (onefile) pour le pipeline, en embarquant les 4 scripts :
#   python -m PyInstaller --onefile --console --name evalnat-pipeline \
#       --add-data "split_4C.py:." \
#       --add-data "merge_parents_4e.py:." \
#       --add-data "build_mailmerge_4e_from_merged_v5.py:." \
#       --add-data "tb_mailmerge_mac.py:." \
#       pipeline_evalnat.py

Usage (exemple) :
  python3 pipeline_evalnat.py \
    --classe 4D \
    --annee "2025-2026" \
    --input-pdf "/Users/julien/Downloads/export_126892_all.pdf" \
    --out-dir "/Users/julien/Downloads/Publipostage_4D" \
    --parents "exportCSVExtraction4A.csv" "exportCSVExtraction4B.csv" \
              "exportCSVExtraction4C.csv" "exportCSVExtraction4D.csv" \
    --run-tb --sleep 0.7
    # --message-text "Votre message aux parents..."  # (ou) --message-file "/chemin/message.txt"

Par défaut :
- ouvre des brouillons Thunderbird avec col. corps = 'CorpsMessage' (conforme à build_mailmerge_v5)
- tu peux passer --dry-run pour tester sans ouvrir TB
- tu peux fournir un message commun aux parents via --message-text ou --message-file (remplit/écrase la colonne 'CorpsMessage').
"""

import argparse, sys, os, subprocess, shutil, csv, re, unicodedata, runpy, importlib, textwrap, warnings, locale
# --- Forcer l'inclusion de pandas dans le binaire (utilisé par merge_parents_4e) ---
try:
    import pandas  # type: ignore  # noqa: F401
except Exception:
    pass
from pathlib import Path

# ========= Sortie console robuste (Windows/UTF-8) =========
try:
    # Force UTF-8 + remplacement des caractères non mappables (évite 'charmap')
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass
os.environ.setdefault("PYTHONUTF8", "1")
os.environ.setdefault("PYTHONIOENCODING", "UTF-8")
# Si Windows, tenter de basculer la console en UTF-8 (évite 'charmap')
if os.name == "nt":
    try:
        # Change codepage to UTF-8 for the current console session
        os.system("chcp 65001 >NUL")
    except Exception:
        pass

# Tenter de forcer un environnement UTF‑8 même sous Windows (inoffensif ailleurs)
os.environ.setdefault("LANG", "C.UTF-8")
os.environ.setdefault("LC_ALL", "C.UTF-8")
try:
    locale.setlocale(locale.LC_ALL, "")
except Exception:
    pass

# Supprimer les FutureWarning bruyants (ex: pandas groupby.apply)
with warnings.catch_warnings():
    warnings.simplefilter("ignore", FutureWarning)

# Assainir l’affichage des warnings (utilise la même logique _safe que print)
_orig_showwarning = warnings.showwarning
def _safe_showwarning(message, category, filename, lineno, file=None, line=None):
    try:
        _orig_showwarning(message, category, filename, lineno, file=file, line=line)
    except Exception:
        # Repli sans caractères problématiques
        msg = f"{category.__name__}: {message} [{filename}:{lineno}]"
        _bi.print(_safe(msg), file=file or sys.stderr, flush=True)
warnings.showwarning = _safe_showwarning

def _safe(s: str) -> str:
    enc = (sys.stdout.encoding or 'utf-8')
    try:
        s.encode(enc)
        return s
    except Exception:
        return s.encode('ascii', 'ignore').decode('ascii')

# Impression "safe" globale (tous les print deviennent robustes)
import builtins as _bi
def _safe_print(*args, **kw):
    # Toujours forcer flush pour que la progression s'affiche en temps réel
    if "flush" not in kw:
        kw["flush"] = True
    try:
        _bi.print(*[_safe(str(a)) for a in args], **kw)
    except Exception:
        # Repli durci : éviter 'file' non sérialisable et forcer ASCII
        safe_kw = {k: v for k, v in kw.items() if k != "file"}
        _bi.print(*[str(a).encode("ascii", "ignore").decode("ascii") for a in args], **safe_kw)

print = _safe_print  # remplace le print globalement


# ========= Icônes / ASCII fallback =========
USE_ASCII = (os.name == 'nt')  # Toujours ASCII sous Windows pour éviter les soucis "charmap"
ICON_SCAN = 'SCAN' if USE_ASCII else '🔎'
ICON_OK   = 'OK'   if USE_ASCII else '✅'
ICON_ERR  = 'ERR'  if USE_ASCII else '❌'
ICON_WARN  = 'WARN' if USE_ASCII else '⚠️'
ICON_INFO  = 'INFO' if USE_ASCII else 'ℹ️'
ICON_SKIP  = 'SKIP' if USE_ASCII else '⏭️'
ICON_PLAY  = 'RUN'  if USE_ASCII else '▶'
ICON_ARROW = '->'   if USE_ASCII else '→'
BULLET     = '*'    if USE_ASCII else '•'

# ========= Progression simple (pourcentages) =========
PROG_POINTS = {0: "Démarrage", 25: "Split", 50: "Fusion parents", 75: "MailMerge", 90: "Thunderbird", 100: "Terminé"}
def progress(pct: int, extra: str = ""):
    try:
        label = PROG_POINTS.get(pct, "")
        msg = f"[{pct:>3}%] {label}"
        if extra:
            msg += f" — {extra}"
        print(msg, flush=True)
    except Exception:
        # repli ultra-sûr
        print(f"{pct}% {extra}", flush=True)

def p(msg: str):
    print(_safe(msg))

BASE_DIR = Path(__file__).resolve().parent  # dossier contenant le pipeline

# --- Helpers pour exécution en mode gelé (PyInstaller onefile) -------------
def _resource_path(rel: str = "") -> str:
    base = getattr(sys, "_MEIPASS", str(BASE_DIR))
    return str(Path(base) / rel)

def _ensure_meipass_on_syspath():
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass and meipass not in sys.path:
        sys.path.insert(0, meipass)

# --- Helper générique pour importer un module embarqué ou via fichier -------
def import_generic_module(modname: str, filename: str):
    try:
        return importlib.import_module(modname)
    except Exception as e_first:
        _ensure_meipass_on_syspath()
        try:
            return importlib.import_module(modname)
        except Exception:
            import importlib.util
            path = Path(_resource_path(filename))
            if path.exists():
                spec = importlib.util.spec_from_file_location(modname, str(path))
                if spec and spec.loader:
                    mod = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(mod)  # type: ignore
                    sys.modules[modname] = mod
                    return mod
            raise ImportError(
                f"[ERREUR] Module '{modname}' introuvable. Rebuild requis en embarquant '{filename}' via --add-data."
            ) from e_first

# Modules embarqués (PyInstaller) — appelés par import/runpy
SPLIT_MODULE    = "split_4C"
MERGE_MODULE    = "merge_parents_4e"
BUILD_MM_MODULE = "build_mailmerge_4e_from_merged_v5"
TB_MODULE       = "tb_mailmerge_mac"

# Helper to run a module with argv
def run_module_with_argv(modname: str, argv: list[str]) -> int:
    old_argv = sys.argv[:]
    try:
        sys.argv = [modname] + argv
        runpy.run_module(modname, run_name="__main__")
        return 0
    finally:
        sys.argv = old_argv

# --- Thunderbird binary resolver (Windows/macOS) ----------------------------
def resolve_tb_binary(user_path: str|None) -> str|None:
    """
    Return a usable Thunderbird binary path.
    Priority:
      1) user-provided path (if exists)
      2) common install locations per-OS
      3) whatever is on PATH (shutil.which)
    """
    import shutil as _sh
    cand: list[str|None] = []
    if user_path:
        cand.append(user_path)

    if os.name == "nt":
        cand += [
            r"C:\Program Files\Mozilla Thunderbird\thunderbird.exe",
            r"C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe",
            _sh.which("thunderbird"),
        ]
    else:
        # macOS default app bundle + PATH
        cand += [
            "/Applications/Thunderbird.app/Contents/MacOS/thunderbird",
            _sh.which("thunderbird"),
        ]

    for c in cand:
        if c and os.path.exists(c):
            return c
    return None

# --- Helpers encodage/CSV ---------------------------------------------------
def detect_sep(p: Path) -> str:
    try:
        with open(p, "r", encoding="utf-8-sig", newline="") as f:
            s = f.read(4096)
    except UnicodeDecodeError:
        with open(p, "r", encoding="cp1252", errors="replace", newline="") as f:
            s = f.read(4096)
    return ";" if s.count(";") >= s.count(",") else ","

def read_text_robust(p: Path) -> str:
    try:
        return p.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        return p.read_text(encoding="cp1252", errors="replace")

def open_csv_reader(path: Path):
    try:
        return open(path, "r", encoding="utf-8-sig", newline="")
    except UnicodeDecodeError:
        return open(path, "r", encoding="cp1252", errors="replace", newline="")

def open_csv_writer(path: Path):
    # On écrit toujours en UTF-8 (Mail Merge/Thunderbird savent le lire)
    return open(path, "w", encoding="utf-8", newline="")

def nfd(s: str) -> str:
    return unicodedata.normalize("NFD", s or "")

def squash(s: str) -> str:
    s = nfd(s).lower()
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"[^a-z0-9]+", "", s)

def count_pdfs_by_disc(base: Path, classe: str, annee: str) -> dict:
    discs_fr = ("Français","Francais","Franais")
    discs_ma = ("Mathématiques","Mathematiques","Mathmatiques","Maths")
    n_fr = n_ma = 0
    for pth in base.glob(f"{classe}_*_{annee}.pdf"):
        name = pth.name
        parts = name[:-4].split("_")
        if len(parts) >= 4:
            disc = parts[-2]
            if disc in discs_fr: n_fr += 1
            if disc in discs_ma: n_ma += 1
    return {"francais": n_fr, "maths": n_ma}

def scan_pdf_labels(base: Path) -> tuple[set[str], set[str]]:
    classes: set[str] = set()
    annees: set[str] = set()
    for pth in base.glob("*.pdf"):
        name = pth.name[:-4]
        parts = name.split("_")
        if len(parts) >= 2:
            classes.add(parts[0])
            annees.add(parts[-1])
    return classes, annees

# --- OCR helpers -----------------------------------------------------------
def quick_text_ratio(pdf_path: Path, max_pages: int = 6) -> float:
    try:
        from pdfminer.high_level import extract_text  # type: ignore
        from pdfminer.pdfpage import PDFPage  # type: ignore
        txt = extract_text(str(pdf_path)) or ""
        pages = 0
        with open(pdf_path, "rb") as f:
            for i, _ in enumerate(PDFPage.get_pages(f)):
                if i >= max_pages:
                    break
                pages += 1
        pages = max(pages, 1)
        chars = len((txt or "").strip())
        return chars / pages
    except Exception:
        pass
    try:
        from PyPDF2 import PdfReader  # type: ignore
        reader = PdfReader(str(pdf_path))
        pages = min(len(reader.pages), max_pages)
        chars = 0
        for i in range(pages):
            try:
                t = reader.pages[i].extract_text() or ""
            except Exception:
                t = ""
            chars += len(t.strip())
        return chars / max(pages, 1)
    except Exception:
        return 0.0

def have_cmd(cmd: str) -> bool:
    from shutil import which
    return which(cmd) is not None

def run_ocrmypdf(src: Path, dst: Path, lang: str = "fra"):
    cmd = ["ocrmypdf", "--force-ocr", "--rotate-pages", "--deskew",
           "--clean-final", "--skip-text", f"--language={lang}",
           str(src), str(dst)]
    print(f"{ICON_PLAY} ocrmypdf:", " ".join(cmd))
    subprocess.check_call(cmd)

def import_split_module():
    try:
        return importlib.import_module(SPLIT_MODULE)
    except Exception as e_first:
        _ensure_meipass_on_syspath()
        try:
            return importlib.import_module(SPLIT_MODULE)
        except Exception:
            import importlib.util
            split_path = Path(_resource_path("split_4C.py"))
            if split_path.exists():
                spec = importlib.util.spec_from_file_location(SPLIT_MODULE, str(split_path))
                if spec and spec.loader:
                    mod = importlib.util.module_from_spec(spec)  # type: ignore[attr-defined]
                    spec.loader.exec_module(mod)  # type: ignore
                    sys.modules[SPLIT_MODULE] = mod
                    return mod
            raise ImportError(
                "[ERREUR] split_4C introuvable. "
                "Rebuild du pipeline en embarquant le fichier :\n"
                "  python -m PyInstaller --onefile --console --name evalnat-pipeline \\\n"
                "    --add-data 'split_4C.py:.' pipeline_evalnat.py"
            ) from e_first

def import_merge_parents_module():
    return import_generic_module(MERGE_MODULE, "merge_parents_4e.py")

def run_build_mailmerge(inp_csv: Path, pdf_base: Path, annee: str, out_csv: Path, missing_csv: Path):
    argv = [
        "--in", str(inp_csv),
        "--pdf-base", str(pdf_base),
        "--out", str(out_csv),
        "--annee", annee,
        "--missing", str(missing_csv),
    ]
    print(f"{ICON_PLAY} build_mailmerge (module):", BUILD_MM_MODULE, " ".join(argv))
    try:
        import_generic_module(BUILD_MM_MODULE, "build_mailmerge_4e_from_merged_v5.py")
        run_module_with_argv(BUILD_MM_MODULE, argv)
    except ImportError:
        path = Path(_resource_path("build_mailmerge_4e_from_merged_v5.py"))
        if not path.exists():
            raise SystemExit("build_mailmerge introuvable (module et fichier). Rebuild requis en embarquant build_mailmerge_4e_from_merged_v5.py")
        old_argv = sys.argv[:]
        try:
            sys.argv = [str(path)] + argv
            runpy.run_path(str(path), run_name="__main__")
        finally:
            sys.argv = old_argv

def run_tb(csv_file: Path, sleep: float, limit: int|None, skip: int, dry_run: bool, tb_binary: str|None):
    print(f"[DEBUG] TB lira ce CSV : {csv_file}")
    # Resolve Thunderbird binary robustly (works on Windows and macOS)
    resolved_tb = resolve_tb_binary(tb_binary)
    if not resolved_tb and not dry_run:
        print("[ERREUR] Binaire Thunderbird introuvable. Installe-le (ou indique --tb-binary \"chemin/vers/thunderbird\").")
        sys.exit(1)
    argv = [
        "--csv", str(csv_file),
        "--col-body", "CorpsMessage",
        "--sleep", str(sleep or 0.6),
    ]
    eff_limit = limit if (isinstance(limit, int) and limit >= 0) else 1_000_000
    argv += ["--limit", str(eff_limit)]
    argv += ["--skip", str(skip or 0)]
    if dry_run:
        argv.append("--dry-run")
    # Pass the resolved Thunderbird path when available
    if resolved_tb:
        argv += ["--tb-binary", resolved_tb]
    print(f"{ICON_PLAY} tb_mailmerge (module):", TB_MODULE, " ".join(argv))
    try:
        import_generic_module(TB_MODULE, "tb_mailmerge_mac.py")
        run_module_with_argv(TB_MODULE, argv)
    except ImportError:
        path = Path(_resource_path("tb_mailmerge_mac.py"))
        if not path.exists():
            raise SystemExit("tb_mailmerge introuvable (module et fichier). Rebuild requis en embarquant tb_mailmerge_mac.py")
        old_argv = sys.argv[:]
        try:
            sys.argv = [str(path)] + argv
            runpy.run_path(str(path), run_name="__main__")
        finally:
            sys.argv = old_argv

def ensure_same_year(annee: str):
    if not annee or "-" not in annee:
        raise ValueError(f"Année invalide : {annee}. Exemple attendu: 2025-2026")

def main():
    ap = argparse.ArgumentParser(description="Pipeline EvalNat (orchestration)")
    ap.add_argument("--classe", required=True, help="Ex: 4D")
    ap.add_argument("--annee", required=True, help='Ex: "2025-2026"')
    ap.add_argument("--input-pdf", help="PDF OCR classe (export all)")
    ap.add_argument("--out-dir", required=True, help="Dossier des PDFs par élève (sera scanné à l'étape 3)")
    ap.add_argument("--parents", nargs="+", help="Exports SIECLE à fusionner (4A/4B/4C/4D...)")
    ap.add_argument("--keep-accents", action="store_true", help="Garder accents dans les noms de fichiers split")
    # Options avancées
    ap.add_argument("--no-split", action="store_true", help="Ne pas exécuter l'étape 1 (réutiliser un dossier de PDFs déjà prêt)")
    ap.add_argument("--no-merge", action="store_true", help="Ne pas exécuter l'étape 2 (fournir --csv-in)")
    ap.add_argument("--csv-in", default=None, help="CSV parents déjà prêt (canonisé ou *_mailmerge.csv). Si absent, on utilisera la sortie de merge_parents.")
    ap.add_argument("--csv-tb", default=None, help="CSV à passer explicitement à Thunderbird (par défaut: mailmerge_<classe>_with_emails.csv s'il existe)")
    ap.add_argument("--strict", action="store_true", help="Stopper avant TB si des pièces jointes manquent")
    ap.add_argument("--preflight-threshold", type=float, default=0.8,
                    help="Seuil minimal (0-1) du rapport (PDF présents / effectif classe) par discipline avant le build. Alerte si en-dessous, arrêt si --strict.")
    # OCR options
    ap.add_argument("--auto-ocr", action="store_true",
                    help="Tenter un pré-traitement OCR si le PDF est non-texte/peu-texte")
    ap.add_argument("--ocr-lang", default="fra",
                    help="Langue OCR (codes tesseract/ocrmypdf), ex: fra, fra+eng (défaut: fra)")
    # Thunderbird (optionnel)
    ap.add_argument("--run-tb", action="store_true", help="Ouvrir les brouillons Thunderbird")
    ap.add_argument("--sleep", type=float, default=0.6, help="Pause entre brouillons TB (sec)")
    ap.add_argument("--limit", type=int, default=None, help="Limiter à N lignes lors de l’ouverture TB (par défaut: tous)")
    ap.add_argument("--skip", type=int, default=0, help="Ignorer N lignes au début lors de l’ouverture TB")
    ap.add_argument("--dry-run", action="store_true", help="TB: n’ouvre rien (test)")
    ap.add_argument("--tb-binary", default=None, help="Chemin explicite Thunderbird (ex: Windows: C:\\Program Files\\Mozilla Thunderbird\\thunderbird.exe | macOS: /Applications/Thunderbird.app/Contents/MacOS/thunderbird)")

    # Message commun aux parents
    ap.add_argument("--message-text", default=None, help="Texte du message parents à répliquer dans la colonne 'CorpsMessage' pour toutes les lignes")
    ap.add_argument("--message-file", default=None, help="Fichier texte (UTF-8) contenant le message parents à répliquer dans 'CorpsMessage'")

    args = ap.parse_args()
    classe = args.classe
    annee  = args.annee
    out_dir= Path(args.out_dir)
    progress(0, "Préparation")
    # Pré-définitions robustes (utilisées aussi dans le résumé final)
    cwd = Path.cwd()
    mailmerge_csv = cwd / f"mailmerge_{classe}.csv"

    ensure_same_year(annee)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Préparer le message commun (optionnel)
    message_common: str | None = None
    if args.message_text and args.message_file:
        ap.error("Utiliser soit --message-text soit --message-file, pas les deux.")
    if args.message_text:
        message_common = args.message_text
    elif args.message_file:
        pmsg = Path(args.message_file)
        if not pmsg.exists():
            ap.error(f"--message-file introuvable: {pmsg}")
        message_common = read_text_robust(pmsg)
    if message_common is not None:
        message_common = message_common.replace("\r\n", "\n").replace("\r", "\n").strip("\n")

    # Valider combinaisons d’options
    if not args.no_split and not args.input_pdf:
        ap.error("--input-pdf requis sauf si --no-split est utilisé.")
    if not args.no_merge and not (args.parents and len(args.parents) > 0):
        if not args.csv_in:
            ap.error("--parents requis (ou bien utiliser --no-merge avec --csv-in)")

    in_pdf = Path(args.input_pdf) if args.input_pdf else None
    parents_paths = [Path(p) for p in (args.parents or [])]

    if in_pdf and not in_pdf.exists():
        ap.error(f"--input-pdf introuvable: {in_pdf}")
    for pth in parents_paths:
        if not pth.exists():
            ap.error(f"CSV SIECLE introuvable: {pth}")

    # --- Pré-OCR éventuel --------------------------------------------------------
    ocr_pdf = None
    try:
        ocrh = import_generic_module("ocr_helper", "ocr_helper.py")
        ok_ocr, msg_ocr = ocrh.ensure_ocr_stack(
            force_install=True,
            base_dir=Path(getattr(sys, "_MEIPASS", BASE_DIR)),
            lang=args.ocr_lang,
        )
    except Exception as e:
        ok_ocr, msg_ocr = False, f"OCR: module indisponible ({e})"
    print((ICON_OK + " " if ok_ocr else ICON_WARN + " ") + _safe(msg_ocr))
    if in_pdf and args.auto_ocr:
        ratio = 0.0
        try:
            ratio = quick_text_ratio(in_pdf)
        except Exception:
            ratio = 0.0
        print(f"{ICON_SCAN} Sondage texte PDF: ~{ratio:.0f} caractères/page (échantillon)")
        if ratio < 50:
            if have_cmd("ocrmypdf"):
                ocr_pdf = in_pdf.with_name(in_pdf.stem + "_OCR.pdf")
                try:
                    run_ocrmypdf(in_pdf, ocr_pdf, lang=args.ocr_lang)
                    if ocr_pdf.exists() and ocr_pdf.stat().st_size > 0:
                        print(f"{ICON_OK} OCR ok {ICON_ARROW} {ocr_pdf.name}")
                        in_pdf = ocr_pdf
                    else:
                        print(f"{ICON_WARN} OCR a produit un fichier vide. On conserve le PDF original.")
                except subprocess.CalledProcessError as e:
                    print(f"{ICON_WARN} OCR échoué (ocrmypdf code {e.returncode}).")
            else:
                print(f"{ICON_WARN} OCR non disponible (ocrmypdf introuvable).")
                print(f"   {ICON_ARROW} Solution immédiate : ouvrir le PDF dans Adobe Acrobat et appliquer")
                print(f"     'Reconnaissance de texte (OCR)', puis relancer le pipeline avec --no-split.")
        else:
            print(f"{ICON_OK} PDF semble déjà textuel {ICON_ARROW} pas d’OCR nécessaire.")

    # === Étape 1: split ===
    print("=== Étape 1/4 — Split PDF par élève ===")
    if args.no_split:
        print(f"{ICON_SKIP}  Étape 1 ignorée (--no-split). On réutilise les PDFs présents dans:", out_dir)
    else:
        progress(25, f"Classe={classe} — Split en cours")
        _ensure_meipass_on_syspath()
        split_mod = import_split_module()
        split_mod.CLASS_LABEL = classe
        split_mod.SCHOOL_YEAR = annee
        split_mod.INPUT_PDF   = str(in_pdf)
        split_mod.OUTPUT_DIR  = str(out_dir)
        split_mod.KEEP_ACCENTS_IN_FILENAME = bool(args.keep_accents)
        split_mod.split_pdf()
        progress(25, "OK — split terminé")

    generated = list(out_dir.rglob("*.pdf"))
    if not generated:
        raise SystemExit("Aucun PDF trouvé dans le dossier de sortie. Vérifie --out-dir et l'étape de split.")
    classes_seen, years_seen = scan_pdf_labels(out_dir)
    this_class_pdfs = [pth for pth in generated if pth.name.startswith(f"{classe}_")]
    if not this_class_pdfs:
        msg = []
        msg.append("Aucun PDF ne correspond à la classe demandée.")
        msg.append(f"  {BULLET} Classe demandée : {classe}")
        if classes_seen:
            msg.append(f"  {BULLET} Classes trouvées : {', '.join(sorted(classes_seen))}")
        if years_seen:
            msg.append(f"  {BULLET} Années vues dans les fichiers : {', '.join(sorted(years_seen))}")
        msg.append("Conseils :")
        msg.append("  - Vérifie que --classe correspond bien aux PDFs générés (ex: 6A vs 4D).")
        msg.append("  - Ou relance le split avec la bonne classe/année.")
        raise SystemExit("\n".join(msg))
    other_classes = sorted(c for c in classes_seen if c != classe)
    if other_classes:
        raise SystemExit(
            "Garde anti-mismatch : des PDFs d'une autre classe sont présents dans le dossier de sortie.\n"
            f"  {BULLET} Classe demandée : {classe}\n"
            f"  {BULLET} Autres classes trouvées : {', '.join(other_classes)}\n"
            "  → Utilise un dossier dédié à la classe, ou nettoie les PDFs parasites."
        )
    print(f"→ {len(this_class_pdfs)} PDF(s) de {classe} visibles dans {out_dir}")
    progress(25, f"OK — {len(this_class_pdfs)} PDF(s) détectés")

    print("=== Étape 2/4 — Fusion exports SIECLE (parents) ===")
    cwd = Path.cwd()
    out_std  = cwd / "parents_4e_merged.csv"
    out_mail = cwd / "parents_4e_mailmerge.csv"

    if args.no_merge and args.csv_in:
        parents_csv = Path(args.csv_in)
        if not parents_csv.exists():
            raise SystemExit(f"--csv-in introuvable: {parents_csv}")
        print(f"{ICON_SKIP}  Étape 2 ignorée (--no-merge). CSV parents fourni: {parents_csv}")
    else:
        progress(50, "Fusion des exports SIECLE")
        mp = import_merge_parents_module()
        mp.merge_files([str(pth) for pth in parents_paths])
        if not out_std.exists():
            raise SystemExit("parents_4e_merged.csv non généré (voir logs merge_parents_4e.py).")
        parents_csv = out_std
        print(f"{ICON_OK} Fusion OK → {parents_csv}")
        progress(50, "OK — fusion terminée")

    # Canonisation/filtrage
    canon_csv = cwd / f"parents_{classe}_canon.csv"
    sep = detect_sep(parents_csv)
    kept = 0
    with open_csv_reader(parents_csv) as f, open_csv_writer(canon_csv) as g:
        rdr = csv.DictReader(f, delimiter=sep)
        fields = rdr.fieldnames or []
        w = csv.DictWriter(g, fieldnames=fields, delimiter=sep)
        w.writeheader()
        for r in rdr:
            div = (r.get("Division") or r.get("Classe") or "").strip()
            m = re.match(r'^=\s*"(.+)"\s*$', div)
            if m: div = m.group(1)
            divN = unicodedata.normalize("NFD", div).upper()
            divN = "".join(ch for ch in divN if unicodedata.category(ch) != "Mn")
            divN = divN.replace("ÈME","E").replace("EME","E")
            divN = re.sub(r"[\s\-.]+","", divN)
            if divN == classe.upper():
                r["Division"] = classe
                w.writerow(r); kept += 1

    if kept == 0:
        divs = set()
        with open_csv_reader(parents_csv) as f:
            rdr = csv.DictReader(f, delimiter=sep)
            for r in rdr:
                d = (r.get("Division") or r.get("Classe") or "").strip()
                if d:
                    dN = unicodedata.normalize("NFD", d).upper()
                    dN = "".join(ch for ch in dN if unicodedata.category(ch) != "Mn")
                    dN = dN.replace("ÈME","E").replace("EME","E")
                    dN = re.sub(r"[\s\-.]+","", dN)
                    divs.add(dN)
        hint = f"Divisions présentes dans le CSV: {', '.join(sorted(divs))}" if divs else "Aucune division détectée dans le CSV."
        raise SystemExit(
            "Garde anti-mismatch : aucune ligne de parents ne correspond à la classe demandée.\n"
            f"  {BULLET} Classe demandée : {classe}\n"
            f"  {BULLET} Fichier analysé : {parents_csv}\n"
            f"  {BULLET} {hint}\n"
            "  → Fourni(s) le(s) CSV de la bonne classe ou change --classe pour correspondre au CSV fourni."
        )
    else:
        print(f"→ {kept} lignes 'Division={classe}' dans {canon_csv}")
        progress(50, f"OK — {kept} lignes pour {classe}")

    # --- Préflight -----------------------------------------------------------
    if kept > 0:
        counts = count_pdfs_by_disc(out_dir, classe, annee)
        eff = kept
        ratio_fr = counts["francais"] / eff if eff else 0.0
        ratio_ma = counts["maths"] / eff if eff else 0.0
        print(f"{ICON_SCAN} Préflight: effectif={eff} | PDFs: Français={counts['francais']} ({ratio_fr:.0%}) | Maths={counts['maths']} ({ratio_ma:.0%})")
        alerts = []
        if ratio_fr < args.preflight_threshold:
            alerts.append(f"Français sous le seuil ({ratio_fr:.0%} < {int(args.preflight_threshold*100)}%)")
        if ratio_ma < args.preflight_threshold:
            alerts.append(f"Maths sous le seuil ({ratio_ma:.0%} < {int(args.preflight_threshold*100)}%)")
        if alerts:
            print(f"{ICON_WARN} Préflight: " + " ; ".join(alerts))
            if args.strict:
                raise SystemExit("Arrêt (--strict) : split incomplet détecté par le préflight.")

    print("=== Étape 3/4 — Construction CSV Mail Merge ===")
    progress(75, "Construction du CSV MailMerge")
    # `mailmerge_csv` a déjà été défini plus haut pour éviter NameError
    missing_csv   = cwd / f"missing_{classe}.csv"
    run_build_mailmerge(inp_csv=canon_csv, pdf_base=out_dir, annee=annee,
                        out_csv=mailmerge_csv, missing_csv=missing_csv)

    if missing_csv.exists() and missing_csv.stat().st_size > 0:
        print(f"{ICON_WARN}  Des pièces jointes manquent : {missing_csv}")

    # === Étape 3 bis — Emails + CorpsMessage =================================
    mm_with_emails = mailmerge_csv.with_name(f"mailmerge_{classe}_with_emails.csv")

    emails_index = {}
    sep_canon = detect_sep(canon_csv)
    with open_csv_reader(canon_csv) as f:
        rdr = csv.DictReader(f, delimiter=sep_canon)
        for r in rdr:
            div = (r.get("Division") or r.get("Classe") or "").strip()
            nom = (r.get("Nom de famille") or r.get("Nom") or "").strip()
            pre = (r.get("Prénom 1") or r.get("Prénom") or r.get("Prenom") or "").strip()
            if not (div and nom and pre): 
                continue
            key = (squash(div), squash(nom), squash(pre))
            e1 = (r.get("Courriel repr. légal") or "").strip()
            e2 = (r.get("Courriel autre repr. légal") or "").strip()
            em = "; ".join([e for e in [e1, e2] if e])
            if em:
                emails_index[key] = em

    sep_mm = detect_sep(mailmerge_csv)
    rows = []
    filled = empty = 0
    with open_csv_reader(mailmerge_csv) as f:
        rdr = csv.DictReader(f, delimiter=sep_mm)
        fields = list(rdr.fieldnames or [])
        if "CorpsMessage" not in fields:
            fields.insert(0, "CorpsMessage")
        if "Emails" not in fields:
            fields.insert(0, "Emails")
        if "Objet" not in fields:
            fields.insert(0, "Objet")
        for r in rdr:
            div = (r.get("Classe") or r.get("Division") or "").strip()
            nom = (r.get("Nom") or "").strip()
            pre = (r.get("Prénom") or r.get("Prenom") or "").strip()

            # Clé pour retrouver les emails depuis le CSV canon
            key = (squash(div), squash(nom), squash(pre))

            # Compléter les emails manquants
            if not (r.get("Emails") or "").strip():
                r["Emails"] = emails_index.get(key, "")

            if r["Emails"]:
                filled += 1
            else:
                empty += 1

            # Corps du message : soit commun, soit existant
            if message_common is not None:
                r["CorpsMessage"] = message_common
            else:
                r.setdefault("CorpsMessage", r.get("CorpsMessage", ""))

            # Objet : remplir si vide → "Évaluations nationales - NOM Prénom (Classe)"
            if not (r.get("Objet") or "").strip():
                classe_for_subject = (div or "").strip() or classe
                sujet = f"Évaluations nationales - {nom} {pre} ({classe_for_subject})".strip()
                r["Objet"] = sujet

            rows.append(r)

        with open_csv_writer(mm_with_emails) as g:
            w = csv.DictWriter(g, fieldnames=fields, delimiter=sep_mm)
            w.writeheader(); w.writerows(rows)

        print(f"{ICON_OK} Emails remplis: {filled} | manquants: {empty} → {mm_with_emails}")
        if message_common is not None:
            preview = (message_common[:80] + "…") if len(message_common) > 80 else message_common
            print(f"{ICON_OK} Message commun appliqué à toutes les lignes (Colonne 'CorpsMessage'). Aperçu: {preview!r}")
        progress(75, "OK — CSV prêt pour Thunderbird")

        # --- 3ter : CSV spécial Thunderbird (délimiteur = ',') -----------------------
        mm_for_tb = mm_with_emails.with_name(f"mailmerge_{classe}_for_tb.csv")

        def _write_csv_for_tb(src_path: Path, dst_path: Path):
            # Relit le CSV mailmerge (peut être en ';') et réécrit en ',' pour TB.
            sep_src = detect_sep(src_path)
            with open_csv_reader(src_path) as f_in:
                rdr = csv.DictReader(f_in, delimiter=sep_src)
                fieldnames = list(rdr.fieldnames or [])
                # S'assure que les colonnes clés existent
                for needed in ("Emails", "Objet", "CorpsMessage", "PJ_francais", "PJ_math"):
                    if needed not in fieldnames:
                        fieldnames.append(needed)

                # Écrit en CSV virgule, encodé UTF-8
                with open(dst_path, "w", encoding="utf-8", newline="") as f_out:
                    w = csv.DictWriter(f_out, fieldnames=fieldnames, delimiter=",", quoting=csv.QUOTE_MINIMAL)
                    w.writeheader()
                    for row in rdr:
                        # Normalise la colonne Emails : remplace ';' par ',' (TB s'attend à des virgules)
                        em = (row.get("Emails") or "").strip()
                        em = em.replace(";", ",")
                        # Nettoie des doubles espaces/virgules
                        em = ",".join([e.strip() for e in em.split(",") if e.strip()])
                        row["Emails"] = em
                        # CorpsMessage : garantit des \n Unix
                        if "CorpsMessage" in row and row["CorpsMessage"]:
                            row["CorpsMessage"] = row["CorpsMessage"].replace("\r\n", "\n").replace("\r", "\n")
                        w.writerow(row)

        _write_csv_for_tb(mm_with_emails, mm_for_tb)
        print(f"{ICON_OK} CSV spécial TB écrit → {mm_for_tb}")

        # Vérification des pièces jointes
        missing_pj = []
        for r in rows:
            for col in ("PJ_francais", "PJ_math", "Attachments"):
                pj = r.get(col, "")
                if pj:
                    for path in pj.split(","):
                        pth = Path(path.strip())
                        if path and not pth.exists():
                            missing_pj.append((r.get('Nom','?'), pth.name))
        if missing_pj:
            print(f"{ICON_WARN}  {len(missing_pj)} pièces jointes introuvables (extraits) :")
            for n, f in missing_pj[:5]:
                print("   -", n, "→", f)
            if args.strict:
                raise SystemExit("Arrêt (--strict) : des pièces jointes sont manquantes.")

        csv_for_tb = Path(args.csv_tb) if args.csv_tb else (mm_for_tb if mm_for_tb.exists() else (mm_with_emails if mm_with_emails.exists() else mailmerge_csv))

        # === Étape 4: TB brouillons (optionnel) ===
        if args.run_tb:
            label_os = "Windows" if os.name == "nt" else "macOS"
            print(f"=== Étape 4/4 — Ouverture des brouillons Thunderbird ({label_os}) ===")
            progress(90, "Ouverture des brouillons Thunderbird")
            run_tb(csv_file=csv_for_tb,
                   sleep=args.sleep, limit=args.limit, skip=args.skip,
                   dry_run=args.dry_run, tb_binary=args.tb_binary)
        else:
            print(f"{ICON_INFO}  Étape TB non exécutée (utilise --run-tb pour ouvrir les brouillons).")

        progress(100, "Pipeline terminé")
        print("\n{ok} Pipeline terminé.".format(ok=ICON_OK))
        print(f"   {BULLET} PDFs par élève   : {out_dir}")
        print(f"   {BULLET} Parents fusionné : {out_std}")
        print(f"   {BULLET} CSV Mail Merge   : {mailmerge_csv}")
        if message_common is not None:
            print("   {b} Message commun   : (colonne 'CorpsMessage' remplie)".format(b=BULLET))
        if args.run_tb:
            print(f"   {BULLET} CSV pour TB      : {csv_for_tb}")
            if args.limit is None:
                print("   {b} TB: ouverture de tous les brouillons (pas de limite)".format(b=BULLET))
            else:
                print(f"   {BULLET} TB: limite d'ouverture = {args.limit}")
        # Résumé : être tolérant si `mailmerge_csv` n'est pas visible (vieux binaire, code path interrompue)
        mm_path = Path(mailmerge_csv) if 'mailmerge_csv' in locals() else (Path.cwd() / f"mailmerge_{classe}.csv")
        if mm_path.exists():
            print(f"   {ICON_ARROW} Ouvre ce CSV dans Thunderbird (extension Mail Merge) ou relance le pipeline avec --run-tb.")

if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as e:
        p(f"\n[ERREUR] Commande externe a échoué (code {e.returncode}).")
        sys.exit(e.returncode)
    except Exception as e:
        p(f"\n[ERREUR] {e}")
        sys.exit(1)