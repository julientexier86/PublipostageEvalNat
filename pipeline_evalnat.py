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

import argparse, sys, os, subprocess, shutil, csv, re, unicodedata, runpy, importlib, textwrap
# --- Forcer l'inclusion de pandas dans le binaire (utilisé par merge_parents_4e) ---
# Lorsque merge_parents_4e.py est chargé depuis un fichier embarqué (--add-data),
# PyInstaller n'analyse pas ses imports → pandas ne serait pas collecté.
# On fait un import "leurre" ici pour que PyInstaller embarque pandas (et ses deps)
# dans l'exécutable onefile `evalnat-pipeline`.
try:
    import pandas  # type: ignore  # noqa: F401
except Exception:
    # En exécution non gelée sans pandas installé, on laisse merge_parents_4e
    # gérer l'erreur. En exécution gelée, pandas sera présent.
    pass
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent  # dossier contenant le pipeline

# --- Helpers pour exécution en mode gelé (PyInstaller onefile) -------------
def _resource_path(rel: str = "") -> str:
    """
    Retourne le chemin absolu d'une ressource embarquée lorsque le script est gelé.
    En mode normal, revient sur le dossier du fichier courant.
    """
    base = getattr(sys, "_MEIPASS", str(BASE_DIR))
    return str(Path(base) / rel)

def _ensure_meipass_on_syspath():
    """
    Ajoute le dossier d'extraction PyInstaller (_MEIPASS) au sys.path si nécessaire
    pour pouvoir importer des modules embarqués (ex: split_4C.py).
    """
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass and meipass not in sys.path:
        sys.path.insert(0, meipass)

# --- Helper générique pour importer un module embarqué ou via fichier -------
def import_generic_module(modname: str, filename: str):
    """
    Importe un module potentiellement embarqué en onefile.
    Stratégie:
      1) import normal
      2) ajouter _MEIPASS au sys.path et retenter
      3) charger explicitement depuis un fichier embarqué via --add-data
    `filename` est le nom du fichier embarqué (ex: 'merge_parents_4e.py').
    """
    # 1) import direct si dispo
    try:
        return importlib.import_module(modname)
    except Exception as e_first:
        # 2) tenter depuis _MEIPASS
        _ensure_meipass_on_syspath()
        try:
            return importlib.import_module(modname)
        except Exception:
            # 3) chargement explicite depuis un chemin fichier
            import importlib.util
            path = Path(_resource_path(filename))
            if path.exists():
                spec = importlib.util.spec_from_file_location(modname, str(path))
                if spec and spec.loader:
                    mod = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(mod)  # type: ignore
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
    """
    Exécute un module comme s'il était appelé en ligne de commande.
    Simule: python -m modname <argv...>
    """
    old_argv = sys.argv[:]
    try:
        sys.argv = [modname] + argv
        runpy.run_module(modname, run_name="__main__")
        return 0
    finally:
        sys.argv = old_argv

# --- Helpers ---------------------------------------------------------------
def detect_sep(p: Path) -> str:
    with open(p, "r", encoding="utf-8-sig", newline="") as f:
        s = f.read(4096)
    return ";" if s.count(";") >= s.count(",") else ","

def nfd(s: str) -> str:
    return unicodedata.normalize("NFD", s or "")

def squash(s: str) -> str:
    s = nfd(s).lower()
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"[^a-z0-9]+", "", s)

def count_pdfs_by_disc(base: Path, classe: str, annee: str) -> dict:
    """
    Retourne un dict {'francais': n, 'maths': n} en scannant base.
    Tolère les variantes 'Français/Francais/Franais' & 'Mathématiques/Mathematiques/Mathmatiques/Maths'.
    """
    discs_fr = ("Français","Francais","Franais")
    discs_ma = ("Mathématiques","Mathematiques","Mathmatiques","Maths")
    n_fr = n_ma = 0
    for p in base.glob(f"{classe}_*_{annee}.pdf"):
        name = p.name
        # on regarde l'avant-dernier segment (discipline) si possible
        parts = name[:-4].split("_")  # drop .pdf
        if len(parts) >= 4:
            disc = parts[-2]
            if disc in discs_fr: n_fr += 1
            if disc in discs_ma: n_ma += 1
    return {"francais": n_fr, "maths": n_ma}

# --- Helper pour inspecter les classes/années présentes dans les PDFs générés
def scan_pdf_labels(base: Path) -> tuple[set[str], set[str]]:
    """
    Analyse les noms de fichiers dans 'base' pour en extraire:
      - les classes vues (prefixe avant premier '_')
      - les années vues (dernier segment AAAA-AAAA)
    Retourne (classes, annees).
    """
    classes: set[str] = set()
    annees: set[str] = set()
    for p in base.glob("*.pdf"):
        name = p.name[:-4]  # sans .pdf
        parts = name.split("_")
        if len(parts) >= 2:
            classes.add(parts[0])
        # année attendue en dernier segment
        if len(parts) >= 2:
            annees.add(parts[-1])
    return classes, annees

# --- OCR helpers -----------------------------------------------------------
def quick_text_ratio(pdf_path: Path, max_pages: int = 6) -> float:
    """
    Renvoie un ratio 'caractères extraits / pages examinées'.
    Si ~0 → PDF probablement scanné (pas de texte).
    Essaie pdfminer; à défaut, PyPDF2.
    """
    # Tentative avec pdfminer
    try:
        from pdfminer.high_level import extract_text  # type: ignore
        from pdfminer.pdfpage import PDFPage  # type: ignore
        chars = 0
        pages = 0
        # On lit rapidement le texte global puis on approxime par page vue
        txt = extract_text(str(pdf_path)) or ""
        # Compter un petit nombre de pages pour l'échantillon
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
    # Fallback PyPDF2
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
    """
    OCR avec ocrmypdf si installé.
    - --force-ocr : force l'OCR même si un peu de texte est détecté (scan partiel)
    - --rotate-pages / --deskew : oriente et redresse
    - --clean-final : nettoie le fond
    - --skip-text : skip si déjà texte (optimisation)
    """
    cmd = ["ocrmypdf", "--force-ocr", "--rotate-pages", "--deskew",
           "--clean-final", "--skip-text", f"--language={lang}",
           str(src), str(dst)]
    print("▶ ocrmypdf:", " ".join(cmd))
    subprocess.check_call(cmd)

def import_split_module():
    """
    Importe le module de découpage, y compris lorsqu'on est packagé en onefile.
    Ordre d'essai:
      1) import normal (module présent dans l'environnement)
      2) import depuis le répertoire d'extraction PyInstaller (_MEIPASS)
         où le fichier 'split_4C.py' peut avoir été embarqué via --add-data.
    """
    # 1) import direct si disponible
    try:
        return importlib.import_module(SPLIT_MODULE)
    except Exception as e_first:
        # 2) tenter depuis _MEIPASS
        _ensure_meipass_on_syspath()
        try:
            return importlib.import_module(SPLIT_MODULE)
        except Exception:
            # 3) chargement explicite depuis un chemin fichier
            import importlib.util
            split_path = Path(_resource_path("split_4C.py"))
            if split_path.exists():
                spec = importlib.util.spec_from_file_location(SPLIT_MODULE, str(split_path))
                if spec and spec.loader:
                    mod = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(mod)  # type: ignore
                    return mod
            # si on arrive ici, le module est vraiment introuvable
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
    print("▶ build_mailmerge (module):", BUILD_MM_MODULE, " ".join(argv))
    # Essayer import module; sinon exécuter le fichier embarqué
    try:
        import_generic_module(BUILD_MM_MODULE, "build_mailmerge_4e_from_merged_v5.py")
        run_module_with_argv(BUILD_MM_MODULE, argv)
    except ImportError:
        # Fallback: exécution par chemin fichier
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
    argv = [
        "--csv", str(csv_file),
        "--col-body", "CorpsMessage",
        "--sleep", str(sleep or 0.6),
    ]
    # Par défaut, TB ouvrait parfois seulement ~10 brouillons (valeur par défaut interne).
    # On force un "illimité" raisonnable quand limit n'est pas spécifié.
    eff_limit = limit if (isinstance(limit, int) and limit >= 0) else 1_000_000
    argv += ["--limit", str(eff_limit)]
    argv += ["--skip", str(skip or 0)]
    if dry_run:
        argv.append("--dry-run")
    if tb_binary:
        argv += ["--tb-binary", tb_binary]
    print("▶ tb_mailmerge (module):", TB_MODULE, " ".join(argv))
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
    # Format attendu "YYYY-YYYY"
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
    ap.add_argument("--tb-binary", default=None, help="Chemin explicite TB (macOS)")

    # Message commun aux parents
    ap.add_argument("--message-text", default=None, help="Texte du message parents à répliquer dans la colonne 'CorpsMessage' pour toutes les lignes")
    ap.add_argument("--message-file", default=None, help="Fichier texte (UTF-8) contenant le message parents à répliquer dans 'CorpsMessage'")

    args = ap.parse_args()
    classe = args.classe
    annee  = args.annee
    out_dir= Path(args.out_dir)

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
        message_common = pmsg.read_text(encoding="utf-8")
    # Normaliser (supprimer BOM/retours Windows)
    if message_common is not None:
        message_common = message_common.replace("\r\n", "\n").replace("\r", "\n").strip("\n")

    # Valider combinaisons d'options
    if not args.no_split and not args.input_pdf:
        ap.error("--input-pdf requis sauf si --no-split est utilisé.")
    if not args.no_merge and not (args.parents and len(args.parents) > 0):
        if not args.csv_in:
            ap.error("--parents requis (ou bien utiliser --no-merge avec --csv-in)")

    in_pdf = Path(args.input_pdf) if args.input_pdf else None
    parents_paths = [Path(p) for p in (args.parents or [])]

    if in_pdf and not in_pdf.exists():
        ap.error(f"--input-pdf introuvable: {in_pdf}")
    for p in parents_paths:
        if not p.exists():
            ap.error(f"CSV SIECLE introuvable: {p}")

    # --- Pré-OCR éventuel --------------------------------------------------------
    ocr_pdf = None
    if in_pdf and args.auto_ocr:
        ratio = 0.0
        try:
            ratio = quick_text_ratio(in_pdf)
        except Exception:
            ratio = 0.0
        print(f"🔎 Sondage texte PDF: ~{ratio:.0f} caractères/page (échantillon)")
        # Seuil empirique: < 50 caractères/page → très probablement scanné
        if ratio < 50:
            if have_cmd("ocrmypdf"):
                ocr_pdf = in_pdf.with_name(in_pdf.stem + "_OCR.pdf")
                try:
                    run_ocrmypdf(in_pdf, ocr_pdf, lang=args.ocr_lang)
                    if ocr_pdf.exists() and ocr_pdf.stat().st_size > 0:
                        print(f"✅ OCR ok → {ocr_pdf.name}")
                        in_pdf = ocr_pdf  # on bascule le split sur ce fichier OCRisé
                    else:
                        print("⚠️ OCR a produit un fichier vide. On conserve le PDF original.")
                except subprocess.CalledProcessError as e:
                    print(f"⚠️ OCR échoué (ocrmypdf code {e.returncode}).")
            else:
                print("⚠️ OCR non disponible (ocrmypdf introuvable).")
                print("   → Solution immédiate : ouvrir le PDF dans Adobe Acrobat et appliquer")
                print("     'Reconnaissance de texte (OCR)', puis relancer le pipeline avec --no-split.")
        else:
            print("✅ PDF semble déjà textuel → pas d’OCR nécessaire.")

    # === Étape 1: split (importer le module et surcharger ses constantes) ===
    print("=== Étape 1/4 — Split PDF par élève ===")
    if args.no_split:
        print("⏭️  Étape 1 ignorée (--no-split). On réutilise les PDFs présents dans:", out_dir)
    else:
        # rendre le répertoire d'extraction PyInstaller importable si besoin
        _ensure_meipass_on_syspath()
        split_mod = import_split_module()
        split_mod.CLASS_LABEL = classe
        split_mod.SCHOOL_YEAR = annee
        split_mod.INPUT_PDF   = str(in_pdf)
        split_mod.OUTPUT_DIR  = str(out_dir)
        split_mod.KEEP_ACCENTS_IN_FILENAME = bool(args.keep_accents)
        split_mod.split_pdf()

    # Vérif: on doit avoir des PDFs pour la classe, et uniquement cette classe
    generated = list(out_dir.rglob("*.pdf"))
    if not generated:
        raise SystemExit("Aucun PDF trouvé dans le dossier de sortie. Vérifie --out-dir et l'étape de split.")
    # classes/années réellement vues
    classes_seen, years_seen = scan_pdf_labels(out_dir)
    this_class_pdfs = [p for p in generated if p.name.startswith(f"{classe}_")]
    if not this_class_pdfs:
        msg = []
        msg.append("Aucun PDF ne correspond à la classe demandée.")
        msg.append(f"  • Classe demandée : {classe}")
        if classes_seen:
            msg.append(f"  • Classes trouvées : {', '.join(sorted(classes_seen))}")
        if years_seen:
            msg.append(f"  • Années vues dans les fichiers : {', '.join(sorted(years_seen))}")
        msg.append("Conseils :")
        msg.append("  - Vérifie que --classe correspond bien aux PDFs générés (ex: 6A vs 4D).")
        msg.append("  - Ou relance le split avec la bonne classe/année.")
        raise SystemExit("\n".join(msg))
    # si d'autres classes sont présentes, on bloque pour éviter mélange
    other_classes = sorted(c for c in classes_seen if c != classe)
    if other_classes:
        raise SystemExit(
            "Garde anti-mismatch : des PDFs d'une autre classe sont présents dans le dossier de sortie.\n"
            f"  • Classe demandée : {classe}\n"
            f"  • Autres classes trouvées : {', '.join(other_classes)}\n"
            "  → Utilise un dossier dédié à la classe, ou nettoie les PDFs parasites."
        )
    print(f"→ {len(this_class_pdfs)} PDF(s) de {classe} visibles dans {out_dir}")

    # === Étape 2/4 — Fusion exports SIECLE (parents) ===
    print("=== Étape 2/4 — Fusion exports SIECLE (parents) ===")
    cwd = Path.cwd()
    out_std  = cwd / "parents_4e_merged.csv"
    out_mail = cwd / "parents_4e_mailmerge.csv"

    if args.no_merge and args.csv_in:
        parents_csv = Path(args.csv_in)
        if not parents_csv.exists():
            raise SystemExit(f"--csv-in introuvable: {parents_csv}")
        print(f"⏭️  Étape 2 ignorée (--no-merge). CSV parents fourni: {parents_csv}")
    else:
        mp = import_merge_parents_module()
        mp.merge_files([str(p) for p in parents_paths])
        if not out_std.exists():
            raise SystemExit("parents_4e_merged.csv non généré (voir logs merge_parents_4e.py).")
        parents_csv = out_std
        print(f"✅ Fusion OK → {parents_csv}")

    # Canoniser/filtrer le CSV parents pour la classe
    canon_csv = cwd / f"parents_{classe}_canon.csv"
    sep = detect_sep(parents_csv)
    kept = 0
    with open(parents_csv, "r", encoding="utf-8-sig", newline="") as f, \
         open(canon_csv, "w", encoding="utf-8", newline="") as g:
        rdr = csv.DictReader(f, delimiter=sep)
        fields = rdr.fieldnames or []
        # on garde toutes les colonnes + on normalise Division sur '4D' etc.
        w = csv.DictWriter(g, fieldnames=fields, delimiter=sep)
        w.writeheader()
        for r in rdr:
            div = (r.get("Division") or r.get("Classe") or "").strip()
            # normalisation "4 D"/="4 D"/4ème D → 4D
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
        # garde anti-mismatch: on s'arrête ici avec un message explicite
        # pour éviter de construire un CSV vide et de perdre du temps
        # (ex: CSV 6A fourni alors que --classe=4D)
        # Diagnostic: lister les divisions distinctes vues dans le CSV source
        divs = set()
        with open(parents_csv, "r", encoding="utf-8-sig", newline="") as f:
            rdr = csv.DictReader(f, delimiter=sep)
            for r in rdr:
                d = (r.get("Division") or r.get("Classe") or "").strip()
                if d:
                    # normaliser rapidement pour indiquer quelles valeurs existent
                    dN = unicodedata.normalize("NFD", d).upper()
                    dN = "".join(ch for ch in dN if unicodedata.category(ch) != "Mn")
                    dN = dN.replace("ÈME","E").replace("EME","E")
                    dN = re.sub(r"[\s\-.]+","", dN)
                    divs.add(dN)
        hint = f"Divisions présentes dans le CSV: {', '.join(sorted(divs))}" if divs else "Aucune division détectée dans le CSV."
        raise SystemExit(
            "Garde anti-mismatch : aucune ligne de parents ne correspond à la classe demandée.\n"
            f"  • Classe demandée : {classe}\n"
            f"  • Fichier analysé : {parents_csv}\n"
            f"  • {hint}\n"
            "  → Fourni(s) le(s) CSV de la bonne classe ou change --classe pour correspondre au CSV fourni."
        )
    else:
        print(f"→ {kept} lignes 'Division={classe}' dans {canon_csv}")

    # --- Préflight de comptage (avant build) -----------------------------------
    if kept > 0:
        counts = count_pdfs_by_disc(out_dir, classe, annee)
        eff = kept  # effectif attendu pour la classe (lignes parents)
        ratio_fr = counts["francais"] / eff if eff else 0.0
        ratio_ma = counts["maths"] / eff if eff else 0.0
        print(f"🔎 Préflight: effectif={eff} | PDFs: Français={counts['francais']} ({ratio_fr:.0%}) | Maths={counts['maths']} ({ratio_ma:.0%})")
        alerts = []
        if ratio_fr < args.preflight_threshold:
            alerts.append(f"Français sous le seuil ({ratio_fr:.0%} &lt; {int(args.preflight_threshold*100)}%)")
        if ratio_ma < args.preflight_threshold:
            alerts.append(f"Maths sous le seuil ({ratio_ma:.0%} &lt; {int(args.preflight_threshold*100)}%)")
        if alerts:
            print("⚠️  Préflight:", " ; ".join(alerts))
            if args.strict:
                raise SystemExit("Arrêt (--strict) : split incomplet détecté par le préflight.")

    print("=== Étape 3/4 — Construction CSV Mail Merge ===")
    mailmerge_csv = cwd / f"mailmerge_{classe}.csv"
    missing_csv   = cwd / f"missing_{classe}.csv"
    run_build_mailmerge(inp_csv=canon_csv, pdf_base=out_dir, annee=annee,
                        out_csv=mailmerge_csv, missing_csv=missing_csv)

    # Rappel utilisateur si manquants
    if missing_csv.exists() and missing_csv.stat().st_size > 0:
        print(f"⚠️  Des pièces jointes manquent : {missing_csv}")

    # === Étape 3 bis — Compléter Emails et vérifier PJ ========================
    mm_with_emails = mailmerge_csv.with_name(f"mailmerge_{classe}_with_emails.csv")

    # Construire un index Emails depuis canon_csv
    emails_index = {}
    sep_canon = detect_sep(canon_csv)
    with open(canon_csv, "r", encoding="utf-8-sig", newline="") as f:
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
    with open(mailmerge_csv, "r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f, delimiter=sep_mm)
        fields = list(rdr.fieldnames or [])
        # S'assurer que les colonnes 'CorpsMessage' et 'Emails' existent
        if "CorpsMessage" not in fields:
            fields.insert(0, "CorpsMessage")
        if "Emails" not in fields:
            fields.insert(0, "Emails")
        for r in rdr:
            div = (r.get("Classe") or r.get("Division") or "").strip()
            nom = (r.get("Nom") or "").strip()
            pre = (r.get("Prénom") or r.get("Prenom") or "").strip()
            key = (squash(div), squash(nom), squash(pre))
            if not (r.get("Emails") or "").strip():
                r["Emails"] = emails_index.get(key, "")
            if r["Emails"]: filled += 1
            else:           empty  += 1
            # Injecter le message commun si demandé
            if message_common is not None:
                r["CorpsMessage"] = message_common
            else:
                # S'assurer que la clé existe (même vide) pour la compatibilité TB
                r.setdefault("CorpsMessage", r.get("CorpsMessage", ""))
            rows.append(r)

    with open(mm_with_emails, "w", encoding="utf-8", newline="") as g:
        w = csv.DictWriter(g, fieldnames=fields, delimiter=sep_mm)
        w.writeheader(); w.writerows(rows)

    print(f"✅ Emails remplis: {filled} | manquants: {empty} → {mm_with_emails}")
    if message_common is not None:
        preview = (message_common[:80] + "…") if len(message_common) > 80 else message_common
        print(f"✅ Message commun appliqué à toutes les lignes (Colonne 'CorpsMessage'). Aperçu: {preview!r}")

    # Vérification des pièces jointes
    missing_pj = []
    for r in rows:
        for col in ("PJ_francais", "PJ_math", "Attachments"):
            pj = r.get(col, "")
            if pj:
                for path in pj.split(","):
                    p = Path(path.strip())
                    if path and not p.exists():
                        missing_pj.append((r.get('Nom','?'), p.name))
    if missing_pj:
        print(f"⚠️  {len(missing_pj)} pièces jointes introuvables (extraits) :")
        for n, f in missing_pj[:5]:
            print("   -", n, "→", f)
        if args.strict:
            raise SystemExit("Arrêt (--strict) : des pièces jointes sont manquantes.")

    # Choisir le CSV pour TB
    csv_for_tb = Path(args.csv_tb) if args.csv_tb else (mm_with_emails if mm_with_emails.exists() else mailmerge_csv)

    # === Étape 4: TB brouillons (optionnel) ===
    if args.run_tb:
        print("=== Étape 4/4 — Ouverture des brouillons Thunderbird (macOS) ===")
        run_tb(csv_file=csv_for_tb,
               sleep=args.sleep, limit=args.limit, skip=args.skip,
               dry_run=args.dry_run, tb_binary=args.tb_binary)
    else:
        print("ℹ️  Étape TB non exécutée (utilise --run-tb pour ouvrir les brouillons).")

    print("\n✅ Pipeline terminé.")
    print(f"   • PDFs par élève   : {out_dir}")
    print(f"   • Parents fusionné : {out_std}")
    print(f"   • CSV Mail Merge   : {mailmerge_csv}")
    if message_common is not None:
        print("   • Message commun   : (colonne 'CorpsMessage' remplie)")
    if args.run_tb:
        print(f"   • CSV pour TB      : {csv_for_tb}")
        if args.limit is None:
            print("   • TB: ouverture de tous les brouillons (pas de limite)")
        else:
            print(f"   • TB: limite d'ouverture = {args.limit}")
    if (Path(mailmerge_csv).exists()):
        print("   → Ouvre ce CSV dans Thunderbird (extension Mail Merge) ou lance le pipeline avec --run-tb.")

if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as e:
        print(f"\n[ERREUR] Commande externe a échoué (code {e.returncode}).", file=sys.stderr)
        sys.exit(e.returncode)
    except Exception as e:
        print(f"\n[ERREUR] {e}", file=sys.stderr)
        sys.exit(1)