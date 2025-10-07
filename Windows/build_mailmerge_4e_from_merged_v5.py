# -*- coding: utf-8 -*-
from __future__ import annotations
import sys, os, csv, re, unicodedata
from pathlib import Path

# ===== Sortie console robuste (Windows/UTF-8) =====
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass
os.environ.setdefault("PYTHONUTF8", "1")
os.environ.setdefault("PYTHONIOENCODING", "UTF-8")

def _safe(s: str) -> str:
    enc = (sys.stdout.encoding or 'utf-8')
    try:
        s.encode(enc)
        return s
    except Exception:
        return s.encode('ascii', 'ignore').decode('ascii')

import builtins as _bi
def print(*args, **kw):
    try:
        _bi.print(*[_safe(str(a)) for a in args], **kw)
    except Exception:
        kw = {k:v for k,v in kw.items() if k != "file"}
        _bi.print(*[str(a).encode("ascii","ignore").decode("ascii") for a in args], **kw)

# ===== Utils encodage/CSV =====
def detect_sep(path: Path) -> str:
    # Détection du séparateur via lecture binaire pour éviter la corruption d'accents
    with open(path, "rb") as f:
        s = f.read(4096)
    sc = s.count(b";")
    cc = s.count(b",")
    return ";" if sc >= cc else ","

def open_csv_reader(path: Path):
    # Lecture robuste avec détection simple d'encodage.
    # 1) On tente UTF-8/UTF-8-SIG.
    # 2) Si échec, on bascule sur CP1252.
    # 3) Si la lecture en CP1252 "réussit" mais contient des séquences de mojibake 'ï¿½',
    #    on retente en UTF-8-SIG.
    raw = Path(path).read_bytes()
    # Heuristique: si décodage utf-8 fonctionne → on l'utilise
    try:
        _ = raw.decode("utf-8")
        return open(path, "r", encoding="utf-8-sig", newline="")
    except UnicodeDecodeError:
        pass
    # Sinon, essaye CP1252
    try:
        txt = raw.decode("cp1252")
        # Si on voit des séquences typiques de mojibake UTF-8 lues en CP1252, reprends en UTF-8-SIG
        if "ï¿½" in txt or "Ã©" in txt or "Ã¨" in txt or "Ãª" in txt or "Ã«" in txt or "Ã" in txt:
            try:
                _ = raw.decode("utf-8")
                return open(path, "r", encoding="utf-8-sig", newline="")
            except UnicodeDecodeError:
                pass
        return open(path, "r", encoding="cp1252", newline="")
    except UnicodeDecodeError:
        # Dernier secours
        return open(path, "r", encoding="utf-8-sig", newline="")

def open_csv_writer(path: Path):
    # on écrit toujours en UTF-8
    return open(path, "w", encoding="utf-8", errors="replace", newline="")

def nfd(s: str) -> str:
    return unicodedata.normalize("NFD", s or "")

def strip_accents_lower(s: str) -> str:
    s = nfd(s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.lower()

def squash_key(*parts: str) -> str:
    raw = "_".join(p or "" for p in parts)
    s = strip_accents_lower(raw)
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s

def canon_div(s: str) -> str:
    s = (s or "").strip()
    m = re.match(r'^=\s*"(.+)"\s*$', s)
    if m:
        s = m.group(1)
    sN = nfd(s).upper()
    sN = "".join(ch for ch in sN if unicodedata.category(ch) != "Mn")
    sN = sN.replace("ÈME","E").replace("EME","E")
    sN = re.sub(r"[\s\-.]+","", sN)
    return sN

def _normalize_header_key(h: str) -> str:
    """
    Normalise une entête:
    - retire BOM/espaces insécables
    - enlève accents
    - met en minuscules
    - supprime espaces/ponctuation
    Gère aussi des cas de mojibake ("Prï¿½nom" etc.).
    """
    h = (h or "").strip().strip("\ufeff").replace("\xa0", " ")
    # Réparation de quelques séquences de mojibake fréquentes
    h = (h
         .replace("Prï¿½nom", "Prénom")
         .replace("prï¿½nom", "prénom")
         .replace("PrÃ©nom", "Prénom")
         .replace("prÃ©nom", "prénom")
         .replace("repr. lï¿½gal", "repr. légal")
         .replace("lÃ©gal", "légal")
    )
    h = strip_accents_lower(h)
    h = re.sub(r"[^a-z0-9]+", "", h)
    return h

def _resolve_columns(fieldnames: list[str]) -> tuple[str|None, str|None, str|None]:
    """
    Retourne un triplet (col_div, col_nom, col_prenom) en cherchant dans les entêtes
    des variantes connues et robustes (sans accents/espaces). Tolère les cas de mojibake.
    """
    token_by_orig = {fn: _normalize_header_key(fn) for fn in (fieldnames or [])}

    candidates_div   = {"division", "classe", "classeeleve"}
    candidates_nom   = {"nom", "nomdefamille", "nom1", "nomusage", "nomdenaissance"}
    # Tolère erreurs d'encodage: "prenom" → "prnom" (é manquant)
    candidates_pren  = {"prenom", "prenom1", "prenomeleve", "prenompremier", "prnom", "prnom1"}

    col_div = col_nom = col_pre = None
    for orig, tok in token_by_orig.items():
        if col_div is None and tok in candidates_div:
            col_div = orig
            continue
        if col_nom is None and (tok in candidates_nom or tok.startswith("nomde")):
            col_nom = orig
            continue
        if col_pre is None and (tok in candidates_pren or re.search(r"^pr?en?om\d*$", tok)):
            col_pre = orig
            continue

    # Fallback ultime: cherche un token contenant "prenom" même partiellement
    if col_pre is None:
        for orig, tok in token_by_orig.items():
            if "prenom" in tok or "prnom" in tok:
                col_pre = orig
                break

    return col_div, col_nom, col_pre

# ===== Normalisation des entêtes variables =====
def _get_division(r: dict) -> str:
    # variantes possibles selon export SIECLE ; on renvoie la version canonique (ex: "6 A" -> "6A", '="6 A"' -> "6A")
    raw = (
        r.get("Division") or r.get("division") or
        r.get("Classe") or r.get("classe") or
        ""
    )
    return canon_div(str(raw))

def _get_nom(r: dict) -> str:
    # variantes possibles selon export SIECLE
    return (
        r.get("Nom") or r.get("nom") or
        r.get("Nom de famille") or r.get("NOM de famille") or r.get("Nom 1") or
        r.get("Nom de naissance") or r.get("Nom Usage") or r.get("Nom d'usage") or
        ""
    ).strip()

def _get_prenom(r: dict) -> str:
    # variantes possibles selon export SIECLE
    return (
        r.get("Prénom") or r.get("Prenom") or r.get("prénom") or r.get("prenom") or
        r.get("Prénom 1") or r.get("Prenom 1") or
        r.get("Prénom de l'élève") or r.get("Prenom de l'eleve") or
        ""
    ).strip()

# ===== Cœur : build mailmerge =====
def build_mailmerge(inp_csv: Path, pdf_base: Path, annee: str,
                    out_csv: Path, missing_csv: Path) -> None:

    if not inp_csv.exists():
        raise SystemExit(f"CSV d'entrée introuvable: {inp_csv}")
    if not pdf_base.exists():
        raise SystemExit(f"Dossier PDF introuvable: {pdf_base}")

    sep = detect_sep(inp_csv)
    print("Lecture parents:", inp_csv, f"(sep='{sep}')")

    # Petit garde-fou encodage/locale (Windows)
    try:
        _ = str(inp_csv).encode(sys.stdout.encoding or "utf-8", "strict")
    except Exception:
        pass

    # Index des parents: key=(classe,nom,prenom) « souple »
    parents_rows = []
    with open_csv_reader(inp_csv) as f:
        rdr = csv.DictReader(f, delimiter=sep)
        # DEBUG: afficher les entêtes détectées + colonnes résolues
        try:
            print("Entêtes CSV détectées:", rdr.fieldnames)
        except Exception:
            pass
        col_div, col_nom, col_pre = _resolve_columns(rdr.fieldnames or [])
        print(f"Colonnes détectées → Division='{col_div}' | Nom='{col_nom}' | Prénom='{col_pre}'")
        for r in rdr:
            # Normalisation des clés et valeurs pour gérer BOM/espaces insécables etc.
            r = {
                (k or "").strip().strip("\ufeff").replace("\xa0", " "):
                    (v or "").strip()
                for k, v in (r.items() if isinstance(r, dict) else [])
            }
            # Lire depuis les colonnes réellement présentes
            raw_div = (r.get(col_div or "", "") if col_div else "")
            raw_nom = (r.get(col_nom or "", "") if col_nom else "")
            raw_pre = (r.get(col_pre or "", "") if col_pre else "")

            div = canon_div(str(raw_div))
            nom = str(raw_nom).strip()
            pre = str(raw_pre).strip()
            if not (div and nom and pre):
                continue
            parents_rows.append({
                "Division": div,
                "Nom": nom,
                "Prénom": pre,
            })

    if not parents_rows:
        # Aide au debug: réafficher les entêtes et quelques premières lignes du CSV
        try:
            with open_csv_reader(inp_csv) as fdbg:
                fdbg.seek(0)
                preview = "".join([next(fdbg) for _ in range(5)])
        except Exception:
            preview = "(aperçu indisponible)"
        raise SystemExit(
            "Aucune ligne exploitable dans le CSV parents (Division/Nom/Prénom manquants).\n"
            f"→ Vérifiez les entêtes (Division/Nom/Prénom) et le séparateur. Aperçu du fichier :\n{preview}\n"
            f"(Colonnes résolues: Division='{col_div}' | Nom='{col_nom}' | Prénom='{col_pre}')"
        )

    # Scanner les PDFs
    # Format attendu: {classe}_{NOM}_{Prénom}_{Discipline}_{annee}.pdf
    # Ex: 6A_AUPRETRE_Leila_Français_2025-2026.pdf
    pdf_fr_map = {}   # key -> path
    pdf_ma_map = {}   # key -> path

    for p in pdf_base.rglob("*.pdf"):
        name = p.name[:-4]  # sans .pdf
        parts = name.split("_")
        if len(parts) < 5:
            continue
        classe = parts[0].strip()
        nom    = parts[1].strip()
        prenom = parts[2].strip()
        disc   = parts[-2].strip()
        an     = parts[-1].strip()
        if an != annee:
            continue
        key = squash_key(classe, nom, prenom)
        discN = strip_accents_lower(disc)
        if discN.startswith("francais") or discN.startswith("fran") or discN.startswith("fr"):
            pdf_fr_map[key] = str(p)
        elif discN.startswith("mathematiques") or discN.startswith("math") or discN.startswith("mat"):
            pdf_ma_map[key] = str(p)

    print(f"PDF scannés dans {pdf_base} : Français={len(pdf_fr_map)} | Maths={len(pdf_ma_map)}")

    # Composer lignes mailmerge
    out_rows = []
    miss_rows = []
    for r in parents_rows:
        classe = r["Division"]
        nom    = r["Nom"]
        pre    = r["Prénom"]
        key    = squash_key(classe, nom, pre)

        pj_fr = pdf_fr_map.get(key, "")
        pj_ma = pdf_ma_map.get(key, "")
        attaches = ", ".join([p for p in (pj_fr, pj_ma) if p])

        # ligne mailmerge (TB utilisera Emails + CorpsMessage ajoutés par le pipeline)
        out_rows.append({
            "Classe": classe,
            "Nom": nom,
            "Prénom": pre,
            "PJ_francais": pj_fr,
            "PJ_math": pj_ma,
            "Attachments": attaches
        })

        if not pj_fr or not pj_ma:
            miss_rows.append({
                "Classe": classe,
                "Nom": nom,
                "Prénom": pre,
                "ManqueFrançais": "oui" if not pj_fr else "",
                "ManqueMaths": "oui" if not pj_ma else ""
            })

    # Écriture UTF-8 (sep=';')
    with open_csv_writer(out_csv) as g:
        cols = ["Classe","Nom","Prénom","PJ_francais","PJ_math","Attachments"]
        w = csv.DictWriter(g, fieldnames=cols, delimiter=";")
        w.writeheader()
        w.writerows(out_rows)
    print("Ecrit :", out_csv)

    with open_csv_writer(missing_csv) as g:
        cols = ["Classe","Nom","Prénom","ManqueFrançais","ManqueMaths"]
        w = csv.DictWriter(g, fieldnames=cols, delimiter=";")
        w.writeheader()
        w.writerows(miss_rows)
    print("Ecrit :", missing_csv)

# ===== CLI =====
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Build mailmerge CSV à partir des parents + PDFs")
    ap.add_argument("--in", dest="inp", required=True, help="CSV parents canonisé (Division/Nom/Prénom)")
    ap.add_argument("--pdf-base", required=True, help="Dossier contenant les PDFs par élève")
    ap.add_argument("--out", required=True, help="Chemin de sortie du CSV mailmerge")
    ap.add_argument("--annee", required=True, help='Ex: "2025-2026"')
    ap.add_argument("--missing", required=True, help="CSV listant les pièces jointes manquantes")
    args = ap.parse_args()

    build_mailmerge(
        inp_csv=Path(args.inp),
        pdf_base=Path(args.pdf_base),
        annee=args.annee,
        out_csv=Path(args.out),
        missing_csv=Path(args.missing),
    )