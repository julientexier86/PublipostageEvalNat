python3 - <<'PY'
from pathlib import Path
import csv, re, unicodedata

src = Path("/Users/julien/Downloads/Scripts qui semblent OK/parents_4e_merged.csv")
dst = src.with_name("parents_4e_merged_norm.csv")

def normalize_div(s: str) -> str:
    s = (s or "").strip()
    # Enlève l'enrobage Excel ="4 D"
    m = re.match(r'^=\s*"(.+)"\s*$', s)
    if m: s = m.group(1)
    # Unicode -> supprime diacritiques ; majuscules
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn").upper()
    # Uniformise : retire espaces/tirets/points "4 ÈME D" -> "4EMED"
    s = s.replace("ÈME","E").replace("EME","E")
    s = re.sub(r"[\s\-.]+","", s)
    return s

with open(src, "r", encoding="utf-8-sig", newline="") as f:
    sample = f.read(4096)
sep = ";" if sample.count(";") >= sample.count(",") else ","

with open(src, "r", encoding="utf-8-sig", newline="") as f, \
     open(dst, "w", encoding="utf-8", newline="") as g:
    rdr = csv.DictReader(f, delimiter=sep)
    fieldnames = rdr.fieldnames or []
    if "Division" not in fieldnames: raise SystemExit("Colonne 'Division' absente.")
    w = csv.DictWriter(g, fieldnames=fieldnames, delimiter=sep)
    w.writeheader()
    for row in rdr:
        row["Division"] = normalize_div(row.get("Division",""))
        w.writerow(row)

print("✅ Division normalisée →", dst)
PY
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
normalize.py
-------------
Normalise la colonne 'Division' dans un CSV fusionné des parents.
- Lecture robuste des encodages (UTF-8/UTF-8-SIG puis fallback CP1252).
- Détection automatique du séparateur (',' ou ';').
- Sortie UTF-8 sans BOM.
- Messages console ASCII pour éviter les erreurs d'encodage Windows.
"""

from __future__ import annotations
from pathlib import Path
import csv
import re
import sys
import unicodedata
from typing import Tuple, Iterable


def _detect_delimiter(sample: str) -> str:
    return ";" if sample.count(";") >= sample.count(",") else ","


def _try_read_text(p: Path) -> Tuple[str, str]:
    """
    Tente de lire p en 'utf-8-sig' puis fallback 'cp1252'.
    Retourne (texte, encodage_utilise).
    """
    for enc in ("utf-8-sig", "utf-8"):
        try:
            return p.read_text(encoding=enc), enc
        except UnicodeDecodeError:
            pass
    # Fallback CP1252 avec remplacement des caracteres non mappables
    try:
        return p.read_text(encoding="cp1252", errors="replace"), "cp1252"
    except Exception as e:
        raise RuntimeError(f"Impossible de lire le fichier '{p}': {e}")


def normalize_div(s: str) -> str:
    s = (s or "").strip()
    # Enlève l'enrobage Excel ="4 D"
    m = re.match(r'^=\s*"(.+)"\s*$', s)
    if m:
        s = m.group(1)
    # Unicode -> supprime diacritiques ; majuscules
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn").upper()
    # Uniformise : retire espaces/tirets/points "4 ÈME D" -> "4EMED"
    s = s.replace("ÈME", "E").replace("EME", "E")
    s = re.sub(r"[\s\-.]+", "", s)
    return s


def normalize_csv(src_path: Path, dst_path: Path) -> None:
    text, used_enc = _try_read_text(src_path)
    sep = _detect_delimiter(text[:4096])

    # Re-parse via csv avec l'encodage effectivement utilise
    with src_path.open("r", encoding=used_enc, errors="replace", newline="") as f_in, \
         dst_path.open("w", encoding="utf-8", newline="") as f_out:
        rdr = csv.DictReader(f_in, delimiter=sep)
        fieldnames = rdr.fieldnames or []
        if "Division" not in fieldnames:
            raise SystemExit("Colonne 'Division' absente dans le CSV d'entree.")
        w = csv.DictWriter(f_out, fieldnames=fieldnames, delimiter=sep)
        w.writeheader()
        for row in rdr:
            row["Division"] = normalize_div(row.get("Division", ""))
            w.writerow(row)

    # Messages console en ASCII uniquement (pas d'emoji) pour Windows
    print(f"[normalize] Encodage entree: {used_enc} | separateur: '{sep}'")
    print(f"[normalize] Sortie ecrite: {dst_path}")


def _default_paths() -> Tuple[Path, Path]:
    """
    Chemins par defaut si on lance le script sans arguments.
    Compatibles avec l'organisation initiale du projet.
    """
    src = Path("parents_4e_merged.csv")
    dst = src.with_name("parents_4e_merged_norm.csv")
    return src, dst


def main(argv: Iterable[str]) -> int:
    import argparse

    parser = argparse.ArgumentParser(
        description="Normalise la colonne 'Division' d'un CSV (UTF-8/CP1252).",
    )
    parser.add_argument("--in", dest="inp", type=str, default=None,
                        help="Chemin du CSV d'entree (par defaut: parents_4e_merged.csv)")
    parser.add_argument("--out", dest="out", type=str, default=None,
                        help="Chemin du CSV de sortie (par defaut: parents_4e_merged_norm.csv)")

    args = parser.parse_args(list(argv))

    if args.inp:
        src = Path(args.inp)
        dst = Path(args.out) if args.out else src.with_name(src.stem + "_norm.csv")
    else:
        src, dst = _default_paths()

    if not src.exists():
        print(f"[normalize] ERREUR: fichier d'entree introuvable: {src}", file=sys.stderr)
        return 2

    try:
        normalize_csv(src, dst)
    except Exception as e:
        # Eviter les erreurs d'encodage console: convertir en pur ASCII
        msg = str(e).encode("ascii", errors="replace").decode("ascii")
        print(f"[normalize] Exception: {msg}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))