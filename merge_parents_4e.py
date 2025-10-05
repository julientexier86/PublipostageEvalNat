#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Fusionne les 4 CSV d'extraction Siècle (4A,4B,4C,4D) en un seul fichier:
 - parents_4e_merged.csv  : colonnes normalisées (comme Siècle)
 - parents_4e_mailmerge.csv : + colonne Emails (parents 1 et 2 concaténés, séparés par ;)
Tolérant aux encodages (utf-8-sig/latin1/cp1252) et aux séparateurs (; ou ,).
Gère les variantes d'intitulés de colonnes (espaces/accents/casse/petites variations).
Déduplique par (Nom de famille, Prénom 1, Date de naissance).
"""

import sys
import os
import unicodedata
import pandas as pd

TARGET_COLS = [
    "Nom de famille",
    "Prénom 1",
    "Date de naissance",
    "Division",
    "Nom de famille repr. légal",
    "Prénom repr. légal",
    "Courriel repr. légal",
    "Nom de famille autre repr. légal",
    "Prénom autre repr. légal",
    "Courriel autre repr. légal",
]

# Normalisation douce des en-têtes (enlève accents, ponctuation, espaces)
def norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    # remplace tout ce qui n'est pas alphanum par rien
    s = "".join(ch for ch in s if ch.isalnum())
    return s.lower()

# Dictionnaire de correspondance "header normalisé" -> "nom canonique"
CANON_MAP = {
    # Elève
    "nomdefamille": "Nom de famille",
    "nomfamille": "Nom de famille",
    "nom": "Nom de famille",
    "prenom1": "Prénom 1",
    "prenom": "Prénom 1",
    "datedenaissance": "Date de naissance",
    "naissance": "Date de naissance",
    "division": "Division",

    # Représentant légal 1
    "nomdefamillereprlegal": "Nom de famille repr. légal",
    "nomfamillereprlegal": "Nom de famille repr. légal",
    "nomreprlegal": "Nom de famille repr. légal",
    "prenomreprlegal": "Prénom repr. légal",
    "courrielreprlegal": "Courriel repr. légal",
    "emailreprlegal": "Courriel repr. légal",
    "mailreprlegal": "Courriel repr. légal",
    "adresseelectroniquereprlegal": "Courriel repr. légal",

    # Représentant légal 2 (autre)
    "nomdefamilleautrereprlegal": "Nom de famille autre repr. légal",
    "nomfamilleautrereprlegal": "Nom de famille autre repr. légal",
    "nomautrereprlegal": "Nom de famille autre repr. légal",
    "prenomautrereprlegal": "Prénom autre repr. légal",
    "courrielautrereprlegal": "Courriel autre repr. légal",
    "emailautrereprlegal": "Courriel autre repr. légal",
    "mailautrereprlegal": "Courriel autre repr. légal",
    "adresseelectroniqueautrereprlegal": "Courriel autre repr. légal",
}

def read_any_csv(path: str) -> pd.DataFrame:
    last_err = None
    for enc in ("utf-8-sig", "cp1252", "latin1"):
        try:
            # sep=None + engine='python' => Sniffer qui détecte ; ou ,
            df = pd.read_csv(path, sep=None, engine="python", encoding=enc)
            return df
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Impossible de lire {path}: {last_err}")

def standardize_headers(df: pd.DataFrame) -> pd.DataFrame:
    new_cols = {}
    for c in df.columns:
        key = norm(c)
        target = CANON_MAP.get(key)
        if target is None:
            # pas dans la map ? on tente quelques cas particuliers
            if key.startswith("division"):
                target = "Division"
            elif key.startswith("prenom") and "repr" not in key:
                target = "Prénom 1"
            else:
                target = c  # on conserve, mais il ne sera pas pris si non ciblé
        new_cols[c] = target
    df = df.rename(columns=new_cols)
    return df

def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in TARGET_COLS:
        if col not in df.columns:
            df[col] = pd.NA
    # on ne garde que les colonnes cibles, dans l'ordre
    return df[TARGET_COLS].copy()

def clean_values(df: pd.DataFrame) -> pd.DataFrame:
    # Trim espaces
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].astype(str).str.strip().replace({"nan": pd.NA})
    # Dates : laissons le format d’origine (Siècle), on ne force pas l’ISO ici
    # Division : homogénéiser espaces
    if "Division" in df.columns:
        df["Division"] = df["Division"].astype(str).str.strip()
    return df

def combine_emails(row) -> str:
    emails = []
    for c in ("Courriel repr. légal", "Courriel autre repr. légal"):
        v = row.get(c)
        if pd.notna(v) and str(v).strip():
            # si le champ comporte déjà plusieurs emails séparés par ;, on garde tel quel
            parts = [p.strip() for p in str(v).replace(",", ";").split(";") if p.strip()]
            emails.extend(parts)
    # dédoublonne en conservant l'ordre
    seen = set()
    out = []
    for e in emails:
        low = e.lower()
        if low not in seen:
            seen.add(low)
            out.append(e)
    return ";".join(out)

def merge_files(paths):
    frames = []
    for p in paths:
        df = read_any_csv(p)
        df = standardize_headers(df)
        df = ensure_columns(df)
        df = clean_values(df)
        frames.append(df)

    merged = pd.concat(frames, ignore_index=True)

    # Déduplication par (Nom, Prénom, Date de naissance)
    key_cols = ["Nom de famille", "Prénom 1", "Date de naissance"]

    # on regroupe et on combine intelligemment
    def agg_func(group: pd.DataFrame) -> pd.Series:
        # premier non-null pour les champs simples
        res = {}
        for col in TARGET_COLS:
            if col in ("Courriel repr. légal", "Courriel autre repr. légal"):
                # on gardera les deux, mais ici on prend le premier non-null
                res[col] = group[col].dropna().astype(str).replace({"nan": None}).dropna().head(1).tolist()
                res[col] = res[col][0] if res[col] else pd.NA
            else:
                res[col] = group[col].dropna().astype(str).replace({"nan": None}).dropna().head(1).tolist()
                res[col] = res[col][0] if res[col] else pd.NA
        return pd.Series(res)

    merged = merged.groupby(key_cols, dropna=False, as_index=False).apply(agg_func)

    # Colonne Emails (parents 1 + 2)
    merged["Emails"] = merged.apply(combine_emails, axis=1)

    # Sauvegardes
    merged_std = merged[TARGET_COLS].copy()
    merged_mail = merged[
        ["Division", "Nom de famille", "Prénom 1", "Date de naissance",
         "Courriel repr. légal", "Courriel autre repr. légal", "Emails"]
    ].copy()

    # Noms de fichiers en sortie (dans le dossier du script)
    out_std = "parents_4e_merged.csv"
    out_mail = "parents_4e_mailmerge.csv"

    merged_std.to_csv(out_std, index=False, encoding="utf-8-sig")
    merged_mail.to_csv(out_mail, index=False, encoding="utf-8-sig")

    print(f"✅ Fusion OK")
    print(f"   → {out_std} : {len(merged_std)} lignes")
    print(f"   → {out_mail} : {len(merged_mail)} lignes (avec colonne Emails)")

if __name__ == "__main__":
    # Utilisation:
    #   python merge_parents_4e.py [paths...]
    # Sans arguments, lit les fichiers par défaut dans le même dossier:
    #   exportCSVExtraction4A.csv, exportCSVExtraction4B.csv, exportCSVExtraction4C.csv, exportCSVExtraction4D.csv
    if len(sys.argv) > 1:
        paths = sys.argv[1:]
    else:
        paths = [
            "exportCSVExtraction4A.csv",
            "exportCSVExtraction4B.csv",
            "exportCSVExtraction4C.csv",
            "exportCSVExtraction4D.csv",
        ]
    # vérifie l’existence
    missing = [p for p in paths if not os.path.exists(p)]
    if missing:
        print("⚠️ Fichiers introuvables :", ", ".join(missing))
        sys.exit(1)
    merge_files(paths)