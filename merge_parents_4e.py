#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
merge_parents_4e.py — Fusion SIECLE robuste (encodage, séparateur, entêtes)

Fonctions principales :
- Lecture robuste d'un export SIECLE (détection encodage/separateur)
- Normalisation souple des en-têtes (accents/espaces/casse)
- Construction d'une colonne 'Emails' en combinant les adresses des représentants
- Sauvegarde d'un CSV 'parents_4e_merged.csv' (toutes divisions)
- Sauvegarde d'un CSV filtré sur une classe : 'parents_<CLASSE>_canon.csv'

Utilisation CLI directe :
  python merge_parents_4e.py /chemin/export.csv --classe 6A --out-dir /chemin/sortie
"""

from __future__ import annotations
import argparse
import os
import sys
import unicodedata
import re
from typing import Optional, Tuple, List

import pandas as pd

# ============================
# Affichage / Logs minimalistes
# ============================
def info(msg: str): print(f"[INFO] {msg}")
def ok(msg: str):   print(f"[OK]  {msg}")
def warn(msg: str): print(f"[WARN] {msg}")
def err(msg: str):  print(f"[ERREUR] {msg}", file=sys.stderr)

# ============================
# Normalisation d'en-têtes
# ============================
def strip_accents(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def canon_header(h: str) -> str:
    h = h.strip()
    h = strip_accents(h)
    h = re.sub(r'\s+', ' ', h)
    return h.lower()

# Colonnes cibles minimales
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

# Synonymes tolérés -> clé canonique
HEADER_MAP = {
    # division
    "division": ["division", "classe"],
    # nom
    "nom": ["nom de famille", "nom", "nom eleve", "eleve nom", "nomfamille"],
    # prénom
    "prenom": ["prénom 1", "prenom 1", "prénom", "prenom", "eleve prenom", "prenom1"],
    # emails représentants
    "email1": ["courriel repr. légal", "courriel repr. legal", "email repr. légal", "mail repr. légal", "adresse électronique repr. légal", "courriel representant legal", "adresseelectroniquereprlegal"],
    "email2": ["courriel autre repr. légal", "courriel autre repr. legal", "email autre repr. légal", "mail autre repr. légal", "adresse électronique autre repr. légal", "courriel autre representant legal", "adresseelectroniqueautrereprlegal"],
}

def find_col(df_cols: List[str], targets: List[str]) -> Optional[str]:
    """Retourne le nom de colonne original correspondant à une des cibles normalisées."""
    norm = {canon_header(c): c for c in df_cols}
    for t in targets:
        if canon_header(t) in norm:
            return norm[canon_header(t)]
    return None

# ============================
# Lecture robuste du CSV SIECLE
# ============================
def read_siecle_csv(path: str) -> Tuple[pd.DataFrame, str, str]:
    """
    Essaie plusieurs encodages et séparateurs puis retourne (df, encoding, sep).
    Logue les entêtes détectées.
    """
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1"]
    seps = [";", ","]
    last_err = None
    for enc in encodings:
        for sep in seps:
            try:
                df = pd.read_csv(path, sep=sep, encoding=enc, dtype=str, engine="python")
                # drop colonnes vides exactes
                df = df.dropna(how="all", axis=1)
                # trim
                df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
                headers = list(df.columns)
                print(f"Entêtes CSV détectées ({enc}, sep='{sep}') : {headers}")
                return df, enc, sep
            except Exception as e:
                last_err = e
                continue
    raise RuntimeError(f"Impossible de lire le CSV '{path}' avec essais {encodings} et seps {seps}.\nDernière erreur: {last_err}")

# ============================
# Pipeline de fusion
# ============================
def fuse_single(parents_csv: str, classe: Optional[str], out_dir: str) -> Tuple[str, Optional[str]]:
    if not os.path.isfile(parents_csv):
        raise FileNotFoundError(f"Fichier parents introuvable: {parents_csv}")

    # Petit compteur de lignes estimées (sans charger encore)
    try:
        est_lines = max(0, sum(1 for _ in open(parents_csv, 'rb')) - 1)
        print(f"[OK]  -> {os.path.basename(parents_csv)} ({est_lines} lignes estimées)")
    except Exception:
        print(f"[OK]  -> {os.path.basename(parents_csv)}")

    df, used_enc, used_sep = read_siecle_csv(parents_csv)
    print(f"[INFO] Lecture CSV: encodage={used_enc}, sep='{used_sep}'")
    if df.empty:
        raise ValueError("CSV vide.")

    # Résolution tolérante des colonnes nécessaires
    col_div = find_col(df.columns, HEADER_MAP["division"])
    col_nom = find_col(df.columns, HEADER_MAP["nom"])
    col_pre = find_col(df.columns, HEADER_MAP["prenom"])
    col_e1  = find_col(df.columns, HEADER_MAP["email1"])
    col_e2  = find_col(df.columns, HEADER_MAP["email2"])

    # Log debug lisible
    print(f"Colonnes détectées → Division='{col_div}' | Nom='{col_nom}' | Prénom='{col_pre}'")
    if not col_div or not col_nom or not col_pre:
        # Aperçu debug
        sample = df.head(6).to_csv(sep=used_sep, index=False)
        print("→ Vérifiez les entêtes (Division/Nom/Prénom) et le séparateur. Aperçu du fichier :")
        print(sample)
        raise KeyError(f"Colonnes essentielles manquantes. Résolu: Division='{col_div}' | Nom='{col_nom}' | Prénom='{col_pre}'")

    # Construire Emails combinés
    def join_emails(a: Optional[str], b: Optional[str]) -> str:
        parts = []
        for x in [a, b]:
            if isinstance(x, str) and x.strip():
                # Normaliser séparateurs , -> ;
                parts.extend([p.strip() for p in x.replace(",", ";").split(";") if p.strip()])
        # dédoublonner
        out, seen = [], set()
        for e in parts:
            le = e.lower()
            if le not in seen:
                seen.add(le)
                out.append(e)
        return ";".join(out)

    emails = []
    for _, row in df.iterrows():
        a = row.get(col_e1, "")
        b = row.get(col_e2, "")
        emails.append(join_emails(a, b))
    df["Emails"] = emails
    non_empty = (df["Emails"].astype(str).str.strip() != "").sum()
    print(f"[INFO] Emails non vides : {non_empty}/{len(df)}")

    # Construire DataFrame canonique minimal
    out_cols = [col_nom, col_pre, col_div, "Emails"]
    for extra in [col_e1, col_e2]:
        if extra and extra not in out_cols:
            out_cols.append(extra)

    df_out = df[out_cols].copy()
    df_out.rename(columns={
        col_nom: "Nom de famille",
        col_pre: "Prénom 1",
        col_div: "Division",
        col_e1 if col_e1 else "Courriel repr. légal": "Courriel repr. légal",
        col_e2 if col_e2 else "Courriel autre repr. légal": "Courriel autre repr. légal",
    }, inplace=True)

    # Sauvegardes
    os.makedirs(out_dir, exist_ok=True)
    merged_path = os.path.join(out_dir, "parents_4e_merged.csv")
    df_out.to_csv(merged_path, sep=";", index=False, encoding="utf-8")
    ok(f"-> fusion écrite : {merged_path}")

    filtered_path = None
    if classe:
        m = df_out["Division"].astype(str).str.strip().str.upper() == str(classe).strip().upper()
        df_c = df_out[m].copy()
        filtered_path = os.path.join(out_dir, f"parents_{str(classe).strip()}_canon.csv")
        df_c.to_csv(filtered_path, sep=";", index=False, encoding="utf-8")
        print(f"→ {len(df_c)} lignes 'Division={classe}' dans {filtered_path}")

    return merged_path, filtered_path

# ============================
# Entrée CLI
# ============================
def main():
    ap = argparse.ArgumentParser(description="Fusionne un export SIECLE vers CSV canoniques (merge + par classe).")
    ap.add_argument("parents_csv", help="Export SIECLE (CSV unique)")
    ap.add_argument("--classe", help="Classe à filtrer (ex: 6A)", default=None)
    ap.add_argument("--out-dir", help="Dossier de sortie", default=".")
    args = ap.parse_args()

    try:
        merged, filtered = fuse_single(args.parents_csv, args.classe, args.out_dir)
        ok("Fusion OK")
    except Exception as e:
        err(str(e))
        sys.exit(1)

if __name__ == "__main__":
    main()