#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse, csv, os, re, sys, unicodedata
from pathlib import Path
from collections import defaultdict
import pandas as pd

MESSAGE_TYPE = (
    "Madame, Monsieur,\n\n"
    "Veuillez trouver en pièces jointes les comptes rendus des évaluations nationales passées par vos enfants.\n"
    "Les enseignants reviendront dessus lors des remises de bulletins. Vous pourrez poser toutes les questions s'y rapportant lors de ce rendez-vous.\n\n"
    "Bien cordialement,\n"
    "Pour l'équipe de direction,"
)

# ---------- Normalisation ----------
def strip_accents(s: str) -> str:
    if s is None: return ""
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c))

def norm_div(s: str) -> str:
    if not s: return ""
    s = strip_accents(s).upper().replace("\u00A0"," ").strip()
    m = re.search(r"([3-6])\D*([A-Z])", s)
    if m: return f"{m.group(1)}{m.group(2)}"
    return re.sub(r"(?<=\d)\s+(?=[A-Z])", "", s)

def norm_name_token(s: str) -> str:
    s = strip_accents(str(s)).lower().strip()
    s = re.sub(r"['’`^¨~]", "", s)
    s = re.sub(r"[^a-z]", "", s)
    return s

def split_name_field_to_tokens(piece: str):
    # Sépare espaces/traits d’union en tokens
    piece = strip_accents(piece).strip()
    subs = [t for t in re.split(r"[\s\-]+", piece) if t]
    return subs

def surname_key_from_tokens(tokens):
    toks = [t for t in (norm_name_token(x) for x in tokens) if t]
    return "".join(sorted(toks))

def surname_key_from_csv_nom(nom: str) -> str:
    raw = strip_accents(nom).strip()
    tokens = re.split(r"[\s\-]+", raw)
    tokens = [t for t in tokens if t]
    return surname_key_from_tokens(tokens)

def norm_disc(s: str) -> str:
    s = strip_accents(s).lower().strip()
    if s.startswith("franc"): return "francais"
    if s.startswith("math"): return "mathematiques"
    return s

# ---------- Index des PDF ----------
def build_catalog(pdf_base: Path):
    """
    Indexe récursivement tous les PDFs sous pdf_base.

    Nom attendu: Classe _ [bloc noms/prénoms…] _ Discipline _ Année

    v5 : on 'aplanit' le bloc central en sous-tokens (séparés par espaces/traits),
         puis on essaie comme 'prénom' TOUT segment contigu (longueur 1..n),
         en concaténant sans séparateur (ex: Lily + Morgane -> 'lilymorgane').
         Le reste des tokens = set des 'noms' (ordre indifférent).
    """
    catalog = {}
    by_div = defaultdict(list)

    for p in pdf_base.rglob("*.pdf"):
        stem = p.stem
        parts = stem.split("_")
        if len(parts) < 4:
            continue

        div_raw = parts[0]
        annee = parts[-1].strip()
        disc_raw = parts[-2]
        mid_blocks = parts[1:-2]  #

        divN = norm_div(div_raw)
        discN = norm_disc(disc_raw)

        # Aplatir les blocs en sous-tokens
        flat = []
        for blk in mid_blocks:
            flat.extend(split_name_field_to_tokens(blk))

        if not flat:
            flat = mid_blocks[:]

        # Générer tous les segments contigus comme "prénom candidat"
        n = len(flat)
        for i in range(n):
            for j in range(i, n):
                pren_concat = "".join(norm_name_token(t) for t in flat[i:j+1])
                if not pren_concat:
                    continue
                # Les autres tokens (hors i..j) deviennent les 'noms'
                sur_tokens = [flat[k] for k in range(n) if k < i or k > j]
                sur_key = surname_key_from_tokens(sur_tokens)
                key = (divN, pren_concat, sur_key, discN, annee)
                catalog[key] = str(p)

        by_div[divN].append(p.name)

    return catalog, by_div

# ---------- Lecture CSV ----------
def read_input_csv(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, dtype=str).fillna("")
    ren = {}
    if "Nom de famille" in df.columns: ren["Nom de famille"] = "Nom"
    if "Prénom 1" in df.columns: ren["Prénom 1"] = "Prenom"
    if "Prénom" in df.columns and "Prenom" not in df.columns: ren["Prénom"] = "Prenom"
    df = df.rename(columns=ren)
    if "Division" not in df.columns:
        raise ValueError("Colonne 'Division' absente.")
    df["Division"] = df["Division"].apply(norm_div)
    return df

# ---------- Main ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True, help="parents_4e_merged.csv")
    ap.add_argument("--pdf-base", dest="pdf_base", required=True, help="Dossier racine des PDFs (récursif)")
    ap.add_argument("--out", dest="out_csv", required=True, help="CSV de sortie pour Thunderbird")
    ap.add_argument("--annee", dest="annee", default="2025-2026")
    ap.add_argument("--missing", dest="missing_csv", default=None, help="CSV des cas sans PJ")
    args = ap.parse_args()

    base = Path(args.pdf_base)
    if not base.exists():
        print(f"[ERREUR] --pdf-base introuvable : {base}", file=sys.stderr)
        sys.exit(1)

    print("[INFO] Indexation des PDFs…")
    catalog, by_div = build_catalog(base)

    print("[INFO] Lecture du CSV…")
    df = read_input_csv(args.inp)

    rows_out, rows_missing = [], []
    count_fr = count_ma = total = 0

    for _, r in df.iterrows():
        nom_raw = (r.get("Nom","") or "").strip()
        prenom_raw = (r.get("Prenom","") or "").strip()
        div_raw = (r.get("Division","") or "").strip()
        if not (nom_raw and prenom_raw and div_raw):
            continue

        divN = norm_div(div_raw)
        prenN = norm_name_token(prenom_raw)  # ex: "Lily-Morgane" -> "lilymorgane"
        sur_key = surname_key_from_csv_nom(nom_raw)
        annee = args.annee

        key_fr = (divN, prenN, sur_key, "francais", annee)
        key_ma = (divN, prenN, sur_key, "mathematiques", annee)

        pj_fr = catalog.get(key_fr, "")
        pj_ma = catalog.get(key_ma, "")

        # Fallback: essayer chaque morceau isolé du nom composé comme clé "nom"
        if not pj_fr or not pj_ma:
            tokens = [t for t in re.split(r"[\s\-]+", strip_accents(nom_raw).strip()) if t]
            tokens_norm = [norm_name_token(t) for t in tokens]
            alt_sur_keys = {"".join(sorted([t])) for t in tokens_norm if t}
            if not pj_fr:
                for ak in alt_sur_keys:
                    pj_fr = catalog.get((divN, prenN, ak, "francais", annee), pj_fr)
                    if pj_fr: break
            if not pj_ma:
                for ak in alt_sur_keys:
                    pj_ma = catalog.get((divN, prenN, ak, "mathematiques", annee), pj_ma)
                    if pj_ma: break

        emails = ";".join([
            (r.get("Courriel repr. légal","") or "").strip(),
            (r.get("Courriel autre repr. légal","") or "").strip()
        ])
        emails = ";".join([e for e in emails.split(";") if e])

        if pj_fr: count_fr += 1
        if pj_ma: count_ma += 1
        total += 1

        objet = f"Évaluations nationales – {nom_raw.upper()} {prenom_raw} ({divN})"
        attachments = ";".join([p for p in [pj_fr, pj_ma] if p])

        rows_out.append({
            "Classe": divN,
            "Nom": nom_raw.upper(),
            "Prénom": prenom_raw,
            "Emails": emails,
            "PJ_francais": pj_fr,
            "PJ_math": pj_ma,
            "Attachments": attachments,
            "Annee": annee,
            "Objet": objet,
            "CorpsMessage": MESSAGE_TYPE
        })

        if not pj_fr and not pj_ma:
            present = ", ".join(by_div.get(divN, [])[:12])
            nom_tokens = [t for t in re.split(r"[\s\-]+", nom_raw.strip()) if t]
            if len(nom_tokens) >= 2:
                nom1, nom2 = nom_tokens[0].upper(), nom_tokens[-1].upper()
            else:
                nom1 = (nom_tokens[0].upper() if nom_tokens else nom_raw.upper())
                nom2 = ""
            attendu_fr_a = f"{divN}_{nom_raw.upper()}_{prenom_raw}_Français_{annee}.pdf"
            attendu_fr_b = f"{divN}_{nom1}_{prenom_raw}_{nom2}_Français_{annee}.pdf" if nom2 else ""
            attendu_ma_a = f"{divN}_{nom_raw.upper()}_{prenom_raw}_Mathématiques_{annee}.pdf"
            attendu_ma_b = f"{divN}_{nom1}_{prenom_raw}_{nom2}_Mathématiques_{annee}.pdf" if nom2 else ""

            rows_missing.append({
                "Division": divN,
                "Nom": nom_raw,
                "Prénom": prenom_raw,
                "Attendu_Fr_1": attendu_fr_a,
                "Attendu_Fr_2": attendu_fr_b,
                "Attendu_Ma_1": attendu_ma_a,
                "Attendu_Ma_2": attendu_ma_b,
                "ExemplesFichiersDansDivision": present
            })

    out_fields = ["Classe","Nom","Prénom","Emails","PJ_francais","PJ_math","Attachments","Annee","Objet","CorpsMessage"]
    with open(args.out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=out_fields)
        w.writeheader()
        w.writerows(rows_out)

    miss_path = args.missing_csv or str(Path(args.out_csv).with_name("missing_4e.csv"))
    with open(miss_path, "w", newline="", encoding="utf-8") as f:
        if rows_missing:
            w = csv.DictWriter(f, fieldnames=list(rows_missing[0].keys()))
            w.writeheader()
            w.writerows(rows_missing)
        else:
            w = csv.writer(f); w.writerow(["Tout trouvé :)"])

    print(f"[OK] Élèves traités : {total}")
    print(f"     PJ Français trouvées : {count_fr}")
    print(f"     PJ Mathématiques trouvées : {count_ma}")
    print(f"     CSV écrit : {args.out_csv}")
    print(f"     Manquants détaillés : {miss_path}")

if __name__ == "__main__":
    main()
