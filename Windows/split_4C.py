# -*- coding: utf-8 -*-
"""
Split 4C OCR export into per-student PDFs:
- Input:  export_126886_all.pdf (OCR)
- Output: one PDF per élève et par discipline:
          4D_Nom_Prénom_Français_2025-2026.pdf
          4D_Nom_Prénom_Mathématiques_2025-2026.pdf
Testé avec PyPDF2 (texte OCR requis).
"""

import re
import os
import unicodedata
from pathlib import Path

from PyPDF2 import PdfReader, PdfWriter

# === À ADAPTER SI BESOIN ===
CLASS_LABEL = "4D"
SCHOOL_YEAR = "2025-2026"

# Répertoire d'entrée et de sortie (à adapter à ton Mac)
INPUT_PDF = "/Users/julien/Downloads/export_126892_all.pdf"
OUTPUT_DIR = "/Users/julien/Downloads/Publipostage_4D"

# Garde les accents dans les noms de fichiers ?
KEEP_ACCENTS_IN_FILENAME = False  # mets True si tu veux garder les accents

# ===========================

def strip_accents(s: str) -> str:
    return ''.join(
        c for c in unicodedata.normalize('NFD', s)
        if unicodedata.category(c) != 'Mn'
    )

def safe_filename(s: str) -> str:
    s = s.strip()
    # Conserver les accents si demandé, sinon les retirer
    if KEEP_ACCENTS_IN_FILENAME:
        # NFC pour éviter les formes décomposées bizarres et garder é, ç, …
        s = unicodedata.normalize("NFC", s)
    else:
        s = strip_accents(s)  # supprime les accents

    # Remplacer espaces multiples par un underscore
    s = re.sub(r"\s+", "_", s)

    # Autoriser lettres/chiffres/underscore/point/tiret en UNICODE
    # \w inclut les lettres accentuées en Python (flag UNICODE)
    s = re.sub(r"[^\w.\-]", "", s, flags=re.UNICODE)

    # Éviter doubles underscores
    s = re.sub(r"_+", "_", s)

    # Trim underscores résiduels en bord
    s = s.strip("_")
    return s

def _norm_text(t: str) -> str:
    return strip_accents((t or "")).lower()

def score_disciplines(page_text: str) -> tuple[int, int]:
    """
    Renvoie un couple (score_fr, score_ma) fondé sur des mots-clés et des indices simples.
    On cumule des points pour éviter les faux positifs, puis on tranchera au niveau élève.
    """
    t = _norm_text(page_text)
    # jeux de mots-clés (volontairement plus précis que la v1)
    fr_kw = [
        "francais", "langue francaise", "lecture", "comprehension",
        "orthographe", "dictee", "vocabulaire", "grammaire", "conjugaison",
        "maitrise de la langue"
    ]
    ma_kw = [
        "mathematiques", "maths", "nombres", "numeration", "calcul",
        "geometrie", "mesure", "grandeurs", "fractions",
        "proportionnalite", "equation", "probleme", "statistiques",
        "probabilites"
    ]
    # scores mots-clés
    fr = sum(t.count(k) for k in fr_kw)
    ma = sum(t.count(k) for k in ma_kw)
    # indices supplémentaires côté maths: densité de chiffres et symboles
    digits = sum(ch.isdigit() for ch in t)
    ops = sum(ch in "+-×x*/÷=<>≤≥" for ch in t)
    # pondération légère
    ma += (digits // 25) + (ops // 5)
    return fr, ma

def guess_discipline(page_text: str) -> str | None:
    """
    Détection par score. On ne tranche que si un score domine nettement.
    Le reste sera géré au niveau élève (répartition 2 pages → Fr/Math).
    """
    fr, ma = score_disciplines(page_text)
    if fr == 0 and ma == 0:
        return None
    # marge: au moins 2 points d'écart ou un score >=3 et l'autre 0
    if fr >= ma + 2 or (fr >= 3 and ma == 0):
        return "Français"
    if ma >= fr + 2 or (ma >= 3 and fr == 0):
        return "Mathématiques"
    return None

def extract_name(page_text: str) -> str | None:
    """
    Heuristique fiable sur la maquette Eduscol:
    - On trouve la ligne avec "Année scolaire"
    - Le nom complet de l’élève est la première ligne non vide après
      ce bloc, souvent "Prénom NOM" (NOM en majuscules).
    """
    lines = [l.strip() for l in page_text.splitlines()]
    # cherche l'index d’une ligne contenant "Année scolaire"
    idx = None
    for i, l in enumerate(lines):
        if "Année scolaire" in l or "ANNEE SCOLAIRE" in l.upper():
            idx = i
            break
    # fallback si pas trouvé : on prendra la première ligne qui ressemble à "Prénom NOM"
    candidate = None
    if idx is not None:
        for l in lines[idx+1: idx+8]:  # on regarde quelques lignes après
            if not l:
                continue
            # très souvent c'est du type "Robin ARLOT"
            # on valide si on a au moins deux mots et si le dernier a ≥2 lettres majuscules
            parts = l.split()
            if len(parts) >= 2:
                last = parts[-1]
                if re.fullmatch(r"[A-ZÉÈÊËÀÂÄÔÖÛÜÇ\-]{2,}", last):
                    candidate = l
                    break

    if not candidate:
        # Plan B : cherche n’importe quelle ligne "Prénom NOM" (dernier mot tout en maj)
        for l in lines:
            parts = l.split()
            if len(parts) >= 2:
                last = parts[-1]
                if re.fullmatch(r"[A-ZÉÈÊËÀÂÄÔÖÛÜÇ\-]{2,}", last):
                    candidate = l
                    break

    if candidate:
        # Nettoyage fin (certains OCR collent des espaces parasites)
        candidate = re.sub(r"\s{2,}", " ", candidate).strip()
        # Petits ajustements de casse : on conserve tel quel (ex: "Gael HAMON DE ALMEIDA")
        return candidate

    return None

def split_pdf():
    in_path = Path(INPUT_PDF)
    if not in_path.exists():
        raise FileNotFoundError(f"Introuvable: {in_path}")

    out_dir = Path(OUTPUT_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)

    reader = PdfReader(str(in_path))
    total = len(reader.pages)

    # 1) Collecte: infos page par page
    pages_info = []  # [{idx:int, name:str|None, disc:str|None, fr:int, ma:int, page:PageObj, raw:str}]
    for i in range(total):
        page = reader.pages[i]
        try:
            txt = page.extract_text() or ""
        except Exception:
            txt = ""
        name = extract_name(txt)
        fr_s, ma_s = score_disciplines(txt)
        # première hypothèse locale (faible, sera réévaluée au niveau élève)
        disc = guess_discipline(txt)
        pages_info.append({
            "idx": i+1, "name": name, "disc": disc,
            "fr": fr_s, "ma": ma_s, "page": page, "raw": txt
        })

    # 2) Regroupement par élève (indépendant de l'adjacence)
    by_student: dict[str, list[dict]] = {}
    for info in pages_info:
        if not info["name"]:
            continue
        by_student.setdefault(info["name"], []).append(info)

    # 3) Attribution de la discipline par élève
    #    - si 2 pages: on choisit la meilleure en Fr pour "Français", la meilleure en Ma pour "Mathématiques"
    #    - si 1 page: on garde l'hypothèse locale si nette, sinon None
    for name, items in by_student.items():
        # tri stable par index pour une répartition déterministe en cas d'égalité
        items.sort(key=lambda d: d["idx"])
        if len(items) >= 2:
            # choisir page avec meilleur score_fr et meilleur score_ma
            best_fr = max(items, key=lambda d: (d["fr"], -d["idx"]))
            best_ma = max(items, key=lambda d: (d["ma"], -d["idx"]))
            # si c'est la même page qui gagne des deux côtés (rare), on attribue la seconde à l'autre matière
            if best_fr is best_ma and len(items) >= 2:
                # choisir une autre page différente
                alt = next((x for x in items if x is not best_fr), None)
                # comparaison des scores pour savoir qui prend quoi
                if best_fr["fr"] >= best_fr["ma"]:
                    best_fr["disc"] = "Français"
                    if alt: alt["disc"] = "Mathématiques"
                else:
                    best_ma["disc"] = "Mathématiques"
                    if alt: alt["disc"] = "Français"
            else:
                best_fr["disc"] = "Français"
                best_ma["disc"] = "Mathématiques"
            # pour les autres pages éventuelles (si >2), on garde la meilleure proximité de score
            for it in items:
                if it["disc"] is None:
                    it["disc"] = "Français" if it["fr"] >= it["ma"] else "Mathématiques"
        else:
            # 1 seule page: on garde l'hypothèse la plus forte si elle existe
            it = items[0]
            if it["disc"] is None:
                if it["fr"] > it["ma"] and it["fr"] >= 2:
                    it["disc"] = "Français"
                elif it["ma"] > it["fr"] and it["ma"] >= 2:
                    it["disc"] = "Mathématiques"
                # sinon on laisse None (sera listé en missing)

    # 4) Écriture des fichiers + rapport
    extracted = []
    missing = []
    for info in pages_info:
        name = info["name"]
        disc = info["disc"]
        page = info["page"]
        idx  = info["idx"]
        if not name or not disc:
            missing.append({
                "page": idx,
                "discipline": disc,
                "name": name,
                "sample": (info["raw"][:400].replace("\n", " ") if info.get("raw") else ""),
                "scores": (info["fr"], info["ma"])
            })
            continue

        # Normalise "Prénom NOM" -> Nom, Prénom pour le nom de fichier
        parts = name.split()
        prenom = " ".join(parts[:-1]) if len(parts) >= 2 else ""
        nom = parts[-1] if parts else "Inconnu"

        filename = f"{CLASS_LABEL}_{nom}_{prenom}_{disc}_{SCHOOL_YEAR}.pdf"
        filename = safe_filename(filename)
        out_file = out_dir / filename

        writer = PdfWriter()
        writer.add_page(page)
        with open(out_file, "wb") as f:
            writer.write(f)
        extracted.append((idx, name, disc, str(out_file)))

    print(f"Pages traitées: {len(extracted)} / {total}")
    if extracted:
        print("\nExemples de sorties:")
        for p, n, d, fpath in extracted[:5]:
            print(f"  p.{p:>3}  {n}  [{d}] -> {fpath}")

    if missing:
        print("\nATTENTION: pages non résolues (discipline et/ou nom manquants):")
        for m in missing[:10]:
            frs, mas = m.get("scores", (0,0))
            print(f"  p.{m['page']:>3}  discipline={m['discipline']}  name={m['name']}  scores(FR,MA)=({frs},{mas})")
        if len(missing) > 10:
            print(f"  ... {len(missing)-10} autres")

if __name__ == "__main__":
    split_pdf()