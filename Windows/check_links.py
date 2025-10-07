# check_links.py
from pathlib import Path
import re, unicodedata, csv, sys

# === À personnaliser si besoin ===
CSV_IN = Path("/Users/julien/Downloads/Scripts qui semblent OK/parents_4e_merged_norm.csv")
PDF_BASE = Path("/Users/julien/Downloads/Publipostage_4D")
ANNEE    = "2025-2026"

def nfd(s: str) -> str:
    return unicodedata.normalize("NFD", s or "").strip()

def squash(s: str) -> str:
    # Supprime diacritiques, espaces, tirets, underscores
    s = nfd(s).lower()
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"[\s\-_]+", "", s)

# Regex du nom de fichier: Classe_Nom_Prenom_Discipline_Annee.pdf
rx = re.compile(r"""
    ^
    (?P<classe>[^_]+)_
    (?P<nom>[^_]+)_
    (?P<prenom>[^_]+)_
    (?P<disc>Français|Francais|Franais|Mathématiques|Mathematiques|Mathmatiques)_
    (?P<annee>\d{4}-\d{4})
    \.pdf$
""", re.X | re.U)

def detect_sep(path: Path) -> str:
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        sample = f.read(4096)
    # Choisit le plus fréquent entre ; et ,
    semi = sample.count(";")
    comma = sample.count(",")
    return ";" if semi >= comma else ","

def read_csv_rows(path: Path) -> list[dict]:
    sep = detect_sep(path)
    with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as f:
        rdr = csv.DictReader(f, delimiter=sep)
        return list(rdr)

# 1) Indexer les PDFs
indexed = []
not_matched = []
for p in PDF_BASE.glob("*.pdf"):
    m = rx.match(p.name)
    if not m:
        not_matched.append(p.name)
        continue
    d = m.groupdict()
    d["path"] = str(p)
    d["k_classe"] = squash(d["classe"])
    d["k_nom"]    = squash(d["nom"])
    d["k_prenom"] = squash(d["prenom"])
    d["k_annee"]  = d["annee"]
    indexed.append(d)

print(f"PDF indexés (regex OK) : {len(indexed)}")
if not_matched:
    print(f"PDF ignorés par la regex : {len(not_matched)}")
    for ex in not_matched[:10]:
        print("  -", ex)
    if len(not_matched) > 10:
        print("  …")

if not indexed:
    sys.exit(0)

classes_pdf = sorted(set(d["classe"] for d in indexed))
disc_pdf = sorted(set(d["disc"] for d in indexed))
print("Classes vues dans les PDFs:", classes_pdf)
print("Disciplines vues dans les PDFs:", disc_pdf)

# 2) Lire le CSV
rows = read_csv_rows(CSV_IN)
print("Lignes élèves CSV :", len(rows))

def get(row: dict, *names):
    for n in names:
        if n in row and (row[n] or "").strip():
            return row[n].strip()
    return ""

# 3) Diagnostic d’association
ok, no_class, no_year, no_name = 0, 0, 0, 0
samples = {"no_class": [], "no_year": [], "no_name": []}

for r in rows:
    div  = get(r, "Division", "Classe")
    nom  = get(r, "Nom de famille", "Nom")
    pren = get(r, "Prénom 1", "Prénom", "Prenom")

    k_div  = squash(div)
    k_nom  = squash(nom)
    k_pren = squash(pren)

    # a) Classe
    pdf_cls = [d for d in indexed if d["k_classe"] == k_div]
    if not pdf_cls:
        no_class += 1
        if len(samples["no_class"]) < 5:
            samples["no_class"].append((div, nom, pren))
        continue

    # b) Année
    pdf_cls_year = [d for d in pdf_cls if d["k_annee"] == ANNEE]
    if not pdf_cls_year:
        no_year += 1
        if len(samples["no_year"]) < 5:
            samples["no_year"].append((div, nom, pren))
        continue

    # c) Nom/Prénom stricts (après squash)
    cand = [d for d in pdf_cls_year if d["k_nom"] == k_nom and d["k_prenom"] == k_pren]
    if not cand:
        no_name += 1
        if len(samples["no_name"]) < 5:
            samples["no_name"].append((div, nom, pren))
        continue

    ok += 1

print(f"Matchs parfaits : {ok}")
print(f"Bloqués classe  : {no_class}  (exemples: {samples['no_class']})")
print(f"Bloqués année   : {no_year}   (exemples: {samples['no_year']})")
print(f"Bloqués nom/pré : {no_name}   (exemples: {samples['no_name']})")