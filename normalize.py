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