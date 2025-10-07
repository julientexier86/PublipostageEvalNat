# -*- coding: utf-8 -*-

# ===== Sortie console robuste (Windows/UTF-8) =====
from __future__ import annotations
import os, sys

# Forcer un mode UTF-8 "à la Python" autant que possible
os.environ.setdefault("PYTHONUTF8", "1")
os.environ.setdefault("PYTHONIOENCODING", "UTF-8")

# Reconfigurer les flux pour éviter les erreurs charmap/UnicodeEncodeError
for _stream_name in ("stdout", "stderr", "stdin"):
    try:
        _s = getattr(sys, _stream_name)
        # stdin en lecture, stdout/err en écriture
        if _stream_name == "stdin":
            _s.reconfigure(encoding="utf-8", errors="replace")
        else:
            _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Impressions sûres quel que soit l'encodage de la console
import builtins as _bi

def _safe_printable(s: str) -> str:
    try:
        enc = (getattr(sys.stdout, "encoding", None) or "utf-8")
        s.encode(enc, "strict")
        return s
    except Exception:
        # En dernier recours : ASCII sans casser l'exécution
        try:
            return s.encode("ascii", "ignore").decode("ascii")
        except Exception:
            return repr(s)

def _print(*args, **kw):
    try:
        _bi.print(*[_safe_printable(str(a)) for a in args], **kw)
    except Exception:
        kw = {k: v for k, v in kw.items() if k != "file"}
        _bi.print(*[repr(a) for a in args], **kw)

print = _print  # remplace print globalement

# Hook global pour formater proprement les exceptions en UTF‑8
import traceback as _tb
def _safe_excepthook(exc_type, exc, tb):
    try:
        msg = "".join(_tb.format_exception(exc_type, exc, tb))
        print(msg)
    except Exception:
        try:
            _bi.print("".join(_tb.format_exception_only(exc_type, exc)))
        except Exception:
            _bi.print("Exception (formatting failed)")

sys.excepthook = _safe_excepthook

import csv, io, unicodedata, re, warnings
from pathlib import Path

# --- Helpers mojibake + email ---
_EMAIL_RE = re.compile(r"^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$", re.I)

def _fix_mojibake_header(s: str) -> str:
    """
    Corrige quelques séquences mojibake fréquentes vues dans les en-têtes exportés par Excel/Windows,
    avant normalisation (_norm_header).
    """
    if not s:
        return s
    rep = {
        # UTF-8 -> latin-1 mojibake
        "Prï¿½nom": "Prénom",
        "prï¿½nom": "prénom",
        "lï¿½gal": "légal",
        "Lï¿½gal": "Légal",
        "PrÃ©nom": "Prénom",
        "prÃ©nom": "prénom",
        "lÃ©gal": "légal",
        "Courriel repr. lÃ©gal": "Courriel repr. légal",
        "Courriel autre repr. lÃ©gal": "Courriel autre repr. légal",
        # Windows replacement char (U+FFFD) variants commonly seen as "�"
        "Pr�nom": "Prénom",
        "pr�nom": "prénom",
        "l�gal": "légal",
        "L�gal": "Légal",
        "Courriel repr. l�gal": "Courriel repr. légal",
        "Courriel autre repr. l�gal": "Courriel autre repr. légal",
    }
    for k, v in rep.items():
        s = s.replace(k, v)
    return s

def _canon_email_list(*values: str) -> list[str]:
    """Nettoie, déduplique et renvoie une liste d'emails valides."""
    seen = set()
    out = []
    for v in values:
        if not v:
            continue
        # éclater sur , ; / espace
        parts = re.split(r"[;,/\\\s]+", str(v))
        for p in parts:
            p = (p or "").strip().strip(".,;:()[]{}<>")
            if not p:
                continue
            if _EMAIL_RE.match(p) and p.lower() not in seen:
                seen.add(p.lower())
                out.append(p)
    return out

# ===== Pandas requis =====
try:
    import pandas as pd
except Exception as e:
    print("[ERREUR] pandas introuvable dans merge_parents_4e.py :", e)
    sys.exit(1)

# Couper les FutureWarning bruyants (pandas groupby/apply)
warnings.simplefilter("ignore", FutureWarning)

def _read_csv_robust(path: Path) -> "pd.DataFrame":
    """
    Lecture robuste d'un CSV SIECLE :
    - tente utf-8-sig, sinon cp1252 ;
    - dtype=str pour garder tout tel quel ;
    - engine='python' pour tolérer des lignes un peu bancales.
    """
    # 1) Essai direct avec pandas (pandas>=2 : encoding_errors)
    for enc in ("utf-8-sig", "cp1252"):
        try:
            try:
                df = pd.read_csv(path, sep=None, engine="python",
                                 dtype=str, keep_default_na=False, na_values=[],
                                 encoding=enc, encoding_errors="replace")
            except TypeError:
                # pandas < 2 (pas de encoding_errors)
                df = pd.read_csv(path, sep=None, engine="python",
                                 dtype=str, keep_default_na=False, na_values=[],
                                 encoding=enc)
            return df
        except Exception:
            pass

    # 2) Fallback : ouvrir nous-mêmes et passer un buffer texte
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(path, "r", encoding=enc, errors="replace", newline="") as f:
                data = f.read()
            buf = io.StringIO(data)
            df = pd.read_csv(buf, sep=None, engine="python",
                             dtype=str, keep_default_na=False, na_values=[])
            return df
        except Exception:
            pass

    raise RuntimeError(f"Echec lecture CSV (encodage) : {path}")

# Normalisation de colonnes (on garde l’existant si trouvé)
def _norm_header(h: str) -> str:
    h0 = _fix_mojibake_header((h or "").strip())
    hN = unicodedata.normalize("NFD", h0)
    hN = "".join(ch for ch in hN if unicodedata.category(ch) != "Mn")
    hN = hN.lower().replace(" ", "").replace("-", "")
    # correspondances usuelles
    if hN in ("division","classe"): return "Division"
    if hN in ("nomdefamille","nom"): return "Nom de famille"
    if hN in ("prenom1","prenom","prénom","prénom1"): return "Prénom 1"
    if hN in ("courrielreprlegal","courrielrepr.légal","courrielreprelegal","mail1","email1","mail","email"):
        return "Courriel repr. légal"
    if hN in ("courrielautreprlegal","courrielautrereprlegal","mail2","email2"):
        return "Courriel autre repr. légal"
    return h0  # par défaut on ne touche pas

def _rename_columns(df: "pd.DataFrame") -> "pd.DataFrame":
    new_cols = {c: _norm_header(c) for c in df.columns}
    df = df.rename(columns=new_cols)
    # Passe 1.5 : uniformiser encore via _fix_mojibake_header (pour les variantes avec �)
    fixed_cols = {c: _fix_mojibake_header(c) for c in df.columns}
    if any(fixed_cols[c] != c for c in df.columns):
        df = df.rename(columns=fixed_cols)
    # Deuxième passe de renommage si des clés mojibake subsistent
    extra_map = {}
    for c in df.columns:
        if "Prï¿½nom" in c:
            extra_map[c] = c.replace("Prï¿½nom", "Prénom")
        if "lï¿½gal" in c:
            extra_map[c] = c.replace("lï¿½gal", "légal")
        if "lÃ©gal" in c:
            extra_map[c] = c.replace("lÃ©gal", "légal")
        if "PrÃ©nom" in c:
            extra_map[c] = c.replace("PrÃ©nom", "Prénom")
    if extra_map:
        df = df.rename(columns=extra_map)
    # Recherche tolérante des colonnes email si non trouvées
    def _norm_token(x: str) -> str:
        x = unicodedata.normalize("NFD", (x or "")).lower()
        x = "".join(ch for ch in x if unicodedata.category(ch) != "Mn")
        x = x.replace("�", "e")  # remplace le losange noir
        return re.sub(r"[^a-z0-9]+", "", x)

    cols_norm = {c: _norm_token(c) for c in df.columns}
    # candidats contenant "courriel" ET ("repr" OU "representant") ET ("legal" OU "legal")
    email_candidates = []
    for c, n in cols_norm.items():
        if "courriel" in n or "email" in n or "mail" in n:
            if ("repr" in n or "representant" in n or "parent" in n) and ("legal" in n or "legrl" in n or "legl" in n):
                email_candidates.append(c)

    # Si nos colonnes cibles n'existent pas mais qu'on a des candidats, on les mappe
    target_emails = []
    if "Courriel repr. légal" not in df.columns and email_candidates:
        target_emails.append(("Courriel repr. légal", email_candidates[0]))
    if "Courriel autre repr. légal" not in df.columns and len(email_candidates) > 1:
        target_emails.append(("Courriel autre repr. légal", email_candidates[1]))
    for dst, src in target_emails:
        if src in df.columns and dst not in df.columns:
            df[dst] = df[src]
    # S’assurer que les colonnes clés existent
    for col in ("Division", "Nom de famille", "Prénom 1"):
        if col not in df.columns:
            df[col] = ""
    # Assurer présence des colonnes emails (même vides)
    for col in ("Courriel repr. légal", "Courriel autre repr. légal"):
        if col not in df.columns:
            df[col] = ""
    # Harmoniser d'éventuelles variantes de colonnes email vers les deux colonnes cibles
    alt_cols = {
        "Courriel repr. legal": "Courriel repr. légal",
        "Courriel autre repr. legal": "Courriel autre repr. légal",
        "Email repr. legal": "Courriel repr. légal",
        "Email autre repr. legal": "Courriel autre repr. légal",
    }
    for src, dst in alt_cols.items():
        if src in df.columns and dst in df.columns:
            df[dst] = df[dst].fillna("").astype(str)
            df[src] = df[src].fillna("").astype(str)
            df[dst] = df[[dst, src]].agg(lambda x: x[0] if x[0].strip() else x[1], axis=1)
    return df

def _strip_all(df: "pd.DataFrame") -> "pd.DataFrame":
    for c in df.columns:
        if pd.api.types.is_string_dtype(df[c]):
            df[c] = (
                df[c].astype(str)
                      .map(lambda x: unicodedata.normalize("NFC", (x or "").strip()))
                      .map(lambda x: x.encode("utf-8", "ignore").decode("utf-8"))
            )
    return df

def _add_emails_column(df: "pd.DataFrame") -> "pd.DataFrame":
    """Crée/alimente la colonne 'Emails' à partir des colonnes parents connues."""
    # Diagnostic simple : combien de valeurs non vides dans les colonnes sources ?
    for _col in ("Courriel repr. légal", "Courriel autre repr. légal"):
        if _col in df.columns:
            non_empty = int((df[_col].astype(str).str.strip() != "").sum())
            print(f"[DEBUG] {_col} non vides : {non_empty}/{len(df)}")
    c1 = "Courriel repr. légal"
    c2 = "Courriel autre repr. légal"
    if c1 not in df.columns:
        df[c1] = ""
    if c2 not in df.columns:
        df[c2] = ""
    emails = []
    for v1, v2 in zip(df[c1].astype(str), df[c2].astype(str)):
        lst = _canon_email_list(v1, v2)
        emails.append(";".join(lst))
    df["Emails"] = emails
    # Fallback : si 'Emails' est toujours vide mais qu'une autre colonne unique 'Email*' existe, l'utiliser
    if int((df["Emails"].astype(str).str.strip() != "").sum()) == 0:
        for alt in df.columns:
            n = alt.strip().lower()
            if n in ("email", "emails", "courriels", "email parents", "emails parents"):
                df["Emails"] = df[alt].astype(str).fillna("").map(lambda s: ";".join(_canon_email_list(s)))
                break
    return df

def _canon_div(s: str) -> str:
    s = (s or "").strip()
    m = re.match(r'^=\s*"(.+)"\s*$', s)  # Excel export "= "..." "
    if m:
        s = m.group(1)
    sN = unicodedata.normalize("NFD", s).upper()
    sN = "".join(ch for ch in sN if unicodedata.category(ch) != "Mn")
    sN = sN.replace("ÈME","E").replace("EME","E")
    sN = re.sub(r"[\s\-.]+","", sN)
    return sN

def _get_message_text() -> str | None:
    """
    Récupère un message « Message aux parents » à répliquer sur toutes les lignes, si dispo.
    Sources (dans l'ordre) :
      - variable d'environnement EVALNAT_MESSAGE_TEXT
      - variable d'environnement MESSAGE_TEXT
      - un fichier texte UTF-8 présent dans le CWD: message.txt ou message_parents.txt
    Les séquences `\n` sont transformées en retours à la ligne réels.
    """
    # 1) Env vars
    msg = os.environ.get("EVALNAT_MESSAGE_TEXT") or os.environ.get("MESSAGE_TEXT")
    # 2) Fichiers connus
    if not msg:
        for candidate in ("message.txt", "message_parents.txt"):
            p = Path.cwd() / candidate
            if p.exists():
                try:
                    msg = p.read_text(encoding="utf-8", errors="replace")
                    break
                except Exception:
                    pass
    if not msg:
        return None
    # Normaliser les fins de ligne (convertir les séquences littérales "\n")
    msg = msg.replace("\r\n", "\n").replace("\r", "\n")
    msg = msg.encode("utf-8", "ignore").decode("utf-8")  # assainir
    msg = re.sub(r"\\n", "\n", msg)  # transformer le texte "\n" en vraies nouvelles lignes
    return msg.strip()

def merge_files(csv_paths: list[str]) -> None:
    """
    Fusion basique des exports SIECLE.
    Écrit parents_4e_merged.csv (UTF-8, séparateur ';') dans le CWD.
    """
    if not csv_paths:
        print("[ERREUR] Aucun chemin CSV fourni à merge_files().")
        sys.exit(2)

    frames = []
    for p in csv_paths:
        path = Path(p)
        if not path.exists():
            print("[ERREUR] CSV introuvable:", path)
            sys.exit(2)
        try:
            df = _read_csv_robust(path)
            df = _rename_columns(df)
            df = _strip_all(df)
            frames.append(df)
            print("[OK]  ->", path.name, f"({len(df)} lignes)")
        except Exception as e:
            print("[ERREUR] ->", path.name, ":", e)
            sys.exit(2)

    if not frames:
        print("[ERREUR] Aucun CSV lisible.")
        sys.exit(2)

    # Concat + suppression de doublons “souples”
    all_df = pd.concat(frames, ignore_index=True)

    # Clé douce division/nom/prénom pour éviter les doublons stricts
    def keyrow(r):
        div = _canon_div(r.get("Division",""))
        nom = unicodedata.normalize("NFD", r.get("Nom de famille","")).lower()
        nom = "".join(ch for ch in nom if unicodedata.category(ch) != "Mn")
        pre = unicodedata.normalize("NFD", r.get("Prénom 1","")).lower()
        pre = "".join(ch for ch in pre if unicodedata.category(ch) != "Mn")
        return (div, re.sub(r"[^a-z0-9]","",nom), re.sub(r"[^a-z0-9]","",pre))

    all_df["_key"] = all_df.apply(keyrow, axis=1)
    all_df = all_df.drop_duplicates(subset=["_key"]).drop(columns=["_key"], errors="ignore")

    # Sort par Division puis Nom/Prénom si dispo
    sort_cols = [c for c in ("Division","Nom de famille","Prénom 1") if c in all_df.columns]
    if sort_cols:
        all_df = all_df.sort_values(by=sort_cols, kind="stable")

    # Injection optionnelle d'un message commun à toutes les lignes
    # (utile pour le publipostage / "Message aux parents" ou "CorpsMessage")
    msg = _get_message_text()
    if msg:
        # Créer les colonnes si absentes
        for col in ("Message aux parents", "CorpsMessage"):
            if col not in all_df.columns:
                all_df[col] = ""

        # Ne remplir que les cellules vides (NaN ou chaînes blanches)
        for col in ("Message aux parents", "CorpsMessage"):
            s = all_df[col].astype(str)
            mask = s.isna() | (s.str.strip() == "") | (s.str.lower().str.strip().isin(("nan", "none", "nul", "null")))
            all_df.loc[mask, col] = msg

    # Ajout colonne 'Emails' (dédupliquée) à partir des colonnes parents
    all_df = _add_emails_column(all_df)

    # Diagnostic : entêtes finales et exemple de 3 lignes "Emails"
    print("[DEBUG] Entêtes finales :", list(all_df.columns))
    print("[DEBUG] échantillon Emails :", list(all_df["Emails"].head(3)))

    # Écriture UTF-8-SIG (compatible Excel/Windows), séparateur ';'
    out = Path.cwd() / "parents_4e_merged.csv"
    try:
        all_df.to_csv(out, index=False, sep=";", encoding="utf-8-sig", lineterminator="\n")
    except Exception as e:
        # Fallback via csv.writer si jamais
        print("[ERREUR] pandas.to_csv a échoué, fallback csv.writer :", e)
        with open(out, "w", encoding="utf-8-sig", newline="") as g:
            w = csv.writer(g, delimiter=";")
            w.writerow(list(all_df.columns))
            for row in all_df.itertuples(index=False, name=None):
                clean = []
                for x in row:
                    s = "" if x is None else str(x)
                    # normaliser + supprimer les caractères non encodables
                    s = unicodedata.normalize("NFC", s)
                    s = s.encode("utf-8", "ignore").decode("utf-8")
                    clean.append(s)
                w.writerow(clean)

    nb_total = len(all_df)
    nb_with_mail = int((all_df.get("Emails","").astype(str).str.len() > 0).sum())
    print(f"[INFO] Emails non vides : {nb_with_mail}/{nb_total}")

    print("[OK]  -> fusion écrite :", out)

# Exécution directe (debug/dev)
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Fusion exports SIECLE (parents)")
    ap.add_argument("csvs", nargs="+", help="Chemins CSV SIECLE")
    args = ap.parse_args()
    merge_files(args.csvs)