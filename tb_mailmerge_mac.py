#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
tb_mailmerge_mac.py — version binaire ThunderBird (multi-brouillons OK)

Usage typique :
  python3 tb_mailmerge_mac.py \
    --csv "/Users/julien/Downloads/mailmerge_4e.csv" \
    --sleep 0.7

Options utiles :
  --dry-run               : n’ouvre rien, affiche ce qu’il ferait
  --force-default-body    : force le message standard pour tous
  --tb-binary <chemin>    : chemin explicite vers le binaire Thunderbird
                            (sinon tentative auto : /Applications/Thunderbird.app/Contents/MacOS/thunderbird)
Colonnes par défaut : Emails | Objet | Message | PJ_francais | PJ_math
"""

import argparse
import csv
import os
import re
import shlex
import subprocess
import sys
import time

DEFAULT_MESSAGE = (
    "Madame, Monsieur,\n\n"
    "Veuillez trouver en pièces jointes les comptes rendus des évaluations nationales passées par vos enfants.\n"
    "Les enseignants reviendront dessus lors des remises de bulletins. Vous pourrez poser toutes les questions s'y rapportant lors de ce rendez-vous.\n\n"
    "Bien cordialement,\n"
    "Pour l'équipe de direction,\n"
    "Julien Texier"
)

def parse_args():
    p = argparse.ArgumentParser(description="Ouvre des brouillons Thunderbird (macOS).")
    p.add_argument("--csv", required=True, help="Chemin du CSV d'entrée.")
    p.add_argument("--col-emails", default="Emails", help="Nom de la colonne Emails (défaut: Emails).")
    p.add_argument("--col-subject", default="Objet", help="Nom de la colonne Objet (défaut: Objet).")
    p.add_argument("--col-body", default="Message", help="Nom de la colonne Message (défaut: Message).")
    p.add_argument("--col-att1", default="PJ_francais", help="Nom de la colonne PJ #1 (défaut: PJ_francais).")
    p.add_argument("--col-att2", default="PJ_math", help="Nom de la colonne PJ #2 (défaut: PJ_math).")
    p.add_argument("--sleep", type=float, default=0.6, help="Pause entre brouillons (sec).")
    p.add_argument("--limit", type=int, default=None, help="Limiter au N premières lignes.")
    p.add_argument("--skip", type=int, default=0, help="Ignorer les N premières lignes.")
    p.add_argument("--dry-run", action="store_true", help="N'exécute pas, affiche les commandes.")
    p.add_argument("--force-default-body", action="store_true",
                   help="Ignore la colonne Message et utilise le message standard.")
    p.add_argument("--tb-binary", default=None, help="Chemin du binaire Thunderbird.")
    return p.parse_args()

def find_tb_binary(user_choice=None):
    if user_choice:
        return user_choice
    # Chemin par défaut d’installation standard macOS
    cand = "/Applications/Thunderbird.app/Contents/MacOS/thunderbird"
    if os.path.isfile(cand):
        return cand
    # Éventuel chemin dans /Applications (local user)
    cand2 = os.path.expanduser("~/Applications/Thunderbird.app/Contents/MacOS/thunderbird")
    if os.path.isfile(cand2):
        return cand2
    raise FileNotFoundError("Binaire Thunderbird introuvable. Fournis-le via --tb-binary.")

def ensure_tb_running(dry_run=False):
    # Lance l’app si elle n’est pas ouverte (sans args)
    cmd = ["open", "-ga", "Thunderbird"]
    if dry_run:
        print("[DRY-RUN] " + " ".join(shlex.quote(c) for c in cmd))
        return
    subprocess.run(cmd, check=False)
    # petite pause pour laisser l’app démarrer
    time.sleep(1.2)

def norm_recipients(raw: str) -> str:
    if raw is None:
        return ""
    s = raw.strip()
    if not s:
        return ""
    s = s.replace(";", ",")
    s = re.sub(r"\s*,\s*", ",", s)
    # Nettoyage éventuel de "Nom Prénom <mail>"
    parts = []
    for token in s.split(","):
        token = token.strip()
        m = re.search(r"<([^>]+)>", token)
        parts.append(m.group(1) if m else token)
    return ",".join(parts)

def escape_compose_value_single_quotes(val: str) -> str:
    if val is None:
        val = ""
    s = str(val).replace("'", "''")
    return f"'{s}'"

def ensure_abs(path: str) -> str:
    if not path:
        return ""
    return os.path.abspath(os.path.expanduser(path))

def build_compose_arg(to_field: str, subject: str, body: str, attachments_paths):
    parts = [
        f"to={escape_compose_value_single_quotes(to_field)}",
        f"subject={escape_compose_value_single_quotes(subject)}",
        f"body={escape_compose_value_single_quotes(body)}",
    ]
    att_ok = []
    for p in attachments_paths or []:
        if not p:
            continue
        ap = ensure_abs(p)
        if os.path.isfile(ap):
            att_ok.append(ap)
        else:
            print(f"  [WARN] PJ introuvable : {ap}", file=sys.stderr)
    if att_ok:
        parts.append(f"attachment={escape_compose_value_single_quotes(','.join(att_ok))}")
    return ",".join(parts)

def open_draft_with_binary(tb_bin, to_field, subject, body, attachments, dry_run=False):
    compose_str = build_compose_arg(to_field, subject, body, attachments)
    cmd = [tb_bin, "-compose", compose_str]
    if dry_run:
        print("[DRY-RUN] " + " ".join(shlex.quote(c) for c in cmd))
        return 0
    try:
        # Ne pas bloquer : Popen, sans wait
        subprocess.Popen(cmd)
        return 0
    except Exception as e:
        print(f"[ERR] Ouverture compose a échoué : {e}", file=sys.stderr)
        return 1

def read_csv(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))

def main():
    args = parse_args()
    rows = read_csv(args.csv)
    total = len(rows)
    if args.skip:
        rows = rows[args.skip:]
    if args.limit is not None:
        rows = rows[:args.limit]

    print(f"[INFO] Lignes CSV lues : {total} | traitées : {len(rows)} (skip={args.skip}, limit={args.limit})")
    print(f"[INFO] Colonnes : Emails='{args.col_emails}', Objet='{args.col_subject}', Message='{args.col_body}', PJ1='{args.col_att1}', PJ2='{args.col_att2}'")
    if args.force_default_body:
        print("[INFO] Message par défaut forcé (--force-default-body).")

    tb_bin = find_tb_binary(args.tb_binary)
    ensure_tb_running(dry_run=args.dry_run)

    sent = 0
    errors = 0

    for i, r in enumerate(rows, 1):
        raw_to = r.get(args.col_emails, "") or ""
        to_field = norm_recipients(raw_to)
        subject = (r.get(args.col_subject, "") or "").strip()
        csv_body = (r.get(args.col_body, "") or "").strip()
        body = DEFAULT_MESSAGE if (args.force_default_body or not csv_body) else csv_body

        att1 = (r.get(args.col_att1, "") or "").strip()
        att2 = (r.get(args.col_att2, "") or "").strip()
        attachments = [p for p in [att1, att2] if p]

        label = subject if subject else to_field
        print(f"[{i}/{len(rows)}] → {label}")

        if not to_field:
            print("  [WARN] Emails vide → brouillon ignoré.")
            errors += 1
            continue

        rc = open_draft_with_binary(tb_bin, to_field, subject, body, attachments, dry_run=args.dry_run)
        if rc == 0:
            sent += 1
        else:
            errors += 1

        if args.sleep and not args.dry_run:
            time.sleep(args.sleep)

    print("\n===== RÉCAP =====")
    print(f"Brouillons ouverts : {sent}")
    print(f"Avertissements/Erreurs : {errors}")
    if args.dry_run:
        print("(DRY-RUN : aucune fenêtre Thunderbird n'a été ouverte)")

if __name__ == "__main__":
    main()