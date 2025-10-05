#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Ouvre une fenêtre de rédaction Thunderbird pour chaque ligne d'un CSV.

CSV attendu (au moins ces colonnes) :
- Emails            : une ou plusieurs adresses séparées par ; (ou ,)
- Objet             : objet du message
- CorpsMessage      : corps du message (les retours à la ligne sont gérés)
- Attachments       : chemins absolus des PJ, séparés par ; (facultatif)
  (fallback) PJ_francais, PJ_math : si Attachments est vide, on les concatène

Exemples :
python3 tb_mailmerge_open_compose_mac.py --csv "/Users/julien/Downloads/mailmerge_4e.csv"

Options utiles :
--dry-run                  : n’ouvre rien, affiche seulement ce qui serait fait
--throttle 1.2             : pause entre fenêtres (sec)
--start 1 --count 20       : n’ouvrir que 20 messages à partir de la ligne 1
--only-list restants.txt   : n’ouvrir que pour une liste "NOM PRENOM" (une par ligne)
"""

import argparse
import csv
import os
import shlex
import subprocess
import sys
import time
from pathlib import Path

def read_csv(path):
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append(row)
    return rows

def to_file_uri(p):
    try:
        return Path(p).expanduser().resolve().as_uri()
    except Exception:
        # Dernier recours : préfixer file://
        q = os.path.abspath(os.path.expanduser(p))
        return "file://" + q.replace(" ", "%20")

def split_multi(value):
    if not value:
        return []
    # Autoriser séparateurs ; ou ,
    parts = []
    for chunk in str(value).split(";"):
        parts.extend(chunk.split(","))
    return [x.strip() for x in parts if x.strip()]

def escape_compose_value(val):
    """
    Échappe proprement pour le paramètre -compose :
    - entouré de guillemets "
    - échappe " en \"
    - remplace les retours ligne par \n (géré par TB)
    """
    if val is None:
        val = ""
    s = str(val).replace("\\", "\\\\").replace('"', '\\"')
    s = s.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n")
    return f"\"{s}\""

def build_compose_cmd(to_addrs, subject, body, attachments):
    # Thunderbird accepte plusieurs champs dans un seul -compose, séparés par des virgules,
    # les valeurs étant entre guillemets.
    parts = []
    # Destinataires : séparés par virgule dans le champ 'to'
    to_field = ",".join(to_addrs)
    parts.append(f"to={escape_compose_value(to_field)}")
    parts.append(f"subject={escape_compose_value(subject)}")
    parts.append(f"body={escape_compose_value(body)}")

    if attachments:
        # Attachements : liste d’URI file://, séparés par virgules dans UNE valeur
        atts_uri = ",".join(attachments)
        parts.append(f"attachment={escape_compose_value(atts_uri)}")

    compose_str = ",".join(parts)
    # Commande macOS : ouvrir Thunderbird avec des arguments
    # Nota: open -a "Thunderbird" --args -compose "<params>"
    cmd = ["open", "-a", "Thunderbird", "--args", "-compose", compose_str]
    return cmd

def parse_only_list(path):
    if not path:
        return None
    keep = set()
    with open(path, encoding="utf-8") as f:
        for line in f:
            t = line.strip()
            if t:
                keep.add(t)
    return keep

def main():
    ap = argparse.ArgumentParser(description="Ouvre des rédactions Thunderbird depuis un CSV (macOS).")
    ap.add_argument("--csv", required=True, help="Chemin du CSV mail-merge")
    ap.add_argument("--throttle", type=float, default=1.0, help="Pause (s) entre ouvertures de fenêtres")
    ap.add_argument("--dry-run", action="store_true", help="N’ouvre pas Thunderbird, affiche seulement")
    ap.add_argument("--start", type=int, default=1, help="Index de départ (1 = première ligne)")
    ap.add_argument("--count", type=int, default=0, help="Nombre de messages à traiter (0 = tout)")
    ap.add_argument("--only-list", help="Fichier texte: une ligne = 'NOM PRENOM' pour restreindre l’envoi")
    args = ap.parse_args()

    rows = read_csv(args.csv)
    only_list = parse_only_list(args.only_list)

    # Fenêtre d’indexation simple
    start_idx = max(1, args.start)
    end_idx = len(rows) if args.count in (0, None) else min(len(rows), start_idx - 1 + args.count)

    total, opened, skipped, warned = 0, 0, 0, 0
    for i, row in enumerate(rows, 1):
        if i < start_idx or i > end_idx:
            continue

        total += 1

        nom = (row.get("Nom","") or "").strip()
        prenom = (row.get("Prénom","") or row.get("Prenom","") or "").strip()
        classe = (row.get("Classe","") or row.get("Division","") or "").strip()
        display_key = f"{classe} | {nom} {prenom}".strip()

        if only_list:
            # Format attendu dans only-list: "NOM PRENOM" (NOM en maj, prénom respecté)
            key_check = f"{nom.upper()} {prenom}"
            if key_check not in only_list:
                skipped += 1
                continue

        # Destinataires
        emails_raw = (row.get("Emails","") or "").strip()
        to_addrs = split_multi(emails_raw)
        if not to_addrs:
            print(f"[SKIP] {display_key} : Emails manquants", file=sys.stderr)
            skipped += 1
            continue

        # Objet & Corps
        subject = (row.get("Objet","") or "Évaluations nationales").strip()
        body = (row.get("CorpsMessage","") or "Madame, Monsieur,\n\nVeuillez trouver en pièces jointes…").strip()

        # Pièces jointes
        att_paths = []
        if row.get("Attachments",""):
            att_paths = [p for p in split_multi(row["Attachments"]) if p]
        else:
            # Fallback : concat PJ_francais / PJ_math si présents
            pf = (row.get("PJ_francais","") or "").strip()
            pm = (row.get("PJ_math","") or "").strip()
            if pf and pm:
                att_paths = [pf, pm]
            elif pf:
                att_paths = [pf]
            elif pm:
                att_paths = [pm]

        # Vérification des PJ (on continue quand même si manquantes, mais on prévient)
        att_uris = []
        for p in att_paths:
            if not p:
                continue
            if not os.path.exists(os.path.expanduser(p)):
                print(f"[AVERTISSEMENT] PJ introuvable pour {display_key} : {p}", file=sys.stderr)
                warned += 1
                # On peut choisir de ne pas l’ajouter
                continue
            att_uris.append(to_file_uri(p))

        cmd = build_compose_cmd(to_addrs, subject, body, att_uris)

        if args.dry_run:
            print(f"[DRY RUN] {display_key}")
            print("          TO:", ", ".join(to_addrs))
            print("          OBJ:", subject)
            print("          PJ :", ", ".join(att_paths) if att_paths else "(aucune)")
            print("          CMD:", " ".join(shlex.quote(c) for c in cmd))
        else:
            try:
                subprocess.run(cmd, check=True)
                print(f"[OK] Fenêtre ouverte : {display_key} → {', '.join(to_addrs)}")
                opened += 1
                time.sleep(args.throttle)
            except subprocess.CalledProcessError as e:
                print(f"[ERREUR] {display_key} : {e}", file=sys.stderr)

    print(f"\nRésumé : total={total}, ouvertes={opened}, ignorées={skipped}, avertissements PJ={warned}")
    if args.dry_run:
        print("Mode DRY RUN : aucune fenêtre Thunderbird n’a été ouverte.")

if __name__ == "__main__":
    main()