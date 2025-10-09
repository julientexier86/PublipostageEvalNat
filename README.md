# PublipostageEVALNAT

Outil clef-en-main pour préparer l’envoi des **comptes rendus des évaluations nationales** aux responsables légaux :

1. **Découpage PDF** (1 PDF source → 1 PDF par élève, FR/Math)
2. **Fusion SIECLE** (récupère les emails parents depuis l’export CSV)
3. **Génération CSV Mail Merge** (objet + message + 2 pièces jointes)
4. **Ouverture automatique des brouillons** dans **Thunderbird** (extension Mail Merge)

Fonctionne sur **macOS** et **Windows**. Aucune donnée ne quitte votre machine.

---

## 📦 Téléchargement (utilisateurs)

### macOS (recommandé)
- Téléchargez la dernière **Release** sur GitHub : `PublipostageEVALNAT-macOS-*.zip`
- Décompressez → ouvrez `PublipostageEVALNAT.app`
- Si macOS bloque l’ouverture : clic droit → **Ouvrir** (une fois)

### Windows
- Téléchargez `PublipostageEvalNat-Setup-*.exe`
- Lancez l’installeur
- Ouvrez **PublipostageEVALNAT** depuis le menu Démarrer

> ⚠️ Pour l’ouverture des brouillons, **Thunderbird** doit être installé (avec l’extension **Mail Merge**).  
> OCR (optionnel sous Windows) : **Tesseract** avec la langue **fra** pour les PDF scannés.

---

## 🧭 Guide rapide (interface)

L’application comporte 5 onglets :

1) **Contexte**  
- **Classe** (ex. `5B`)  
- **Année scolaire** (ex. `2025-2026`)  
- **Mode verbose** (option) : affiche le journal du pipeline dans cet onglet

2) **Découpage PDF**  
- **PDF source** (si vous souhaitez découper)  
- **Dossier de sortie** (les PDF par élève y seront créés)  
- Option **Ne pas découper** si vous avez déjà les PDF individuels  
- **Langue OCR** (ex. `fra`) — utile uniquement pour des PDF scannés

3) **Récupération mails parents**  
- **Export SIECLE (CSV)** : choisissez le fichier CSV  
  - Colonnes attendues (noms possibles) :  
    - `Division`  
    - `Nom de famille` (ou `Nom 1`)  
    - `Prénom 1` (ou `Prenom 1`)  
    - Emails : `Courriel repr. légal` et/ou `Courriel autre repr. légal`  
  - L’app gère les CSV en `;` (SIECLE) et corrige l’encodage si besoin.

4) **Message aux parents**  
- **Objet (modèle)**, ex. : `Evaluations nationales - {NOM} {Prénom} ({Classe})`  
- **Message** (clic droit → coller possible)

5) **Publipostage**  
- **Ouvrir automatiquement les brouillons Thunderbird** (ON/OFF)  
- **Options avancées** (cocher pour afficher) :  
  - **Chemin de Thunderbird** (si l’app ne le détecte pas)  
  - **Limit / Skip / Sleep (s)** pour piloter Mail Merge

Cliquez sur **“C’est parti”** pour lancer le pipeline. Une barre de progression s’affiche en haut.

---

## 📨 Détails des fichiers attendus

- **PDF source** : export des évaluations nationales (FR/Math) de la classe  
- **CSV SIECLE** : export contenant au moins Division/Nom/Prénom et les emails des représentants légaux

La sortie principale est un **CSV Mail Merge** avec colonnes :
- `Emails` (listes séparées par `;` quand 2 responsables)  
- `Objet` (d’après le modèle)  
- `CorpsMessage` (votre texte commun)  
- `PJ_francais` / `PJ_math` (chemins des PDF trouvés)

---

## 🛠️ Dépannage (FAQ)

- **Thunderbird non détecté (Win)**  
  → Saisissez le chemin dans Options avancées (ex. `C:\Program Files\Mozilla Thunderbird\thunderbird.exe`).

- **0 brouillon créé (Win)**  
  → Vérifiez la colonne `Emails` dans `*_with_emails.csv`. Elle ne doit pas être vide.  
  → L’export SIECLE doit bien contenir au moins un email par élève.

- **PDF scanné (pas de texte)**  
  → Sous Windows, installez **Tesseract** + le pack **fra** (ou décochez l’OCR et faites un OCR manuel).  
  → Sous macOS, l’OCR auto est géré si `ocrmypdf` est présent.

- **CSV encodé bizarrement (accents)**  
  → L’app tente de corriger automatiquement (`latin-1` → `utf-8`).  
  → Si problème persiste, ouvrez le CSV dans un tableur et exportez en **CSV UTF-8 ;** (point-virgule).

- **Journal / logs**  
  - Dans l’onglet **Contexte**, cochez **Mode verbose** pour afficher le journal.  
  - macOS : log fichier `~/Library/Logs/PublipostageEVALNAT.log`.

---

## 👩‍💻 Développement / Build (pour mainteneurs)

### 1) Prérequis
- Python 3.11+ (testé 3.13)
- macOS : Xcode CLT recommandés
- Windows : Inno Setup (si vous faites l’installeur), Tesseract (optionnel)

```bash
python -m venv .venv
source .venv/bin/activate        # (macOS/Linux)
# ou: .\.venv\Scripts\activate   # (Windows PowerShell)
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller
```

### 2) Build “pipeline” (binaire CLI)
```bash
pyinstaller pipeline_evalnat.py --name pipeline_evalnat --onefile --console
```
→ Sortie : `dist/pipeline_evalnat` (macOS/Linux) ou `dist/pipeline_evalnat.exe` (Win)

### 3) Build **app macOS** (GUI + pipeline embarqué)
Deux options :

**A. Simple (2 étapes)**  
- Construire le pipeline (étape 2)  
- Puis :
  ```bash
  pyinstaller app_gui.py --windowed --name PublipostageEVALNAT \
    --add-binary "dist/pipeline_evalnat:."
  ```
  → L’app tentera d’embarquer le binaire.  
  → Zippez `dist/PublipostageEVALNAT.app` pour la release.

**B. Via le fichier `.spec`** (recommandé si fourni dans le dépôt)  
- Utilisez le `.spec` livré pour garantir l’emplacement du binaire dans `Contents/MacOS/`  
  ```bash
  pyinstaller app_mac.spec
  ```

### 4) Build **installeur Windows**
- Générez d’abord `dist/pipeline_evalnat.exe` & `dist/PublipostageEVALNAT.exe`  
- Ouvrez le fichier **Inno Setup** (`*.iss`) inclus dans le dépôt  
- Build → obtient `PublipostageEvalNat-Setup-*.exe`

> **Note** : on ne commit **jamais** `dist/` ni les `.exe/.app` dans Git. On publie les binaires via **GitHub Releases**.

---

## 🚀 Publier une Release GitHub

```bash
# Sur la branche main après merge:
git tag -a v1.4 -m "PublipostageEVALNAT v1.4"
git push --tags
```

- GitHub → **Releases** → **Draft a new release** → Tag `v1.4`  
- Attachez :
  - `PublipostageEVALNAT-macOS-*.zip`
  - `PublipostageEvalNat-Setup-*.exe`
  - (optionnel) `SHA256SUMS.txt`  
- Renseignez les notes de version (changements, correctifs, etc.)

---

## 🗂️ Structure du dépôt (résumé)

```
PublipostageEvalNat/
├─ app_gui.py                  # Interface Tkinter
├─ pipeline_evalnat.py         # Pipeline CLI (split/merge/mailmerge/TB)
├─ merge_parents_4e.py         # Fusion exports SIECLE (emails)
├─ build_mailmerge_*.py        # Construction CSV Mail Merge
├─ ocr_helper.py               # Détection/installation OCR (Win)
├─ requirements.txt
├─ app_mac.spec                # Build macOS .app (si fourni)
├─ windows_installer.iss       # Script Inno Setup (si fourni)
└─ README.md
```

---

## 🔒 Confidentialité

- Aucune donnée n’est envoyée en ligne par l’application.  
- Les CSV et PDF sont lus et produits **localement**.

---

## 📄 Licence

MIT 

---

## 📝 Historique

- v1.4 : macOS **single app** (GUI + pipeline embarqué), options TB avancées, corrections encodage CSV, sujets dynamiques `{NOM} {Prénom} ({Classe})`, barre de progression, journal “verbose” dans l’onglet Contexte.  
- v1.3 : première release stable multi-plateforme.

---

### Remerciements
Merci aux collègues pour les retours et tests ⚡️
