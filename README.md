# 🧩 Publipostage ÉvalNat — V1

**Publipostage ÉvalNat** est une application multiplateforme (Mac / Windows) permettant d’automatiser la préparation des **évaluations nationales** (6e – 5e – 4e) à partir des exports PDF et fichiers SIECLE.

## 🚀 Fonctionnalités principales

- **Découpage automatique** des PDF d’évaluations (français / mathématiques) par élève  
- **Fusion automatique** avec le fichier parents issu de SIECLE (`exportCSVExtractionClasse.csv`)  
- **Génération du publipostage** pour Thunderbird ou autre client mail  
- **Message personnalisé aux parents** (nouvel onglet dédié dans la V1)  
- **Interface graphique** claire et simple  
- **Barre de progression** et mode *verbose* optionnel pour le suivi des étapes  
- **Aucune dépendance externe** : l’application embarque ses scripts et dépendances Python

## 🖥️ Compatibilité

- macOS 12 (Monterey) ou supérieur  
- Windows 10 / 11 (64 bits)

## 📦 Installation

### 🧑‍💻 Méthode la plus simple

1. Téléchargez la dernière version depuis l’onglet **Releases** du dépôt.  
2. Décompressez le dossier téléchargé.  
3. Lancez :
   - Sur **Mac** : `EvalNat-Publipostage.app`
   - Sur **Windows** : `EvalNat-Publipostage.exe`

> ⚠️ Sur macOS, si l’app est bloquée par Gatekeeper, faites clic droit → *Ouvrir* → *Autoriser*.

PublipostageEvalNat/
├── app_gui.py                 ← Interface principale
├── pipeline_evalnat.py        ← Pipeline principal (split + merge + mail)
├── split_4C.py                ← Découpage PDF OCR
├── merge_parents_4e.py        ← Fusion des CSV parents
├── tb_mailmerge_mac.py        ← Génération mails (Thunderbird)
├── tb_mailmerge_open_compose_mac.py
├── build_mailmerge_4e_from_merged_v5.py
├── normalize.py               ← Nettoyage des noms/accents
├── check_links.py             ← Vérification des chemins
└── README.md                  ← Ce fichier


## 🧭 Utilisation rapide
	1.	**Onglet 1 — Paramètres
		-	Sélectionnez la classe (6A, 4B…), l’année et les fichiers source.
		-	Option : cochez Mode verbose pour voir les logs détaillés.
	2.	**Onglet 2 — Publipostage
		-	Le pipeline découpe, fusionne et prépare les fichiers pour l’envoi.
	3.	**Onglet 3 — Message aux parents
		-	Rédigez le message commun à insérer dans chaque mail.
	4.	**Cliquez sur “C’est parti !”
		-	Suivez la progression dans la barre prévue à cet effet.

## 💡 Astuces
	•	L’OCR est appliqué automatiquement si le PDF est image uniquement.
	•	Les accents et prénoms composés sont normalisés automatiquement.
	•	Les fichiers produits suivent la convention :

  Classe_NOM_prénom_Discipline_Année.pdf

  ## 🧱 Distribution

L’application peut être distribuée simplement en transmettant le dossier dist/ :
	•	EvalNat-Publipostage.app (Mac)
	•	EvalNat-Publipostage.exe (Windows) (demain)

Aucune installation de Python n’est requise.
## Thunderbird est indispensable. 

## 🏷️ Version

V1 stable — octobre 2025
Fonctionnalités : GUI complète + pipeline intégré + message parents + barre de progression.

### 🧰 Méthode avancée (développeurs)

Cloner le dépôt et lancer en mode développement :

```bash
git clone https://github.com/julientexier86/PublipostageEvalNat.git
cd PublipostageEvalNat
python3 app_gui.py
