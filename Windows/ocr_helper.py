# -*- coding: utf-8 -*-
from __future__ import annotations
import os, sys, subprocess, shutil
from pathlib import Path

def _which(cmd: str) -> str | None:
    p = shutil.which(cmd)
    return p if p else None

def _run_ok(cmd: list[str]) -> bool:
    try:
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        return True
    except Exception:
        return False

def _set_tessdata_prefix_if_needed() -> None:
    """
    Sur Windows, si Tesseract est installé via UB Mannheim, TESSDATA_PREFIX
    est souvent sous 'C:\\Program Files\\Tesseract-OCR\\tessdata'.
    On n’écrase pas la var si elle existe déjà.
    """
    if os.environ.get("TESSDATA_PREFIX"):
        return
    common = [
        r"C:\Program Files\Tesseract-OCR\tessdata",
        r"C:\Program Files (x86)\Tesseract-OCR\tessdata",
    ]
    for p in common:
        if Path(p).exists():
            os.environ["TESSDATA_PREFIX"] = p
            break

def _try_install_windows_tesseract() -> bool:
    """
    Tentative d’installation silencieuse de Tesseract sur Windows si 'winget' ou 'choco' est dispo.
    On reste best-effort : si rien n’est possible, on retourne False (mais on n’échoue pas bruyamment).
    """
    # winget (UB Mannheim est souvent le plus simple)
    if _which("winget"):
        # Paquets possibles : 'UB-Mannheim.TesseractOCR' ou 'TesseractOCR.Tesseract'
        # On tente Mannheim d’abord, puis l’officiel.
        if _run_ok(["winget", "install", "-e", "--id", "UB-Mannheim.TesseractOCR", "--silent"]):
            return True
        if _run_ok(["winget", "install", "-e", "--id", "TesseractOCR.Tesseract", "--silent"]):
            return True
    # chocolatey
    if _which("choco"):
        if _run_ok(["choco", "install", "-y", "tesseract"]):
            return True
    return False

def _try_install_ocrmypdf() -> bool:
    """
    ocrmypdf ne s’installe pas de façon fiable depuis un exécutable gelé.
    On tente via 'pip' si accessible. Best-effort, sinon False.
    """
    py = sys.executable or "python"
    try:
        subprocess.run([py, "-m", "pip", "install", "--upgrade", "ocrmypdf"], check=True)
        return True
    except Exception:
        return False

def _try_install_ghostscript() -> bool:
    """
    ocrmypdf tire souvent parti de Ghostscript pour le nettoyage.
    On tente winget/choco, best-effort.
    """
    if _which("gs"):
        return True
    if _which("winget"):
        if _run_ok(["winget", "install", "-e", "--id", "ArtifexSoftware.Ghostscript", "--silent"]):
            return True
    if _which("choco"):
        if _run_ok(["choco", "install", "-y", "ghostscript"]):
            return True
    return False

def ensure_ocr_stack(force_install: bool = False, base_dir: Path | None = None, lang: str = "fra") -> tuple[bool, str]:
    """
    Vérifie la présence minimale pour l’OCR :
      - tesseract (binaire)
      - ocrmypdf (pour l’orchestration, si dispo)
    Sous Windows, tente une installation automatique best-effort si force_install=True.

    Retourne (ok, message).
    - ok=True → on a au moins Tesseract; si ocrmypdf manque, le pipeline fonctionnera
               quand même (il préviendra juste que l’OCR auto est indisponible).
    - ok=False → même Tesseract est introuvable (OCR non disponible).
    """
    # 1) Tesseract
    tesseract = _which("tesseract")
    if not tesseract and os.name == "nt" and force_install:
        # tentative d'installation silencieuse
        if _try_install_windows_tesseract():
            tesseract = _which("tesseract")
    if not tesseract:
        return (False, "OCR indisponible : 'tesseract' introuvable. "
                       "Installez Tesseract (winget/choco) ou décochez --auto-ocr.")
    # TESSDATA_PREFIX si possible
    if os.name == "nt":
        _set_tessdata_prefix_if_needed()

    # 2) Paquet langue (best-effort : on suppose 'fra' déjà livré par les builds Mannheim)
    # On n’essaie pas de télécharger les .traineddata ici, ça dépend d’internet & droits.
    # Si l’utilisateur manque une langue, tesseract indiquera l’erreur au moment de l’ocr.

    # 3) ocrmypdf (optionnel)
    ocrmypdf = _which("ocrmypdf")
    if not ocrmypdf and force_install:
        _try_install_ghostscript()  # best-effort
        if _try_install_ocrmypdf():
            ocrmypdf = _which("ocrmypdf")

    if ocrmypdf:
        return (True, "Pile OCR prête : tesseract + ocrmypdf disponibles.")
    else:
        # Dégradé : on a tesseract (ok) mais pas ocrmypdf → le pipeline pourra continuer
        # (il affichera juste un warning si --auto-ocr).
        return (True, "Pile OCR partielle : tesseract OK, ocrmypdf indisponible (OCR auto désactivé).")