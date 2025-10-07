; ============================
; Publipostage EvalNat - Installeur
; ============================

[Setup]
AppId={{8A6D2C19-3B49-4A48-8B0B-EE7C3E51F9B7}}
AppName=Publipostage EvalNat
AppVersion=1.0.0
AppPublisher=Ton établissement
DefaultDirName={autopf}\PublipostageEvalNat
DefaultGroupName=Publipostage EvalNat
DisableDirPage=no
DisableProgramGroupPage=no
LicenseFile=
OutputBaseFilename=PublipostageEvalNat_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
ArchitecturesAllowed=x64
SetupLogging=yes

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "Créer un raccourci sur le Bureau"; GroupDescription: "Raccourcis:"; Flags: unchecked
Name: "installprereqs"; Description: "Installer les prérequis (Thunderbird, Tesseract OCR + français)"; GroupDescription: "Prérequis :"; Flags: checkedonce

[Files]
; ⚠️ Mets ici le chemin vers l'exe généré par PyInstaller
Source: "dist\pipeline_evalnat.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\EvalNat-Publipostage.exe"; DestDir: "{app}"; DestName: "PublipostageEVALNAT.exe"; Flags: ignoreversion

[Icons]
Name: "{group}\PublipostageEVALNAT"; Filename: "{app}\PublipostageEVALNAT.exe"
Name: "{autodesktop}\PublipostageEVALNAT"; Filename: "{app}\PublipostageEVALNAT.exe"; Tasks: desktopicon

[Run]
; Lance le script de prérequis si l'utilisateur a coché l'option
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{tmp}\install_prereqs.ps1"""; \
  Flags: runhidden waituntilterminated; \
  Tasks: installprereqs

; Lancer l'application en fin d'installation (optionnel)
Filename: "{app}\PublipostageEVALNAT.exe"; Description: "Lancer PublipostageEVALNAT maintenant"; Flags: nowait postinstall skipifsilent

[Code]

procedure CurStepChanged(CurStep: TSetupStep);
var
  ScriptPath: string;
  PS: string;
begin
  if CurStep = ssInstall then
  begin
    ScriptPath := ExpandConstant('{tmp}\install_prereqs.ps1');

    { Contenu complet du script PowerShell }
    PS :=
      '$ErrorActionPreference = ''Stop'';'#13#10 +
      'Function Test-Cmd($name){ try { Get-Command $name -ErrorAction Stop | Out-Null; $true } catch { $false } }'#13#10 +
      ''#13#10 +
      'Write-Host "=== Vérification des prérequis ==="'#13#10 +
      'if (-not (Test-Cmd ''winget'')) {'#13#10 +
      '  Write-Host "[WARN] winget introuvable. Installez ""App Installer"" depuis Microsoft Store, puis relancez ce programme." -ForegroundColor Yellow'#13#10 +
      '  exit 0'#13#10 +
      '}'#13#10 +
      ''#13#10 +
      'Function Install-IfMissing { param($CmdName, $WingetId, $Display)'#13#10 +
      '  if (Test-Cmd $CmdName) {'#13#10 +
      '    Write-Host "[OK] $Display est déjà présent."'#13#10 +
      '    return'#13#10 +
      '  }'#13#10 +
      '  Write-Host "[INFO] Installation de $Display via winget..."'#13#10 +
      '  & winget install --id $WingetId -e --source winget --accept-source-agreements --accept-package-agreements'#13#10 +
      '}'#13#10 +
      ''#13#10 +
      'Install-IfMissing -CmdName "thunderbird" -WingetId "Mozilla.Thunderbird" -Display "Mozilla Thunderbird"'#13#10 +
      'Install-IfMissing -CmdName "tesseract"   -WingetId "UB-Mannheim.TesseractOCR" -Display "Tesseract OCR"'#13#10 +
      ''#13#10 +
      '$tessExe = "$env:ProgramFiles\Tesseract-OCR\tesseract.exe"'#13#10 +
      'if (Test-Path $tessExe) {'#13#10 +
      '  $tessdata = Split-Path $tessExe -Parent | Join-Path -ChildPath "tessdata"'#13#10 +
      '  if (-not (Test-Path (Join-Path $tessdata "fra.traineddata"))) {'#13#10 +
      '    Write-Host "[INFO] Téléchargement du pack langue français (fra)..."'#13#10 +
      '    $tmp = Join-Path $env:TEMP "fra.traineddata"'#13#10 +
      '    Invoke-WebRequest -UseBasicParsing -OutFile $tmp "https://github.com/tesseract-ocr/tessdata_fast/raw/main/fra.traineddata"'#13#10 +
      '    Copy-Item $tmp (Join-Path $tessdata "fra.traineddata") -Force'#13#10 +
      '    Remove-Item $tmp -Force'#13#10 +
      '    Write-Host "[OK] Langue fra installée."'#13#10 +
      '  } else {'#13#10 +
      '    Write-Host "[OK] Langue fra déjà présente."'#13#10 +
      '  }'#13#10 +
      '} else {'#13#10 +
      '  Write-Host "[WARN] Tesseract-OCR non détecté après installation. Redémarrez la session si besoin." -ForegroundColor Yellow'#13#10 +
      '}'#13#10 +
      ''#13#10 +
      'Write-Host "=== Prérequis vérifiés ==="'#13#10 ;

    SaveStringToFile(ScriptPath, PS, False);
  end;
end;