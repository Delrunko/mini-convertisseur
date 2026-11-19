# Mini Convertisseur

Application de bureau Windows pour convertir des fichiers PDF, Word, texte et images.

## Installer sur Windows

À chaque push sur `main`, GitHub Actions crée l'installateur `MiniConvertisseur-Setup.exe`.

1. Ouvrez [la dernière exécution du workflow](https://github.com/Delrunko/mini-convertisseur/actions/workflows/build-windows.yml).
2. Cliquez sur l'exécution la plus récente réussie.
3. Dans **Artifacts**, téléchargez **MiniConvertisseur-Installateur**.
4. Extrayez le fichier ZIP et lancez `MiniConvertisseur-Setup.exe`.

L'artefact est conservé pendant 30 jours. Vous pouvez aussi lancer manuellement une
construction depuis l'onglet **Actions** avec **Run workflow**.

## Construire sur Windows

Installez Python 3.12 et Inno Setup 6, puis exécutez dans le dossier du projet :

```powershell
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean --onefile --windowed --name MiniConvertisseur --collect-all pymupdf --collect-all pdf2docx --collect-all tkinterdnd2 mini_convertisseur.py
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
```

L'installateur sera créé dans `installer\MiniConvertisseur-Setup.exe`.

La conversion de Word vers PDF nécessite Microsoft Word ou LibreOffice installé
sur l'ordinateur.
