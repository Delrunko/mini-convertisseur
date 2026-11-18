# Mini Convertisseur

Application de bureau Windows pour convertir des fichiers PDF, Word, texte et images.

## Télécharger l'exécutable Windows

À chaque push sur `main`, GitHub Actions construit `MiniConvertisseur.exe`.

1. Ouvrez l'onglet **Actions** du dépôt GitHub.
2. Sélectionnez le workflow **Build Windows executable** et son exécution la plus récente réussie.
3. Téléchargez l'artefact **MiniConvertisseur-Windows**, puis extrayez le fichier `.exe`.

L'artefact est conservé pendant 30 jours. Vous pouvez aussi lancer manuellement une
construction depuis l'onglet **Actions** avec **Run workflow**.

## Construire sur Windows

Installez Python 3.12, puis exécutez dans le dossier du projet :

```powershell
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean --onefile --windowed --name MiniConvertisseur --collect-all pymupdf --collect-all pdf2docx --collect-all tkinterdnd2 mini_convertisseur.py
```

L'exécutable sera créé dans `dist\MiniConvertisseur.exe`.

La conversion de Word vers PDF nécessite Microsoft Word ou LibreOffice installé
sur l'ordinateur.
