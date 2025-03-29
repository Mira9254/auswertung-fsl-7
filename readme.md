# Automatische Fragebogen Auswertung

## Projekt aufsetzen (Windows)

### 1. Python und virtuelle Umgebung

Stelle sicher, dass Python auf deinem System installiert ist. Du kannst Python von der [offiziellen Seite herunterladen](https://www.python.org/downloads/).

Es wird empfohlen, eine virtuelle Umgebung (venv) zu erstellen und zu aktivieren:

```bash
python -m venv venv
```

Aktiviere die virtuelle Umgebung:

```bash
.\venv\Scripts\activate
```

### 2. Abhängigkeiten installieren

Installiere die benötigten Abhängigkeiten mit:

```bash
pip install -r requirements.txt
```

### 3. Abhängigkeiten aktualisieren (nur für Entwickler)

Aktualisiere die `requirements.txt` mit:

```bash
pip freeze > requirements.txt
```

## Programm starten

Starte das Programm mit:

```bash
python main.py
```

## Code in eine .exe umwandeln

### 1. PyInstaller installieren

Installiere PyInstaller:

```bash
pip install pyinstaller
```

### 2. .exe erstellen

Erstelle die .exe-Datei mit:

```bash
pyinstaller --onefile --windowed --name "Auswertung FSL-7" main.py
```

Die .exe findest du im `dist`-Ordner.