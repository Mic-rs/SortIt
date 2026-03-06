# SortIt — Ordinamento Intelligente

SortIt è un'applicazione desktop per Windows che monitora una cartella e sposta automaticamente i file nelle sottocartelle giuste, in base a regole configurabili per estensione e parole chiave nel contenuto.

---

## Requisiti

- Python 3.10+
- Dipendenze (installa con `pip install -r requirements.txt`):
  - `customtkinter`, `watchdog`, `pdfplumber`, `python-docx`, `Pillow`, `winotify`
  - `pyinstaller` (solo per compilare l'exe)

---

## Avvio

```bash
python main.py
```

Oppure avvia direttamente `SortIt.exe`.

---

## Come funziona

### 1. Seleziona la cartella da monitorare

Clicca **Sfoglia** e scegli la cartella da monitorare (es. `C:\Scuola`). La scelta viene salvata automaticamente e ricaricata al prossimo avvio.

### 2. Configura le regole

Clicca **📋** per aprire la finestra **Gestione Regole**. Ogni regola definisce:

| Campo | Descrizione |
|---|---|
| **Nome categoria** | Nome identificativo della regola |
| **Cartella destinazione** | Sottocartella dove verranno spostati i file (creata automaticamente) |
| **Estensioni** | Estensioni separate da virgola (es. `.pdf, .docx`) |
| **Parole chiave** | Parole da cercare nel contenuto del file (es. `algebra, equazione`) |
| **Template rinomina** | Formato del nuovo nome file (opzionale) |

Un file viene classificato se la sua **estensione** corrisponde, oppure se il suo **contenuto** contiene almeno la percentuale minima di parole chiave configurata nelle Impostazioni.

### 3. Avvia il monitoraggio

Clicca **▶ Avvia** — ogni file che entra nella cartella viene analizzato e spostato automaticamente. Clicca **⏹ Ferma** per interrompere.

### 4. Ordina i file già presenti

Clicca **⚡ Ordina Esistenti** per classificare i file già presenti nella cartella senza aspettare nuovi arrivi.

---

## Opzioni

### 🔍 Dry Run
Simula lo spostamento senza spostare nulla. Nel log appare `[DRY RUN]` con la destinazione prevista. Utile per verificare le regole prima di usarle.

### ✎ Rinomina automatica
Se attivo, i file vengono rinominati secondo il template definito nella regola. Se il template è vuoto, il nome originale viene mantenuto.

---

## Impostazioni

Clicca **⚙️** per aprire la finestra Impostazioni.

### 🌙 Tema scuro
Alterna tra tema scuro e chiaro. La scelta viene salvata.

### 🚀 Avvio con Windows
Aggiunge SortIt al registro di avvio automatico di Windows.

### ▶ Avvia monitoraggio all'apertura
Avvia il monitoraggio automaticamente all'apertura dell'app, se è già configurata una cartella.

### 🎯 Soglia di confidenza
Percentuale minima di parole chiave che devono essere presenti nel contenuto di un file per classificarlo.

- **50%** (default): con 2 parole chiave basta trovarne 1; con 4 ne bastano 2
- **100%**: tutte le parole chiave devono essere presenti
- **0%**: qualsiasi file con estensione corrispondente viene classificato

---

## Template di rinomina

| Variabile | Significato |
|---|---|
| `{date}` | Data corrente (`YYYY-MM-DD`) |
| `{type}` | Nome della categoria |
| `{sender}` | Mittente estratto dal contenuto |
| `{number}` | Numero documento estratto dal contenuto |

Esempio: `{date}_{type}_{number}` → `2026-03-06_Fatture_0042.pdf`

---

## Sicurezza

- **Validazione path**: la cartella di destinazione deve essere sempre dentro la cartella monitorata. Regole con path anomali vengono bloccate automaticamente.
- **Integrità regole**: SortIt calcola un hash SHA256 delle regole ad ogni salvataggio. Se vengono modificate esternamente dal registro, viene mostrato un avviso all'avvio successivo.

---

## Dati salvati

Tutto viene salvato nel **registro di Windows** — nessun file generato sul disco.

| Chiave registro | Contenuto |
|---|---|
| `Software\SortIt\rules` | Regole di classificazione (JSON) |
| `Software\SortIt\theme` | Tema chiaro/scuro |
| `Software\SortIt\last_folder` | Ultima cartella monitorata |
| `Software\SortIt\autostart` | Avvio con Windows |
| `Software\SortIt\automonitor` | Avvio monitoraggio automatico |

Ogni utente ha il proprio registro — le impostazioni non vengono condivise quando si distribuisce l'exe.

---

## Distribuzione

Per condividere l'app basta mandare **solo `SortIt.exe`**. Al primo avvio crea automaticamente le regole di default nel registro del PC destinatario.

---

## Compilare l'exe

```bash
pyinstaller --onefile --windowed --name "SortIt" --icon "icon.ico" ^
  --collect-all customtkinter ^
  --add-data "icon.ico;." --add-data "icon.png;." main.py
```

L'exe viene generato in `dist\SortIt.exe`.

> **Nota**: Windows potrebbe bloccare l'exe alla prima esecuzione. Vai su Proprietà → spunta **Sblocca**, oppure disattiva il Controllo Intelligente delle App.

---

## Struttura del progetto

```
SortIt/
├── main.py            # Interfaccia grafica
├── watcher.py         # Monitoraggio cartella
├── classifier.py      # Logica di classificazione
├── renamer.py         # Logica di rinomina
├── rules_manager.py   # Gestione regole (registro Windows)
├── icon.ico           # Icona applicazione
├── icon.png           # Icona per barra del titolo
└── requirements.txt   # Dipendenze Python
```
