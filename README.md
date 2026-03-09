# SortIt v2.0 — Ordinamento Intelligente

SortIt è un'applicazione desktop per Windows che monitora una cartella (o una cartella Google Drive) e sposta automaticamente i file nelle sottocartelle giuste, in base a regole configurabili per estensione e parole chiave nel contenuto.

---

## Requisiti

- Python 3.10+
- Dipendenze (installa con `pip install -r requirements.txt`):
  - `customtkinter`, `watchdog`, `pdfplumber`, `python-docx`, `Pillow`, `winotify`
  - `google-auth`, `google-auth-oauthlib`, `google-api-python-client`
  - `pyinstaller` (solo per compilare l'exe)

---

## Avvio

```bash
python main.py
```

Oppure avvia direttamente `SortIt.exe`.

---

## Come funziona

### 1. Scegli la sorgente

Usa lo switch **Locale ↔ Google Drive** per scegliere da dove SortIt prende i file:

- **Locale**: monitora una cartella sul tuo PC. Ogni file che vi entra viene classificato e spostato nelle sottocartelle.
- **Google Drive**: esegue il polling di una cartella Drive. Ogni file trovato viene scaricato, classificato nella cartella locale, e **eliminato dalla cartella Drive**.

> Non puoi cambiare sorgente mentre il monitoraggio è attivo. Ferma prima il monitoraggio.

### 2. Configura Google Drive (solo modalità Drive)

1. Clicca **🔗 Connetti Google** — si apre il browser per il login
2. Dopo il login, seleziona la **cartella Drive** da monitorare con **📂 Scegli**
3. Imposta l'**intervallo di polling** (da 1 a 60 minuti)
4. Seleziona anche la **cartella locale** dove i file verranno salvati e organizzati

Il token di accesso viene salvato nel registro Windows — non sarà necessario rifare il login ai prossimi avvii.

> Per usare questa funzione è necessario un account Google autorizzato.


### 3. Seleziona la cartella locale

Clicca **Sfoglia** e scegli la cartella da monitorare (es. `C:\Scuola`) oppure la cartella di destinazione per i file scaricati da Drive.

### 4. Configura le regole

Clicca **📋** per aprire la finestra **Gestione Regole**. Ogni regola definisce:

| Campo | Descrizione |
|---|---|
| **Nome categoria** | Nome identificativo della regola |
| **Cartella destinazione** | Sottocartella dove verranno spostati i file (creata automaticamente) |
| **Estensioni** | Estensioni separate da virgola (es. `.pdf, .docx`) |
| **Parole chiave** | Parole da cercare nel contenuto del file (es. `algebra, equazione`) |
| **Template rinomina** | Formato del nuovo nome file (opzionale) |

### 5. Avvia il monitoraggio

Clicca **▶ Avvia** — SortIt inizia a monitorare la sorgente selezionata. Clicca **⏹ Ferma** per interrompere.

### 6. Ordina i file già presenti

Clicca **⚡ Ordina Esistenti** per classificare i file già presenti nella cartella locale senza aspettare nuovi arrivi.

---

## Opzioni

### 🔍 Dry Run
Simula lo spostamento senza spostare nulla. Nel log appare `[DRY RUN]` con la destinazione prevista.

### ✎ Rinomina automatica
Se attivo, i file vengono rinominati secondo il template definito nella regola.

---

## Impostazioni

Clicca **⚙️** per aprire la finestra Impostazioni.

### 🌙 Tema scuro
Alterna tra tema scuro e chiaro.

### 🚀 Avvio con Windows
Aggiunge SortIt al registro di avvio automatico di Windows.

### ▶ Avvia monitoraggio all'apertura
Avvia il monitoraggio automaticamente all'apertura dell'app.

### 🎯 Soglia di confidenza
Percentuale minima di parole chiave che devono essere presenti nel contenuto per classificare un file.

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

- **Validazione path**: la cartella di destinazione deve essere sempre dentro la cartella monitorata.
- **Integrità regole**: hash SHA256 delle regole ad ogni salvataggio. Se vengono modificate esternamente, viene mostrato un avviso.

---

## Dati salvati

Tutto viene salvato nel **registro di Windows** — nessun file generato sul disco.

| Chiave registro | Contenuto |
|---|---|
| `Software\SortIt\rules` | Regole di classificazione (JSON) |
| `Software\SortIt\theme` | Tema chiaro/scuro |
| `Software\SortIt\last_folder` | Ultima cartella locale selezionata |
| `Software\SortIt\autostart` | Avvio con Windows |
| `Software\SortIt\automonitor` | Avvio monitoraggio automatico |
| `Software\SortIt\source_mode` | Modalità sorgente (`local` o `drive`) |
| `Software\SortIt\google_token` | Token OAuth Google (cifrato da Windows) |
| `Software\SortIt\google_drive_folder_id` | ID cartella Drive selezionata |
| `Software\SortIt\google_drive_folder_name` | Nome cartella Drive selezionata |
| `Software\SortIt\google_poll_interval` | Intervallo polling in minuti |

---

## Distribuzione

Per condividere l'app basta mandare **solo `SortIt.exe`**. Al primo avvio crea automaticamente le regole di default nel registro del PC destinatario.

---

## Compilare l'exe

```bash
pyinstaller --onefile --windowed --name "SortIt" --icon "icon.ico" ^
  --collect-all customtkinter ^
  --add-data "icon.ico;." --add-data "icon.png;." ^
  --add-data "credentials.json;." main.py
```

L'exe viene generato in `dist\SortIt.exe`.

> **Nota**: Windows potrebbe bloccare l'exe alla prima esecuzione. Vai su Proprietà → spunta **Sblocca**, oppure disattiva il Controllo Intelligente delle App.

---

## Struttura del progetto

```
SortIt/
├── main.py                # Interfaccia grafica
├── watcher.py             # Monitoraggio cartella locale
├── classifier.py          # Logica di classificazione
├── renamer.py             # Logica di rinomina
├── rules_manager.py       # Gestione regole (registro Windows)
├── google_drive_sync.py   # Autenticazione OAuth e polling Google Drive
├── credentials.json       # Credenziali OAuth Google (non committare)
├── icon.ico               # Icona applicazione
├── icon.png               # Icona per barra del titolo
└── requirements.txt       # Dipendenze Python
```

---

## Autore

[@_michelecast_](https://instagram.com/_michelecast_)
