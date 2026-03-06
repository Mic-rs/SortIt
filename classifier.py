import os
import re
from typing import Optional, Tuple
from rules_manager import RulesManager


def extract_text(filepath: str) -> str:
    """Estrae testo da PDF, DOCX o file di testo."""
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == ".pdf":
            return _extract_pdf(filepath)
        elif ext in (".docx", ".doc"):
            return _extract_docx(filepath)
        elif ext in (".txt", ".csv", ".xml", ".json", ".html", ".htm"):
            with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
    except Exception:
        pass
    return ""


def _extract_pdf(filepath: str) -> str:
    try:
        import pdfplumber
        with pdfplumber.open(filepath) as pdf:
            pages = [p.extract_text() or "" for p in pdf.pages[:5]]
            return "\n".join(pages)
    except Exception:
        return ""


def _extract_docx(filepath: str) -> str:
    try:
        from docx import Document
        doc = Document(filepath)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception:
        return ""


def classify(filepath: str, rules_manager: RulesManager) -> Tuple[Optional[dict], float]:
    """
    Classifica un file in base alle regole configurate.
    Ritorna (categoria, confidence) oppure (None, 0.0) se non classificato.
    """
    ext = os.path.splitext(filepath)[1].lower()
    filename = os.path.basename(filepath).lower()

    # Prima passa: classificazione per estensione senza contenuto
    extension_only_categories = []
    for cat in rules_manager.categories:
        if ext in [e.lower() for e in cat.get("extensions", [])]:
            if not cat.get("keywords"):
                # Categoria puramente per estensione (immagini, archivi, ecc.)
                return cat, 1.0
            extension_only_categories.append(cat)

    if not extension_only_categories:
        return None, 0.0

    # Seconda passa: analisi contenuto + nome file
    text = extract_text(filepath)
    combined = (filename + " " + text).lower()

    best_cat = None
    best_score = 0.0

    for cat in extension_only_categories:
        keywords = [k.lower() for k in cat.get("keywords", [])]
        if not keywords:
            continue

        hits = sum(1 for kw in keywords if kw in combined)
        score = hits / len(keywords)

        if score > best_score:
            best_score = score
            best_cat = cat

    if best_score >= rules_manager.confidence_threshold:
        return best_cat, best_score

    # Nessuna categoria con confidenza sufficiente
    return None, best_score


def extract_metadata(text: str) -> dict:
    """
    Estrae metadati dal testo: data, mittente, numero documento, tipo.
    Ritorna un dizionario con i campi trovati (stringa vuota se non trovato).
    """
    meta = {
        "date": "",
        "sender": "",
        "number": "",
        "type": ""
    }

    # Data: vari formati comuni
    date_patterns = [
        r'\b(\d{4}-\d{2}-\d{2})\b',                      # 2024-03-15
        r'\b(\d{2}/\d{2}/\d{4})\b',                       # 15/03/2024
        r'\b(\d{1,2}\s+\w+\s+\d{4})\b',                  # 15 marzo 2024
        r'\b(\d{2}\.\d{2}\.\d{4})\b',                     # 15.03.2024
    ]
    for pat in date_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            meta["date"] = _normalize_date(m.group(1))
            break

    # Numero documento
    num_patterns = [
        r'(?:fattura|invoice|n[°.]?|numero)[^\d]*(\d{3,}[/\-]?\d*)',
        r'(?:ricevuta|receipt)[^\d]*(\d+)',
        r'n\.\s*(\d+)',
    ]
    for pat in num_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            meta["number"] = m.group(1).replace("/", "-").replace(" ", "")
            break

    # Mittente / ragione sociale (prima riga o riga con P.IVA vicina)
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    for line in lines[:10]:
        if len(line) > 4 and not re.match(r'^\d', line):
            meta["sender"] = _sanitize_name(line[:40])
            break

    return meta


def _normalize_date(raw: str) -> str:
    """Tenta di convertire vari formati data in YYYY-MM-DD."""
    months_it = {
        "gennaio": "01", "febbraio": "02", "marzo": "03", "aprile": "04",
        "maggio": "05", "giugno": "06", "luglio": "07", "agosto": "08",
        "settembre": "09", "ottobre": "10", "novembre": "11", "dicembre": "12"
    }
    raw_lower = raw.lower()
    for name, num in months_it.items():
        raw_lower = raw_lower.replace(name, num)

    # Prova diversi formati
    formats = ["%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%d %m %Y"]
    from datetime import datetime
    for fmt in formats:
        try:
            return datetime.strptime(raw_lower.strip(), fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return raw[:10].replace("/", "-").replace(".", "-")


def _sanitize_name(s: str) -> str:
    """Rimuove caratteri non validi per nomi file."""
    s = re.sub(r'[\\/:*?"<>|]', '', s)
    s = re.sub(r'\s+', '_', s.strip())
    return s[:30]
