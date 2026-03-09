import os
import re
from classifier import extract_text, extract_metadata


def build_new_name(filepath: str, category: dict) -> str:
    """
    Costruisce il nuovo nome file basandosi sul template della categoria
    e sui metadati estratti dal contenuto.
    Ritorna il nuovo nome (con estensione) oppure il nome originale se fallisce.
    """
    template = category.get("rename_template")
    if not template:
        return os.path.basename(filepath)

    ext = os.path.splitext(filepath)[1]
    text = extract_text(filepath)
    meta = extract_metadata(text)

    # Tipo documento dal nome categoria
    meta["type"] = category.get("name", "Doc")

    # Riempi il template
    new_name = template
    for key, value in meta.items():
        placeholder = "{" + key + "}"
        if placeholder in new_name:
            if value:
                new_name = new_name.replace(placeholder, value)
            else:
                # Rimuovi il placeholder e il separatore precedente se vuoto
                new_name = re.sub(r'[_\-]?\{' + key + r'\}', '', new_name)

    # Pulizia finale
    new_name = re.sub(r'[_\-]{2,}', '_', new_name)  # doppi separatori
    new_name = new_name.strip('_-')
    new_name = re.sub(r'[\\/:*?"<>|]', '', new_name)

    if not new_name:
        return os.path.basename(filepath)

    return new_name + ext


def resolve_conflict(destination: str) -> str:
    """
    Se il file di destinazione esiste già, aggiunge un suffisso numerico.
    Es: fattura.pdf → fattura_1.pdf → fattura_2.pdf
    """
    if not os.path.exists(destination):
        return destination

    base, ext = os.path.splitext(destination)
    counter = 1
    while True:
        candidate = f"{base}_{counter}{ext}"
        if not os.path.exists(candidate):
            return candidate
        counter += 1
