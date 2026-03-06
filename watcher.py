import os
import shutil
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from classifier import classify
from renamer import build_new_name, resolve_conflict


class FileSorterHandler(FileSystemEventHandler):
    def __init__(self, base_folder: str, rules_manager, log_callback=None):
        super().__init__()
        self.base_folder   = base_folder
        self.rules_manager = rules_manager
        self.log_callback  = log_callback or (lambda msg, level: None)
        self._processing   = set()

    def on_created(self, event):
        if event.is_directory:
            return
        filepath = event.src_path
        if filepath in self._processing:
            return
        self._processing.add(filepath)
        time.sleep(1.0)
        self._handle_file(filepath)
        self._processing.discard(filepath)

    def _handle_file(self, filepath: str):
        if not os.path.exists(filepath):
            return
        for cat in self.rules_manager.categories:
            cat_path = os.path.join(self.base_folder, cat["folder"])
            if filepath.startswith(cat_path):
                return
        fallback_path = os.path.join(self.base_folder, self.rules_manager.fallback_folder)
        if filepath.startswith(fallback_path):
            return
        category, confidence = classify(filepath, self.rules_manager)
        self._process(filepath, category, confidence)

    def _process(self, filepath: str, category, confidence: float):
        filename = os.path.basename(filepath)
        dry_run  = self.rules_manager.dry_run

        if category is None:
            dest_folder  = os.path.join(self.base_folder, self.rules_manager.fallback_folder)
            new_filename = filename
            self.log_callback(
                f"[NON CLASSIFICATO] {filename} → {self.rules_manager.fallback_folder} "
                f"(confidenza: {confidence:.0%})", "warning")
        else:
            dest_folder = os.path.join(self.base_folder, category["folder"])
            if self.rules_manager.rename_enabled and category.get("rename_template"):
                new_filename = build_new_name(filepath, category)
            else:
                new_filename = filename
            label = "[DRY RUN] " if dry_run else ""
            self.log_callback(
                f"{label}{filename} → {category['folder']}/{new_filename} "
                f"(confidenza: {confidence:.0%})", "info")

        if dry_run:
            return

        os.makedirs(dest_folder, exist_ok=True)
        destination = resolve_conflict(os.path.join(dest_folder, new_filename))
        try:
            shutil.move(filepath, destination)
        except Exception as e:
            self.log_callback(f"[ERRORE] {filename}: {e}", "error")


class FolderWatcher:
    def __init__(self, base_folder: str, rules_manager, log_callback=None, undo_stack=None):
        self.base_folder   = base_folder
        self.rules_manager = rules_manager
        self.log_callback  = log_callback
        self._observer     = None
        self._handler      = None

    def _make_handler(self):
        return FileSorterHandler(self.base_folder, self.rules_manager, self.log_callback)

    def start(self):
        if self._observer and self._observer.is_alive():
            return
        self._handler  = self._make_handler()
        self._observer = Observer()
        self._observer.schedule(self._handler, self.base_folder, recursive=False)
        self._observer.start()

    def stop(self):
        if self._observer:
            self._observer.stop()
            self._observer.join()
            self._observer = None

    @property
    def is_running(self) -> bool:
        return self._observer is not None and self._observer.is_alive()

    def sort_existing(self):
        if not self._handler:
            self._handler = self._make_handler()
        for fname in os.listdir(self.base_folder):
            fpath = os.path.join(self.base_folder, fname)
            if os.path.isfile(fpath):
                self._handler._handle_file(fpath)
