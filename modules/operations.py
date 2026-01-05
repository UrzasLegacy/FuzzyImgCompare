import threading
from tkinter import messagebox
import modules.status
from modules.settings import settings
from pathlib import Path
import subprocess
import logging
import tempfile
import time

def build_search_string(prefix: str, match_results: list) -> str:
    names = [Path(r.path).name for r in match_results]
    or_expr = "|".join(f'"""{n}"""' for n in names)
    # group OR terms so the prefix applies to the whole group
    return f'path:"{prefix}" {or_expr}' if prefix else f"{or_expr}"


def search_with_everything(query: str):
    everything_path = settings.get("everything_path")
    if not everything_path:
        raise ValueError("Everything path isn't set.")
    if not Path(everything_path).is_file():
        raise FileNotFoundError(f"Everything.exe not found at assigned path: {everything_path}")

    logging.info(f"Searching Everything for: {query}")

    subprocess.Popen(
        [everything_path, "-search", query, "-debug-log"],
        creationflags=getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
    )


def search_with_everything_file_list(match_results):
    everything_path = Path(settings["everything_path"])

    # write full paths (recommended) so Everything can match exactly
    with tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False, encoding="utf-8") as f:
        for r in match_results:
            f.write(str(Path(r.path)) + "\n")
        list_path = f.name

    logging.info(f"Searching Everything using file list: {list_path}")

    # Everything: -search-file-list <filename>
    subprocess.Popen([str(everything_path), "-search-file-list", list_path])
    time.sleep(1)  # give Everything time to read the file
    Path(list_path).unlink(missing_ok=True)


def show_warning(title, message):
    if modules.status.root:
        modules.status.root.after(0, lambda: messagebox.showwarning(title, message, parent=modules.status.root))


def show_error(title, message):
    if modules.status.root:
        modules.status.root.after(0, lambda: messagebox.showerror(title, message, parent=modules.status.root))


def ask_yes_no_cancel(title, message):
    if modules.status.root:
        result_container = []
        event = threading.Event()

        def ask():
            result = messagebox.askyesnocancel(title, message, parent=modules.status.root)
            result_container.append(result)
            event.set()

        modules.status.root.after(0, ask)
        event.wait()
        return result_container[0]
    else:
        return None
