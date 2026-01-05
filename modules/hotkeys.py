import keyboard
import logging
from modules.settings import settings


def start_hotkey_listener():
    """Start the global hotkey listener."""
    logging.info("Starting hotkey listener.")

    # Fuzzy image search hotkey
    def _fuzzy_action():
        try:
            from modules import fuzzy_image_search as _fuzzy

            _fuzzy.run_dialog(top_n=15, threshold=0.53, recursive=True)
        except Exception as e:
            logging.exception("Fuzzy search error.")

    fuzzy_keys = settings["hotkeys"].get("fuzzy_search", ["ctrl+alt+f"])
    if isinstance(fuzzy_keys, str):
        fuzzy_keys = [fuzzy_keys]
    for hk in fuzzy_keys:
        keyboard.add_hotkey(hk, _fuzzy_action)
        logging.info(f"Registered hotkey '{hk}' for fuzzy image search.")

    # Block and wait for events
    keyboard.wait()
