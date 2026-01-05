import threading
import logging
import os
from modules.gui import start_gui
from modules.hotkeys import start_hotkey_listener
from modules.settings import load_settings, save_settings

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_file_path = os.path.join(script_dir, "app.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        filename=log_file_path,
        filemode="a",
    )

    # Load settings at startup
    load_settings()

    # Start the hotkey listener in a separate thread
    hotkey_thread = threading.Thread(target=start_hotkey_listener)
    hotkey_thread.daemon = True

    hotkey_thread.start()

    # Start the GUI

    start_gui()

    # Save settings on exit
    save_settings()
