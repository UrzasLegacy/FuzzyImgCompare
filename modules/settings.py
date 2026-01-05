import json
import os
import logging

# Global settings dictionary

# Default settings
settings = {
    "target_folders": {},
    "hotkeys": {"fuzzy_search": ["ctrl+alt+f"]},
    "stay_on_top": True,
    "file_filter_types": ["*.*"],
    "theme": "winxpblue",
}


CONFIG_FILE = "config.json"
CONFIG_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), CONFIG_FILE)
print(f"Config file located at: {CONFIG_PATH}")


def load_settings():
    """Load settings from the configuration file."""
    global settings
    if os.path.isfile(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                data = json.load(f)
                settings.update(data)
                logging.info("Settings loaded.")
        except json.JSONDecodeError:
            logging.error("Error decoding settings file.")
    else:
        logging.info("No settings file found. Using default settings.")


def save_settings():
    """Save settings to the configuration file."""
    with open(CONFIG_PATH, "w") as f:
        json.dump(settings, f, indent=4)
        logging.info("Settings saved.")
