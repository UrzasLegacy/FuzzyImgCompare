import os
import logging
from tkinter import ttk, messagebox, filedialog, Menu, BooleanVar, StringVar, SUNKEN
from ttkthemes import ThemedTk
from modules.settings import settings, save_settings
import modules.status

VERSION = "v0.8.0"

"""Available themes (put name in config file):
pixmap_themes = [
   "arc",
   "blue",
   "clearlooks",
   "elegance",
   "kroc",
   "plastik",
   "radiance",
   "winxpblue"
]
"""


class GuiApp:
    def __init__(self):
        """Initialize and run the GUI."""
        self.root = ThemedTk(theme=settings.get("theme", "winxpblue"))
        modules.status.root = self.root
        self.root.title(f"Ali Fuzzy Image Comparer {VERSION}")
        self.root.geometry("600x600")
        self.root.resizable(True, True)

        # Menu bar setup
        self.menubar = Menu(self.root)
        self.root.config(menu=self.menubar)

        # Options menu
        self.options_menu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Options", menu=self.options_menu)

        # Stay on top option
        self.stay_on_top_var = BooleanVar(value=settings.get("stay_on_top", False))
        self.options_menu.add_checkbutton(
            label="Stay on Top",
            variable=self.stay_on_top_var,
            command=self.toggle_stay_on_top,
        )

        # Fuzzy Image Search (clipboard -> focused folder)
        def _run_fuzzy():
            from modules import fuzzy_image_search as _fuzzy
            _fuzzy.run_dialog(top_n=5, threshold=0.55, recursive=True)

        self.options_menu.add_command(label="Fuzzy Image Search…", command=_run_fuzzy)

        # Help menu
        self.help_menu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Help", menu=self.help_menu)
        self.help_menu.add_command(label="Instructions", command=self.show_instructions)
        self.help_menu.add_command(label="About", command=self.show_about)

        # Main frame
        self.mainframe = ttk.Frame(self.root, padding="1")
        self.mainframe.pack(fill="both", expand=True)
        modules.status.mainframe = self.mainframe

        # Status bar
        modules.status.status_var = StringVar()
        self.status_var = modules.status.status_var  # Use instance variable

        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=SUNKEN, anchor="w")
        self.status_bar.pack(side="bottom", fill="x")

        # Load settings into GUI
        self.toggle_stay_on_top()

        self.root.mainloop()

    def toggle_stay_on_top(self):
        """Toggle the 'Stay on Top' window attribute."""
        settings["stay_on_top"] = self.stay_on_top_var.get()
        save_settings()
        self.root.attributes("-topmost", self.stay_on_top_var.get())
        modules.status.update_status_message(f"Stay On Top set to {settings['stay_on_top']}")
        logging.info(f"Stay On Top set to {settings['stay_on_top']}")

    def show_instructions(self):
        """Display the instructions dialog."""
        instructions = "\n" "\n"
        instructions = """
Press Ctrl+Alt+F to search for duplicates of the last image on the clipboard in the most recently focused folder OR the folder you set in config.json if there is one."""
        messagebox.showinfo("Instructions", instructions, parent=self.root)

    def show_about(self):
        """Display the about dialog."""
        about_info = "Ali Fuzzy Image Comparer\n" f"Version {VERSION}\n" "Author: UrzasLegacy"
        messagebox.showinfo("About", about_info, parent=self.root)

    def select_target_folder(self, index):
        """Open a directory selection dialog to set the target folder."""
        folder_path = filedialog.askdirectory(title=f"Select Folder {index}")
        if folder_path and os.path.isdir(folder_path):
            settings["search_folder"] = folder_path
            save_settings()
            logging.info(f"Search folder set to: {folder_path}")
            modules.status.update_status_message("Search folder set.")
        else:
            logging.warning("No folder was selected or invalid folder.")


def start_gui():
    """Function to initialize and start the GUI application."""
    GuiApp()
