root = None  # Initialized in gui.start_gui()
mainframe = None
status_var = None
# Optional container used by fuzzy search UI (reused between runs)
fuzzy_results_frame = None
# progress_var = None


def update_status_message(message):
    """Update the status message label in a thread-safe manner.\n
    This function is separated into a unique module to avoid circular imports"""
    if root and status_var:
        root.after(0, lambda: status_var.set(message))


# def update_progress(value):
#     """Update the progress bar in a thread-safe manner.\n
#     This function is separated into a unique module to avoid circular imports"""
#     if root and progress_var:
#         root.after(0, lambda: progress_var.set(value))
