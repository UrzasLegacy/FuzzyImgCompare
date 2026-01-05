import win32gui
#import win32con

def is_text_field_focused():
    """Check if the currently focused control is a text field."""
    try:
        foreground_window = win32gui.GetForegroundWindow()
        if not foreground_window:
            return False

        # Get the thread and process ID of the foreground window
        thread_id, process_id = win32gui.GetWindowThreadProcessId(foreground_window)

        # Get the handle to the focused control in the foreground window
        focused_control = win32gui.GetFocus()
        if not focused_control:
            return False

        # Get the class name of the focused control
        class_name = win32gui.GetClassName(focused_control)

        # List of class names that typically represent text input fields
        text_field_classes = ['Edit', 'RichEdit20A', 'RichEdit20W']

        if class_name in text_field_classes:
            return True
        else:
            return False
    except Exception as e:
        return False

