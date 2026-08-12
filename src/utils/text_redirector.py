"""
System Stream Redirector - Real-time UI logger
----------------------------------------------
Description: Thread-safe text redirector that intercepts standard outputs
             (stdout/stderr) and pipes print feeds directly into a CustomTkinter textbox.

Author: Joseph Adogeri
Version: 1.0.0
Since: 2026-08-12
File: text_redirector.py
License: MIT
"""

class TextRedirector:
    """Safely intercepts standard system print calls and routes them into the CTkTextbox."""

    def __init__(self, textbox_widget, app_context):
        self.textbox = textbox_widget
        self.app = app_context

    def write(self, string):
        if string:  # Ignore completely empty strings
            # Safely schedule the text insertion onto the main GUI thread loop
            self.app.after(0, lambda: self._safe_append(string))

    def _safe_append(self, string):
        try:
            self.textbox.configure(state="normal")
            self.textbox.insert("end", string)
            self.textbox.see("end")  # Dynamic auto-scroll to the bottom of the box
            self.textbox.configure(state="disabled")

            # ⚡ CRITICAL FIX: Force the UI to refresh its text layout instantly
            self.textbox.update_idletasks()
        except Exception:
            pass

    def flush(self):
        pass  # Required for Python standard stream buffer compatibility

