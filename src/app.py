import os
import sys
import json
import datetime
import traceback
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Determine absolute path to the directory containing this script file
BASE_APP_DIR = os.path.dirname(os.path.abspath(sys.argv[0] if sys.argv[0] else __file__))

# Adjust system path so it can find adjacent files inside src/
sys.path.append(BASE_APP_DIR)

# Import legacy processing systems
from utils.converter import convert_txt_to_excel
from pipeline import run_pipeline
from constants.job_config import PropertyType

# Import decoupled UI layer
from utils.ui_components import AppUILayer
from utils.text_redirector import TextRedirector


class ChangeOrderApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Enforce absolute file paths relative to the executable directory
        self.config_file = os.path.join(BASE_APP_DIR, "app_config.json")
        self.log_filename = os.path.join(BASE_APP_DIR, "pipeline_errors.log")

        # Apply saved user preferences layout profile before building elements
        initial_theme = self.load_saved_theme()
        ctk.set_appearance_mode(initial_theme)
        ctk.set_default_color_theme("blue")

        self.title("LTC Change Orders Processing Studio")
        self.geometry("720x575")  # Expanded height slightly to gracefully cushion sticky footer component
        self.resizable(False, False)

        # Internal Data Properties
        self.converted_excel_path = ""
        self.current_step = 1
        self.anim_frame_counter = 0
        self.frames = {}

        # Global Dark/Light Theme Selector Top Bar Frame Setup
        AppUILayer.setup_theme_bar(self)

        # Create a clean master container card frame
        self.card = ctk.CTkFrame(self)
        self.card.pack(padx=30, pady=(15, 15), fill="both", expand=True)

        # Initialize modular layouts securely into tracking dictionary
        AppUILayer.init_step_1_ui(self)
        AppUILayer.init_step_2_ui(self)
        AppUILayer.init_loading_ui(self)
        AppUILayer.init_success_ui(self)

        # Render sticky system corporate footer
        AppUILayer.setup_footer(self)

        # ⚙️ REDIRECT ALL STANDARD PRINTS LIVE INTO THE BOX
        sys.stdout = TextRedirector(self.console_box, self)
        sys.stderr = TextRedirector(self.console_box, self)

        # Render first view
        self.show_step(1)

    def load_saved_theme(self) -> str:
        """Reads configuration properties from a JSON file."""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r") as f:
                    config = json.load(f)
                    return config.get("appearance_mode", "System")
            except Exception:
                pass
        return "System"

    def change_theme(self, choice: str):
        """Sets application mode and writes current value to persistent storage."""
        ctk.set_appearance_mode(choice)
        try:
            with open(self.config_file, "w") as f:
                json.dump({"appearance_mode": choice}, f)
        except Exception as e:
            print(f"Failed to save settings: {e}")

    def write_error_log(self, error_message: str, error_traceback: str):
        """Utility function that writes pipeline crashes safely using absolute pathing."""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = (
            f"========================================================================\n"
            f"TIMESTAMP: {timestamp}\n"
            f"ERROR SUMMARY: {error_message}\n"
            f"DETAILED TRACEBACK:\n{error_traceback}"
            f"========================================================================\n\n"
        )
        try:
            # Enforce explicit UTF-8 and write mode
            with open(self.log_filename, "a", encoding="utf-8") as log_file:
                log_file.write(log_entry)
        except Exception as log_err:
            print(f"Failed writing log file artifact: {log_err}")

    def clear_pipeline_log_file(self):
        """Wipes the local pipeline_errors.log file from disk completely."""
        if not os.path.exists(self.log_filename):
            messagebox.showinfo("Log Cleaner", "No error log file found. Your workspace is already clean!")
            return

        try:
            os.remove(self.log_filename)
            messagebox.showinfo("Log Cleaner", "The pipeline_errors.log file has been wiped successfully!")
        except Exception as err:
            messagebox.showerror("Log Cleaner", f"Failed to delete the log file: {err}")

    def show_step(self, step_key):
        """Hides active modules and renders requested frame steps smoothly."""
        self.current_step = step_key
        for k, frame in self.frames.items():
            if k == step_key:
                frame.pack(fill="both", expand=True, padx=20, pady=20)
            else:
                frame.pack_forget()

    def run_loading_animation(self):
        """Simulates visual fluid processing feedback infinitely while tasks finish."""
        if self.current_step != "loading":
            return

        spinners = ["📁 🔍 📄", "📄 📁 🔍", "🔍 📄 📁"]
        dots = "." * (self.anim_frame_counter % 4)
        current_spinner = spinners[self.anim_frame_counter % len(spinners)]

        self.anim_label.configure(text=f" {current_spinner}   Processing{dots} ")
        self.anim_frame_counter += 1

        self.after(800, self.run_loading_animation)

    def browse_txt_file(self):
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if path:
            self.txt_entry.delete(0, "end")
            self.txt_entry.insert(0, path)

    def browse_excel_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel spreadsheets", "*.xlsx *.xls")])
        if path:
            self.converted_excel_path = path
            self.excel_entry.delete(0, "end")
            self.excel_entry.insert(0, path)

    def execute_txt_conversion(self):
        txt_file = self.txt_entry.get().strip()
        if not txt_file or not os.path.exists(txt_file):
            messagebox.showerror("Error", "Please select a valid source text file first.")
            return

        try:
            self.convert_btn.configure(state="disabled", text="Converting...")
            self.update()

            excel_out = convert_txt_to_excel(txt_file)
            self.converted_excel_path = excel_out

            self.excel_entry.delete(0, "end")
            self.excel_entry.insert(0, excel_out)

            messagebox.showinfo(
                "Success",
                f"Data sheet processed successfully!\nSaved as: {os.path.basename(excel_out)}"
            )
            self.show_step(2)
        except Exception as e:
            err_trace = traceback.format_exc()
            self.write_error_log(str(e), err_trace)
            messagebox.showerror(
                "Fault",
                f"Ingestion transformation crashed. Details saved to {os.path.basename(self.log_filename)}.\nError: {e}"
            )
        finally:
            self.convert_btn.configure(state="normal", text="⚙️ Convert to Excel")

    def skip_to_excel_selection(self):
        self.show_step(2)

    def launch_engine(self, property_type: PropertyType):
        """Validates inputs, transitions UI states, and executes the core pipeline on a separate thread."""
        target_excel = self.excel_entry.get().strip()
        if not target_excel or not os.path.exists(target_excel):
            messagebox.showerror(
                "Selection Error",
                "Please provide or browse a valid Excel file to map change records."
            )
            return

        # 🧹 CLEAR OLD LOG CONSOLE CONTENT BEFORE RUNNING NEW BATCH
        self.console_box.configure(state="normal")
        self.console_box.delete("1.0", "end")
        self.console_box.configure(state="disabled")

        # Unify Enum matching to ensure we use the local definition namespace app.py understands
        resolved_type = property_type
        type_str = str(property_type).upper()
        if "PERSONAL" in type_str or "PP" in type_str:
            resolved_type = PropertyType.PERSONAL
        elif "REAL" in type_str:
            resolved_type = PropertyType.REAL

        # Safely transition over to the loading screen view
        self.load_title.configure(text=f"Generating {resolved_type.name} Layouts...")
        self.show_step("loading")
        self.anim_frame_counter = 0
        self.run_loading_animation()
        self.update()

        # ✅ FIX: Spawn a real background worker thread so the UI remains perfectly fluid
        import threading
        worker_thread = threading.Thread(
            target=self.process_pipeline_worker,
            args=(target_excel, resolved_type),
            daemon=True  # Automatically kills background processing loop if app window closes
        )
        worker_thread.start()

    def process_pipeline_worker(self, excel_path: str, property_type: PropertyType):
        """Worker function running on a separate thread to handle heavy Excel and Word tasks seamlessly."""
        try:
            # 🏢 CRITICAL WINDOWS COM INITIALIZATION FIX:
            # Background threads must explicitly initialize a COM context area
            # before spawning or driving MS Word instances via win32com.client
            import pythoncom
            pythoncom.CoInitialize()

            base_dir = os.path.dirname(excel_path)

            # Execute central core processing pipeline script
            run_pipeline(
                excel_file_path=excel_path,
                property_type=property_type,
                base_output_dir=base_dir,
                ui_app=self  # Enables real-time UI logging feedback loops
            )

            # ✅ FIX: Safely route UI updates back to the main thread loop using self.after
            self.after(0, lambda: self.success_title.configure(text=f"{property_type.name} Letters Ready!"))
            self.after(0, lambda: self.success_desc.configure(
                text="All files correctly exported right next to your tracking excel spreadsheet folder location."
            ))

            # Transition window view state to final confirmation checkmark screen
            self.after(400, lambda: self.show_step("success"))

        except Exception as e:
            err_trace = traceback.format_exc()
            self.write_error_log(str(e), err_trace)

            # ✅ FIX: Use string key "step2" instead of integer 2 to eliminate the KeyError layout issue
            self.after(0, lambda: self.show_step("step2"))
            self.after(0, lambda: messagebox.showerror(
                "Pipeline Failure",
                f"A fatal mapping disruption occurred. Error traceback dump appended to {os.path.basename(self.log_filename)}.\n\n"
                f"Error: {e}"
            ))

        finally:
            # Safely release COM allocations from thread pool memory
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


if __name__ == "__main__":
    app = ChangeOrderApp()
    app.mainloop()
