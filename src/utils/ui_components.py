import os
import datetime
import customtkinter as ctk

from src.constants.job_config import PropertyType


class AppUILayer:
    """Handles layout initializations and UI setup for the Change Order Studio."""

    @staticmethod
    def setup_theme_bar(app: ctk.CTk):
        app.top_bar = ctk.CTkFrame(app, fg_color="transparent")
        app.top_bar.pack(fill="x", padx=30, pady=(15, 0))

        app.theme_label = ctk.CTkLabel(app.top_bar, text="🎨 Theme Switcher:", font=ctk.CTkFont(size=11))
        app.theme_label.pack(side="left", padx=(0, 5))

        saved_mode = app.load_saved_theme()

        app.theme_menu = ctk.CTkOptionMenu(
            app.top_bar,
            values=["System", "Light", "Dark"],
            command=app.change_theme,
            width=100,
            height=24,
            font=ctk.CTkFont(size=11)
        )
        app.theme_menu.set(saved_mode)
        app.theme_menu.pack(side="left")

    @staticmethod
    def setup_footer(app: ctk.CTk):
        """Builds a dynamic corporate system footer centered perfectly at the bottom of the window."""
        current_year = datetime.datetime.now().year

        app.footer_bar = ctk.CTkFrame(app, fg_color="transparent", height=25)
        app.footer_bar.pack(side="bottom", fill="x", padx=30, pady=(0, 10))

        app.footer_text = ctk.CTkLabel(
            app.footer_bar,
            text=f"© {current_year} Orleans Parish Assessor's Office. All Rights Reserved.",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        # ✅ FIXED: Removed side="left" and added expand=True to center it cleanly
        app.footer_text.pack(expand=True)

    @staticmethod
    def init_step_1_ui(app: ctk.CTk):
        f = ctk.CTkFrame(app.card, fg_color="transparent")
        app.frames[1] = f  # 🧠 SAVED TO KEY 1

        title = ctk.CTkLabel(f, text="Step 1: Text Data Conversion (Optional)",
                             font=ctk.CTkFont(size=18, weight="bold"))
        title.pack(pady=(15, 10))

        desc = ctk.CTkLabel(f,
                            text="If you have a raw source .txt file, convert it to Excel below.\nOtherwise, skip straight to Step 2 if you already have a formatted spreadsheet.",
                            text_color="gray")
        desc.pack(pady=5)

        row = ctk.CTkFrame(f, fg_color="transparent")
        row.pack(fill="x", padx=40, pady=25)

        app.txt_entry = ctk.CTkEntry(row, placeholder_text="Browse a source text file...", width=420)
        app.txt_entry.grid(row=0, column=0, padx=(0, 10))

        browse_btn = ctk.CTkButton(row, text="Browse TXT", width=100, command=app.browse_txt_file)
        browse_btn.grid(row=0, column=1)

        actions = ctk.CTkFrame(f, fg_color="transparent")
        actions.pack(fill="x", padx=40, pady=10)

        app.convert_btn = ctk.CTkButton(actions, text="⚙️ Convert to Excel", height=40,
                                        command=app.execute_txt_conversion)
        app.convert_btn.pack(side="left", expand=True, fill="x", padx=(0, 10))

        skip_btn = ctk.CTkButton(actions, text="Skip to Step 2 ➔", height=40, fg_color="gray", hover_color="#555555",
                                 command=app.skip_to_excel_selection)
        skip_btn.pack(side="right", expand=True, fill="x")

    @staticmethod
    def init_step_2_ui(app: ctk.CTk):
        f = ctk.CTkFrame(app.card, fg_color="transparent")
        app.frames[2] = f  # 🧠 SAVED TO KEY 2

        title = ctk.CTkLabel(f, text="Step 2: Select Property Change Order Pipeline",
                             font=ctk.CTkFont(size=18, weight="bold"))
        title.pack(pady=(15, 10))

        app.step2_desc = ctk.CTkLabel(f,
                                      text="Select the active tracking spreadsheet file, then choose the single engine layout to process.",
                                      text_color="gray")
        app.step2_desc.pack(pady=5)

        row = ctk.CTkFrame(f, fg_color="transparent")
        row.pack(fill="x", padx=40, pady=20)

        app.excel_entry = ctk.CTkEntry(row, placeholder_text="No Excel spreadsheet loaded yet...", width=420)
        app.excel_entry.grid(row=0, column=0, padx=(0, 10))

        browse_btn = ctk.CTkButton(row, text="Browse Excel", width=100, command=app.browse_excel_file)
        browse_btn.grid(row=0, column=1)

        btn_wrapper = ctk.CTkFrame(f, fg_color="transparent")
        btn_wrapper.pack(fill="x", padx=40, pady=15)

        personal_btn = ctk.CTkButton(btn_wrapper, text="💼 Process Personal Property (PP)",
                                     font=ctk.CTkFont(size=13, weight="bold"), height=50,
                                     command=lambda: app.launch_engine(PropertyType.PERSONAL))
        personal_btn.grid(row=0, column=0, padx=(0, 10), sticky="ew")

        real_btn = ctk.CTkButton(btn_wrapper, text="🏠 Process Real Property (REAL)",
                                 font=ctk.CTkFont(size=13, weight="bold"), height=50, fg_color="#2b7a78",
                                 hover_color="#1f5f5d", command=lambda: app.launch_engine(PropertyType.REAL))
        real_btn.grid(row=0, column=1, sticky="ew")

        btn_wrapper.grid_columnconfigure(0, weight=1)
        btn_wrapper.grid_columnconfigure(1, weight=1)

        back_btn = ctk.CTkButton(f, text="📋 Go Back to Step 1", text_color="gray", fg_color="transparent", hover=False,
                                 command=lambda: app.show_step(1))
        back_btn.pack(pady=10)

    @staticmethod
    def init_loading_ui(app: ctk.CTk):
        f = ctk.CTkFrame(app.card, fg_color="transparent")
        app.frames["loading"] = f  # 🧠 SAVED TO KEY loading

        app.load_title = ctk.CTkLabel(f, text="Processing Request...", font=ctk.CTkFont(size=18, weight="bold"))
        app.load_title.pack(pady=(45, 15))

        app.anim_label = ctk.CTkLabel(f, text=" [ ⏳ ] Spawning Word application layers... ",
                                      font=ctk.CTkFont(size=22, family="Courier"))
        app.anim_label.pack(pady=20)

        app.load_detail = ctk.CTkLabel(f,
                                       text="Generating Word mailmerges, drawing Code128 barcodes, and compiling PDF sheets.\nPlease do not close this window...",
                                       text_color="gray")
        app.load_detail.pack(pady=15)


        # 📋 ✅ FIX: Explicitly create and attach the console box to 'app' parameter context
        app.console_box = ctk.CTkTextbox(
            f,
            height=220,
            font=ctk.CTkFont(family="Courier", size=11),
            state="disabled", # Protects logs data from being edited by hand
            wrap="word",
            fg_color=("#f5f5f5", "#1c1c1c"),
            text_color=("#222222", "#b5cea8")
        )
        app.console_box.pack(fill="both", expand=True, padx=25, pady=(10, 5))


    @staticmethod
    def init_success_ui(app: ctk.CTk):
        f = ctk.CTkFrame(app.card, fg_color="transparent")
        app.frames["success"] = f  # 🧠 SAVED TO KEY success

        badge = ctk.CTkLabel(f, text="✨ 🏆 ✨", font=ctk.CTkFont(size=48))
        badge.pack(pady=(30, 10))
        app.success_title = ctk.CTkLabel(f, text="Generation Cycles Complete!",
                                         font=ctk.CTkFont(size=20, weight="bold"), text_color="#2b7a78")
        app.success_title.pack(pady=10)
        app.success_desc = ctk.CTkLabel(f, text="All letter records generated and matched successfully.",
                                        text_color="gray")
        app.success_desc.pack(pady=5)

        btn_row = ctk.CTkFrame(f, fg_color="transparent")
        btn_row.pack(pady=30)

        reset_btn = ctk.CTkButton(btn_row, text="🔂 Process Another Property", font=ctk.CTkFont(weight="bold"),
                                  height=40, command=lambda: app.show_step(2))
        reset_btn.grid(row=0, column=0, padx=10)

        clean_log_btn = ctk.CTkButton(btn_row, text="🗑️ Clear Error Logs", fg_color="#cb4154", hover_color="#9e2a2b",
                                      font=ctk.CTkFont(weight="bold"), height=40, command=app.clear_pipeline_log_file)
        clean_log_btn.grid(row=0, column=1, padx=10)
