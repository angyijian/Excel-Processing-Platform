import os
import sys
import threading
import importlib.util
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES

# ----------------------------------------------------------------------
# 1. Compatibility Setup
# ----------------------------------------------------------------------
class Tk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# ----------------------------------------------------------------------
# 2. Universal GUI Application Class
# ----------------------------------------------------------------------
class UniversalExcelApp(Tk):
    def __init__(self):
        super().__init__()

        self.title("Universal Excel Batch Processor")
        self.geometry("620x730")
        self.resizable(False, False)

        self.file_paths = []
        self.output_file_path = None
        self.scripts_dict = {}  # Stores loaded modules: {"Script Name": module}

        # Load dynamic scripts from 'scripts' folder
        self._load_external_scripts()

        # Build UI
        self._build_ui()

    def _load_external_scripts(self):
        """Scans 'scripts' directory and imports modules dynamically"""
        scripts_dir = os.path.join(os.path.dirname(__file__), "scripts")
        if not os.path.exists(scripts_dir):
            os.makedirs(scripts_dir)

        for file in os.listdir(scripts_dir):
            if file.endswith(".py") and not file.startswith("__"):
                module_name = file[:-3]
                file_path = os.path.join(scripts_dir, file)

                # Dynamic Import
                spec = importlib.util.spec_from_file_location(module_name, file_path)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)

                # Check standard interface function
                if hasattr(module, "process_data"):
                    display_name = getattr(module, "SCRIPT_NAME", module_name)
                    self.scripts_dict[display_name] = module

    def _build_ui(self):
        # Header Title
        ctk.CTkLabel(
            self, text="Universal Excel Processing Platform", font=ctk.CTkFont(size=20, weight="bold")
        ).pack(pady=(15, 5))

        # Script Selection Box
        script_frame = ctk.CTkFrame(self, fg_color="transparent")
        script_frame.pack(fill="x", padx=25, pady=5)

        ctk.CTkLabel(script_frame, text="Select Task Script:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0, 10))

        script_options = list(self.scripts_dict.keys()) if self.scripts_dict else ["No scripts found in /scripts"]
        self.script_combo = ctk.CTkOptionMenu(
            script_frame, values=script_options, width=320
        )
        self.script_combo.pack(side="right", fill="x", expand=True)

        # Drag & Drop Zone
        self.drop_frame = ctk.CTkFrame(self, height=110, corner_radius=10, border_width=2, border_color="#3B82F6")
        self.drop_frame.pack(fill="x", padx=25, pady=10)
        self.drop_frame.pack_propagate(False)

        self.drop_label = ctk.CTkLabel(
            self.drop_frame,
            text="📂 Drag & Drop Excel files here\nor click 'Select Files' below",
            font=ctk.CTkFont(size=14)
        )
        self.drop_label.pack(expand=True)

        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self._on_file_drop)

        # File Controls
        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=25, pady=5)

        self.btn_select = ctk.CTkButton(self.btn_frame, text="Select Files", command=self._select_files, width=120)
        self.btn_select.pack(side="left")

        self.btn_clear = ctk.CTkButton(
            self.btn_frame, text="Clear List", command=self._clear_files, 
            fg_color="#EF4444", hover_color="#DC2626", width=100
        )
        self.btn_clear.pack(side="right")

        # File List Display
        self.file_list_label = ctk.CTkLabel(self, text="Selected Files (0):", anchor="w")
        self.file_list_label.pack(fill="x", padx=25, pady=(5, 2))

        self.file_list_box = ctk.CTkTextbox(self, height=140, corner_radius=8)
        self.file_list_box.pack(fill="x", padx=25, pady=5)
        self.file_list_box.configure(state="disabled")

        # Progress and Status
        self.status_label = ctk.CTkLabel(self, text="Status: Ready", anchor="w", text_color="gray")
        self.status_label.pack(fill="x", padx=25, pady=(5, 2))

        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.pack(fill="x", padx=25, pady=5)
        self.progress_bar.set(0)

        # Execute Button
        self.btn_process = ctk.CTkButton(
            self, text="🚀 Execute Script", font=ctk.CTkFont(size=16, weight="bold"),
            height=45, command=self._start_processing_thread
        )
        self.btn_process.pack(fill="x", padx=25, pady=15)

        # Results Buttons
        self.result_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.result_frame.pack(fill="x", padx=25, pady=5)

        self.btn_open_file = ctk.CTkButton(
            self.result_frame, text="📄 Open Result File", command=self._open_result_file,
            state="disabled", fg_color="#10B981", hover_color="#059669"
        )
        self.btn_open_file.pack(side="left", expand=True, padx=(0, 5))

        self.btn_open_folder = ctk.CTkButton(
            self.result_frame, text="📁 Open Output Folder", command=self._open_result_folder,
            state="disabled", fg_color="#6B7280", hover_color="#4B5563"
        )
        self.btn_open_folder.pack(side="right", expand=True, padx=(5, 0))

    # ----------------------------------------------------------------------
    # Event Handlers
    # ----------------------------------------------------------------------
    def _on_file_drop(self, event):
        files = self.tk.splitlist(event.data)
        valid_files = [f for f in files if f.lower().endswith(('.xlsx', '.xls'))]
        if valid_files:
            self._add_files(valid_files)

    def _select_files(self):
        files = ctk.filedialog.askopenfilenames(
            title="Select Excel Files", filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if files:
            self._add_files(files)

    def _add_files(self, new_files):
        for f in new_files:
            if f not in self.file_paths:
                self.file_paths.append(f)
        self._update_file_list_ui()

    def _clear_files(self):
        self.file_paths.clear()
        self._update_file_list_ui()
        self.progress_bar.set(0)
        self.status_label.configure(text="Status: Waiting for files...", text_color="gray")
        self.btn_open_file.configure(state="disabled")
        self.btn_open_folder.configure(state="disabled")

    def _update_file_list_ui(self):
        self.file_list_box.configure(state="normal")
        self.file_list_box.delete("1.0", "end")
        for index, path in enumerate(self.file_paths, 1):
            self.file_list_box.insert("end", f"{index}. {os.path.basename(path)}  ({path})\n")
        self.file_list_box.configure(state="disabled")
        self.file_list_label.configure(text=f"Selected Files ({len(self.file_paths)}):")

    # ----------------------------------------------------------------------
    # Execution Logic
    # ----------------------------------------------------------------------
    def _start_processing_thread(self):
        if not self.file_paths:
            self.status_label.configure(text="Status: Please select Excel files first!", text_color="#EF4444")
            return

        selected_script_name = self.script_combo.get()
        if selected_script_name not in self.scripts_dict:
            self.status_label.configure(text="Status: No valid script selected!", text_color="#EF4444")
            return

        self.btn_process.configure(state="disabled")
        self.btn_select.configure(state="disabled")
        self.btn_clear.configure(state="disabled")

        threading.Thread(target=self._run_script, args=(selected_script_name,), daemon=True).start()

    def _run_script(self, script_name):
        try:
            target_module = self.scripts_dict[script_name]

            # Callback function passed to external script
            def update_progress(progress_val: float, message: str):
                self.progress_bar.set(progress_val)
                self.status_label.configure(text=f"Status: {message}", text_color="#3B82F6")

            # Execute the script's entry function
            self.output_file_path = target_module.process_data(self.file_paths, update_progress)

            self.status_label.configure(text="🎉 Script execution completed!", text_color="#10B981")
            self.btn_open_file.configure(state="normal")
            self.btn_open_folder.configure(state="normal")

        except Exception as e:
            self.status_label.configure(text=f"❌ Script Error: {str(e)}", text_color="#EF4444")

        finally:
            self.btn_process.configure(state="normal")
            self.btn_select.configure(state="normal")
            self.btn_clear.configure(state="normal")

    # ----------------------------------------------------------------------
    # Helper Actions
    # ----------------------------------------------------------------------
    def _open_result_file(self):
        if self.output_file_path and os.path.exists(self.output_file_path):
            abs_path = os.path.abspath(self.output_file_path)
            if sys.platform == "win32":
                os.startfile(abs_path)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.call(["open", abs_path])

    def _open_result_folder(self):
        if self.output_file_path and os.path.exists(self.output_file_path):
            folder = os.path.dirname(os.path.abspath(self.output_file_path))
            if sys.platform == "win32":
                os.startfile(folder)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.call(["open", folder])

if __name__ == "__main__":
    app = UniversalExcelApp()
    app.mainloop()