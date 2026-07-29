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


def resolve_required_file_count(module) -> int:
    """Return the number of files a script expects based on its metadata."""
    count = getattr(module, "REQUIRED_FILE_COUNT", None)
    if count is None:
        return 1

    try:
        parsed_count = int(count)
    except (TypeError, ValueError):
        return 1

    return max(1, parsed_count)


# ----------------------------------------------------------------------
# 2. Universal GUI Application Class
# ----------------------------------------------------------------------
class UniversalExcelApp(Tk):
    def __init__(self):
        super().__init__()

        self.title("Universal Excel Batch Processor")
        self.geometry("620x900")
        self.resizable(True, True)

        self.file_paths = []
        self.output_file_path = None
        self.scripts_dict = {}  # Stores loaded modules: {"Script Name": module}
        self.input_paths = []
        self.required_file_count = 1
        self.input_slot_path_vars = []
        self.input_slot_buttons = []
        self.input_slots_frame = None

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
            script_frame, values=script_options, width=320, command=self._on_script_changed
        )
        self.script_combo.pack(side="right", fill="x", expand=True)

        # File Input Section
        self.input_section_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.input_section_frame.pack(fill="x", padx=25, pady=(8, 6))

        ctk.CTkLabel(self.input_section_frame, text="Required File Inputs:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(0, 6))

        self.input_slots_frame = ctk.CTkFrame(self.input_section_frame, fg_color="transparent")
        self.input_slots_frame.pack(fill="x")

        # File Controls
        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=25, pady=5)

        self.btn_clear = ctk.CTkButton(
            self.btn_frame, text="Clear All", command=self._clear_files,
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

        self._refresh_input_slots()

        if script_options and script_options[0] != "No scripts found in /scripts":
            self.script_combo.set(script_options[0])
            self._on_script_changed(script_options[0])

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
    # Dynamic Input Slot UI
    # ----------------------------------------------------------------------
    def _get_slot_label(self, slot_index: int) -> str:
        if self.required_file_count == 2:
            if slot_index == 0:
                return "Base File"
            if slot_index == 1:
                return "Compare File"

        return f"Input {slot_index + 1}"

    def _sync_input_paths_for_count(self):
        if not self.input_paths:
            self.input_paths = [None] * self.required_file_count
        elif len(self.input_paths) < self.required_file_count:
            self.input_paths.extend([None] * (self.required_file_count - len(self.input_paths)))
        elif len(self.input_paths) > self.required_file_count:
            self.input_paths = self.input_paths[:self.required_file_count]

    def _refresh_input_slots(self):
        if self.input_slots_frame is None:
            return

        for child in self.input_slots_frame.winfo_children():
            child.destroy()

        self.input_slot_path_vars = []
        self.input_slot_buttons = []
        self._sync_input_paths_for_count()

        for slot_index in range(self.required_file_count):
            slot_title = self._get_slot_label(slot_index)
            slot_frame = ctk.CTkFrame(self.input_slots_frame, corner_radius=10, border_width=1, border_color="#4B5563")
            slot_frame.pack(fill="x", pady=4)

            ctk.CTkLabel(slot_frame, text=slot_title, font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(8, 2))

            drop_frame = ctk.CTkFrame(slot_frame, height=78, corner_radius=8, border_width=1, border_color="#3B82F6")
            drop_frame.pack(fill="x", padx=10, pady=(0, 6))
            drop_frame.pack_propagate(False)

            ctk.CTkLabel(
                drop_frame,
                text="📂 Drag Excel file here",
                font=ctk.CTkFont(size=12)
            ).pack(expand=True)

            drop_frame.drop_target_register(DND_FILES)
            drop_frame.dnd_bind('<<Drop>>', lambda event, idx=slot_index: self._on_slot_drop(event, idx))

            button_row = ctk.CTkFrame(slot_frame, fg_color="transparent")
            button_row.pack(fill="x", padx=10, pady=(0, 8))

            select_button = ctk.CTkButton(
                button_row,
                text=f"Select {slot_title}",
                command=lambda idx=slot_index: self._select_file_for_slot(idx),
                width=140
            )
            select_button.pack(side="left")
            self.input_slot_buttons.append(select_button)

            path_var = ctk.StringVar(value=self.input_paths[slot_index] or "")
            path_entry = ctk.CTkEntry(button_row, textvariable=path_var, state="disabled")
            path_entry.pack(side="right", fill="x", expand=True, padx=(10, 0))
            self.input_slot_path_vars.append(path_var)

        self._refresh_selected_files()

    def _refresh_selected_files(self):
        self.file_paths = [path for path in self.input_paths if path]
        if hasattr(self, "file_list_box") and hasattr(self, "file_list_label"):
            self._update_file_list_ui()
        self._update_execute_button_state()

    def _update_execute_button_state(self):
        if not hasattr(self, "btn_process"):
            return

        is_ready = len(self.file_paths) == self.required_file_count and self.required_file_count > 0
        if is_ready:
            self.btn_process.configure(state="normal")
        else:
            self.btn_process.configure(state="disabled")

    def _update_slot_display(self, slot_index: int):
        if 0 <= slot_index < len(self.input_slot_path_vars):
            self.input_slot_path_vars[slot_index].set(self.input_paths[slot_index] or "")
        self._refresh_selected_files()

    def _is_supported_excel_file(self, file_path: str) -> bool:
        return os.path.splitext(file_path)[1].lower() in {".xlsx", ".xls", ".xlsm", ".xlsb"}

    def _on_script_changed(self, _selected_script_name):
        selected_module = self.scripts_dict.get(_selected_script_name)
        self.required_file_count = resolve_required_file_count(selected_module)
        self._sync_input_paths_for_count()
        self._refresh_input_slots()
        self.status_label.configure(text=f"Status: Ready for {self.required_file_count} input file(s).", text_color="gray")

    def _on_slot_drop(self, event, slot_index: int):
        files = self.tk.splitlist(event.data)
        valid_files = [f for f in files if self._is_supported_excel_file(f)]
        if valid_files:
            self._set_input_file(slot_index, valid_files[0])

    def _select_file_for_slot(self, slot_index: int):
        files = ctk.filedialog.askopenfilenames(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm *.xlsb")]
        )
        if files:
            self._set_input_file(slot_index, files[0])

    def _set_input_file(self, slot_index: int, file_path: str):
        if slot_index >= len(self.input_paths):
            self._sync_input_paths_for_count()
        self.input_paths[slot_index] = file_path
        self._update_slot_display(slot_index)

    # ----------------------------------------------------------------------
    # Event Handlers
    # ----------------------------------------------------------------------
    def _clear_files(self):
        self.input_paths = [None] * self.required_file_count
        self._refresh_selected_files()
        self.progress_bar.set(0)
        self.status_label.configure(text="Status: Waiting for files...", text_color="gray")
        self.btn_open_file.configure(state="disabled")
        self.btn_open_folder.configure(state="disabled")
        for path_var in self.input_slot_path_vars:
            path_var.set("")

    def _update_file_list_ui(self):
        if not hasattr(self, "file_list_box") or not hasattr(self, "file_list_label"):
            return

        self.file_list_box.configure(state="normal")
        self.file_list_box.delete("1.0", "end")
        for index, path in enumerate(self.file_paths, 1):
            slot_label = self._get_slot_label(index - 1)
            self.file_list_box.insert("end", f"{index}. {slot_label}: {os.path.basename(path)}\n")
        self.file_list_box.configure(state="disabled")
        self.file_list_label.configure(text=f"Selected Files ({len(self.file_paths)}):")

    # ----------------------------------------------------------------------
    # Execution Logic
    # ----------------------------------------------------------------------
    def _start_processing_thread(self):
        selected_script_name = self.script_combo.get()
        if selected_script_name not in self.scripts_dict:
            self.status_label.configure(text="Status: No valid script selected!", text_color="#EF4444")
            return

        self._refresh_selected_files()
        if len(self.file_paths) != self.required_file_count:
            self.status_label.configure(
                text=f"Status: This script requires {self.required_file_count} file(s).",
                text_color="#EF4444"
            )
            return

        self.btn_process.configure(state="disabled")
        self.btn_clear.configure(state="disabled")
        for button in self.input_slot_buttons:
            button.configure(state="disabled")

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
            self._update_execute_button_state()
            self.btn_clear.configure(state="normal")
            for button in self.input_slot_buttons:
                button.configure(state="normal")

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