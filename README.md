# Excel Processing Platform

To ensure seamless integration with the **Universal Excel Batch Processing Platform**, all backend processing scripts placed under the `scripts/` directory must adhere strictly to the following standards.

---

## 1. Directory & File Naming Conventions

* **Directory**: Place all script files in the `./scripts/` directory relative to the main application.
* **File Format**: Files must end with `.py` (e.g., `clean_sales_data.py`).
* **Exclusions**: Files starting with dual underscores `__` (such as `__init__.py`) will be ignored automatically by the dynamic plugin engine.

---

## 2. Standard Interface Requirements

Every script module **MUST** implement the following global variable and function signature:

### 2.1 Script Title Identifier (`SCRIPT_NAME`)

A string variable used by the GUI dropdown menu to display a human-readable title.

```python
SCRIPT_NAME = "01. Merge Monthly Sales Reports"
```

### 2.2 Required Input Count (`REQUIRED_FILE_COUNT`)

Optional metadata used by the GUI to render the correct number of file input slots. If
omitted, the GUI defaults to one input slot.

```python
REQUIRED_FILE_COUNT = 2
```

Use this for scripts with a fixed number of inputs, such as a comparison that needs
one base file and one comparison file.

### 2.3 Multiple File Selection (`ALLOW_MULTIPLE_FILES`)

Optional metadata used when a script accepts one or more files through a single input
slot. Set it to `True` to enable multi-select in the file picker and multi-file drag
and drop in the GUI:

```python
ALLOW_MULTIPLE_FILES = True
```

When enabled:

* The GUI accepts one or more Excel files for that script.
* `process_data` receives every selected file in `file_paths`.
* The execute button is enabled once at least one valid file is selected.
* The selected file list shows all files that will be processed.

Do not use `ALLOW_MULTIPLE_FILES = True` for scripts that require a fixed number of
separate roles. For example, `compare_po.py` uses `REQUIRED_FILE_COUNT = 2` so the
GUI can display separate Base File and Compare File inputs.

### 2.4 Entry Function (`process_data`)

The main execution entry point called by the GUI thread.

```python
def process_data(file_paths: list[str], update_progress_callback=None) -> str:
    """
    Executes the custom Excel processing logic.

    Args:
        file_paths (list[str]): A list of absolute or relative file paths
                                selected/dropped into the GUI.
        update_progress_callback (function, optional): A callback used to report
                                                       progress to the UI.
                                                       Signature:
                                                       callback(progress: float, message: str)

    Returns:
        str: Path to the generated result Excel file.

    Raises:
        Exception: Raise standard or custom Python exceptions when processing fails.
                   The GUI captures and displays the error.
    """
```

## 3. Progress and Status Feedback Mechanism

To keep the GUI progress bar and status text responsive during batch tasks, invoke `update_progress_callback` inside your iteration loops:

- **Progress Value**: A `float` between `0.0` (0%) and `1.0` (100%).
- **Status Message**: A short `str` describing the current action.

### Example:

```python
if update_progress_callback:
    update_progress_callback(0.50, "Cleaning duplicate SKUs in file 2...")
```

## 4. GUI Integration Rules

The application dynamically loads every `.py` file in `./scripts/` that:

* does not start with `__`, and
* defines a callable `process_data` function.

The GUI uses `SCRIPT_NAME` for the task dropdown. It passes the selected input paths
to `process_data(file_paths, update_progress_callback)`, then uses the returned path
to enable the Open Result File and Open Output Folder actions.

Input files selected through the GUI must use one of these extensions:
`.xlsx`, `.xls`, `.xlsm`, or `.xlsb`. Scripts should still validate their own content
and raise a clear exception when a workbook does not meet the script's requirements.

The GUI supports both file-picker selection and drag-and-drop. The selected paths are
provided as a list even when only one file is selected.

## 5. Standard Script Template

Copy and modify this starter template when writing new automation tasks:

```python
import os
import pandas as pd

# 1. UI Display Name
SCRIPT_NAME = "01. Sample - Combine Excels"
REQUIRED_FILE_COUNT = 1
ALLOW_MULTIPLE_FILES = True

def process_data(file_paths: list[str], update_progress_callback=None) -> str:
    """
    Standard template for processing input Excel files.
    """
    total_files = len(file_paths)
    data_frames = []

    for index, path in enumerate(file_paths, 1):
        file_name = os.path.basename(path)

        if update_progress_callback:
            progress = index / total_files
            update_progress_callback(progress, f"Reading file ({index}/{total_files}): {file_name}")

        df = pd.read_excel(path)
        data_frames.append(df)

    if update_progress_callback:
        update_progress_callback(1.0, "Exporting generated result file...")

    combined_df = pd.concat(data_frames, ignore_index=True)
    output_path = "Merged_Result.xlsx"
    combined_df.to_excel(output_path, index=False)
    return output_path
```

## 6. Error Handling Guidelines

- **Do not** use `sys.exit()` or suppress critical exceptions silently using blank `except:` blocks.
- If file validation or processing fails, raise an exception directly:

```python
if "SKU" not in df.columns:
    raise ValueError(f"Missing required 'SKU' column in file: {file_name}")
```