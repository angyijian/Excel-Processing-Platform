# Excel Processing Script Specification

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
Optional metadata used by the GUI to render the correct number of file input slots.

```python
REQUIRED_FILE_COUNT = 2
```

If this variable is omitted, the GUI defaults to 1 input slot.

### 2.3 Entry Function (`process_data`)
The main execution entry point called by the GUI thread.

```python
def process_data(file_paths: list[str], update_progress_callback=None) -> str:
    """
    Executes the custom Excel processing logic.

    Args:
        file_paths (list[str]): A list of absolute or relative file paths 
                                selected/dropped into the GUI.
        update_progress_callback (function, optional): A thread-safe callback function 
                                                     to report progress back to UI.
                                                     Signature: callback(progress: float, message: str)

    Returns:
        str: Path to the generated result Excel file.

    Raises:
        Exception: Throw standard or custom Python exceptions when processing fails.
                   The GUI will capture and render errors in the interface.
    """
```

## 3. Progress and Status Feedback Mechanism
To keep the GUI progress bar and status text responsive during batch tasks, invoke `update_progress_callback` inside your iteration loops:

- **Progress Value**: A `float` between `0.0` (0%) and `1.0` (100%).
- **Status Message**: A short `str` describing the current action.

### Example:

```python
if update_progress_callback:
    # Example: 50% completed
    update_progress_callback(0.50, "Cleaning duplicate SKUs in file 2...")
```

## 4. Standard Script Template
Copy and modify this starter template when writing new automation tasks:

```python
import os
import pandas as pd

# 1. UI Display Name
SCRIPT_NAME = "01. Sample - Combine Excels"

def process_data(file_paths: list[str], update_progress_callback=None) -> str:
    """
    Standard template for processing input Excel files.
    """
    total_files = len(file_paths)
    data_frames = []

    for index, path in enumerate(file_paths, 1):
        file_name = os.path.basename(path)
        
        # Report Progress
        if update_progress_callback:
            progress = index / total_files
            update_progress_callback(progress, f"Reading file ({index}/{total_files}): {file_name}")

        # Read Excel data
        df = pd.read_excel(path)
        
        # --- Apply your Excel logic here ---
        # df['Processed_By'] = 'Python Automation'
        # ------------------------------------

        data_frames.append(df)

    # Combine DataFrames
    if update_progress_callback:
        update_progress_callback(1.0, "Exporting generated result file...")

    combined_df = pd.concat(data_frames, ignore_index=True)
    
    # Save output file
    output_path = "Merged_Result.xlsx"
    combined_df.to_excel(output_path, index=False)

    # Return output file location
    return output_path
```

## 5. Error Handling Guidelines

- Do **not** use `sys.exit()` or suppress critical exceptions silently using blank `except:` blocks.
- If file validation or processing fails, raise an exception directly:

```python
if "SKU" not in df.columns:
    raise ValueError(f"Missing required 'SKU' column in file: {file_name}")
```
