import os
import pandas as pd

SCRIPT_NAME = "Merge Multiple Excels into One"

def process_data(file_paths: list[str], update_progress_callback=None) -> str:
    total = len(file_paths)
    data_frames = []

    for index, path in enumerate(file_paths, 1):
        # 1. Update UI progress
        if update_progress_callback:
            progress = index / total
            file_name = os.path.basename(path)
            update_progress_callback(progress, f"Reading ({index}/{total}): {file_name}")

        # 2. Excel processing logic
        df = pd.read_excel(path)
        data_frames.append(df)

    # 3. Combine and export
    if update_progress_callback:
        update_progress_callback(1.0, "Exporting final output file...")

    output_path = "Merged_Result.xlsx"
    combined_df = pd.concat(data_frames, ignore_index=True)
    combined_df.to_excel(output_path, index=False)

    # 4. Return output path to GUI
    return output_path