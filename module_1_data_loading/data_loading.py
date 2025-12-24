import pandas as pd
import os

# ---------------------------------------------------
# DATA LOADING PIPELINE
# ---------------------------------------------------

def load_excel_data(file_path):
    """
    Loads all sheets from an Excel file safely
    """

    try:
        # STEP 1: SOURCE & PATH
        if not os.path.exists(file_path):
            raise FileNotFoundError("Excel file not found at given path.")

        # STEP 2: LOAD
        excel_file = pd.ExcelFile(file_path)
        sheets = {}

        for sheet in excel_file.sheet_names:
            sheets[sheet] = pd.read_excel(file_path, sheet_name=sheet)

        # STEP 3: VALIDATE
        if not sheets:
            raise ValueError("No sheets found in the Excel file.")

        print("✅ Data loaded successfully.")
        print("📄 Available sheets:", list(sheets.keys()))

        return sheets

    except Exception as error:
        print("❌ Error during data loading:", error)
        return None


def save_loaded_data(sheets, output_file):
    """
    STEP 5: SAVE
    Saves loaded raw data into a new Excel file
    """
    try:
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            for sheet_name, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print("✅ Raw data saved successfully.")

    except Exception as error:
        print("❌ Error while saving raw data:", error)


# ---------------------------------------------------
# MAIN EXECUTION (THIS PART WAS MISSING)
# ---------------------------------------------------

if __name__ == "__main__":

    # ✅ FILE PATH GOES HERE
    file_path = r"E:\vehicle_routing_system\data\ghmc_waste_data.xlsx"

    # CALL THE LOADING FUNCTION
    loaded_sheets = load_excel_data(file_path)

    # SAVE ONLY IF DATA LOADED SUCCESSFULLY
    if loaded_sheets:
        output_file = r"E:\vehicle_routing_system\data\loaded_data.xlsx"
        save_loaded_data(loaded_sheets, output_file)

