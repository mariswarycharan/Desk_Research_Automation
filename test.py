import pandas as pd
from pathlib import Path


def excel_to_raw_text(excel_file_path, output_file_path="output.txt", sheet_name=1, no_of_rows=None):
    """
    Convert structured Excel data into unstructured raw text.

    Parameters:
        excel_file_path (str): Path to Excel file
        output_file_path (str): Output text file path
        sheet_name (str/int): Sheet name or sheet index
        no_of_rows (int): Number of rows to process (None for all rows)
    """

    # Read Excel file
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

    # Limit the number of rows if specified
    if no_of_rows is not None:
        df = df.head(no_of_rows)

    # Remove completely empty rows/columns
    df = df.dropna(how="all").dropna(axis=1, how="all")

    # Get column headers
    columns = df.columns.tolist()

    # Start building text output
    content = []

    # Intro section
    content.append("This file contains structured Excel data converted into raw text format.\n")
    content.append("The following columns are available in the Excel file:\n")

    for col in columns:
        content.append(f"- {col}")

    content.append("\nBelow is the row-wise data in the same order:\n")

    # Iterate row-wise
    for index, row in df.iterrows():
        content.append(f"\n========== Row {index + 1} ==========")

        for col in columns:
            value = row[col]

            # Handle NaN values
            if pd.isna(value):
                value = "N/A"

            content.append(f"{col}: {value}")

    # Convert list into raw text string
    final_text = "\n".join(content)

    # Write to output file
    with open(output_file_path, "w", encoding="utf-8") as file:
        file.write(final_text)

    print(f"Raw text file created successfully: {output_file_path}")

    return final_text


# Example Usage
if __name__ == "__main__":
    excel_path = "House of Trends - IVD(China & EU)  sources -WIP.xlsx"      # Your Excel file
    output_path = "xlsx_to_txt_content.txt"

    no_of_rows = 10

    excel_to_raw_text(excel_path, output_path, sheet_name=1, no_of_rows=no_of_rows)