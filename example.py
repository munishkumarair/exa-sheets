import os
import pandas as pd
import exa_py
from typing import Any, Dict

# Initialize the Exa client (replace with your actual key or config)
exa = exa_py.Exa("ae4398e2-f2e3-4041-a352-cc8dbe72d967")

def fetch_value(company: str, data_point: str) -> str:
    """
    Ask Exa for a single data point on a company.
    Returns the plain-text answer or "NA" if not found / on error.
    """
    prompt = f"What is the {data_point} of {company} in 2024? Return only the value or 'NA' if unavailable."
    try:
        response = exa.answer(prompt)
        print(response.answer.strip()
        return response.answer.strip() or "NA"
    except Exception:
        return "NA"

def populate_sheet(
    input_path: str,
    output_path: str,
    sheet_name: str = None
) -> pd.DataFrame:
    """
    Reads the input Excel, populates all cells via Exa calls,
    and writes the result to a new Excel file.
    """
    # Load input
    df_in = pd.read_excel(input_path)
    # Assume first column is company names
    company_col = df_in.columns[0]
    data_points = list(df_in.columns[1:])

    # Prepare output DataFrame
    df_out = pd.DataFrame(columns=[company_col] + data_points)
    df_out[company_col] = df_in[company_col]

    # Iterate companies
    for idx, company in df_in[company_col].items():
        for dp in data_points:
            df_out.at[idx, dp] = fetch_value(company, dp)

    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    df_out.to_excel(output_path, index=False)
    print(f"Output written to: {output_path}")
    return df_out

if __name__ == "__main__":
    INPUT_FILE  = r"C:\Users\m.kumar\Downloads\Exa-test-1.xlsx"
    OUTPUT_FILE = r"C:\Users\m.kumar\Downloads\Exa-test-1-output.xlsx"
    populate_sheet(INPUT_FILE, OUTPUT_FILE)
