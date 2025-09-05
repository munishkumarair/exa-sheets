import os
from dotenv import load_dotenv
import logging
from typing import List, Optional

import pandas as pd
import exa_py

load_dotenv()
logger = logging.getLogger(__name__)

EXA_API_KEY: Optional[str] = os.environ.get("EXA_API_KEY")
exa_client: Optional[exa_py.Exa] = exa_py.Exa(EXA_API_KEY)



def get_exa_client() -> exa_py.Exa:
    """Return a singleton Exa client using the EXA_API_KEY env var."""
    global exa_client
    if exa_client is None:
        if not EXA_API_KEY:
            logger.error("EXA_API_KEY environment variable is not set")
            raise RuntimeError("EXA_API_KEY environment variable is not set")
        exa_client = exa_py.Exa(EXA_API_KEY)
        logger.debug("Initialized Exa client")
    return exa_client


def fetch_value(company: str, data_point: str) -> str:
    """Ask Exa for a single data point on a company. Return 'NA' on error."""
    prompt = (
        f"What is the {data_point} of {company}? "
        "Return only the value or 'NA' if unavailable."
    )
    try:
        client = get_exa_client()
        logger.debug("Sending prompt to Exa: %s", prompt)
        response = client.answer(prompt)
        value = (response.answer or "").strip()
        if not value:
            logger.info(
                "Exa returned empty answer for company='%s', data_point='%s' -> NA",
                company,
                data_point,
            )
            return "NA"
        lower_val = value.lower()
        if lower_val in {"na", "n/a", "not available", "not applicable"}:
            logger.info(
                "Exa explicitly responded with '%s' for company='%s', data_point='%s'",
                value,
                company,
                data_point,
            )
            return "NA"
        return value
    except Exception as exc:
        logger.exception(
            "Exa API call failed for company='%s', data_point='%s': %s",
            company,
            data_point,
            exc,
        )
        return "NA"


def populate_dataframe(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Given an input DataFrame where the first column is the company name and the
    subsequent columns are data points, return a populated DataFrame.
    """
    if df_in.empty or df_in.shape[1] < 2:
        return df_in.copy()

    company_col = df_in.columns[0]
    data_points: List[str] = list(df_in.columns[1:])

    df_out = pd.DataFrame(columns=[company_col] + data_points)
    df_out[company_col] = df_in[company_col]

    for idx, company in df_in[company_col].items():
        for dp in data_points:
            df_out.at[idx, dp] = fetch_value(str(company), str(dp))

    return df_out


def fill_rows(df_out: pd.DataFrame, df_in: pd.DataFrame, row_indices: List[int]) -> pd.DataFrame:
    """
    Fill only the specified row indices in df_out using df_in+Exa.
    df_out must have the same columns as df_in (company + data points).
    Returns df_out (same object) for convenience.
    """
    if df_in.empty or df_in.shape[1] < 2:
        return df_out

    company_col = df_in.columns[0]
    data_points: List[str] = list(df_in.columns[1:])

    for idx in row_indices:
        if idx not in df_in.index:
            continue
        company = df_in.at[idx, company_col]
        for dp in data_points:
            df_out.at[idx, dp] = fetch_value(str(company), str(dp))

    return df_out

def populate_sheet(input_path: str, output_path: str) -> pd.DataFrame:
    """Read Excel, populate via Exa, write to output Excel, and return DataFrame."""
    df_in = pd.read_excel(input_path)
    df_out = populate_dataframe(df_in)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    df_out.to_excel(output_path, index=False)
    return df_out


