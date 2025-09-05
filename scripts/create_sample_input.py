import argparse
from pathlib import Path
from typing import List

import pandas as pd


DEFAULT_COMPANIES: List[str] = [
    "Apple",
    "Microsoft",
    "Google",
    "Amazon",
    "Meta",
    "Netflix",
    "NVIDIA",
    "Tesla",
    "IBM",
    "Intel",
    "Oracle",
    "Salesforce",
    "Adobe",
    "Cisco",
    "Samsung",
    "Uber",
    "Airbnb",
    "Shopify",
    "Spotify",
    "PayPal",
]

DEFAULT_DATA_POINTS: List[str] = [
    "Revenue",
    "Headcount",
    "CEO",
    "Headquarters",
    "Market Cap",
    "Ticker",
    "Founded Year",
    "Website",
]


def create_sample_input_dataframe(
    companies: List[str], data_points: List[str]
) -> pd.DataFrame:
    """Create a DataFrame with the first column as companies and remaining as data points."""
    if not companies:
        companies = DEFAULT_COMPANIES
    if not data_points:
        data_points = DEFAULT_DATA_POINTS

    columns = ["Company"] + data_points
    df = pd.DataFrame(columns=columns)
    df["Company"] = companies

    # Leave data point cells empty for population later
    return df


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Create a sample Excel input file. First column is 'Company', subsequent columns are data points."
        )
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(r"C:\Users\Munish Kumar\Downloads\exa-sheets\sample_input.xlsx"),
        help="Path to write the Excel file (default: sample_input.xlsx)",
    )
    parser.add_argument(
        "--companies",
        type=str,
        default=",".join(DEFAULT_COMPANIES),
        help=(
            "Comma-separated company names. Default: "
            + ", ".join(DEFAULT_COMPANIES)
        ),
    )
    parser.add_argument(
        "--data-points",
        type=str,
        default=",".join(DEFAULT_DATA_POINTS),
        help=(
            "Comma-separated data points (columns after Company). Default: "
            + ", ".join(DEFAULT_DATA_POINTS)
        ),
    )

    args = parser.parse_args()
    companies = [c.strip() for c in args.companies.split(",") if c.strip()]
    data_points = [d.strip() for d in args["data-points"].split(",") if d.strip()] if isinstance(args, dict) else [d.strip() for d in getattr(args, "data_points").split(",") if d.strip()]

    df = create_sample_input_dataframe(companies, data_points)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(args.output, index=False)
    print(f"Wrote sample input to: {args.output.resolve()}")


if __name__ == "__main__":
    main()


