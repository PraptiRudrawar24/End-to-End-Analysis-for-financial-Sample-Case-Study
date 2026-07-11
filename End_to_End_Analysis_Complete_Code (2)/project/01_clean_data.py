import pandas as pd

RAW_PATH = "data/Sample_data.xlsx"
CLEAN_CSV_PATH = "data/cleaned_financial_data.csv"


def load_raw(path: str = RAW_PATH) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [c.strip() for c in df.columns]
    return df


def validate_and_clean(df: pd.DataFrame) -> pd.DataFrame:
    print(f"Raw shape: {df.shape}")
    print(f"Duplicate rows: {df.duplicated().sum()}")
    print("\nMissing values per column:")
    print(df.isna().sum()[df.isna().sum() > 0])

    
    missing_band = df["Discount Band"].isna()
    all_zero_discount = (df.loc[missing_band, "Discounts"] == 0).all()
    print(f"\nAll {missing_band.sum()} blank Discount Band rows have Discounts == 0: "
          f"{all_zero_discount}")

    df.loc[missing_band, "Discount Band"] = "None"

    
    gross_check = (df["Units Sold"] * df["Sale Price"] - df["Gross Sales"]).abs() > 0.01
    sales_check = (df["Gross Sales"] - df["Discounts"] - df["Sales"]).abs() > 0.01
    profit_check = (df["Sales"] - df["COGS"] - df["Profit"]).abs() > 0.01

    print(f"\nGross Sales mismatches: {gross_check.sum()}")
    print(f"Sales mismatches: {sales_check.sum()}")
    print(f"Profit mismatches: {profit_check.sum()}")

    
    band_order = ["None", "Low", "Medium", "High"]
    df["Discount Band"] = pd.Categorical(df["Discount Band"], categories=band_order, ordered=True)

    df["Date"] = pd.to_datetime(df["Date"])

    return df


def main():
    df = load_raw()
    df = validate_and_clean(df)
    df.to_csv(CLEAN_CSV_PATH, index=False)
    print(f"\nCleaned dataset saved to {CLEAN_CSV_PATH} -> shape {df.shape}")


if __name__ == "__main__":
    main()
