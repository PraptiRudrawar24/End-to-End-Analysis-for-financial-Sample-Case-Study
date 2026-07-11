"""
03_visualize.py
----------------
Step 3 of the pipeline: reproduce every chart used in the final report with
Matplotlib / Seaborn, saved to charts/.
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import pandas as pd
import seaborn as sns

sns.set_style("whitegrid")

CLEAN_CSV_PATH = "data/cleaned_financial_data.csv"
CHART_DIR = "charts"
BAND_ORDER = ["None", "Low", "Medium", "High"]


def load_data() -> pd.DataFrame:
    df = pd.read_csv(CLEAN_CSV_PATH, parse_dates=["Date"], keep_default_na=False,
                      na_values=[""])
    df["Discount Band"] = pd.Categorical(df["Discount Band"], categories=BAND_ORDER, ordered=True)
    return df


def chart_profit_by_segment(df: pd.DataFrame):
    seg = df.groupby("Segment", observed=True)["Profit"].sum().sort_values(ascending=False)
    colors = ["#d62728" if v < 0 else "#1f77b4" for v in seg.values]

    fig, ax = plt.subplots(figsize=(9, 6))
    ax.barh(seg.index[::-1], seg.values[::-1], color=colors[::-1])
    ax.axvline(0, color="black", linewidth=1)
    ax.set_title("Total Profit by Segment", fontsize=15, fontweight="bold")
    ax.set_xlabel("Total Profit ($)")
    ax.set_ylabel("Segment")
    fig.tight_layout()
    fig.savefig(f"{CHART_DIR}/01_profit_by_segment.png", dpi=150)
    plt.close(fig)


def chart_margin_by_discount_band(df: pd.DataFrame):
    band = df.groupby("Discount Band", observed=True).apply(
        lambda g: 100 * g["Profit"].sum() / g["Sales"].sum()
    ).reindex(BAND_ORDER)

    fig, ax = plt.subplots(figsize=(9, 6))
    bars = ax.bar(band.index, band.values, color=sns.color_palette("rocket", len(band)))
    for bar, val in zip(bars, band.values):
        ax.text(bar.get_x() + bar.get_width() / 2, val + 0.4, f"{val:.1f}%",
                ha="center", fontweight="bold")
    ax.set_title("Profit Margin % by Discount Band", fontsize=15, fontweight="bold")
    ax.set_xlabel("Discount Band")
    ax.set_ylabel("Profit Margin (%)")
    fig.tight_layout()
    fig.savefig(f"{CHART_DIR}/02_margin_by_discount_band.png", dpi=150)
    plt.close(fig)


def chart_heatmap_segment_discount(df: pd.DataFrame):
    pivot = df.pivot_table(index="Segment", columns="Discount Band",
                            values=["Profit", "Sales"], aggfunc="sum", observed=True)
    margin = 100 * pivot["Profit"] / pivot["Sales"]
    margin = margin[BAND_ORDER]
    margin = margin.reindex(margin.mean(axis=1).sort_values(ascending=False).index)

    fig, ax = plt.subplots(figsize=(9, 6))
    sns.heatmap(margin, annot=True, fmt=".1f", cmap="RdYlGn", center=0, ax=ax,
                cbar_kws={"label": "Profit Margin %"})
    ax.set_title("Profit Margin % — Segment vs Discount Band", fontsize=15, fontweight="bold")
    ax.set_xlabel("Discount Band")
    ax.set_ylabel("Segment")
    fig.tight_layout()
    fig.savefig(f"{CHART_DIR}/03_heatmap_segment_discount.png", dpi=150)
    plt.close(fig)


def chart_sales_profit_by_product(df: pd.DataFrame):
    prod = df.groupby("Product")[["Sales", "Profit"]].sum().sort_values("Sales", ascending=False)

    fig, ax = plt.subplots(figsize=(9, 6))
    ax.bar(prod.index, prod["Sales"], label="Sales", color="#9ecae1")
    ax.bar(prod.index, prod["Profit"], label="Profit", color="#1f77b4")
    ax.set_title("Sales vs Profit by Product", fontsize=15, fontweight="bold")
    ax.legend()
    fig.tight_layout()
    fig.savefig(f"{CHART_DIR}/04_sales_profit_by_product.png", dpi=150)
    plt.close(fig)


def chart_sales_by_country(df: pd.DataFrame):
    country = df.groupby("Country").apply(
        lambda g: pd.Series({"Sales": g["Sales"].sum(),
                              "Margin": 100 * g["Profit"].sum() / g["Sales"].sum()})
    ).sort_values("Sales", ascending=False)

    fig, ax = plt.subplots(figsize=(9, 6))
    bars = ax.bar(country.index, country["Sales"], color=sns.color_palette("mako", len(country)))
    for bar, m in zip(bars, country["Margin"]):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 2e5, f"{m:.1f}%",
                ha="center", fontweight="bold", fontsize=9)
    ax.set_title("Total Sales by Country (label = Profit Margin %)", fontsize=15, fontweight="bold")
    ax.set_ylabel("Sales ($)")
    plt.xticks(rotation=20, ha="right")
    fig.tight_layout()
    fig.savefig(f"{CHART_DIR}/05_sales_by_country.png", dpi=150)
    plt.close(fig)


def chart_monthly_trend(df: pd.DataFrame):
    monthly = df.set_index("Date").resample("MS")[["Sales", "Profit"]].sum()

    fig, ax1 = plt.subplots(figsize=(10, 6))
    ax2 = ax1.twinx()
    ax1.plot(monthly.index, monthly["Sales"], marker="o", color="tab:blue", label="Sales")
    ax2.plot(monthly.index, monthly["Profit"], marker="s", color="tab:green", label="Profit")
    ax1.set_ylabel("Sales ($)", color="tab:blue")
    ax2.set_ylabel("Profit ($)", color="tab:green")
    ax1.set_title("Monthly Sales & Profit Trend", fontsize=15, fontweight="bold")
    plt.xticks(monthly.index, [d.strftime("%Y-%m") for d in monthly.index], rotation=45, ha="right")
    fig.tight_layout()
    fig.savefig(f"{CHART_DIR}/06_monthly_trend.png", dpi=150)
    plt.close(fig)


def main():
    df = load_data()
    chart_profit_by_segment(df)
    chart_margin_by_discount_band(df)
    chart_heatmap_segment_discount(df)
    chart_sales_profit_by_product(df)
    chart_sales_by_country(df)
    chart_monthly_trend(df)
    print("All 6 charts saved to charts/")


if __name__ == "__main__":
    main()
