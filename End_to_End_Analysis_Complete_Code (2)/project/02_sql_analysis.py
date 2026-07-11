"""
02_sql_analysis.py
-------------------
Step 2 of the pipeline: load the cleaned dataset into a SQLite database and
run the SQL queries that power the report's findings (profitability by
segment, discount-band impact, product/geo performance, monthly trend).
"""

import sqlite3
import pandas as pd

CLEAN_CSV_PATH = "data/cleaned_financial_data.csv"
DB_PATH = "data/financial_data.db"


def build_database(csv_path: str = CLEAN_CSV_PATH, db_path: str = DB_PATH) -> sqlite3.Connection:
    # keep_default_na=False so the literal string "None" (a valid Discount
    # Band label) is not silently converted to a missing value
    df = pd.read_csv(csv_path, parse_dates=["Date"], keep_default_na=False,
                      na_values=[""])
    conn = sqlite3.connect(db_path)
    df.to_sql("sales", conn, if_exists="replace", index=False)
    return conn


def run_query(conn, label: str, sql: str) -> pd.DataFrame:
    print(f"\n--- {label} ---")
    result = pd.read_sql_query(sql, conn)
    print(result.to_string(index=False))
    return result


def main():
    conn = build_database()

    q_segment = """
        SELECT Segment,
               ROUND(SUM(Sales), 2)  AS Sales,
               ROUND(SUM(Profit), 2) AS Profit,
               ROUND(100.0 * SUM(Profit) / SUM(Sales), 2) AS Margin_Pct
        FROM sales
        GROUP BY Segment
        ORDER BY Profit DESC;
    """
    run_query(conn, "Profitability by Segment", q_segment)

    q_discount_band = """
        SELECT "Discount Band" AS Discount_Band,
               ROUND(100.0 * SUM(Profit) / SUM(Sales), 2) AS Margin_Pct
        FROM sales
        GROUP BY "Discount Band"
        ORDER BY CASE "Discount Band"
                    WHEN 'None' THEN 1 WHEN 'Low' THEN 2
                    WHEN 'Medium' THEN 3 WHEN 'High' THEN 4 END;
    """
    run_query(conn, "Margin % by Discount Band", q_discount_band)

    q_segment_band = """
        SELECT Segment, "Discount Band" AS Discount_Band,
               ROUND(100.0 * SUM(Profit) / SUM(Sales), 2) AS Margin_Pct
        FROM sales
        GROUP BY Segment, "Discount Band"
        ORDER BY Segment;
    """
    run_query(conn, "Margin % by Segment x Discount Band", q_segment_band)

    q_enterprise_losses = """
        SELECT "Discount Band" AS Discount_Band, COUNT(*) AS Loss_Orders
        FROM sales
        WHERE Segment = 'Enterprise' AND Profit < 0
        GROUP BY "Discount Band";
    """
    run_query(conn, "Enterprise Loss-Making Orders by Discount Band", q_enterprise_losses)

    q_product = """
        SELECT Product,
               ROUND(SUM(Sales), 2) AS Sales,
               ROUND(SUM(Profit), 2) AS Profit,
               ROUND(100.0 * SUM(Profit) / SUM(Sales), 2) AS Margin_Pct
        FROM sales
        GROUP BY Product
        ORDER BY Sales DESC;
    """
    run_query(conn, "Sales & Profit by Product", q_product)

    q_country = """
        SELECT Country,
               ROUND(SUM(Sales), 2) AS Sales,
               ROUND(100.0 * SUM(Profit) / SUM(Sales), 2) AS Margin_Pct
        FROM sales
        GROUP BY Country
        ORDER BY Sales DESC;
    """
    run_query(conn, "Sales & Margin by Country", q_country)

    q_monthly = """
        SELECT strftime('%Y-%m', Date) AS Month,
               ROUND(SUM(Sales), 2) AS Sales,
               ROUND(SUM(Profit), 2) AS Profit
        FROM sales
        GROUP BY Month
        ORDER BY Month;
    """
    run_query(conn, "Monthly Sales & Profit Trend", q_monthly)

    conn.close()
    print(f"\nSQLite database written to {DB_PATH}")


if __name__ == "__main__":
    main()
