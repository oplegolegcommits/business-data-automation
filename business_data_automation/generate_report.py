from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, LineChart, Reference


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_FILE = OUTPUT_DIR / "report.xlsx"

REQUIRED_COLUMNS = [
    "date",
    "product",
    "category",
    "price",
    "quantity",
    "revenue",
    "customer",
    "region",
]


# ============================================================
# DATA LOADING
# ============================================================

def load_data():
    csv_files = sorted(DATA_DIR.glob("*.csv"))

    if not csv_files:
        raise FileNotFoundError(
            f"No CSV files found in {DATA_DIR}"
        )

    print(f"Found {len(csv_files)} CSV files.")

    frames = []

    for file in csv_files:
        print(f"  Loading {file.name}...")

        df = pd.read_csv(file)

        missing_columns = set(REQUIRED_COLUMNS) - set(df.columns)

        if missing_columns:
            raise ValueError(
                f"{file.name} is missing columns: "
                f"{', '.join(sorted(missing_columns))}"
            )

        frames.append(df)

    data = pd.concat(frames, ignore_index=True)

    return data


# ============================================================
# CLEANING
# ============================================================

def clean_data(df):
    print("Cleaning data...")

    df = df.copy()

    # Dates
    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    # Numeric columns
    df["price"] = pd.to_numeric(df["price"], errors="coerce")
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")

    # Text columns
    text_columns = [
        "product",
        "category",
        "customer",
        "region",
    ]

    for column in text_columns:
        df[column] = df[column].astype(str).str.strip()

    # Remove invalid rows
    before = len(df)

    df = df.dropna(
        subset=[
            "date",
            "product",
            "category",
            "price",
            "quantity",
            "customer",
            "region",
        ]
    )

    # Remove impossible values
    df = df[
        (df["price"] >= 0) &
        (df["quantity"] > 0)
    ]

    # Remove duplicate transactions
    df = df.drop_duplicates()

    removed = before - len(df)

    # Recalculate revenue instead of trusting the CSV value
    df["revenue"] = df["price"] * df["quantity"]

    # Add month
    df["month"] = df["date"].dt.to_period("M").astype(str)

    print(f"  Removed {removed} invalid/duplicate rows.")
    print(f"  Final rows: {len(df)}")

    return df


# ============================================================
# METRICS
# ============================================================

def calculate_metrics(df):
    print("Calculating business metrics...")

    total_revenue = df["revenue"].sum()

    total_orders = df["customer"].count()

    units_sold = df["quantity"].sum()

    average_order_value = (
        total_revenue / total_orders
        if total_orders > 0
        else 0
    )

    # Monthly metrics
    monthly = (
        df.groupby("month")
        .agg(
            revenue=("revenue", "sum"),
            orders=("customer", "count"),
            units=("quantity", "sum"),
        )
        .reset_index()
    )

    monthly["aov"] = (
        monthly["revenue"] / monthly["orders"]
    )

    monthly["mom_growth"] = (
        monthly["revenue"].pct_change()
    )

    # Product analysis
    products = (
        df.groupby(["product", "category"])
        .agg(
            revenue=("revenue", "sum"),
            units=("quantity", "sum"),
            orders=("customer", "count"),
        )
        .reset_index()
        .sort_values("revenue", ascending=False)
    )

    products["revenue_share"] = (
        products["revenue"] / total_revenue * 100
    )

    # Region analysis
    regions = (
        df.groupby("region")
        .agg(
            revenue=("revenue", "sum"),
            orders=("customer", "count"),
            units=("quantity", "sum"),
        )
        .reset_index()
        .sort_values("revenue", ascending=False)
    )

    regions["aov"] = (
        regions["revenue"] / regions["orders"]
    )

    regions["revenue_share"] = (
        regions["revenue"] / total_revenue * 100
    )

    # Category analysis
    categories = (
        df.groupby("category")
        .agg(
            revenue=("revenue", "sum"),
            orders=("customer", "count"),
            units=("quantity", "sum"),
        )
        .reset_index()
        .sort_values("revenue", ascending=False)
    )

    categories["revenue_share"] = (
        categories["revenue"] / total_revenue * 100
    )

    metrics = {
        "total_revenue": total_revenue,
        "total_orders": total_orders,
        "units_sold": units_sold,
        "average_order_value": average_order_value,
        "monthly": monthly,
        "products": products,
        "regions": regions,
        "categories": categories,
    }

    return metrics


# ============================================================
# EXCEL REPORT
# ============================================================

def create_excel_report(df, metrics):
    print("Creating Excel report...")

    OUTPUT_DIR.mkdir(exist_ok=True)

    monthly = metrics["monthly"]
    products = metrics["products"]
    regions = metrics["regions"]
    categories = metrics["categories"]

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        # ----------------------------------------------------
        # DASHBOARD
        # ----------------------------------------------------

        dashboard_data = pd.DataFrame({
            "Metric": [
                "Total Revenue",
                "Total Orders",
                "Units Sold",
                "Average Order Value",
            ],
            "Value": [
                metrics["total_revenue"],
                metrics["total_orders"],
                metrics["units_sold"],
                metrics["average_order_value"],
            ],
        })

        dashboard_data.to_excel(
            writer,
            sheet_name="Dashboard",
            index=False,
            startrow=3,
        )

        # Monthly data
        monthly.to_excel(
            writer,
            sheet_name="Dashboard",
            index=False,
            startrow=10,
        )

        # ----------------------------------------------------
        # PRODUCT ANALYSIS
        # ----------------------------------------------------

        products.to_excel(
            writer,
            sheet_name="Products",
            index=False,
        )

        # ----------------------------------------------------
        # REGIONAL ANALYSIS
        # ----------------------------------------------------

        regions.to_excel(
            writer,
            sheet_name="Regions",
            index=False,
        )

        # ----------------------------------------------------
        # CATEGORY ANALYSIS
        # ----------------------------------------------------

        categories.to_excel(
            writer,
            sheet_name="Categories",
            index=False,
        )

        # ----------------------------------------------------
        # CLEANED DATA
        # ----------------------------------------------------

        df.to_excel(
            writer,
            sheet_name="Data",
            index=False,
        )

    # ========================================================
    # FORMAT WORKBOOK
    # ========================================================

    workbook = load_workbook(OUTPUT_FILE)

    for worksheet in workbook.worksheets:

        # Freeze header
        if worksheet.title != "Dashboard":
            worksheet.freeze_panes = "A2"

        # Header styling
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(
                horizontal="center"
            )

        # Auto column width
        for column in worksheet.columns:

            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:
                try:
                    length = len(str(cell.value))
                    max_length = max(max_length, length)
                except Exception:
                    pass

            worksheet.column_dimensions[
                column_letter
            ].width = min(max_length + 2, 35)

    # Dashboard title
    dashboard = workbook["Dashboard"]

    dashboard["A1"] = "BUSINESS PERFORMANCE REPORT"
    dashboard["A1"].font = Font(
        bold=True,
        size=18,
    )

    dashboard["A2"] = "Automatically generated from CSV sales data"
    dashboard["A2"].font = Font(
        italic=True,
    )

    # KPI formatting
    dashboard["B4"].number_format = '$#,##0.00'
    dashboard["B5"].number_format = '#,##0'
    dashboard["B6"].number_format = '#,##0'
    dashboard["B7"].number_format = '$#,##0.00'

    # Monthly formatting
    monthly_start_row = 11

    for row in range(
        monthly_start_row,
        monthly_start_row + len(monthly)
    ):
        dashboard[f"B{row}"].number_format = '$#,##0.00'
        dashboard[f"C{row}"].number_format = '#,##0'
        dashboard[f"D{row}"].number_format = '#,##0'
        dashboard[f"E{row}"].number_format = '$#,##0.00'
        dashboard[f"F{row}"].number_format = '0.0%'

    # Product formatting
    products_ws = workbook["Products"]

    for row in range(2, len(products) + 2):
        products_ws[f"C{row}"].number_format = '$#,##0.00'
        products_ws[f"D{row}"].number_format = '#,##0'
        products_ws[f"E{row}"].number_format = '#,##0'
        products_ws[f"F{row}"].number_format = '0.0%'

    # Region formatting
    regions_ws = workbook["Regions"]

    for row in range(2, len(regions) + 2):
        regions_ws[f"B{row}"].number_format = '$#,##0.00'
        regions_ws[f"C{row}"].number_format = '#,##0'
        regions_ws[f"D{row}"].number_format = '#,##0'
        regions_ws[f"E{row}"].number_format = '$#,##0.00'
        regions_ws[f"F{row}"].number_format = '0.0%'

    # Category formatting
    categories_ws = workbook["Categories"]

    for row in range(2, len(categories) + 2):
        categories_ws[f"B{row}"].number_format = '$#,##0.00'
        categories_ws[f"C{row}"].number_format = '#,##0'
        categories_ws[f"D{row}"].number_format = '#,##0'
        categories_ws[f"E{row}"].number_format = '0.0%'

    # ========================================================
    # CHARTS
    # ========================================================

    # 1. Monthly Revenue
    revenue_chart = LineChart()

    revenue_chart.title = "Revenue by Month"
    revenue_chart.y_axis.title = "Revenue"
    revenue_chart.x_axis.title = "Month"

    revenue_data = Reference(
        dashboard,
        min_col=2,
        min_row=10,
        max_row=10 + len(monthly),
    )

    revenue_categories = Reference(
        dashboard,
        min_col=1,
        min_row=11,
        max_row=10 + len(monthly),
    )

    revenue_chart.add_data(
        revenue_data,
        titles_from_data=True,
    )

    revenue_chart.set_categories(
        revenue_categories
    )

    dashboard.add_chart(
        revenue_chart,
        "H3",
    )

    # 2. Top products
    product_chart = BarChart()

    product_chart.type = "bar"
    product_chart.style = 10
    product_chart.title = "Top 10 Products by Revenue"
    product_chart.y_axis.title = "Product"
    product_chart.x_axis.title = "Revenue"

    top_products = products.head(10)

    product_data = Reference(
        products_ws,
        min_col=3,
        min_row=1,
        max_row=len(top_products) + 1,
    )

    product_categories = Reference(
        products_ws,
        min_col=1,
        min_row=2,
        max_row=len(top_products) + 1,
    )

    product_chart.add_data(
        product_data,
        titles_from_data=True,
    )

    product_chart.set_categories(
        product_categories
    )

    product_chart.height = 8
    product_chart.width = 14

    dashboard.add_chart(
        product_chart,
        "H20",
    )

    # 3. Revenue by region
    region_chart = BarChart()

    region_chart.type = "col"
    region_chart.title = "Revenue by Region"
    region_chart.y_axis.title = "Revenue"
    region_chart.x_axis.title = "Region"

    region_data = Reference(
        regions_ws,
        min_col=2,
        min_row=1,
        max_row=len(regions) + 1,
    )

    region_categories = Reference(
        regions_ws,
        min_col=1,
        min_row=2,
        max_row=len(regions) + 1,
    )

    region_chart.add_data(
        region_data,
        titles_from_data=True,
    )

    region_chart.set_categories(
        region_categories
    )

    region_chart.height = 8
    region_chart.width = 14

    dashboard.add_chart(
        region_chart,
        "X3",
    )

    # 4. Revenue by category
    category_chart = BarChart()

    category_chart.type = "col"
    category_chart.title = "Revenue by Category"
    category_chart.y_axis.title = "Revenue"
    category_chart.x_axis.title = "Category"

    category_data = Reference(
        categories_ws,
        min_col=2,
        min_row=1,
        max_row=len(categories) + 1,
    )

    category_categories = Reference(
        categories_ws,
        min_col=1,
        min_row=2,
        max_row=len(categories) + 1,
    )

    category_chart.add_data(
        category_data,
        titles_from_data=True,
    )

    category_chart.set_categories(
        category_categories
    )

    category_chart.height = 8
    category_chart.width = 14

    dashboard.add_chart(
        category_chart,
        "X20",
    )

    workbook.save(OUTPUT_FILE)


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 55)
    print("BUSINESS DATA AUTOMATION")
    print("=" * 55)

    try:

        # 1. Load
        df = load_data()

        # 2. Clean
        df = clean_data(df)

        # 3. Calculate metrics
        metrics = calculate_metrics(df)

        # 4. Generate Excel
        create_excel_report(
            df,
            metrics,
        )

        print()
        print("=" * 55)
        print("REPORT GENERATED SUCCESSFULLY")
        print("=" * 55)
        print(f"File: {OUTPUT_FILE}")
        print()
        print(
            f"Revenue:       ${metrics['total_revenue']:,.2f}"
        )
        print(
            f"Orders:        {metrics['total_orders']:,}"
        )
        print(
            f"Units sold:    {metrics['units_sold']:,}"
        )
        print(
            f"Average order: ${metrics['average_order_value']:,.2f}"
        )
        print("=" * 55)

    except Exception as error:

        print()
        print("ERROR:")
        print(error)


if __name__ == "__main__":
    main()