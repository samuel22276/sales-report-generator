import pandas as pd
from pathlib import Path
import sys

def load_sales_data(file_path):
    sales = pd.read_csv(file_path)
    return sales

def validate_data(sales):
    required_columns = [
        "Date",
        "Product",
        "Category",
        "Region",
        "Quantity",
        "Price"
    ]

    missing_columns = [
        column for column in required_columns
        if column not in sales.columns
    ]

    if missing_columns:
        print("ERROR: Missing columns:", missing_columns)
        return False

    if sales[required_columns].isnull().sum().sum() > 0:
        print("WARNING: Missing values detected.")

    if (sales["Quantity"] < 0).any():
        print("ERROR: Negative quantities detected.")
        return False

    if (sales["Price"] < 0).any():
        print("ERROR: Negative prices detected.")
        return False

    print("Data validation complete.")
    return True

def calculate_metrics(sales):
    sales["Revenue"] = sales["Quantity"] * sales["Price"]

    total_revenue = sales["Revenue"].sum()
    total_units = sales["Quantity"].sum()
    total_orders = len(sales)
    average_order_value = total_revenue / total_orders

    return (
        sales,
        total_revenue,
        total_units,
        total_orders,
        average_order_value
    )

def analyze_products(sales):
    product_sales = sales.groupby("Product")["Revenue"].sum()
    return product_sales.sort_values(ascending=False)


def analyze_regions(sales):
    regional_sales = sales.groupby("Region")["Revenue"].sum()
    return regional_sales.sort_values(ascending=False)

def analyze_monthly_sales(sales):
    sales["Date"] = pd.to_datetime(sales["Date"])

    monthly_sales = (
        sales.groupby(sales["Date"].dt.to_period("M"))["Revenue"]
        .sum()
    )

    return monthly_sales



if len(sys.argv) > 1:
    file_path = Path(sys.argv[1])
else:
    file_path = Path.home() / "Documents" / "sales-report-generator" / "data" / "sales.csv"
sales = load_sales_data(file_path)

if not validate_data(sales):
    exit()

# -----------------------------
# DATA VALIDATION
# -----------------------------

required_columns = [
    "Date",
    "Product",
    "Category",
    "Region",
    "Quantity",
    "Price"
]

# Check required columns
missing_columns = [
    column for column in required_columns
    if column not in sales.columns
]

if missing_columns:
    print("ERROR: Missing columns:", missing_columns)
    exit()

# Check for missing values
missing_values = sales[required_columns].isnull().sum()

if missing_values.sum() > 0:
    print("WARNING: Missing values detected:")
    print(missing_values[missing_values > 0])

# Check for invalid quantities
if (sales["Quantity"] < 0).any():
    print("ERROR: Negative quantities detected.")
    exit()

# Check for invalid prices
if (sales["Price"] < 0).any():
    print("ERROR: Negative prices detected.")
    exit()

print("Data validation complete.")

print(sales)

(
    sales,
    total_revenue,
    total_units,
    total_orders,
    average_order_value
) = calculate_metrics(sales)

# -----------------------------
# MONTHLY SALES ANALYSIS
# -----------------------------

monthly_sales = analyze_monthly_sales(sales)

print("\nMonthly Revenue:")
print(monthly_sales)

print("Average Order Value:", average_order_value)

product_sales = analyze_products(sales)

print(product_sales)

regional_sales = analyze_regions(sales)

print("\nRevenue by Region:")
print(regional_sales)

import matplotlib.pyplot as plt

regional_sales.plot(kind="bar")

plt.title("Revenue by Region")
plt.xlabel("Region")
plt.ylabel("Revenue ($)")
plt.tight_layout()

plt.savefig("../output/revenue_by_region.png")

plt.show()

product_sales.plot(kind="bar")

plt.title("Revenue by Product")
plt.xlabel("Product")
plt.ylabel("Revenue ($)")
plt.xticks(rotation=45)
plt.tight_layout()

plt.savefig("../output/revenue_by_product.png")

plt.show()

excel_file = "../output/sales_report.xlsx"

with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
    sales.to_excel(writer, sheet_name="Sales Data", index=False)
    product_sales.to_excel(writer, sheet_name="Product Revenue")
    regional_sales.to_excel(writer, sheet_name="Regional Revenue")
    monthly_sales.to_excel(writer, sheet_name="Monthly Revenue")    

print("Excel report created:", excel_file)

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.chart import BarChart, Reference

# Load the Excel workbook
workbook = load_workbook(excel_file)

# Create Dashboard sheet
dashboard = workbook.create_sheet("Dashboard", 0)

# Title
dashboard["A1"] = "SALES PERFORMANCE DASHBOARD"
dashboard["A1"].font = Font(size=14, bold=True)
dashboard["A1"].alignment = Alignment(horizontal="center")

# KPI labels
dashboard["A3"] = "Total Revenue"
dashboard["C3"] = "Total Units"
dashboard["E3"] = "Total Orders"
dashboard["G3"] = "Average Order Value"

# KPI values
dashboard["A4"] = total_revenue
dashboard["C4"] = total_units
dashboard["E4"] = total_orders
dashboard["G4"] = average_order_value

# Make KPI values bold
for cell in ["A4", "C4", "E4", "G4"]:
    dashboard[cell].font = Font(size=16, bold=True)

# Currency formatting
dashboard["A4"].number_format = '$#,##0.00'
dashboard["G4"].number_format = '$#,##0.00'

# Column widths
for column in ["A", "C", "E", "G"]:
    dashboard.column_dimensions[column].width = 22

# Save workbook
workbook.save(excel_file)

print("Dashboard added successfully!")

# Create product revenue chart
chart = BarChart()

chart.title = "Revenue by Product"
chart.y_axis.title = "Revenue ($)"
chart.x_axis.title = "Product"

data = Reference(
    workbook["Product Revenue"],
    min_col=2,
    min_row=1,
    max_row=6
)

categories = Reference(
    workbook["Product Revenue"],
    min_col=1,
    min_row=2,
    max_row=6
)

chart.add_data(data, titles_from_data=True)
chart.set_categories(categories)

chart.height = 8
chart.width = 14

dashboard.add_chart(chart, "A7")

# Save workbook
workbook.save(excel_file)

print("Dashboard chart added!")

# Create regional revenue chart
regional_chart = BarChart()

regional_chart.title = "Revenue by Region"
regional_chart.y_axis.title = "Revenue ($)"
regional_chart.x_axis.title = "Region"

data = Reference(
    workbook["Regional Revenue"],
    min_col=2,
    min_row=1,
    max_row=4
)

categories = Reference(
    workbook["Regional Revenue"],
    min_col=1,
    min_row=2,
    max_row=4
)

regional_chart.add_data(data, titles_from_data=True)
regional_chart.set_categories(categories)

regional_chart.height = 8
regional_chart.width = 14

dashboard.add_chart(regional_chart, "P7")

# Save workbook

from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# -----------------------------
# PROFESSIONAL FORMATTING
# -----------------------------

# Dashboard title
dashboard.merge_cells("A1:G1")
dashboard["A1"].font = Font(size=20, bold=True)
dashboard["A1"].alignment = Alignment(horizontal="center")

# Format KPI labels
for cell in ["A3", "C3", "E3", "G3"]:
    dashboard[cell].font = Font(bold=True, size=11)
    dashboard[cell].alignment = Alignment(horizontal="center")

# Format KPI values
for cell in ["A4", "C4", "E4", "G4"]:
    dashboard[cell].font = Font(bold=True, size=16)
    dashboard[cell].alignment = Alignment(horizontal="center")

# Currency formatting
dashboard["A4"].number_format = '$#,##0.00'
dashboard["G4"].number_format = '$#,##0.00'

# Borders for KPI cells
thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

for cell in ["A3", "A4", "C3", "C4", "E3", "E4", "G3", "G4"]:
    dashboard[cell].border = thin_border

# Center KPI columns
for column in ["A", "C", "E", "G"]:
    dashboard.column_dimensions[column].width = 22

# -----------------------------
# FORMAT SALES DATA
# -----------------------------

sales_sheet = workbook["Sales Data"]

sales_sheet.freeze_panes = "A2"

# Header formatting
for cell in sales_sheet[1]:
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center")

# Currency column
for cell in sales_sheet["F"][1:]:
    cell.number_format = '$#,##0.00'

# Revenue column
for cell in sales_sheet["G"][1:]:
    cell.number_format = '$#,##0.00'

# Auto-size columns
for column in sales_sheet.columns:
    max_length = 0
    column_letter = get_column_letter(column[0].column)

    for cell in column:
        if cell.value is not None:
            max_length = max(max_length, len(str(cell.value)))

    sales_sheet.column_dimensions[column_letter].width = max_length + 2

# -----------------------------
# FORMAT PRODUCT REVENUE
# -----------------------------

product_sheet = workbook["Product Revenue"]

product_sheet.freeze_panes = "A2"

for cell in product_sheet[1]:
    cell.font = Font(bold=True)

for cell in product_sheet["B"][1:]:
    cell.number_format = '$#,##0.00'

product_sheet.column_dimensions["A"].width = 20
product_sheet.column_dimensions["B"].width = 18

# -----------------------------
# FORMAT REGIONAL REVENUE
# -----------------------------

region_sheet = workbook["Regional Revenue"]

region_sheet.freeze_panes = "A2"

for cell in region_sheet[1]:
    cell.font = Font(bold=True)

for cell in region_sheet["B"][1:]:
    cell.number_format = '$#,##0.00'

region_sheet.column_dimensions["A"].width = 20
region_sheet.column_dimensions["B"].width = 18

# Save final workbook

print("Professional formatting complete!")

# -----------------------------
# DASHBOARD LAYOUT
# -----------------------------

# Row heights
dashboard.row_dimensions[1].height = 35
dashboard.row_dimensions[3].height = 25
dashboard.row_dimensions[4].height = 30

# Center the KPI area
for row in dashboard.iter_rows(min_row=3, max_row=4, min_col=1, max_col=7):
    for cell in row:
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

# Add spacing between KPI sections
dashboard.column_dimensions["B"].width = 4
dashboard.column_dimensions["D"].width = 4
dashboard.column_dimensions["F"].width = 4

# Make dashboard title larger
dashboard["A1"].font = Font(
    size=22,
    bold=True
)

# Add a subtitle
dashboard["A2"] = "Automated Sales Performance Report"
dashboard.merge_cells("A2:G2")

dashboard["A2"].font = Font(
    size=11,
    italic=True
)

dashboard["A2"].alignment = Alignment(
    horizontal="center"
)

# Make the KPI labels larger
for cell in ["A3", "C3", "E3", "G3"]:
    dashboard[cell].font = Font(
        size=12,
        bold=True
    )

# Make KPI numbers stand out
for cell in ["A4", "C4", "E4", "G4"]:
    dashboard[cell].font = Font(
        size=18,
        bold=True
    )

workbook.save(excel_file)

print("Regional chart added!")
