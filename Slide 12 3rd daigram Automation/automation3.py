import pandas as pd

# =========================================================
# INPUT / OUTPUT FILE PATH
# =========================================================
input_file = "PRP Sample Jun (2).xlsx"
output_file = "Risk_Output.xlsx"
source_sheet = "OneTrust - Risk Export"

# =========================================================
# LOAD DATA
# =========================================================
df = pd.read_excel(input_file, sheet_name=source_sheet, engine="openpyxl")

# Clean column names
df.columns = df.columns.astype(str).str.strip()

# =========================================================
# REQUIRED COLUMN VALIDATION
# =========================================================
required_columns = ["Organization", "ID", "Stage", "Aging"]

missing_columns = [col for col in required_columns if col not in df.columns]

if missing_columns:
    raise Exception(f"Missing required columns in input file: {missing_columns}")

# =========================================================
# CLEAN DATA
# =========================================================
df["Organization"] = df["Organization"].fillna("").astype(str).str.strip()
df["Stage"] = df["Stage"].fillna("").astype(str).str.strip()
df["Aging"] = df["Aging"].fillna("").astype(str).str.strip()

# Normalize columns
df["Stage_Normalized"] = df["Stage"].str.lower().str.strip()
df["Aging_Normalized"] = df["Aging"].str.lower().str.strip()

# Convert Aging values into Open / Overdue buckets
def aging_bucket(value):
    value = str(value).lower().strip()

    if "overdue" in value:
        return "Overdue"
    elif "open" in value:
        return "Open"
    elif value == "":
        return "Blank"
    else:
        return value.title()

df["Aging_Bucket"] = df["Aging_Normalized"].apply(aging_bucket)

print("Unique Aging values from Excel:")
print(df["Aging"].unique())

print("Unique Aging buckets used:")
print(df["Aging_Bucket"].unique())

# =========================================================
# CREATE FIRST PIVOT TABLE
# =========================================================
pivot = (
    df.groupby("Organization", as_index=False)["ID"]
    .count()
    .rename(columns={"ID": "Count of ID"})
)

pivot = pivot[pivot["Organization"] != ""]

# =========================================================
# FILTER CONDITIONS
# =========================================================
total_stage = ["evaluation", "identified", "treatment", "monitoring"]
open_stage = ["evaluation", "identified", "treatment"]

# =========================================================
# CREATE ZONE SUMMARY TABLE
# =========================================================
total_df = df[df["Stage_Normalized"].isin(total_stage)].copy()
open_df = df[df["Stage_Normalized"].isin(open_stage)].copy()

total_counts = (
    total_df.groupby("Organization", as_index=False)["ID"]
    .count()
    .rename(columns={"ID": "Total Risks"})
)

open_counts = (
    open_df.groupby("Organization", as_index=False)["ID"]
    .count()
    .rename(columns={"ID": "Open Risks"})
)

zones_df = pd.merge(
    total_counts,
    open_counts,
    on="Organization",
    how="left"
)

zones_df["Open Risks"] = zones_df["Open Risks"].fillna(0)
zones_df["Total Risks"] = zones_df["Total Risks"].astype(int)
zones_df["Open Risks"] = zones_df["Open Risks"].astype(int)

zones_df = zones_df.rename(columns={"Organization": "Zones"})
zones_df = zones_df[zones_df["Zones"] != ""]

# =========================================================
# OPEN VS OVERDUE TABLE
# =========================================================
open_df_age = df[df["Stage_Normalized"].isin(open_stage)].copy()

if len(open_df_age) > 0:
    pivot_open_overdue = (
        open_df_age.groupby(["Organization", "Aging_Bucket"])["ID"]
        .count()
        .unstack(fill_value=0)
        .reset_index()
    )
else:
    pivot_open_overdue = pd.DataFrame(columns=["Organization", "Open", "Overdue"])

pivot_open_overdue.columns = [str(col).strip() for col in pivot_open_overdue.columns]

# Force columns to exist
if "Open" not in pivot_open_overdue.columns:
    pivot_open_overdue["Open"] = 0

if "Overdue" not in pivot_open_overdue.columns:
    pivot_open_overdue["Overdue"] = 0

# =========================================================
# MAP ORGANIZATION TO ZONE CODES
# =========================================================
zone_map = {
    "Africa": "AFR",
    "APAC": "APAC",
    "Europe": "EUR",
    "GHQ": "GHQ",
    "Middle America Zone": "MAZ",
    "North America Zone": "NAZ",
    "South America Zone": "SAZ",
    "BEES": "GRO",
    "BEES | FINTECH": "GRO"
}

pivot_open_overdue["Zones"] = pivot_open_overdue["Organization"].map(zone_map)

zones_open_overdue = pivot_open_overdue[["Zones", "Open", "Overdue"]].copy()
zones_open_overdue = zones_open_overdue.dropna(subset=["Zones"])

zones_open_overdue["Open"] = zones_open_overdue["Open"].fillna(0).astype(int)
zones_open_overdue["Overdue"] = zones_open_overdue["Overdue"].fillna(0).astype(int)

zones_open_overdue = (
    zones_open_overdue.groupby("Zones", as_index=False)[["Open", "Overdue"]]
    .sum()
)

# =========================================================
# WRITE INTO EXCEL
# =========================================================
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:

    pivot.to_excel(writer, sheet_name="Pivot Table", index=False)
    zones_df.to_excel(writer, sheet_name="Zone Summary", index=False)
    zones_open_overdue.to_excel(writer, sheet_name="Open vs Overdue", index=False)

    workbook = writer.book

    # =====================================================
    # FORMATS
    # =====================================================
    header_format = workbook.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": "#000000",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    cell_format = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    # =====================================================
    # FORMAT PIVOT TABLE SHEET
    # =====================================================
    ws_pivot = writer.sheets["Pivot Table"]

    for col_num, column_name in enumerate(pivot.columns):
        ws_pivot.write(0, col_num, column_name, header_format)

    ws_pivot.set_column("A:A", 35, cell_format)
    ws_pivot.set_column("B:B", 15, cell_format)

    # =====================================================
    # FORMAT ZONE SUMMARY SHEET
    # =====================================================
    ws_zone = writer.sheets["Zone Summary"]

    for col_num, column_name in enumerate(zones_df.columns):
        ws_zone.write(0, col_num, column_name, header_format)

    ws_zone.set_column("A:A", 30, cell_format)
    ws_zone.set_column("B:C", 15, cell_format)

    # =====================================================
    # FORMAT OPEN VS OVERDUE SHEET
    # =====================================================
    ws_open_overdue = writer.sheets["Open vs Overdue"]

    for col_num, column_name in enumerate(zones_open_overdue.columns):
        ws_open_overdue.write(0, col_num, column_name, header_format)

    ws_open_overdue.set_column("A:A", 15, cell_format)
    ws_open_overdue.set_column("B:C", 15, cell_format)

    # =====================================================
    # FIRST CHART - ZONE SUMMARY
    # =====================================================
    if len(zones_df) > 0:

        chart = workbook.add_chart({
            "type": "column",
            "subtype": "stacked"
        })

        max_row = len(zones_df) + 1

        categories = f"='Zone Summary'!$A$2:$A${max_row}"
        total_values = f"='Zone Summary'!$B$2:$B${max_row}"
        open_values = f"='Zone Summary'!$C$2:$C${max_row}"

        chart.add_series({
            "name": "Total Risks",
            "categories": categories,
            "values": total_values,
            "fill": {"color": "#1f6a8a"},
            "border": {"none": True},
            "data_labels": {
                "value": True,
                "font": {"color": "white"}
            }
        })

        chart.add_series({
            "name": "Open Risks",
            "categories": categories,
            "values": open_values,
            "fill": {"color": "#f26c23"},
            "border": {"none": True},
            "data_labels": {
                "value": True,
                "font": {"color": "white"}
            }
        })

        chart.set_title({
            "name": "Zone wise Risks",
            "name_font": {
                "color": "white",
                "bold": True,
                "size": 14
            }
        })

        chart.set_x_axis({
            "num_font": {"color": "white"},
            "line": {"color": "white"}
        })

        # IMPORTANT FIX: "visible": True added
        chart.set_y_axis({
            "num_font": {"color": "white"},
            "line": {"color": "white"},
            "major_gridlines": {
                "visible": True,
                "line": {"color": "#444444"}
            }
        })

        chart.set_legend({
            "position": "bottom",
            "font": {"color": "white"}
        })

        chart.set_chartarea({
            "fill": {"color": "black"}
        })

        chart.set_plotarea({
            "fill": {"color": "black"}
        })

        chart.set_size({
            "width": 720,
            "height": 420
        })

        ws_zone.insert_chart("E2", chart)

    # =====================================================
    # SECOND CHART - OPEN VS OVERDUE
    # =====================================================
    if len(zones_open_overdue) > 0:

        chart2 = workbook.add_chart({
            "type": "column",
            "subtype": "stacked"
        })

        max_row2 = len(zones_open_overdue) + 1

        categories2 = f"='Open vs Overdue'!$A$2:$A${max_row2}"
        open_values2 = f"='Open vs Overdue'!$B$2:$B${max_row2}"
        overdue_values2 = f"='Open vs Overdue'!$C$2:$C${max_row2}"

        chart2.add_series({
            "name": "Open",
            "categories": categories2,
            "values": open_values2,
            "fill": {"color": "#1f6a8a"},
            "border": {"none": True},
            "data_labels": {
                "value": True,
                "font": {"color": "white"}
            }
        })

        chart2.add_series({
            "name": "Overdue",
            "categories": categories2,
            "values": overdue_values2,
            "fill": {"color": "#c00000"},
            "border": {"none": True},
            "data_labels": {
                "value": True,
                "font": {"color": "white"}
            }
        })

        chart2.set_title({
            "name": "Open vs Overdue Risks",
            "name_font": {
                "color": "white",
                "bold": True,
                "size": 14
            }
        })

        chart2.set_x_axis({
            "num_font": {"color": "white"},
            "line": {"color": "white"}
        })

        # IMPORTANT FIX: "visible": True added
        chart2.set_y_axis({
            "num_font": {"color": "white"},
            "line": {"color": "white"},
            "major_gridlines": {
                "visible": True,
                "line": {"color": "#444444"}
            }
        })

        chart2.set_legend({
            "position": "bottom",
            "font": {"color": "white"}
        })

        chart2.set_chartarea({
            "fill": {"color": "black"}
        })

        chart2.set_plotarea({
            "fill": {"color": "black"}
        })

        chart2.set_size({
            "width": 720,
            "height": 420
        })

        ws_open_overdue.insert_chart("E2", chart2)

print("Excel file created successfully:", output_file)