import pandas as pd
import xlsxwriter

# =========================
# INPUT / OUTPUT FILE PATH
# =========================
input_file = "PRP Sample Jun (2).xlsx"
output_file = "Risk_Output.xlsx"

# =========================
# LOAD DATA
# =========================
df = pd.read_excel(input_file, sheet_name="OneTrust - Risk Export")
df.columns = df.columns.str.strip()
print(df["Aging"].unique())

# =========================
# CREATE FIRST PIVOT TABLE
# =========================
pivot = pd.pivot_table(
    df,
    index="Organization",
    values="ID",
    aggfunc="count"
).reset_index()

pivot.rename(columns={"ID": "Count of ID"}, inplace=True)

# =========================
# FILTER CONDITIONS
# =========================
total_stage = ["Evaluation", "Identified", "Treatment", "Monitoring"]
open_stage = ["Evaluation", "Identified", "Treatment"]

# =========================
# CREATE ZONES TABLE
# =========================
total_df = df[df["Stage"].isin(total_stage)]
open_df = df[df["Stage"].isin(open_stage)]

total_counts = total_df.groupby("Organization")["ID"].count().reset_index()
open_counts = open_df.groupby("Organization")["ID"].count().reset_index()

total_counts.rename(columns={"ID": "Total Risks"}, inplace=True)
open_counts.rename(columns={"ID": "Open Risks"}, inplace=True)

zones_df = pd.merge(total_counts, open_counts, on="Organization", how="left").fillna(0)
zones_df.rename(columns={"Organization": "Zones"}, inplace=True)

# =========================================================
# ✅ NEW: OPEN vs OVERDUE PIVOT (FROM OPEN RISKS ONLY)
# =========================================================
open_df_age = df[df["Stage"].isin(open_stage)]

pivot_open_overdue = pd.pivot_table(
    open_df_age,
    index="Organization",
    columns="Aging",   # IMPORTANT column
    values="ID",
    aggfunc="count",
    fill_value=0
).reset_index()

pivot_open_overdue.columns.name = None

# Map Zones
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

zones_open_overdue = pivot_open_overdue[["Zones", "Open", "Overdue"]].dropna()

# =========================
# WRITE INTO EXCEL
# =========================
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:

    pivot.to_excel(writer, sheet_name="Pivot Table", index=False)
    zones_df.to_excel(writer, sheet_name="Zone Summary", index=False)
    zones_open_overdue.to_excel(writer, sheet_name="Open vs Overdue", index=False)

    workbook = writer.book

    # =========================
    # ✅ FIRST CHART (ZONE SUMMARY)
    # =========================
    worksheet = writer.sheets["Zone Summary"]

    if len(zones_df) > 0:
        chart = workbook.add_chart({"type": "column", "subtype": "stacked"})
        max_row = len(zones_df) + 1

        categories = f"='Zone Summary'!$A$2:$A${max_row}"

        chart.add_series({
            "name": "Tier-1 Supplier",
            "categories": categories,
            "values": f"='Zone Summary'!$B$2:$B${max_row}",
            "fill": {"color": "#1f6a8a"},
            "data_labels": {"value": True, "font": {"color": "white"}}
        })

        chart.add_series({
            "name": "Supplier Added by Zone",
            "categories": categories,
            "values": f"='Zone Summary'!$C$2:$C${max_row}",
            "fill": {"color": "#f26c23"},
            "data_labels": {"value": True, "font": {"color": "white"}}
        })

        chart.set_title({"name": "Zone wise Tier 1 Suppliers", "name_font": {"color": "white"}})

        chart.set_x_axis({"num_font": {"color": "white"}, "line": {"color": "white"}})
        chart.set_y_axis({"num_font": {"color": "white"},
                          "major_gridlines": {"visible": True, "line": {"color": "#444444"}}})

        chart.set_legend({"position": "bottom", "font": {"color": "white"}})

        chart.set_chartarea({"fill": {"color": "black"}})
        chart.set_plotarea({"fill": {"color": "black"}})

        worksheet.insert_chart("E2", chart)

    # =========================
    # ✅ SECOND CHART (OPEN vs OVERDUE)
    # =========================
    worksheet2 = writer.sheets["Open vs Overdue"]

    if len(zones_open_overdue) > 0:
        chart2 = workbook.add_chart({"type": "column", "subtype": "stacked"})
        max_row2 = len(zones_open_overdue) + 1

        categories2 = f"='Open vs Overdue'!$A$2:$A${max_row2}"

        chart2.add_series({
            "name": "Open",
            "categories": categories2,
            "values": f"='Open vs Overdue'!$B$2:$B${max_row2}",
            "fill": {"color": "#1f6a8a"},
            "data_labels": {"value": True, "font": {"color": "white"}}
        })

        chart2.add_series({
            "name": "Overdue",
            "categories": categories2,
            "values": f"='Open vs Overdue'!$C$2:$C${max_row2}",
            "fill": {"color": "#c00000"},
            "data_labels": {"value": True, "font": {"color": "white"}}
        })

        chart2.set_title({"name": "Open vs Overdue Risks", "name_font": {"color": "white"}})

        chart2.set_x_axis({"num_font": {"color": "white"}, "line": {"color": "white"}})
        chart2.set_y_axis({"num_font": {"color": "white"},
                           "major_gridlines": {"visible": True, "line": {"color": "#444444"}}})

        chart2.set_legend({"position": "bottom", "font": {"color": "white"}})

        chart2.set_chartarea({"fill": {"color": "black"}})
        chart2.set_plotarea({"fill": {"color": "black"}})

        worksheet2.insert_chart("E2", chart2)

print("✅ Excel file created successfully:", output_file)