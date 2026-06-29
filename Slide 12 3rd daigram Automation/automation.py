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
df = pd.read_excel(input_file, sheet_name="OneTrust - Risk Export")  # update sheet if needed

# Clean column names (remove hidden characters like \r)
df.columns = df.columns.str.strip()

# =========================
# CREATE PIVOT TABLE
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

# Merge both
zones_df = pd.merge(total_counts, open_counts, on="Organization", how="left")
zones_df.fillna(0, inplace=True)

# Rename Organization → Zones (as per your image)
zones_df.rename(columns={"Organization": "Zones"}, inplace=True)

# =========================
# WRITE INTO EXCEL + CHART
# =========================
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    pivot.to_excel(writer, sheet_name="Pivot Table", index=False)
    zones_df.to_excel(writer, sheet_name="Zone Summary", index=False)

    workbook = writer.book
    worksheet = writer.sheets["Zone Summary"]

    # =========================
    # CREATE BAR CHART
    # =========================
    # =========================
# CREATE DARK THEME CHART
# =========================
chart = workbook.add_chart({"type": "column", "subtype": "stacked"})

categories = f"=Zone Summary!$A$2:$A${len(zones_df)+1}"

# Series 1 → Total Risks (Blue)
chart.add_series({
    "name": "Tier-1 Supplier",
    "categories": categories,
    "values": f"=Zone Summary!$B$2:$B${len(zones_df)+1}",
    "fill": {"color": "#1f6a8a"},   # dark blue
    "data_labels": {
        "value": True,
        "font": {"color": "white"}
    }
})

# Series 2 → Open Risks (Orange)
chart.add_series({
    "name": "Supplier Added by Zone",
    "categories": categories,
    "values": f"=Zone Summary!$C$2:$C${len(zones_df)+1}",
    "fill": {"color": "#f26c23"},   # orange
    "data_labels": {
        "value": True,
        "font": {"color": "white"}
    }
})

# Title
chart.set_title({
    "name": "Zone wise Tier 1 Suppliers",
    "name_font": {"color": "white", "bold": True, "size": 16}
})

# Axis styling
chart.set_x_axis({
    "name": "Zones",
    "name_font": {"color": "white"},
    "num_font": {"color": "white"},
    "line": {"color": "white"}
})

chart.set_y_axis({
    "name": "",
    "name_font": {"color": "white"},
    "num_font": {"color": "white"},
    "major_gridlines": {
        "visible": True,
        "line": {"color": "#444444"}
    }
})

# Legend
chart.set_legend({
    "position": "bottom",
    "font": {"color": "white"}
})

# Chart background
chart.set_chartarea({
    "fill": {"color": "black"},
    "border": {"none": True}
})

# Plot area (inside)
chart.set_plotarea({
    "fill": {"color": "black"}
})

# Insert chart
worksheet.insert_chart("E2", chart)

    

print("✅ Excel file created successfully:", output_file)