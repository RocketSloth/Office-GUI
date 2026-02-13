import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def style_html_table(html: str) -> str:
    """
    Take the raw HTML from pandas.DataFrame.to_html() and inject
    inline styles so the table looks good in an email client
    (Atmos-style blues, simple borders, readable font).
    """
    # Style the <table> tag
    html = html.replace(
        '<table border="0" class="dataframe">',
        '<table style="border-collapse:collapse; font-family:Arial, sans-serif; '
        'font-size:12px; color:#333333;">'
    )

    # Style header cells
    html = html.replace(
        "<th>",
        '<th style="background-color:#005596; color:#ffffff; padding:4px 8px; '
        'border:1px solid #cccccc; text-align:left;">'
    )

    # Style body cells
    html = html.replace(
        "<td>",
        '<td style="padding:4px 8px; border:1px solid #cccccc; text-align:left;">'
    )

    return html

def main():
    # -----------------------------
    # 1) Set up Tkinter (no main window)
    # -----------------------------
    root = tk.Tk()
    root.withdraw()

    # -----------------------------
    # 2) Ask user to select the input Excel file
    # -----------------------------
    messagebox.showinfo("Select File", "Select the Excel file you want to analyze.")
    input_path = filedialog.askopenfilename(
        title="Select the Pavement Export Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls")]
    )

    if not input_path:
        messagebox.showerror("Cancelled", "No file selected. Exiting.")
        return

    # -----------------------------
    # 3) Load the selected Excel file
    # -----------------------------
    try:
        df = pd.read_excel(input_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file:\n{e}")
        return

    # -----------------------------
    # 4) Check that required columns exist
    # -----------------------------
    required_cols = [
        "InvoiceNumber",
        "DateAssigned",
        "DateCut",
        "CutAtmosSupervisor",
        "PavementKey",
        "Address",
    ]

    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        messagebox.showerror(
            "Error",
            f"The following required columns are missing from the file:\n{', '.join(missing_cols)}"
        )
        return

    # -----------------------------
    # 5) Filter rows
    #    - Remove rows where InvoiceNumber = "Cancelled"
    #    - Keep only rows where DateAssigned is blank
    # -----------------------------
    df_filtered = df[df["InvoiceNumber"] != "Cancelled"].copy()
    df_missing = df_filtered[df_filtered["DateAssigned"].isna()].copy()

    # -----------------------------
    # 6) Add DaysMissing based on DateCut
    #     DaysMissing = TODAY - DateCut
    # -----------------------------
    today = pd.Timestamp.today()

    # Make sure DateCut is treated as a datetime
    df_missing["DateCut"] = pd.to_datetime(df_missing["DateCut"], errors="coerce")

    # Calculate days difference where DateCut is valid
    df_missing["DaysMissing"] = (today - df_missing["DateCut"]).dt.days

    # Treat blank supervisors as "Unassigned" for reporting
    df_missing["CutAtmosSupervisor"] = df_missing["CutAtmosSupervisor"].fillna("Unassigned")

    # Filter to only jobs missing 14+ days
    df_missing = df_missing[
        (df_missing["DaysMissing"] >= 14) &
        (df_missing["DaysMissing"] < 90)
        ]
    
    # If no rows remain, notify user
    if df_missing.empty:
        messagebox.showinfo(
            "No Results",
            "No jobs have been missing DateAssigned between 14 and 365 days."
        )
        return





    # -----------------------------
    # 7) Build summary table by supervisor
    #     - MissingJobs: how many jobs missing DateAssigned
    #     - AvgDaysMissing: average age of missing jobs
    #     - OldestMissing: max age of missing jobs
    # -----------------------------
    summary = (
        df_missing.groupby("CutAtmosSupervisor")
        .agg(
            MissingJobs=("PavementKey", "count"),
            AvgDaysMissing=("DaysMissing", "mean"),
            OldestMissing=("DaysMissing", "max"),
        )
        .reset_index()
        .sort_values("MissingJobs", ascending=False)
    )

    # Optional: round AvgDaysMissing to 1 decimal
    summary["AvgDaysMissing"] = summary["AvgDaysMissing"].round(1)

    # -----------------------------
    # 8) Build detail table
    #     - PavementKey
    #     - Address
    #     - CutAtmosSupervisor
    #     - DateCut
    #     - DaysMissing
    # -----------------------------
    detail = df_missing[
        ["PavementKey", "Address", "CutAtmosSupervisor", "DateCut", "DaysMissing"]
    ].sort_values("DaysMissing", ascending=False)

    # -----------------------------
    # 9) Ask user where to save output Excel
    # -----------------------------
    messagebox.showinfo("Save File", "Choose where to save the output Excel report.")
    output_path = filedialog.asksaveasfilename(
        title="Save Output Excel File",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not output_path:
        messagebox.showerror("Cancelled", "No save location selected. Exiting.")
        return

    # -----------------------------
    # 10) Write results to Excel (openpyxl engine)
    # -----------------------------
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            summary.to_excel(writer, sheet_name="Supervisor_Missing_Count", index=False)
            detail.to_excel(writer, sheet_name="Missing_Job_Details", index=False)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save Excel file:\n{e}")
        return

    # -----------------------------
    # 11) Build HTML for email body (Atmos colors)
    # -----------------------------
    # Limit detail in email to top 20 oldest jobs so it doesn't get huge
    detail_email = detail.head(20).copy()

    # Convert DataFrames to basic HTML
    summary_html_raw = summary.to_html(index=False, border=0)
    detail_html_raw = detail_email.to_html(index=False, border=0)

    # Style the tables
    summary_html = style_html_table(summary_html_raw)
    detail_html = style_html_table(detail_html_raw)

    # Build the full email body
    email_body = f"""
<div style="font-family:Arial, sans-serif; font-size:13px; color:#333333;">

  <p>Below is this week's summary of paving jobs that are missing
  <strong>DateAssigned</strong>. These jobs <strong>cannot</strong> continue until supervisors complete their updates.</p>


  <h4 style="color:#0072CE; margin-bottom:4px;">Missing Jobs by Supervisor</h4>
  {summary_html}

  <h4 style="color:#0072CE; margin-top:16px; margin-bottom:4px;">Jobs Missing Data (Top 20 Oldest)</h4>
  {detail_html}

  Supervisors listed above should update <strong>DateAssigned</strong> in CM+ as soon as possible.</p>

</div>
"""

    # Save HTML next to the Excel file
    html_output_path = os.path.splitext(output_path)[0] + "_email.html"
    try:
        with open(html_output_path, "w", encoding="utf-8") as f:
            f.write(email_body)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save HTML email body:\n{e}")
        return

    # -----------------------------
    # 12) Done
    # -----------------------------
    messagebox.showinfo(
        "Success",
        f"Report generated successfully!\n\n"
        f"Excel report:\n{output_path}\n\n"
        f"Email HTML file:\n{html_output_path}"
    )


if __name__ == "__main__":
    main()
