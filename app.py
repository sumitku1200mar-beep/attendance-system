from flask import Flask, render_template, request, send_file
import pandas as pd

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        summary_file = request.files["summary"]
        punch_file = request.files["punch"]

        # Read Excel files
        summary_df = pd.read_excel(summary_file)
        punch_df = pd.read_excel(punch_file)

        # Punching se sirf ARCO ID (Column 1)
        present_ids = punch_df.iloc[:, 0].dropna().unique()

        # Summary me data row 4 se start
        summary_df = summary_df.iloc[3:].copy()

        summary_df.columns = range(summary_df.shape[1])

        # Important Columns (adjust if needed)
        arco_col = 1      # ARCO ID column in summary
        name_col = 2      # Name column
        desig_col = 3     # Designation column

        # Status assign
        summary_df["Status"] = summary_df[arco_col].apply(
            lambda x: "P" if x in present_ids else "A"
        )

        # Present & Absent Sheets
        present_df = summary_df[summary_df["Status"] == "P"]
        absent_df = summary_df[summary_df["Status"] == "A"]

        # Keep only required columns
        present_df = present_df[[arco_col, name_col, desig_col, "Status"]]
        absent_df = absent_df[[arco_col, name_col, desig_col, "Status"]]

        # Add Serial Number
        present_df.insert(0, "Sr No", range(1, len(present_df)+1))
        absent_df.insert(0, "Sr No", range(1, len(absent_df)+1))

        # Rename columns
        present_df.columns = ["Sr No", "ARCO ID", "Name", "Designation", "Status"]
        absent_df.columns = ["Sr No", "ARCO ID", "Name", "Designation", "Status"]

        # Manpower Summary
        manpower = summary_df.groupby([summary_df[desig_col], "Status"]).size().unstack(fill_value=0)
        manpower["Total"] = manpower.sum(axis=1)

        # Grand Total
        grand_present = manpower["P"].sum() if "P" in manpower.columns else 0
        grand_absent = manpower["A"].sum() if "A" in manpower.columns else 0
        grand_total = grand_present + grand_absent

        manpower.loc["SUB TOTAL"] = [grand_present, grand_absent, grand_total]

        output_path = "Attendance_Result.xlsx"

        with pd.ExcelWriter(output_path) as writer:
            present_df.to_excel(writer, sheet_name="Present", index=False)
            absent_df.to_excel(writer, sheet_name="Absent", index=False)
            manpower.to_excel(writer, sheet_name="Manpower Summary")

        return send_file(output_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run()
