from flask import Flask, render_template, request, send_file
import pandas as pd

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        summary_file = request.files["summary"]
        punch_file = request.files["punch"]

        summary_df = pd.read_excel(summary_file)
        punch_df = pd.read_excel(punch_file)

        present_ids = punch_df.iloc[:, 0].dropna().unique()
        summary_df = summary_df.iloc[3:]

        summary_df["Status"] = summary_df.iloc[:, 1].apply(
            lambda x: "Present" if x in present_ids else "Absent"
        )

        present_df = summary_df[summary_df["Status"] == "Present"]
        absent_df = summary_df[summary_df["Status"] == "Absent"]

        manpower = summary_df.groupby(
            [summary_df.iloc[:, 3], "Status"]
        ).size().unstack(fill_value=0)

        manpower["Total"] = manpower.sum(axis=1)

        output_path = "Attendance_Result.xlsx"

        with pd.ExcelWriter(output_path) as writer:
            present_df.to_excel(writer, sheet_name="Present", index=False)
            absent_df.to_excel(writer, sheet_name="Absent", index=False)
            manpower.to_excel(writer, sheet_name="Manpower Summary")

        return send_file(output_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run()
