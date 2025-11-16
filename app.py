from flask import Flask, send_file
from flask_cors import CORS
import openpyxl
from io import BytesIO

app = Flask(__name__)
CORS(app)

@app.route("/api/download-template", methods=["GET"])
def download_template():
    # Create a simple Excel workbook in memory
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Template"
    ws["A1"] = "Hello!"
    ws["A2"] = "This is your attendance template."

    # Save Excel to memory
    file_stream = BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name="Attendance_Template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/")
def home():
    return "Backend is running!"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
