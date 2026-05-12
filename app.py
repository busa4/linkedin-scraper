from flask import Flask, jsonify, request, render_template
from pathlib import Path
import openpyxl

app = Flask(__name__)
TRACKER = Path(__file__).parent / "tracker.xlsx"

STATUSES = ["New", "Contacted", "Interested", "Not Interested", "Follow Up", "Closed Deal"]


def load_companies():
    wb = openpyxl.load_workbook(TRACKER)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    companies = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        r = dict(zip(headers, row))
        companies.append({
            "row": i,
            "id": int(r.get("#") or i - 1),
            "field": r.get("Field") or "",
            "name": r.get("Name") or "",
            "url": r.get("LinkedIn URL") or "",
            "industry": r.get("Industry") or "",
            "location": r.get("Location") or "",
            "description": r.get("Description") or "",
            "score": int(r.get("Automation Interest Score (1-10)") or 0),
            "needs": r.get("Potential Needs") or "",
            "contact_role": r.get("Recommended Contact Role") or "",
            "contact_person": r.get("Contact Person") or "",
            "status": r.get("Status") or "",
        })
    return companies


def save_field(row_num: int, col_name: str, value):
    wb = openpyxl.load_workbook(TRACKER)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    col_idx = headers.index(col_name) + 1
    ws.cell(row=row_num, column=col_idx, value=value)
    wb.save(TRACKER)


@app.route("/")
def index():
    return render_template("index.html", statuses=STATUSES)


@app.route("/api/companies")
def api_companies():
    return jsonify(load_companies())


@app.route("/api/companies/<int:row_num>", methods=["PATCH"])
def api_update(row_num: int):
    data = request.json
    allowed = {"Status": "status", "Contact Person": "contact_person"}
    for col, key in allowed.items():
        if key in data:
            save_field(row_num, col, data[key] or None)
    return jsonify({"ok": True})


@app.route("/api/statuses")
def api_statuses():
    return jsonify(STATUSES)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
