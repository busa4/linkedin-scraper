from flask import Flask, jsonify, request, render_template
from pathlib import Path
import openpyxl

app = Flask(__name__)
TRACKER = Path(__file__).parent / "tracker.xlsx"

STATUSES = ["New", "Contacted", "Interested", "Not Interested", "Follow Up", "Closed Deal", "Manually Declined"]

FIELD_MAP = {
    'field':         'Field',
    'name':          'Name',
    'url':           'LinkedIn URL',
    'industry':      'Industry',
    'location':      'Location',
    'description':   'Description',
    'score':         'Automation Interest Score (1-10)',
    'needs':         'Potential Needs',
    'contact_role':  'Recommended Contact Role',
    'contact_person':'Contact Person',
    'status':        'Status',
}


def load_companies():
    wb = openpyxl.load_workbook(TRACKER)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    companies = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        r = dict(zip(headers, row))
        score_raw = r.get("Automation Interest Score (1-10)")
        try:
            score = int(score_raw) if score_raw not in (None, '', 'Not set') else 0
        except (ValueError, TypeError):
            score = 0
        companies.append({
            "row":          i,
            "id":           int(r.get("#") or i - 1),
            "field":        r.get("Field") or "",
            "name":         r.get("Name") or "",
            "url":          r.get("LinkedIn URL") or "",
            "industry":     r.get("Industry") or "",
            "location":     r.get("Location") or "",
            "description":  r.get("Description") or "",
            "score":        score,
            "needs":        r.get("Potential Needs") or "",
            "contact_role": r.get("Recommended Contact Role") or "",
            "contact_person": r.get("Contact Person") or "",
            "status":       r.get("Status") or "",
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
    for key, col_name in FIELD_MAP.items():
        if key in data:
            val = data[key]
            if key == 'score':
                try:
                    val = int(val) if val not in (None, '', 'Not set') else None
                except (ValueError, TypeError):
                    val = None
            save_field(row_num, col_name, val if val else None)
    return jsonify({"ok": True})


@app.route("/api/statuses")
def api_statuses():
    return jsonify(STATUSES)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
