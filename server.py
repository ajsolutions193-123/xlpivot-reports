"""
Backend for the Excel to XLPivot & Reporting System page.

Each report button on the HTML page sends a POST request to /api/run-report
with:
  - file:       the Excel/CSV file picked by the user
  - report_id:  a string identifying which button was clicked

This server routes each report_id to its own Python function below.
Replace the body of each function with your converted VBA logic.

LOCAL USE (just you, on your own computer):
    pip install -r requirements.txt
    python server.py
    -> open http://127.0.0.1:5000 in the browser

DEPLOYING SO OTHERS CAN USE IT (see the deployment notes at the bottom
of this file / the accompanying instructions) runs this instead with:
    gunicorn server:app
"""

from flask import Flask, request, jsonify, send_file, send_from_directory
import os
import uuid
import tempfile
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment

# Serves login_form.html / dashboard.html directly from this same app,
# so the whole thing is one deployable unit (one URL, no separate hosting
# for the HTML files, and no cross-origin/CORS headaches).
app = Flask(__name__, static_folder=".", static_url_path="")

BASE_TMP_DIR = tempfile.mkdtemp(prefix="xlpivot_")


# ---------------------------------------------------------------------------
# One function per report button.
# Each takes the path to the uploaded source file and must return EITHER:
#   - a dict, e.g. {"message": "Done, saved to X"}   -> shown as a status message
#   - a file path (string) to an Excel/PDF/CSV file  -> sent back for download
# ---------------------------------------------------------------------------


def salary_pivot_bu_wise(input_path):
    # TODO: put your converted VBA logic here.
    # Example shape of what you'll likely do:
    #   import pandas as pd
    #   df = pd.read_excel(input_path)
    #   pivot = df.pivot_table(index="BU", values="Salary", aggfunc="sum")
    #   out_path = os.path.join(OUTPUT_DIR, "Salary_Pivot_BU_Wise.xlsx")
    #   pivot.to_excel(out_path)
    #   return out_path
    return {"message": "salary_pivot_bu_wise: logic not implemented yet."}


def salary_pivot_bu_wise_pdf(input_path):
    # TODO: same as above, but export as PDF instead of Excel.
    return {"message": "salary_pivot_bu_wise_pdf: logic not implemented yet."}


def gross_salary_comparision(input_path):
    # TODO: convert your "Gross Salary Comparision (A-B)" VBA macro here.
    return {"message": "gross_salary_comparision: logic not implemented yet."}


def net_salary_comparision(input_path):
    # TODO: convert your "Net Salary Comparision (A-B)" VBA macro here.
    return {"message": "net_salary_comparision: logic not implemented yet."}


def tds_section_wise(input_path):
    # TODO: convert your "TDS Section wise" VBA macro here.
    return {"message": "tds_section_wise: logic not implemented yet."}


def tds_pivot_bu_wise(input_path):
    # TODO: convert your "TDS Pivot BU Wise" VBA macro here.
    return {"message": "tds_pivot_bu_wise: logic not implemented yet."}


def gst_pivot_summary(input_path):
    # TODO: convert your "GST Pivot Summary" VBA macro here.
    return {"message": "gst_pivot_summary: logic not implemented yet."}


def sheet_exists(name, wb):
    """Equivalent of the VBA SheetExists helper."""
    return name.lower() in [s.lower() for s in wb.sheetnames]


def amount_received(input_path):
    """
    Converted from VBA: CreatePivotFromSourceData_Adv

    Builds a static "pivot-style" summary table (openpyxl/pandas cannot
    create a native, refreshable Excel PivotTable object -- this
    reproduces the same grouping/sums/formatting as a plain table instead).
    """
    wb = load_workbook(input_path)

    # --- Find SourceData sheet -------------------------------------------------
    if "SourceData" not in wb.sheetnames:
        raise ValueError("Sheet 'SourceData' not found in the selected file.")
    ws_data = wb["SourceData"]

    # --- Dynamic range: header row fixed at 13, last row/col detected ----------
    header_row = 13

    last_row = header_row
    for r in range(ws_data.max_row, header_row - 1, -1):
        if ws_data.cell(row=r, column=1).value not in (None, ""):
            last_row = r
            break

    last_col = 1
    for c in range(ws_data.max_column, 0, -1):
        if ws_data.cell(row=header_row, column=c).value not in (None, ""):
            last_col = c
            break

    # --- Read header + data rows into a DataFrame -------------------------------
    headers = [ws_data.cell(row=header_row, column=c).value for c in range(1, last_col + 1)]

    data_rows = []
    for r in range(header_row + 1, last_row + 1):
        data_rows.append([ws_data.cell(row=r, column=c).value for c in range(1, last_col + 1)])

    df = pd.DataFrame(data_rows, columns=headers)

    required_cols = [
        "Status",
        "TOTAL REC TILL DATE BASIC",
        "On Account Amount",
        "TOTAL REC TILL DATE BASIC2",
        "On Account Amount2",
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(
            "Missing expected column(s) in SourceData: " + ", ".join(missing)
        )

    value_cols = [
        "TOTAL REC TILL DATE BASIC",
        "On Account Amount",
        "TOTAL REC TILL DATE BASIC2",
        "On Account Amount2",
    ]
    for col in value_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # --- Group by Status, sum the value columns (the actual "pivot") -----------
    pivot = df.groupby("Status", dropna=False)[value_cols].sum()

    # --- New sheet name: Sheet1, Sheet2, Sheet3 ... -----------------------------
    i = 1
    while sheet_exists(f"Sheet{i}", wb):
        i += 1
    new_sheet_name = f"Sheet{i}"
    ws_pivot = wb.create_sheet(new_sheet_name)

    # --- Write the summary table starting at A3 (like TableDestination A3) -----
    start_row = 3
    start_col = 1  # column A

    header_labels = ["Row Labels"] + [f"Sum of {c}" for c in value_cols]
    for j, label in enumerate(header_labels):
        ws_pivot.cell(row=start_row, column=start_col + j, value=label)

    data_start_row = start_row + 1
    row_idx = data_start_row
    for status_val, row_data in pivot.iterrows():
        ws_pivot.cell(row=row_idx, column=start_col, value=status_val)
        for j, col in enumerate(value_cols):
            ws_pivot.cell(row=row_idx, column=start_col + 1 + j, value=float(row_data[col]))
        row_idx += 1

    grand_total_row = row_idx
    ws_pivot.cell(row=grand_total_row, column=start_col, value="Grand Total")
    for j, col in enumerate(value_cols):
        col_letter = get_column_letter(start_col + 1 + j)
        ws_pivot.cell(
            row=grand_total_row,
            column=start_col + 1 + j,
            value=f"=SUM({col_letter}{data_start_row}:{col_letter}{grand_total_row - 1})",
        )

    # --- G Total column (row-wise sum), like F3/F4/F5/F6 in the VBA ------------
    g_total_col = start_col + 1 + len(value_cols)  # column F when there are 4 value cols
    g_total_col_letter = get_column_letter(g_total_col)
    ws_pivot.cell(row=start_row, column=g_total_col, value="G Total")

    first_value_col_letter = get_column_letter(start_col + 1)
    last_value_col_letter = get_column_letter(start_col + len(value_cols))

    for r in range(data_start_row, grand_total_row):
        ws_pivot.cell(
            row=r,
            column=g_total_col,
            value=f"=SUM({first_value_col_letter}{r}:{last_value_col_letter}{r})",
        )

    ws_pivot.cell(
        row=grand_total_row,
        column=g_total_col,
        value=f"=SUM({g_total_col_letter}{data_start_row}:{g_total_col_letter}{grand_total_row - 1})",
    )

    # --- Borders + fill on the G Total column (matches the VBA formatting) -----
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")

    for r in range(start_row, grand_total_row + 1):
        cell = ws_pivot.cell(row=r, column=g_total_col)
        cell.border = border
        cell.fill = fill

    # --- Save back into the same workbook (like wbSource.Save) ------------------
    wb.save(input_path)

    return input_path


def daybook_difference(input_path):
    # TODO: convert your "Difference of Day Book A-B" VBA macro here.
    return {"message": "daybook_difference: logic not implemented yet."}


def creditor_outstanding(input_path):
    """
    Converted from VBA: GenerateCreditorReport

    Builds a two-section report on a new sheet:
      Section 1: "Outstanding Bills Amount" -> rows where Bill Outstanding <> 0
      Section 2: "Advance to Creditors"      -> rows where On A/C <> 0
    Each section gets a title row, a bold/shaded header row, its filtered
    data rows with a light grid, and a bold/shaded Total row.
    """

    # --- Same constants as the top of the VBA module ---------------------------
    HEADER_ROW = 6
    HDR_NO_OF_BILLS = "No. Of Bills"
    HDR_BILL_OUTSTANDING = "Bill Outstanding"
    HDR_ON_AC = "On A/C"
    HDR_NET_OUTSTANDING = "Net Outstanding"
    GAP_ROWS = 2

    wb = load_workbook(input_path)

    if "SourceData" not in wb.sheetnames:
        raise ValueError("Sheet 'SourceData' was not found in the selected file.")
    src_ws = wb["SourceData"]

    # --- last column (from header row) / last row (from column A) --------------
    last_col = 1
    for c in range(src_ws.max_column, 0, -1):
        if src_ws.cell(row=HEADER_ROW, column=c).value not in (None, ""):
            last_col = c
            break

    last_row = HEADER_ROW
    for r in range(src_ws.max_row, 0, -1):
        if src_ws.cell(row=r, column=1).value not in (None, ""):
            last_row = r
            break

    if last_row <= HEADER_ROW:
        raise ValueError("No data found below the headers in SourceData.")

    # --- find the needed columns by header text ---------------------------------
    col_no_of_bills = find_header_column(src_ws, HEADER_ROW, last_col, HDR_NO_OF_BILLS)
    col_bill_outstanding = find_header_column(src_ws, HEADER_ROW, last_col, HDR_BILL_OUTSTANDING)
    col_on_ac = find_header_column(src_ws, HEADER_ROW, last_col, HDR_ON_AC)
    col_net_outstanding = find_header_column(src_ws, HEADER_ROW, last_col, HDR_NET_OUTSTANDING)

    if col_bill_outstanding == 0 or col_on_ac == 0 or col_net_outstanding == 0:
        raise ValueError(
            f"Could not find '{HDR_BILL_OUTSTANDING}', '{HDR_ON_AC}' or '{HDR_NET_OUTSTANDING}' "
            f"headers in row {HEADER_ROW}. Please check the header name constants."
        )

    # --- columns to copy: every column except "No. Of Bills" --------------------
    src_cols = build_column_list(last_col, col_no_of_bills)

    # --- new sheet, placed right after SourceData --------------------------------
    dest_ws = add_next_available_sheet(wb, "SourceData")

    # --- timestamp -----------------------------------------------------------------
    dest_ws.cell(row=1, column=2, value="Report Generated: " + datetime.now().strftime("%d-%b-%Y %I:%M:%S %p"))
    dest_ws.cell(row=1, column=2).font = Font(bold=True, italic=True)

    section_start_row = 3  # leave row 2 blank, same as the VBA

    # --- Section 1: Outstanding Bills Amount ------------------------------------
    next_start_row = write_section(
        dest_ws, src_ws, HEADER_ROW, last_row, src_cols,
        col_bill_outstanding, "Outstanding Bills Amount", section_start_row,
        col_bill_outstanding, col_on_ac, col_net_outstanding,
    )

    # --- Section 2: Advance to Creditors ----------------------------------------
    next_start_row = write_section(
        dest_ws, src_ws, HEADER_ROW, last_row, src_cols,
        col_on_ac, "Advance to Creditors", next_start_row + GAP_ROWS,
        col_bill_outstanding, col_on_ac, col_net_outstanding,
    )

    # --- reasonable column widths (openpyxl has no true "AutoFit") -------------
    for col_idx in range(2, len(src_cols) + 2):
        dest_ws.column_dimensions[get_column_letter(col_idx)].width = 16

    wb.save(input_path)
    return input_path


def find_header_column(ws, header_row, last_col, header_text):
    """Equivalent of VBA FindHeaderColumn: case/whitespace-insensitive header match."""
    target = header_text.strip().lower()
    for c in range(1, last_col + 1):
        val = ws.cell(row=header_row, column=c).value
        if val is not None and str(val).strip().lower() == target:
            return c
    return 0


def build_column_list(last_col, exclude_col):
    """Equivalent of VBA BuildColumnList: all source columns except one."""
    return [c for c in range(1, last_col + 1) if c != exclude_col]


def position_in_list(col_list, src_col_idx):
    """Equivalent of VBA PositionInList: 1-based position of a source column in col_list."""
    for i, c in enumerate(col_list, start=1):
        if c == src_col_idx:
            return i
    return 0


def add_next_available_sheet(wb, after_sheet_name):
    """Equivalent of VBA AddNextAvailableSheet: first unused 'SheetN' name, placed after a given sheet."""
    n = 1
    while f"Sheet{n}" in wb.sheetnames:
        n += 1
    new_name = f"Sheet{n}"
    idx = wb.sheetnames.index(after_sheet_name) + 1
    return wb.create_sheet(new_name, idx)


def write_section(dest_ws, src_ws, header_row, last_row, col_list, filter_src_col,
                   title_text, start_row, src_col_bill_outstanding, src_col_on_ac,
                   src_col_net_outstanding):
    """Equivalent of VBA WriteSection: title row + header row + filtered data rows + total row."""

    thin = Side(style="thin")
    thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    medium = Side(style="medium")

    title_fill = PatternFill(start_color="FBE0CE", end_color="FBE0CE", fill_type="solid")   # ~accent2, tint .6
    header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")  # ~accent1, tint .6
    total_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")   # ~accent1, tint .8

    last_dest_col = len(col_list)
    end_col = 1 + last_dest_col  # data occupies columns B..end_col

    pos_bill = position_in_list(col_list, src_col_bill_outstanding)
    pos_onac = position_in_list(col_list, src_col_on_ac)
    pos_net = position_in_list(col_list, src_col_net_outstanding)

    # --- Title row (merged, centered, shaded) -----------------------------------
    dest_ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=end_col)
    title_cell = dest_ws.cell(row=start_row, column=2, value=title_text)
    title_cell.alignment = Alignment(horizontal="center", vertical="bottom")
    for col in range(2, end_col + 1):
        cell = dest_ws.cell(row=start_row, column=col)
        cell.fill = title_fill
        cell.border = thin_border

    # --- Header row ---------------------------------------------------------------
    header_dest_row = start_row + 1
    dest_col = 2
    for src_c in col_list:
        dest_ws.cell(row=header_dest_row, column=dest_col, value=src_ws.cell(row=header_row, column=src_c).value)
        dest_col += 1
    for col in range(2, end_col + 1):
        cell = dest_ws.cell(row=header_dest_row, column=col)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

    # --- Data rows: only where filter_src_col value is numeric and non-zero ----
    dest_row = start_row + 2
    first_data_row = dest_row
    for r in range(header_row + 1, last_row + 1):
        val = src_ws.cell(row=r, column=filter_src_col).value
        if isinstance(val, (int, float)) and val != 0:
            dest_col = 2
            for src_c in col_list:
                dest_ws.cell(row=dest_row, column=dest_col, value=src_ws.cell(row=r, column=src_c).value)
                dest_col += 1
            dest_row += 1
    last_data_row = dest_row - 1

    if last_data_row >= first_data_row:
        for r in range(first_data_row, last_data_row + 1):
            for col in range(2, end_col + 1):
                dest_ws.cell(row=r, column=col).border = thin_border

    # --- Total row -----------------------------------------------------------------
    if last_data_row >= first_data_row:
        total_label_cell = dest_ws.cell(row=dest_row, column=2, value="Total")
        total_label_cell.font = Font(bold=True)

        for pos in (pos_bill, pos_onac, pos_net):
            if pos > 0:
                col_num = pos + 1
                letter = get_column_letter(col_num)
                cell = dest_ws.cell(row=dest_row, column=col_num)
                cell.value = f"=SUM({letter}{first_data_row}:{letter}{last_data_row})"
                cell.font = Font(bold=True)

        for col in range(2, end_col + 1):
            cell = dest_ws.cell(row=dest_row, column=col)
            cell.fill = total_fill
            cell.border = Border(left=thin, right=thin, top=medium, bottom=thin)

    return dest_row


# Maps the report_id sent from the HTML page to the function that handles it.
REPORT_HANDLERS = {
    "salary_pivot_bu_wise": salary_pivot_bu_wise,
    "salary_pivot_bu_wise_pdf": salary_pivot_bu_wise_pdf,
    "gross_salary_comparision": gross_salary_comparision,
    "net_salary_comparision": net_salary_comparision,
    "tds_section_wise": tds_section_wise,
    "tds_pivot_bu_wise": tds_pivot_bu_wise,
    "gst_pivot_summary": gst_pivot_summary,
    "amount_received": amount_received,
    "daybook_difference": daybook_difference,
    "creditor_outstanding": creditor_outstanding,
}


@app.route("/")
def serve_login():
    return send_from_directory(".", "login_form.html")


@app.route("/dashboard.html")
def serve_dashboard():
    return send_from_directory(".", "dashboard.html")


@app.route("/api/run-report", methods=["POST"])
def run_report():
    report_id = request.form.get("report_id")
    uploaded_file = request.files.get("file")

    if not report_id or report_id not in REPORT_HANDLERS:
        return jsonify({"message": f"Unknown report_id: {report_id}"}), 400

    if not uploaded_file:
        return jsonify({"message": "No file was uploaded."}), 400

    # Each request gets its own private folder, so two people uploading a
    # file with the same name at the same time never collide or overwrite
    # each other's data.
    request_dir = os.path.join(BASE_TMP_DIR, uuid.uuid4().hex)
    os.makedirs(request_dir, exist_ok=True)

    input_path = os.path.join(request_dir, uploaded_file.filename)
    uploaded_file.save(input_path)

    handler = REPORT_HANDLERS[report_id]

    try:
        result = handler(input_path)
    except Exception as exc:
        return jsonify({"message": f"Error while running {report_id}: {exc}"}), 500

    # If the handler returned a file path, send that file back for download.
    if isinstance(result, str) and os.path.isfile(result):
        return send_file(result, as_attachment=True)

    # Otherwise treat it as a JSON status message.
    return jsonify(result)


if __name__ == "__main__":
    # For local testing only. When deployed, a production server
    # (gunicorn) runs this app instead -- see deployment notes.
    app.run(host="127.0.0.1", port=5000, debug=True)
