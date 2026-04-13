import streamlit as st
from datetime import date, datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

st.set_page_config(page_title="仮払申請書メーカー", page_icon="💴", layout="wide")


# ---------- Helpers ----------
def fmt_date(value):
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y/%m/%d")
    if isinstance(value, date):
        return value.strftime("%Y/%m/%d")
    return str(value)


def fmt_currency(value):
    try:
        if value is None or value == "":
            return ""
        return f"¥{float(value):,.0f}"
    except Exception:
        return str(value)


def safe_text(value):
    return "" if value is None else str(value)


def build_workbook(data: dict) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "仮払申請書"

    # ---------- Theme ----------
    dark = "1F4E78"
    mid = "D9E7F5"
    light = "F7FAFC"
    note_fill = "FFF4CC"
    white = "FFFFFF"
    black = "222222"

    thin_gray = Side(style="thin", color="9AA4B2")
    medium_dark = Side(style="medium", color="4B5563")

    title_font = Font(name="Meiryo", size=16, bold=True, color=black)
    section_font = Font(name="Meiryo", size=11, bold=True, color=white)
    label_font = Font(name="Meiryo", size=10, bold=True, color=black)
    value_font = Font(name="Meiryo", size=10, color=black)
    note_font = Font(name="Meiryo", size=9, italic=True, color=black)
    sign_font = Font(name="Meiryo", size=9, bold=True, color=black)
    mini_font = Font(name="Meiryo", size=8, color="666666")

    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    top_left_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)
    center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def set_range_border(cell_range, border):
        rows = ws[cell_range]
        for row in rows:
            for cell in row:
                cell.border = border

    def merge_value(label_row, label_col, label, value, value_end_col, height=22):
        ws.cell(label_row, label_col, label)
        ws.cell(label_row, label_col).font = label_font
        ws.cell(label_row, label_col).alignment = left
        ws.cell(label_row, label_col).fill = PatternFill("solid", fgColor=light)
        ws.merge_cells(start_row=label_row, start_column=label_col + 1, end_row=label_row, end_column=value_end_col)
        vcell = ws.cell(label_row, label_col + 1, safe_text(value))
        vcell.font = value_font
        vcell.alignment = left
        ws.row_dimensions[label_row].height = height
        for col in range(label_col, value_end_col + 1):
            ws.cell(label_row, col).border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    # ---------- Layout ----------
    widths = {
        "A": 14, "B": 14, "C": 14, "D": 14,
        "E": 14, "F": 14, "G": 14, "H": 14,
        "I": 14, "J": 14, "K": 14, "L": 14,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A1"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.45, bottom=0.45, header=0.2, footer=0.2)
    ws.print_title_rows = "$1:$3"
    ws.print_area = "A1:L42"

    # ---------- Title ----------
    ws.merge_cells("A1:J2")
    ws["A1"] = "仮払申請書"
    ws["A1"].font = title_font
    ws["A1"].alignment = center

    ws.merge_cells("K1:L2")
    ws["K1"] = f"管理番号\n{safe_text(data['control_no'])}"
    ws["K1"].font = label_font
    ws["K1"].alignment = center_wrap
    ws["K1"].fill = PatternFill("solid", fgColor=mid)
    set_range_border("K1:L2", Border(left=medium_dark, right=medium_dark, top=medium_dark, bottom=medium_dark))
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 22

    def section_header(row, title):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
        c = ws.cell(row, 1, title)
        c.font = section_font
        c.fill = PatternFill("solid", fgColor=dark)
        c.alignment = left
        set_range_border(f"A{row}:L{row}", Border(left=medium_dark, right=medium_dark, top=medium_dark, bottom=medium_dark))
        ws.row_dimensions[row].height = 22

    def signature_block(start_row, role1, role2, role3):
        labels = ["Prepared by", "Confirmed by", "Approved by"]
        roles = [role1, role2, role3]
        starts = [1, 5, 9]
        for idx, start_col in enumerate(starts):
            ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + 3)
            ws.cell(start_row, start_col, labels[idx])
            ws.cell(start_row, start_col).font = mini_font
            ws.cell(start_row, start_col).alignment = left

            ws.merge_cells(start_row=start_row + 1, start_column=start_col, end_row=start_row + 2, end_column=start_col + 3)
            cell = ws.cell(start_row + 1, start_col, roles[idx])
            cell.font = sign_font
            cell.alignment = center
            rng = f"{get_column_letter(start_col)}{start_row+1}:{get_column_letter(start_col+3)}{start_row+2}"
            set_range_border(rng, Border(top=medium_dark))
        ws.row_dimensions[start_row].height = 18
        ws.row_dimensions[start_row + 1].height = 20
        ws.row_dimensions[start_row + 2].height = 18

    # ---------- Section 1 ----------
    section_header(4, "1. 申請")
    merge_value(5, 1, "申請者", data["applicant"], 4)
    merge_value(5, 5, "所属", data["department"], 8)
    merge_value(5, 9, "オフィス", data["office"], 12)

    merge_value(6, 1, "申請日", fmt_date(data["application_date"]), 4)
    merge_value(6, 5, "借入日", fmt_date(data["borrow_date"]), 8)
    merge_value(6, 9, "精算予定日", fmt_date(data["planned_settlement_date"]), 12)

    merge_value(7, 1, "申請金額", fmt_currency(data["application_amount"]), 12)

    ws.cell(8, 1, "申請理由")
    ws.cell(8, 1).font = label_font
    ws.cell(8, 1).alignment = left
    ws.cell(8, 1).fill = PatternFill("solid", fgColor=light)
    ws.merge_cells("B8:L10")
    ws["B8"] = safe_text(data["application_reason"])
    ws["B8"].font = value_font
    ws["B8"].alignment = top_left_wrap
    set_range_border("A8:L10", Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray))
    ws.row_dimensions[8].height = 24
    ws.row_dimensions[9].height = 24
    ws.row_dimensions[10].height = 24

    ws.merge_cells("A11:L11")
    ws["A11"] = "※ 清算予定日を過ぎた場合、次回給与から天引きされることに同意します。"
    ws["A11"].font = note_font
    ws["A11"].alignment = left
    ws["A11"].fill = PatternFill("solid", fgColor=note_fill)
    set_range_border("A11:L11", Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray))
    signature_block(12, "仮受者", "上長", "Casher")

    # ---------- Section 2 ----------
    section_header(16, "2. 精算書")
    merge_value(17, 1, "清算日", fmt_date(data["settlement_date"]), 12)
    merge_value(18, 1, "仮払金額", fmt_currency(data["cash_advance_amount"]), 4)
    merge_value(18, 5, "領収書合計", fmt_currency(data["receipt_total"]), 8)
    merge_value(18, 9, "お釣り", fmt_currency(data["change_amount"]), 12)
    merge_value(19, 1, "未清算金", fmt_currency(data["unsettled_amount"]), 12)

    ws.merge_cells("A20:L20")
    ws["A20"] = "※ 上記の通り精算したことを証明します。"
    ws["A20"].font = note_font
    ws["A20"].alignment = left
    ws["A20"].fill = PatternFill("solid", fgColor=note_fill)
    set_range_border("A20:L20", Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray))
    signature_block(21, "仮受者", "Casher", "CFO")

    # ---------- Section 3 ----------
    section_header(25, "3. 返済書")
    merge_value(26, 1, "返済者", data["repayer"], 4)
    merge_value(26, 5, "所属", data["repayer_department"], 8)
    merge_value(26, 9, "オフィス", data["repayer_office"], 10)
    ws.cell(26, 11, "記入日")
    ws.cell(26, 11).font = label_font
    ws.cell(26, 11).fill = PatternFill("solid", fgColor=light)
    ws.cell(26, 11).alignment = left
    ws.cell(26, 12, fmt_date(data["repayment_entry_date"]))
    ws.cell(26, 12).font = value_font
    ws.cell(26, 12).alignment = left
    set_range_border("A26:L26", Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray))

    merge_value(27, 1, "返済額合計", fmt_currency(data["repayment_total"]), 12)
    merge_value(28, 1, "返済方法", data["repayment_method"], 12)
    merge_value(29, 1, "分割回数", data["installments"], 12)
    merge_value(30, 1, "初回引落日", fmt_date(data["first_deduction_date"]), 12)

    ws.merge_cells("A31:L31")
    ws["A31"] = "※ 上記の通り返済することを誓います。"
    ws["A31"].font = note_font
    ws["A31"].alignment = left
    ws["A31"].fill = PatternFill("solid", fgColor=note_fill)
    set_range_border("A31:L31", Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray))
    signature_block(32, "仮受者", "Casher", "CFO")

    # Footer memo area
    ws.merge_cells("A36:L37")
    ws["A36"] = "備考"
    ws["A36"].font = label_font
    ws["A36"].alignment = top_left_wrap
    ws["A36"].fill = PatternFill("solid", fgColor=light)
    set_range_border("A36:L37", Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray))

    # Outer frame
    set_range_border("A4:L34", Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray))

    # Keep everything readable as one sheet
    for row in range(38, 43):
        ws.row_dimensions[row].height = 6

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


def default_state():
    today = date.today()
    return {
        "control_no": f"CA-{today.strftime('%Y%m%d')}-001",
        "applicant": "",
        "department": "",
        "office": "",
        "application_date": today,
        "borrow_date": today,
        "planned_settlement_date": today,
        "application_amount": 0,
        "application_reason": "",
        "settlement_date": today,
        "cash_advance_amount": 0,
        "receipt_total": 0,
        "change_amount": 0,
        "repayer": "",
        "repayer_department": "",
        "repayer_office": "",
        "repayment_entry_date": today,
        "repayment_total": 0,
        "repayment_method": "給与天引き",
        "installments": 1,
        "first_deduction_date": today,
    }


if "form_data" not in st.session_state:
    st.session_state.form_data = default_state()


def set_sample_data():
    today = date.today()
    st.session_state.form_data = {
        "control_no": f"CA-{today.strftime('%Y%m%d')}-001",
        "applicant": "山田 太郎",
        "department": "営業部",
        "office": "鹿児島",
        "application_date": today,
        "borrow_date": today,
        "planned_settlement_date": today,
        "application_amount": 30000,
        "application_reason": "出張に伴う交通費・宿泊費・立替経費の仮払い申請。",
        "settlement_date": today,
        "cash_advance_amount": 30000,
        "receipt_total": 24500,
        "change_amount": 3000,
        "repayer": "山田 太郎",
        "repayer_department": "営業部",
        "repayer_office": "鹿児島",
        "repayment_entry_date": today,
        "repayment_total": 2500,
        "repayment_method": "給与天引き",
        "installments": 1,
        "first_deduction_date": today,
    }


# ---------- UI ----------
st.title("💴 仮払申請書メーカー")
st.caption("申請・精算・返済を1画面で入力し、Excelで1シートにまとめて出力します。")

col_a, col_b = st.columns([3, 1])
with col_b:
    if st.button("サンプル入力", use_container_width=True):
        set_sample_data()
        st.rerun()

with st.form("cash_advance_form"):
    st.subheader("基本情報")
    c1, c2 = st.columns(2)
    with c1:
        control_no = st.text_input("管理番号", value=st.session_state.form_data["control_no"])
    with c2:
        applicant = st.text_input("申請者", value=st.session_state.form_data["applicant"])

    c3, c4 = st.columns(2)
    with c3:
        department = st.text_input("所属", value=st.session_state.form_data["department"])
    with c4:
        office = st.text_input("オフィス", value=st.session_state.form_data["office"])

    st.subheader("1. 申請")
    d1, d2, d3 = st.columns(3)
    with d1:
        application_date = st.date_input("申請日", value=st.session_state.form_data["application_date"])
    with d2:
        borrow_date = st.date_input("借入日", value=st.session_state.form_data["borrow_date"])
    with d3:
        planned_settlement_date = st.date_input("精算予定日", value=st.session_state.form_data["planned_settlement_date"])

    application_amount = st.number_input(
        "申請金額",
        min_value=0,
        step=1000,
        value=int(st.session_state.form_data["application_amount"]),
    )
    application_reason = st.text_area(
        "申請理由",
        value=st.session_state.form_data["application_reason"],
        height=100,
        placeholder="例：出張費、物品購入、現場対応の立替経費など",
    )

    st.subheader("2. 精算書")
    settlement_date = st.date_input("清算日", value=st.session_state.form_data["settlement_date"])
    e1, e2, e3 = st.columns(3)
    with e1:
        cash_advance_amount = st.number_input(
            "仮払金額", min_value=0, step=1000, value=int(st.session_state.form_data["cash_advance_amount"])
        )
    with e2:
        receipt_total = st.number_input(
            "領収書合計", min_value=0, step=1000, value=int(st.session_state.form_data["receipt_total"])
        )
    with e3:
        change_amount = st.number_input(
            "お釣り", min_value=0, step=1000, value=int(st.session_state.form_data["change_amount"])
        )

    unsettled_amount = cash_advance_amount - receipt_total - change_amount
    if unsettled_amount > 0:
        st.warning(f"未清算金：{fmt_currency(unsettled_amount)}")
    elif unsettled_amount < 0:
        st.info(f"追加精算が必要です：{fmt_currency(abs(unsettled_amount))}")
    else:
        st.success("未清算金はありません。")

    st.subheader("3. 返済書")
    f1, f2, f3 = st.columns(3)
    with f1:
        repayer = st.text_input("返済者", value=st.session_state.form_data["repayer"])
    with f2:
        repayer_department = st.text_input("返済者の所属", value=st.session_state.form_data["repayer_department"])
    with f3:
        repayer_office = st.text_input("返済者のオフィス", value=st.session_state.form_data["repayer_office"])

    g1, g2, g3, g4 = st.columns(4)
    with g1:
        repayment_entry_date = st.date_input("返済書の記入日", value=st.session_state.form_data["repayment_entry_date"])
    with g2:
        repayment_total = st.number_input(
            "返済額合計", min_value=0, step=1000, value=int(st.session_state.form_data["repayment_total"])
        )
    with g3:
        repayment_method = st.selectbox(
            "返済方法",
            ["コープ", "給与天引き"],
            index=["コープ", "給与天引き"].index(st.session_state.form_data["repayment_method"]),
        )
    with g4:
        installments = st.number_input(
            "分割回数", min_value=1, step=1, value=int(st.session_state.form_data["installments"])
        )

    first_deduction_date = st.date_input(
        "初回引落日",
        value=st.session_state.form_data["first_deduction_date"],
    )

    submitted = st.form_submit_button("入力内容を更新", use_container_width=True)

if submitted:
    st.session_state.form_data = {
        "control_no": control_no,
        "applicant": applicant,
        "department": department,
        "office": office,
        "application_date": application_date,
        "borrow_date": borrow_date,
        "planned_settlement_date": planned_settlement_date,
        "application_amount": application_amount,
        "application_reason": application_reason,
        "settlement_date": settlement_date,
        "cash_advance_amount": cash_advance_amount,
        "receipt_total": receipt_total,
        "change_amount": change_amount,
        "unsettled_amount": unsettled_amount,
        "repayer": repayer,
        "repayer_department": repayer_department,
        "repayer_office": repayer_office,
        "repayment_entry_date": repayment_entry_date,
        "repayment_total": repayment_total,
        "repayment_method": repayment_method,
        "installments": installments,
        "first_deduction_date": first_deduction_date,
    }
    st.success("入力内容を更新しました。")

# Ensure computed field exists before preview/download
preview_data = st.session_state.form_data.copy()
preview_data["unsettled_amount"] = (
    preview_data.get("cash_advance_amount", 0)
    - preview_data.get("receipt_total", 0)
    - preview_data.get("change_amount", 0)
)

st.divider()
st.subheader("プレビュー")

p1, p2 = st.columns(2)
with p1:
    st.markdown("#### 申請")
    st.write(f"**管理番号**: {preview_data['control_no']}")
    st.write(f"**申請者**: {preview_data['applicant']}")
    st.write(f"**所属 / オフィス**: {preview_data['department']} / {preview_data['office']}")
    st.write(f"**申請日**: {fmt_date(preview_data['application_date'])}")
    st.write(f"**借入日**: {fmt_date(preview_data['borrow_date'])}")
    st.write(f"**精算予定日**: {fmt_date(preview_data['planned_settlement_date'])}")
    st.write(f"**申請金額**: {fmt_currency(preview_data['application_amount'])}")
    st.write(f"**申請理由**: {preview_data['application_reason']}" or "-")

with p2:
    st.markdown("#### 精算 / 返済")
    st.write(f"**清算日**: {fmt_date(preview_data['settlement_date'])}")
    st.write(f"**仮払金額**: {fmt_currency(preview_data['cash_advance_amount'])}")
    st.write(f"**領収書合計**: {fmt_currency(preview_data['receipt_total'])}")
    st.write(f"**お釣り**: {fmt_currency(preview_data['change_amount'])}")
    st.write(f"**未清算金**: {fmt_currency(preview_data['unsettled_amount'])}")
    st.write(f"**返済者**: {preview_data['repayer']}")
    st.write(f"**返済方法**: {preview_data['repayment_method']}")
    st.write(f"**返済額合計**: {fmt_currency(preview_data['repayment_total'])}")
    st.write(f"**分割回数**: {preview_data['installments']} 回")
    st.write(f"**初回引落日**: {fmt_date(preview_data['first_deduction_date'])}")

excel_bytes = build_workbook(preview_data)
st.download_button(
    label="📥 Excelをダウンロード（1シート版）",
    data=excel_bytes,
    file_name=f"cash_advance_{preview_data['control_no'] or 'form'}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

st.info("出力ファイルはシート分割なしで、A4印刷を意識した1シート構成です。")
