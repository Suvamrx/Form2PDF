import io
from datetime import date
from typing import Any

import streamlit as st
from num2words import num2words
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# --- Page Configuration ---
st.set_page_config(page_title="Bank Transfer Letter PDF Generator", page_icon="PDF", layout="centered")

# Place beneficiary count outside the form so changing it reruns and updates form widgets
beneficiary_count = st.number_input("How many accounts?", min_value=1, max_value=10, value=1, step=1)

# --- Helper Functions ---
def amount_to_words(amount: float) -> str:
    """Convert amount to Indian rupees in words."""
    rupees = int(amount)
    paise = round((amount - rupees) * 100)
    
    rupees_words = num2words(rupees, lang='en_IN').capitalize() if rupees > 0 else "Zero"
    
    if paise > 0:
        paise_words = num2words(paise, lang='en_IN').lower()
        return f"Rupees {rupees_words} and {paise_words} paise only"
    else:
        return f"Rupees {rupees_words} only"


def build_letter_data(form_values: dict[str, Any]) -> dict[str, Any]:
    beneficiaries = []
    for index in range(form_values["beneficiary_count"]):
        beneficiaries.append(
            {
                "name": form_values[f"beneficiary_name_{index}"].strip(),
                "bank_branch": form_values[f"bank_branch_{index}"].strip(),
                "account_number": form_values[f"account_number_{index}"].strip(),
                "ifsc": form_values[f"ifsc_{index}"].strip(),
                "amount": float(form_values[f"amount_{index}"]),
            }
        )

    return {
        "bank_name": form_values["bank_name"].strip(),
        "branch_name": form_values["branch_name"].strip(),
        "date": form_values["date"],
        "sender_name": form_values["sender_name"].strip(),
        "sender_designation": form_values.get("sender_designation", "").strip(),
        "sender_phone": form_values.get("sender_phone", "").strip(),
        "sender_address": form_values["sender_address"].strip(),
        "program_account_name": form_values.get("program_account_name", "").strip(),
        "program_account_number": form_values.get("program_account_number", "").strip(),
        "cheque_number": form_values["cheque_number"].strip(),
        "reference": form_values["reference"].strip(),
        "beneficiaries": beneficiaries,
    }

# --- Core PDF Generation (with precise left-alignment) ---
def generate_pdf(data: dict[str, Any]) -> bytes:
    buffer = io.BytesIO()
    # Printable width: 210mm (A4) - 18mm (Left) - 18mm (Right) = 174mm
    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
    )

    styles = getSampleStyleSheet()
    body_style = ParagraphStyle(
        "BodyStyle",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=11,
        leading=16,
        textColor=colors.HexColor("#111827"),
    )
    label_style = ParagraphStyle(
        "LabelStyle",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=16,
        textColor=colors.HexColor("#111827"),
    )
    right_style = ParagraphStyle(
        "RightStyle",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=11,
        leading=16,
        textColor=colors.HexColor("#111827"),
        alignment=2,  # Right align for Date and Sign-off
    )
    story = []
    
    # 1. Header with From (Flush Left) and Date (Flush Right)
    sender_address_block = data["sender_address"].replace("\n", "<br/>")
    sender_html = f"From: <br/>{data['sender_designation']}<br/>{sender_address_block}"
    from_and_date_data = [
        [
            Paragraph(sender_html, body_style),
            Paragraph(f"Date: {data['date'].strftime('%d %b %Y')}", right_style),
        ]
    ]
    # Total width equals exactly 174mm
    from_and_date_table = Table(from_and_date_data, colWidths=[120 * mm, 54 * mm], hAlign='LEFT')
    from_and_date_table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (0, -1), 0),    # Forces 'From' to touch the exact left margin
                ("RIGHTPADDING", (-1, 0), (-1, -1), 0), # Forces 'Date' to touch the exact right margin
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story.append(from_and_date_table)
    story.append(Spacer(1, 5))
    
    # 2. To Block (Naturally Flush Left)
    story.append(Paragraph(f"To,<br/>{data['bank_name']}<br/>{data['branch_name']}", body_style))
    story.append(Spacer(1, 6))
    
    # 3. Subject (Naturally Flush Left)
    story.append(Paragraph("Subject: Transfer of money electronically", label_style))
    story.append(Spacer(1, 6))
    
    # Reference if available
    if data["reference"]:
        story.append(Paragraph(f"Reference: {data['reference']}", body_style))
        story.append(Spacer(1, 6))
    
    story.append(Paragraph("Dear Sir/Madam,<br/>", body_style))
    story.append(Spacer(1, 4))
    
    # Calculate totals and amounts in words
    total_amount = sum(item["amount"] for item in data["beneficiaries"])
    amount_words = amount_to_words(total_amount)
    
    # 4. Main Message (Naturally Flush Left)
    program_account_info = f", bearing account no. {data.get('program_account_number', '')}" if data.get('program_account_number') else ""

    base_message = (
        f"With due respect, please credit electronically to the accounts below with the amount "
        f"Rs. {total_amount:,.2f} /- ({amount_words}) from {data.get('program_account_name', 'Program Account')}"
        f"{program_account_info}"
    )

    if data.get('cheque_number'):
        main_message = base_message + f" and cheque no. {data['cheque_number']}."
    else:
        main_message = base_message + "."

    story.append(Paragraph(main_message, body_style))
    story.append(Spacer(1, 10))
    
    # 5. Beneficiary Table (Forced Flush Left)
    table_header_style = ParagraphStyle(
        "TableHeaderStyle",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=12,
        textColor=colors.black,
    )
    table_cell_style = ParagraphStyle(
        "TableCellStyle",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=12,
        textColor=colors.HexColor("#111827"),
    )
    
    table_data = [
        [
            Paragraph("S.No.", table_header_style),
            Paragraph("Beneficiary", table_header_style),
            Paragraph("Bank name", table_header_style),
            Paragraph("Account Number", table_header_style),
            Paragraph("IFSC", table_header_style),
            Paragraph("Amount (INR)", table_header_style),
        ]
    ]
    for index, item in enumerate(data["beneficiaries"], 1):
        table_data.append(
            [
                Paragraph(str(index), table_cell_style),
                Paragraph(item["name"], table_cell_style),
                Paragraph(item["bank_branch"], table_cell_style),
                Paragraph(item["account_number"], table_cell_style),
                Paragraph(item["ifsc"], table_cell_style),
                Paragraph(f"{item['amount']:,.2f}", table_cell_style),
            ]
        )
    table_data.append(
        [
            Paragraph("", table_cell_style),
            Paragraph("Total", table_cell_style),
            Paragraph("", table_cell_style),
            Paragraph("", table_cell_style),
            Paragraph("", table_cell_style),
            Paragraph(f"{total_amount:,.2f}", table_cell_style),
        ]
    )
    
    # Total width equals exactly 174mm, aligned left so the grid border matches paragraph text
    table = Table(table_data, colWidths=[12 * mm, 34 * mm, 38 * mm, 38 * mm, 27 * mm, 25 * mm], hAlign='LEFT')
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e5e7eb")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#9ca3af")),
                ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
                ("VALIGN", (0, 1), (-1, -1), "TOP"),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("ALIGN", (4, 1), (4, -1), "RIGHT"),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#f3f4f6")),
            ]
        )
    )
    story.append(table)
    story.append(Spacer(1, 14))
    
    # 6. Regards block (Sign-off name and phone flush right)
    regards_data = [
        [
            "", # Empty left column spacer
            Paragraph(
                f"<br/><br/>{data['sender_name']}"
                + (f"<br/>Phone: {data['sender_phone']}" if data.get("sender_phone") else ""),
                right_style,
            ),
        ]
    ]
    # Total width equals exactly 174mm
    regards_table = Table(regards_data, colWidths=[80 * mm, 94 * mm], hAlign='LEFT')
    regards_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("RIGHTPADDING", (-1, 0), (-1, -1), 0), # Forces text to touch the exact right margin
    ]))
    story.append(regards_table)
    
    document.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

# --- Streamlit User Interface ---
st.title("Bank Transfer Letter PDF Generator")
st.write("Fill in the details below and download a ready-to-print PDF letter.")

with st.form("transfer_letter_form"):
    st.subheader("Letter Details")
    bank_name = st.text_input("To", value="The BM")
    branch_name = st.text_input("Branch name", value="SBI, Ranpur, Dist.- Nayagarh")
    sender_name = st.text_input("Your name / Company name", placeholder="ABC Enterprises")
    sender_designation = st.text_input("From (Designation)", value="M.O I/C")
    sender_address = st.text_area("Your address", value="PHC Darpanarayanpur, Dist.- Nayagarh")
    sender_phone = st.text_input("Your phone (optional)", placeholder="+91-98765-43210")
    cheque_number = st.text_input("Cheque number (optional)", placeholder="000123")
    letter_date = st.date_input("Letter date", value=date.today())
    program_account_name = st.text_input("Program Account name", value="HWC PHC Darpanarayanpur")
    program_account_number = st.text_input("Program Account number", placeholder="1234567890")
    reference = st.text_input("Reference / Notes", placeholder="Optional reference")

    st.subheader("Beneficiary Accounts")
    beneficiary_inputs: dict[str, Any] = {}
    for index in range(int(beneficiary_count)):
        st.markdown(f"**Account {index + 1}**")
        beneficiary_inputs[f"beneficiary_name_{index}"] = st.text_input(
            f"Beneficiary name {index + 1}",
            key=f"beneficiary_name_{index}",
            placeholder="Recipient name",
        )
        beneficiary_inputs[f"bank_branch_{index}"] = st.text_input(
            f"Bank name with branch {index + 1}",
            key=f"bank_branch_{index}",
            placeholder="SBI, Ranpur Branch",
        )
        beneficiary_inputs[f"account_number_{index}"] = st.text_input(
            f"Account number {index + 1}",
            key=f"account_number_{index}",
            placeholder="1234567890",
        )
        beneficiary_inputs[f"ifsc_{index}"] = st.text_input(
            f"IFSC {index + 1}",
            key=f"ifsc_{index}",
            placeholder="SBIN0000001",
        )
        beneficiary_inputs[f"amount_{index}"] = st.number_input(
            f"Amount (INR) {index + 1}",
            min_value=0.0,
            value=0.0,
            step=1000.0,
            format="%.2f",
            key=f"amount_{index}",
        )

    submitted = st.form_submit_button("Generate PDF")

if submitted:
    values = {
        "bank_name": bank_name,
        "branch_name": branch_name,
        "sender_name": sender_name,
        "sender_designation": sender_designation,
        "sender_phone": sender_phone,
        "sender_address": sender_address,
        "cheque_number": cheque_number,
        "date": letter_date,
        "program_account_name": program_account_name,
        "program_account_number": program_account_number,
        "reference": reference,
        "beneficiary_count": int(beneficiary_count),
        **beneficiary_inputs,
    }

    missing_fields = []
    for key in ["bank_name", "branch_name", "sender_name", "sender_designation"]:
        if not values[key].strip():
            missing_fields.append(key.replace("_", " "))

    if missing_fields:
        st.error(f"Please fill in the required fields: {', '.join(missing_fields)}.")
    else:
        letter_data = build_letter_data(values)
        pdf_bytes = generate_pdf(letter_data)
        st.success("PDF generated successfully.")
        st.download_button(
            label="Download PDF",
            data=pdf_bytes,
            file_name="bank_transfer_letter.pdf",
            mime="application/pdf",
        )
        st.info(
            "Tip: review the PDF before sending it to the bank, especially the beneficiary account numbers and amounts."
        )