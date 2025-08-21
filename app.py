import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from datetime import datetime
from io import BytesIO

st.title("Factuur Generator")

# Inputs
leveranciersnummer = st.text_input("Supplier Number (e.g., 1322 or 277)", "1322")
factuurnummer = st.text_input("Invoice Number", "INV-000")
shipping = st.number_input("Shipping & Handling (EUR)", min_value=0.0, value=20.0, step=0.01)
uploaded_file = st.file_uploader("Upload XLSX File (e.g., gecombineerd_resultaat-1322.xlsx)", type="xlsx")

if uploaded_file and st.button("Generate Invoice"):
    # Lees XLSX uit upload
    df = pd.read_excel(uploaded_file, usecols="A:D", header=None, names=["Part Number", "Description", "Quantity", "Price"])
    df["Total"] = df["Quantity"] * df["Price"]

    # PDF buffer
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=3*cm, bottomMargin=2*cm)
    elements = []
    styles = getSampleStyleSheet()

    # Stijlen
    company_header = ParagraphStyle(name='CompanyHeader', parent=styles['Title'], fontSize=18, alignment=1, spaceAfter=8, textColor=colors.red, fontName='Helvetica-Bold', backColor=colors.HexColor("#f0f0f0"))
    bold_style = ParagraphStyle(name='Bold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1)
    normal_style = ParagraphStyle(name='Normal', parent=styles['Normal'], fontSize=10, leading=12, alignment=1)
    header_style = ParagraphStyle(name='Header', parent=styles['Title'], fontSize=16, spaceAfter=10, alignment=1, textColor=colors.red)
    footer_bold = ParagraphStyle(name='FooterBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, alignment=1, spaceAfter=8)
    footer_style = ParagraphStyle(name='Footer', parent=styles['Normal'], fontSize=9, leading=11, alignment=1, textColor=colors.grey)

    # Bedrijfsgegevens
    company_data = [
        [Paragraph("CLASSIC SUZUKI PARTS NL", company_header)],
        [HRFlowable(width="50%", thickness=1, lineCap='round', color=colors.black, hAlign='CENTER')],
        [Paragraph("Vlaandere Motoren - de Marne 136 B - 8701MC - Bolsward - Tel: 00316-41484547", bold_style)],
        [Paragraph("IBAN NL49 RABO 0372 0041 64 - SWIFT RABONL2U", normal_style)],
        [Paragraph("VATnumber 8077 51 911 B01 - C.O.C.number 01018576", normal_style)]
    ]
    company_table = Table(company_data, colWidths=[doc.width], hAlign='CENTER')
    company_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0, colors.transparent),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(company_table)
    elements.append(Spacer(1, 20))

    # Factuurgegevens
    elements.append(Paragraph("INVOICE", header_style))
    elements.append(HRFlowable(width="100%", thickness=1, color=colors.black, spaceAfter=10))
    bill_to_data = [
        ["Bill To:", "Invoice Number:", "Supplier Number:", "Invoice Date:"],
        ["CMS", factuurnummer, leveranciersnummer, datetime.today().strftime("%d-%m-%Y")],
        ["Artemisweg 245, 8239 DD Lelystad, Netherlands", "", "", ""]
    ]
    bill_table = Table(bill_to_data, colWidths=[4.5*cm, 4.5*cm, 4.5*cm, 4.5*cm])
    bill_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0, colors.transparent),
    ]))
    elements.append(bill_table)
    elements.append(Spacer(1, 24))
    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey, spaceAfter=12))

    # Producttabel
    data = [["Part Number", "Description", "Quantity", "Price", "Total"]]
    for _, row in df.iterrows():
        data.append([str(row["Part Number"]), str(row["Description"]), int(row["Quantity"]), f"€ {row['Price']:.2f}", f"€ {row['Total']:.2f}"])

    # Totalen
    subtotal = df["Total"].sum()
    total_ex_vat = subtotal + shipping
    vat = total_ex_vat * 0.21
    grand_total = total_ex_vat + vat
    data.append(["", "", "", "Subtotal", f"€ {subtotal:.2f}"])
    data.append(["", "", "", "Shipping & Handling", f"€ {shipping:.2f}"])
    data.append(["", "", "", "Total Ex. VAT", f"€ {total_ex_vat:.2f}"])
    data.append(["", "", "", "VAT 21%", f"€ {vat:.2f}"])
    data.append(["", "", "", "Grand Total", f"€ {grand_total:.2f}"])

    table = Table(data, colWidths=[3.5*cm, 8*cm, 2*cm, 2.5*cm, 3*cm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4a4a4a")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (2, 1), (-1, -1), "CENTER"),
        ("ALIGN", (-2, 1), (-1, -1), "RIGHT"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
        ("BACKGROUND", (-2, -5), (-1, -1), colors.HexColor("#f0f0f0")),
        ("FONTNAME", (-2, -5), (-1, -1), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 36))
    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey, spaceAfter=12))

    # Footer
    elements.append(Paragraph("Thank you for your order!", footer_bold))
    elements.append(Paragraph("Payment is due within 14 days.", footer_style))
    elements.append(Paragraph("Returns are allowed without reason offer return within 15 days after receiving the item.", footer_style))
    elements.append(Paragraph("CLASSIC SUZUKI PARTS NL | IBAN: NL49 RABO 0372 0041 64", footer_style))
    elements.append(Paragraph("VATnumber 8077 51 911 B01 | C.O.C.number 01018576", footer_style))
    elements.append(Paragraph("Thank you for doing business with us.", footer_style))

    # Bouw PDF en download
    doc.build(elements)
    buffer.seek(0)
    st.download_button("Download Invoice PDF", buffer, file_name=f"invoice_{leveranciersnummer}.pdf", mime="application/pdf")