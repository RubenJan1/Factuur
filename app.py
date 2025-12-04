import streamlit as st 
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from datetime import datetime
from io import BytesIO

st.title("Factuur Generator - Classic Suzuki Parts NL")

# Step 1: Upload and Combine XLSX Files
st.header("Step 1: Upload and Combine XLSX Files")
leveranciersgroep = st.selectbox("Select Supplier Group", ["1322", "277"])

uploaded_files = st.file_uploader("Upload XLSX Files for Combination (e.g., 1322-2084856.xlsx)", type="xlsx", accept_multiple_files=True)

combined_df = None
if uploaded_files and st.button("Combine Files"):
    all_dataframes = []
    for file in uploaded_files:
        try:
            df = pd.read_excel(file, header=None, usecols=[0, 1, 2, 3])
            df = df.dropna(how='all')
            if df.shape[1] != 4:
                st.error(f"File {file.name} does not have exactly 4 columns.")
                continue
            all_dataframes.append(df)
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
    
    if all_dataframes:
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        combined_df = combined_df.sort_values(by=0)  # Sort by Part Number
        combined_df.columns = ["Part Number", "Description", "Quantity", "Price"]
        combined_df["Description"] = combined_df["Description"].fillna("N/A")

        try:
            combined_df["Quantity"] = pd.to_numeric(combined_df["Quantity"], errors='coerce')
            combined_df["Price"] = pd.to_numeric(combined_df["Price"], errors='coerce')
            combined_df = combined_df.dropna(subset=["Quantity", "Price"])
            if combined_df.empty:
                st.error("Combined list is empty or contains no valid numeric data.")
            else:
                st.session_state['combined_df'] = combined_df
                st.success("Files combined successfully!")
                st.write("Combined List Preview:")
                st.dataframe(combined_df)
        except Exception as e:
            st.error(f"Error processing combined data: {str(e)}")
    else:
        st.error("No valid files processed.")

# Step 2: Edit Combined List
st.header("Step 2: Edit Combined List")
if 'combined_df' in st.session_state:
    df = st.session_state['combined_df']
    edited_df = st.data_editor(df, num_rows="dynamic")
    if st.button("Save Edited List"):
        if edited_df.empty:
            st.error("Edited list is empty. Please ensure there are valid items.")
        else:
            edited_df["Description"] = edited_df["Description"].fillna("N/A")
            st.session_state['edited_df'] = edited_df
            st.success("List saved for invoice generation!")
else:
    st.info("Combine files first.")

# Step 3: Generate Invoice
st.header("Step 3: Generate Invoice")

leveranciersnummer = st.text_input("Supplier Number", value=leveranciersgroep, key="lev_nr")
factuurnummer = st.text_input("Invoice Number", value="INV-000", key="fact_nr")
shipping = st.number_input("Shipping & Handling (EUR)", min_value=0.0, value=20.0, step=0.01)

# ← NIEUW: invoerveld voor bestandsnaam
default_filename = f"Factuur_{leveranciersnummer}_{factuurnummer}.pdf"
pdf_filename = st.text_input(
    "Bestandsnaam voor de PDF (incl. .pdf)",
    value=default_filename,
    help="Je kunt de naam aanpassen. Plaats {leverancier} of {factuur} als je die automatisch wilt invullen."
)

# Vervang eventuele placeholders
pdf_filename = pdf_filename.replace("{leverancier}", leveranciersnummer).replace("{factuur}", factuurnummer)

# Forceer .pdf-extensie
if not pdf_filename.lower().endswith(".pdf"):
    pdf_filename += ".pdf"
if 'edited_df' in st.session_state and st.button("Generate Invoice"):
    df = st.session_state['edited_df']

    required_cols = ["Part Number", "Description", "Quantity", "Price"]
    if df.empty or df[required_cols].isna().all().any():
        st.error("Invalid or empty data in the edited list. Please check and try again.")
    else:
        try:
            df["Total"] = df["Quantity"].astype(float) * df["Price"].astype(float)
            if df["Total"].isna().any():
                st.error("Some totals could not be calculated due to invalid Quantity or Price values.")
            else:
                buffer = BytesIO()
                doc = SimpleDocTemplate(
                    buffer,
                    pagesize=A4,
                    rightMargin=1.5*cm,
                    leftMargin=1.5*cm,
                    topMargin=3*cm,
                    bottomMargin=2*cm
                )
                elements = []
                styles = getSampleStyleSheet()

                # Stijlen
                company_header = ParagraphStyle(
                    name='CompanyHeader', parent=styles['Title'],
                    fontSize=22, alignment=1, spaceAfter=14,
                    textColor=colors.red, fontName='Helvetica-Bold'
                )
                footer_style = ParagraphStyle(
                    name='Footer', parent=styles['Normal'], fontSize=9, leading=11, alignment=1, textColor=colors.grey
                )
                bold_style = ParagraphStyle(
                    name='Bold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=12
                )
                normal_style = ParagraphStyle(
                    name='Normal', parent=styles['Normal'], fontSize=10, leading=12
                )
                footer_bold = ParagraphStyle(
                    name='FooterBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, alignment=1
                )

                # Titel + Bedrijfsgegevens
                company_data = [
                    [Paragraph("CLASSIC SUZUKI PARTS NL", company_header)],
                    [Paragraph("Vlaandere Motoren - de Marne 136 B - 8701MC - Bolsward - Tel: 00316-41484547", footer_style)],
                    [Paragraph("IBAN NL49 RABO 0372 0041 64 - SWIFT RABONL2U", footer_style)],
                    [Paragraph("VATnumber 8077 51 911 B01 - C.O.C.number 01018576", footer_style)]
                ]
                company_table = Table(company_data, colWidths=[doc.width], hAlign='CENTER')
                company_table.setStyle(TableStyle([
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                ]))
                elements.append(company_table)
                elements.append(Spacer(1, 20))

                # Factuurgegevens
                bill_to_data = [
                    ["Bill To:", "Invoice Number:", "Supplier Number:", "Invoice Date:"],
                    ["CMS", factuurnummer, leveranciersnummer, datetime.today().strftime("%d-%m-%Y")],
                    [Paragraph("Artemisweg 245, 8239 DD Lelystad, Netherlands", normal_style), "", "", ""]
                ]
                bill_table = Table(bill_to_data, colWidths=[4.5*cm]*4)
                bill_table.setStyle(TableStyle([
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("SPAN", (0, 2), (3, 2)),
                ]))
                elements.append(bill_table)
                elements.append(Spacer(1, 24))

                # Producttabel
                data = [["Part Number", "Description", "Quantity", "Price", "Total"]]
                for _, row in df.iterrows():
                    data.append([
                        str(row["Part Number"]) if pd.notna(row["Part Number"]) else "N/A",
                        str(row["Description"]) if pd.notna(row["Description"]) else "N/A",
                        int(row["Quantity"]),
                        f"€ {row['Price']:.2f}",
                        f"€ {row['Total']:.2f}"
                    ])

                subtotal = df["Total"].sum()
                total_ex_vat = subtotal + shipping
                vat = total_ex_vat * 0.21
                grand_total = total_ex_vat + vat

                data.extend([
                    ["", "", "", "", ""],
                    ["", "", "", "Subtotal", f"€ {subtotal:.2f}"],
                    ["", "", "", "Shipping & Handling", f"€ {shipping:.2f}"],
                    ["", "", "", "Total Ex. VAT", f"€ {total_ex_vat:.2f}"],
                    ["", "", "", "VAT 21%", f"€ {vat:.2f}"],
                    ["", "", "", "Grand Total", f"€ {grand_total:.2f}"]
                ])

                table = Table(data, colWidths=[3.5*cm, 8*cm, 2*cm, 3*cm, 3*cm])
                table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4a4a4a")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("ALIGN", (2, 1), (-1, -1), "CENTER"),
                    ("ALIGN", (-2, 1), (-1, -1), "RIGHT"),
                    ("GRID", (0, 0), (-1, -6), 0.5, colors.grey),
                    ("BOX", (0, 0), (-1, -1), 1, colors.black),
                    ("FONTNAME", (-2, -5), (-1, -1), "Helvetica-Bold"),
                ]))
                elements.append(table)
                elements.append(Spacer(1, 36))

                # Footer
                elements.append(Paragraph("Thank you for your order!", footer_bold))
                elements.append(Paragraph("Payment is due within 14 days.", footer_style))
                elements.append(Paragraph("Returns allowed within 15 days after receiving the item.", footer_style))
                elements.append(Paragraph("Vlaandere Motoren | IBAN: NL49 RABO 0372 0041 64", footer_style))
                elements.append(Paragraph("VATnumber 8077 51 911 B01 | C.O.C.number 01018576", footer_style))
                elements.append(Paragraph("Thank you for doing business with us.", footer_bold))

                # Bouw PDF
                doc.build(elements)
                buffer.seek(0)
                st.download_button("Download Invoice PDF", buffer, file_name=f"invoice_{leveranciersnummer}.pdf", mime="application/pdf")

        except Exception as e:
            st.error(f"Error generating PDF: {str(e)}")
