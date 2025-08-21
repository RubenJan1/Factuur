import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
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
            if df.shape[1] != 4:  # Controleer of er 4 kolommen zijn
                st.error(f"File {file.name} does not have exactly 4 columns.")
                continue
            all_dataframes.append(df)
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
    
    if all_dataframes:
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        combined_df = combined_df.sort_values(by=0)  # Sort by Part Number
        combined_df.columns = ["Part Number", "Description", "Quantity", "Price"]
        # Controleer op geldige numerieke waarden
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
        # Controleer of de bewerkte lijst nog geldig is
        if edited_df.empty:
            st.error("Edited list is empty. Please ensure there are valid items.")
        else:
            st.session_state['edited_df'] = edited_df
            st.success("List saved for invoice generation!")
else:
    st.info("Combine files first.")

# Step 3: Generate Invoice
st.header("Step 3: Generate Invoice")
leveranciersnummer = st.text_input("Supplier Number", leveranciersgroep)
factuurnummer = st.text_input("Invoice Number", "INV-000")
shipping = st.number_input("Shipping & Handling (EUR)", min_value=0.0, value=20.0, step=0.01)

if 'edited_df' in st.session_state and st.button("Generate Invoice"):
    df = st.session_state['edited_df']
    # Controleer of de DataFrame geldige gegevens heeft
    if df.empty or df[["Part Number", "Description", "Quantity", "Price"]].isna().all().any():
        st.error("Invalid or empty data in the edited list. Please check and try again.")
    else:
        try:
            df["Total"] = df["Quantity"] * df["Price"]  # Berekening gebeurt hier
            if df["Total"].isna().any():
                st.error("Some totals could not be calculated due to invalid Quantity or Price values.")
            else:
                # PDF buffer
                buffer = BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=3*cm, bottomMargin=2*cm)
                elements = []
                styles = getSampleStyleSheet()

                # Stijlen
                company_header = ParagraphStyle(name='CompanyHeader', parent=styles['Title'], fontSize=18, alignment=1, spaceAfter=8, textColor=colors.red, fontName='Helvetica-Bold', backColor=colors.HexColor("#f0f0f0"))
                bold_style = ParagraphStyle(name='Bold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1)
                normal_style = ParagraphStyle(name='Normal', parent=styles['Normal'], fontSize=10, leading=12, alignment=1)
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
                bill_to_data = [
                    ["Bill To:", "Invoice Number:", "Supplier Number:", "Invoice Date:"],
                    ["CMS", factuurnummer, leveranciersnummer, datetime.today().strftime("%d-%m-%Y")],
                    [Paragraph("Artemisweg 245, 8239 DD Lelystad, Netherlands", normal_style), "", "", ""]
                ]
                bill_table = Table(bill_to_data, colWidths=[4.5*cm, 4.5*cm, 4.5*cm, 4.5*cm])
                bill_table.setStyle(TableStyle([
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 10),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("GRID", (0, 0), (-1, -1), 0, colors.transparent),
                    ("SPAN", (0, 2), (3, 2)),  # Adres over alle kolommen
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
                data.extend([
                    ["", "", "", "", ""],  # Lege rij voor scheiding
                    ["", "", "", "Subtotal", f"€ {subtotal:.2f}"],
                    ["", "", "", "Shipping & Handling", f"€ {shipping:.2f}"],
                    ["", "", "", "Total Ex. VAT", f"€ {total_ex_vat:.2f}"],
                    ["", "", "", "VAT 21%", f"€ {vat:.2f}"],
                    ["", "", "", "Grand Total", f"€ {grand_total:.2f}"]
                ])

                table = Table(data, colWidths=[3.5*cm, 8*cm, 2*cm, 2.5*cm, 3*cm])
                table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4a4a4a")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 10),
                    ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                    ("ALIGN", (2, 1), (-1, -1), "CENTER"),
                    ("ALIGN", (-2, 1), (-1, -1), "RIGHT"),
                    ("GRID", (0, 0), (-1, -5), 0.5, colors.grey),  # Grid alleen voor producten
                    ("BOX", (0, 0), (-1, -1), 1, colors.black),
                    ("BACKGROUND", (-2, -4), (-1, -1), colors.HexColor("#f0f0f0")),
                    ("FONTNAME", (-2, -4), (-1, -1), "Helvetica-Bold"),
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

                # Bouw PDF
                doc.build(elements)
                buffer.seek(0)
                st.download_button("Download Invoice PDF", buffer, file_name=f"invoice_{leveranciersnummer}.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"Error generating PDF: {str(e)}")