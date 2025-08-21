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
        # Vervang lege beschrijvingen door "N/A"
        combined_df["Description"] = combined_df["Description"].fillna("N/A")
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
        # Controleer op geldige bewerkte lijst
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
leveranciersnummer = st.text_input("Supplier Number", leveranciersgroep)
factuurnummer = st.text_input("Invoice Number", "INV-000")
shipping = st.number_input("Shipping & Handling (EUR)", min_value=0.0, value=20.0, step=0.01)

if 'edited_df' in st.session_state and st.button("Generate Invoice"):
    df = st.session_state['edited_df']
    if df.empty or df[["Part Number", "Description", "Quantity", "Price"]].isna().all().any():
        st.error("Invalid or empty data in the edited list. Please check and try again.")
    else:
        try:
            df["Total"] = df["Quantity"] * df["Price"]
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=2.5*cm, bottomMargin=2*cm)
            elements = []
            styles = getSampleStyleSheet()

            # Styles
            company_header = ParagraphStyle(
                name='CompanyHeader', fontSize=20, alignment=1, textColor=colors.HexColor("#e60000"),
                fontName='Helvetica-Bold', spaceAfter=6
            )
            bold_center = ParagraphStyle(name='BoldCenter', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, alignment=1)
            normal_left = ParagraphStyle(name='NormalLeft', parent=styles['Normal'], fontSize=10, leading=12, alignment=0)
            footer_style = ParagraphStyle(name='Footer', parent=styles['Normal'], fontSize=9, leading=11, alignment=1, textColor=colors.grey)

            # Company Header
            company_info = [
                [Paragraph("CLASSIC SUZUKI PARTS NL", company_header)],
                [Paragraph("Vlaandere Motoren - de Marne 136 B - 8701MC - Bolsward", normal_left)],
                [Paragraph("Tel: 00316-41484547 | IBAN: NL49 RABO 0372 0041 64 | VAT: 807751911B01", normal_left)],
                [Paragraph("C.O.C. Number: 01018576", normal_left)]
            ]
            company_table = Table(company_info, colWidths=[doc.width])
            company_table.setStyle(TableStyle([("ALIGN", (0,0), (-1,-1), "CENTER"), ("BOTTOMPADDING", (0,0), (-1,-1), 4)]))
            elements.append(company_table)
            elements.append(Spacer(1, 15))

            # Invoice & Bill To
            bill_data = [
                ["Bill To:", "Invoice Number:", "Supplier Number:", "Invoice Date:"],
                ["CMS", factuurnummer, leveranciersnummer, datetime.today().strftime("%d-%m-%Y")],
                ["Artemisweg 245, 8239 DD Lelystad, Netherlands", "", "", ""]
            ]
            bill_table = Table(bill_data, colWidths=[5*cm, 4*cm, 4*cm, 4*cm])
            bill_table.setStyle(TableStyle([
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE", (0,0), (-1,-1), 10),
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("SPAN", (0,2), (3,2)),
            ]))
            elements.append(bill_table)
            elements.append(Spacer(1, 12))

            # Products Table
            data = [["Part Number", "Description", "Quantity", "Price", "Total"]]
            for i, row in df.iterrows():
                data.append([
                    str(row["Part Number"]),
                    str(row["Description"]),
                    int(row["Quantity"]),
                    f"€ {row['Price']:.2f}",
                    f"€ {row['Total']:.2f}"
                ])

            subtotal = df["Total"].sum()
            total_ex_vat = subtotal + shipping
            vat = total_ex_vat * 0.21
            grand_total = total_ex_vat + vat

            # Add totals
            data.extend([
                ["", "", "", "Subtotal", f"€ {subtotal:.2f}"],
                ["", "", "", "Shipping & Handling", f"€ {shipping:.2f}"],
                ["", "", "", "Total Ex. VAT", f"€ {total_ex_vat:.2f}"],
                ["", "", "", "VAT 21%", f"€ {vat:.2f}"],
                ["", "", "", "Grand Total", f"€ {grand_total:.2f}"]
            ])

            table = Table(data, colWidths=[3*cm, 8*cm, 2*cm, 2.5*cm, 3*cm])
            table.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#4a4a4a")),
                ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                ("ALIGN", (2,1), (-1,-1), "CENTER"),
                ("ALIGN", (-2,1), (-1,-1), "RIGHT"),
                ("GRID", (0,0), (-1,-6), 0.5, colors.grey),
                ("BOX", (0,0), (-1,-1), 1, colors.black),
                ("BACKGROUND", (-2,-5), (-1,-1), colors.HexColor("#f0f0f0")),
                ("FONTNAME", (-2,-5), (-1,-1), "Helvetica-Bold"),
            ]))
            elements.append(table)
            elements.append(Spacer(1, 25))

            # Footer
            elements.append(Paragraph("Thank you for your order!", bold_center))
            elements.append(Paragraph("Payment due within 14 days. Returns accepted within 15 days.", footer_style))
            elements.append(Paragraph("CLASSIC SUZUKI PARTS NL | IBAN: NL49 RABO 0372 0041 64 | VAT: 807751911B01", footer_style))

            doc.build(elements)
            buffer.seek(0)
            st.download_button("Download Invoice PDF", buffer, file_name=f"invoice_{leveranciersnummer}.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"Error generating PDF: {str(e)}")
