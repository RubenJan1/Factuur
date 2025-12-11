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

                # ------------------------------
                #   PROFESSIONELE FACTUUR UI
                # ------------------------------

                PRIMARY = colors.HexColor("#007B80")     # Petrol / blauwgroen
                TEXT_DARK = colors.HexColor("#333333")   # Antraciet
                ROW_LIGHT = colors.HexColor("#F7F7F7")   # Lichtgrijze rij

                buffer = BytesIO()
                doc = SimpleDocTemplate(
                    buffer,
                    pagesize=A4,
                    rightMargin=1.5*cm,
                    leftMargin=1.5*cm,
                    topMargin=2.2*cm,
                    bottomMargin=1.5*cm
                )

                elements = []
                styles = getSampleStyleSheet()

                # ------------------------------
                #   STIJLEN
                # ------------------------------
                title_style = ParagraphStyle(
                    name="Title",
                    fontSize=32,
                    textColor=PRIMARY,
                    fontName="Helvetica-Bold",
                    leading=34,
                )

                section_header = ParagraphStyle(
                    name="SectionHeader",
                    fontSize=12,
                    textColor=PRIMARY,
                    fontName="Helvetica-Bold",
                    spaceAfter=6
                )

                normal = ParagraphStyle(
                    name="NormalText",
                    fontSize=10,
                    textColor=TEXT_DARK,
                )

                bold = ParagraphStyle(
                    name="BoldText",
                    parent=normal,
                    fontName="Helvetica-Bold"
                )

                # ------------------------------
                #   TITEL + HORIZONTALE LIJN
                # ------------------------------
                elements.append(Paragraph("FACTUUR", title_style))

                elements.append(Table(
                    [[]],
                    colWidths=[doc.width],
                    style=[
                        ("LINEBELOW", (0,0), (-1,0), 3, PRIMARY)
                    ]
                ))
                elements.append(Spacer(1, 18))

                # ------------------------------
                #   FACTUUR INFO BLOK (RECHTS)
                # ------------------------------
                invoice_info = [
                    ["Factuurdatum:", datetime.today().strftime("%d-%m-%Y")],
                    ["Factuurnummer:", factuurnummer],
                    ["Leveranciersnummer:", leveranciersnummer],
                    ["Vervaldatum:", datetime.today().strftime("%d-%m-%Y")],
                ]

                info_table = Table(
                    invoice_info,
                    colWidths=[5*cm, 5*cm],
                    style=[
                        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
                        ("ALIGN", (0,0), (-1,-1), "LEFT"),
                        ("TEXTCOLOR", (0,0), (-1,-1), TEXT_DARK),
                        ("BOTTOMPADDING", (0,0), (-1,-1), 4)
                    ]
                )
                elements.append(info_table)
                elements.append(Spacer(1, 15))

                # ------------------------------
                #   BILL TO BLOK (LINKS)
                # ------------------------------
                elements.append(Paragraph("Factuur aan", section_header))

                bill_to = [
                    ["CMS"],
                    ["Artemisweg 245"],
                    ["8239 DD Lelystad"],
                    ["Netherlands"],
                ]

                bill_table = Table(
                    bill_to,
                    colWidths=[doc.width/2],
                    style=[
                        ("ALIGN", (0,0), (-1,-1), "LEFT"),
                        ("TEXTCOLOR", (0,0), (-1,-1), TEXT_DARK),
                        ("BOTTOMPADDING", (0,0), (-1,-1), 2)
                    ]
                )

                elements.append(bill_table)
                elements.append(Spacer(1, 20))

                # ------------------------------
                #   PRODUCTTABEL
                # ------------------------------
                elements.append(Paragraph("Productspecificaties", section_header))

                table_data = [["Part Number", "Description", "Qty", "Price", "Total"]]

                # tabelregels invullen
                for i, row in df.iterrows():
                    table_data.append([
                        str(row["Part Number"]),
                        str(row["Description"]),
                        int(row["Quantity"]),
                        f"€ {row['Price']:.2f}",
                        f"€ {row['Total']:.2f}"
                    ])

                # Totaalberekeningen
                subtotal = df["Total"].sum()
                total_ex_vat = subtotal + shipping
                vat = total_ex_vat * 0.21
                grand_total = total_ex_vat + vat

                # tabel toevoegen
                product_table = Table(
                    table_data,
                    colWidths=[3.5*cm, 8*cm, 2*cm, 3*cm, 3*cm],
                    style=[
                        ("BACKGROUND", (0,0), (-1,0), PRIMARY),
                        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                        ("GRID", (0,1), (-1,-1), 0.3, colors.grey),
                        ("ALIGN", (2,1), (-1,-1), "RIGHT"),
                        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, ROW_LIGHT]),
                        ("TOPPADDING", (0,0), (-1,0), 6),
                        ("BOTTOMPADDING", (0,0), (-1,0), 6),
                    ]
                )
                elements.append(product_table)
                elements.append(Spacer(1, 25))

                # ------------------------------
                #   TOTALEN BLOK (RECHTS)
                # ------------------------------
                totals = [
                    ["Subtotal:", f"€ {subtotal:.2f}"],
                    ["Shipping:", f"€ {shipping:.2f}"],
                    ["Total Ex. VAT:", f"€ {total_ex_vat:.2f}"],
                    ["VAT (21%):", f"€ {vat:.2f}"],
                    ["Grand Total:", f"€ {grand_total:.2f}"],
                ]

                totals_table = Table(
                    totals,
                    colWidths=[6*cm, 4*cm],
                    style=[
                        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
                        ("ALIGN", (1,0), (-1,-1), "RIGHT"),
                        ("TEXTCOLOR", (0,-1), (-1,-1), PRIMARY),
                        ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),
                        ("BOTTOMPADDING", (0,0), (-1,-1), 4)
                    ]
                )
                elements.append(Spacer(1, 10))
                elements.append(totals_table)
                elements.append(Spacer(1, 35))

                # ------------------------------
                #   FOOTER
                # ------------------------------
                footer_text = [
                    "Thank you for your order.",
                    "Payment term: 14 days.",
                    "Classic Suzuki Parts NL – IBAN NL49 RABO 0372 0041 64",
                    "VATnumber 8077 51 911 B01 | C.O.C.number 01018576"
                ]

            for line in footer_text:
                elements.append(Paragraph(line, normal))

                # Bouw PDF
                doc.build(elements)
                buffer.seek(0)
                st.download_button(
                    label="Download Invoice PDF",
                    data=buffer,
                    file_name=pdf_filename,
                    mime="application/pdf"
                )
                

        except Exception as e:
            st.error(f"Error generating PDF: {str(e)}")
