# app.py
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PIL import Image
import base64

# -----------------------
# App config & layout
# -----------------------
st.set_page_config(page_title="Factuur Generator - Classic Suzuki Parts NL", layout="wide")
st.title("Factuur Generator — Classic Suzuki Parts NL")
st.write("Maak professionele, uitgelijnde PDF-facturen met logo, thema en slimme instellingen.")

# -----------------------
# Helper functions
# -----------------------
def currency(eur):
    """Return formatted euro string with two decimals."""
    try:
        return f"€ {float(eur):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return f"€ {eur}"

def pil_image_to_bytesio(pil_image, fmt="PNG"):
    bio = BytesIO()
    pil_image.save(bio, format=fmt)
    bio.seek(0)
    return bio

def read_excel_file(uploaded_file):
    # Allow various shapes: try to read with/without header
    try:
        df = pd.read_excel(uploaded_file)
        # If standard columns exist, try to map
        expected = ["Part Number", "Description", "Quantity", "Price"]
        cols = list(df.columns)
        if set(expected).issubset(set(cols)):
            df = df[expected]
        else:
            # If file has no headers, try reading first 4 columns
            df = pd.read_excel(uploaded_file, header=None, usecols=[0,1,2,3])
            df.columns = expected
    except Exception:
        # fallback: try reading without header
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=None, usecols=[0,1,2,3])
        df.columns = ["Part Number", "Description", "Quantity", "Price"]
    return df

# -----------------------
# Sidebar: settings
# -----------------------
st.sidebar.header("Instellingen & Thema")
with st.sidebar.form("settings_form"):
    primary_color = st.color_picker("Primaire kleur (accent)", "#007B80")
    text_color = st.color_picker("Tekstkleur", "#333333")
    light_row = st.color_picker("Licht rij-achtergrond", "#F7F7F7")
    default_vat = st.number_input("Standaard BTW percentage", min_value=0.0, max_value=100.0, value=21.0, step=0.5)
    default_shipping = st.number_input("Standaard verzendkosten (EUR)", min_value=0.0, value=20.0, step=0.01)
    default_payment_days = st.number_input("Standaard betalingstermijn (dagen)", min_value=0, value=14, step=1)
    submit_settings = st.form_submit_button("Opslaan instellingen")

# -----------------------
# Step 1: Upload / combine files
# -----------------------
st.header("Stap 1 — Upload parts (XLSX) of voer handmatig in")
uploaded_files = st.file_uploader("Upload één of meerdere XLSX bestanden (kolommen: Part Number, Description, Quantity, Price)", type=["xlsx"], accept_multiple_files=True)

if 'combined_df' not in st.session_state:
    st.session_state['combined_df'] = pd.DataFrame(columns=["Part Number", "Description", "Quantity", "Price"])

if uploaded_files:
    all_dfs = []
    for f in uploaded_files:
        try:
            df = read_excel_file(f)
            all_dfs.append(df)
        except Exception as e:
            st.error(f"Kon bestand {f.name} niet lezen: {e}")
    if all_dfs:
        combined = pd.concat(all_dfs, ignore_index=True)
        # sanitize / keep only necessary columns
        combined = combined[["Part Number", "Description", "Quantity", "Price"]]
        combined["Description"] = combined["Description"].fillna("N/A")
        combined["Quantity"] = pd.to_numeric(combined["Quantity"], errors='coerce').fillna(0).astype(int)
        combined["Price"] = pd.to_numeric(combined["Price"], errors='coerce').fillna(0.0).astype(float)
        st.session_state['combined_df'] = combined
        st.success("Bestanden gecombineerd en geladen.")

st.write("Voorvertoning gecombineerde lijst (je kunt rijen handmatig aanpassen):")
edited_df = st.data_editor(st.session_state['combined_df'], num_rows="dynamic", key="editor")

if st.button("Opslaan lijst"):
    # validate
    if edited_df.empty:
        st.error("Lijst is leeg, voeg producten toe.")
    else:
        edited_df["Description"] = edited_df["Description"].fillna("N/A")
        edited_df["Quantity"] = pd.to_numeric(edited_df["Quantity"], errors='coerce').fillna(0).astype(int)
        edited_df["Price"] = pd.to_numeric(edited_df["Price"], errors='coerce').fillna(0.0).astype(float)
        st.session_state['edited_df'] = edited_df
        st.success("Lijst opgeslagen voor factuurgeneratie.")

# -----------------------
# Step 2: Invoice metadata
# -----------------------
st.header("Stap 2 — Factuurgegevens & extra opties")
col1, col2, col3 = st.columns([2,2,1])
with col1:
    supplier_number = st.text_input("Leveranciersnummer", value="1322")
    invoice_number = st.text_input("Factuurnummer", value=f"INV-{datetime.now().strftime('%Y%m%d%H%M')}")
    invoice_date = st.date_input("Factuurdatum", value=datetime.today().date())
    payment_days = st.number_input("Betalingstermijn (dagen)", min_value=0, value=int(default_payment_days))
    due_date = st.date_input("Vervaldatum", value=(invoice_date + timedelta(days=int(payment_days))))
with col2:
    bill_to_name = st.text_input("Naam klant / factuur aan", value="CMS")
    bill_to_address = st.text_area("Adres klant", value="Artemisweg 245\n8239 DD Lelystad\nNetherlands", height=100)
    reference = st.text_input("Opmerking / referentie", value="")
with col3:
    # logo upload
    st.write("Logo (optioneel)")
    uploaded_logo = st.file_uploader("Upload logo (png/jpg). Wordt linksboven in de factuur gebruikt.", type=["png","jpg","jpeg"])
    preview_logo = None
    if uploaded_logo:
        try:
            img = Image.open(uploaded_logo)
            # create small preview for UI
            img_thumb = img.copy()
            img_thumb.thumbnail((300, 150))
            st.image(img_thumb, use_column_width=True)
            preview_logo = img
        except Exception as e:
            st.error(f"Kon logo niet laden: {e}")
    st.write("")
    vat_percent = st.number_input("BTW (%)", min_value=0.0, max_value=100.0, value=float(default_vat), step=0.5)
    shipping = st.number_input("Verzendkosten (EUR)", min_value=0.0, value=float(default_shipping), step=0.01)
    file_name = st.text_input("Bestandsnaam PDF", value=f"Factuur_{supplier_number}_{invoice_number}.pdf")

# -----------------------
# Step 3: Preview & Generate
# -----------------------
st.header("Stap 3 — Preview & Genereer PDF")
if 'edited_df' not in st.session_state:
    st.info("Sla eerst de bewerkte lijst op (Stap 1).")
else:
    df = st.session_state['edited_df'].copy()
    if df.empty:
        st.warning("Geen regels om te factureren. Voeg producten toe.")
    else:
        # calculate totals
        df["Total"] = df["Quantity"].astype(float) * df["Price"].astype(float)
        subtotal = float(df["Total"].sum())
        total_ex_vat = subtotal + float(shipping)
        vat = total_ex_vat * (float(vat_percent) / 100.0)
        grand_total = total_ex_vat + vat

        # show quick preview
        st.subheader("Factuuroverzicht")
        preview_col1, preview_col2 = st.columns([2,1])
        with preview_col1:
            st.markdown(f"**Factuurnummer:** {invoice_number}  ")
            st.markdown(f"**Factuurdatum:** {invoice_date.strftime('%d-%m-%Y')}")
            st.markdown(f"**Vervaldatum:** {due_date.strftime('%d-%m-%Y')}")
            st.markdown(f"**Factuur aan:** {bill_to_name}  ")
            for line in bill_to_address.splitlines():
                st.write(line)
        with preview_col2:
            st.markdown(f"**Subtotaal:** {currency(subtotal)}")
            st.markdown(f"**Verzending:** {currency(shipping)}")
            st.markdown(f"**Totaal excl. btw:** {currency(total_ex_vat)}")
            st.markdown(f"**BTW ({vat_percent}%):** {currency(vat)}")
            st.markdown(f"**Totaal:** **{currency(grand_total)}**")

        st.dataframe(df[["Part Number", "Description", "Quantity", "Price", "Total"]])

        if st.button("Genereer PDF en download"):
            try:
                # Build PDF
                buffer = BytesIO()
                doc = SimpleDocTemplate(
                    buffer,
                    pagesize=A4,
                    rightMargin=1.5*cm,
                    leftMargin=1.5*cm,
                    topMargin=2.2*cm,
                    bottomMargin=1.5*cm
                )
                styles = getSampleStyleSheet()

                # Colors from settings
                PRIMARY = colors.HexColor(primary_color)
                TEXT_DARK = colors.HexColor(text_color)
                ROW_LIGHT = colors.HexColor(light_row)

                # Styles
                title_style = ParagraphStyle("Title", parent=styles["Heading1"], fontSize=28, textColor=PRIMARY, leading=30, spaceAfter=6)
                comp_style = ParagraphStyle("Company", parent=styles["Normal"], fontSize=9, textColor=TEXT_DARK)
                section_header = ParagraphStyle("SectionHeader", parent=styles["Heading4"], fontSize=11, textColor=PRIMARY, leading=12)
                normal = ParagraphStyle("Normal", parent=styles["Normal"], fontSize=10, textColor=TEXT_DARK)
                small = ParagraphStyle("Small", parent=styles["Normal"], fontSize=8, textColor=TEXT_DARK)
                bold = ParagraphStyle("Bold", parent=styles["Normal"], fontSize=10, textColor=TEXT_DARK, fontName="Helvetica-Bold")

                elems = []

                # Header: logo left, company info right
                # Prepare logo image if present
                logo_cell = []
                if preview_logo is not None:
                    # Fit logo in a max box
                    logo_io = pil_image_to_bytesio(preview_logo.convert("RGBA"), fmt="PNG")
                    # RLImage expects either a filename or BytesIO
                    rl_logo = RLImage(logo_io, width=6*cm, height=2.6*cm, kind='proportional')
                    logo_cell = [rl_logo]
                else:
                    # fallback: company name as Paragraph
                    logo_cell = [Paragraph("<b>CLASSIC SUZUKI PARTS NL</b>", title_style)]

                company_lines = [
                    Paragraph("<b>CLASSIC SUZUKI PARTS NL</b>", title_style),
                    Paragraph("Vlaandere Motoren - de Marne 136 B", comp_style),
                    Paragraph("8701 MC - Bolsward", comp_style),
                    Paragraph("Tel: 00316-41484547", comp_style),
                    Paragraph("IBAN: NL49 RABO 0372 0041 64", comp_style),
                    Paragraph("VAT: 8077 51 911 B01 | C.O.C: 01018576", comp_style),
                ]

                header_table = Table(
                    [[logo_cell[0], company_lines]],
                    colWidths=[6.5*cm, doc.width - 6.5*cm]
                )
                header_table.setStyle(TableStyle([
                    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                    ("ALIGN", (1,0), (1,0), "RIGHT"),
                    ("LEFTPADDING", (0,0), (-1,-1), 0),
                    ("RIGHTPADDING", (0,0), (-1,-1), 0),
                ]))
                elems.append(header_table)
                # horizontal accent line
                elems.append(Spacer(1, 8))
                elems.append(Table([[" "]], colWidths=[doc.width], style=[("LINEBELOW", (0,0), (-1,0), 4, PRIMARY)]))
                elems.append(Spacer(1, 12))

                # Invoice & BillTo block side-by-side
                inv_info = [
                    ["Factuurnummer:", invoice_number],
                    ["Factuurdatum:", invoice_date.strftime("%d-%m-%Y")],
                    ["Vervaldatum:", due_date.strftime("%d-%m-%Y")],
                    ["Leveranciersnr:", supplier_number],
                ]
                inv_table = Table(inv_info, colWidths=[4*cm, 6*cm])
                inv_table.setStyle(TableStyle([
                    ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
                    ("ALIGN", (0,0), (-1,-1), "LEFT"),
                    ("VALIGN", (0,0), (-1,-1), "TOP"),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 4)
                ]))

                bill_lines = [Paragraph(f"<b>{bill_to_name}</b>", bold)]
                for ln in bill_to_address.splitlines():
                    bill_lines.append(Paragraph(ln, normal))

                bill_table = Table([[bill_lines]], colWidths=[doc.width - 10*cm])
                bill_table.setStyle(TableStyle([
                    ("LEFTPADDING", (0,0), (-1,-1), 0),
                    ("VALIGN", (0,0), (-1,-1), "TOP"),
                ]))

                side_table = Table([[bill_table, inv_table]], colWidths=[doc.width - 8.5*cm, 8.5*cm])
                side_table.setStyle(TableStyle([
                    ("VALIGN", (0,0), (-1,-1), "TOP"),
                    ("LEFTPADDING", (0,0), (-1,-1), 0),
                    ("RIGHTPADDING", (0,0), (-1,-1), 0),
                ]))
                elems.append(side_table)
                elems.append(Spacer(1, 18))

                # Product table header and rows
                table_data = [["Artikelnummer", "Omschrijving", "Aantal", "Prijs", "Totaal"]]
                for _, r in df.iterrows():
                    pn = str(r["Part Number"]) if pd.notna(r["Part Number"]) else "-"
                    desc = str(r["Description"]) if pd.notna(r["Description"]) else "-"
                    qty = int(r["Quantity"])
                    price = float(r["Price"])
                    total_row = float(qty) * float(price)
                    table_data.append([pn, desc, f"{qty}", f"€ {price:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), f"€ {total_row:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")])

                # Add empty row before totals
                table_data.append(["", "", "", "", ""])

                # Totals rows (as part of the same table to keep alignment)
                table_data.append(["", "", "", "Subtotal", f"€ {subtotal:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")])
                table_data.append(["", "", "", "Shipping", f"€ {float(shipping):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")])
                table_data.append(["", "", "", "Total Excl. BTW", f"€ {total_ex_vat:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")])
                table_data.append(["", "", "", f"BTW ({vat_percent}%)", f"€ {vat:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")])
                table_data.append(["", "", "", "Totaal", f"€ {grand_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")])

                col_widths = [3.5*cm, doc.width - (3.5+2+3+3)*cm, 2*cm, 3*cm, 3*cm]
                product_table = Table(table_data, colWidths=col_widths, repeatRows=1)
                product_table.setStyle(TableStyle([
                    ("BACKGROUND", (0,0), (-1,0), PRIMARY),
                    ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                    ("ALIGN", (2,1), (-1,-1), "RIGHT"),
                    ("GRID", (0,1), (-1,-6), 0.25, colors.lightgrey),
                    ("ROWBACKGROUNDS", (0,1), (-1,-6), [colors.white, ROW_LIGHT]),
                    ("SPAN", (0, len(table_data)-6), (2, len(table_data)-6)),  # empty row span
                    ("SPAN", (0, len(table_data)-5), (2, len(table_data)-5)),
                    ("SPAN", (0, len(table_data)-4), (2, len(table_data)-4)),
                    ("SPAN", (0, len(table_data)-3), (2, len(table_data)-3)),
                    ("SPAN", (0, len(table_data)-2), (2, len(table_data)-2)),
                    ("SPAN", (0, len(table_data)-1), (2, len(table_data)-1)),
                    ("FONTNAME", (3, len(table_data)-1), (3, len(table_data)-1), "Helvetica-Bold"),
                    ("TEXTCOLOR", (3, len(table_data)-1), (4, len(table_data)-1), PRIMARY),
                    ("ALIGN", (4, len(table_data)-5), (4, len(table_data)-1), "RIGHT"),
                    ("BOTTOMPADDING", (0,0), (-1,0), 8),
                    ("TOPPADDING", (0,0), (-1,0), 8),
                ]))
                elems.append(product_table)
                elems.append(Spacer(1, 24))

                # Footer notes & bank details
                notes = []
                if reference:
                    notes.append(Paragraph(f"<b>Referentie:</b> {reference}", normal))
                notes.append(Paragraph("Payment term: {} days.".format(payment_days), normal))
                notes.append(Paragraph("Returns allowed within 15 days after receiving the item.", small))
                notes.append(Spacer(1,6))
                notes.append(Paragraph("Classic Suzuki Parts NL — IBAN: NL49 RABO 0372 0041 64", small))
                notes.append(Paragraph("VATnumber 8077 51 911 B01 | C.O.C.number 01018576", small))
                for el in notes:
                    elems.append(el)

                # Build PDF
                doc.build(elems)
                buffer.seek(0)

                # Download button
                st.download_button(
                    label="Download factuur (PDF)",
                    data=buffer,
                    file_name=file_name if file_name.lower().endswith(".pdf") else file_name + ".pdf",
                    mime="application/pdf"
                )
                st.success("PDF gegenereerd.")
            except Exception as e:
                st.error(f"Fout bij PDF generatie: {e}")
