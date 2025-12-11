# app.py
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PIL import Image

# -----------------------
# BASIC APP SETTINGS
# -----------------------
st.set_page_config(page_title="Invoice Generator", layout="wide")
st.title("Invoice Generator")

PRIMARY = colors.HexColor("#ED1B2D")
TEXT_DARK = colors.HexColor("#333333")
ROW_LIGHT = colors.HexColor("#F7F7F7")

VAT_PERCENT = 21.0  # fixed VAT


# -----------------------
# HELPERS
# -----------------------
def currency(value):
    try:
        return f"€ {float(value):,.2f}"
    except:
        return f"€ {value}"


def load_logo():
    try:
        img = Image.open("logo.png")
        bio = BytesIO()
        img.save(bio, format="PNG")
        bio.seek(0)
        return RLImage(bio, width=12*cm, height=5.2*cm, kind='proportional')
    except:
        return None


def read_excel_file(file):
    try:
        df = pd.read_excel(file)
        expected = ["Part Number", "Description", "Quantity", "Price"]
        if set(expected).issubset(df.columns):
            return df[expected]
        df = pd.read_excel(file, header=None, usecols=[0,1,2,3])
        df.columns = expected
        return df
    except:
        df = pd.read_excel(file, header=None, usecols=[0,1,2,3])
        df.columns = ["Part Number", "Description", "Quantity", "Price"]
        return df


# -----------------------
# STEP 1 — PRODUCT LIST
# -----------------------
st.header("Step 1 — Upload or Enter Items")

uploaded_files = st.file_uploader("Upload .xlsx invoice part lists", type=["xlsx"], accept_multiple_files=True)

if "combined_df" not in st.session_state:
    st.session_state["combined_df"] = pd.DataFrame(columns=["Part Number", "Description", "Quantity", "Price"])

if uploaded_files:
    dfs = []
    for f in uploaded_files:
        try:
            df = read_excel_file(f)
            dfs.append(df)
        except Exception as e:
            st.error(f"Error reading file {f.name}: {e}")

    if dfs:
        merged = pd.concat(dfs, ignore_index=True)
        merged["Description"] = merged["Description"].fillna("N/A")
        merged["Quantity"] = pd.to_numeric(merged["Quantity"], errors="coerce").fillna(0).astype(int)
        merged["Price"] = pd.to_numeric(merged["Price"], errors="coerce").fillna(0.0).astype(float)
        st.session_state["combined_df"] = merged
        st.success("Files combined successfully!")

st.write("Edit items below:")

# --- Sorting UI ---
sort_column = st.selectbox(
    "Sort by column",
    ["Part Number", "Description", "Quantity", "Price"]
)

ascending = st.radio(
    "Sort direction",
    ["Ascending", "Descending"],
    horizontal=True
) == "Ascending"

# Apply sorting
sorted_df = st.session_state["combined_df"].sort_values(
    sort_column,
    ascending=ascending
)

# Editable table
edited_df = st.data_editor(sorted_df, num_rows="dynamic", key="editor")


if st.button("Save Item List"):
    if edited_df.empty:
        st.error("The list is empty.")
    else:
        edited_df["Description"] = edited_df["Description"].fillna("N/A")
        edited_df["Quantity"] = pd.to_numeric(edited_df["Quantity"], errors='coerce').fillna(0).astype(int)
        edited_df["Price"] = pd.to_numeric(edited_df["Price"], errors='coerce').fillna(0.0).astype(float)
        st.session_state["edited_df"] = edited_df
        st.success("Item list saved!")


# -----------------------
# STEP 2 — INVOICE INFO
# -----------------------
st.header("Step 2 — Invoice Details")

colA, colB = st.columns(2)

with colA:
    invoice_number = st.text_input("Invoice Number", value=f"INV-{datetime.now().strftime('%Y%m%d%H%M')}")
    invoice_date = st.date_input("Invoice Date", value=datetime.today().date())
    supplier_number = st.text_input("Supplier Number", value="")
    shipping_cost = st.number_input("Shipping Cost (EUR)", value=20.00, step=0.50)

with colB:
    bill_to_name = st.text_input("Bill To", value="CMS")
    bill_to_address = st.text_area("Billing Address", value="Artemisweg 245\n8239 DD Lelystad\nNetherlands", height=100)
    reference = st.text_input("Reference", value="")
    file_name = st.text_input("PDF File Name", value=f"Invoice_{invoice_number}.pdf")


# -----------------------
# STEP 3 — PDF GENERATION
# -----------------------
st.header("Step 3 — Generate Invoice")

if "edited_df" not in st.session_state or st.session_state["edited_df"].empty:
    st.warning("Please save an item list first.")
else:
    df = st.session_state["edited_df"]
    df["Total"] = df["Quantity"] * df["Price"]

    subtotal = df["Total"].sum()
    total_ex_vat = subtotal + shipping_cost
    vat_amount = total_ex_vat * (VAT_PERCENT / 100)
    grand_total = total_ex_vat + vat_amount

    st.write("Preview:")
    st.dataframe(df)

    if st.button("Generate PDF Invoice"):
        try:
            buffer = BytesIO()
            doc = SimpleDocTemplate(
                buffer,
                pagesize=A4,
                rightMargin=1.5*cm,
                leftMargin=1.5*cm,
                topMargin=2.0*cm,
                bottomMargin=1.5*cm
            )

            styles = getSampleStyleSheet()
            title_style = ParagraphStyle("Title", parent=styles["Heading1"], fontSize=26, textColor=PRIMARY)
            normal = ParagraphStyle("Normal", parent=styles["Normal"], fontSize=10, textColor=TEXT_DARK)
            bold = ParagraphStyle("Bold", parent=styles["Normal"], fontSize=11, textColor=TEXT_DARK, fontName="Helvetica-Bold")

            elems = []

            # -----------------------
            # HEADER (Centered Logo + Left Company Info)
            # -----------------------

            logo = load_logo()

            # Centered Logo
            if logo:
                logo_table = Table(
                    [[logo]],
                    colWidths=[doc.width]
                )
                logo_table.setStyle([
                    ("ALIGN", (0,0), (-1,-1), "CENTER"),
                    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 12),
                ])
                elems.append(logo_table)
            else:
                elems.append(Paragraph("Company Logo Missing", bold))

            # Company Info underneath (left aligned)
            company_info_block = [
                Paragraph("Vlaandere Motoren - de Marne 136 B", normal),
                Paragraph("8701 MC - Bolsward", normal),
                Paragraph("Tel: +316-41484547", normal),
                Paragraph("IBAN: NL49 RABO 0372 0041 64", normal),
                Paragraph("VAT: 8077 51 911 B01 | C.O.C: 01018576", normal),
            ]

            company_table = Table([[company_info_block]], colWidths=[doc.width])
            company_table.setStyle([
                ("ALIGN", (0,0), (-1,-1), "LEFT"),
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("LEFTPADDING", (0,0), (-1,-1), 0),
                ("BOTTOMPADDING", (0,0), (-1,-1), 12),
            ])

            elems.append(company_table)

            # Divider line below
            elems.append(Table([[" "]], colWidths=[doc.width], style=[
                ("LINEBELOW", (0,0), (-1,0), 4, PRIMARY)
            ]))
            elems.append(Spacer(1, 15))


            # INVOICE & BILLING INFO
            bill_lines = [Paragraph(f"<b>{bill_to_name}</b>", bold)]
            for line in bill_to_address.split("\n"):
                bill_lines.append(Paragraph(line, normal))

            invoice_info = [
                ["Invoice Number:", invoice_number],
                ["Supplier Number:", supplier_number],
                ["Invoice Date:", invoice_date.strftime("%d-%m-%Y")],
                ["VAT:", f"{VAT_PERCENT}%"],
            ]


            inv_table = Table(invoice_info, colWidths=[4*cm, 6*cm])
            inv_table.setStyle([
                ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
                ("BOTTOMPADDING", (0,0), (-1,-1), 3),
            ])

            bill_table = Table([[bill_lines]], colWidths=[doc.width - 8*cm])

            # Combine billing info + invoice info side by side, top-aligned and same height
            side_by_side = Table(
                [
                    [
                        bill_table,   # Left block (Bill To)
                        inv_table     # Right block (Invoice Info)
                    ]
                ],
                colWidths=[doc.width * 0.55, doc.width * 0.45],
                style=[
                    ("VALIGN", (0,0), (-1,-1), "TOP"),       # << ensures equal height start
                    ("ALIGN", (0,0), (-1,-1), "LEFT"),
                    ("LEFTPADDING", (0,0), (-1,-1), 0),
                    ("RIGHTPADDING", (0,0), (-1,-1), 0),
                    ("TOPPADDING", (0,0), (-1,-1), 0),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 0),
                ]
            )

            elems.append(side_by_side)
            elems.append(Spacer(1, 18))

            # PRODUCT TABLE
            table_data = [["Part Number", "Description", "Qty", "Price", "Total"]]
            for _, item in df.iterrows():
                table_data.append([
                    str(item["Part Number"]),
                    str(item["Description"]),
                    str(item["Quantity"]),
                    currency(item["Price"]),
                    currency(item["Total"])
                ])

            table_data.append(["", "", "", "Subtotal", currency(subtotal)])
            table_data.append(["", "", "", "Shipping", currency(shipping_cost)])
            table_data.append(["", "", "", "Total Excl. VAT", currency(total_ex_vat)])
            table_data.append(["", "", "", f"VAT ({VAT_PERCENT}%)", currency(vat_amount)])
            table_data.append(["", "", "", "Grand Total", currency(grand_total)])

            col_widths = [3.5*cm, doc.width - (3.5+2+3+3)*cm, 2*cm, 3*cm, 3*cm]

            product_table = Table(table_data, colWidths=col_widths, repeatRows=1)
            product_table.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), PRIMARY),
                ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                ("ALIGN", (2,1), (-1,-1), "RIGHT"),
                ("GRID", (0,1), (-1,-5), 0.25, colors.lightgrey),
                ("ROWBACKGROUNDS", (0,1), (-1,-6), [colors.white, ROW_LIGHT]),
            ]))

            elems.append(product_table)
            elems.append(Spacer(1, 20))

            # FOOTER
            elems.append(Spacer(1, 15))
            elems.append(Paragraph("Payment term: 14 days.", normal))
            elems.append(Paragraph("Returns allowed within 15 days after receiving the item.", normal))
            elems.append(Paragraph("Vlaandere Motoren — IBAN: NL49 RABO 0372 0041 64", normal))
            elems.append(Paragraph("VATnumber 8077 51 911 B01 | C.O.C.number 01018576", normal))


            doc.build(elems)
            buffer.seek(0)

            st.download_button(
                label="Download PDF Invoice",
                data=buffer,
                file_name=file_name if file_name.endswith(".pdf") else file_name + ".pdf",
                mime="application/pdf"
            )

            st.success("PDF generated successfully!")

        except Exception as e:
            st.error(f"PDF generation error: {e}")
