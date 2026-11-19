import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.drawing.image import Image as XLImage
import os
from datetime import date
from PIL import Image as PILImage
import io

st.set_page_config(page_title="ECA HO Sheet Generator", layout="wide")
st.title("ECA HO Sheet Generator")

# ---- BASE_DIR se path pakka ----
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE_DIR, "HO_template.xlsx")

TMP_DIR = os.path.join(BASE_DIR, "tmp_images")
os.makedirs(TMP_DIR, exist_ok=True)

if not os.path.exists(TEMPLATE):
    st.error(f"Template nahi mila: {TEMPLATE}\nApni Excel ko 'HO_template.xlsx' naam de kar is folder me rakho:\n{BASE_DIR}")
    st.stop()

# ---- INC snip ki starting row ----
INC_START_ROW = 24

# ---------- FORM ----------
with st.form("ho_form"):
    st.subheader("Shift Details")
    col1, col2, col3 = st.columns(3)
    shift_date = col1.date_input("Date", value=date.today())
    shift = col2.selectbox("Shift", ["Morning", "Evening", "Night"])
    engineer = col3.text_input("Engineer Name", "Vipul")

    st.subheader("Counts")
    c1, c2 = st.columns(2)
    inc_internal = c1.text_input("Incidents - Internal TT Count", "00")
    inc_external = c2.text_input("Incidents - External TT Count", "00")
    tc_internal = c1.text_input("Ticket Creation - Internal TT", "00")
    tc_external = c2.text_input("Ticket Creation - External TT", "00")
    service_req = st.text_input("Service Request", "HO")
    crq_external = c1.text_input("External CRQ", "00")
    crq_internal = c2.text_input("Internal CRQ", "00")

    st.subheader("Status Fields")
    s1, s2, s3 = st.columns(3)
    ageing = s1.text_input("Ageing Tracker", "Nil")
    ims_hc = s2.text_input("IMS HC", "Nil")
    vas_hc = s3.text_input("VAS HC", "Nil")
    ip_hc = s1.text_input("IP HC", "Nil")
    tools = s2.text_input("Tools (ITSM, ENM, OMC)", "checked")
    traffic = s3.text_input("Traffic Files", "Nil")

    st.subheader("HO Sheet Snips")
    inc_imgs = st.file_uploader(
        "ECA HO Incidents (open INC) ka snip - row 24 se",
        type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="inc"
    )
    console_imgs = st.file_uploader(
        "Incident Console snip (incidents ke neeche)",
        type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="console"
    )
    wo_imgs = st.file_uploader(
        "WO snip (Incident Console ke neeche)",
        type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="wo"
    )

    st.subheader("Snips (Screenshots) - alag sheets")
    sup_imgs = st.file_uploader(
        "Supervision Alarm ke snips (ek ya zyada)",
        type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="sup"
    )
    crq_imgs = st.file_uploader(
        "CRQ ke snips (ek ya zyada)",
        type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="crq"
    )

    submitted = st.form_submit_button("Generate HO Sheet")

# ---------- HELPER: image daalna (FULL QUALITY) ----------
def add_images_to_sheet(ws, uploaded_files, start_row=2, start_col="A",
                         display_width=700, prefix=""):
    """Image full pixels (original quality) me daalti hai.
    Sirf Excel me DISPLAY size chhoti hoti hai, quality kam nahi hoti."""
    anchor_row = start_row

    if not uploaded_files:
        return anchor_row

    for up in uploaded_files:
        try:
            img_bytes = up.getvalue()
            pil = PILImage.open(io.BytesIO(img_bytes))

            # Transparency fix (PNG/palette)
            if pil.mode in ("RGBA", "P"):
                bg = PILImage.new("RGB", pil.size, (255, 255, 255))
                if pil.mode == "RGBA":
                    bg.paste(pil, mask=pil.split()[3])
                else:
                    bg.paste(pil.convert("RGBA"))
                pil = bg

            # Full pixels rakho - koi resize NAHI
            w, h = pil.size

            buf = io.BytesIO()
            pil.save(buf, format="PNG")  # lossless, poora pixel data
            buf.seek(0)

            xl_img = XLImage(buf)

            # Sirf DISPLAY size chhota, aspect ratio maintain
            if w > display_width:
                scale = display_width / w
                xl_img.width = display_width
                xl_img.height = int(h * scale)
            else:
                xl_img.width = w
                xl_img.height = h

            ws.add_image(xl_img, f"{start_col}{anchor_row}")

            # Row advance - display height ke hisaab se
            anchor_row += int(xl_img.height / 18) + 3

        except Exception as e:
            st.error(f"[{prefix}] ERROR: {e}")

    return anchor_row

# ---------- GENERATION ----------
if submitted:
    d = shift_date.strftime("%d_%b_%Y")
    shift_short = shift

    fname = f"ECA HO sheet {d}_{shift_short}_shift-{engineer}.xlsx"

    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["HO"]

    std_font = Font(name="Ericsson Hilda", size=9)
    wrap = Alignment(wrap_text=True, vertical="top")

    def set_cell(ref, value):
        ws[ref] = value
        ws[ref].font = std_font
        ws[ref].alignment = wrap

    # ----- HO sheet data -----
    set_cell("A1", f"{shift} Shift")
    set_cell("C3", f"Internal TT Count: {inc_internal}\nExternal TT Count:{inc_external}")
    set_cell("C4", f"Internal TT Count: {tc_internal}\nExternal TT Count: {tc_external}")
    set_cell("C5", service_req)
    set_cell("C6", f"External CRQ :{crq_external}  Internal CRQ : {crq_internal}")
    set_cell("C7", ageing)
    set_cell("C8", ims_hc)
    set_cell("C9", vas_hc)
    set_cell("C10", ip_hc)
    set_cell("C11", tools)
    set_cell("C12", traffic)

    # ----- HO sheet: INC snip -----
    next_row = INC_START_ROW

    if inc_imgs:
        next_row = add_images_to_sheet(
            ws, inc_imgs, start_row=INC_START_ROW,
            start_col="A", display_width=1400, prefix="inc"
        )

    # ----- Incident Console -----
    if console_imgs:
        set_cell(f"A{next_row}", "Incident Console:")
        next_row = add_images_to_sheet(
            ws, console_imgs, start_row=next_row + 1,
            start_col="A", display_width=1400, prefix="console"
        )

    # ----- WO -----
    if wo_imgs:
        set_cell(f"A{next_row}", "WO:")
        next_row = add_images_to_sheet(
            ws, wo_imgs, start_row=next_row + 1,
            start_col="A", display_width=1400, prefix="wo"
        )

    # ----- Supervision Alarm -----
    if sup_imgs:
        ws_sup = wb["Supervision Alarm"]
        add_images_to_sheet(
            ws_sup, sup_imgs, start_row=2,
            start_col="B", display_width=1400, prefix="sup"
        )

    # ----- CRQ -----
    if crq_imgs:
        crq_name = next(
            (s for s in wb.sheetnames if s.strip() == "CRQ"), None
        )
        if crq_name:
            ws_crq = wb[crq_name]
            add_images_to_sheet(
                ws_crq, crq_imgs, start_row=2,
                start_col="B", display_width=700, prefix="crq"
            )

    # ----- Excel memory mein generate karo -----
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("HO Sheet generated successfully!")

    st.download_button(
        label="Download HO Sheet",
        data=output.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )