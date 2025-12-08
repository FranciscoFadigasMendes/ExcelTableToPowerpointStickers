import math
import time
from typing import Any

# ----------------------------------------------------------------------------- #
# CONFIGURATION
# ----------------------------------------------------------------------------- #

EXCEL_PATH = r"C:\LOTO PLACARDS FC08\LOTO Updating Tool FCO8.xlsm"
PPT_PATH = r"C:\LOTO PLACARDS FC08\PLC50\02. Energy Point Tags\PLC50_LOTO_EnergyTags_133_Template_Italian.pptx"

SHEET_NAME = "EnergyTags"  # Adjust sheet name
TOTAL_TAGS = 133
TAGS_PER_SLIDE = 10
START_ROW = 2  # Excel row G2 = first tag
EXCEL_COL = "I"  # Column containing the tag values
FONT_SIZE = 20  # Desired font size


# ----------------------------------------------------------------------------- #
# SAFE EXCEL CELL ACCESS
# ----------------------------------------------------------------------------- #

def safe_cell(ws, row, col, retries=5, delay=0.25):
    """Safely get a cell value from Excel with retry if COM call is rejected."""
    for _ in range(retries):
        try:
            return ws.Cells(row, col).Value
        except Exception as e:
            if "Call was rejected by callee" in str(e):
                time.sleep(delay)
                continue
            else:
                raise
    raise RuntimeError(f"Excel failed COM access: Cell({row},{col})")


# ----------------------------------------------------------------------------- #
# IMPORT WIN32 COM
# ----------------------------------------------------------------------------- #

try:
    import win32com.client as win32
except Exception:
    class _Win32Stub:
        def __getattr__(self, _):
            raise RuntimeError("pywin32 required: pip install pywin32")
    win32 = _Win32Stub()


# ----------------------------------------------------------------------------- #
# HELPERS
# ----------------------------------------------------------------------------- #

def attach_office(app_name: str):
    """Attach to a running Office COM instance or launch one."""
    try:
        app = win32.GetActiveObject(f"{app_name}.Application")
    except Exception:
        app = win32.gencache.EnsureDispatch(f"{app_name}.Application")
    app.Visible = True
    return app

def safe_shape(slide: Any, name: str):
    try:
        return slide.Shapes(name)
    except Exception:
        return None

def excel_col_letter_to_number(col: str) -> int:
    """Convert Excel column letter to number (G -> 7)."""
    col = col.upper()
    num = 0
    for c in col:
        num = num * 26 + (ord(c) - 64)
    return num

def set_font_size(shp, size_pt: int):
    """Safely set font size of shape text."""
    try:
        tf = shp.TextFrame
        if tf is not None:
            rng = tf.TextRange
            if rng is not None and hasattr(rng, "Font"):
                rng.Font.Size = size_pt
    except Exception:
        pass

def center_text(shp):
    """Center text horizontally and vertically in the shape."""
    try:
        tf = shp.TextFrame
        if tf is not None:
            tf.HorizontalAnchor = 2  # center
            tf.VerticalAnchor = 2    # middle
            tf.TextRange.ParagraphFormat.Alignment = 2  # center
    except Exception:
        pass

def format_shape(shp, size_pt=FONT_SIZE):
    """Set font size and center text."""
    set_font_size(shp, size_pt)
    center_text(shp)


# ----------------------------------------------------------------------------- #
# MAIN
# ----------------------------------------------------------------------------- #

def main():
    print("Starting PLC50 Energy Tag update...")

    # Attach Excel
    excel = attach_office("Excel")

    # Open workbook
    try:
        wb = excel.Workbooks.Open(EXCEL_PATH)
    except Exception as e:
        raise RuntimeError("Failed to open Excel file: " + str(e))

    ws = wb.Sheets(SHEET_NAME)
    col_index = excel_col_letter_to_number(EXCEL_COL)

    # Attach PowerPoint
    ppt = attach_office("PowerPoint")
    try:
        pres = ppt.Presentations.Open(PPT_PATH, WithWindow=True)
    except Exception as e:
        raise RuntimeError("Failed to open PowerPoint file: " + str(e))

    # Iterate tags
    for tag_index in range(1, TOTAL_TAGS + 1):
        shape_name = f"LOTO {tag_index:02d}"
        row = START_ROW + tag_index
        slide_index = math.ceil(tag_index / TAGS_PER_SLIDE)
        pos_print = f"[Tag {tag_index:03d} / Slide {slide_index}]"

        if slide_index > pres.Slides.Count:
            print(f"⚠️ Slide {slide_index} missing — {pos_print}")
            continue

        slide = pres.Slides(slide_index)
        shp = safe_shape(slide, shape_name)

        if shp is None:
            print(f"⚠️ Missing shape \"{shape_name}\" {pos_print}")
            continue

        val = safe_cell(ws, row, col_index)
        new_text = "" if val in (None, "", "nan") else str(val).strip()

        try:
            if shp.TextFrame.Orientation not in (3, 4):  # skip vertical text
                shp.TextFrame.TextRange.Text = new_text
                format_shape(shp, FONT_SIZE)
        except Exception:
            pass

        print(f"  🔹 Set {shape_name} → \"{new_text}\" {pos_print}")

    print("DONE! 133 Energy Tags Updated Successfully!")
    print("Remember to SAVE your PowerPoint before closing.")


# ----------------------------------------------------------------------------- #
# ENTRY POINT
# ----------------------------------------------------------------------------- #

if __name__ == "__main__":
    main()
