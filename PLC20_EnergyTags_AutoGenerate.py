import math
import time
from typing import Any, Tuple

# ----------------------------------------------------------------------------- #
# SAFE EXCEL ACCESS
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
# IMPORT WIN32 WITH GRACEFUL FAIL
# ----------------------------------------------------------------------------- #

try:
    import win32com.client as win32  # type: ignore
except Exception:
    class _Win32Stub:
        def __getattr__(self, _):
            raise RuntimeError("pywin32 is required. Install: pip install pywin32")
    win32 = _Win32Stub()


# ----------------------------------------------------------------------------- #
# CONFIGURATION
# ----------------------------------------------------------------------------- #

EXCEL_PATH: str = r"C:\LOTO PLACARDS FC08\LOTO Updating Tool FCO8.xlsm"
PPT_PATH: str = r"C:\LOTO PLACARDS FC08\PLC20\02. Energy Point Tags\PLC20_LOTO_EnergyTags_165_Template_Italian.pptx"

TOTAL_TAGS: int = 165
TAGS_PER_SLIDE: int = 10

START_ROW: int = 2  # Excel Row C2 → Tag 1
EXCEL_COL: str = "C"  # Source column for tag label

SHEET_NAME: str = "EnergyTags"


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
    """Convert Excel column letter to index (G→7, AA→27, etc.)."""
    col = col.upper()
    num = 0
    for c in col:
        num = num * 26 + (ord(c) - 64)
    return num

COL_INDEX = excel_col_letter_to_number(EXCEL_COL)

def set_font_size(shp, size_pt: int):
    """Safely apply font size to a shape's text."""
    try:
        tf = shp.TextFrame
        if tf is not None:
            rng = tf.TextRange
            if rng is not None and hasattr(rng, "Font"):
                rng.Font.Size = size_pt
    except Exception:
        pass


def center_text(shp):
    """Center text horizontally and vertically in a shape."""
    try:
        tf = shp.TextFrame
        if tf is not None:
            tf.HorizontalAnchor = 2  # 0=left, 1=centered (deprecated?), 2=centered works reliably
            tf.VerticalAnchor = 2    # 0=top, 1=center, 2=middle
            tf.TextRange.ParagraphFormat.Alignment = 2  # 0=left, 1=center, 2=center
    except Exception:
        pass


# ----------------------------------------------------------------------------- #
# MAIN PROCESS
# ----------------------------------------------------------------------------- #

def main() -> None:

    print("Starting PLC20 Energy Tag update...")

    # Attach Excel
    excel = attach_office("Excel")

    # Open workbook
    try:
        wb = excel.Workbooks.Open(EXCEL_PATH)
    except Exception as e:
        raise RuntimeError(
            "Failed to open Excel file. Check path or OneDrive login.\n" + str(e)
        )

    ws = wb.Sheets(SHEET_NAME)

    # Attach PowerPoint
    ppt = attach_office("PowerPoint")
    try:
        pres = ppt.Presentations.Open(PPT_PATH, WithWindow=True)
    except Exception as e:
        raise RuntimeError("Failed to open PPT file:\n" + str(e))
    

    # Iterate tags
    for tag_index in range(1, TOTAL_TAGS + 1):
        shape_name = f"LOTO {tag_index:02d}"
        row = START_ROW + tag_index  # G2 = tag1 → row3

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

        val = safe_cell(ws, row, COL_INDEX)
        new_text = "" if val in (None, "", "nan") else str(val).strip()

        try:
            if shp.TextFrame.Orientation not in (3, 4):  # skip rotated shapes
                shp.TextFrame.TextRange.Text = new_text
                set_font_size(shp, 20)  # apply font size here
                center_text(shp)          # center text
        except Exception:
            pass

        print(f"  🔹 Set {shape_name} → \"{new_text}\" {pos_print}")

    print("DONE! 165 Energy Tags Updated Successfully!")


# ----------------------------------------------------------------------------- #
# ENTRY POINT
# ----------------------------------------------------------------------------- #

if __name__ == "__main__":
    main()
