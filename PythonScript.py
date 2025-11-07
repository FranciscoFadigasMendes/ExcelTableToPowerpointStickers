import math

# -----------------------------------------------------------------------------

import time

def safe_cell(ws, row, col, retries=5, delay=0.2):
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
    raise RuntimeError(f"Excel did not respond after {retries} retries for cell ({row}, {col})")

# -----------------------------------------------------------------------------

from typing import Tuple, Any

# -----------------------------------------------------------------------------

try:
    import win32com.client as win32  # type: ignore
except Exception:
    # Clear runtime error if pywin32 not installed
    class _Win32Stub:
        def __getattr__(self, name):
            raise RuntimeError(
                "pywin32 (win32com) is required. Install with: pip install pywin32"
            )
    win32 = _Win32Stub()

# --- Configuration (edit) ---------------------------------------------------

# Excel workbook (set "| None" if you want to use the active workbook)
excel_file_path: str = r"C:\LOTO PLACARDS FC08\LOTO Updating Tool FCO8.xlsm"

# PowerPoint file (can be local path or a URL if PowerPoint can open it)
ppt_file_url: str = r"C:\LOTO PLACARDS FC08\PLC30 - Inbound 1\03. LOTO Information Sticker\PLC30 LOTO Information 114 Stickers_Template.pptx"

stickers_per_slide: int = 6
total_stickers: int = 120

# Force coordinates/size and font sizes (toggle)
FORCE_COORDS: bool = False  # if True, apply computed left/top/width/height

# 1 point = 0.0352778 centimeters
# Left X for Column 0 and 1 --> Position: Horizontal= 0,07 cm = 0.002469446 points || Horizontal= 17,85 cm = 0.62970873 points 
# Top Y for rows 0..2 (points) --> Position: Vertical 0 = 2,17 cm = 0.076552826 points || Vertical 1 = 8,46 cm = 0.298450188 points || Vertical 2 = 14,79 cm = 0.521758662 points

# Layout: 2 columns x 3 rows (positions 1..6)
BASE_LEFTS = [1.98, 506.78]      # left X for column 0 and 1 (points)
BASE_TOPS = [61.80, 239.27, 419.15]  # top Y for rows 0..2 (points)

# Width/Height for each shape type (points)
POINT_SIZE: Tuple[float, float] = (450.43, 34.02)
LOTO_SIZE: Tuple[float, float] = (32.03, 41.10)
CABINET_SIZE: Tuple[float, float] = (134.12, 21.83)

# Font sizes (points)
POINT_FONT_SIZE = 28.06   # ~28pt
LOTO_FONT_SIZE = 22.11    # ~22pt
CABINET_FONT_SIZE = 11.91 # ~12pt


# ----------------------------------------------------------------------------

def attach_office(app_name: str):
    """
    Attach to a running COM application or start it using EnsureDispatch (generated wrappers).
    Returns a COM application object.
    """
    try:
        return win32.GetActiveObject(f"{app_name}.Application")
    except Exception:
        # Use generated wrappers to avoid AttributeError on dynamic dispatch
        try:
            app = win32.gencache.EnsureDispatch(f"{app_name}.Application")
            # make visible so user can see errors/auth prompts (optional)
            app.Visible = True
            return app
        except Exception as e:
            raise RuntimeError(f"Failed to start/attach to {app_name}: {e}")


def coords_for_position(pos_index: int, size: Tuple[float, float]) -> Tuple[float, float, float, float]:
    idx = pos_index - 1
    col = idx % 2
    row = idx // 2
    left = BASE_LEFTS[col]
    top = BASE_TOPS[row]
    width, height = size
    return (left, top, width, height)

def apply_coords(shp, coords):
    left, top, width, height = coords
    shp.Left = left
    shp.Top = top
    shp.Width = width
    shp.Height = height

def apply_font_size(shp: Any, size_pt: int) -> None:
    """Apply font size to the shape text range if possible."""
    try:
        tf = shp.TextFrame
        if tf is not None:
            rng = tf.TextRange
            if rng is not None and hasattr(rng, "Font"):
                rng.Font.Size = size_pt
    except Exception:
        pass

def safe_shape(slide: Any, name: str):
    """Return shape by name, or None if not present (avoid exceptions)."""
    try:
        return slide.Shapes(name)
    except Exception:
        return None
    
# ----------------------------------------------------------------------------

import time
import math

def main() -> None:
    # Attach to Excel
    excel = attach_office("Excel")

    # Open or get workbook
    try:
        if excel_file_path:
            wb = excel.Workbooks.Open(excel_file_path)
        else:
            wb = excel.ActiveWorkbook
    except Exception as e:
        raise RuntimeError(
            "Failed to open workbook. If it's on SharePoint/OneDrive, "
            "sync locally or ensure Office is authenticated. " + str(e)
        )

    ws = wb.Sheets("Info_Tags_PLC30_FCO8")

    # Attach to PowerPoint and open presentation
    ppt = attach_office("PowerPoint")
    try:
        ppt_pres = ppt.Presentations.Open(ppt_file_url, WithWindow=True)
    except Exception as e:
        raise RuntimeError(
            "Failed to open PowerPoint file. Ensure path/URL is reachable by PowerPoint: " + str(e)
        )

    # Iterate rows and update shapes
    for data_row in range(3, total_stickers + 3):  # rows 3 .. total_stickers+2
        sticker_index = data_row - 2
        sticker_str = f"{sticker_index:02d}"
        slide_index = math.ceil(sticker_index / stickers_per_slide)

        if slide_index > ppt_pres.Slides.Count:
            print(f"Slide {slide_index} missing for sticker {sticker_str}")
            continue

        slide = ppt_pres.Slides(slide_index)
        pos_in_slide = ((sticker_index - 1) % stickers_per_slide) + 1  # 1..6

        # POINTS (Point ##.01 .. Point ##.04) from Excel cols J..M (10..13)
        for col in range(10, 14):
            idx_str = f"{col - 9:02d}"  # 01..04
            name = f"Point {sticker_str}.{idx_str}"
            shp = safe_shape(slide, name)
            if shp is None:
                print(f"Missing shape: {name} on slide {slide_index}")
                continue

            # Read value safely
            try:
                val = ws.Cells(data_row, col).Value
            except Exception:
                time.sleep(0.1)
                val = ws.Cells(data_row, col).Value

            if val is None:
                val_text = ""
            else:
                val_text = str(val).strip()

            # Skip overwriting rotated text boxes (vertical)
            try:
                orientation = shp.TextFrame.Orientation
            except Exception:
                orientation = 1  # Default horizontal

            if orientation not in (3, 4):  # Skip vertical orientations
                shp.TextFrame.TextRange.Text = val_text

            # Apply coords if forced
            if FORCE_COORDS:
                apply_coords(shp, coords_for_position(pos_in_slide, POINT_SIZE))

            apply_font_size(shp, POINT_FONT_SIZE)

        # LOTO Amount (Column N = 14)
        name = f"LOTO Amount {sticker_str}"
        shp = safe_shape(slide, name)
        if shp is None:
            print(f"Missing shape: {name} on slide {slide_index}")
        else:
            try:
                val = ws.Cells(data_row, 14).Value
            except Exception:
                time.sleep(0.1)
                val = ws.Cells(data_row, 14).Value

            val_text = "" if val is None else str(val).strip()
            try:
                if shp.TextFrame.Orientation not in (3, 4):
                    shp.TextFrame.TextRange.Text = val_text
            except Exception:
                pass

            if FORCE_COORDS:
                apply_coords(shp, coords_for_position(pos_in_slide, LOTO_SIZE))
            apply_font_size(shp, LOTO_FONT_SIZE)

        # Cabinet (Column O = 15)
        name = f"Cabinet {sticker_str}"
        shp = safe_shape(slide, name)
        if shp is None:
            print(f"Missing shape: {name} on slide {slide_index}")
        else:
            try:
                val = ws.Cells(data_row, 15).Value
            except Exception:
                time.sleep(0.1)
                val = ws.Cells(data_row, 15).Value

            val_text = "" if val is None else str(val).strip()
            try:
                if shp.TextFrame.Orientation not in (3, 4):
                    shp.TextFrame.TextRange.Text = val_text
            except Exception:
                pass

            if FORCE_COORDS:
                apply_coords(shp, coords_for_position(pos_in_slide, CABINET_SIZE))
            apply_font_size(shp, CABINET_FONT_SIZE)

    print("✅ LOTO stickers updated successfully!")


if __name__ == "__main__":
    main()
