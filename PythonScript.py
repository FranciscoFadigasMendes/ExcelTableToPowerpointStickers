import math
import time
import win32com.client
from typing import Tuple, Any

# ----------------------------------------------------------------------------- #
# SAFE EXCEL ACCESS
# ----------------------------------------------------------------------------- #

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

# ----------------------------------------------------------------------------- #
# IMPORTS
# ----------------------------------------------------------------------------- #

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

# ----------------------------------------------------------------------------- #
# CONFIGURATION
# ----------------------------------------------------------------------------- #

excel_file_path: str = r"C:\LOTO PLACARDS FC08\LOTO Updating Tool FCO8.xlsm"
ppt_file_url: str = r"C:\LOTO PLACARDS FC08\PLC30 - Inbound 1\03. LOTO Information Sticker\PLC30 LOTO Information 114 Stickers_Template.pptx"

stickers_per_slide: int = 6
total_stickers: int = 120

FORCE_COORDS: bool = False  # if True, apply computed left/top/width/height

# Layout for A4 (approx, points)
BASE_LEFTS = [2, 507]      # left X for column 0 and 1 (points)
BASE_TOPS = [63, 245, 420]  # top Y for rows 0..2 (points)

POINT_SIZE: Tuple[float, float] = (450.43, 34.02)
LOTO_SIZE: Tuple[float, float] = (32.03, 41.10)
CABINET_SIZE: Tuple[float, float] = (134.12, 21.83)

POINT_FONT_SIZE = 20
LOTO_FONT_SIZE = 22
CABINET_FONT_SIZE = 10

# ----------------------------------------------------------------------------- #
# HELPERS
# ----------------------------------------------------------------------------- #

import shutil, os, win32com.client

genpy = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp", "gen_py")
if os.path.exists(genpy):
    print("🧹 Clearing old gen_py cache...")
    shutil.rmtree(genpy, ignore_errors=True)

win32com.client.gencache.Rebuild()

def attach_office(app_name: str):
    """Attach to a running COM application or start it using EnsureDispatch."""
    try:
        return win32.GetActiveObject(f"{app_name}.Application")
    except Exception:
        try:
            app = win32.gencache.EnsureDispatch(f"{app_name}.Application")
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
    """Return shape by name, or None if not present."""
    try:
        return slide.Shapes(name)
    except Exception:
        return None

# ----------------------------------------------------------------------------- #
# MAIN
# ----------------------------------------------------------------------------- #

def main() -> None:

    win32com.client.gencache.Rebuild()

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
            print(f"⚠️ Slide {slide_index} missing for sticker {sticker_str}")
            continue

        slide = ppt_pres.Slides(slide_index)
        pos_in_slide = ((sticker_index - 1) % stickers_per_slide) + 1  # 1..6

        # POINTS (Point ##.01 .. Point ##.04) from Excel cols I..L (9..12)
        for col, point_num in zip(range(9, 13), range(1, 5)):
            idx_str = f"{point_num:02d}"  # 01..04
            name = f"Point {sticker_str}.{idx_str}"
            shp = safe_shape(slide, name)
            if shp is None:
                print(f"⚠️ Missing shape: {name} on slide {slide_index}")
                continue

            val = safe_cell(ws, data_row, col)

            # --- CLEAN VALUE ---
            val_text = ""
            if val is not None and str(val).strip() not in ["", "nan", "None"]:
                # Convert numbers like 1.0 → 1, 2.50 → 2.5
                if isinstance(val, (int, float)):
                    if float(val).is_integer():
                        val_text = str(int(val))
                    else:
                        val_text = str(round(val, 2))
                else:
                    val_text = str(val).strip()

            print(f"  ✅ Set {name} (row {data_row}, col {col}) → \"{val_text}\"")

            # Skip overwriting vertical text boxes
            try:
                if shp.TextFrame.Orientation not in (3, 4):
                    shp.TextFrame.TextRange.Text = val_text
            except Exception:
                pass

            if FORCE_COORDS:
                # Widen the cabinet text box to avoid text wrapping
                left, top, width, height = coords_for_position(pos_in_slide, CABINET_SIZE)
                apply_coords(shp, (left, top, width + 40, height))  # add ~40 points width change this if needed

        # LOTO Amount (Column M = 13)
        name = f"LOTO Amount {sticker_str}"
        shp = safe_shape(slide, name)
        if shp is None:
            print(f"⚠️ Missing shape: {name} on slide {slide_index}")
        else:
            val = safe_cell(ws, data_row, 13)
            val_text = ""
            if val is not None and str(val).strip() not in ["", "nan", "None"]:
                if isinstance(val, (int, float)):
                    if float(val).is_integer():
                        val_text = str(int(val))
                    else:
                        val_text = str(round(val, 2))
                else:
                    val_text = str(val).strip()

            print(f"  ✅ Set {name} (row {data_row}, col 13) → \"{val_text}\"")

            try:
                if shp.TextFrame.Orientation not in (3, 4):
                    shp.TextFrame.TextRange.Text = val_text
            except Exception:
                pass

            if FORCE_COORDS:
                apply_coords(shp, coords_for_position(pos_in_slide, LOTO_SIZE))
                apply_font_size(shp, LOTO_FONT_SIZE)

        # Cabinet (Column N = 14)
        name = f"Cabinet {sticker_str}"
        shp = safe_shape(slide, name)
        if shp is None:
            print(f"⚠️ Missing shape: {name} on slide {slide_index}")
        else:
            val = safe_cell(ws, data_row, 14)
            val_text = ""
            if val is not None and str(val).strip() not in ["", "nan", "None"]:
                if isinstance(val, (int, float)):
                    if float(val).is_integer():
                        val_text = str(int(val))
                    else:
                        val_text = str(round(val, 2))
                else:
                    val_text = str(val).strip()

            print(f"  ✅ Set {name} (row {data_row}, col 14) → \"{val_text}\"")

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
