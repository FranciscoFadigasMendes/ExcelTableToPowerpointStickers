import win32com
import win32com.client as win32
import math

excel_file_path: str = r"C:\LOTO PLACARDS FC08\LOTO Updating Tool FCO8.xlsm"
ppt_file_url: str = r"C:\LOTO PLACARDS FC08\PLC20\02. Energy Point Tags\PLC20_LOTO_EnergyTags_250_Template_Italian.pptx"

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

def update_energy_tags():
    ppt_path = r"C:\LOTO PLACARDS FC08\PLC20\02. Energy Point Tags\PLC20_LOTO_EnergyTags_250_Template_Italian.pptx"
    
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

    ws = wb.Sheets("Info_Tags_PLC20_FCO8_147")

    # Launch PowerPoint
    try:
        ppt = win32.GetActiveObject("PowerPoint.Application")
    except:
        ppt = win32.Dispatch("PowerPoint.Application")
    
    ppt.Visible = True
    presentation = ppt.Presentations.Open(ppt_path)

    # Loop through 251 values
    for i in range(1, 251):
        shape_name = f"LOTO {i:02d}"
        cell_value = sheet.Range(f"G{i+1}").Value
        
        slide_number = math.ceil(i / 10)  # 10 per slide
        
        if slide_number <= presentation.Slides.Count:
            slide = presentation.Slides(slide_number)

            try:
                shape = slide.Shapes(shape_name)
                shape.TextFrame.TextRange.Text = str(cell_value)
            except Exception as e:
                print(f"Shape not found: {shape_name} on slide {slide_number}")
        else:
            print(f"Slide {slide_number} does not exist.")

    print("250 Energy Point Stickers Updated!")

if __name__ == "__main__":
    update_energy_tags()
