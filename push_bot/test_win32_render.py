import os
import time
import win32com.client
from PIL import ImageGrab

def excel_to_image_win32(xlsx_path, out_png_path):
    abs_path = os.path.abspath(xlsx_path)
    print(f"Opening Excel with file: {abs_path}")
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = None
    try:
        wb = excel.Workbooks.Open(abs_path)
        ws = wb.ActiveSheet
        print("Copying used range as picture...")
        ws.UsedRange.CopyPicture(1, 2)
        time.sleep(0.5) # wait for clipboard
        
        print("Getting image from clipboard...")
        img = None
        for i in range(5):
            img = ImageGrab.grabclipboard()
            if img:
                print(f"Image retrieved from clipboard on attempt {i+1}!")
                break
            time.sleep(0.5)
            
        if not img:
            raise ValueError("Clipboard was empty or copy failed.")
            
        print(f"Saving image to: {out_png_path}")
        img.save(out_png_path, "PNG")
        return True
    except Exception as e:
        print(f"Failed: {e}")
        return False
    finally:
        if wb:
            wb.Close(SaveChanges=False)
        excel.Quit()

xlsx_path = r"test_out\Report_SVAP001_Pending_14_07_2026_000000.xlsx"
out_png = "svap001_pending_win32_test.png"
excel_to_image_win32(xlsx_path, out_png)
