import os
import excel_to_image

xlsx_path = r"test_out\Report_SVAP001_Pending_14_07_2026_000000.xlsx"
out_png = "svap001_pending_test.png"

print(f"Loading and rendering: {xlsx_path}")
if not os.path.exists(xlsx_path):
    print("Error: file does not exist.")
    exit(1)

try:
    img_buf = excel_to_image.excel_to_image(xlsx_path)
    with open(out_png, "wb") as f:
        f.write(img_buf.getvalue())
    print("Successfully generated BATP001 Pending test image!")
except Exception as e:
    print(f"Error: {e}")
