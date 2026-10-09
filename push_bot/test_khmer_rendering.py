import os
import json
import generate_report
import excel_to_image

print("Loading config...")
with open("config.json", encoding="utf-8") as f:
    cfg = json.load(f)

src_xlsx = "test_detail.xlsx"
ref_path = "post_office_lookup.csv"
out_dir = "test_out"

print(f"Source file: {src_xlsx}")
print(f"Ref path: {ref_path}")
print(f"Output dir: {out_dir}")

if not os.path.exists(src_xlsx):
    print(f"Error: {src_xlsx} does not exist.")
    exit(1)

try:
    print("Generating reports...")
    result = generate_report.generate_reports_from_data(
        src_xlsx, ref_path, out_dir, return_metadata=True, mode="wide"
    )
    print("Report generation metadata:")
    print("Overall counts:", result["overall_counts"])
    
    # Let's find one of the generated files to render
    handle_results = result["handle_results"]
    if not handle_results:
        print("No handle results found.")
        exit(1)
        
    first_hr = handle_results[0]
    print(f"First handle: {first_hr['handle']}")
    if not first_hr["handle_files"]:
        print("No files for first handle.")
        exit(1)
        
    first_file = first_hr["handle_files"][0]["path"]
    print(f"Rendering: {first_file}")
    
    img_buf = excel_to_image.excel_to_image(first_file)
    out_png = "khmer_render_test.png"
    with open(out_png, "wb") as f_img:
        f_img.write(img_buf.getvalue())
    print(f"SUCCESS: Rendered test image successfully saved to {out_png}")
except Exception as e:
    print(f"Error occurred: {e}")
    import traceback
    traceback.print_exc()
