# # import os
# # import yaml
# # import sys

# # from scripts.pdf_to_json import extract_with_gemini, save_to_json
# # from scripts.json_to_excel import json_to_excel
# # from scripts.excel_postprocess import process_customs_excel


# # # --------------------------------------------------
# # # CONFIG LOADER
# # # --------------------------------------------------
# # def load_config():
# #     config_path = os.path.join(
# #         os.path.dirname(__file__), "..", "config.yaml"
# #     )
# #     with open(config_path, "r") as f:
# #         return yaml.safe_load(f)


# # # --------------------------------------------------
# # # PATH RESOLVER
# # # --------------------------------------------------
# # def resolve(base, relative):
# #     return os.path.join(base, relative)


# # # --------------------------------------------------
# # # PIPELINE
# # # --------------------------------------------------
# # def run():
# #     print("\n--- Indonesian Customs Automation Pipeline ---")

# #     cfg = load_config()
# #     base_dir = cfg["base_dir"]

# #     # Inputs
# #     invoice_pdf = resolve(base_dir, cfg["data"]["input"]["invoice_pdf"])
# #     packing_pdf = resolve(base_dir, cfg["data"]["input"]["packing_pdf"])
    
# #     # Intermediate & References
# #     json_path = resolve(base_dir, cfg["data"]["intermediate"]["extracted_json"])
# #     template_excel = resolve(base_dir, cfg["data"]["templates"]["pib_template"])
# #     customer_ref = resolve(base_dir, cfg["data"]["reference"]["customer_list"])
# #     hs_code_ref = resolve(base_dir, cfg["data"]["reference"]["hs_code"])
# #     output_dir = resolve(base_dir, cfg["data"]["output"]["final_excel_dir"])

# #     os.makedirs(os.path.dirname(json_path), exist_ok=True)
# #     os.makedirs(output_dir, exist_ok=True)

# #     # =========================================================================
# #     # STEP 1: EXTRACT (SKIPPED)
# #     # =========================================================================
# #     print("\n[1/3] PDF → JSON (SKIPPED)")
# #     extracted = extract_with_gemini(invoice_pdf, packing_pdf)
# #     json_path = save_to_json(extracted, output_dir=os.path.dirname(json_path))

# #     # SAFETY CHECK: Ensure JSON exists before Step 2
# #     if not os.path.exists(json_path):
# #         print(f"\n❌ CRITICAL ERROR: JSON file not found at: {json_path}")
# #         print("   Since you skipped Step 1, you must have a previously generated JSON file.")
# #         print("   Please run Step 1 at least once.")
# #         sys.exit(1)

# #     # =========================================================================
# #     # STEP 2: GENERATE EXCEL
# #     # =========================================================================
# #     print("\n[2/3] JSON → EXCEL")
# #     print(f"   > Reading from: {os.path.basename(json_path)}")
# #     populated_excel = json_to_excel(json_path, template_excel)

# #     # =========================================================================
# #     # STEP 3: POST-PROCESS
# #     # =========================================================================
# #     print("\n[3/3] EXCEL POST-PROCESS")
# #     final_output = process_customs_excel(
# #         populated_excel,
# #         customer_ref,
# #         hs_code_ref
# #     )

# #     print("\n✅ PIPELINE COMPLETED")
# #     print(f"Final file → {final_output}")


# # # --------------------------------------------------
# # # ENTRY
# # # --------------------------------------------------
# # if __name__ == "__main__":
# #     run()







# # #####for fastapi compatible
# # import os
# # import yaml
# # import sys

# # # Change imports to relative or absolute based on your execution context
# # from scripts.pdf_to_json import extract_with_gemini, save_to_json
# # from scripts.json_to_excel import json_to_excel
# # from scripts.excel_postprocess import process_customs_excel

# # def load_config():
# #     config_path = os.path.join(os.path.dirname(__file__), "..", "config.yaml")
# #     with open(config_path, "r") as f:
# #         return yaml.safe_load(f)

# # def resolve(base, relative):
# #     return os.path.join(base, relative)

# # # --- NEW FUNCTION FOR FASTAPI ---
# # def run_custom_pipeline(invoice_pdf_path, packing_pdf_path):
# #     print("\n--- Triggering API Pipeline ---")
# #     cfg = load_config()
# #     base_dir = cfg["base_dir"]

# #     # Intermediate & References from Config
# #     json_dir = resolve(base_dir, "data/intermediate")
# #     template_excel = resolve(base_dir, cfg["data"]["templates"]["pib_template"])
# #     customer_ref = resolve(base_dir, cfg["data"]["reference"]["customer_list"])
# #     hs_code_ref = resolve(base_dir, cfg["data"]["reference"]["hs_code"])
# #     output_dir = resolve(base_dir, cfg["data"]["output"]["final_excel_dir"])

# #     # STEP 1: PDF → JSON
# #     extracted = extract_with_gemini(invoice_pdf_path, packing_pdf_path)
# #     json_path = save_to_json(extracted, output_dir=json_dir)

# #     # STEP 2: JSON → EXCEL
# #     populated_excel = json_to_excel(json_path, template_excel)

# #     # STEP 3: POST-PROCESS
# #     final_output_path = process_customs_excel(
# #         populated_excel,
# #         customer_ref,
# #         hs_code_ref
# #     )
# #     final_output_path = process_customs_excel(
# #     populated_excel,
# #     customer_ref,
# #     hs_code_ref
# # )

# #     # 🔥 STEP 4: Excel Compliance Fix
# #     final_output_path = fix_entitas_nomor_aju_to_text(final_output_path)

# #     return final_output_path

# # # --- New Pipeline with 4 script
# # import os
# # import yaml
# # import sys

# # # Change imports to relative or absolute based on your execution context
# # from scripts.pdf_to_json import extract_with_gemini, save_to_json
# # from scripts.json_to_excel import json_to_excel
# # from scripts.excel_postprocess import process_customs_excel

# # def load_config():
# #     config_path = os.path.join(os.path.dirname(__file__), "..", "config.yaml")
# #     with open(config_path, "r") as f:
# #         return yaml.safe_load(f)

# # def resolve(base, relative):
# #     return os.path.join(base, relative)

# # # --- NEW FUNCTION FOR FASTAPI ---
# # def run_custom_pipeline(invoice_pdf_path, packing_pdf_path):
# #     print("\n--- Triggering API Pipeline ---")
# #     cfg = load_config()
# #     base_dir = cfg["base_dir"]

# #     # Intermediate & References from Config
# #     json_dir = resolve(base_dir, "data/intermediate")
# #     template_excel = resolve(base_dir, cfg["data"]["templates"]["pib_template"])
# #     customer_ref = resolve(base_dir, cfg["data"]["reference"]["customer_list"])
# #     hs_code_ref = resolve(base_dir, cfg["data"]["reference"]["hs_code"])
# #     output_dir = resolve(base_dir, cfg["data"]["output"]["final_excel_dir"])

# #     # STEP 1: PDF → JSON
# #     extracted = extract_with_gemini(invoice_pdf_path, packing_pdf_path)
# #     json_path = save_to_json(extracted, output_dir=json_dir)

# #     # STEP 2: JSON → EXCEL
# #     populated_excel = json_to_excel(json_path, template_excel)

# #     # STEP 3: POST-PROCESS
# #     final_output_path = process_customs_excel(
# #         populated_excel,
# #         customer_ref,
# #         hs_code_ref
# #     )


# #     return final_output_path

# import os
# import yaml
# import sys

# # Change imports to relative or absolute based on your execution context
# from scripts.pdf_to_json import extract_with_gemini, save_to_json
# from scripts.json_to_excel import json_to_excel
# from scripts.excel_postprocess import process_customs_excel
# # --- NEW IMPORT ---
# from scripts.excel_fix import fix_entitas_nomor_aju_to_text

# def load_config():
#     config_path = os.path.join(os.path.dirname(__file__), "..", "config.yaml")
#     with open(config_path, "r") as f:
#         return yaml.safe_load(f)

# def resolve(base, relative):
#     return os.path.join(base, relative)

# # --- NEW FUNCTION FOR FASTAPI ---
# def run_custom_pipeline(invoice_pdf_path, packing_pdf_path):
#     print("\n--- Triggering API Pipeline ---")
#     cfg = load_config()
#     base_dir = cfg["base_dir"]

#     # Intermediate & References from Config
#     json_dir = resolve(base_dir, "data/intermediate")
#     template_excel = resolve(base_dir, cfg["data"]["templates"]["pib_template"])
#     customer_ref = resolve(base_dir, cfg["data"]["reference"]["customer_list"])
#     hs_code_ref = resolve(base_dir, cfg["data"]["reference"]["hs_code"])
#     output_dir = resolve(base_dir, cfg["data"]["output"]["final_excel_dir"])

#     # STEP 1: PDF → JSON
#     print("...Running Step 1: Extraction")
#     extracted = extract_with_gemini(invoice_pdf_path, packing_pdf_path)
#     json_path = save_to_json(extracted, output_dir=json_dir)

#     # STEP 2: JSON → EXCEL
#     print("...Running Step 2: Populating Excel")
#     populated_excel = json_to_excel(json_path, template_excel)

#     # STEP 3: POST-PROCESS (Calculations & Logic)
#     print("...Running Step 3: Post-Processing")
#     post_processed_excel = process_customs_excel(
#         populated_excel,
#         customer_ref,
#         hs_code_ref
#     )

#     # STEP 4: EXCEL FIX (Text Formatting for ENTITAS)
#     print("...Running Step 4: Formatting Fixes")
#     final_output_path = fix_entitas_nomor_aju_to_text(post_processed_excel)

#     print(f"Pipeline Complete. Output: {final_output_path}")
#     return final_output_path
import os
import yaml
import sys

# Change imports to relative or absolute based on your execution context
from scripts.pdf_to_json import extract_with_gemini, save_to_json
from scripts.json_to_excel import json_to_excel
from scripts.excel_postprocess import process_customs_excel
# --- NEW IMPORT ---
from scripts.excel_fix import fix_entitas_nomor_aju_to_text

def load_config():
    config_path = os.path.join(os.path.dirname(__file__), "..", "config.yaml")
    with open(config_path, "r") as f:
        return yaml.safe_load(f)

def resolve(base, relative):
    return os.path.join(base, relative)

# --- NEW FUNCTION FOR FASTAPI ---
# CHANGE: Added serial_number=None parameter
def run_custom_pipeline(invoice_pdf_path, packing_pdf_path, serial_number=None):
    print("\n--- Triggering API Pipeline ---")
    cfg = load_config()
    base_dir = cfg["base_dir"]

    # Intermediate & References from Config
    json_dir = resolve(base_dir, "data/intermediate")
    template_excel = resolve(base_dir, cfg["data"]["templates"]["pib_template"])
    customer_ref = resolve(base_dir, cfg["data"]["reference"]["customer_list"])
    hs_code_ref = resolve(base_dir, cfg["data"]["reference"]["hs_code"])
    output_dir = resolve(base_dir, cfg["data"]["output"]["final_excel_dir"])

    # STEP 1: PDF → JSON
    print("...Running Step 1: Extraction")
    extracted = extract_with_gemini(invoice_pdf_path, packing_pdf_path)
    json_path = save_to_json(extracted, output_dir=json_dir)

    # STEP 2: JSON → EXCEL
    print("...Running Step 2: Populating Excel")
    # CHANGE: Passed user_serial to json_to_excel
    populated_excel = json_to_excel(json_path, template_excel, user_serial=serial_number)

    # STEP 3: POST-PROCESS (Calculations & Logic)
    print("...Running Step 3: Post-Processing")
    post_processed_excel = process_customs_excel(
        populated_excel,
        customer_ref,
        hs_code_ref
    )

    # STEP 4: EXCEL FIX (Text Formatting for ENTITAS)
    print("...Running Step 4: Formatting Fixes")
    final_output_path = fix_entitas_nomor_aju_to_text(post_processed_excel)

    print(f"Pipeline Complete. Output: {final_output_path}")
    return final_output_path