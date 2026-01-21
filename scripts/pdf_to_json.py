import json
import os
import re  # Added for filename sanitization
from dotenv import load_dotenv
from google import genai
from google.genai import types

# --------------------------------------------------
# ENV SETUP
# --------------------------------------------------
load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY")
if not API_KEY:
    raise EnvironmentError("GEMINI_API_KEY not found in environment")

client = genai.Client(api_key=API_KEY)


# --------------------------------------------------
# PROMPT
# --------------------------------------------------
def build_prompt():
    return """
You are an expert Indonesian customs documentation officer (PIB / CEISA).

You will be provided with:
1) A Commercial Invoice PDF
2) A Packing List PDF

Your task:
Extract, normalize, and infer information from these documents and populate ALL sheets.

==================================================
MANDATORY OUTPUT STRUCTURE (NESTED LISTS)
==================================================
Return ONE valid JSON object. 
Each key is the Sheet Name. 
The value for each key must be a LIST OF LISTS. 
The first sub-list contains the Column Headers. 
Following sub-lists contain the actual data rows.

Example:
{
  "HEADER": [
    ["CIF", "BRUTO", "NETTO"],
    ["5000", "120", "100"]
  ],
  "BARANG": [
    ["HS", "KODE BARANG", "URAIAN"],
    ["85171300", "SAP001", "Smartphone A"],
    ["85171300", "SAP002", "Smartphone B"]
  ]
}

==================================================
SHEET DEFINITIONS AND COLUMNS
==================================================
- HEADER: CIF, BRUTO, NETTO, TANGGAL PERNYATAAN, KODE VALUTA
- ENTITAS: NAMA ENTITAS, ALAMAT ENTITAS
- DOKUMEN: SERI, NOMOR DOKUMEN, TANGGAL
- PENGANGKUT: NAMA PENGANGKUT
- BARANG: HS, KODE BARANG, URAIAN, KODE SATUAN, JUMLAH SATUAN, NETTO, CIF

==================================================
CRITICAL RULES
==================================================
- Date Format: YYYY-MM-DD
- Numeric values: Digits only (no commas or currency symbols)
- SAP CODE: Look for MA00041841 or similar at the bottom of descriptions and map to KODE BARANG.
- Multiple Rows: BARANG have a sub-list for EVERY item found and DOKUMEN should have 3 sub-list for each documents like Invoice, packing list and GRN. GRN entries should be same as invoice.
- Seri - If seri infomation is not there then use serial number starting from 1. 
- HS CODE: Look for 39094010 or similar at the bottom of descriptions and map to HS.
- NAMA PENGANGKUT will be like TRUCK etc
- In KODE SATUAN check for these entries {'ST', 'KGM', 'MTR','RO'}
- In KODE SATUAN if "SHT" then replace it with 'ST'
- In ENTITAS, Receipents of goods and only one row entry expected
- If missing: Use ""
- HEADER: BRUTO - Gross weight or quatity from Invoice 
- HEADER: NETTO - Net Weight or quantity from Invoice

"""


# --------------------------------------------------
# GEMINI EXTRACTION
# --------------------------------------------------
def extract_with_gemini(invoice_pdf, packing_pdf):
    # Validate paths
    if not os.path.exists(invoice_pdf):
        raise FileNotFoundError(f"Invoice PDF not found → {invoice_pdf}")
    if not os.path.exists(packing_pdf):
        raise FileNotFoundError(f"Packing List PDF not found → {packing_pdf}")

    # Upload PDF files to Gemini Files API
    invoice_file = client.files.upload(file=invoice_pdf, config={"mime_type": "application/pdf"})
    packing_file = client.files.upload(file=packing_pdf, config={"mime_type": "application/pdf"})

    # Generate structured content with uploaded file references
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            build_prompt(),       # user prompt text
            invoice_file,         # uploaded invoice
            packing_file,         # uploaded packing list
        ],
        config=types.GenerateContentConfig(
            temperature=0,
            response_mime_type="application/json",
        ),
    )

    return json.loads(response.text)


# --------------------------------------------------
# SAVE JSON (MODIFIED)
# --------------------------------------------------
def save_to_json(data, output_dir, filename=None):
    """
    Saves JSON data. 
    If filename is None, it extracts 'NOMOR DOKUMEN' from the data to name the file.
    """
    os.makedirs(output_dir, exist_ok=True)

    final_filename = filename

    # LOGIC: If no filename is forced, try to extract it from 'DOKUMEN' -> 'NOMOR DOKUMEN'
    if final_filename is None:
        try:
            # 1. Access DOKUMEN sheet
            dok_sheet = data.get("DOKUMEN", [])
            
            # 2. Check if we have data (Index 0 is Header, Index 1 is Data)
            if len(dok_sheet) > 1:
                # The prompt structure is ["SERI", "NOMOR DOKUMEN", "TANGGAL"]
                # So we want index 1 of the data row.
                first_data_row = dok_sheet[1]
                
                if len(first_data_row) > 1:
                    raw_doc_num = str(first_data_row[1]).strip()
                    
                    if raw_doc_num:
                        # 3. Sanitize filename (Replace / \ : * ? " < > | with underscores)
                        # Example: "INV/2024/001" -> "INV_2024_001.json"
                        safe_name = re.sub(r'[\\/*?:"<>|]', '_', raw_doc_num)
                        final_filename = f"{safe_name}.json"
        except Exception as e:
            print(f"⚠️ Warning: Could not extract document number for filename: {e}")

    # Fallback if extraction failed and no filename was provided
    if not final_filename:
        final_filename = "extracted_data.json"

    output_path = os.path.join(output_dir, final_filename)
    
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
        
    return output_path


# --------------------------------------------------
# ENTRY POINT (BLOCKED)
# --------------------------------------------------
if __name__ == "__main__":
    raise RuntimeError(
        "This module is pipeline-only. Use scripts.run_pipeline"
    )