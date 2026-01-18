import os
import shutil
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

from scripts.run_pipeline import run_custom_pipeline

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Root Directory Setup
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "data", "input")
STATE_DIR = os.path.join(BASE_DIR, "state")
SERIAL_FILE = os.path.join(STATE_DIR, "serial_tracker.txt")


def update_serial_tracker(serial_number: str):
    os.makedirs(STATE_DIR, exist_ok=True)
    with open(SERIAL_FILE, "w") as f:
        f.write(serial_number.strip())


@app.post("/api/process-docs")
async def process_documents(
    serial_number: str = Form(...),
    invoice: UploadFile = File(...),
    packing_list: UploadFile = File(...)
):
    try:
        # 1. Save serial number
        update_serial_tracker(serial_number)

        # 2. Prepare input directory
        os.makedirs(INPUT_DIR, exist_ok=True)

        inv_filename = os.path.basename(invoice.filename)
        pl_filename = os.path.basename(packing_list.filename)

        invoice_path = os.path.join(INPUT_DIR, inv_filename)
        pl_path = os.path.join(INPUT_DIR, pl_filename)

        # 3. Save uploaded files
        with open(invoice_path, "wb") as buffer:
            shutil.copyfileobj(invoice.file, buffer)

        with open(pl_path, "wb") as buffer:
            shutil.copyfileobj(packing_list.file, buffer)

        # 4. Run pipeline
        final_excel_path = run_custom_pipeline(invoice_path, pl_path)

        if not os.path.exists(final_excel_path):
            raise HTTPException(status_code=500, detail="Pipeline failed to create output.")

        # 5. Return output file
        return FileResponse(
            path=final_excel_path,
            filename=os.path.basename(final_excel_path),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        print("Error:", str(e))
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
