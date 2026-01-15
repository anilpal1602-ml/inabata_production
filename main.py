import os
import shutil
import uuid
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

# Import the new function we just created
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

@app.post("/api/process-docs")
async def process_documents(
    invoice: UploadFile = File(...), 
    packing_list: UploadFile = File(...)
):
    try:
        # 1. Create unique filenames for this specific request
        job_id = str(uuid.uuid4())[:8]
        os.makedirs(INPUT_DIR, exist_ok=True)
        
        invoice_path = os.path.join(INPUT_DIR, f"{job_id}_inv.pdf")
        pl_path = os.path.join(INPUT_DIR, f"{job_id}_pl.pdf")

        # 2. Save the uploaded files to disk
        with open(invoice_path, "wb") as buffer:
            shutil.copyfileobj(invoice.file, buffer)
        with open(pl_path, "wb") as buffer:
            shutil.copyfileobj(packing_list.file, buffer)

        # 3. Trigger the Pipeline with the saved files
        # This calls the Step 1, 2, and 3 logic in sequence
        final_excel_path = run_custom_pipeline(invoice_path, pl_path)

        if not os.path.exists(final_excel_path):
            raise HTTPException(status_code=500, detail="Pipeline failed to create output.")

        # 4. Return the final Excel file to the React UI
        return FileResponse(
            path=final_excel_path, 
            filename=os.path.basename(final_excel_path),
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        print(f"Error: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    # Run from terminal with: python main.py
    uvicorn.run(app, host="0.0.0.0", port=8000)