from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import uuid
import os

from docx_processor import protect_docx, protect_xlsx

app = FastAPI(title="Document Template Locker")

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "temp"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Serve static files (HTML)
@app.get("/")
async def serve_ui():
    html_path = os.path.join(os.path.dirname(__file__), "index.html")
    return FileResponse(html_path, media_type="text/html")

@app.post("/lock-document/")
async def lock_document(file: UploadFile = File(...)):
    file_ext = file.filename.split('.')[-1].lower()
    
    if file_ext not in ["docx", "xlsx", "xls"]:
        return {"error": "Only .docx, .xlsx, and .xls files are supported"}
    
    input_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}.{file_ext}")
    output_path = os.path.join(UPLOAD_DIR, f"locked_{uuid.uuid4()}.{file_ext}")

    try:
        with open(input_path, "wb") as f:
            f.write(await file.read())

        if file_ext in ["xlsx", "xls"]:
            protect_xlsx(
                input_path=input_path,
                output_path=output_path
            )
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else:  # docx
            protect_docx(
                input_path=input_path,
                output_path=output_path
            )
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        return FileResponse(
            output_path,
            filename=f"protected_{file.filename}",
            media_type=mime_type
        )
    
    except Exception as e:
        return {"error": str(e)}, 500
    
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)
