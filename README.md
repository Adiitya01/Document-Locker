# Document Template Locker 🔒

A web application for protecting and locking Word (.docx) and Excel (.xlsx, .xls) documents. This tool allows you to lock template fields while keeping the rest of the document editable.

## Features

- **DOCX Protection**: Lock specific content in Word documents using content controls
- **XLSX/XLS Protection**: Protect Excel sheets while allowing editing in specific cells
- **Web Interface**: Easy-to-use FastAPI web application
- **File Upload**: Drag-and-drop file upload functionality
- **Download Protected Files**: Get your locked documents instantly

## Tech Stack

- **Backend**: FastAPI, Python 3.11
- **Document Processing**: python-docx, openpyxl, xlrd, lxml
- **Server**: Gunicorn with Uvicorn workers
- **Frontend**: HTML5

## Installation

### Local Setup

1. Clone the repository:
```bash
git clone https://github.com/Aditya01/Document-Locker.git
cd Document-Locker
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python main.py
```

The app will be available at `http://localhost:8000`

## Deployment

### Deploy on Render

This project is configured for deployment on Render using the `render.yaml` configuration file.

1. Push your code to GitHub
2. Go to [Render.com](https://render.com)
3. Click "New" → "Web Service"
4. Connect your GitHub repository
5. Select the `Document-Locker` repo
6. Render will auto-detect `render.yaml` and deploy automatically

Your app will be live at `https://document-locker-xxxx.onrender.com`

## Project Structure

```
├── main.py                 # FastAPI application entry point
├── docx_processor.py       # Document processing logic
├── index.html              # Frontend UI
├── requirements.txt        # Python dependencies
├── render.yaml             # Render deployment config
├── build.sh                # Build script
└── README.md               # This file
```

## Usage

1. Open the web application
2. Upload a .docx or .xlsx file
3. The application will process and lock the document
4. Download the protected file

## Requirements

- Python 3.11+
- FastAPI
- Uvicorn
- python-docx
- openpyxl
- lxml
- python-multipart

## License

This project is open source and available under the MIT License.

## Support

For issues or questions, please open an issue on the GitHub repository.

---

**Live Demo**: https://document-locker.onrender.com
