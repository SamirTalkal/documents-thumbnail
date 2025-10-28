# Document Thumbnail Service

A FastAPI service that generates thumbnails from uploaded documents (PDF, DOC, DOCX, PPT, PPTX).

## Features

- Upload documents and get first-page thumbnails
- Supports PDF, DOC, DOCX, PPT, PPTX formats
- RESTful API with automatic documentation
- File management (list, delete)
- Static file serving for thumbnails and originals

## Quick Start (Development)

```bash
# Install dependencies
pip install -r requirements.txt

# Run server
uvicorn app:app --reload
```

## Production Deployment

### Prerequisites
- Ubuntu/Debian server
- Python 3.11+
- LibreOffice installed

### Deploy to Server

1. **Copy files to server:**
```bash
scp -r . user@your-server:/var/app/document-thumbnail/
```

2. **Run deployment script:**
```bash
ssh user@your-server
cd /var/app/document-thumbnail
sudo ./deploy.sh
```

### Manual Setup

1. **Install dependencies:**
```bash
sudo apt-get update
sudo apt-get install -y python3 python3-pip python3-venv libreoffice
```

2. **Setup application:**
```bash
cd /var/app/document-thumbnail
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

3. **Configure environment:**
```bash
cp env.production .env
# Edit .env with your settings
```

4. **Install as system service:**
```bash
sudo cp document-thumbnail.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable document-thumbnail
sudo systemctl start document-thumbnail
```

## API Endpoints

- `POST /upload` - Upload document and generate thumbnail
- `GET /files` - List all uploaded files
- `DELETE /delete/{doc_id}` - Delete file and thumbnail
- `GET /thumbnails/{doc_id}.png` - View thumbnail
- `GET /documents/{doc_id}.ext` - Download original file
- `GET /docs` - API documentation
- `GET /health` - Health check

## Configuration

Environment variables (set in `.env`):

- `BASE_STORAGE_DIR` - Storage directory path
- `MAX_FILE_SIZE` - Max upload size in bytes
- `SOFFICE_CMD` - LibreOffice executable path
- `CORS_ALLOW_ORIGINS` - Allowed CORS origins

## Service Management

```bash
# Start/stop service
sudo systemctl start document-thumbnail
sudo systemctl stop document-thumbnail
sudo systemctl restart document-thumbnail

# Check status
sudo systemctl status document-thumbnail

# View logs
sudo journalctl -u document-thumbnail -f
```

## Security Notes

- Service runs as `www-data` user
- Storage directory has restricted permissions
- CORS can be configured for specific domains
- File size limits prevent abuse
