# File Converter Application

A Python application for converting various file formats (Markdown, JSON, HTML, CSV, Python) to Office formats (DOCX, XLSX, PPTX).

## Features

- **GUI Interface**: Desktop application with text input and file upload options
- **API Server**: RESTful API for programmatic access (Qwen integration)
- **Supported Conversions**:
  - **To DOCX**: Markdown (.md), JSON (.json), HTML (.html), Python (.py)
  - **To XLSX**: CSV (.csv), JSON (.json), HTML (.html), Python (.py)
  - **To PPTX**: Markdown (.md), JSON (.json), HTML (.html), Python (.py)

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### GUI Application

Run the desktop application:

```bash
python file_converter.py
```

The GUI provides:
- Text input area for pasting content
- File upload option for selecting files
- Output format selection (DOCX, XLSX, PPTX)
- Convert and Save buttons
- Built-in API server for Qwen integration

### API Server (for Qwen Integration)

Run the standalone API server:

```bash
python file_converter_api.py --host 0.0.0.0 --port 5000
```

#### API Endpoints

**POST /api/convert**
Convert content to target format.

Request body (JSON):
```json
{
    "content": "# Hello\nThis is markdown content",
    "source_format": "md",
    "target_format": "docx",
    "filename": "my_document"
}
```

Response:
```json
{
    "success": true,
    "filename": "my_document.docx",
    "content_base64": "<base64-encoded-file-content>",
    "message": "Conversion successful"
}
```

**GET /api/health**
Health check endpoint.

Response:
```json
{
    "status": "healthy",
    "service": "file-converter-api",
    "version": "1.0.0"
}
```

**GET /**
API documentation and usage examples.

## Qwen Integration

To enable Qwen to use this converter tool:

1. **Start the API server**:
   ```bash
   python file_converter_api.py --port 5000
   ```

2. **Configure Qwen** to make HTTP requests to `http://localhost:5000/api/convert`

3. **Example Qwen function call**:
   ```python
   import requests
   import base64
   
   def convert_to_docx(content, source_format='md', filename='document'):
       response = requests.post('http://localhost:5000/api/convert', json={
           'content': content,
           'source_format': source_format,
           'target_format': 'docx',
           'filename': filename
       })
       result = response.json()
       if result.get('success'):
           # Decode and save file
           file_content = base64.b64decode(result['content_base64'])
           with open(result['filename'], 'wb') as f:
               f.write(file_content)
           return result['filename']
       else:
           raise Exception(result.get('error', 'Conversion failed'))
   ```

4. **Supported parameters**:
   - `content`: The text content to convert (required)
   - `source_format`: One of `md`, `json`, `html`, `csv`, `py` (default: `md`)
   - `target_format`: One of `docx`, `xlsx`, `pptx` (default: `docx`)
   - `filename`: Output filename without extension (default: `converted`)

## Examples

### Convert Markdown to DOCX via API
```bash
curl -X POST http://localhost:5000/api/convert \
  -H "Content-Type: application/json" \
  -d '{
    "content": "# Hello World\n\nThis is a test.",
    "source_format": "md",
    "target_format": "docx",
    "filename": "hello"
  }'
```

### Convert JSON to XLSX via API
```bash
curl -X POST http://localhost:5000/api/convert \
  -H "Content-Type: application/json" \
  -d '{
    "content": "[{\"name\": \"Alice\", \"age\": 30}]",
    "source_format": "json",
    "target_format": "xlsx",
    "filename": "data"
  }'
```

### Convert Python Code to PPTX via API
```bash
curl -X POST http://localhost:5000/api/convert \
  -H "Content-Type: application/json" \
  -d '{
    "content": "def hello():\n    print(\"Hello\")",
    "source_format": "py",
    "target_format": "pptx",
    "filename": "code_presentation"
  }'
```

## Dependencies

- `python-docx` - For DOCX file creation
- `openpyxl` - For XLSX file creation
- `python-pptx` - For PPTX file creation
- `flask` - For API server
- `flask-cors` - For CORS support in API
- `requests` - For API testing

## License

MIT License
