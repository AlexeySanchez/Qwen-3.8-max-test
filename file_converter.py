"""
File Converter Application
Converts various formats (md, json, html, csv, py) to docx, xlsx, pptx
Supports text input and file upload
Includes API endpoint for Qwen integration
"""

import os
import sys
import json
import tempfile
import base64
from pathlib import Path
from typing import Optional, Dict, Any

# GUI imports
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# Web server imports for Qwen integration
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import threading

# Document processing imports
try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    Document = None

try:
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
except ImportError:
    Presentation = None

import html
import csv
import io


class FileConverter:
    """Core conversion logic for various file formats"""
    
    @staticmethod
    def md_to_docx(content: str, output_path: str) -> str:
        """Convert Markdown to DOCX"""
        if Document is None:
            raise ImportError("python-docx not installed")
        
        doc = Document()
        
        lines = content.split('\n')
        current_paragraph = []
        
        for line in lines:
            stripped = line.strip()
            
            # Headers
            if stripped.startswith('# '):
                if current_paragraph:
                    doc.add_paragraph(' '.join(current_paragraph))
                    current_paragraph = []
                doc.add_heading(stripped[2:], level=1)
            elif stripped.startswith('## '):
                if current_paragraph:
                    doc.add_paragraph(' '.join(current_paragraph))
                    current_paragraph = []
                doc.add_heading(stripped[3:], level=2)
            elif stripped.startswith('### '):
                if current_paragraph:
                    doc.add_paragraph(' '.join(current_paragraph))
                    current_paragraph = []
                doc.add_heading(stripped[4:], level=3)
            # Code blocks
            elif stripped.startswith('```'):
                if current_paragraph:
                    doc.add_paragraph(' '.join(current_paragraph))
                    current_paragraph = []
                # Skip code block markers, add content as monospace
            elif stripped.startswith('- ') or stripped.startswith('* '):
                if current_paragraph:
                    doc.add_paragraph(' '.join(current_paragraph))
                    current_paragraph = []
                p = doc.add_paragraph(stripped[2:], style='List Bullet')
            elif stripped.isdigit() or (stripped and stripped[0].isdigit() and '. ' in stripped):
                if current_paragraph:
                    doc.add_paragraph(' '.join(current_paragraph))
                    current_paragraph = []
                # Simple numbered list handling
                p = doc.add_paragraph(stripped, style='List Number')
            elif stripped:
                current_paragraph.append(stripped)
            else:
                if current_paragraph:
                    doc.add_paragraph(' '.join(current_paragraph))
                    current_paragraph = []
                doc.add_paragraph('')  # Empty line
        
        if current_paragraph:
            doc.add_paragraph(' '.join(current_paragraph))
        
        doc.save(output_path)
        return output_path
    
    @staticmethod
    def json_to_docx(content: str, output_path: str) -> str:
        """Convert JSON to DOCX"""
        if Document is None:
            raise ImportError("python-docx not installed")
        
        try:
            data = json.loads(content)
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON: {e}")
        
        doc = Document()
        doc.add_heading('JSON Data', level=1)
        
        def add_json_data(data, parent=None, indent=0):
            if isinstance(data, dict):
                for key, value in data.items():
                    p = doc.add_paragraph()
                    runner = p.add_run(f"{'  ' * indent}{key}: ")
                    runner.bold = True
                    add_json_data(value, p, indent + 1)
            elif isinstance(data, list):
                for i, item in enumerate(data):
                    p = doc.add_paragraph()
                    runner = p.add_run(f"{'  ' * indent}[{i}]: ")
                    add_json_data(item, p, indent + 1)
            else:
                p = doc.add_paragraph(f"{'  ' * indent}{data}")
        
        add_json_data(data)
        doc.save(output_path)
        return output_path
    
    @staticmethod
    def html_to_docx(content: str, output_path: str) -> str:
        """Convert HTML to DOCX"""
        if Document is None:
            raise ImportError("python-docx not installed")
        
        doc = Document()
        
        # Simple HTML parsing (basic tags)
        # Remove script and style tags
        import re
        content = re.sub(r'<script[^>]*>.*?</script>', '', content, flags=re.DOTALL | re.IGNORECASE)
        content = re.sub(r'<style[^>]*>.*?</style>', '', content, flags=re.DOTALL | re.IGNORECASE)
        
        # Decode HTML entities
        content = html.unescape(content)
        
        # Extract text from basic tags
        lines = []
        current_line = []
        
        # Split by block elements
        blocks = re.split(r'</(?:p|div|h[1-6]|br|li)?>', content, flags=re.IGNORECASE)
        
        for block in blocks:
            # Remove remaining tags
            text = re.sub(r'<[^>]+>', ' ', block)
            text = text.strip()
            if text:
                lines.append(text)
        
        for line in lines:
            if line:
                doc.add_paragraph(line)
        
        doc.save(output_path)
        return output_path
    
    @staticmethod
    def py_to_docx(content: str, output_path: str) -> str:
        """Convert Python code to DOCX"""
        if Document is None:
            raise ImportError("python-docx not installed")
        
        doc = Document()
        doc.add_heading('Python Code', level=1)
        
        # Add code with monospace formatting
        p = doc.add_paragraph()
        runner = p.add_run(content)
        runner.font.name = 'Courier New'
        runner.font.size = Pt(10)
        
        doc.save(output_path)
        return output_path
    
    @staticmethod
    def md_to_pptx(content: str, output_path: str) -> str:
        """Convert Markdown to PPTX"""
        if Presentation is None:
            raise ImportError("python-pptx not installed")
        
        prs = Presentation()
        
        lines = content.split('\n')
        current_slide_content = []
        slide_title = "Slide"
        
        def create_slide(title, content_lines):
            if len(prs.slides) == 0:
                layout = prs.slide_layouts[0]  # Title slide
                slide = prs.slides.add_slide(layout)
                slide.shapes.title.text = title
                if content_lines:
                    subtitle = slide.placeholders[1]
                    subtitle.text = '\n'.join(content_lines[:5])
            else:
                layout = prs.slide_layouts[1]  # Title and Content
                slide = prs.slides.add_slide(layout)
                slide.shapes.title.text = title
                body = slide.placeholders[1]
                tf = body.text_frame
                tf.clear()
                
                for i, line in enumerate(content_lines):
                    if i == 0:
                        p = tf.paragraphs[0]
                    else:
                        p = tf.add_paragraph()
                    p.text = line
                    p.level = 0
        
        for line in lines:
            stripped = line.strip()
            
            if stripped.startswith('# '):
                if current_slide_content:
                    create_slide(slide_title, current_slide_content)
                    current_slide_content = []
                slide_title = stripped[2:]
            elif stripped.startswith('## '):
                if current_slide_content:
                    create_slide(slide_title, current_slide_content)
                    current_slide_content = []
                slide_title = stripped[3:]
            elif stripped.startswith('- ') or stripped.startswith('* '):
                current_slide_content.append(stripped[2:])
            elif stripped:
                current_slide_content.append(stripped)
        
        if current_slide_content or slide_title != "Slide":
            create_slide(slide_title, current_slide_content)
        
        # Ensure at least one slide
        if len(prs.slides) == 0:
            layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = "Presentation"
        
        prs.save(output_path)
        return output_path
    
    @staticmethod
    def json_to_pptx(content: str, output_path: str) -> str:
        """Convert JSON to PPTX"""
        if Presentation is None:
            raise ImportError("python-pptx not installed")
        
        try:
            data = json.loads(content)
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON: {e}")
        
        prs = Presentation()
        
        # Create title slide
        layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text = "JSON Data Presentation"
        
        def add_data_to_slides(data, title="Data"):
            if isinstance(data, dict):
                layout = prs.slide_layouts[1]
                slide = prs.slides.add_slide(layout)
                slide.shapes.title.text = title
                
                body = slide.placeholders[1]
                tf = body.text_frame
                tf.clear()
                
                for i, (key, value) in enumerate(data.items()):
                    if i == 0:
                        p = tf.paragraphs[0]
                    else:
                        p = tf.add_paragraph()
                    p.text = f"{key}: {value}"
                    
                    if isinstance(value, (dict, list)):
                        add_data_to_slides(value, key)
            
            elif isinstance(data, list):
                layout = prs.slide_layouts[1]
                slide = prs.slides.add_slide(layout)
                slide.shapes.title.text = title
                
                body = slide.placeholders[1]
                tf = body.text_frame
                tf.clear()
                
                for i, item in enumerate(data):
                    if i == 0:
                        p = tf.paragraphs[0]
                    else:
                        p = tf.add_paragraph()
                    p.text = f"Item {i}: {item}"
        
        if isinstance(data, dict):
            add_data_to_slides(data)
        elif isinstance(data, list):
            add_data_to_slides(data, "List Items")
        
        prs.save(output_path)
        return output_path
    
    @staticmethod
    def html_to_pptx(content: str, output_path: str) -> str:
        """Convert HTML to PPTX"""
        if Presentation is None:
            raise ImportError("python-pptx not installed")
        
        prs = Presentation()
        
        # Simple HTML parsing
        import re
        content = re.sub(r'<script[^>]*>.*?</script>', '', content, flags=re.DOTALL | re.IGNORECASE)
        content = re.sub(r'<style[^>]*>.*?</style>', '', content, flags=re.DOTALL | re.IGNORECASE)
        content = html.unescape(content)
        
        # Extract headings and paragraphs
        headings = re.findall(r'<h[1-6][^>]*>(.*?)</h[1-6]>', content, re.IGNORECASE)
        paragraphs = re.findall(r'<p[^>]*>(.*?)</p>', content, re.IGNORECASE)
        
        # Create slides from headings
        if headings:
            for i, heading in enumerate(headings):
                heading = re.sub(r'<[^>]+>', '', heading).strip()
                slide_content = []
                
                if i < len(paragraphs):
                    para = re.sub(r'<[^>]+>', '', paragraphs[i]).strip()
                    slide_content.append(para)
                
                if i == 0:
                    layout = prs.slide_layouts[0]
                    slide = prs.slides.add_slide(layout)
                    slide.shapes.title.text = heading
                else:
                    layout = prs.slide_layouts[1]
                    slide = prs.slides.add_slide(layout)
                    slide.shapes.title.text = heading
                    
                    if slide_content:
                        body = slide.placeholders[1]
                        tf = body.text_frame
                        tf.paragraphs[0].text = slide_content[0]
        else:
            # No headings, create from paragraphs
            layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = "HTML Content"
            
            if paragraphs:
                body = slide.placeholders[1]
                body.text = re.sub(r'<[^>]+>', '', paragraphs[0]).strip()
        
        prs.save(output_path)
        return output_path
    
    @staticmethod
    def py_to_pptx(content: str, output_path: str) -> str:
        """Convert Python code to PPTX"""
        if Presentation is None:
            raise ImportError("python-pptx not installed")
        
        prs = Presentation()
        
        # Title slide
        layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text = "Python Code"
        subtitle = slide.placeholders[1]
        subtitle.text = "Code Presentation"
        
        # Code slide
        layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text = "Source Code"
        
        body = slide.placeholders[1]
        tf = body.text_frame
        tf.clear()
        
        lines = content.split('\n')[:20]  # Limit to 20 lines
        for i, line in enumerate(lines):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = line
            p.font.name = 'Courier New'
        
        prs.save(output_path)
        return output_path
    
    @staticmethod
    def csv_to_xlsx(content: str, output_path: str) -> str:
        """Convert CSV to XLSX"""
        if Workbook is None:
            raise ImportError("openpyxl not installed")
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        
        # Parse CSV
        reader = csv.reader(io.StringIO(content))
        
        for row_idx, row in enumerate(reader, 1):
            for col_idx, value in enumerate(row, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)
        
        # Auto-adjust column widths
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column].width = adjusted_width
        
        wb.save(output_path)
        return output_path
    
    @staticmethod
    def json_to_xlsx(content: str, output_path: str) -> str:
        """Convert JSON to XLSX"""
        if Workbook is None:
            raise ImportError("openpyxl not installed")
        
        try:
            data = json.loads(content)
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON: {e}")
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        
        def flatten_json(data, prefix='', result=None):
            if result is None:
                result = {}
            
            if isinstance(data, dict):
                for key, value in data.items():
                    new_key = f"{prefix}.{key}" if prefix else key
                    flatten_json(value, new_key, result)
            elif isinstance(data, list):
                for i, item in enumerate(data):
                    flatten_json(item, f"{prefix}[{i}]", result)
            else:
                result[prefix] = data
            
            return result
        
        if isinstance(data, list):
            # List of objects
            flattened_data = [flatten_json(item) for item in data]
            all_keys = set()
            for item in flattened_data:
                all_keys.update(item.keys())
            
            headers = sorted(list(all_keys))
            
            # Write headers
            for col_idx, header in enumerate(headers, 1):
                ws.cell(row=1, column=col_idx, value=header)
            
            # Write data
            for row_idx, item in enumerate(flattened_data, 2):
                for col_idx, header in enumerate(headers, 1):
                    ws.cell(row=row_idx, column=col_idx, value=item.get(header, ''))
        else:
            # Single object
            flattened = flatten_json(data)
            for row_idx, (key, value) in enumerate(flattened.items(), 1):
                ws.cell(row=row_idx, column=1, value=key)
                ws.cell(row=row_idx, column=2, value=value)
        
        wb.save(output_path)
        return output_path
    
    @staticmethod
    def html_to_xlsx(content: str, output_path: str) -> str:
        """Convert HTML to XLSX"""
        if Workbook is None:
            raise ImportError("openpyxl not installed")
        
        wb = Workbook()
        ws = wb.active
        ws.title = "HTML Data"
        
        # Try to extract tables
        import re
        tables = re.findall(r'<table[^>]*>(.*?)</table>', content, re.DOTALL | re.IGNORECASE)
        
        if tables:
            table = tables[0]
            rows = re.findall(r'<tr[^>]*>(.*?)</tr>', table, re.DOTALL | re.IGNORECASE)
            
            for row_idx, row in enumerate(rows, 1):
                cells = re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', row, re.DOTALL | re.IGNORECASE)
                for col_idx, cell in enumerate(cells, 1):
                    cell_text = re.sub(r'<[^>]+>', '', cell).strip()
                    cell_text = html.unescape(cell_text)
                    ws.cell(row=row_idx, column=col_idx, value=cell_text)
        else:
            # No tables, extract text
            content = re.sub(r'<script[^>]*>.*?</script>', '', content, flags=re.DOTALL | re.IGNORECASE)
            content = re.sub(r'<style[^>]*>.*?</style>', '', content, flags=re.DOTALL | re.IGNORECASE)
            content = html.unescape(content)
            text = re.sub(r'<[^>]+>', '\n', content)
            
            lines = text.split('\n')
            for row_idx, line in enumerate(lines[:100], 1):  # Limit to 100 rows
                if line.strip():
                    ws.cell(row=row_idx, column=1, value=line.strip())
        
        wb.save(output_path)
        return output_path
    
    @staticmethod
    def py_to_xlsx(content: str, output_path: str) -> str:
        """Convert Python code to XLSX"""
        if Workbook is None:
            raise ImportError("openpyxl not installed")
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Python Code"
        
        ws.cell(row=1, column=1, value="Line Number")
        ws.cell(row=1, column=2, value="Code")
        
        lines = content.split('\n')
        for row_idx, line in enumerate(lines, 2):
            ws.cell(row=row_idx, column=1, value=row_idx - 1)
            ws.cell(row=row_idx, column=2, value=line)
        
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 80
        
        wb.save(output_path)
        return output_path


class ConverterApp:
    """Main GUI Application"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("File Converter - MD/JSON/HTML/CSV/PY to DOCX/XLSX/PPTX")
        self.root.geometry("900x700")
        
        self.converter = FileConverter()
        self.current_file_path = None
        
        self.setup_ui()
        self.start_api_server()
    
    def setup_ui(self):
        """Setup the user interface"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Input type selection
        ttk.Label(main_frame, text="Input Type:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_type = tk.StringVar(value="text")
        ttk.Radiobutton(main_frame, text="Text Input", variable=self.input_type, 
                       value="text", command=self.toggle_input_mode).grid(row=0, column=1, sticky=tk.W)
        ttk.Radiobutton(main_frame, text="File Upload", variable=self.input_type, 
                       value="file", command=self.toggle_input_mode).grid(row=0, column=2, sticky=tk.W)
        
        # Output format selection
        ttk.Label(main_frame, text="Output Format:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_format = tk.StringVar(value="docx")
        output_formats = [
            ("DOCX", "docx"),
            ("XLSX", "xlsx"),
            ("PPTX", "pptx")
        ]
        for i, (text, value) in enumerate(output_formats):
            ttk.Radiobutton(main_frame, text=text, variable=self.output_format, 
                          value=value).grid(row=1, column=i+1, sticky=tk.W)
        
        # Text input area
        self.text_frame = ttk.LabelFrame(main_frame, text="Input Content", padding="5")
        self.text_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        self.text_frame.columnconfigure(0, weight=1)
        self.text_frame.rowconfigure(0, weight=1)
        
        self.text_input = scrolledtext.ScrolledText(self.text_frame, wrap=tk.WORD, width=80, height=20)
        self.text_input.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # File input frame (hidden by default)
        self.file_frame = ttk.LabelFrame(main_frame, text="File Input", padding="5")
        self.file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        self.file_frame.grid_remove()
        
        ttk.Button(self.file_frame, text="Browse File", command=self.browse_file).grid(row=0, column=0, padx=5)
        self.file_label = ttk.Label(self.file_frame, text="No file selected")
        self.file_label.grid(row=0, column=1, sticky=tk.W)
        
        # Convert button
        ttk.Button(main_frame, text="Convert", command=self.convert).grid(row=3, column=0, pady=10)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready")
        self.status_label.grid(row=3, column=1, sticky=tk.W, pady=10)
        
        # Save button
        self.save_btn = ttk.Button(main_frame, text="Save As...", command=self.save_file, state=tk.DISABLED)
        self.save_btn.grid(row=3, column=2, pady=10)
        
        # API Info
        api_frame = ttk.LabelFrame(main_frame, text="API Access for Qwen", padding="5")
        api_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.api_url_label = ttk.Label(api_frame, text="API Server: http://localhost:5000/api/convert")
        self.api_url_label.grid(row=0, column=0, sticky=tk.W)
        
        ttk.Button(api_frame, text="Test API", command=self.test_api).grid(row=0, column=1, padx=10)
    
    def toggle_input_mode(self):
        """Toggle between text and file input modes"""
        if self.input_type.get() == "text":
            self.text_frame.grid()
            self.file_frame.grid_remove()
        else:
            self.text_frame.grid_remove()
            self.file_frame.grid()
    
    def browse_file(self):
        """Open file browser dialog"""
        filetypes = [
            ("All supported files", "*.md *.json *.html *.csv *.py"),
            ("Markdown files", "*.md"),
            ("JSON files", "*.json"),
            ("HTML files", "*.html *.htm"),
            ("CSV files", "*.csv"),
            ("Python files", "*.py"),
            ("All files", "*.*")
        ]
        
        filepath = filedialog.askopenfilename(
            title="Select file to convert",
            filetypes=filetypes
        )
        
        if filepath:
            self.current_file_path = filepath
            filename = os.path.basename(filepath)
            self.file_label.config(text=filename)
            
            # Optionally load file content
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
                    self.text_input.delete(1.0, tk.END)
                    self.text_input.insert(tk.END, content)
            except Exception as e:
                messagebox.showerror("Error", f"Could not read file: {e}")
    
    def convert(self):
        """Perform the conversion"""
        try:
            # Get input content
            if self.input_type.get() == "text":
                content = self.text_input.get(1.0, tk.END).strip()
                if not content:
                    messagebox.showwarning("Warning", "Please enter some content to convert")
                    return
            else:
                if not self.current_file_path:
                    messagebox.showwarning("Warning", "Please select a file first")
                    return
                with open(self.current_file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            
            # Determine source format (auto-detect or use file extension)
            source_format = self.detect_source_format(content, self.current_file_path)
            output_format = self.output_format.get()
            
            # Check if conversion is supported
            if not self.is_conversion_supported(source_format, output_format):
                messagebox.showerror(
                    "Error", 
                    f"Conversion from {source_format} to {output_format} is not supported.\n\n"
                    f"Supported conversions:\n"
                    f"- DOCX: md, json, html, py\n"
                    f"- PPTX: md, json, html, py\n"
                    f"- XLSX: csv, json, html, py"
                )
                return
            
            # Create temporary output file
            temp_dir = tempfile.gettempdir()
            output_filename = f"converted_{os.urandom(4).hex()}.{output_format}"
            output_path = os.path.join(temp_dir, output_filename)
            
            # Perform conversion
            self.status_label.config(text="Converting...")
            self.root.update()
            
            if output_format == "docx":
                if source_format == "md":
                    self.converter.md_to_docx(content, output_path)
                elif source_format == "json":
                    self.converter.json_to_docx(content, output_path)
                elif source_format == "html":
                    self.converter.html_to_docx(content, output_path)
                elif source_format == "py":
                    self.converter.py_to_docx(content, output_path)
            
            elif output_format == "pptx":
                if source_format == "md":
                    self.converter.md_to_pptx(content, output_path)
                elif source_format == "json":
                    self.converter.json_to_pptx(content, output_path)
                elif source_format == "html":
                    self.converter.html_to_pptx(content, output_path)
                elif source_format == "py":
                    self.converter.py_to_pptx(content, output_path)
            
            elif output_format == "xlsx":
                if source_format == "csv":
                    self.converter.csv_to_xlsx(content, output_path)
                elif source_format == "json":
                    self.converter.json_to_xlsx(content, output_path)
                elif source_format == "html":
                    self.converter.html_to_xlsx(content, output_path)
                elif source_format == "py":
                    self.converter.py_to_xlsx(content, output_path)
            
            self.current_output_path = output_path
            self.save_btn.config(state=tk.NORMAL)
            self.status_label.config(text=f"Conversion successful! Ready to save.")
            messagebox.showinfo("Success", "Conversion completed successfully!")
            
        except Exception as e:
            self.status_label.config(text="Conversion failed")
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")
    
    def detect_source_format(self, content: str, filepath: Optional[str] = None) -> str:
        """Detect the source format from content or file extension"""
        if filepath:
            ext = os.path.splitext(filepath)[1].lower()
            ext_map = {
                '.md': 'md',
                '.markdown': 'md',
                '.json': 'json',
                '.html': 'html',
                '.htm': 'html',
                '.csv': 'csv',
                '.py': 'py'
            }
            if ext in ext_map:
                return ext_map[ext]
        
        # Try to detect from content
        content_stripped = content.strip()
        
        # JSON detection
        if content_stripped.startswith('{') or content_stripped.startswith('['):
            try:
                json.loads(content_stripped)
                return 'json'
            except:
                pass
        
        # CSV detection (has commas and multiple lines)
        if ',' in content and '\n' in content:
            try:
                reader = csv.reader(io.StringIO(content))
                rows = list(reader)
                if len(rows) > 1 and all(len(row) > 1 for row in rows):
                    return 'csv'
            except:
                pass
        
        # HTML detection
        if '<' in content and ('<' in content.split('>')[0] if '>' in content else False):
            import re
            if re.search(r'<[a-z][\s\S]*>', content, re.IGNORECASE):
                return 'html'
        
        # Python detection (has def, class, import, etc.)
        python_keywords = ['def ', 'class ', 'import ', 'from ', 'if __name__']
        if any(kw in content for kw in python_keywords):
            return 'py'
        
        # Default to markdown
        return 'md'
    
    def is_conversion_supported(self, source: str, target: str) -> bool:
        """Check if conversion is supported"""
        supported = {
            'docx': ['md', 'json', 'html', 'py'],
            'pptx': ['md', 'json', 'html', 'py'],
            'xlsx': ['csv', 'json', 'html', 'py']
        }
        return source in supported.get(target, [])
    
    def save_file(self):
        """Save the converted file"""
        if not hasattr(self, 'current_output_path') or not self.current_output_path:
            return
        
        output_format = self.output_format.get()
        filetypes = [
            (f"{output_format.upper()} files", f"*.{output_format}"),
            ("All files", "*.*")
        ]
        
        filepath = filedialog.asksaveasfilename(
            title="Save converted file",
            defaultextension=f".{output_format}",
            filetypes=filetypes
        )
        
        if filepath:
            try:
                import shutil
                shutil.copy2(self.current_output_path, filepath)
                self.status_label.config(text=f"Saved to: {filepath}")
                messagebox.showinfo("Success", f"File saved to:\n{filepath}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not save file: {e}")
    
    def test_api(self):
        """Test the API endpoint"""
        try:
            import requests
            test_data = {
                "content": "# Test\nThis is a test",
                "source_format": "md",
                "target_format": "docx"
            }
            response = requests.post("http://localhost:5000/api/convert", json=test_data, timeout=5)
            if response.status_code == 200:
                messagebox.showinfo("API Test", "API is working correctly!")
            else:
                messagebox.showwarning("API Test", f"API returned status {response.status_code}")
        except requests.exceptions.ConnectionError:
            messagebox.showwarning("API Test", "API server is not running")
        except Exception as e:
            messagebox.showerror("API Test", f"Error: {e}")
    
    def start_api_server(self):
        """Start Flask API server in background thread"""
        app = Flask(__name__)
        CORS(app)
        
        @app.route('/api/convert', methods=['POST'])
        def api_convert():
            try:
                data = request.get_json()
                
                if not data:
                    return jsonify({"error": "No JSON data provided"}), 400
                
                content = data.get('content', '')
                source_format = data.get('source_format', 'md')
                target_format = data.get('target_format', 'docx')
                filename = data.get('filename', 'converted')
                
                if not content:
                    return jsonify({"error": "No content provided"}), 400
                
                # Validate formats
                if not self.is_conversion_supported(source_format, target_format):
                    return jsonify({
                        "error": f"Unsupported conversion: {source_format} to {target_format}"
                    }), 400
                
                # Create temporary file
                temp_dir = tempfile.gettempdir()
                output_filename = f"{filename}.{target_format}"
                output_path = os.path.join(temp_dir, output_filename)
                
                # Perform conversion
                if target_format == "docx":
                    if source_format == "md":
                        self.converter.md_to_docx(content, output_path)
                    elif source_format == "json":
                        self.converter.json_to_docx(content, output_path)
                    elif source_format == "html":
                        self.converter.html_to_docx(content, output_path)
                    elif source_format == "py":
                        self.converter.py_to_docx(content, output_path)
                
                elif target_format == "pptx":
                    if source_format == "md":
                        self.converter.md_to_pptx(content, output_path)
                    elif source_format == "json":
                        self.converter.json_to_pptx(content, output_path)
                    elif source_format == "html":
                        self.converter.html_to_pptx(content, output_path)
                    elif source_format == "py":
                        self.converter.py_to_pptx(content, output_path)
                
                elif target_format == "xlsx":
                    if source_format == "csv":
                        self.converter.csv_to_xlsx(content, output_path)
                    elif source_format == "json":
                        self.converter.json_to_xlsx(content, output_path)
                    elif source_format == "html":
                        self.converter.html_to_xlsx(content, output_path)
                    elif source_format == "py":
                        self.converter.py_to_xlsx(content, output_path)
                
                # Read file and encode as base64
                with open(output_path, 'rb') as f:
                    file_content = base64.b64encode(f.read()).decode('utf-8')
                
                # Clean up
                os.remove(output_path)
                
                return jsonify({
                    "success": True,
                    "filename": output_filename,
                    "content_base64": file_content,
                    "message": "Conversion successful"
                })
                
            except Exception as e:
                return jsonify({"error": str(e)}), 500
        
        @app.route('/api/health', methods=['GET'])
        def health_check():
            return jsonify({"status": "healthy", "service": "file-converter"})
        
        def run_server():
            app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)
        
        # Start server in background thread
        server_thread = threading.Thread(target=run_server, daemon=True)
        server_thread.start()
        
        print("API Server started on http://localhost:5000")


def main():
    """Main entry point"""
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
