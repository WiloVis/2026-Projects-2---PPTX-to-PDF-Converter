# 2026-Projects-2---PPTX-to-PDF-Converter
PPTX/PPT to PDF Converter (Windows)
A lightweight Python tool that converts PowerPoint files (.pptx, .ppt) to PDF format using a simple GUI dialog interface.

Features
- Graphical file selection – Native Windows dialogs for selecting input files and output folder
- Batch conversion – Select multiple PowerPoint files at once for conversion
- DPI awareness – Handles high-DPI displays to ensure proper UI scaling on modern Windows versions
- Progress feedback – Console output indicates when conversion is complete

Requirements
- Windows OS
- Microsoft PowerPoint (installed and licensed) – the script uses COM automation
- Python packages: comtypes, pptxtopdf, tkinter

How It Works
- Prompts you to select one or more PowerPoint files
- Prompts you to choose an output folder for the PDFs
- Converts each selected file to PDF in the chosen destination

Usage
- python converter.py

Notes
- The conversion relies on PowerPoint's native export functionality via COM
- PowerPoint must be installed and accessible on the system
- Designed specifically for Windows (uses ctypes.windll and comtypes.client)
