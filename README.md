# automated-excel-to-screen-converter
Multiple files excel to screen converter

## General Information
This Python script **automates capturing screenshots** of predefined areas of the screen, specifically targeting certain Excel cells.  
The workflow ensures **high-quality images** by saving the screenshots to PDF, and then extracting images as PNG.  

## Features
- Capture specific Excel worksheet regions (defined print area) for all Excel files in a selected directory
- Export the selected region directly from Excel as a PDF
- Convert the PDF to high-resolution PNG images
- Automatically trim empty margins/background from the images
- Fully automated process for efficiency and consistency
- Designed to preserve maximum image clarity and resolution

## Software / Tech Stack

- **Python 3.x**  
- **win32com.client** (`pywin32`) – Excel automation and PDF export  
- **Pillow (PIL)** – image handling and trimming  
- **pdf2image** – PDF-to-image conversion  
- **os & datetime** – file operations and automation  
- **psutil** – managing/killing Excel processes  
- **tkinter** – folder selection GUI

## How It Works
1. Define the **screen area** (print area) corresponding to Excel cells to capture.  
2. The script automatically processes **all Excel files in the selected folder**.  
3. The defined area is **exported directly from Excel as a PDF**.  
4. The PDF is **converted to high-resolution PNG images**.  
5. Images are **automatically trimmed** to remove empty margins/background.  
6. Output images maintain **maximum clarity and quality** for reporting or documentation.

## Possible Improvements / Future Work
- Add configuration for **dynamic screen areas**  
- Automate naming and organization of output images  
- Add GUI for easier selection of regions and file paths  
