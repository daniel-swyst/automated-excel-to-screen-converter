import win32com.client as win32
import os
from PIL import Image, ImageChops
import pdf2image
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import psutil

def kill_excel_processes():
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == 'EXCEL.EXE':
            proc.kill()

def trim(image):
    bg = Image.new(image.mode, image.size, image.getpixel((0, 0)))
    diff = ImageChops.difference(image, bg)
    bbox = diff.getbbox()
    if bbox:
        return image.crop(bbox)
    return image

root = tk.Tk()
root.withdraw()

directory_path = filedialog.askdirectory(title="Select Directory with .xlsm Files")

if directory_path:
    kill_excel_processes()

    for file_name in os.listdir(directory_path):
        if file_name.endswith('.xlsm'):
            file_path = os.path.join(directory_path, file_name)

            excel = win32.Dispatch('Excel.Application')
            excel.Visible = False

            try:
                wb = excel.Workbooks.Open(file_path)
                sheet = wb.Sheets('Leakage Border Template')

                print_area = "A1:Q45"           ##cells
                sheet.PageSetup.PrintGridlines = True
                sheet.PageSetup.PrintArea = print_area
                sheet.PageSetup.Orientation = 2
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = 1
                sheet.PageSetup.LeftMargin = 0
                sheet.PageSetup.RightMargin = 0
                sheet.PageSetup.TopMargin = 0
                sheet.PageSetup.BottomMargin = 0


                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                pdf_file_name = f"{os.path.splitext(file_name)[0]}_{timestamp}.pdf"
                desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
                pdf_path = os.path.join(desktop_path, pdf_file_name)

                try:
                    sheet.ExportAsFixedFormat(Type=0, Filename=pdf_path, Quality=0)
                    print(f"PDF saved at {pdf_path}")
                except Exception as e:
                    print(f"Error saving PDF to {pdf_path}: {e}")
                    wb.Close(SaveChanges=False)
                    excel.Quit()
                    continue

                poppler_path = r'C:\Program Files (x86)\poppler-0.68.0\bin'
                images = pdf2image.convert_from_path(pdf_path, dpi=300, poppler_path=poppler_path)

                image_file_name = f"{os.path.splitext(file_name)[0]}_{timestamp}.png"
                image_path = os.path.join(directory_path, image_file_name)
                images[0].save(image_path, 'PNG')

                image = Image.open(image_path)
                trimmed_image = trim(image)
                trimmed_image.save(image_path, 'PNG')

                print(f"Image saved at {image_path}")

                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                    print(f"PDF file {pdf_path} deleted.")

            except Exception as e:
                print(f"An error occurred while processing {file_name}: {e}")

            finally:
                wb.Close(SaveChanges=False)
                excel.Quit()

    kill_excel_processes()

else:
    print("No directory selected.")
