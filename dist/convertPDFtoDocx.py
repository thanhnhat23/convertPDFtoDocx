import PyPDF2
import tkinter as tk
from tkinter import messagebox, filedialog, Tk
from docx import Document

def convertPDF(pdf_file):
    try:
        with open(pdf_file, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ''
            for page in reader.pages:
                text += page.extract_text()
            if not text.strip():
                raise ValueError("Can't extract text from PDF.")
            return text
    except Exception as e:
        messagebox.showerror("Error", f"Can't read file PDF: {str(e)}")
        return None

def convertDOCX(pdf_file, output_docx):
    text = convertPDF(pdf_file)
    if text:
        try:
            doc = Document()
            doc.add_paragraph(text)
            doc.save(output_docx)
            messagebox.showinfo("Completed", f"File DOCX has been saved: {output_docx}")
        except Exception as e:
            messagebox.showerror("Error", f"Can't save DOCX: {str(e)}")

def selectPDFandDOCX():
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_file:
        output_docx = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
        if output_docx:
            convertDOCX(pdf_file, output_docx)

root = Tk()
root.title("Convert PDF to DOCX")
root.geometry("400x200")

title_label = tk.Label(root, text="Convert PDF to DOCX", font=("Arial", 16))
title_label.pack(pady=20)

select_button = tk.Button(root, text="Select PDF", command=selectPDFandDOCX)
select_button.pack(pady=10)

root.mainloop()
