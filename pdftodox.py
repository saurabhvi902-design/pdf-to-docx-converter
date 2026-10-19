
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter


def convert_pdf():
    pdf_path = filedialog.askopenfilename(
        title="Select PDF File",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not pdf_path:
        return

    output_path = filedialog.asksaveasfilename(
        title="Save Word Document",
        defaultextension=".docx",
        initialfile=os.path.splitext(
            os.path.basename(pdf_path)
        )[0] + ".docx",
        filetypes=[("Word Documents", "*.docx")]
    )

    if not output_path:
        return

    converter = None

    try:
        status_label.config(text="Converting PDF...")
        root.update_idletasks()

        converter = Converter(pdf_path)
        converter.convert(output_path)

        status_label.config(text="Conversion completed!")
        messagebox.showinfo(
            "Success",
            f"Word file saved at:\n{output_path}"
        )

    except Exception as error:
        status_label.config(text="Conversion failed.")
        messagebox.showerror(
            "Error",
            f"Could not convert PDF:\n{error}"
        )

    finally:
        if converter is not None:
            converter.close()


root = tk.Tk()
root.title("PDF to DOCX Converter")
root.geometry("420x220")
root.resizable(False, False)

title_label = tk.Label(
    root,
    text="PDF to Word Converter",
    font=("Arial", 18, "bold")
)
title_label.pack(pady=25)

convert_button = tk.Button(
    root,
    text="Select PDF & Convert",
    command=convert_pdf,
    font=("Arial", 12),
    padx=15,
    pady=8
)
convert_button.pack()

status_label = tk.Label(
    root,
    text="Ready to convert",
    font=("Arial", 10)
)
status_label.pack(pady=20)

root.mainloop()
