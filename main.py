import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from pdf2docx import Converter
import PyPDF2
from PIL import Image
import fitz  # PyMuPDF
import os
import pyttsx3
from reportlab.pdfgen import canvas

class PDFUtilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SmartPDF ToolBox")
        self.root.geometry("1000x750")
        self.root.configure(bg="#f7e9d7")  # Light beige background

        # Title
        heading = tk.Label(root, text="SmartPDF ToolBox", font=("Georgia", 36, "bold"), bg="#f7e9d7", fg="#6a1b9a")
        heading.pack(pady=30)

        # Subtitle
        subheading = tk.Label(
            root,
            text="Your all-in-one PDF management tool",
            font=("Georgia", 18, "italic"),
            bg="#f7e9d7",
            fg="#283593",
        )
        subheading.pack(pady=10)

        # Frame for buttons
        button_frame = tk.Frame(root, bg="#f7e9d7")
        button_frame.pack(pady=10)

        # Button configuration
        button_info = [
            ("PDF to DOCX", self.pdf_to_docx, "#7e57c2"),
            ("Merge PDFs", self.merge_pdfs, "#8bc34a"),
            ("Image to PDF", self.image_to_pdf, "#ff5722"),
            ("Compress PDF", self.compress_pdf, "#9e9e9e"),
            ("Add Page Numbers", self.add_page_numbers, "#03a9f4"),
            ("Delete Pages from PDF", self.delete_pages, "#f44336"),
            ("Text to PDF", self.text_to_pdf, "#009688"),
            ("Encrypt PDF", self.encrypt_pdf, "#8e44ad"),
            ("Decrypt PDF", self.decrypt_pdf, "#e74c3c"),
            ("Extract Images from PDF", self.extract_images, "#ffa726"),
            ("PDF to Audio", self.pdf_to_audio, "#ff69b4"),
            ("Split PDF", self.split_pdf, "#00bcd4"),
        ]

        for idx, (text, command, color) in enumerate(button_info):
            button = tk.Button(
                button_frame,
                text=text,
                command=command,
                font=("Georgia", 14, "bold"),
                width=25,
                height=3,
                bg=color,
                fg="white",
                activebackground="#f7e9d7",
                activeforeground=color,
                bd=0,
            )
            button.grid(row=idx // 3, column=idx % 3, padx=20, pady=20)
            button.bind("<Enter>", lambda e, btn=button: btn.config(bg="white", fg=color))
            button.bind("<Leave>", lambda e, btn=button: btn.config(bg=color, fg="white"))

    # ------------------- Methods -------------------

    def pdf_to_docx(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filepath:
            save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
            if save_path:
                try:
                    cv = Converter(filepath)
                    cv.convert(save_path)
                    cv.close()
                    messagebox.showinfo("Success", f"PDF successfully converted to DOCX at {save_path}")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {e}")

    def merge_pdfs(self):
        filepaths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if filepaths:
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if save_path:
                try:
                    pdf_writer = PyPDF2.PdfWriter()
                    for path in filepaths:
                        pdf_reader = PyPDF2.PdfReader(path)
                        for page in pdf_reader.pages:
                            pdf_writer.add_page(page)
                    with open(save_path, "wb") as output_pdf:
                        pdf_writer.write(output_pdf)
                    messagebox.showinfo("Success", f"PDFs merged and saved at {save_path}")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {e}")

    def image_to_pdf(self):
        filepaths = filedialog.askopenfilenames(filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif")])
        if filepaths:
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if save_path:
                try:
                    images = [Image.open(fp).convert("RGB") for fp in filepaths]
                    images[0].save(save_path, save_all=True, append_images=images[1:])
                    messagebox.showinfo("Success", f"Images converted to PDF and saved at {save_path}")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {e}")

    def compress_pdf(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filepath:
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if save_path:
                try:
                    pdf = fitz.open(filepath)
                    pdf.save(save_path, deflate=True)
                    pdf.close()
                    messagebox.showinfo("Success", f"PDF compressed and saved at {save_path}")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {e}")

    def add_page_numbers(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filepath:
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if save_path:
                try:
                    pdf_document = fitz.open(filepath)
                    for page_num in range(len(pdf_document)):
                        page = pdf_document[page_num]
                        page.insert_text((500, 750), f"Page {page_num + 1}", fontsize=12, color=(0, 0, 0))
                    pdf_document.save(save_path)
                    pdf_document.close()
                    messagebox.showinfo("Success", f"Page numbers added and saved at {save_path}")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {e}")

    def delete_pages(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filepath:
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if save_path:
                try:
                    pdf_reader = PyPDF2.PdfReader(filepath)
                    pdf_writer = PyPDF2.PdfWriter()
                    pages_to_delete = simpledialog.askstring("Delete Pages", "Enter page numbers to delete (comma-separated):")
                    if pages_to_delete:
                        pages_to_delete = [int(p) - 1 for p in pages_to_delete.split(",")]
                        for i in range(len(pdf_reader.pages)):
                            if i not in pages_to_delete:
                                pdf_writer.add_page(pdf_reader.pages[i])
                        with open(save_path, "wb") as output_pdf:
                            pdf_writer.write(output_pdf)
                        messagebox.showinfo("Success", f"Selected pages deleted and saved at {save_path}")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {e}")

    def text_to_pdf(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if save_path:
            try:
                text_content = simpledialog.askstring("Text to PDF", "Enter the text to be added to the PDF:")
                if text_content:
                    c = canvas.Canvas(save_path, pagesize=(595, 842))
                    c.drawString(50, 800, text_content)
                    c.save()
                    messagebox.showinfo("Success", f"Text PDF saved at {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")

    def encrypt_pdf(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filepath:
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if save_path:
                password = simpledialog.askstring("Encrypt PDF", "Enter a password to secure the PDF:", show="*")
                try:
                    reader = PyPDF2.PdfReader(filepath)
                    writer = PyPDF2.PdfWriter()
                    for page in reader.pages:
                        writer.add_page(page)
                    writer.encrypt(password)
                    with open(save_path, "wb") as out_file:
                        writer.write(out_file)
                    messagebox.showinfo("Success", f"PDF encrypted and saved at {save_path}")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {e}")

    def decrypt_pdf(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filepath:
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            password = simpledialog.askstring("Decrypt PDF", "Enter the password:", show="*")
            try:
                reader = PyPDF2.PdfReader(filepath)
                if reader.is_encrypted:
                    reader.decrypt(password)
                writer = PyPDF2.PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)
                with open(save_path, "wb") as out_file:
                    writer.write(out_file)
                messagebox.showinfo("Success", f"PDF decrypted and saved at {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")

    def extract_images(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filepath:
            save_dir = filedialog.askdirectory(title="Select Folder to Save Images")
            if save_dir:
                try:
                    pdf_document = fitz.open(filepath)
                    image_count = 0
                    for i in range(len(pdf_document)):
                        for img in pdf_document.get_page_images(i):
                            xref = img[0]
                            base_image = pdf_document.extract_image(xref)
                            image_bytes = base_image["image"]
                            img_name = os.path.join(save_dir, f"extracted_image_{image_count + 1}.png")
                            with open(img_name, "wb") as img_file:
                                img_file.write(image_bytes)
                            image_count += 1
                    messagebox.showinfo("Success", f"{image_count} images extracted and saved.")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {e}")

    def pdf_to_audio(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filepath:
            save_path = filedialog.asksaveasfilename(defaultextension=".mp3", filetypes=[("Audio files", "*.mp3")])
            if save_path:
                try:
                    reader = PyPDF2.PdfReader(filepath)
                    text_content = ""
                    for page in reader.pages:
                        text_content += page.extract_text() + " "
                    if text_content.strip():
                        engine = pyttsx3.init()
                        engine.save_to_file(text_content, save_path)
                        engine.runAndWait()
                        messagebox.showinfo("Success", f"PDF converted to audio at {save_path}")
                    else:
                        messagebox.showwarning("Warning", "No text found in PDF.")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {e}")

    def split_pdf(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filepath:
            save_dir = filedialog.askdirectory(title="Select Folder to Save Split PDFs")
            if save_dir:
                try:
                    reader = PyPDF2.PdfReader(filepath)
                    for i, page in enumerate(reader.pages):
                        writer = PyPDF2.PdfWriter()
                        writer.add_page(page)
                        split_path = os.path.join(save_dir, f"page_{i+1}.pdf")
                        with open(split_path, "wb") as f:
                            writer.write(f)
                    messagebox.showinfo("Success", f"PDF split into {len(reader.pages)} files and saved at {save_dir}")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {e}")

# ------------------- Run App -------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFUtilityApp(root)
    root.mainloop()