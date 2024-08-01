import os
import shutil
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import messagebox

# Diretórios de entrada e saída
pato = os.path.dirname(os.path.abspath(__file__))
inpato = '../digitalizados'
outpato = '../registrados'
input_dir = os.path.join(pato, inpato)
output_dir = os.path.join(pato, outpato)
excel_file = os.path.join(pato, '../xlsx/extracted_data.xlsx')

notas_fiscais_dir = os.path.join(output_dir, "Notas Fiscais")
notas_transporte_dir = os.path.join(output_dir, "Notas de Transporte")

# Criar diretorios que não existem:
if not os.path.exists(input_dir):
    os.makedirs(input_dir)

if not os.path.exists(notas_fiscais_dir):
    os.makedirs(notas_fiscais_dir)

if not os.path.exists(notas_transporte_dir):
    os.makedirs(notas_transporte_dir)

# Initialize global variables
pdf_files = []
current_pdf_index = 0
history = []

def construct_database():
    # Convert Excel File to Panda Dataframe
    global df
    df = pd.read_excel(excel_file)
    print(df)

def load_pdf_files():
    global pdf_files
    pdf_files = [os.path.join(input_dir, f) for f in os.listdir(input_dir) if f.endswith(".pdf")]
    pdf_files.sort()  # Optional: Sort files if you want a specific order

def display_pdf_first_page(pdf_path, canvas):
    # Open the PDF
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    
    # Render the page as an image
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    
    # Display the image
    tk_img = ImageTk.PhotoImage(img)
    canvas.config(width=800, height=600)  # Set a fixed size for the canvas
    canvas.create_image(0, 0, anchor=tk.NW, image=tk_img)
    canvas.image = tk_img  # Keep a reference to avoid garbage collection
    
    # Update the canvas scroll region
    canvas.config(scrollregion=canvas.bbox("all"))
    
    doc.close()

def process_next_pdf():
    global current_pdf_index
    if current_pdf_index >= len(pdf_files):
        # Clear the canvas to indicate completion
        canvas.delete("all")
        # Inform the user
        messagebox.showinfo("Completo", "Todos os PDFs foram processados!")
        return
    
    file_path = pdf_files[current_pdf_index]
    display_pdf_first_page(file_path, canvas)
    
    barcode_entry.config(state=tk.NORMAL)
    assign_button.config(state=tk.NORMAL)


def assign_barcode(event=None):
    global current_pdf_index
    if current_pdf_index >= len(pdf_files):
        messagebox.showinfo("Completo", "Todos os PDFs foram processados!")
        return
    
    file_path = pdf_files[current_pdf_index]
    key = barcode_entry_var.get().replace(" ", "").strip()
    print("Barcode entry:", key)  # Debug: Print barcode entry
    if not key:
        messagebox.showwarning("Erro de entrada", "Por favor digite um código de barras.")
        return
    
    target_dir = ""
    if key in df['chNFe'].values:
        target_dir = notas_fiscais_dir
    elif key in df['chNTR'].values:
        target_dir = notas_transporte_dir
    else:
        response = messagebox.askyesno("Aviso", "Esse código não está na base de dados, deseja prosseguir mesmo assim?")
        if not response:
            return
        target_dir = output_dir
    
    original_name = os.path.basename(file_path)
    extension = os.path.splitext(original_name)[1]
    new_file_path = os.path.join(target_dir, f"{key}{extension}")
    
    if os.path.exists(new_file_path):
        response = messagebox.askyesno("Nota já registrada", "Nota já registrada, deseja incorporar o documento como página adicional?")
        if response:
            merge_pdfs(new_file_path, file_path)
        else:
            return
    else:
        shutil.move(file_path, new_file_path)
    
    # Add action to history for undo functionality
    history.append((file_path, new_file_path))
    
    # Clear the entry and reset UI
    barcode_entry_var.set("")
    barcode_entry.config(state=tk.DISABLED)
    assign_button.config(state=tk.DISABLED)
    
    # Move to the next PDF file
    current_pdf_index += 1
    process_next_pdf()

def merge_pdfs(existing_file, new_file):
    # Open the existing PDF
    existing_pdf = fitz.open(existing_file)
    new_pdf = fitz.open(new_file)
    
    # Append new PDF to existing PDF
    existing_pdf.insert_pdf(new_pdf)
    
    # Save the merged PDF to a temporary file
    temp_file = existing_file + ".tmp"
    existing_pdf.save(temp_file)
    existing_pdf.close()
    new_pdf.close()
    
    # Replace the original file with the merged file
    os.replace(temp_file, existing_file)
    
    # Delete the additional pages file
    os.remove(new_file)

def undo_last_action(event=None):
    if history:
        original_file, moved_file = history.pop()
        shutil.move(moved_file, original_file)
        messagebox.showinfo("Undo", "Última ação desfeita.")
        
        # Step back to the previous PDF
        global current_pdf_index
        current_pdf_index -= 1
        display_pdf_first_page(original_file, canvas)
        
        # Reset UI to reflect the undone state
        barcode_entry.config(state=tk.NORMAL)
        assign_button.config(state=tk.NORMAL)
    else:
        messagebox.showinfo("Undo", "Nenhuma ação para desfazer.")

def format_barcode_entry(*args):
    value = barcode_entry_var.get()
    value = value.replace(" ", "")  # Remove any existing spaces
    formatted_value = " ".join(value[i:i+4] for i in range(0, len(value), 4))
    barcode_entry_var.set(formatted_value)

def refresh_files():
    global current_pdf_index
    load_pdf_files()
    current_pdf_index = 0
    process_next_pdf()

# Carrega a base de dados do pandas
construct_database()
load_pdf_files()  # Load files initially

# Setup the GUI
window = tk.Tk()
window.title("Sakana Capture")

# Set fixed size for the window
window.geometry("620x768")  # Adjust as needed
window.minsize(600, 600)  # Set minimum size to ensure scrollbars work

# Create a frame to hold the canvas and scrollbar
canvas_frame = tk.Frame(window)
canvas_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# Create a canvas with a vertical scrollbar
canvas = tk.Canvas(canvas_frame, bg='white')
scrollbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=canvas.yview)
canvas.config(yscrollcommand=scrollbar.set)

# Pack scrollbar on the left side
scrollbar.pack(side=tk.LEFT, fill=tk.Y)
canvas.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

frame = tk.Frame(window)
frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

tk.Label(frame, text="Código de Barras:").grid(row=0, column=0, sticky=tk.W)


barcode_entry_var = tk.StringVar()
barcode_entry = tk.Entry(frame, textvariable=barcode_entry_var, state=tk.DISABLED, width=49)
barcode_entry.grid(row=0, column=1, padx=5)

barcode_entry_var.trace("w", format_barcode_entry)

# Button to the right of the entry
assign_button = tk.Button(frame, text="Associar Código", command=assign_barcode, state=tk.DISABLED)
assign_button.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)

'''
# Add a tiny button to trigger barcode reading
barcode_button = tk.Button(window, text="🔍", command=run_barcode_reader, width=2, height=1, font=("Arial", 14))
barcode_button.pack(side=tk.TOP, anchor=tk.W, padx=0, pady=0)  # Ensure it's not obscured
'''

# Add refresh button with a refresh symbol and place it in the top-right corner
refresh_button = tk.Button(window, text="🔄", command=refresh_files, width=2, height=1, font=("Arial", 14))
refresh_button.pack(side=tk.TOP, anchor=tk.W, padx=0, pady=0)  # Place in the top-right corner

# Bind Enter key to assign_barcode function
window.bind('<Return>', assign_barcode)
# Bind Ctrl+Z to undo_last_action function
window.bind('<Control-z>', undo_last_action)

# Start processing the first PDF
process_next_pdf()

# Prevent window from reopening after closing
window.protocol("WM_DELETE_WINDOW", window.quit)

window.mainloop()
