import os
import shutil
import pandas as pd
from pypdf import PdfWriter
import tkinter as tk
from tkinter import ttk, scrolledtext

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger")

        # Configure the main window
        self.root.geometry('600x400')

        # Create a frame for the progress bar
        self.progress_frame = tk.Frame(self.root)
        self.progress_frame.pack(fill=tk.X, padx=10, pady=5)

        # Create a progress bar
        self.progress = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, length=580, mode='determinate')
        self.progress.pack(pady=5)

        # Create a frame for the logger
        self.logger_frame = tk.Frame(self.root)
        self.logger_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create a scrolled text widget for the log
        self.log_text = scrolledtext.ScrolledText(self.logger_frame, wrap=tk.WORD, height=15, width=70, bg='#2A2B2A', fg='white')
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Create a button to start the merging process
        self.start_button = tk.Button(self.root, text="Começar Mesclagem", command=self.start_merging)
        self.start_button.pack(pady=10)

    def log(self, message):
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.yview(tk.END)
        
    def update_progress(self, value, maximum):
        self.progress['value'] = value
        self.progress['maximum'] = maximum
        self.root.update_idletasks()

    def start_merging(self):
        # Paths and columns configuration
        excel_file = os.path.join(script_dir, '../xlsx/extracted_data.xlsx')
        folder_path = os.path.join(script_dir, '../registrados')
        output_folder = os.path.join(script_dir, '../mesclados')
        archived_folder = os.path.join(folder_path, '../residuo')
        column1 = 'chNFe'
        column2 = 'chNTR'
        folder_column = 'xMun'
        suffix_column1 = 'FOR'
        suffix_column2 = 'vProd'

        # Ensure the output and archived folders exist
        os.makedirs(output_folder, exist_ok=True)
        os.makedirs(archived_folder, exist_ok=True)

        # Get the total number of files in the input folder
        total_files = len([f for f in os.listdir(folder_path) if f.endswith(".pdf")])

        # Initialize the progress bar
        self.update_progress(0, total_files)
        
        # Start the merging process
        find_and_merge_pdfs(
            excel_file, folder_path, column1, column2, output_folder, 
            folder_column, suffix_column1, suffix_column2, 
            abbrev_length=3, log_callback=self.log, 
            progress_callback=self.update_progress, total_files=total_files
        )

def merge_pdfs(pdf_list, output_path):
    """Merges PDF files from pdf_list into a single PDF at output_path."""
    merger = PdfWriter()
    for pdf in pdf_list:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()

def count_pdfs_in_subfolders(base_folder):
    """Counts the number of PDFs in each subfolder of base_folder."""
    folder_counts = {}
    for root, dirs, files in os.walk(base_folder):
        for dir in dirs:
            subfolder_path = os.path.join(root, dir)
            num_files = len([f for f in os.listdir(subfolder_path) if f.endswith(".pdf")])
            folder_counts[subfolder_path] = num_files
    return folder_counts

def find_and_merge_pdfs(excel_file, folder_path, column1, column2, output_folder, folder_column, suffix_column1, suffix_column2, abbrev_length=5, log_callback=None, progress_callback=None, total_files=0):
    """Finds, merges, and names PDF files based on Excel data."""

    # Read the Excel file
    df = pd.read_excel(excel_file)
    
    # Create a dictionary to map file names to their full paths
    pdf_files = {}
    
    # Traverse directories and subdirectories to find all PDF files
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".pdf"):
                file_name = os.path.splitext(file)[0]  # Remove the .pdf extension
                file_path = os.path.join(root, file)
                pdf_files[file_name] = file_path

    # Create 'arquivados' folder if it doesn't exist
    archived_folder = os.path.join(folder_path, '../residuo')
    os.makedirs(archived_folder, exist_ok=True)

    # Prepare progress bar
    processed_files = 0

    # Iterate over the rows of the DataFrame
    for index, row in df.iterrows():
        file1 = str(row[column1])  # Ensure file names are treated as strings
        file2 = str(row[column2])

        # Use only the last 8 characters of file1 and file2 for naming
        file1_last8 = file1[-8:] if len(file1) >= 8 else file1
        file2_last8 = file2[-8:] if len(file2) >= 8 else file2
        
        # Get folder and suffix names from specified columns
        subfolder_name = str(row[folder_column])  # Convert to string for valid folder name

        # Generate abbreviations for suffixes
        suffix1 = str(row[suffix_column1])[:abbrev_length]  # Abbreviate to first few characters

        # Convert the value in suffix_column2 to a formatted string with two decimal places
        value = row[suffix_column2]
        suffix2 = f"{value:.2f}"  # Format as a string with two decimal places

        # Replace any potential invalid characters in the suffix
        suffix2 = suffix2.replace('.', '-')

        # Construct the full file paths
        file1_path = pdf_files.get(file1)
        file2_path = pdf_files.get(file2)

        # Check if both PDF files exist
        if file1_path and file2_path:
            # Create subfolder if it doesn't exist
            subfolder_path = os.path.join(output_folder, subfolder_name)
            os.makedirs(subfolder_path, exist_ok=True)

            # Define the output file path with the abbreviated suffixes
            output_filename = f"{suffix1}_{file1_last8}_{file2_last8}_{suffix2}.pdf"
            output_path = os.path.join(subfolder_path, output_filename)

            # Merge the PDFs
            merge_pdfs([file1_path, file2_path], output_path)
            if log_callback:
                log_callback(f"Mesclando {file1} e {file2}")
            
            # Move the original PDF files to the 'arquivados' folder
            shutil.move(file1_path, os.path.join(archived_folder, os.path.basename(file1_path)))
            shutil.move(file2_path, os.path.join(archived_folder, os.path.basename(file2_path)))

            processed_files += 2  # Count each file processed

            # Update progress based on number of processed files
            if progress_callback:
                progress_callback(processed_files, total_files)

    # Print the remaining files count in each subfolder
    folder_counts = count_pdfs_in_subfolders(folder_path)
    for subfolder_path, count in folder_counts.items():
        if log_callback:
            log_callback(f"Sobraram {count} arquivos em {os.path.basename(subfolder_path)}")

    if log_callback:
        log_callback(f"Total de arquivos mesclados: {processed_files // 2}")

# Set up the GUI
if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
