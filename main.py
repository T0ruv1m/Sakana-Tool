import tkinter as tk
import subprocess
import os
import sys

# Define paths for subprograms
path = os.path.dirname(os.path.abspath(__file__))
Capture = "Sakana_Capture.py"
Fusion = "Sakana_Fusion.py"
xml = "extract_xml2.py"
install_packages = "install_packages.py"

Sakana_C_Dir = os.path.join(path, Capture)
Sakana_F_Dir = os.path.join(path, Fusion)
extract_xml2dir = os.path.join(path, xml)
install_packages_dir = os.path.join(path, install_packages)

def update_pip():
    try:
        # Running the command to upgrade pip
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
        print("Pip has been successfully updated.")
    except subprocess.CalledProcessError as e:
        print(f"Error occurred while updating pip: {e}")

def run_subprocess(script_path):
    # Hide the main window
    root.withdraw()
    try:
        # Run the subprocess and wait for it to complete
        subprocess.run(["python", script_path], check=True)
    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e}")
        # Optionally, add error handling here
    finally:
        # Show the main window again after the subprocess finishes
        root.deiconify()

def sub_program1():
    print("Subprogram 1 executed")
    run_subprocess(Sakana_C_Dir)

def sub_program2():
    print("Subprogram 2 executed")
    run_subprocess(Sakana_F_Dir)

# Create the main application window
root = tk.Tk()
root.title("Sakana Tool")
root.geometry("200x200")

# Define the path to the .xlsx file
xlsx_file = "../xlsx/extracted_data.xlsx"  # Replace with the actual path to the .xlsx file

# Run the subprocess for extract_xml2 only if the .xlsx file does not exist
if not os.path.exists(xlsx_file):
    run_subprocess(extract_xml2dir)

# Run the subprocess for installing dependencies
# run_subprocess(install_packages_dir)

# Create and place the buttons
button1 = tk.Button(root, text="Registrar Notas", command=sub_program1, width=20, height=2)
button1.pack(pady=20)

button2 = tk.Button(root, text="Mesclar Notas", command=sub_program2, width=20, height=2)
button2.pack(pady=20)

# Run the main event loop
root.mainloop()
