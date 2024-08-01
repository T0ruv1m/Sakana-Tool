import subprocess
import sys
import os

def install_packages(requirements_file='requirements.txt'):
    requirements_path = os.path.join(os.path.dirname(__file__), requirements_file)
    print(f"Procurando por: {requirements_path}")
    try:
        with open(requirements_path, 'r') as file:
            packages = file.readlines()
            for package in packages:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', package.strip()])
        print("All packages installed successfully.")
    except FileNotFoundError:
        print(f"Error: {requirements_file} not found.")
    except subprocess.CalledProcessError as e:
        print(f"Error occurred while installing packages: {e}")

if __name__ == "__main__":
    install_packages()
