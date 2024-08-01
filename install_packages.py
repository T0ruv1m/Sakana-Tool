import subprocess
import sys
import os
import shutil

def is_tool(name):
    """Check if a tool is installed and available in PATH."""
    from shutil import which
    return which(name) is not None

def install_pip():
    """Ensure pip is installed."""
    try:
        # Try importing pip
        import pip
        print("pip is already installed.")
    except ImportError:
        print("pip is not installed. Installing pip...")
        # Download get-pip.py script and execute it to install pip
        try:
            subprocess.check_call([sys.executable, '-m', 'ensurepip', '--upgrade'])
            print("pip installed successfully.")
        except subprocess.CalledProcessError as e:
            print(f"Error occurred while installing pip: {e}")
            # Fall back to get-pip.py method if ensurepip fails
            import urllib.request
            get_pip_url = 'https://bootstrap.pypa.io/get-pip.py'
            get_pip_script = 'get-pip.py'
            urllib.request.urlretrieve(get_pip_url, get_pip_script)
            subprocess.check_call([sys.executable, get_pip_script])
            os.remove(get_pip_script)
            print("pip installed successfully via get-pip.py.")

def install_packages(requirements_file='requirements.txt'):
    """Install packages listed in the requirements file."""
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

def setup_environment():
    """Setup Python environment by ensuring Python, pip are installed, and packages are up to date."""
    if is_tool("python"):
        print(f"Python is installed at: {shutil.which('python')}")
    else:
        print("Python is not installed. Please install Python first.")
        return

    # Ensure pip is installed
    install_pip()

    # Install required packages
    install_packages()

if __name__ == "__main__":
    setup_environment()
