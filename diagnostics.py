"""
Complete Dependency Diagnostic Tool for EduText Application
Checks all required libraries, system tools, and configuration
"""

import sys
import subprocess
import os
import platform
from importlib import import_module
from pathlib import Path

class Color:
    """ANSI color codes for terminal output"""
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    BOLD = '\033[1m'
    END = '\033[0m'

def print_header(text):
    """Print formatted section header"""
    print(f"\n{Color.BOLD}{Color.BLUE}{'=' * 70}{Color.END}")
    print(f"{Color.BOLD}{Color.BLUE}{text.center(70)}{Color.END}")
    print(f"{Color.BOLD}{Color.BLUE}{'=' * 70}{Color.END}\n")

def print_subheader(text):
    """Print formatted subsection header"""
    print(f"\n{Color.BOLD}{'-' * 70}{Color.END}")
    print(f"{Color.BOLD}{text}{Color.END}")
    print(f"{Color.BOLD}{'-' * 70}{Color.END}")

def check_success(name, version=""):
    """Print success message"""
    version_text = f" (version {version})" if version else ""
    print(f"{Color.GREEN}✓{Color.END} {name}: {Color.GREEN}INSTALLED{Color.END}{version_text}")

def check_failure(name, message=""):
    """Print failure message"""
    msg_text = f" - {message}" if message else ""
    print(f"{Color.RED}✗{Color.END} {name}: {Color.RED}NOT FOUND{Color.END}{msg_text}")

def check_warning(name, message):
    """Print warning message"""
    print(f"{Color.YELLOW}⚠{Color.END} {name}: {Color.YELLOW}WARNING{Color.END} - {message}")

def check_python_library(module_name, display_name=None):
    """
    Check if a Python library is installed

    Args:
        module_name: Name of the module to import
        display_name: Optional display name (defaults to module_name)

    Returns:
        tuple: (success: bool, version: str or None)
    """
    display_name = display_name or module_name

    try:
        module = import_module(module_name)

        # Try to get version
        version = None
        for attr in ['__version__', 'VERSION', 'version']:
            if hasattr(module, attr):
                version = getattr(module, attr)
                if isinstance(version, tuple):
                    version = '.'.join(map(str, version))
                break

        check_success(display_name, version)
        return True, version

    except ImportError as e:
        check_failure(display_name, str(e))
        return False, None

def check_system_command(cmd, name=None):
    """
    Check if a system command is available

    Args:
        cmd: Command to check
        name: Display name (defaults to cmd)

    Returns:
        bool: True if command is available
    """
    name = name or cmd

    try:
        result = subprocess.run(
            [cmd, '--version'],
            capture_output=True,
            text=True,
            timeout=5
        )

        output = (result.stdout + result.stderr).strip()

        # Try -v flag if --version doesn't work
        if result.returncode != 0:
            result = subprocess.run(
                [cmd, '-v'],
                capture_output=True,
                text=True,
                timeout=5
            )
            output = (result.stdout + result.stderr).strip()

        if result.returncode == 0 or output:
            # Extract version if available
            version_line = output.split('\n')[0][:80]
            check_success(name, version_line if version_line else "")
            return True
        else:
            check_failure(name)
            return False

    except FileNotFoundError:
        check_failure(name, "Command not found in PATH")
        return False
    except subprocess.TimeoutExpired:
        check_warning(name, "Command found but timed out")
        return True
    except Exception as e:
        check_failure(name, str(e))
        return False

def check_file_exists(filepath, description):
    """
    Check if a file exists

    Args:
        filepath: Path to check
        description: Description of the file

    Returns:
        bool: True if file exists
    """
    if Path(filepath).exists():
        check_success(description, f"Found at {filepath}")
        return True
    else:
        check_failure(description, f"Not found at {filepath}")
        return False

def main():
    """Run complete diagnostic"""

    print_header("EDUTEXT DEPENDENCY DIAGNOSTIC TOOL")

    # System information
    print_subheader("System Information")
    print(f"Operating System: {platform.system()} {platform.release()}")
    print(f"Platform: {platform.platform()}")
    print(f"Architecture: {platform.machine()}")
    print(f"Python Version: {sys.version}")
    print(f"Python Executable: {sys.executable}")

    # Required Python libraries
    print_subheader("Core Python Libraries (Required)")

    core_libs = {
        'tkinter': 'tkinter (Built-in GUI)',
        'PIL': 'Pillow (Image processing)',
        'ttkbootstrap': 'ttkbootstrap (Modern UI)',
    }

    core_results = {}
    for module, display in core_libs.items():
        success, version = check_python_library(module, display)
        core_results[module] = success

    # Document processing libraries
    print_subheader("Document Processing Libraries (Required)")

    doc_libs = {
        'csv': 'csv (Built-in)',
        'docx': 'python-docx (Word documents)',
        'pypdf': 'pypdf (PDF reading)',
    }

    doc_results = {}
    for module, display in doc_libs.items():
        success, version = check_python_library(module, display)
        doc_results[module] = success

    # Optional libraries
    print_subheader("Optional Libraries (Enhanced Features)")

    optional_libs = {
        'pdf2image': 'pdf2image (PDF image rendering)',
    }

    optional_results = {}
    for module, display in optional_libs.items():
        success, version = check_python_library(module, display)
        optional_results[module] = success

    # System tools (for PDF image rendering)
    print_subheader("System Tools (For PDF Image Rendering)")

    poppler_cmds = ['pdftoppm', 'pdfinfo', 'pdftocairo']
    poppler_found = False

    for cmd in poppler_cmds:
        if check_system_command(cmd, f"Poppler - {cmd}"):
            poppler_found = True

    # Check for resource files
    print_subheader("Resource Files")

    resource_files = {
        'logo.ico': 'Application Icon',
        'logo.png': 'Logo Image',
        'label.png': 'Label Image',
    }

    resource_results = {}
    for filename, description in resource_files.items():
        resource_results[filename] = check_file_exists(filename, description)

    # PATH information
    print_subheader("System PATH")

    path_dirs = os.environ.get('PATH', '').split(os.pathsep)
    print(f"\nTotal directories in PATH: {len(path_dirs)}")

    # Highlight important paths
    poppler_paths = [p for p in path_dirs if 'poppler' in p.lower()]
    if poppler_paths:
        print(f"\n{Color.GREEN}Poppler-related paths found:{Color.END}")
        for p in poppler_paths:
            print(f"  • {p}")

    # Summary
    print_header("DIAGNOSTIC SUMMARY")

    total_core = len(core_libs)
    passed_core = sum(core_results.values())

    total_doc = len(doc_libs)
    passed_doc = sum(doc_results.values())

    total_optional = len(optional_libs)
    passed_optional = sum(optional_results.values())

    total_resources = len(resource_files)
    passed_resources = sum(resource_results.values())

    print(f"\n{Color.BOLD}Core Libraries:{Color.END} {passed_core}/{total_core} installed")
    print(f"{Color.BOLD}Document Libraries:{Color.END} {passed_doc}/{total_doc} installed")
    print(f"{Color.BOLD}Optional Libraries:{Color.END} {passed_optional}/{total_optional} installed")
    print(f"{Color.BOLD}System Tools (Poppler):{Color.END} {'✓ Found' if poppler_found else '✗ Not Found'}")
    print(f"{Color.BOLD}Resource Files:{Color.END} {passed_resources}/{total_resources} found")

    # Overall status
    all_required = (passed_core == total_core and
                   passed_doc == total_doc)

    print("\n" + "=" * 70)

    if all_required:
        print(f"{Color.GREEN}{Color.BOLD}✓ ALL REQUIRED DEPENDENCIES INSTALLED{Color.END}")
        print(f"{Color.GREEN}Your EduText application should work!{Color.END}")

        if not poppler_found or passed_optional < total_optional:
            print(f"\n{Color.YELLOW}Note: PDF image rendering requires pdf2image and Poppler.{Color.END}")
            print(f"{Color.YELLOW}Text extraction fallback will be used for PDFs.{Color.END}")
    else:
        print(f"{Color.RED}{Color.BOLD}✗ MISSING REQUIRED DEPENDENCIES{Color.END}")
        print(f"{Color.RED}Please install missing packages before running EduText.{Color.END}")

    print("=" * 70)

    # Installation instructions
    if not all_required or passed_optional < total_optional or not poppler_found:
        print_header("INSTALLATION INSTRUCTIONS")

        missing_packages = []

        # Core packages
        for module, success in core_results.items():
            if not success:
                if module == 'PIL':
                    missing_packages.append('Pillow')
                elif module == 'tkinter':
                    print(f"{Color.YELLOW}tkinter is built-in but may need system package:{Color.END}")
                    if platform.system() == "Linux":
                        print("  Ubuntu/Debian: sudo apt-get install python3-tk")
                        print("  Fedora: sudo dnf install python3-tkinter")
                else:
                    missing_packages.append(module)

        # Document packages
        for module, success in doc_results.items():
            if not success and module != 'csv':
                if module == 'docx':
                    missing_packages.append('python-docx')
                else:
                    missing_packages.append(module)

        # Optional packages
        for module, success in optional_results.items():
            if not success:
                missing_packages.append(module)

        if missing_packages:
            print(f"\n{Color.BOLD}Install missing Python packages:{Color.END}")
            print(f"pip install {' '.join(missing_packages)}")

        if not poppler_found:
            print(f"\n{Color.BOLD}Install Poppler (for PDF image rendering):{Color.END}\n")

            if platform.system() == "Windows":
                print("Windows:")
                print("  1. Download from: https://github.com/oschwartz10612/poppler-windows/releases/")
                print("  2. Extract to C:\\poppler")
                print("  3. Add C:\\poppler\\Library\\bin to PATH")
                print("  4. Restart terminal/IDE")

            elif platform.system() == "Darwin":
                print("macOS:")
                print("  brew install poppler")

            else:
                print("Linux:")
                print("  Ubuntu/Debian: sudo apt-get install poppler-utils")
                print("  Fedora: sudo dnf install poppler-utils")
                print("  Arch: sudo pacman -S poppler")

        if not all(resource_results.values()):
            print(f"\n{Color.YELLOW}Missing resource files:{Color.END}")
            for filename, found in resource_results.items():
                if not found:
                    print(f"  • {filename}")
            print("\nThese are optional. App will work without icons/images.")

    print("\n" + "=" * 70)
    print(f"{Color.BOLD}Diagnostic complete!{Color.END}")
    print("=" * 70 + "\n")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\n{Color.YELLOW}Diagnostic interrupted by user.{Color.END}")
    except Exception as e:
        print(f"\n\n{Color.RED}Error during diagnostic: {e}{Color.END}")
    finally:
        input("\nPress Enter to exit...")