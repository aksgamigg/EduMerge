"""
EduText Automatic Dependency Installer
Automatically installs and configures all required dependencies
"""

import sys
import subprocess
import os
import platform
import urllib.request
import zipfile
import shutil
from pathlib import Path
import tempfile


class Color:
    """ANSI color codes for terminal output"""
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    BOLD = '\033[1m'
    END = '\033[0m'


class DependencyInstaller:
    """Handles automatic installation of all dependencies"""

    def __init__(self):
        self.system = platform.system()
        self.errors = []
        self.warnings = []
        self.installed = []

    def print_header(self, text):
        """Print formatted section header"""
        print(f"\n{Color.BOLD}{Color.BLUE}{'=' * 70}{Color.END}")
        print(f"{Color.BOLD}{Color.BLUE}{text.center(70)}{Color.END}")
        print(f"{Color.BOLD}{Color.BLUE}{'=' * 70}{Color.END}\n")

    def print_step(self, text):
        """Print step information"""
        print(f"\n{Color.CYAN}▶ {text}...{Color.END}")

    def print_success(self, text):
        """Print success message"""
        print(f"{Color.GREEN}✓ {text}{Color.END}")
        self.installed.append(text)

    def print_error(self, text):
        """Print error message"""
        print(f"{Color.RED}✗ {text}{Color.END}")
        self.errors.append(text)

    def print_warning(self, text):
        """Print warning message"""
        print(f"{Color.YELLOW}⚠ {text}{Color.END}")
        self.warnings.append(text)

    def run_command(self, cmd, description, check=True):
        """
        Run a shell command

        Args:
            cmd: Command to run (string or list)
            description: Description of what's being done
            check: Whether to check return code

        Returns:
            tuple: (success: bool, output: str)
        """
        try:
            if isinstance(cmd, str):
                cmd = cmd.split()

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                check=check
            )

            if result.returncode == 0:
                self.print_success(description)
                return True, result.stdout
            else:
                self.print_error(f"{description} - {result.stderr}")
                return False, result.stderr

        except subprocess.CalledProcessError as e:
            self.print_error(f"{description} - {e.stderr}")
            return False, str(e)
        except FileNotFoundError:
            self.print_error(f"{description} - Command not found")
            return False, "Command not found"
        except Exception as e:
            self.print_error(f"{description} - {str(e)}")
            return False, str(e)

    def check_pip(self):
        """Check if pip is available"""
        self.print_step("Checking pip installation")

        try:
            result = subprocess.run(
                [sys.executable, "-m", "pip", "--version"],
                capture_output=True,
                text=True
            )

            if result.returncode == 0:
                self.print_success("pip is available")
                return True
            else:
                self.print_error("pip is not available")
                return False
        except Exception as e:
            self.print_error(f"pip check failed: {e}")
            return False

    def install_pip_package(self, package, display_name=None):
        """
        Install a Python package using pip

        Args:
            package: Package name to install
            display_name: Display name (defaults to package name)

        Returns:
            bool: True if installation successful
        """
        display_name = display_name or package
        self.print_step(f"Installing {display_name}")

        try:
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", package, "--upgrade"],
                capture_output=True,
                text=True
            )

            if result.returncode == 0:
                self.print_success(f"{display_name} installed successfully")
                return True
            else:
                self.print_error(f"Failed to install {display_name}: {result.stderr}")
                return False

        except Exception as e:
            self.print_error(f"Failed to install {display_name}: {str(e)}")
            return False

    def install_python_packages(self):
        """Install all required Python packages"""
        self.print_header("INSTALLING PYTHON PACKAGES")

        packages = [
            ('Pillow', 'Pillow (Image Processing)'),
            ('ttkbootstrap', 'ttkbootstrap (Modern UI)'),
            ('python-docx', 'python-docx (Word Documents)'),
            ('pypdf', 'pypdf (PDF Reading)'),
            ('pdf2image', 'pdf2image (PDF Image Rendering)'),
        ]

        success_count = 0
        for package, display in packages:
            if self.install_pip_package(package, display):
                success_count += 1

        print(f"\n{Color.BOLD}Python packages: {success_count}/{len(packages)} installed{Color.END}")
        return success_count == len(packages)

    def install_poppler_windows(self):
        """Install Poppler on Windows"""
        self.print_header("INSTALLING POPPLER FOR WINDOWS")

        # Check if already installed
        try:
            result = subprocess.run(['pdftoppm', '-v'], capture_output=True, timeout=5)
            if result.returncode == 0 or result.stderr:
                self.print_success("Poppler already installed")
                return True
        except FileNotFoundError:
            pass
        except Exception:
            pass

        self.print_step("Downloading Poppler for Windows")

        # Poppler download URL (latest stable version)
        poppler_url = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip"

        temp_dir = tempfile.mkdtemp()
        zip_path = os.path.join(temp_dir, "poppler.zip")

        try:
            # Download Poppler
            print("  Downloading... (this may take a few minutes)")
            urllib.request.urlretrieve(poppler_url, zip_path)
            self.print_success("Poppler downloaded")

            # Extract Poppler
            self.print_step("Extracting Poppler")
            extract_path = "C:\\poppler"

            if os.path.exists(extract_path):
                self.print_warning(f"Directory {extract_path} already exists")
                response = input(f"  Remove existing directory? (y/n): ").lower()
                if response == 'y':
                    shutil.rmtree(extract_path)
                else:
                    extract_path = "C:\\poppler_new"
                    print(f"  Using alternative path: {extract_path}")

            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_path)

            self.print_success(f"Poppler extracted to {extract_path}")

            # Add to PATH
            self.print_step("Adding Poppler to system PATH")

            poppler_bin = os.path.join(extract_path, "poppler-24.08.0", "Library", "bin")

            if not os.path.exists(poppler_bin):
                # Try to find the bin directory
                for root, dirs, files in os.walk(extract_path):
                    if 'pdftoppm.exe' in files:
                        poppler_bin = root
                        break

            if not os.path.exists(poppler_bin):
                self.print_error(f"Could not find Poppler bin directory in {extract_path}")
                return False

            # Add to user PATH
            try:
                import winreg

                # Open registry key
                key = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    'Environment',
                    0,
                    winreg.KEY_ALL_ACCESS
                )

                # Get current PATH
                try:
                    current_path, _ = winreg.QueryValueEx(key, 'Path')
                except FileNotFoundError:
                    current_path = ''

                # Check if already in PATH
                if poppler_bin.lower() not in current_path.lower():
                    new_path = current_path + ';' + poppler_bin if current_path else poppler_bin
                    winreg.SetValueEx(key, 'Path', 0, winreg.REG_EXPAND_SZ, new_path)
                    winreg.CloseKey(key)

                    self.print_success("Poppler added to PATH")
                    self.print_warning("Please restart your terminal/IDE for PATH changes to take effect")
                else:
                    self.print_success("Poppler already in PATH")

                return True

            except Exception as e:
                self.print_error(f"Could not modify PATH automatically: {e}")
                print(f"\n{Color.YELLOW}Please add manually:{Color.END}")
                print(f"  1. Press Win + X and select 'System'")
                print(f"  2. Click 'Advanced system settings'")
                print(f"  3. Click 'Environment Variables'")
                print(f"  4. Edit 'Path' variable")
                print(f"  5. Add: {poppler_bin}")
                print(f"  6. Click OK and restart terminal/IDE")
                return False

        except Exception as e:
            self.print_error(f"Failed to install Poppler: {str(e)}")
            return False
        finally:
            # Cleanup
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

    def install_poppler_macos(self):
        """Install Poppler on macOS"""
        self.print_header("INSTALLING POPPLER FOR MACOS")

        # Check if Homebrew is installed
        self.print_step("Checking for Homebrew")

        try:
            result = subprocess.run(['brew', '--version'], capture_output=True, text=True)
            if result.returncode == 0:
                self.print_success("Homebrew is installed")
            else:
                self.print_error("Homebrew is not installed")
                print(f"\n{Color.YELLOW}Please install Homebrew first:{Color.END}")
                print('/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"')
                return False
        except FileNotFoundError:
            self.print_error("Homebrew is not installed")
            print(f"\n{Color.YELLOW}Please install Homebrew first:{Color.END}")
            print('/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"')
            return False

        # Install Poppler
        return self.run_command(
            ['brew', 'install', 'poppler'],
            "Installing Poppler via Homebrew"
        )[0]

    def install_poppler_linux(self):
        """Install Poppler on Linux"""
        self.print_header("INSTALLING POPPLER FOR LINUX")

        # Detect package manager
        if shutil.which('apt-get'):
            # Debian/Ubuntu
            print("Detected Debian/Ubuntu system")
            return self.run_command(
                ['sudo', 'apt-get', 'update'],
                "Updating package list"
            )[0] and self.run_command(
                ['sudo', 'apt-get', 'install', '-y', 'poppler-utils'],
                "Installing poppler-utils"
            )[0]

        elif shutil.which('dnf'):
            # Fedora
            print("Detected Fedora system")
            return self.run_command(
                ['sudo', 'dnf', 'install', '-y', 'poppler-utils'],
                "Installing poppler-utils"
            )[0]

        elif shutil.which('pacman'):
            # Arch
            print("Detected Arch system")
            return self.run_command(
                ['sudo', 'pacman', '-S', '--noconfirm', 'poppler'],
                "Installing poppler"
            )[0]

        elif shutil.which('yum'):
            # CentOS/RHEL
            print("Detected CentOS/RHEL system")
            return self.run_command(
                ['sudo', 'yum', 'install', '-y', 'poppler-utils'],
                "Installing poppler-utils"
            )[0]

        else:
            self.print_error("Could not detect package manager")
            print(f"\n{Color.YELLOW}Please install poppler-utils manually for your distribution{Color.END}")
            return False

    def install_poppler(self):
        """Install Poppler based on OS"""
        if self.system == "Windows":
            return self.install_poppler_windows()
        elif self.system == "Darwin":
            return self.install_poppler_macos()
        elif self.system == "Linux":
            return self.install_poppler_linux()
        else:
            self.print_warning(f"Unsupported OS: {self.system}")
            return False

    def verify_installation(self):
        """Verify all installations"""
        self.print_header("VERIFYING INSTALLATION")

        verification_tests = []

        # Test Python packages
        self.print_step("Testing Python packages")

        packages_to_test = [
            ('PIL', 'Pillow'),
            ('ttkbootstrap', 'ttkbootstrap'),
            ('docx', 'python-docx'),
            ('pypdf', 'pypdf'),
            ('pdf2image', 'pdf2image'),
        ]

        for module, display in packages_to_test:
            try:
                __import__(module)
                self.print_success(f"{display} import successful")
                verification_tests.append(True)
            except ImportError:
                self.print_error(f"{display} import failed")
                verification_tests.append(False)

        # Test Poppler
        self.print_step("Testing Poppler")

        try:
            result = subprocess.run(['pdftoppm', '-v'], capture_output=True, timeout=5)
            if result.returncode == 0 or result.stderr:
                self.print_success("Poppler is accessible")
                verification_tests.append(True)
            else:
                self.print_error("Poppler test failed")
                verification_tests.append(False)
        except FileNotFoundError:
            self.print_error("Poppler not found in PATH")
            self.print_warning("You may need to restart your terminal/IDE")
            verification_tests.append(False)
        except Exception as e:
            self.print_error(f"Poppler test failed: {e}")
            verification_tests.append(False)

        passed = sum(verification_tests)
        total = len(verification_tests)

        print(f"\n{Color.BOLD}Verification: {passed}/{total} tests passed{Color.END}")

        return all(verification_tests)

    def create_resource_files(self):
        """Create placeholder resource files if they don't exist"""
        self.print_header("CHECKING RESOURCE FILES")

        resources = ['logo.ico', 'logo.png', 'label.png']

        for resource in resources:
            if not os.path.exists(resource):
                self.print_warning(f"{resource} not found - app will work without it")
            else:
                self.print_success(f"{resource} found")

    def print_summary(self):
        """Print installation summary"""
        self.print_header("INSTALLATION SUMMARY")

        print(f"\n{Color.BOLD}Successfully Installed:{Color.END}")
        if self.installed:
            for item in self.installed:
                print(f"  {Color.GREEN}✓{Color.END} {item}")
        else:
            print("  None")

        if self.warnings:
            print(f"\n{Color.BOLD}Warnings:{Color.END}")
            for warning in self.warnings:
                print(f"  {Color.YELLOW}⚠{Color.END} {warning}")

        if self.errors:
            print(f"\n{Color.BOLD}Errors:{Color.END}")
            for error in self.errors:
                print(f"  {Color.RED}✗{Color.END} {error}")

        print("\n" + "=" * 70)

        if not self.errors:
            print(f"{Color.GREEN}{Color.BOLD}✓ INSTALLATION COMPLETE!{Color.END}")
            print(f"{Color.GREEN}Your EduText application is ready to use.{Color.END}")
        else:
            print(f"{Color.YELLOW}{Color.BOLD}⚠ INSTALLATION COMPLETED WITH ERRORS{Color.END}")
            print(f"{Color.YELLOW}Please review the errors above and fix them manually.{Color.END}")

        print("=" * 70 + "\n")

    def run(self):
        """Run the complete installation process"""
        self.print_header("EDUTEXT AUTOMATIC INSTALLER")

        print(f"System: {self.system}")
        print(f"Python: {sys.version}")
        print(f"\nThis installer will:")
        print("  1. Check pip availability")
        print("  2. Install required Python packages")
        print("  3. Install Poppler (for PDF rendering)")
        print("  4. Verify all installations")

        response = input(f"\n{Color.BOLD}Continue? (y/n): {Color.END}").lower()
        if response != 'y':
            print("Installation cancelled.")
            return

        # Check pip
        if not self.check_pip():
            print(f"\n{Color.RED}pip is required but not available. Please install pip first.{Color.END}")
            return

        # Install Python packages
        self.install_python_packages()

        # Install Poppler
        print(f"\n{Color.BOLD}Install Poppler for PDF image rendering?{Color.END}")
        print("(Required for best PDF viewing experience)")
        response = input("Install Poppler? (y/n): ").lower()

        if response == 'y':
            self.install_poppler()
        else:
            self.print_warning("Skipping Poppler installation")
            self.print_warning("PDF text extraction fallback will be used")

        # Check resource files
        self.create_resource_files()

        # Verify installation
        self.verify_installation()

        # Print summary
        self.print_summary()


def main():
    """Main entry point"""

    # Check if running with appropriate privileges
    if platform.system() == "Linux" and os.geteuid() == 0:
        print(f"{Color.YELLOW}Warning: Running as root. Consider running as regular user.{Color.END}")

    installer = DependencyInstaller()

    try:
        installer.run()
    except KeyboardInterrupt:
        print(f"\n\n{Color.YELLOW}Installation interrupted by user.{Color.END}")
    except Exception as e:
        print(f"\n\n{Color.RED}Unexpected error: {e}{Color.END}")
        import traceback
        traceback.print_exc()
    finally:
        input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()