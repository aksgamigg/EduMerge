# Student Hackpad 2025

---

## EduMerge Text Editor

**EduMerge** is an advanced text editor designed for students and educators. It combines a powerful text editing environment with a customizable **mail merge system**, enabling users to create, edit, and automate documents efficiently. Built with Python with **ttkbootstrap**, EduMerge is lightweight, and versatile.

---

## Features

- **Full-Featured Text Editor**
    - Rich text editing (bold, italics, font size adjustments)
    - Syntax-aware editing for multiple document types
    - Supports `.txt`, `.docx` and `.pdf` files
- **Mail Merge Integration**
    - Easily import CSV or Excel files
    - Generate personalized documents for multiple recipients
    - Supports font and style formatting for merged documents
- **Document Export**
    - Export to **DOCX** and **PDF**
    - Maintain styles and formatting
- **Custom Windows & Splash Screens**
    - Modern UI using ttkbootstrap
    - Splash screen on startup for a professional look
    - Modal windows for smooth workflows
- **User-Friendly Interface**
    - Intuitive layout with buttons, menus, and table views
    - Side-by-side editing and previews
    - Easy-to-use dialogs for file management, color selection, and notifications
- **Future-Proof Design**
    - Built with OOP for easy extension
    - Designed to integrate additional features like Markdown support, AI-driven suggestions, and more

---

## Installation

**Note:** For user-level installation, admin rights are usually not required. Admin rights may be needed only for system-level PATH modifications.

1. **Clone the repository:**
    
    ```bash
    git clone https://github.com/aksgamigg/EduMerge.git
    cd EduMerge
    ```
    
2. **Install dependencies:**
    
    ```bash
    python dependency-installer.py
    # An installer for all dependencies that also configures poppler to add it to path for you
    
    ```
    
    Required libraries include:
    
    - `ttkbootstrap`
    - `Pillow`
    - `python-docx`
    - `pdf2image`
    - `pypdf` and many other standard python libraries
    
    Restart your terminal after running the installer to ensure all dependencies are recognized
    
    **Note:** If the installer fails to recognize a dependency, restart your terminal and run `diagnostics.py` to troubleshoot.
    
    ```bash
    python diagnostics.py
    # A troubleshooter to see if all your dependencies were installed correctly.
    ```
    
3. **Run EduMerge:**
    
    ```bash
    python EduMerge.py
    # The main Python file
    ```
    

---

## Usage

1. Open **EduMerge**.
2. Use the text editor to create or open documents.
3. For **mail merge**, import a txt file and see the magic.
4. Merge documents and export as DOCX or PDF.
5. Customize your workspace using themes and font settings.

---

## Project Architecture

- **EduMerge.py** – Entry point of the application
- **dependency-installer.py** – Installs dependencies for the app
- **diagnostics.py** – Restart terminal after the installer to check if all the dependencies were installed

---

## Future Enhancements

- Direct Markdown editing and preview
- Integrated AI suggestions for text and mail merge automation
- Multi-language support
- Cloud sync and document collaboration

---

## Contributing

EduMerge is an open-source project for students and educators. Contributions are welcome!

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/my-feature`)
3. Commit your changes (`git commit -m 'Add new feature'`)
4. Push to the branch (`git push origin feature/my-feature`)
5. Open a Pull Request

---

## License

This project is licensed under the MIT License. See the LICENSE file for details.

---

## Contact

Created by **Akshaj Goel** –

Email: akshajgoel@bnpsramvihar.edu.in

GitHub: [https://github.com/aksgamigg](https://github.com/aksgamigg)