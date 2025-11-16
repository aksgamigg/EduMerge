# =============================================================================
# IMPORTS
# =============================================================================
from tkinter import filedialog, messagebox, colorchooser
from tkinter import Text, END, LEFT, RIGHT, BOTTOM, TOP, X, Y, BOTH, W, E
import ttkbootstrap as ttk
from ttkbootstrap import Toplevel, Label, PhotoImage
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap.widgets.tableview import Tableview        # fixed deprecated import
from ttkbootstrap.constants import *
from pathlib import Path
import os, platform, subprocess
from PIL import Image, ImageTk
import csv
from docx import Document
from docx.shared import Pt, RGBColor
import pypdf

# =============================================================================
# CUSTOM DIALOG BOXES
# =============================================================================
class StartingWindow(Toplevel):
    """Starting Screen"""
    def __init__(self):
        super().__init__()
        self.result = None

        self.geometry("960x540")
        self.title("Mail Merge")
        self.resizable(True, True)
        self.attributes("-topmost", True)
        self.overrideredirect(True) # Removes Title Bar

        self.place_window_center()

        self.photo = PhotoImage(file="logo.png")
        image_label = Label(self, image=self.photo)
        image_label.place(x=250, y=45)

        text = ttk.Label(self, text="EduText: A Student Software", font=("Segoe UI", 20, "bold"))
        text.pack(side = "top", pady = 20)
        text2 = ttk.Label(self, text="By: Akshaj Goel", font=("Segoe UI", 14))
        text2.pack(side="bottom", pady=20)

        buttons = ["📧Mail Merge", "📝Text Editor", "❌Exit"]

        bottom_frame = ttk.Frame(self)
        bottom_frame.place(relx=0, rely=1, anchor="sw", relwidth=1)

        for text in buttons:
            btn1 = ttk.Button(bottom_frame, text=text, command=lambda v=text: self.finish(v))

            btn1.pack(side="left", expand=True, fill="x", padx=10, pady=10)

        self.bind("<Button-1>", self.start_move)
        self.bind("<B1-Motion>", self.do_move)

        self.mainloop()

    def start_move(self, event):
        self._x = event.x
        self._y = event.y

    def do_move(self, event):
        dx = event.x - self._x
        dy = event.y - self._y
        x = self.winfo_x() + dx
        y = self.winfo_y() + dy
        self.geometry(f"+{x}+{y}")

    def finish(self, value):
        self.result = value
        self.destroy()
        self.quit()

class CustomBox(Toplevel):
    def __init__(self, buttons, message, image_y):
        super().__init__()
        self.result = None

        self.geometry("480x350")
        self.title("Mail Merge")
        self.resizable(True, True)
        self.attributes("-topmost", True)
        self.overrideredirect(True)

        self.photo = PhotoImage(file="label.png")
        image_label = Label(self, image=self.photo)
        image_label.place(x=128, y= image_y)

        bottom_frame = ttk.Frame(self)
        bottom_frame.place(relx=0, rely=1, anchor="sw", relwidth=1)

        for text in buttons:
            btn1 = ttk.Button(bottom_frame, text=text, command=lambda v=text: self.finish(v))

            btn1.pack(side="left", expand=True, fill="x", padx=5, pady=5)

        text = ttk.Label(self, text=message, font=("Segoe UI", 10))
        text.pack(side="top", pady = 15)

        self.place_window_center()

        self.bind("<Button-1>", self.start_move)
        self.bind("<B1-Motion>", self.do_move)

        self.mainloop()

    def start_move(self, event):
        self._x = event.x
        self._y = event.y

    def do_move(self, event):
        dx = event.x - self._x
        dy = event.y - self._y
        x = self.winfo_x() + dx
        y = self.winfo_y() + dy
        self.geometry(f"+{x}+{y}")

    def finish(self, value):
        self.result = value
        self.destroy()
        self.quit()

class IntBox(Toplevel):
    def __init__(self, buttons, message):
        super().__init__()

        self.result = None
        self.choice = ""

        self.spinbox = ttk.Spinbox(
            self,
            from_=2,
            to=1000,
            width=10
        )
        self.spinbox.pack(side="bottom", pady=60)


        self.geometry("580x400")
        self.title("Mail Merge")
        self.resizable(True, True)
        self.attributes("-topmost", True)
        self.overrideredirect(True)

        self.photo = PhotoImage(file="label.png")
        image_label = Label(self, image=self.photo)
        image_label.place(x=168, y=40)

        bottom_frame = ttk.Frame(self)
        bottom_frame.place(relx=0, rely=1, anchor="sw", relwidth=1)

        for text in buttons:
            btn1 = ttk.Button(bottom_frame, text=text, command=lambda v=text: self.finish(v))

            btn1.pack(side="left", expand=True, fill="x", padx=5, pady=5)

        text = ttk.Label(self, text=message, font=("Segoe UI", 10))
        text.place(relx=0.5, y=20, anchor='n')

        self.place_window_center()

        self.bind("<Button-1>", self.start_move)
        self.bind("<B1-Motion>", self.do_move)

        self.mainloop()

    def start_move(self, event):
        self._x = event.x
        self._y = event.y

    def do_move(self, event):
        dx = event.x - self._x
        dy = event.y - self._y
        x = self.winfo_x() + dx
        y = self.winfo_y() + dy
        self.geometry(f"+{x}+{y}")

    def finish(self, value):
        self.result = int(self.spinbox.get())
        self.choice = value
        self.destroy()
        self.quit()


class StringBox(Toplevel):
    def __init__(self, buttons, message):
        super().__init__()

        self.result = None
        self.choice = ""

        self.entry = ttk.Entry(self)
        self.entry.pack(side="bottom", pady=60)

        self.geometry("580x400")
        self.title("Mail Merge")
        self.resizable(True, True)
        self.attributes("-topmost", True)
        self.overrideredirect(True)

        self.photo = PhotoImage(file="label.png")
        image_label = Label(self, image=self.photo)
        image_label.place(x=168, y=40)

        bottom_frame = ttk.Frame(self)
        bottom_frame.place(relx=0, rely=1, anchor="sw", relwidth=1)

        for text in buttons:
            btn1 = ttk.Button(bottom_frame, text=text, command=lambda v=text: self.finish(v))

            btn1.pack(side="left", expand=True, fill="x", padx=5, pady=5)

        text = ttk.Label(self, text=message, font=("Segoe UI", 10))
        text.place(relx=0.5, y=20, anchor='n')

        self.place_window_center()

        self.bind("<Button-1>", self.start_move)
        self.bind("<B1-Motion>", self.do_move)

        self.mainloop()

    def start_move(self, event):
        self._x = event.x
        self._y = event.y

    def do_move(self, event):
        dx = event.x - self._x
        dy = event.y - self._y
        x = self.winfo_x() + dx
        y = self.winfo_y() + dy
        self.geometry(f"+{x}+{y}")

    def finish(self, value):
        self.result = self.entry.get()
        self.choice = value
        self.destroy()
        self.quit()

class TextBox(Toplevel):
    def __init__(self, buttons, message):
        super().__init__()

        self.result = None
        self.choice = ""

        self.text_box = ttk.Text(self, width=300, height=100)
        self.text_box.pack(padx=50, pady=50)

        self.geometry("960x540")
        self.title("Mail Merge")
        self.resizable(True, True)
        self.attributes("-topmost", True)
        self.overrideredirect(True)

        bottom_frame = ttk.Frame(self)
        bottom_frame.place(relx=0, rely=1, anchor="sw", relwidth=1)

        for text in buttons:
            btn1 = ttk.Button(bottom_frame, text=text, command=lambda v=text: self.finish(v))

            btn1.pack(side="left", expand=True, fill="x", padx=5, pady=5)

        text = ttk.Label(self, text=message, font=("Segoe UI", 10))
        text.place(relx=0.5, y=20, anchor='n')

        self.place_window_center()

        self.bind("<Button-1>", self.start_move)
        self.bind("<B1-Motion>", self.do_move)

        self.mainloop()

    def start_move(self, event):
        self._x = event.x
        self._y = event.y

    def do_move(self, event):
        dx = event.x - self._x
        dy = event.y - self._y
        x = self.winfo_x() + dx
        y = self.winfo_y() + dy
        self.geometry(f"+{x}+{y}")

    def finish(self, value):
        self.result = self.get_text()
        self.choice = value
        self.destroy()
        self.quit()

    def get_text(self):
        content = self.text_box.get("1.0", "end-1c")  # '1.0' = start, 'end-1c' = remove final newline
        return content




# =============================================================================
# LETTER CLASS
# =============================================================================
class Letter:
    def __init__(self):
        self.letter_body = ""
        self.output_directory = ""
        self.letter_file_location = ""
        self.is_placeholder_not_present = False

    def show_placeholder_instructions(self):
        """Displays instructions about using the [name] placeholder."""
        Messagebox.show_info(title="Mail Merge",
                             message="Make sure that you've entered the [name] placeholder in your letter")

    def generate_personalized_letters(self):
        """
        Creates personalized letter files for each recipient.
        Replaces [name] placeholder with actual names and saves files.
        """

        CustomBox(buttons=["Continue", "Exit"],
                  message="Now please select the folders in which you want to save your mails.",
                  image_y=40
                  )

        self.output_directory = filedialog.askdirectory(
            title="Where do you want to save the mails?"
        )

        # Create individual letter file for each recipient
        for recipient in name_manager.recipient_names:
            try:
                output_file = open(f"{self.output_directory}/{recipient}'s Mail.txt", "r")
                output_file.close()

                with open(f"{self.output_directory}/{recipient}'s Mail(1).txt", "r") as output_file:
                    personalized_content = self.letter_body.replace("[name]", recipient)
                    output_file.write(personalized_content)

            except FileNotFoundError:
                with open(f"{self.output_directory}/{recipient}'s Mail.txt", "r") as output_file:
                    personalized_content = self.letter_body.replace("[name]", recipient)
                    output_file.write(personalized_content)

        CustomBox(["Exit"], "Your Mail Merge was successful, congratulations!", 40)

        # Open mail folder
        if platform.system() == "Windows":
            os.startfile(self.output_directory)
        elif platform.system() == "Darwin":  # macOS
            subprocess.call(["open", self.output_directory])
        else:
            subprocess.call(["xdg-open", self.output_directory])

    def process_letter_content(self):
        """
        Handles letter content input based on user's choice (Browse or Type).
        Validates placeholder presence and triggers letter generation.
        """

        letter_choice_dialog = CustomBox(["Browse", "Type Letter Content"],
                                         "How do you want to enter the letter content?\nAs a file or type it?",
                                         50)
        file_input_choice = letter_choice_dialog.result

        if file_input_choice is None:
            show_exit_confirmation(self.process_letter_content)
            return

        if file_input_choice == "Browse":
            self.show_placeholder_instructions()

            # Get letter file from user
            self.letter_file_location = filedialog.askopenfilename(
                title="Select your mail",
                filetypes=(("Text files", "*.txt"),)
            )

            if not self.letter_file_location:
                show_exit_confirmation(self.process_letter_content)
                return

            with open(self.letter_file_location) as letter_file:
                self.letter_body = letter_file.read()

                # Validate [name] placeholder exists
                name_place = self.letter_body.find("[name]")
                if name_place < 0:
                    CustomBox(["Try Again"], "Warning: Placeholder '[name]' not present in file.", 40)
                    self.is_placeholder_not_present = True
                else:
                    self.is_placeholder_not_present = False

                if self.is_placeholder_not_present:
                    # Re-prompt for file input method
                    self.process_letter_content()
                    return

                self.generate_personalized_letters()

        elif file_input_choice == "Type Letter Content":
            self.show_placeholder_instructions()

            # Get letter content via text input
            letter_body_dialog = TextBox(["Continue", "Exit"], "Enter the letter Content")
            self.letter_body = letter_body_dialog.result
            continue_choice = letter_body_dialog.choice


            if continue_choice == "Exit":
                show_exit_confirmation(self.process_letter_content)
                return

            elif continue_choice == "":
                Messagebox.show_warning(title="Mail Merge", message="The letter is empty, please try again.")
                self.process_letter_content()
                return

            # Validate [name] placeholder exists
            name_place = self.letter_body.find("[name]")
            if name_place < 0:
                Messagebox.show_warning(
                    title="Mail Merge",
                    message="Warning: Placeholder '[name]' is not in the letter body"
                )
                self.is_placeholder_not_present = True
            else:
                self.is_placeholder_not_present = False

            if self.is_placeholder_not_present:
                # Re-prompt for file input method
                self.process_letter_content()
                return

            self.generate_personalized_letters()


# =============================================================================
# NAME MANAGER CLASS
# =============================================================================
class NameManager:
    def __init__(self):
        self.recipient_names = []
        self.name_input_method = ""
        self.names_file_location = ""
        self.name_input_method_prompt = "Do you want to enter names manually, or through a text file?"
        self.recipient_count_prompt = "How many people do you want send a mail to?"
        self.name_prompt = "Please enter the name:"
        self.total_recipients = 0

    def name_collection(self):
        """Ask user for name input method"""
        while True:
            name_input_method_dialog = CustomBox(buttons=["Manual Insert", "Text File", "Exit"],
                                                 message=self.name_input_method_prompt,
                                                 image_y=40)
            self.name_input_method = name_input_method_dialog.result

            if self.name_input_method is None:
                exit_confirmation()

            if self.name_input_method == "Manual Insert":
                self.ask_recipient_count()
                self.collect_names_manually()

            elif self.name_input_method == "Text File":
                self.names_text_file_processing()

            else:
                exit_confirmation()

            self.confirm_names_proceed()
            break


    def ask_recipient_count(self):
        """Prompts user to enter the number of recipients."""
        while True:
            total_recipients_dialog_box = IntBox(["Continue", "Exit"], self.recipient_count_prompt)
            self.total_recipients = total_recipients_dialog_box.result

            # Handle cancellation of recipient count dialog
            if total_recipients_dialog_box.choice == "Exit":
                self.total_recipients = 0
                exit_confirmation()

            if self.total_recipients >= 2:
                break

    def collect_names_manually(self):
        """
        Collects names from user through dialog boxes.
        Validates that each name contains only alphabetic characters, spaces, and hyphens.
        """

        for _ in range(self.total_recipients - len(self.recipient_names)):
            while True:
                entered_name_dialog = StringBox(["Continue", "Exit"], self.name_prompt)
                entered_name = entered_name_dialog.result
                name_dialog_exit_choice = entered_name_dialog.choice

                # Handle dialog cancellation
                if name_dialog_exit_choice == "Exit":
                    show_exit_confirmation(self.collect_names_manually)
                    return

                # Validate name contains only letters, spaces, and hyphens
                if entered_name.replace(" ", "").replace("-", "").isalpha():
                    self.name_prompt = "Please enter the name:"
                    self.recipient_names.append(entered_name.title())
                    break
                else:
                    self.name_prompt = "Please enter a valid name:"

    def ask_name_file(self):
        while True:
            self.names_file_location = ""
            self.names_file_location = filedialog.askopenfilename(
                title="Select the names text file",
                filetypes=(("Text files", "*.txt"),)
            )

            if not self.names_file_location:
                exit_confirmation()
            else:
                break

    def names_text_file_processing(self):
        """Get names file from user"""
        while True:
            self.ask_name_file()

            try:
                with open(self.names_file_location) as names_file:
                    names_content = names_file.read()

            except Exception as e:
                Messagebox.show_warning(title = "Error occurred",
                                        message = f"Error: {e}")
                self.name_collection()
                return

            # Validate comma separation
            if ", " not in names_content:
                Messagebox.show_warning(
                    title="Mail Merge",
                    message="Please enter the names separated by comma."
                )

            elif ", " in names_content:
                break

        self.recipient_names = names_content.split(", ")

    def confirm_names_proceed(self):
        """
        Shows collected names to user for confirmation.
        Allows re-entry or continuation to file selection.
        Returns user's file input choice (Browse or Type).
        """
        global confirmation_choice

        names_confirmation_dialog = CustomBox(
            message=f"Names entered:\n{', '.join(self.recipient_names)}",
            buttons=["Continue", "Re-enter Names", "Exit"],
            image_y=60
        )

        confirmation_choice = names_confirmation_dialog.result

        if confirmation_choice == "Continue":
            return confirmation_choice

        elif confirmation_choice == "Re-enter Names":
            self.recipient_names = []
            self.name_collection()
            return
        else:
            exit_confirmation()

# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def center_window(window):
    """Centers the given window on the screen."""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')


def show_exit_confirmation(callback_function):
    """
    Shows exit confirmation dialog and handles user's choice.
    If user chooses 'No', calls the provided callback function.
    """
    exit_dialog = CustomBox(["Yes", "No"], "Do you want to exit the app?", 40)

    user_choice = exit_dialog.result

    if user_choice == "Yes":
        exit()
    else:
        callback_function()

def exit_confirmation():
    """
    Shows exit confirmation dialog and handles user's choice.
    If user chooses 'No', calls the provided callback function.
    """
    exit_dialog = CustomBox(["Yes", "No"], "Do you want to exit the app?", 40)

    user_choice = exit_dialog.result

    if user_choice == "Yes":
        exit()
    else:
        pass


# =============================================================================
# INITIALIZATION & MAIN FLOW
# =============================================================================
# Loading Screen
app_window = ttk.Window(themename="darkly")
app_window.attributes('-alpha', 0)
app_window.iconbitmap("logo.ico")
center_window(app_window)

# Initialize variables
confirmation_choice = ""

# Ask user for name input method
input_method_choice = ""

# Objects
name_manager = NameManager()
letter = Letter()

while True:
    start_screen = StartingWindow()
    software_choice = start_screen.result
    if software_choice == "📧Mail Merge":
        break
    elif software_choice == "📝Text Editor":
        break
    elif software_choice == "❌Exit":
        exit_confirmation()

if software_choice == "📧Mail Merge":
    name_manager.name_collection()
    letter.process_letter_content()

elif software_choice == "📝Text Editor":
    textEditor_launch()
