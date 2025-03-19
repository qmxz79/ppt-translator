def select_file():
    """Open a file dialog to select a PowerPoint file."""
    from tkinter import filedialog
    from tkinter import Tk

    root = Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title="选择要翻译的课件", filetypes=[("PowerPoint Files", "*.pptx;*.ppt")])
    return file_path

def save_file():
    """Open a file dialog to select the location and name for saving the translated file."""
    from tkinter import filedialog
    from tkinter import Tk

    root = Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.asksaveasfilename(title="另存为", defaultextension=".pptx", filetypes=[("PowerPoint Files", "*.pptx")])
    return file_path

def is_valid_file(file_path):
    """Check if the selected file is a valid PowerPoint file."""
    return file_path.endswith(('.pptx', '.ppt')) and os.path.isfile(file_path)

def create_directory_if_not_exists(directory):
    """Create a directory if it does not exist."""
    import os
    if not os.path.exists(directory):
        os.makedirs(directory)