from tkinter import Tk, filedialog

def select_input_file():
    Tk().withdraw()
    return filedialog.askopenfilename(title="Select Excel File")

def get_save_path(default_name="report.docx"):
    Tk().withdraw()
    return filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word", "*.docx")],
        initialfile=default_name
    )