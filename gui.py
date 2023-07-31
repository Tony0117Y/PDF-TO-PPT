import tkinter as tk
from tkinter import filedialog, ttk
from convert import convert_pdf_to_pptx
from PIL import Image, ImageTk

pdf_file = ''
def select_pdf_file():
    global pdf_file
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    return pdf_file

def convert_button_clicked():
    global pdf_file
    pdf_file = select_pdf_file()
    output_pptx = filedialog.asksaveasfilename(defaultextension=".pptx")
    pptx_title = title_entry.get()
    pptx_sub_title = sub_title_entry.get()
    slide_theme = select_theme()
    convert_pdf_to_pptx(pdf_file, output_pptx, pptx_title, pptx_sub_title, slide_theme)
    root.destroy()

def select_theme():
    theme = theme_var.get()
    bg_theme = {
        "Bubbles": "images/bubbles.jpg",
        "Nature": "images/Nature.jpg",
        "School": "images/school.jpg",
        "Triangles": "images/triangles.jpg",
    }[theme]
    return bg_theme

root = tk.Tk()
root.title("PDF-to-PPT")

window_width, window_height = 414, 736
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
center_x = (screen_width - window_width) // 2
center_y = (screen_height - window_height) // 2
root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

background_image = Image.open("./images/gg.jpg")
background_image = background_image.resize((window_width, window_height), Image.BILINEAR)
background_image = ImageTk.PhotoImage(background_image)

canvas = tk.Canvas(root, width=window_width, height=window_height)
canvas.pack(fill="both", expand=True)

canvas.create_image(0, 0, image=background_image, anchor="nw")

theme_var = tk.StringVar(root)
theme_var.set("Bubbles")

theme_box = canvas.create_text(window_width / 2, 20, anchor="center", text="Select The Theme For Powerpoint:", fill="white")
dropdown_theme = ttk.Combobox(root, textvariable=theme_var, values=["Bubbles", "Nature", "School", "Triangles"])
canvas.create_window(window_width / 2, 60, window=dropdown_theme)

title_label = canvas.create_text(window_width / 2, 120, anchor="center", text="Enter the title of the PPTX:", fill="white")
title_entry = tk.Entry(root)
canvas.create_window(window_width / 2, 160, window=title_entry)

sub_title_label = canvas.create_text(window_width / 2, 220, anchor="center", text="Enter the subtitle for the PPTX slides:", fill="white")
sub_title_entry = tk.Entry(root)
canvas.create_window(window_width / 2, 260, window=sub_title_entry)

convert_button = tk.Button(root, text="Select PDF file and Convert", command=convert_button_clicked)
canvas.create_window(window_width / 2, 320, window=convert_button)


root.mainloop()
