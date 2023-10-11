import customtkinter as ctk # pip install customtkinter | https://customtkinter.tomschimansky.com/
from customtkinter import filedialog
import os
from docx import Document # pip install python-docx | https://python-docx.readthedocs.io/en/latest/user/install.html
from docx.shared import Pt, Cm, Mm
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
import natsort # pip install natsort | https://pypi.org/project/natsort/#installation
import threading

# tkinter mainloop is blocking. handle frozen gui with threading
def threaded_func(func):
    def inner(*args, **kwargs):
        thread = threading.Thread(target=func, args=args, kwargs=kwargs)
        thread.start()
    return inner

@threaded_func
def create_word_document(folder_path, pic_width, pic_height, start_count, default_caption, pic_extensions, save_file):
    print("Creating Word Document..")
    
    # Get picture dimensions
    pic_width = float(pic_width)
    pic_height = float(pic_height)

    # Create a new Word document
    doc = Document()
    
    # Set landscape orientation and two columns layout
    section = doc.sections[0]
    section.page_width = Mm(297)
    section.page_height = Mm(210)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.left_margin = Mm(17.5)
    section.right_margin = Mm(17.5)
    section.top_margin = Mm(12.5)
    section.bottom_margin = Mm(10.0)
    section.header_distance = Mm(12.7)
    section.footer_distance = Mm(12.7)
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')[0]
    cols.set(qn('w:num'),'2')
    
    # Set document style
    style = doc.styles['Normal']
    paragraph_format = style.paragraph_format
    paragraph_format.space_after = Pt(8)
    paragraph_format.line_spacing = 1.1

    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)

    # Get a list of files with the specified extension in the folder
    pic_files = [f for f in os.listdir(folder_path) if any(f.upper().endswith(e) for e in pic_extensions)]
    pic_files = natsort.natsorted(pic_files, alg=natsort.PATH)
    # Initialize photo counter
    count = int(start_count)
    
    # Iterate through the picture files and insert them into the document
    for pic_file in pic_files:     
        doc.add_paragraph()
        doc.paragraphs[-1].add_run().add_picture(os.path.join(folder_path, pic_file), width=Cm(pic_width), height=Cm(pic_height))
        
        # Add a caption below the image
        doc.add_paragraph()
        caption = f"{default_caption} {count}"
        doc.paragraphs[-1].add_run(caption)
        status_label.configure(text=f"Inserting Photo {count}..", fg_color="yellow")
        print(f"Inserting Photo {count}..")
        count += 1

    # Save the Word document
    status_label.configure(text=f"Saving document..")
    print("Saving document..")
    doc.save(save_file)
    print("Word document created successfully.")
    status_label.configure(text="FINISHED SUCCESSFULLY", fg_color="lightgreen")

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_path_entry.delete(0, ctk.END)
    folder_path_entry.insert(0, folder_path)
    folder_check()

def browse_save():
    desktop = os.path.normpath(os.path.expanduser("~/Desktop")).replace('\\','/')
    save_file =  filedialog.asksaveasfilename(initialdir=desktop, initialfile="Photos.docx", defaultextension=".docx", title="Select Save File Location",filetypes=(("Word Document","*.docx"),))
    return save_file

def getExtensions():
    extensions = set()
    if jpg_extension:
        extensions.add('.JPG')
    if jpeg_extension:
        extensions.add('.JPEG')
    if png_extension:
        extensions.add('.PNG')
    return extensions

def folder_check():
    if folder_path_entry.get():
        if not os.path.exists(folder_path_entry.get()):
            folder_path_entry.configure(border_color="pink")
            status_label.configure(text="Folder specified does not exist. Please reselect.", fg_color="pink")
            return False
        status_label.configure(text="Configure options and click 'Submit' to begin collating.", fg_color="yellow")
        submit_button.configure(state="active")
        folder_path_entry.configure(fg_color=default_colour, border_color=default_border_colour)
        if not number_entry_checks():
            status_label.configure(text="Invalid number.", fg_color="pink")
        return True
    status_label.configure(text="Set folder location of photos.", fg_color="yellow")
    submit_button.configure(state="disabled")
    return False

def number_entry_check(number_widget):
    if not isfloat(number_widget.get()):
        number_widget.configure(border_color="pink")
        return False
    else:
        number_widget.configure(fg_color=default_colour, border_color=default_border_colour)
        return True

def number_entry_checks():
    valid_entry = True
    number_entry_widgets = [pic_width_entry, pic_height_entry, start_count_entry]
    for w in number_entry_widgets:
        if not number_entry_check(w):
            valid_entry = False
    return valid_entry

def validate_entries():
    valid_entries = number_entry_checks()
    if not caption_entry.get():
        valid_entries = False
    if not folder_check():
        valid_entries = False
    if not valid_entries:
        submit_button.configure(state="disabled")
    else:
        submit_button.configure(state="normal", hover=True, hover_color=default_hover_colour)
    return valid_entries

def isfloat(str):
    try:
        if not str:
            raise ValueError
        float(str)
        return True
    except ValueError:
        return False

@threaded_func
def update_label(text:str, *args, **kwargs):
    status_label.configure(text=text, *args, **kwargs)

def submit():
    update_label(text="Processing..", fg_color="yellow")
    if not validate_entries():
        return
    pic_extension = getExtensions()
    save_file = browse_save()
    if not save_file:
        status_label.configure(text="Configure options and click 'Submit' to begin collating.", fg_color="yellow")
        return
    
    try:
        create_word_document(folder_path_entry.get(), pic_width_entry.get(), pic_height_entry.get(), start_count_entry.get(), caption_entry.get(), pic_extension, save_file)
    except (ValueError, AttributeError) as e:
        print(type(e), e)
        status_label.configure(text="AN ERROR OCCURRED. PLEASE TRY AGAIN.", fg_color="red")
    except FileNotFoundError:
        status_label.configure(text="Folder location is not valid. Please reselect.", fg_color="red")

# Create the main application window
app = ctk.CTk()
app.title("Photo Collator")
app.resizable(False, False)

# Create and place widgets
folder_path_label = ctk.CTkLabel(app, text="Folder Location of Photos:").grid(row=0, column=0, padx=10, pady=5, sticky='E')

folder_path_entry = ctk.CTkEntry(app)
folder_path_entry.grid(row=0, column=1, padx=10, pady=5)
browse_button = ctk.CTkButton(app, text="Browse", command=browse_folder).grid(row=0, column=2, padx=10, pady=5, sticky='EW')
folder_path_entry.bind("<FocusOut>", lambda event: validate_entries())

pic_width_label = ctk.CTkLabel(app, text="Picture Width (cm):").grid(row=1, column=0, padx=10, pady=5, sticky='E')
pic_width_entry = ctk.CTkEntry(app)
pic_width_entry.grid(row=1, column=1, padx=10, pady=5, sticky='EW')
pic_width_entry.insert(0, "12.52")
pic_width_entry.bind("<FocusOut>", lambda event: validate_entries())

pic_height_label = ctk.CTkLabel(app, text="Picture Height (cm):").grid(row=2, column=0, padx=10, pady=5, sticky='E')
pic_height_entry = ctk.CTkEntry(app)
pic_height_entry.grid(row=2, column=1, padx=10, pady=5, sticky='EW')
pic_height_entry.insert(0, "8.3")
pic_height_entry.bind("<FocusOut>", lambda event: validate_entries())

start_count_label = ctk.CTkLabel(app, text="Start Count:").grid(row=3, column=0, padx=10, pady=5, sticky='E')
start_count_entry = ctk.CTkEntry(app)
start_count_entry.grid(row=3, column=1, padx=10, pady=5, sticky='EW')
start_count_entry.insert(0, "1")
start_count_entry.bind("<FocusOut>", lambda event: validate_entries())

caption_label = ctk.CTkLabel(app, text="Default Caption:").grid(row=4, column=0, padx=10, pady=5, sticky='E')
caption_entry = ctk.CTkEntry(app)
caption_entry.grid(row=4, column=1, padx=10, pady=5, sticky='EW')
caption_entry.insert(0, "Photo")

pic_extension_label = ctk.CTkLabel(app, text="Picture Extension:").grid(row=5, column=0, padx=10, pady=5, sticky='E')
jpg_extension = ctk.BooleanVar()
jpeg_extension = ctk.BooleanVar()
png_extension = ctk.BooleanVar()
jpg_checkbox = ctk.CTkCheckBox(app, text=".JPG", variable=jpg_extension, onvalue=True, offvalue=False)
jpg_checkbox.grid(row=5, column=1, columnspan=1, padx=10, pady=5, sticky='W')
jpeg_checkbox = ctk.CTkCheckBox(app, text=".JPEG", variable=jpeg_extension, onvalue=True, offvalue=False)
jpeg_checkbox.grid(row=5, column=1, columnspan=2, padx=80, pady=5, sticky='W')
png_checkbox = ctk.CTkCheckBox(app, text=".PNG", variable=png_extension, onvalue=True, offvalue=False)
png_checkbox.grid(row=5, column=2, columnspan=2, padx=0, pady=5, sticky='W')
jpg_checkbox.select()
png_checkbox.select()

submit_button = ctk.CTkButton(app, text="Submit", command=submit, state="disabled")
submit_button.grid(row=7, column=0, columnspan=3, padx=10, pady=10)

status_label = ctk.CTkLabel(app, text="Set folder location of photos.", corner_radius=5, padx=20, fg_color="yellow", text_color="black")
status_label.grid(row=8, column=0, columnspan=3, padx=10, pady=5)
default_colour = folder_path_entry.cget("fg_color")
default_border_colour = folder_path_entry.cget("border_color")
default_hover_colour = submit_button.cget("hover_color")

# Start the Tkinter main loop
app.mainloop()