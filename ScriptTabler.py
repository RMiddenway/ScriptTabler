import pdftotext
from docx import Document
import re
from docx.shared import Cm, Inches
import os

from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import tkinter.scrolledtext as tkst

target_directory = ''

max_line_width = 60
includes_title_page = True

def set_col_widths(table):
    widths = (Inches(1), Inches(1.5), Inches(4))
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

def check_for_dialogue_name(line):#checks for name only
    first_char_index = 0
    white_space = check_white_space(line)
    line_length = len(line)

    if white_space[0] < 15:# dialogue names usually indented 25 spaces
        return False
    if any(c for c in line if c.islower()):#dialogue names usually in caps
        return False
    if line_length > 50:# if line finishes on the right, name not centered
        return False
    return True

#checks for end of dialogue by looking for line left justified
def check_for_dialogue_end(line):
    white_space = check_white_space(line)
    if white_space[0] < 8: #dialogue usually indented 12 spaces
        return True
    if white_space[0] > 30:
        return True
    elif check_for_dialogue_name(line):
        return True
    else:
        return False

def check_white_space(line):
    white_space = []
    length = 0
    if line[0] is not ' ':
        white_space.append(length)
    for i in range(len(line)):
        if i < (len(line)):
            if line[i] is ' ':
                length += 1
            elif length is not 0:
                white_space.append (length)
                length = 0
        else:
            white_space.append(length)
    if line[0] != ' ':
        white_space.append(length)
    return white_space

root = Tk()
root.title("Script Tabler")

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

#debug_window = ttk.Label(mainframe, text="Debug output")
#debug_window.grid(column=2, row=3, sticky=S)

def convert_to_table(pdf, output_name, output_directory):
    page_array = [p for p in pdf]
    # removes title page
    if includes_title_page:
        page_array.pop(0)

    application_path = ''
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, the pyInstaller bootloader
        # extends the sys module by a flag frozen=True and sets the app
        # path into variable _MEIPASS'.
        application_path = sys._MEIPASS
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    print(application_path)

    document = Document(application_path + '/default.docx')
    table = document.add_table(rows=1, cols=3)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Timecode'
    hdr_cells[1].text = 'Char'
    hdr_cells[2].text = 'Line'

    document.save('%s.docx' % (output_directory + '/' + os.path.splitext(output_name)[0]))
    lines = {}
    line = ''
    is_dialogue = False
    row_cells = table.rows[0].cells
    current_dialogue = ''

    for page in page_array:
        document.save('%s.docx' % (output_directory + '/' + os.path.splitext(output_name)[0]))
        line_array = [page for page in page.split('\n') if page != '']
        for line in line_array:
            if is_dialogue:
                if check_for_dialogue_name(line):
                    row_cells = table.add_row().cells
                    row_cells[1].text = re.sub(r'\(.*\)', '', line).lstrip()

                elif check_for_dialogue_end(line):
                    is_dialogue = False

                else:
                    row_cells[2].text += line.lstrip() + ' '

            else:
                if check_for_dialogue_name(line):
                    row_cells = table.add_row().cells
                    row_cells[1].text = re.sub(r'\(.*\)', '', line).lstrip()
                    is_dialogue = True

    set_col_widths(table)
    document.save('%s.docx' % (output_directory + '/' + os.path.splitext(output_name)[0]))


def convert_pdfs():
    global dest_dir_name
    dest_dir_name = original_dir_name = 'Tabled Scripts'
    appended_number = 1
    while dest_dir_name in os.listdir(target_directory):
        appended_number += 1
        dest_dir_name = original_dir_name + str(appended_number)
    os.makedirs(target_directory + '/' + dest_dir_name)

    files = [target_directory + '/' + file for file in os.listdir(target_directory)]
    for file in files:
        if file.endswith('.pdf'):
            pdfFile = open(file, 'rb')

            pdf = pdftotext.PDF(pdfFile)
            convert_to_table(pdf, os.path.basename(file), target_directory + '/' + dest_dir_name)
    b_convert.config(state="disabled")



file_display = tkst.ScrolledText(mainframe, width=40, height=10)
file_display.grid(column=2, row=1, sticky=N)

def show_contents(files):
    file_display.delete('1.0', END)
    file_display.insert('1.0',files)

def select_folder():
    dirname = filedialog.askdirectory()
    global target_directory
    target_directory = dirname
    file_list = os.listdir(dirname)
    show_contents('\n'.join([file for file in file_list if file.endswith('.pdf')]))
    b_convert.config(state="normal")

def convert_contents():
    convert_pdfs()
    show_contents(['DONE'])

b_select_folder = ttk.Button(mainframe, text="Select Folder", command=select_folder)
b_select_folder.grid(column=1, row=3, sticky=S)

b_convert = Button(mainframe, text="Convert", command=convert_contents, disabledforeground='grey') #removed ttk to allow greying
b_convert.grid(column=3, row=3, sticky=S)
b_convert.config(state="disabled")
button_includes_titlepage = Checkbutton(mainframe, text="Includes title page", variable=includes_title_page, onvalue = 1, offvalue = 0)
button_includes_titlepage.grid(column=1, row=2, sticky=S)
button_includes_titlepage.select()

root.mainloop()
