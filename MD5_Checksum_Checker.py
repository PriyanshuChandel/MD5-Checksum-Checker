from hashlib import md5
from pathlib import Path
from tkinter import Tk, filedialog, Label, Button, Entry
from threading import Thread
from time import sleep
from openpyxl import Workbook

window = Tk()
window.title('MD5_Checksum_Checker')
window.minsize(width=349, height=160)
window.maxsize(width=349, height=160)


def threading_btn1():
    thread_btn1 = Thread(target=func_btn1)
    thread_btn1.start()


def func_btn1():
    global user_path_selection
    user_path_selection = filedialog.askdirectory()
    ent1.insert(0, user_path_selection)


def threading_btn2():
    thread_btn2 = Thread(target=mainfunc)
    thread_btn2.start()


def secondary_function():
    path_ = Path(user_path_selection)
    excel_file = Workbook()
    work_sheet = excel_file.active
    work_sheet.title = 'Checksum'
    work_sheet.append(['Directory', 'File-Name', 'Checksum'])
    print('start')
    for full_path_of_file in path_.rglob("*"):
        file_name = str(full_path_of_file).split('\\')[-1]
        if '.' in file_name:
            with open(full_path_of_file, 'rb') as getmd5:
                data = getmd5.read()
                gethash = md5(data).hexdigest()
                work_sheet.append([str(full_path_of_file), file_name, gethash])
                excel_file.save(f"MD5_Checksum.xlsx")
    labl4.config(text='Checksum checked successfully, MD5_Checksum.xlsx created in current working directory',
                 wraplength=347, justify="left")


def mainfunc():
    labl4.config(text='Checking MD5 checksum...')
    sleep(2)
    secondary_function()


labl1 = Label(window, text='Welcome to MD5 checksum checker', font=(None, 10, 'bold')).place(x=60, y=8)
labl2 = Label(window, text="Path").place(x=0, y=38)

ent1 = Entry(window, bd=4, width=44, bg='lavender')
ent1.place(x=40, y=36)

btn1 = Button(window, text='...', bg='green', command=threading_btn1)
btn1.place(x=320, y=36)

btn2 = Button(window, text='Submit', bg='green', command=threading_btn2)
btn2.place(x=145, y=65)

labl3 = Label(window, text='After submission wait for success message').place(x=0, y=95)

labl4 = Label(window)
labl4.place(x=0, y=120)

window.mainloop()
