from hashlib import md5
from pathlib import Path
from tkinter import Tk, filedialog, Label, Button, Entry, Toplevel
from tkinter.ttk import Progressbar, Style
from threading import Thread
from openpyxl import Workbook
from datetime import datetime
from os.path import join, dirname, isfile, exists
from warnings import filterwarnings
from time import time

fileHandler = open(f"logs_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt", 'a')

iconFile = join(dirname(__file__), 'md5.ico')
aboutIcon = join(dirname(__file__), 'info.ico')


sourcePathFolder = ''

def threadingFileDialog(progressBar, progressStyle, pathEntry, fileDialogBtn, submitBtn, messageText):
    threadFileDialog = Thread(target=FileDialog, args=(progressBar, progressStyle, pathEntry, fileDialogBtn, submitBtn,
                                                       messageText))
    threadFileDialog.start()


def FileDialog(progressBar, progressStyle, pathEntry, fileDialogBtn, submitBtn, messageText):
    global sourcePathFolder
    progressBar.config(value=0)
    progressStyle.configure("Custom.Horizontal.TProgressbar", text='0 %')
    sourcePathFolder = filedialog.askdirectory()
    pathEntry.config(state='normal')
    pathEntry.delete(0, 'end')
    pathEntry.insert(0, sourcePathFolder)
    if len(pathEntry.get()) > 0:
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} [{sourcePathFolder}] is selected as source path\n')
        submitBtn.config(state='normal', bg='green')
        fileDialogBtn.config(state='disabled', bg='light grey')
    else:
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} No source path selected')
        messageText.config(text='Please select source path')
    pathEntry.config(state='disabled')


def threadingMD5Checksum(submitBtn, fileDialogBtn, messageText, progressBar, window, progressStyle):
    threadMD5Checksum = Thread(target=MD5Checksum, args=(submitBtn, fileDialogBtn, messageText, progressBar, window,
                                                         progressStyle))
    threadMD5Checksum.start()


def MD5Checksum(submitBtn, fileDialogBtn, messageText, progressBar, window, progressStyle):
    startTime = time()
    submitBtn.config(state='disabled', bg='light grey')
    messageText.config(text='Checking...')
    filterwarnings("ignore", category=DeprecationWarning)
    excelFile = Workbook()
    workSheet = excelFile.active
    workSheet.title = 'Checksum'
    fileHandler.write(f'{datetime.now().replace(microsecond=0)}[Checksum] worksheet created\n')
    workSheet.append(['Directory', 'Checksum'])
    for column in workSheet.columns:
        for cell in column:
            alignmentObj = cell.alignment.copy(horizontal='center', vertical='center')
            cell.alignment = alignmentObj
    fileHandler.write(f'{datetime.now().replace(microsecond=0)} Checking...\n')
    directories = [fileDir for fileDir in Path(sourcePathFolder).rglob("*") if isfile(fileDir)]
    grandTotal = len(directories)
    checkSuccess = 0
    if grandTotal > 0:
        for path in directories:
            excelPathString = str(path)
            with open(path, 'rb') as getMd5:
                data = getMd5.read()
                getHash = md5(data).hexdigest()
                checkSuccess = checkSuccess + 1
                workSheet.append([excelPathString, getHash])
                updateProgressSaveExcel(checkSuccess, progressBar, grandTotal, window, progressStyle)
        excelFileName = "MD5_Checksum.xlsx"
        excelFIleCounter = 1
        while exists(excelFileName):
            excelFileName = f"MD5_Checksum_{excelFIleCounter}.xlsx"
            excelFIleCounter = excelFIleCounter + 1
        excelFile.save(excelFileName)
        fileHandler.write(f'{datetime.now().replace(microsecond=0)}Listing Done for\nGrand_Total: [{checkSuccess}]\n')
        fileHandler.write(f"{datetime.now().replace(microsecond=0)}[{excelFileName}] saved in current directory.\n")
        messageText.config(text=f'Done, {excelFileName} created')
    else:
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} [{sourcePathFolder}] is empty\n')
        messageText.config(text='Error! check logs')
    endTime = time()
    fileHandler.write(f'{datetime.now().replace(microsecond=0)} ELAPSED TIME TO COMPLETE THE PROCESS IS '
                      f'{((endTime - startTime) / 60)} Minutes\n')
    fileDialogBtn.config(state='normal', bg='green')

def updateProgressSaveExcel(listingSuccess, progressBar, totalFiles, window, progressStyle):
    resultVal = (listingSuccess / totalFiles) * 100
    progressBar['value'] = resultVal
    progressStyle.configure("Custom.Horizontal.TProgressbar", text='{:g} %'.format(resultVal))
    window.update()


def mainGUI():
    window = Tk()
    window.config(bg='light grey')
    window.title('File Integrity Checker v1.0')
    window.geometry('349x180')
    window.resizable(False, False)
    window.iconbitmap(iconFile)
    mainLabel = Label(window, text='MD5 Checksum', font=('Arial', 15, 'bold'), fg='blue', bg='light grey')
    mainLabel.place(x=180, y=14, anchor='center')
    sourcePathLabel = Label(window, text='Path:', font=('Arial', 8, 'bold italic'), bg='light grey').place(x=2, y=40)
    pathEntry = Entry(window, bd=4, width=46, bg='white', state='disabled')
    pathEntry.place(x=40, y=40)
    fileDialogBtn = Button(window, text='...', bg='green', fg='white', font=('Arial', 8),
                           command=lambda: threadingFileDialog(progress, progressStyle, pathEntry, fileDialogBtn,
                                                               submitBtn, messageLabel))
    fileDialogBtn.place(x=328, y=40)
    submitBtn = Button(window, text='Submit', bg='light grey', fg='white', font=('Arial', 12, 'bold'),
                       command=lambda: threadingMD5Checksum(submitBtn, fileDialogBtn, messageLabel, progress, window,
                                                            progressStyle), state='disabled')
    submitBtn.place(x=150, y=75)
    progress = Progressbar(window, length=340, mode="determinate", style="Custom.Horizontal.TProgressbar")
    progress.place(x=5, y=120)
    progressStyle = Style()
    progressStyle.theme_use('default')
    progressStyle.configure("Custom.Horizontal.TProgressbar", thickness=20, troughcolor='#E0E0E0', background='#4CAF50',
                            troughrelief='flat', relief='flat', text='0 %')
    progressStyle.layout('Custom.Horizontal.TProgressbar', [('Horizontal.Progressbar.trough',
                                                             {'children': [('Horizontal.Progressbar.pbar',
                                                                            {'side': 'left', 'sticky': 'ns'})],
                                                              'sticky': 'nswe'}),
                                                            ('Horizontal.Progressbar.label', {'sticky': ''})])

    messageLabel = Label(window, font=('Arial', 11, 'bold'), bg='light grey')
    messageLabel.place(x=5, y=150)
    aboutBtn = Button(window, text='?', bg='brown', command=lambda: aboutWindow(window))
    aboutBtn.place(x=330, y=150)
    window.mainloop()

def aboutWindow(mainWin):
    aboutWin = Toplevel(mainWin)
    aboutWin.grab_set()
    aboutWin.geometry('285x90')
    aboutWin.resizable(False, False)
    aboutWin.title('About')
    aboutWin.iconbitmap(aboutIcon)
    aboutWinLabel = Label(aboutWin, text=f'Version - 1.0\nDeveloped by Priyanshu\nFor any improvement please reach on '
                                         f'below email\nEmail : chandelpriyanshu8@outlook.com\nMobile : '
                                         f'+91-8285775109 '
                                         f'', font=('Helvetica', 9)).place(x=1, y=6)


mainGUI()
