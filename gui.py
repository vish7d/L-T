# !/usr/bin/python3
import os
import sys
from tkinter import *
from tkinter import scrolledtext
from PIL import Image, ImageTk, ImageOps
from tkinter import ttk


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def run(i):
    progressBar['maximum'] = 100
    x = 100 / length
    progressBar["value"] = i * x
    progressBar.update()


root = Tk()
root.state('zoomed')
root.configure(background="#dadfe3")
root.title("TP Break Up")

# ========================================== Image URL =================================================================

title = resource_path("img/titlevprs.png")
img2 = resource_path("img/side.jpg")
lin = resource_path("img/line.png")
chfi = resource_path("img/choosefile.png")
ee1 = resource_path("img/edit.png")
chek = resource_path("img/check.png")

# =========================================== Opening Image ============================================================

image1 = Image.open(title)
imgr = Image.open(img2)
lini = Image.open(lin)
chck = Image.open(chek)
chofi = Image.open(chfi)
edi = Image.open(ee1)

# =========================================== Resizing =================================================================

photo1 = ImageTk.PhotoImage(image1)
image1 = image1.resize((1400, 179))
copy_of_image1 = image1.copy()
photo1 = ImageTk.PhotoImage(image1)

imgri = ImageTk.PhotoImage(imgr)
imgr = imgr.resize((460, 345))
copy_of_imager = imgr.copy()
imgri = ImageTk.PhotoImage(imgr)

line = ImageTk.PhotoImage(lini)
lini = lini.resize((600, 15))
copy_of_imagel = lini.copy()
line = ImageTk.PhotoImage(lini)

checkk = ImageTk.PhotoImage(chck)
chck = chck.resize((166, 41))
copy_of_imagec = chck.copy()
checkk = ImageTk.PhotoImage(chck)

chosfi = ImageTk.PhotoImage(chofi)
chofi = chofi.resize((139, 37))
copy_of_imagech = chofi.copy()
chosfi = ImageTk.PhotoImage(chofi)

editt = ImageTk.PhotoImage(edi)
edi = edi.resize((40, 40))
copy_of_imagee = edi.copy()
editt = ImageTk.PhotoImage(edi)

# =============================================== Frames ===============================================================

workzone = Frame(root, border='5', bg='grey')
workzone.grid(row=1, column=0, sticky='ENW', padx=(80, 10), pady=(20, 0), rowspan=2)

progress = Frame(root, border='5', bg='#dadfe3')
progress.grid(row=6, column=0, sticky='NS', padx=(10, 80))

group1 = LabelFrame(root, text="Details", padx=5, pady=5, bg='#dadfe3')
group1.grid(row=6, column=1, padx=(0, 80), pady=(0, 45), sticky='NSEW')

# ============================================ Logo and Side Image =====================================================

title = Label(root, bg='#dadfe3', image=photo1)
title.grid(row=0, column=0, columnspan=2, sticky=N, pady=20)

sideimage = Label(root, bg='#dadfe3', image=imgri)
sideimage.grid(row=1, column=1, sticky='ENWS', padx=(0, 100), pady=(0, 0), rowspan=5)

# ============================================ Buttons and Labels ======================================================

lb1t = Label(workzone, text='Select Files', bg='#a0a0a0', font=("Arial Rounded MT Bold", 15), fg="white")
lb1t.grid(row=0, column=0, sticky='ENWS', padx=(0, 5))
lb2t = Label(workzone, text='Selected File', bg='#a0a0a0', font=("Arial Rounded MT Bold", 15), fg="white")
lb2t.grid(row=0, column=1, sticky='ENWS', padx=(0, 5), columnspan=2)
lb21 = Label(workzone, text='Status', bg='#a0a0a0', font=("Arial Rounded MT Bold", 15), fg="white")
lb21.grid(row=0, column=3, sticky='ENWS')

# =========================================== Required by Design =======================================================

lb11 = Label(workzone, text='CS11\n[BOM]', bg='#dadfe3', border='5', font=("Arial Rounded MT Bold", 12),
             foreground="#173d5c")
lb11.grid(row=1, column=0, sticky='ENWS', padx=(0, 5), pady=(5, 0))
label1 = Label(workzone, text='No File Selected', bg='#dadfe3', font=("Arial Rounded MT Narrow", 8), fg="#173d5c")
label1.grid(row=1, column=2, sticky='ENWS', padx=(0, 5), pady=(5, 0))
buttonrd = Button(workzone, image=chosfi, bg='#dadfe3', border='0')
buttonrd.grid(row=1, column=1, padx=(0, 5), pady=(5, 0), sticky='NEWS')
lb1 = Label(workzone, text='', bg='#dadfe3', font=("Arial Rounded MT Bold", 12), foreground="green")
lb1.grid(row=1, column=3, sticky='ENWS', pady=(5, 0))

# ============================================== Posted [CJI3] =========================================================

lb12 = Label(workzone, text='CJI3\n[Posted]', bg='#dadfe3', border='5', font=("Arial Rounded MT Bold", 12),
             foreground="#173d5c")
lb12.grid(row=2, column=0, sticky='ENWS', padx=(0, 5), pady=(5, 0))
label2 = Label(workzone, text='No File Selected', bg='#dadfe3', font=("Arial Rounded MT Narrow", 8), fg="#173d5c")
label2.grid(row=2, column=2, sticky='ENWS', padx=(0, 5), pady=(5, 0))
buttonpo = Button(workzone, image=chosfi, bg='#dadfe3', border='0')
buttonpo.grid(row=2, column=1, padx=(0, 5), pady=(5, 0), sticky='NEWS')
lb2 = Label(workzone, text='', bg='#dadfe3', font=("Arial Rounded MT Bold", 12), foreground="green")
lb2.grid(row=2, column=3, sticky='ENWS', pady=(5, 0))

# =============================================  MBBS [Pending Posting]  ===============================================

lb13 = Label(workzone, text='MBBS \n[Pending Posting]', bg='#dadfe3', border='5', font=("Arial Rounded MT Bold", 12),
             fg="#173d5c")
lb13.grid(row=3, column=0, sticky='ENWS', padx=(0, 5), pady=(5, 0))
label3 = Label(workzone, text='No File Selected', bg='#dadfe3', font=("Arial Rounded MT Narrow", 8), fg="#173d5c")
label3.grid(row=3, column=2, sticky='ENWS', padx=(0, 5), pady=(5, 0))
buttonps = Button(workzone, image=chosfi, bg='#dadfe3', border='0')
buttonps.grid(row=3, column=1, padx=(0, 5), pady=(5, 0), sticky='NEWS')
lb3 = Label(workzone, text='', bg='#dadfe3', font=("Arial Rounded MT Bold", 12), foreground="green")
lb3.grid(row=3, column=3, sticky='ENWS', pady=(5, 0))

# ============================================= OPEN PO [ME2J] =========================================================

lb14 = Label(workzone, text='ME2J\n[OPEN PO]', bg='#dadfe3', border='5', font=("Arial Rounded MT Bold", 12),
             foreground="#173d5c")
lb14.grid(row=4, column=0, sticky='ENWS', padx=(0, 5), pady=(5, 0))
label4 = Label(workzone, text='No File Selected', bg='#dadfe3', font=("Arial Rounded MT Narrow", 8), foreground="#173d5c")
label4.grid(row=4, column=2, sticky='ENWS', padx=(0, 5), pady=(5, 0))
buttonopo = Button(workzone, image=chosfi, bg='#dadfe3', border='0')
buttonopo.grid(row=4, column=1, padx=(0, 5), pady=(5, 0), sticky='NEWS')
lb4 = Label(workzone, text='', bg='#dadfe3', font=("Arial Rounded MT Bold", 12), foreground="green")
lb4.grid(row=4, column=3, sticky='ENWS', pady=(5, 0))

# ======================================================================================================================

buttonbu = Button(root, image=checkk, bg='#dadfe3', border='0')
buttonbu.grid(row=3, column=0, pady=(20, 20), sticky='NS')

lbl = Label(root, image=line, bg='#dadfe3')
lbl.grid(row=4, column=0, sticky='ENWS', pady=20)

lbp = Label(progress, text='Progress: ', bg='#dadfe3')
lbp.grid(row=0, column=0, sticky='ENWS', pady=0)
progressBar = ttk.Progressbar(progress, orient="horizontal", length=286, mode="determinate")
progressBar.grid(row=0, column=1, pady=5)

# Create the textbox
txtbox = scrolledtext.ScrolledText(group1, width=30, height=30)
txtbox.grid(row=0, column=0, padx=0, pady=(0, 0), sticky='NSEW')

# ============================================= Column Weights =========================================================

root.columnconfigure(0, weight=15)
root.columnconfigure(1, weight=2)
workzone.columnconfigure(0, weight=2)
workzone.columnconfigure(1, weight=1)
workzone.columnconfigure(2, weight=3)
workzone.columnconfigure(3, weight=1)
group1.rowconfigure(0, weight=1)
group1.columnconfigure(0, weight=1)
root.rowconfigure(6, weight=1)

# ======================================================================================================================

root.mainloop()
