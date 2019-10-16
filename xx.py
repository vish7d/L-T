# !/usr/bin/python3
import os
import sys
from tkinter import *

from PIL import Image, ImageTk, ImageOps
from subprocess import call
import threading


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def quit(root):
    root.destroy()


def mainmenu():
    title = resource_path("img/logo1.png")
    b1 = resource_path("img/MBBS.png")
    b2 = resource_path("img/TP.png")
    b3 = resource_path("img/vprs.png")
    b4 = resource_path("img/vpcom.png")
    m1 = resource_path("img/menu.png")
    img2 = resource_path("img/right.png")
    img1 = resource_path("img/left.png")

    image1 = Image.open(title)
    but1 = Image.open(b1)
    but2 = Image.open(b2)
    but3 = Image.open(b3)
    but4 = Image.open(b4)
    men = Image.open(m1)
    imgr = Image.open(img2)
    imgl = Image.open(img1)

    root = Tk()
    root.state('zoomed')
    root.configure(background="#dadfe3")
    root.title("Main Menu")

    imgri = ImageTk.PhotoImage(imgr)
    imgr = imgr.resize((358, 447))
    copy_of_imager = imgr.copy()
    imgri = ImageTk.PhotoImage(imgr)

    imgle = ImageTk.PhotoImage(imgl)
    imgl = imgl.resize((480, 320))
    copy_of_imagel = imgl.copy()
    imgle = ImageTk.PhotoImage(imgl)

    photo1 = ImageTk.PhotoImage(image1)
    image1 = image1.resize((1400, 203))
    copy_of_image1 = image1.copy()
    photo1 = ImageTk.PhotoImage(image1)

    men1 = ImageTk.PhotoImage(men)
    men = men.resize((275, 70))
    copy_of_image0 = men.copy()
    men1 = ImageTk.PhotoImage(men)

    p1 = ImageTk.PhotoImage(but1)
    but1 = but1.resize((275, 70))
    copy_of_image2 = but1.copy()
    p1 = ImageTk.PhotoImage(but1)

    p2 = ImageTk.PhotoImage(but2)
    but2 = but2.resize((275, 70))
    copy_of_image3 = but2.copy()
    p2 = ImageTk.PhotoImage(but2)

    p3 = ImageTk.PhotoImage(but3)
    but3 = but3.resize((275, 70))
    copy_of_image4 = but3.copy()
    p3 = ImageTk.PhotoImage(but3)

    p4 = ImageTk.PhotoImage(but4)
    but4 = but4.resize((275, 70))
    copy_of_image5 = but4.copy()
    p4 = ImageTk.PhotoImage(but4)

    label1 = Label(root, image=photo1, background="#dadfe3")
    label1.grid(row=0, column=0, columnspan=3,pady=(20,40))

    buttonr = Label(root, image=imgri, bg="#dadfe3")
    buttonr.grid(row=1, column=2, rowspan=6, padx=(10, 100),pady=(40,0))
    buttonl = Label(root, image=imgle, bg="#dadfe3")
    buttonl.grid(row=1, column=0, rowspan=6, padx=0)

    button0 = Label(root, image=men1, bg="#dadfe3")
    button0.grid(row=1, column=1)
    button1 = Button(root, image=p1, bg="#dadfe3", border='0', command=mbbs)
    button1.grid(row=2, column=1, pady=2)
    button2 = Button(root, image=p2, bg="#dadfe3", border='0', command=tpbreak)
    button2.grid(row=3, column=1, pady=2)
    button3 = Button(root, image=p3, bg="#dadfe3", border='0', command=vprsbup)
    button3.grid(row=4, column=1, pady=2)
    button4 = Button(root, image=p4, bg="#dadfe3", border='0', command=vpcom)
    button4.grid(row=5, column=1, pady=2)

    root.columnconfigure(0, weight=1)
    root.columnconfigure(1, weight=1)
    root.columnconfigure(2, weight=1)

    root.mainloop()


def mbbs():
    call(["python", "stage3.py"])


def tpbreak():
    call(["python", "p.py"])


def vprsbup():
    pass


def vpcom():
    pass


if __name__ == "__main__":
    # creating thread
    global t1
    t1 = threading.Thread(target=mainmenu)
    t2 = threading.Thread(target=mbbs)
    t3 = threading.Thread(target=tpbreak)
    t4 = threading.Thread(target=vprsbup)
    t5 = threading.Thread(target=vpcom)

    # starting thread 1
    t1.start()

    # wait until thread 1 is completely executed
    t1.join()

    try:
        t2.join()
    except:
        print("AYYYOOO!!")
    try:
        t3.join()
    except:
        pass
    try:
        t4.join()
    except:
        pass
    try:
        t5.join()
    except:
        pass

    print("Done!")
