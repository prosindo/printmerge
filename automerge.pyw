"""
auto merge printing v2
python 2.7.14 tested
require module pypiwin32: pip install pypiwin32
require module pil: pip install pillow
"""

import os
import subprocess
from datetime import datetime
from Tkinter import Tk, Frame, Button, Label, Checkbutton, IntVar, StringVar, \
    TOP, LEFT, RIGHT, BOTTOM
from ttk import Combobox
import win32print
from printlib import PrintLib
from PIL import Image

DEFAULT_PRINTER = "EPSON M100 Series MD"
SRC_FOLDER = "PNG"
MERGED_FOLDER  = "PNGMERGED"
ARCHIVE_FOLDER = "PNGARCHIVE"

FULLPAPERIFONE = False
PRINTPERPAGE = 4  # 2 or 4
if PRINTPERPAGE == 2:
    IMAGE_ORDER = 0   # 0 = RIGHT THEN BELOW, 1 = BELOW THEN RIGHT
    IMAGE_ROTATE = 90
else:
    IMAGE_ORDER = 1
    IMAGE_ROTATE = 0
SOURCE_WIDTH = 1275
SOURCE_HEIGHT = 1650

WINDOW_WIDTH = 250
WINDOW_HEIGHT = 115


class App(object):
    """main app"""
    # pylint: disable=too-many-instance-attributes
    # pylint: disable=too-many-locals
    # pylint: disable=no-member

    @staticmethod
    def getimages(path, maxfind):
        """get images"""
        imglist = []

        for fpath in os.listdir(path):
            ext = os.path.splitext(fpath)[1]
            if ext == ".png":
                imglist.append(fpath)
                # print fname
            if len(imglist) == maxfind:
                break
        return imglist

    @staticmethod
    def movefiles(srcpath, dstpath, filenames):
        """move files"""
        for fname in filenames:
            os.rename(srcpath + "/" + fname, dstpath + "/" + fname)

    def mergeimages(self, srcpath, dstpath, acvpath):
        """merge images"""
        imglist = self.getimages(srcpath, PRINTPERPAGE)
        if not imglist:  # empty list
            self.lbst.configure(text="No document found")
            return

        # source images size
        basew = SOURCE_WIDTH
        baseh = SOURCE_HEIGHT

        # crop this size (about 73.5% width)
        imgwcrop = int(basew / 2 / 0.68)
        imghcrop = int(baseh / 4 / 0.68)

        # paste this size
        if IMAGE_ROTATE == 90 or IMAGE_ROTATE == 270:
            imgw = int(basew / 2)
            imgh = int(baseh / 2)
        else:
            imgw = int(basew / 2)
            imgh = int(baseh / 4)

        # full paper if only 1 image
        if FULLPAPERIFONE and len(imglist) == 1:
            imgw = int(basew)
            imgh = int(baseh / 2)

        # new dimension of paper
        baseh = int(baseh / 2)  # half source A4 = A5
        print "Creating image", basew, baseh
        baseim = Image.new("RGB", (basew, baseh), (255, 255, 255))

        left = 0
        top = 0
        for fname in imglist:
            img = Image.open(srcpath + "/" + fname, "r")
            imgcrop = img.crop((0, 0, imgwcrop, imghcrop))
            print imgcrop.width, imgcrop.height
            if IMAGE_ROTATE != 0:
                imgcrop = imgcrop.rotate(IMAGE_ROTATE, expand=True)
            print imgcrop.width, imgcrop.height
            imgcrop = imgcrop.resize((imgw, imgh), Image.LANCZOS)
            baseim.paste(imgcrop, (left, top, left + imgw, top + imgh))

            if IMAGE_ORDER == 0:  # right then below
                if left == 0:
                    left += imgw
                else:
                    left = 0
                    top += imgh
            else:  # below then right
                if top == 0:
                    top += imgh
                else:
                    top = 0
                    left += imgw

        dstfn = dstpath + r"\merge-" + \
            datetime.now().strftime("%Y%m%d%H%M%S") + ".png"
        baseim = baseim.convert('P', colors=256, palette=Image.ADAPTIVE)
        baseim.save(dstfn, optimize=True)
        self.movefiles(srcpath, acvpath, imglist)
        self.lbst.configure(text=dstfn)
        print "New file:", dstfn
        return dstfn

    @staticmethod
    def printdocument(filename, printername):
        """print document"""
        prnlib = PrintLib()
        prnlib.printfile(filename, printername)

    def buttonprinthandler(self):
        """button print handler"""
        self.lbst.configure(text="Merging images")
        dst = self.mergeimages(self.srcpath, self.dstpath, self.acvpath)
        self.lbst.configure(text="Printing document")
        self.printdocument(dst, self.printername)

    def buttonclearhandler(self):
        """button clear handler"""
        imglist = self.getimages(self.srcpath, 100)
        self.movefiles(self.srcpath, self.acvpath, imglist)
        return

    def documentcounthandler(self):
        """document count handler"""
        docfound = len(self.getimages(self.srcpath, 100))
        self.lbdoc.configure(text="Document found: " + str(docfound))
        if self.chautoprintval.get() == 1 and docfound >= PRINTPERPAGE:
            self.lbst.configure(text="Auto print now")
            self.buttonprinthandler()
        else:
            self.lbst.configure(text="Ready: " + self.printername)
        if self.chprintallval.get() == 1 and docfound > 0:
            self.lbst.configure(text="Print all now")
            self.buttonprinthandler()
        else:
            self.chprintallval.set(0)
            self.lbst.configure(text="Ready: " + self.printername)
        self.main.after(2000, self.documentcounthandler)

    def buttonexplorehandler(self):
        """button explore"""
        print 'explorer "' + os.getcwd() + '\\' + self.srcpath + '"'
        subprocess.Popen('explorer "' + os.getcwd() +
                         '\\' + self.srcpath + '"')
        return

    @staticmethod
    def populateprinters(defprinter):
        """populate printers"""
        defidx = 0
        printer_info = win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )
        printer_names = []
        for pidx in range(0, len(printer_info)):
            printer_names.append(printer_info[pidx][2])
            if printer_info[pidx][2] == defprinter:
                defidx = pidx

        return (defidx, printer_names)

    def printerselecthandler(self, event):
        """printer select"""
        print event
        self.printername = self.cbprinterval.get()
        return

    def __init__(self):
        self.srcpath = SRC_FOLDER
        self.dstpath = MERGED_FOLDER
        self.acvpath = ARCHIVE_FOLDER
        self.printername = win32print.GetDefaultPrinter()
        if DEFAULT_PRINTER != "":
            self.printername = DEFAULT_PRINTER
        self.main = Tk()
        screenw = self.main.winfo_screenwidth()
        screenh = self.main.winfo_screenheight()
        self.main.title("PrintMerge")
        self.main.wm_attributes("-topmost", 1)
        self.main.geometry("%dx%d%+d%+d" %
                           (WINDOW_WIDTH, WINDOW_HEIGHT,
                            screenw - WINDOW_WIDTH, screenh - 70 - WINDOW_HEIGHT))
        self.lbdoc = Label(self.main, text="Reading documents..")
        self.lbdoc.pack(side=TOP)
        fmain = Frame(self.main)  # first frame, button frame
        self.btprint = Button(fmain, text="Print",
                              command=self.buttonprinthandler)
        self.btprint.pack(side=LEFT)
        self.btclear = Button(fmain, text="Clear",
                              command=self.buttonclearhandler)
        self.btclear.pack(side=RIGHT)
        self.btexplore = Button(
            fmain, text="Explore", command=self.buttonexplorehandler)
        self.btexplore.pack(side=RIGHT)
        fmain.pack(side=TOP)
        self.lbst = Label(self.main, text="Ready: " + self.printername)
        self.lbst.pack(side=BOTTOM)
        fsec = Frame(self.main)  # second frame, checkbox frame
        self.chautoprintval = IntVar()
        self.chautoprint = Checkbutton(
            fsec, text="Auto Print", variable=self.chautoprintval)
        self.chautoprint.pack(side=LEFT)
        self.chprintallval = IntVar()
        self.chprintall = Checkbutton(
            fsec, text="Print ALL", variable=self.chprintallval)
        self.chprintall.pack(side=RIGHT)
        self.cbprinterval = StringVar()
        self.cbprinter = Combobox(
            self.main, textvariable=self.cbprinterval, state='readonly')
        (self.defidx, self.printerlist) = self.populateprinters(self.printername)
        self.cbprinter['values'] = self.printerlist
        self.cbprinter.current(self.defidx)
        self.cbprinter.bind("<<ComboboxSelected>>", self.printerselecthandler)
        self.cbprinter.pack(side=BOTTOM)
        fsec.pack(side=BOTTOM)

        self.main.after(2000, self.documentcounthandler)
        self.main.mainloop()


App()
