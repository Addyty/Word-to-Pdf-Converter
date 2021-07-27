#import Library in module
from docx2pdf import convert
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import os
from random import randint

#Define class for creating os_window.
class Root(Tk):
    def __init__(self):
        super(Root,self).__init__()
        self.title("Word_to_PDF Converter.")
        self.minsize(300,300)

        self.wm_iconbitmap('icon.ico')
        self.config(bg= '#0059b3')
        self.resizable(width=False, height=False)


        self.lableFrame = ttk.LabelFrame(self, text = "                           Open your Word File",relief= "groove")
        self.lableFrame.grid(column = 1, row = 1, padx =20,pady = 20,sticky=N + S + E + W)
        self.lable = ttk.Label(self, text="Develop by Addy",relief= "groove")
        self.lable.place(x=100, y=306)
        self.button()
        self.make_dir()
        self.button1()

    #define Function for Button1.
    def button(self):
        self.button = ttk.Button(self.lableFrame, text = "Browse a File", command = self.fileDialog)
        self.img = PhotoImage(file="btn1.jpg")  # make sure to add "/" not "\"
        self.button.config(image=self.img)
        self.button.grid(column =1, row = 1)

    # define Function for Button2.
    def button1(self):
        self.button1 = ttk.Button(self.lableFrame, text = "Convert File", command = self.convert)
        self.img1 = PhotoImage(file="btn2.jpg")  # make sure to add "/" not "\"
        self.button1.config(image=self.img1)
        self.button1.grid(column =1, row = 2, padx= 20, pady= 50)

    # define Function for Dialog box for file.
    def fileDialog(self):
        self.filename = filedialog.askopenfilename(initialdir = "/", title = "Select a File", filetype = (("docx", "*.docx"),("All Files", "*.*")))
        self.lable = ttk.Label(self.lableFrame, text = "")
        print(self.filename)
    #define function to create a new folder for output of the file.
    def make_dir(self):
        path = 'D:/Doc_2_PDF (Output)'
        try:
            os.mkdir(path)
        except OSError as error:
            print()
    #define function to convert the docx file into pdf.
    def convert(self):
        i = str(randint(1,1000))
        self.input_file = self.filename
        self.output_file = ('D:/Doc_2_PDF (Output)/output' + i + '.pdf')
        convert(self.input_file,self.output_file)


if __name__ == '__main__':
    root =Root()
    root.mainloop()
