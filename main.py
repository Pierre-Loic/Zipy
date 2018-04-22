# Package imports
import zipfile
import subprocess
import os
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
from tkinter import *


class Window:
    def __init__(self, master):
        # Variables
        self.file_sel = ""

        master.title("ZIP")
        # General structure of the window
        self.frame1 = ttk.LabelFrame(master, height=100, width=400, text="Actions")
        self.frame1.pack(padx=10, pady=10)
        self.frame2 = ttk.LabelFrame(master, height=100, width=400, text="Information about zip file")
        self.frame2.pack(padx=10, pady=10)
        self.frame3 = ttk.LabelFrame(master, height=100, width=400, text="List of files and folders inside zip file")
        self.frame3.pack(padx=10, pady=10)
        # Frame 1 : buttons
        # Button 1
        self.button1 = ttk.Button(self.frame1, text="Browse zip file", command=self.button_file)
        self.button1.pack(side=LEFT, padx=10, pady=10)
        self.logo1 = PhotoImage(file="D:\\A_graver\\micro-entreprise\\Mentorat_OC\\Projets_informatiques\\Zip\\res\\zipfile.png")
        self.small_logo1 = self.logo1.subsample(6, 6)
        self.button1.config(image=self.small_logo1, compound=LEFT)
        # Button 2
        self.button2 = ttk.Button(self.frame1, text="Unzip directory", command=self.change_dir)
        self.button2.pack(side=LEFT, padx=10, pady=10)
        self.logo2 = PhotoImage(file="D:\\A_graver\\micro-entreprise\\Mentorat_OC\\Projets_informatiques\\Zip\\res\\folder.png")
        self.small_logo2 = self.logo2.subsample(6, 6)
        self.button2.config(image=self.small_logo2, compound=LEFT)
        # Button 3
        self.button3 = ttk.Button(self.frame1, text="ExtractAll")
        self.button3.pack(side=LEFT, padx=10, pady=10)
        self.logo3 = PhotoImage(file="D:\\A_graver\\micro-entreprise\\Mentorat_OC\\Projets_informatiques\\Zip\\res\\extract_all.png")
        self.small_logo3 = self.logo3.subsample(6, 6)
        self.button3.config(image=self.small_logo3, compound=LEFT, state="disabled", command=self.button_extract_all)
        # Button 4
        self.button4 = ttk.Button(self.frame1, text="Extract Selection")
        self.button4.pack(side=LEFT, padx=10, pady=10)
        self.logo4 = PhotoImage(file="D:\\A_graver\\micro-entreprise\\Mentorat_OC\\Projets_informatiques\\Zip\\res\\select.png")
        self.small_logo4 = self.logo4.subsample(12, 12)
        self.button4.config(image=self.small_logo4, compound=LEFT, state="disabled", command=self.button_extract_sel)
        # Frame 2 : data
        self.label_frame_2 = ttk.Label(self.frame2, text="Select a zip file", width=100)
        self.label_frame_2.pack(padx=10, pady=10)
        # Frame 3 : files list
        self.treeview = ttk.Treeview(self.frame3, show="tree", selectmode="extended")
        self.treeview.pack(padx=10, pady=10)
        self.treeview["column"]=("one")
        self.treeview.column("one", width=400)
        self.treeview.bind("<<TreeviewSelect>>", self.callback)
        # Bottom label
        self.label1 = ttk.Label(master)
        self.label1.pack()


    def callback(self, event):
        """ Enable the button 4 """
        self.button4.config(state="enabled")
        print(self.treeview.selection())

    def button_file(self):
        """ Select zip file """
        self.file_sel = self.find_file("file")
        if self.is_zip(self.file_sel):
            zip1 = ZipData(self.file_sel)
            self.update_label(1)
            self.label_frame_2.config(text="You have selected the zip file " + self.file_sel + "\n" + str(zip1.len_zip()) + " items found in the zip file")
            self.button3.config(state="enabled")
            print(zip1.info())
            for elt in zip1.info():
                elt_split = elt.split("/")
                if "" in elt_split:
                    elt_split.remove("")
                if len(elt_split)==1:
                    self.treeview.insert("", "end", elt_split[0], text=elt_split[0])
                elif len(elt_split)>1:
                    self.treeview.insert("/".join(elt_split[0:-1]), "end", "/".join(elt_split), text=elt_split[-1])

    def button_extract_all(self):
        """ Extract all files in the zip file """
        if self.file_sel != "":
            zip1 = ZipData(self.file_sel)
            zip1.extract_all()
            self.clean_window()
        else:
            messagebox.showerror("Zip file missing", "You need to select a zip file")

    def button_extract_sel(self):
        """ Extract selected files in the zip file """
        if self.file_sel != "":
            zip1 = ZipData(self.file_sel)
            zip1.extract_sel(self.treeview.selection())
            self.clean_window()
        else:
            messagebox.showerror("Zip file missing", "You need to select a zip file")

    def clean_window(self):
        """ Clean the main window after extraction """
        self.update_label(2)
        self.label_frame_2.config(text="")
        for i in self.treeview.get_children():
            self.treeview.delete(i)
        subprocess.Popen("explorer " + os.getcwd())

    @staticmethod
    def find_file(f_type):
        """ Return the path of the zip file or folder """
        f_name = ""
        if f_type == "file":
            file_name = filedialog.askopenfile()
            f_name = file_name.name
        if f_type == "folder":
            f_name = filedialog.askdirectory()
        return f_name

    @staticmethod
    def find_dir():
        """ Find the current directory """
        cwd = os.getcwd()
        return cwd

    @staticmethod
    def is_zip(zip_path):
        """ Check if the selected file is a zip file """
        if zip_path[-4:]!=".zip":
            messagebox.showerror("Type file error", "You need to select a zip file")
            return False
        else:
            return True

    def change_dir(self):
        """ Change the current directory """
        dir_path = self.find_file("folder")
        os.chdir(dir_path)
        self.update_label(1)

    def update_label(self, val):
        """ Update the text of the label at the bottom of the window """
        if val==1:
            self.label1.config(text="The zip file will be extracted in the directory " + self.find_dir())
        if val==2:
            self.label1.config(text="Zip file successfully extracted")

class ZipData:
    def __init__(self, file):
        self.file = file

    def extract_all(self):
        """ Extract all the files of the zip folder """
        with zipfile.ZipFile(self.file, 'r') as z:
            z.extractall()

    def extract_sel(self, file_name):
        """ Extract selected files """
        with zipfile.ZipFile(self.file, 'r') as z:
            for f_name in file_name:
                z.extract(f_name)

    def len_zip(self):
        """ Count the number of items in the zip file """
        with zipfile.ZipFile(self.file, 'r') as z:
            nb_items = len(z.namelist())
            return nb_items

    def info(self):
        """ Return the list of items inside the zip file """
        with zipfile.ZipFile(self.file, 'r') as z:
            return z.namelist()


if __name__ == "__main__":
    root = Tk()
    win = Window(root)
    root.mainloop()
