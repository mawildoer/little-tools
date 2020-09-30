import pathlib
import hashlib

import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog


class GUI:
    def __init__(self, master):
        self.master = master
        master.title("FancyFill")

        self.instructions = tk.Message(master, text="Select two files and it'll tell "
                                                    "you whether they're identical", width=300)
        self.instructions.pack(side=tk.TOP)

        self.filepath_1_button = tk.Button(master, text="Select File", command=self.select_1).pack()
        self.filepath_1_label_text = tk.StringVar(master)
        self.filepath_1_label = tk.Label(master, textvariable=self.filepath_1_label_text)
        self.filepath_1_label_text.set("Copy the data you want to fill.")
        self.filepath_1_label.pack()
        self.filepath_1 = None

        self.filepath_2_button = tk.Button(master, text="Select File", command=self.select_2).pack()
        self.filepath_2_label_text = tk.StringVar(master)
        self.filepath_2_label = tk.Label(master, textvariable=self.filepath_2_label_text)
        self.filepath_2_label_text.set("Copy the data you want to fill.")
        self.filepath_2_label.pack()
        self.filepath_2 = None


    def run_hashing(self):
        f1_sha1 = hashlib.sha1()
        f2_sha1 = hashlib.sha1()

        with self.filepath_1.open(mode="rb") as f1:
            while True:
                data = f1.read(65535)
                if not data:
                    break
                f1_sha1.update(data)

        with self.filepath_2.open(mode="rb") as f2:
            while True:
                data = f2.read(65535)
                if not data:
                    break
                f2_sha1.update(data)

        if f1_sha1.hexdigest() == f2_sha1.hexdigest():
            messagebox.showinfo("", "Same Same")
        else:
            messagebox.showinfo("", "Different")

    def select_1(self):
        try:
            self.filepath_1 = pathlib.Path(
                 filedialog.askopenfilename(title="Select a file"))
        except ValueError as e:
            messagebox.showinfo("", "Bad file selection. Please try again.")
        else:
            if self.filepath_2 is not None:
                self.run_hashing()

    def select_2(self):
        try:
            self.filepath_2 = pathlib.Path(
                 filedialog.askopenfilename(title="Select another file"))
        except ValueError as e:
            messagebox.showinfo("", "Bad file selection. Please try again.")
        else:
            if self.filepath_1 is not None:
                self.run_hashing()


if __name__ == '__main__':
    root = tk.Tk()
    gui = GUI(root)
    root.mainloop()
