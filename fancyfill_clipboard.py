import clipboard
import keyboard
import re
import time

import tkinter as tk
from tkinter import messagebox


class GUI:
    def __init__(self, master):
        self.master = master
        master.title("FancyFill")

        self.instructions = tk.Message(master, text="'>>' will sent a key for example 'enter' or 'tab'. \n"
                                                    "Each new cell (excluding '>>') is prefaced as a 'tab'. \n"
                                                    "After hitting 'fill', you've got 3 seconds", width=300)
        self.instructions.pack()

        self.label_text = tk.StringVar()
        self.label = tk.Label(master, textvariable=self.label_text)
        self.label_text.set("Copy the data you want to fill.")
        self.label.pack()

        self.start_button = tk.Button(master, text="Fill", command=self.fill)
        self.start_button.pack()

    def fill(self):
        data = re.split(r'\r\n|\t', clipboard.paste())
        if len(data) == 0:
            messagebox.showinfo("Nothing to paste.",
                                "There's nothing in your clipboard. Go to excel and copy something first.")
            return

        for tminus in range(3, 0, -1):
            self.label_text.set("Starting in {}...".format(tminus))
            time.sleep(1)
        self.label_text.set("Press ESC to stop.")

        hold_tab = True
        for c in data:
            if keyboard.is_pressed('esc'):
                break

            if c.startswith('>>'):
                sequence = [s.strip() for s in c[2:].split(",")]
                for k in sequence:
                    if k.startswith("pause"):
                        t = float(k[5:])
                        time.sleep(t/1000)
                    else:
                        keyboard.press_and_release(k)
                hold_tab = True
            else:
                if hold_tab:
                    hold_tab = False
                else:
                    keyboard.press_and_release('tab')
                keyboard.write(str(c))
            time.sleep(0.05)


if __name__ == '__main__':
    root = tk.Tk()
    gui = GUI(root)
    root.mainloop()
