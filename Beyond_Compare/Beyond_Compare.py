import sys
import difflib
import tkinter as tk
from tkinter import filedialog, scrolledtext

class FileComparisonTool:
    def __init__(self, master):
        self.master = master
        master.title('File Comparison Tool')
        master.geometry('1000x600')

        self.text_area1 = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=50, height=30)
        self.text_area2 = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=50, height=30)

        load_button1 = tk.Button(master, text='Load File 1', command=lambda: self.load_file(1))
        load_button2 = tk.Button(master, text='Load File 2', command=lambda: self.load_file(2))
        compare_button = tk.Button(master, text='Compare Files', command=self.compare_files)

        self.text_area1.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.text_area2.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        load_button1.grid(row=1, column=0, padx=5, pady=5)
        load_button2.grid(row=1, column=1, padx=5, pady=5)
        compare_button.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        master.grid_columnconfigure(0, weight=1)
        master.grid_columnconfigure(1, weight=1)
        master.grid_rowconfigure(0, weight=1)

    def load_file(self, text_area_num):
        file_path = filedialog.askopenfilename()
        if file_path:
            with open(file_path, 'r') as file:
                content = file.read()
                if text_area_num == 1:
                    self.text_area1.delete('1.0', tk.END)
                    self.text_area1.insert(tk.END, content)
                else:
                    self.text_area2.delete('1.0', tk.END)
                    self.text_area2.insert(tk.END, content)

    def compare_files(self):
        text1 = self.text_area1.get('1.0', tk.END).splitlines()
        text2 = self.text_area2.get('1.0', tk.END).splitlines()

        differ = difflib.Differ()
        diff = list(differ.compare(text1, text2))

        self.text_area1.delete('1.0', tk.END)
        self.text_area2.delete('1.0', tk.END)

        for line in diff:
            if line.startswith('  '):
                self.text_area1.insert(tk.END, line[2:] + '\n')
                self.text_area2.insert(tk.END, line[2:] + '\n')
            elif line.startswith('- '):
                self.text_area1.insert(tk.END, line[2:] + '\n', 'removed')
                self.text_area2.insert(tk.END, '\n')
            elif line.startswith('+ '):
                self.text_area1.insert(tk.END, '\n')
                self.text_area2.insert(tk.END, line[2:] + '\n', 'added')

        self.text_area1.tag_configure('removed', background='#ffcccc')
        self.text_area2.tag_configure('added', background='#ccffcc')

if __name__ == '__main__':
    root = tk.Tk()
    app = FileComparisonTool(root)
    root.mainloop()
