import sys
import difflib
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk

class FileComparisonTool:
    def __init__(self, master):
        self.master = master
        master.title('File Comparison Tool')
        master.geometry('1000x600')

        # Create frames for search functionality
        self.search_frame1 = tk.Frame(master)
        self.search_frame2 = tk.Frame(master)

        # Create search entries and buttons
        self.search_var1 = tk.StringVar()
        self.search_var2 = tk.StringVar()
        
        self.search_entry1 = ttk.Entry(self.search_frame1, textvariable=self.search_var1)
        self.search_button1 = ttk.Button(self.search_frame1, text='Find', command=lambda: self.find_text(1))
        self.next_button1 = ttk.Button(self.search_frame1, text='Next Diff', command=lambda: self.jump_to_next_diff(1))
        
        self.search_entry2 = ttk.Entry(self.search_frame2, textvariable=self.search_var2)
        self.search_button2 = ttk.Button(self.search_frame2, text='Find', command=lambda: self.find_text(2))
        self.next_button2 = ttk.Button(self.search_frame2, text='Next Diff', command=lambda: self.jump_to_next_diff(2))

        # Pack search components
        self.search_entry1.pack(side=tk.LEFT, padx=5)
        self.search_button1.pack(side=tk.LEFT, padx=5)
        self.next_button1.pack(side=tk.LEFT, padx=5)
        
        self.search_entry2.pack(side=tk.LEFT, padx=5)
        self.search_button2.pack(side=tk.LEFT, padx=5)
        self.next_button2.pack(side=tk.LEFT, padx=5)

        self.text_area1 = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=50, height=30)
        self.text_area2 = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=50, height=30)

        load_button1 = tk.Button(master, text='Load File 1', command=lambda: self.load_file(1))
        load_button2 = tk.Button(master, text='Load File 2', command=lambda: self.load_file(2))
        compare_button = tk.Button(master, text='Compare Files', command=self.compare_files)

        # Grid layout
        self.search_frame1.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.search_frame2.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.text_area1.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        self.text_area2.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
        load_button1.grid(row=2, column=0, padx=5, pady=5)
        load_button2.grid(row=2, column=1, padx=5, pady=5)
        compare_button.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

        # Configure grid weights
        master.grid_columnconfigure(0, weight=1)
        master.grid_columnconfigure(1, weight=1)
        master.grid_rowconfigure(1, weight=1)

        # Initialize search variables
        self.last_search_index1 = '1.0'
        self.last_search_index2 = '1.0'
        self.diff_positions1 = []  # Store positions of differences in text_area1
        self.diff_positions2 = []  # Store positions of differences in text_area2
        self.current_diff_index1 = 0
        self.current_diff_index2 = 0

    def find_text(self, text_area_num):
        if text_area_num == 1:
            text_area = self.text_area1
            search_text = self.search_var1.get()
            self.last_search_index1 = self.find_next(text_area, search_text, self.last_search_index1)
        else:
            text_area = self.text_area2
            search_text = self.search_var2.get()
            self.last_search_index2 = self.find_next(text_area, search_text, self.last_search_index2)

    def find_next(self, text_area, search_text, last_index):
        text_area.tag_remove('search', '1.0', tk.END)
        
        if search_text:
            pos = text_area.search(search_text, last_index, nocase=1, stopindex=tk.END)
            if not pos:
                pos = text_area.search(search_text, '1.0', nocase=1, stopindex=tk.END)
                if not pos:
                    return '1.0'
            
            end_pos = f"{pos}+{len(search_text)}c"
            text_area.tag_add('search', pos, end_pos)
            text_area.tag_config('search', background='yellow')
            text_area.see(pos)
            
            return end_pos
        return '1.0'

    def jump_to_next_diff(self, text_area_num):
        if text_area_num == 1:
            positions = self.diff_positions1
            text_area = self.text_area1
            if positions:
                self.current_diff_index1 = (self.current_diff_index1 + 1) % len(positions)
                text_area.see(positions[self.current_diff_index1])
        else:
            positions = self.diff_positions2
            text_area = self.text_area2
            if positions:
                self.current_diff_index2 = (self.current_diff_index2 + 1) % len(positions)
                text_area.see(positions[self.current_diff_index2])

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
        
        # Reset diff positions
        self.diff_positions1 = []
        self.diff_positions2 = []
        self.current_diff_index1 = 0
        self.current_diff_index2 = 0

        current_pos1 = '1.0'
        current_pos2 = '1.0'

        for line in diff:
            if line.startswith('  '):
                self.text_area1.insert(tk.END, line[2:] + '\n')
                self.text_area2.insert(tk.END, line[2:] + '\n')
                current_pos1 = self.text_area1.index('end-1c')
                current_pos2 = self.text_area2.index('end-1c')
            elif line.startswith('- '):
                self.diff_positions1.append(current_pos1)
                self.text_area1.insert(tk.END, line[2:] + '\n', 'removed')
                self.text_area2.insert(tk.END, '\n')
                current_pos1 = self.text_area1.index('end-1c')
                current_pos2 = self.text_area2.index('end-1c')
            elif line.startswith('+ '):
                self.diff_positions2.append(current_pos2)
                self.text_area1.insert(tk.END, '\n')
                self.text_area2.insert(tk.END, line[2:] + '\n', 'added')
                current_pos1 = self.text_area1.index('end-1c')
                current_pos2 = self.text_area2.index('end-1c')

        self.text_area1.tag_configure('removed', background='#ffcccc')
        self.text_area2.tag_configure('added', background='#ccffcc')

if __name__ == '__main__':
    root = tk.Tk()
    app = FileComparisonTool(root)
    root.mainloop()
