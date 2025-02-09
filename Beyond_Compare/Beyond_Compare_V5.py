#!/usr/bin/env python3
import sys
import difflib
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk

class FileComparisonTool:
    def __init__(self, master):
        self.master = master
        master.title('File Comparison Tool')
        master.geometry('1000x600')

        # Initialize sync scrolling variables and a flag to avoid recursion.
        self.sync_enabled = tk.BooleanVar(value=False)
        self.current_widget = None
        self.sync_in_progress = False

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

        # Create text areas using ScrolledText
        self.text_area1 = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=50, height=30)
        self.text_area2 = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=50, height=30)

        # Set tag configurations once.
        self.text_area1.tag_configure("removed", background="#ffcccc")
        self.text_area2.tag_configure("added", background="#ccffcc")
        self.text_area1.tag_configure("search", background="yellow")
        self.text_area2.tag_configure("search", background="yellow")

        # Add mouse enter bindings for scroll sync
        self.text_area1.bind('<Enter>', lambda e: setattr(self, 'current_widget', self.text_area1))
        self.text_area2.bind('<Enter>', lambda e: setattr(self, 'current_widget', self.text_area2))

        # Create control buttons
        load_button1 = tk.Button(master, text='Load File 1', command=lambda: self.load_file(1))
        load_button2 = tk.Button(master, text='Load File 2', command=lambda: self.load_file(2))
        compare_button = tk.Button(master, text='Compare Files', command=self.compare_files)

        # Create control frame for the sync scrolling checkbutton
        self.control_frame = tk.Frame(master)
        self.sync_button = ttk.Checkbutton(
            self.control_frame,
            text="Sync Scrolling",
            variable=self.sync_enabled,
            command=self.toggle_sync_scroll
        )
        self.sync_button.pack(side=tk.LEFT, padx=5)

        # Layout with grid
        self.search_frame1.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.search_frame2.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.text_area1.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        self.text_area2.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
        load_button1.grid(row=2, column=0, padx=5, pady=5)
        load_button2.grid(row=2, column=1, padx=5, pady=5)
        compare_button.grid(row=2, column=0, columnspan=2, padx=5, pady=5)
        self.control_frame.grid(row=3, column=0, columnspan=2, pady=5)

        # Configure grid weights for proper resizing
        master.grid_columnconfigure(0, weight=1)
        master.grid_columnconfigure(1, weight=1)
        master.grid_rowconfigure(1, weight=1)

        # Initialize search and diff variables
        self.last_search_index1 = '1.0'
        self.last_search_index2 = '1.0'
        self.diff_positions1 = []
        self.diff_positions2 = []
        self.current_diff_index1 = 0
        self.current_diff_index2 = 0

    def sync_scroll(self, *args):
        # Use a guard flag to avoid recursive callbacks.
        if self.sync_in_progress:
            return
        if self.sync_enabled.get():
            self.sync_in_progress = True
            # args[0] is the first argument (the fraction) passed by the text widget.
            if self.current_widget == self.text_area1:
                self.text_area2.yview_moveto(args[0])
            elif self.current_widget == self.text_area2:
                self.text_area1.yview_moveto(args[0])
            self.sync_in_progress = False

    def toggle_sync_scroll(self):
        if self.sync_enabled.get():
            # When enabling sync scrolling, assign our custom yscrollcommand function
            self.text_area1.config(yscrollcommand=lambda *args: self.sync_scroll(*args))
            self.text_area2.config(yscrollcommand=lambda *args: self.sync_scroll(*args))
        else:
            # Restore default yscrollcommand when sync scrolling is disabled.
            self.text_area1.config(yscrollcommand=self.text_area1.yview)
            self.text_area2.config(yscrollcommand=self.text_area2.yview)

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
            # You could add an encoding parameter if needed (e.g. encoding="utf-8")
            with open(file_path, 'r') as file:
                content = file.read()
                if text_area_num == 1:
                    self.text_area1.delete('1.0', tk.END)
                    self.text_area1.insert(tk.END, content)
                else:
                    self.text_area2.delete('1.0', tk.END)
                    self.text_area2.insert(tk.END, content)

    def compare_files(self):
        # Use "end-1c" so that a stray final newline isn’t included.
        lines1 = self.text_area1.get('1.0', 'end-1c').splitlines()
        lines2 = self.text_area2.get('1.0', 'end-1c').splitlines()

        # Clear both text areas and reset difference positions/indexes
        self.text_area1.delete('1.0', tk.END)
        self.text_area2.delete('1.0', tk.END)
        self.diff_positions1 = []
        self.diff_positions2 = []
        self.current_diff_index1 = 0
        self.current_diff_index2 = 0

        # Compare the two files line‐by‐line using SequenceMatcher
        sm = difflib.SequenceMatcher(None, lines1, lines2)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                for line in lines1[i1:i2]:
                    self.text_area1.insert(tk.END, line + "\n")
                    self.text_area2.insert(tk.END, line + "\n")
            elif tag == "replace":
                n = max(i2 - i1, j2 - j1)
                for k in range(n):
                    old_line = lines1[i1 + k] if i1 + k < i2 else ""
                    new_line = lines2[j1 + k] if j1 + k < j2 else ""
                    self.diff_line_pair(old_line, new_line)
            elif tag == "delete":
                for line in lines1[i1:i2]:
                    pos = self.text_area1.index(tk.END)
                    self.diff_positions1.append(pos)
                    self.text_area1.insert(tk.END, line, "removed")
                    self.text_area1.insert(tk.END, "\n")
                    self.text_area2.insert(tk.END, "\n")
            elif tag == "insert":
                for line in lines2[j1:j2]:
                    pos = self.text_area2.index(tk.END)
                    self.diff_positions2.append(pos)
                    self.text_area1.insert(tk.END, "\n")
                    self.text_area2.insert(tk.END, line, "added")
                    self.text_area2.insert(tk.END, "\n")
        # (The tag styles are already configured in __init__)

    def diff_line_pair(self, old_line, new_line):
        ta1 = self.text_area1
        ta2 = self.text_area2

        # Record the first diff position encountered for this pair, if any.
        diff_found = False

        matcher = difflib.SequenceMatcher(None, old_line, new_line)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                segment = old_line[i1:i2]
                ta1.insert(tk.END, segment)
                ta2.insert(tk.END, segment)
            else:
                if not diff_found:
                    pos1 = ta1.index(tk.END)
                    pos2 = ta2.index(tk.END)
                    self.diff_positions1.append(pos1)
                    self.diff_positions2.append(pos2)
                    diff_found = True
                if tag in ("replace", "delete"):
                    segment_old = old_line[i1:i2]
                    ta1.insert(tk.END, segment_old, "removed")
                if tag in ("replace", "insert"):
                    segment_new = new_line[j1:j2]
                    ta2.insert(tk.END, segment_new, "added")
        ta1.insert(tk.END, "\n")
        ta2.insert(tk.END, "\n")
        # (No need to re-configure the tags here since they were set earlier.)

if __name__ == '__main__':
    root = tk.Tk()
    app = FileComparisonTool(root)
    root.mainloop()
