import os
import glob
import base64
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

# The TF icon in Base64
ICON_B64 = b"""
iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAABAUlEQVR4nO3asS4EURhH8bOHSnSegEThXUTlJTzONsQzUHkNajovgIpsIbF0mk3GLInNYX/lnf+dOzfz5X5TzGT39JkyiZM4iZM4iZM4iZM4iZM4iZM4iZM4iZM4iZM4iZM4iZM4iZO4zZG5+5Ptb9x97+xl2enTm9fp9ev4Jf7NGzi8mC0OOuHqeGsgsOjocjZ/Hwo8zAYvf3sDt4/zxcENvwgsunuav40K/psSkjiJkziJk7jNX17vYMeBPnA77ixe5QY+G9/wp8d46xJa0v75y7oT/60SkjiJkziJkziJk7jJ+n+hFZM4iZM4iZM4iZM4iZM4iZM4iZM4V/0AP/UBSlAuQGVruTEAAAAASUVORK5CYII=
"""

def parse_dat_file(filepath):
    """Parses a single DAT file and returns a list of dictionaries."""
    data = []
    try:
        with open(filepath, 'r') as f:
            for line in f:
                line = line.strip('\r\n')
                if not line:
                    continue
                try:
                    ref_no = line[10:24].strip()
                    current_read = line[46:63].strip()
                    units = line[63:71].strip()
                    tf_id = line[-11:].strip()
                except IndexError:
                    continue
                
                try:
                    current_read = float(current_read)
                except ValueError:
                    pass
                
                try:
                    units = int(units)
                except ValueError:
                    pass
                
                data.append({
                    "Ref No.": ref_no,
                    "Current Read": current_read,
                    "Units": units,
                    "TF ID": tf_id
                })
        return data
    except Exception as e:
        print(f"Error parsing {filepath}: {str(e)}")
        return []

class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("TF Readings")
        self.geometry("500x380")
        self.eval('tk::PlaceWindow . center')
        
        # Set suitable icon
        try:
            icon_img = tk.PhotoImage(data=ICON_B64)
            self.iconphoto(True, icon_img)
        except Exception:
            pass # fallback if icon fails
        
        self.files_to_process = []
        
        # UI Elements
        self.lbl_title = tk.Label(self, text="TF Readings", font=("Helvetica", 18, "bold"), fg="#1976D2")
        self.lbl_title.pack(pady=15)
        
        # Drag and Drop Area
        self.lbl_drop = tk.Label(
            self, 
            text="Drag and Drop .DAT Files or Folders Here\n\nor use buttons below",
            bg="#f0f0f0", 
            relief="groove",
            width=50,
            height=6,
            font=("Helvetica", 10)
        )
        self.lbl_drop.pack(pady=5)
        
        # Enable Drag and Drop
        self.lbl_drop.drop_target_register(DND_FILES)
        self.lbl_drop.dnd_bind('<<Drop>>', self.handle_drop)
        
        # Frame for buttons
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)
        
        self.btn_select_file = tk.Button(btn_frame, text="Select File(s)", command=self.select_files, width=15)
        self.btn_select_file.grid(row=0, column=0, padx=10)
        
        self.btn_select_folder = tk.Button(btn_frame, text="Select Folder", command=self.select_folder, width=15)
        self.btn_select_folder.grid(row=0, column=1, padx=10)
        
        self.lbl_status = tk.Label(self, text="0 files selected", fg="gray", font=("Helvetica", 9, "bold"))
        self.lbl_status.pack(pady=5)
        
        self.btn_export = tk.Button(
            self, text="Export to single Excel file", 
            command=self.export_excel, 
            state=tk.DISABLED, 
            width=25, 
            bg="#4CAF50", fg="white", activebackground="#45a049",
            font=("Helvetica", 10, "bold")
        )
        self.btn_export.pack(pady=5)
        
        self.lbl_footer = tk.Label(self, text="Developed by Noman Asad", font=("Helvetica", 9, "italic"), fg="gray")
        self.lbl_footer.pack(side=tk.BOTTOM, pady=10)
        
    def handle_drop(self, event):
        files = self.tk.splitlist(event.data)
        self.process_selected_paths(files)
        
    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select DAT files",
            filetypes=(("DAT Files", "*.DAT;*.dat"), ("All Files", "*.*"))
        )
        if files:
            self.process_selected_paths(files)
            
    def select_folder(self):
        folder = filedialog.askdirectory(title="Select Folder containing DAT files")
        if folder:
            self.process_selected_paths([folder])
            
    def process_selected_paths(self, paths):
        found_files = set()
        for path in paths:
            # TkinterDnD sometimes adds curly braces around paths with spaces
            path = path.strip('{}') 
            
            if os.path.isdir(path):
                # Search for dat files in folder
                dat_files = glob.glob(os.path.join(path, "*.dat")) + glob.glob(os.path.join(path, "*.DAT"))
                found_files.update(dat_files)
            elif os.path.isfile(path) and path.lower().endswith('.dat'):
                found_files.add(path)
                
        if found_files:
            # Append new files without duplicates
            for f in found_files:
                if f not in self.files_to_process:
                    self.files_to_process.append(f)
                    
            self.update_status()
        else:
            messagebox.showinfo("No Files", "No .DAT files found in the selection.")
            
    def update_status(self):
        count = len(self.files_to_process)
        self.lbl_status.config(text=f"{count} file(s) ready for export", fg="#1976D2")
        if count > 0:
            self.btn_export.config(state=tk.NORMAL)
            self.lbl_drop.config(bg="#e8f5e9") # Greenish to indicate success
        else:
            self.btn_export.config(state=tk.DISABLED)
            self.lbl_drop.config(bg="#f0f0f0")
            self.lbl_status.config(fg="gray")
            
    def export_excel(self):
        if not self.files_to_process:
            return
            
        all_data = []
        for file in self.files_to_process:
            file_data = parse_dat_file(file)
            all_data.extend(file_data)
            
        if not all_data:
            messagebox.showwarning("Warning", "No valid data could be extracted from the selected files.")
            return
            
        try:
            df = pd.DataFrame(all_data)
            
            # Ask where to save
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile="TF_Readings_Export.xlsx",
                title="Save Excel File As",
                filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*"))
            )
            
            if save_path:
                df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"Successfully exported {len(all_data)} records to Excel!")
                # Reset after successful export
                self.files_to_process = []
                self.update_status()
                
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
