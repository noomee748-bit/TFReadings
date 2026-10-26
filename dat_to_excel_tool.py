import os
import glob
import base64
import re
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

# The TF icon in Base64
ICON_B64 = b"""
iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAABAUlEQVR4nO3asS4EURhH8bOHSnSegEThXUTlJTzONsQzUHkNajovgIpsIbF0mk3GLInNYX/lnf+dOzfz5X5TzGT39JkyiZM4iZM4iZM4iZM4iZM4iZM4iZM4iZM4iZM4iZM4iZO4zZG5+5Ptb9x97+xl2enTm9fp9ev4Jf7NGzi8mC0OOuHqeGsgsOjocjZ/Hwo8zAYvf3sDt4/zxcENvwgsunuav40K/psSkjiJkziJk7jNX17vYMeBPnA77ixe5QY+G9/wp8d46xJa0v75y7oT/60SkjiJkziJkziJk7jJ+n+hFZM4iZM4iZM4iZM4iZM4iZM4iZM4V/0AP/UBSlAuQGVruTEAAAAASUVORK5CYII=
"""

def extract_tf_ids_from_excel(filepath):
    """
    Reads an Excel (.xlsx, .xls) or CSV file and extracts unique TF IDs.
    Auto-detects the column named 'TF_ID', 'TF ID', 'TFID', or looks for column with matching IDs.
    """
    if filepath.lower().endswith('.csv'):
        df = pd.read_csv(filepath, dtype=str)
    else:
        df = pd.read_excel(filepath, dtype=str)
        
    col_found = None
    # 1. Match by column name
    for col in df.columns:
        norm = re.sub(r'[\s_\-]', '', str(col)).lower()
        if norm in ['tfid', 'tfids', 'tf', 'transformerid', 'transformer']:
            col_found = col
            break
            
    # 2. Match by values if column name doesn't match
    if col_found is None:
        best_col = None
        best_score = 0
        for col in df.columns:
            vals = df[col].dropna().astype(str).str.strip()
            if len(vals) == 0:
                continue
            score = sum(1 for v in vals if re.match(r'^[A-Za-z0-9]{7,15}$', v))
            if score > best_score:
                best_score = score
                best_col = col
        col_found = best_col if best_col is not None else df.columns[0]
        
    vals = df[col_found].dropna().astype(str).str.strip()
    result = set(v.upper() for v in vals if v and v.lower() not in ['nan', 'none'])
    
    # If the header row itself was actually an ID
    header_str = str(col_found).strip().upper()
    if re.match(r'^[A-Za-z0-9]{7,15}$', header_str) and header_str.startswith('L'):
        result.add(header_str)
        
    return result, str(col_found)

def parse_dat_for_excel(filepath):
    """Parses a single DAT file and returns a list of dictionaries for Excel export."""
    data = []
    try:
        with open(filepath, 'r', encoding='latin1', errors='ignore') as f:
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

class TFReadingsApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("TF Readings - Manager")
        self.geometry("640x700")
        self.minsize(600, 650)
        self.eval('tk::PlaceWindow . center')
        
        # Set icon
        try:
            icon_img = tk.PhotoImage(data=ICON_B64)
            self.iconphoto(True, icon_img)
        except Exception:
            pass
            
        # State variables for Tab 1 (Filter DAT by Excel)
        self.excel_path_var = tk.StringVar(value="")
        self.excel_tf_ids = set()
        self.excel_col_name = ""
        self.dest_folder_var = tk.StringVar(value="")
        self.filter_dat_files = []
        self.renumber_var = tk.BooleanVar(value=False)
        
        # State variables for Tab 2 (Export DAT to Excel)
        self.export_dat_files = []
        
        self.setup_ui()
        
    def setup_ui(self):
        # Header
        header_frame = tk.Frame(self, bg="#1565C0", pady=10)
        header_frame.pack(fill=tk.X)
        
        lbl_heading = tk.Label(
            header_frame, text="TF Readings", 
            font=("Helvetica", 20, "bold"), fg="white", bg="#1565C0"
        )
        lbl_heading.pack()
        
        lbl_subheading = tk.Label(
            header_frame, text="Filter .DAT by Excel TF IDs & Export Readings", 
            font=("Helvetica", 10), fg="#BBDEFB", bg="#1565C0"
        )
        lbl_subheading.pack()
        
        # Notebook for Tabs
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TNotebook.Tab", font=("Helvetica", 10, "bold"), padding=[12, 6])
        
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.tab_filter = tk.Frame(self.notebook, bg="#FAFAFA")
        self.tab_export = tk.Frame(self.notebook, bg="#FAFAFA")
        
        self.notebook.add(self.tab_filter, text="  1. Filter .DAT by Excel TF IDs  ")
        self.notebook.add(self.tab_export, text="  2. Convert .DAT to Excel  ")
        
        self.build_filter_tab()
        self.build_export_tab()
        
        # Footer
        footer_frame = tk.Frame(self, bg="#ECEFF1", pady=6)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        lbl_footer = tk.Label(
            footer_frame, text="Developed by Noman Asad", 
            font=("Helvetica", 9, "bold italic"), fg="#455A64", bg="#ECEFF1"
        )
        lbl_footer.pack()

    # =========================================================================
    # TAB 1: Filter .DAT by Excel TF IDs
    # =========================================================================
    def build_filter_tab(self):
        parent = self.tab_filter
        
        # --- Section 1: Excel File Source ---
        sec1 = tk.LabelFrame(parent, text=" 1. Excel File (TF IDs Source) ", font=("Helvetica", 10, "bold"), bg="#FAFAFA", padx=10, pady=8)
        sec1.pack(fill=tk.X, padx=10, pady=6)
        
        row1 = tk.Frame(sec1, bg="#FAFAFA")
        row1.pack(fill=tk.X, pady=2)
        
        self.ent_excel = tk.Entry(row1, textvariable=self.excel_path_var, font=("Helvetica", 9), state="readonly")
        self.ent_excel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        
        btn_browse_excel = tk.Button(row1, text="Browse Excel...", command=self.browse_excel, bg="#E0E0E0", width=14)
        btn_browse_excel.pack(side=tk.RIGHT)
        
        self.lbl_excel_status = tk.Label(sec1, text="No Excel file loaded (Browse or Drag & Drop .xlsx/.xls/.csv)", fg="#757575", font=("Helvetica", 9), bg="#FAFAFA", anchor="w")
        self.lbl_excel_status.pack(fill=tk.X, pady=(4, 0))
        
        # Enable dropping on Excel frame
        sec1.drop_target_register(DND_FILES)
        sec1.dnd_bind('<<Drop>>', self.handle_excel_drop)
        self.ent_excel.drop_target_register(DND_FILES)
        self.ent_excel.dnd_bind('<<Drop>>', self.handle_excel_drop)
        
        # --- Section 2: Input .DAT Files ---
        sec2 = tk.LabelFrame(parent, text=" 2. Input .DAT Files to Filter ", font=("Helvetica", 10, "bold"), bg="#FAFAFA", padx=10, pady=8)
        sec2.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        
        # Drop Zone for DAT
        self.lbl_drop_filter = tk.Label(
            sec2, 
            text="Drag and Drop .DAT Files or Folders Here\n(or click buttons below)",
            bg="#E8EAF6", fg="#283593", relief="groove", height=3, font=("Helvetica", 9, "bold")
        )
        self.lbl_drop_filter.pack(fill=tk.X, pady=4)
        self.lbl_drop_filter.drop_target_register(DND_FILES)
        self.lbl_drop_filter.dnd_bind('<<Drop>>', self.handle_filter_dat_drop)
        
        # Button row
        btn_row = tk.Frame(sec2, bg="#FAFAFA")
        btn_row.pack(fill=tk.X, pady=4)
        
        btn_add_files = tk.Button(btn_row, text="Select .DAT Files", command=self.select_filter_files, width=15)
        btn_add_files.pack(side=tk.LEFT, padx=(0, 6))
        
        btn_add_folder = tk.Button(btn_row, text="Select Folder", command=self.select_filter_folder, width=15)
        btn_add_folder.pack(side=tk.LEFT, padx=(0, 6))
        
        btn_clear_dat = tk.Button(btn_row, text="Clear List", command=self.clear_filter_files, width=10)
        btn_clear_dat.pack(side=tk.RIGHT)
        
        # Listbox showing selected DAT files
        list_frame = tk.Frame(sec2)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=4)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.lst_filter_files = tk.Listbox(list_frame, font=("Courier", 9), yscrollcommand=scrollbar.set, selectmode=tk.EXTENDED)
        self.lst_filter_files.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.lst_filter_files.yview)
        
        self.lbl_filter_dat_count = tk.Label(sec2, text="0 .DAT file(s) selected", fg="#757575", font=("Helvetica", 9), bg="#FAFAFA", anchor="w")
        self.lbl_filter_dat_count.pack(fill=tk.X)
        
        # --- Section 3: Destination Folder & Options ---
        sec3 = tk.LabelFrame(parent, text=" 3. Destination Folder & Options ", font=("Helvetica", 10, "bold"), bg="#FAFAFA", padx=10, pady=8)
        sec3.pack(fill=tk.X, padx=10, pady=6)
        
        row3 = tk.Frame(sec3, bg="#FAFAFA")
        row3.pack(fill=tk.X, pady=2)
        
        self.ent_dest = tk.Entry(row3, textvariable=self.dest_folder_var, font=("Helvetica", 9), state="readonly")
        self.ent_dest.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        
        btn_browse_dest = tk.Button(row3, text="Browse Folder...", command=self.browse_dest_folder, bg="#E0E0E0", width=14)
        btn_browse_dest.pack(side=tk.RIGHT)
        
        chk_renumber = tk.Checkbutton(
            sec3, text="Renumber record counter (0000000001, 0000000002...)", 
            variable=self.renumber_var, bg="#FAFAFA", font=("Helvetica", 9)
        )
        chk_renumber.pack(anchor="w", pady=(4, 0))
        
        # --- Section 4: Export Action ---
        self.btn_export_filter = tk.Button(
            parent, 
            text="Export Filtered .DAT Files", 
            command=self.run_filter_export,
            bg="#2E7D32", fg="white", activebackground="#1B5E20", activeforeground="white",
            font=("Helvetica", 11, "bold"), height=2, state=tk.DISABLED
        )
        self.btn_export_filter.pack(fill=tk.X, padx=10, pady=(6, 10))

    # =========================================================================
    # TAB 1 Event Handlers & Methods
    # =========================================================================
    def browse_excel(self):
        f = filedialog.askopenfilename(
            title="Select Excel File with TF IDs",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if f:
            self.load_excel_file(f)
            
    def handle_excel_drop(self, event):
        paths = self.tk.splitlist(event.data)
        for p in paths:
            p = p.strip('{}')
            if p.lower().endswith(('.xlsx', '.xls', '.csv')):
                self.load_excel_file(p)
                break
                
    def load_excel_file(self, filepath):
        try:
            ids, col_name = extract_tf_ids_from_excel(filepath)
            if not ids:
                messagebox.showwarning("No TF IDs Found", "Could not find valid TF IDs in the selected file.")
                return
            self.excel_tf_ids = ids
            self.excel_col_name = col_name
            self.excel_path_var.set(filepath)
            self.lbl_excel_status.config(
                text=f"✔ Loaded {len(ids)} unique TF IDs from column '{col_name}'", 
                fg="#2E7D32", font=("Helvetica", 9, "bold")
            )
            self.check_filter_export_ready()
        except Exception as e:
            messagebox.showerror("Excel Error", f"Failed to read Excel file:\n{str(e)}")
            
    def browse_dest_folder(self):
        d = filedialog.askdirectory(title="Select Destination Folder for Filtered .DAT Files")
        if d:
            self.dest_folder_var.set(d)
            self.check_filter_export_ready()
            
    def select_filter_files(self):
        files = filedialog.askopenfilenames(
            title="Select DAT Files to Filter",
            filetypes=[("DAT Files", "*.DAT;*.dat"), ("All Files", "*.*")]
        )
        if files:
            self.add_filter_dat_paths(files)
            
    def select_filter_folder(self):
        folder = filedialog.askdirectory(title="Select Folder Containing .DAT Files")
        if folder:
            self.add_filter_dat_paths([folder])
            
    def handle_filter_dat_drop(self, event):
        paths = self.tk.splitlist(event.data)
        excel_found = False
        dat_paths = []
        for p in paths:
            clean = p.strip('{}')
            if clean.lower().endswith(('.xlsx', '.xls', '.csv')) and not self.excel_tf_ids:
                self.load_excel_file(clean)
                excel_found = True
            else:
                dat_paths.append(clean)
        if dat_paths:
            self.add_filter_dat_paths(dat_paths)
            
    def add_filter_dat_paths(self, paths):
        found = set()
        for p in paths:
            clean = p.strip('{}')
            if os.path.isdir(clean):
                found.update(glob.glob(os.path.join(clean, "*.dat")))
                found.update(glob.glob(os.path.join(clean, "*.DAT")))
            elif os.path.isfile(clean) and clean.lower().endswith('.dat'):
                found.add(clean)
                
        for f in sorted(found):
            if f not in self.filter_dat_files:
                self.filter_dat_files.append(f)
                self.lst_filter_files.insert(tk.END, f"{os.path.basename(f)}  ({f})")
                
        self.lbl_filter_dat_count.config(
            text=f"✔ {len(self.filter_dat_files)} .DAT file(s) ready", 
            fg="#1565C0" if self.filter_dat_files else "#757575"
        )
        self.check_filter_export_ready()
        
    def clear_filter_files(self):
        self.filter_dat_files.clear()
        self.lst_filter_files.delete(0, tk.END)
        self.lbl_filter_dat_count.config(text="0 .DAT file(s) selected", fg="#757575")
        self.check_filter_export_ready()
        
    def check_filter_export_ready(self):
        if bool(self.excel_tf_ids) and bool(self.filter_dat_files) and bool(self.dest_folder_var.get()):
            self.btn_export_filter.config(state=tk.NORMAL)
        else:
            self.btn_export_filter.config(state=tk.DISABLED)
            
    def run_filter_export(self):
        dest_dir = self.dest_folder_var.get()
        if not dest_dir or not os.path.exists(dest_dir):
            messagebox.showerror("Error", "Please select a valid destination folder.")
            return
            
        if not self.excel_tf_ids:
            messagebox.showerror("Error", "Please load an Excel file with TF IDs.")
            return
            
        if not self.filter_dat_files:
            messagebox.showerror("Error", "No .DAT files selected.")
            return
            
        renumber = self.renumber_var.get()
        results_summary = []
        total_kept_all = 0
        total_deleted_all = 0
        
        try:
            for file_path in self.filter_dat_files:
                fname = os.path.basename(file_path)
                out_path = os.path.join(dest_dir, fname)
                
                with open(file_path, 'r', encoding='latin1', errors='ignore') as f:
                    raw_lines = f.readlines()
                    
                total_in_file = 0
                kept_lines = []
                
                for line in raw_lines:
                    clean = line.strip('\r\n')
                    if not clean:
                        continue
                    total_in_file += 1
                    
                    # Last 11 characters are the TF ID
                    tf_id = clean[-11:].strip().upper()
                    if tf_id in self.excel_tf_ids:
                        kept_lines.append(clean)
                        
                deleted_in_file = total_in_file - len(kept_lines)
                total_kept_all += len(kept_lines)
                total_deleted_all += deleted_in_file
                
                # Apply renumbering if selected
                final_lines = []
                if renumber:
                    for idx, line in enumerate(kept_lines, start=1):
                        new_line = f"{idx:010d}" + line[10:]
                        final_lines.append(new_line)
                else:
                    final_lines = kept_lines
                    
                # Write to destination file
                with open(out_path, 'w', encoding='latin1', newline='') as f:
                    for line in final_lines:
                        f.write(line + '\r\n')
                        
                results_summary.append(
                    f"• {fname}:\n   {len(kept_lines)} kept, {deleted_in_file} deleted (Total: {total_in_file})"
                )
                
            report_msg = (
                f"Export Complete!\n\n"
                f"Destination: {dest_dir}\n"
                f"Files Processed: {len(self.filter_dat_files)}\n"
                f"Total Lines Kept: {total_kept_all}\n"
                f"Total Lines Deleted: {total_deleted_all}\n\n"
                f"File Breakdown:\n" + "\n".join(results_summary)
            )
            messagebox.showinfo("Export Successful", report_msg)
            
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred during export:\n{str(e)}")

    # =========================================================================
    # TAB 2: Convert .DAT to Excel (Preserved original functionality)
    # =========================================================================
    def build_export_tab(self):
        parent = self.tab_export
        
        lbl_info = tk.Label(
            parent, text="Extract Ref No., Current Read, Units, & TF ID to Excel",
            font=("Helvetica", 11, "bold"), fg="#1565C0", bg="#FAFAFA"
        )
        lbl_info.pack(pady=10)
        
        self.lbl_drop_export = tk.Label(
            parent, 
            text="Drag and Drop .DAT Files or Folders Here\n\nor use buttons below",
            bg="#E8EAF6", fg="#283593", relief="groove", width=50, height=6, font=("Helvetica", 10)
        )
        self.lbl_drop_export.pack(pady=5, padx=20, fill=tk.X)
        self.lbl_drop_export.drop_target_register(DND_FILES)
        self.lbl_drop_export.dnd_bind('<<Drop>>', self.handle_export_drop)
        
        btn_frame = tk.Frame(parent, bg="#FAFAFA")
        btn_frame.pack(pady=10)
        
        btn_sel_files = tk.Button(btn_frame, text="Select File(s)", command=self.select_export_files, width=15)
        btn_sel_files.grid(row=0, column=0, padx=8)
        
        btn_sel_folder = tk.Button(btn_frame, text="Select Folder", command=self.select_export_folder, width=15)
        btn_sel_folder.grid(row=0, column=1, padx=8)
        
        btn_clear = tk.Button(btn_frame, text="Clear", command=self.clear_export_files, width=10)
        btn_clear.grid(row=0, column=2, padx=8)
        
        self.lbl_export_status = tk.Label(parent, text="0 file(s) selected", fg="gray", font=("Helvetica", 10, "bold"), bg="#FAFAFA")
        self.lbl_export_status.pack(pady=10)
        
        self.btn_run_excel_export = tk.Button(
            parent, text="Export to single Excel file", 
            command=self.run_excel_export, 
            state=tk.DISABLED, 
            width=25, 
            bg="#2E7D32", fg="white", activebackground="#1B5E20",
            font=("Helvetica", 11, "bold"), pady=6
        )
        self.btn_run_excel_export.pack(pady=10)
        
    def handle_export_drop(self, event):
        files = self.tk.splitlist(event.data)
        self.process_export_paths(files)
        
    def select_export_files(self):
        files = filedialog.askopenfilenames(
            title="Select DAT files",
            filetypes=(("DAT Files", "*.DAT;*.dat"), ("All Files", "*.*"))
        )
        if files:
            self.process_export_paths(files)
            
    def select_export_folder(self):
        folder = filedialog.askdirectory(title="Select Folder containing DAT files")
        if folder:
            self.process_export_paths([folder])
            
    def clear_export_files(self):
        self.export_dat_files.clear()
        self.lbl_export_status.config(text="0 file(s) selected", fg="gray")
        self.btn_run_excel_export.config(state=tk.DISABLED)
        self.lbl_drop_export.config(bg="#E8EAF6")
        
    def process_export_paths(self, paths):
        found = set()
        for p in paths:
            clean = p.strip('{}')
            if os.path.isdir(clean):
                found.update(glob.glob(os.path.join(clean, "*.dat")))
                found.update(glob.glob(os.path.join(clean, "*.DAT")))
            elif os.path.isfile(clean) and clean.lower().endswith('.dat'):
                found.add(clean)
                
        for f in found:
            if f not in self.export_dat_files:
                self.export_dat_files.append(f)
                
        count = len(self.export_dat_files)
        if count > 0:
            self.lbl_export_status.config(text=f"{count} file(s) ready for Excel export", fg="#1565C0")
            self.btn_run_excel_export.config(state=tk.NORMAL)
            self.lbl_drop_export.config(bg="#E8F5E9")
        else:
            self.clear_export_files()
            
    def run_excel_export(self):
        if not self.export_dat_files:
            return
            
        all_data = []
        for f in self.export_dat_files:
            file_data = parse_dat_for_excel(f)
            all_data.extend(file_data)
            
        if not all_data:
            messagebox.showwarning("Warning", "No valid data could be extracted.")
            return
            
        try:
            df = pd.DataFrame(all_data)
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile="TF_Readings_Export.xlsx",
                title="Save Excel File As",
                filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*"))
            )
            if save_path:
                df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"Successfully exported {len(all_data)} records to:\n{save_path}")
                self.clear_export_files()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

if __name__ == "__main__":
    app = TFReadingsApp()
    app.mainloop()
