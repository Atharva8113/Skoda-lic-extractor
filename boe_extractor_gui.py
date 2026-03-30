import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import fitz # PyMuPDF
import re
import pandas as pd
import os
import threading
from PIL import Image, ImageTk
import io

# --- Nagarkot Brand Constants ---
PRIMARY_BLUE = "#1F3F6E"
ACCENT_RED = "#D8232A"
LIGHT_BG = "#F4F6F8"
PANEL_WHITE = "#FFFFFF"
TEXT_DARK = "#1E1E1E"
MUTED_GRAY = "#6B7280"
HOVER_BLUE = "#2A528F"
BORDER_GRAY = "#E5E7EB"

import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class BOEExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Skoda License Extractor")
        self.root.state('zoomed') # Full screen
        self.root.configure(bg=LIGHT_BG)
        
        self.selected_files = []
        self.extracted_data = []
        
        self.setup_ui()
        
    def setup_ui(self):
        # --- Root Config ---
        self.root.state('zoomed')
        self.root.configure(bg=LIGHT_BG)

        # --- Header ---
        self.header = tk.Frame(self.root, bg=PANEL_WHITE, height=75) # Balanced Brand Height
        self.header.pack(side="top", fill="x")
        self.header.pack_propagate(False)

        # Logo: Top-Left, Fixed Height = 20 Units (~30px visual height)
        self.logo_container = tk.Frame(self.header, bg=PANEL_WHITE)
        self.logo_container.pack(side="left", padx=(30, 0))
        
        try:
            logo_path = resource_path("Nagarkot Logo.png")
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                # Workflow Rule 2.1: Fixed height = 20 units. Scaling exactly to 30px height.
                h_target = 30 
                w_target = int(img.width * (h_target / img.height))
                img = img.resize((w_target, h_target), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
                self.logo_label = tk.Label(self.logo_container, image=self.logo_img, bg=PANEL_WHITE)
                self.logo_label.pack(pady=22) # Centered in header vertically
            else:
                self.logo_label = tk.Label(self.logo_container, text="NAGARKOT", fg=PRIMARY_BLUE, 
                                         font=("Segoe UI", 16, "bold"), bg=PANEL_WHITE)
                self.logo_label.pack(pady=20)
        except:
            pass

        # Title: Visual Center of Window (Absolute Positioning)
        self.title_label = tk.Label(self.header, text="BOE LICENSE EXTRACTOR", fg=PRIMARY_BLUE, 
                                  font=("Segoe UI", 24, "bold"), bg=PANEL_WHITE)
        self.title_label.place(relx=0.5, rely=0.5, anchor="center")

        # --- Body ---
        self.body = tk.Frame(self.root, bg=LIGHT_BG)
        self.body.pack(fill="both", expand=True, padx=40, pady=(20, 10))

        # Control Panel
        self.control_frame = tk.Frame(self.body, bg=LIGHT_BG)
        self.control_frame.pack(fill="x", pady=(0, 25))

        self.upload_btn = tk.Button(self.control_frame, text="Upload BOE PDFs", command=self.select_files,
                                  bg=PRIMARY_BLUE, fg="white", font=("Segoe UI", 11, "bold"),
                                  padx=30, pady=10, relief="flat", cursor="hand2")
        self.upload_btn.pack(side="left")
        self.upload_btn.bind("<Enter>", lambda e: self.upload_btn.configure(bg=HOVER_BLUE))
        self.upload_btn.bind("<Leave>", lambda e: self.upload_btn.configure(bg=PRIMARY_BLUE))

        self.clear_btn = tk.Button(self.control_frame, text="Clear All", command=self.clear_data,
                                 bg="white", fg=PRIMARY_BLUE, font=("Segoe UI", 11),
                                 padx=20, pady=10, relief="flat", highlightbackground=PRIMARY_BLUE, 
                                 highlightthickness=1, cursor="hand2")
        self.clear_btn.pack(side="left", padx=25)

        self.export_btn = tk.Button(self.control_frame, text="Generate Excel", command=self.generate_excel,
                                  bg=ACCENT_RED, fg="white", font=("Segoe UI", 11, "bold"),
                                  padx=30, pady=10, relief="flat", state="disabled", cursor="hand2")
        self.export_btn.pack(side="right")

        # Progress
        self.progress_var = tk.DoubleVar()
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Nagarkot.Horizontal.TProgressbar", thickness=12, troughcolor=PANEL_WHITE, background=PRIMARY_BLUE)
        self.progress = ttk.Progressbar(self.body, variable=self.progress_var, maximum=100, style="Nagarkot.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=15)

        # Table Area
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=32, background="white", fieldbackground="white")
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"), foreground=PRIMARY_BLUE, background="#F8FAFC")
        
        self.tree_frame = tk.Frame(self.body, bg=PANEL_WHITE, highlightbackground=BORDER_GRAY, highlightthickness=1)
        self.tree_frame.pack(fill="both", expand=True)

        columns = ("File Name", "BE Number", "BE Date", "License No", "Debit Duty", "Status")
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show="headings", selectmode="browse")

        for col in columns:
            self.tree.heading(col, text=col.upper())
            self.tree.column(col, width=150, anchor="center")

        self.tree.column("File Name", width=350, anchor="w")
        self.tree.column("Status", width=120)

        scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # --- Footer ---
        self.footer = tk.Frame(self.root, bg=LIGHT_BG, height=40)
        self.footer.pack(side="bottom", fill="x")
        self.footer_label = tk.Label(self.footer, text="Nagarkot Forwarders Pvt. Ltd. ©", 
                                   fg=MUTED_GRAY, font=("Segoe UI", 9), bg=LIGHT_BG)
        self.footer_label.pack(side="left", padx=45, pady=8)

    def log(self, file_name, be_no="", be_date="", lic_no="", duty="", status="Processing"):
        item = self.tree.insert("", "end", values=(file_name, be_no, be_date, lic_no, duty, status))
        self.tree.see(item)
        self.root.update_idletasks()
        return item

    def update_log(self, item, values):
        self.tree.item(item, values=values)
        self.root.update_idletasks()

    def select_files(self):
        files = filedialog.askopenfilenames(title="Select BOE PDFs", filetypes=[("PDF files", "*.pdf")])
        if files:
            self.selected_files = list(files)
            self.upload_btn.configure(state="disabled")
            threading.Thread(target=self.process_files, daemon=True).start()

    def clear_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.extracted_data = []
        self.selected_files = []
        self.export_btn.configure(state="disabled")
        self.upload_btn.configure(state="normal")
        self.progress_var.set(0)

    def process_files(self):
        total = len(self.selected_files)
        for idx, file_path in enumerate(self.selected_files):
            file_name = os.path.basename(file_path)
            log_item = self.log(file_name)
            
            try:
                data, found_table = self.extract_data(file_path)
                if not found_table:
                    self.update_log(log_item, (file_name, "-", "-", "-", "-", "No License Table Found"))
                elif not data:
                    self.update_log(log_item, (file_name, "-", "-", "-", "-", "No Data Extracted"))
                else:
                    self.extracted_data.extend(data)
                    first = data[0]
                    self.update_log(log_item, (file_name, first['BE Number'], first['BE Date'], first['Licence No'], first['Debit Duty'], f"Extracted {len(data)} items"))
                    
                    if len(data) > 1:
                        for d in data[1:]:
                            self.log(file_name, d['BE Number'], d['BE Date'], d['Licence No'], d['Debit Duty'], "Success")
            except Exception as e:
                self.update_log(log_item, (file_name, "Error", "", "", "", str(e)))
            
            self.progress_var.set(((idx + 1) / total) * 100)
            
        if self.extracted_data:
            self.export_btn.configure(state="normal")
        
        self.upload_btn.configure(state="normal")
        messagebox.showinfo("Done", "Processing complete!")

    def extract_data(self, pdf_path):
        doc = fitz.open(pdf_path)
        be_no = ""
        be_date = ""
        
        # 1. Header Extraction - Zero-Assumption Brute Force
        for i in range(min(2, len(doc))): # Header is on page 1 or 2
            page = doc[i]
            text = page.get_text()
            
            # Find all 7-8 digit numbers (BE No) and DD/MM/YYYY (BE Date) in the page
            all_digits = re.findall(r'\b\d{7,8}\b', text)
            all_dates = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', text)
            
            # Prioritize matches near "BE" keyword
            if not be_no:
                # Search specifically for BE No pattern
                m = re.search(r'BE\s*No\.?\s*[|: ]*\s*(\d{7,8})', text, re.I)
                if m: 
                    be_no = m.group(1)
                elif all_digits:
                    # Fallback to first 7-8 digit number found (usually at top)
                    be_no = all_digits[0]
            
            if not be_date:
                # Search specifically for BE Date pattern
                m = re.search(r'BE\s*Date\s*[|: ]*\s*(\d{2}/\d{2}/\d{4})', text, re.I)
                if m: 
                    be_date = m.group(1)
                elif all_dates:
                    # Fallback to first date found (usually at top)
                    be_date = all_dates[0]

            if be_no and be_date: break

        # 2. Section F Extraction
        records = []
        found_section_ever = False
        in_section = False
        
        lic_no_x = -1
        debit_duty_x = -1
        section_y = -1 
        
        for page in doc:
            title_rects = page.search_for("F. LICENCE DETAILS")
            if title_rects:
                found_section_ever = True
                in_section = True
                section_y = title_rects[0].y1
            
            if in_section:
                words = page.get_text("words")
                words.sort(key=lambda w: (w[1], w[0]))
                
                rows = []
                current_row = []
                last_y = -1
                for w in words:
                    if section_y != -1 and w[3] < section_y: continue
                    if last_y == -1 or abs(w[1] - last_y) <= 3:
                        current_row.append(w)
                    else:
                        rows.append(current_row)
                        current_row = [w]
                    last_y = w[1]
                if current_row: rows.append(current_row)
                
                for r in rows:
                    r.sort(key=lambda w: w[0])
                    line_text = " ".join([w[4] for w in r])
                    
                    if any(x in line_text for x in ["G. RE-EXPORT DETAILS", "H. IGST", "PART - V", "I. IGST", "G.RE-EXPORT DETAILS"]):
                        in_section = False
                        break
                    
                    header_check = line_text.replace(" ", "").upper()
                    if "4.LICNO" in header_check or ("4.LIC" in header_check and "11.DEBIT" in header_check):
                        for w in r:
                            w_text = w[4].replace(" ", "").upper()
                            if "4.LIC" in w_text: lic_no_x = w[0]
                            if "11.DEBIT" in w_text or "11.DEBT" in w_text or "11.DEBITDUTY" in w_text: debit_duty_x = w[0]
                        continue

                    if lic_no_x != -1:
                        lic_val = ""
                        duty_val = ""
                        tol = 35 
                        for w in r:
                            if abs(w[0] - lic_no_x) < tol:
                                if re.match(r'^([A-Z]\d{9}|\d{10})$', w[4].strip()):
                                    lic_val = w[4].strip()
                            if debit_duty_x != -1 and abs(w[0] - debit_duty_x) < tol:
                                if re.match(r'^\d+\.?\d*$', w[4].strip()):
                                    duty_val = w[4].strip()
                        
                        if lic_val:
                            if not duty_val:
                                ns = re.findall(r'\d+\.\d{2}', line_text)
                                if ns: duty_val = ns[-1]
                            records.append({
                                "BE Number": be_no, "BE Date": be_date,
                                "Licence No": lic_val, "Debit Duty": duty_val if duty_val else "0"
                            })
                    else:
                        if "INVSNO" in line_text or "ITMSNO" in line_text: continue
                        lic_match = re.search(r'([A-Z]\d{9}|\d{10})', line_text)
                        if lic_match:
                            lic_no = lic_match.group(1)
                            if lic_no.startswith("0000"): continue 
                            ns = re.findall(r'\d+\.\d{2}', line_text)
                            records.append({
                                "BE Number": be_no, "BE Date": be_date,
                                "Licence No": lic_no, "Debit Duty": ns[-1] if ns else "0"
                            })
                section_y = -1 
                if not in_section: break
        return records, found_section_ever

    def generate_excel(self):
        if not self.extracted_data: return
        sp = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if sp:
            try:
                pd.DataFrame(self.extracted_data).to_excel(sp, index=False)
                messagebox.showinfo("Success", f"Report saved to:\n{sp}")
            except Exception as e: messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = BOEExtractorApp(root)
    root.mainloop()
