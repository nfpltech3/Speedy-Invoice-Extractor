import os
import sys
import threading
import queue
import re
from datetime import datetime
from dataclasses import dataclass
from tkinter import (
    Tk, Frame, Label, Text, filedialog, messagebox,
    ttk, END, WORD, BOTH, Y, X, BOTTOM, RIGHT, LEFT
)
import pandas as pd
import pdfplumber

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


# ==========================================
# 1. PARSER MODULE
# ==========================================
@dataclass
class InvoiceData:
    invoice_number: str = ""
    invoice_date: str = ""
    sac_code: str = ""
    shipping_bill_no: str = ""
    base_amount: float = 0.0
    cgst_amount: float = 0.0
    sgst_amount: float = 0.0
    total_amount: float = 0.0
    tds_amount: float = 0.0
    extraction_errors: list = None

    def __post_init__(self):
        if self.extraction_errors is None:
            self.extraction_errors = []

def parse_date(date_str: str) -> str:
    try:
        dt = datetime.strptime(date_str, "%d/%m/%Y")
        return dt.strftime("%d-%b-%Y")
    except ValueError:
        return date_str

def parse_amount(amt_str: str) -> float:
    if not amt_str: return 0.0
    cleaned = re.sub(r'[₹$,\s]', '', str(amt_str))
    try: return float(cleaned)
    except ValueError: return 0.0

def parse_invoice(pdf_path: str) -> InvoiceData:
    data = InvoiceData()
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            tables = page.extract_tables()
            
        inv_no_match = re.search(r'Invoice No\s*[:]\s*([^\s]+)', text, re.IGNORECASE)
        if inv_no_match:
            data.invoice_number = inv_no_match.group(1).strip()
            
        inv_date_match = re.search(r'Invoice Date\s*[:]\s*(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
        if inv_date_match:
            data.invoice_date = parse_date(inv_date_match.group(1).strip())
            
        sac_match = re.search(r'(\d+)\s+CONTAINER HANDLING SERVICE\s+([\d,]+\.?\d*)', text)
        if sac_match:
            data.sac_code = sac_match.group(1).strip()
            data.base_amount = parse_amount(sac_match.group(2))
            
        cgst_match = re.search(r'CGST\s*-\s*\d+(?:\.\d+)?%\s+([\d,]+\.?\d*)', text)
        if cgst_match:
            data.cgst_amount = parse_amount(cgst_match.group(1))
            
        sgst_match = re.search(r'SGST\s*-\s*\d+(?:\.\d+)?%\s+([\d,]+\.?\d*)', text)
        if sgst_match:
            data.sgst_amount = parse_amount(sgst_match.group(1))
            
        inv_amt_match = re.search(r'Inv Amt:\s*₹?\s*([\d,]+\.?\d*)', text)
        if inv_amt_match:
            data.total_amount = parse_amount(inv_amt_match.group(1))
            
        tds_match = re.search(r'TDS\s+([\d,]+\.?\d*)', text)
        if tds_match:
            data.tds_amount = parse_amount(tds_match.group(1))
        else:
            data.tds_amount = round(data.base_amount * 0.02, 2)
            
        # Extract Shipping Bill No from Tables
        for table in tables:
            for i, row in enumerate(table):
                # Search for Shipping Bill No header
                str_row = [str(x) for x in row]
                if 'Shipping Bill No' in str_row:
                    # It's usually in the row right below the header row
                    if i + 1 < len(table):
                        next_row = [str(x) for x in table[i+1]]
                        # Assuming the 2nd column (index 1) is Shipping Bill No as seen in extraction
                        if len(next_row) > 1 and next_row[1].isdigit():
                            data.shipping_bill_no = next_row[1].strip()
                        elif len(next_row) > 2 and next_row[2].isdigit():
                            data.shipping_bill_no = next_row[2].strip()

        # Regex fallback for Shipping Bill if table parsing failed
        if not data.shipping_bill_no:
            sb_fallback = re.search(r'Shipping Bill Details\s*\n.*?\n.*?\n1\s+(\d{5,10})', text)
            if sb_fallback:
                data.shipping_bill_no = sb_fallback.group(1)
            
        if not data.invoice_number:
            data.extraction_errors.append("Invoice No not found in PDF.")
            
    except Exception as e:
        data.extraction_errors.append(f"Error reading PDF: {str(e)}")
        
    return data

def get_job_no_from_df(shipping_bill: str, job_df: pd.DataFrame) -> str:
    if not shipping_bill or job_df is None or job_df.empty:
        return ""
    try:
        # Match "SB No." with the shipping bill. Note column name is 'SB No.' 
        # but let's be flexible and find the right col
        sb_col = None
        job_col = None
        for col in job_df.columns:
            clean_col = str(col).strip().lower()
            if 'sb no' in clean_col or 'shipping bill' in clean_col:
                sb_col = col
            if 'job no' in clean_col:
                job_col = col
                
        if not sb_col or not job_col:
            return ""

        job_df[sb_col] = job_df[sb_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        shipping_bill_clean = str(shipping_bill).strip()
        
        match = job_df[job_df[sb_col] == shipping_bill_clean][job_col]
        if not match.empty:
            return str(match.iloc[0])
    except Exception:
        return ""
    return ""

# ==========================================
# 2. GENERATOR MODULE
# ==========================================
CSV_HEADERS = [
    "Entry Date", "Posting Date", "Organization", "Organization Branch",
    "Vendor Inv No", "Vendor Inv Date", "Currency", "ExchRate", "Narration",
    "Due Date", "Charge or GL", "Charge or GL Name", "Charge or GL Amount",
    "DR or CR", "Cost Center", "Branch", " Charge Narration", "TaxGroup",
    "Tax Type", "SAC or HSN", "Taxcode1", "Taxcode1 Amt", "Taxcode2",
    "Taxcode2 Amt", "Taxcode3", "Taxcode3 Amt", "Taxcode4", "Taxcode4 Amt",
    "Avail Tax Credit", "LOB", "Ref Type", "Ref No", "Amount ", "Start Date",
    "End Date", "WH Tax Code", "WH Tax Percentage", "WH Tax Taxable",
    "WH Tax Amount", "Round Off", "CC Code"
]

def invoice_to_csv_row(inv: InvoiceData, job_no: str) -> dict:
    today = datetime.now().strftime("%d-%b-%Y")
    
    # Narration replacement
    narration_suffix = job_no if job_no else ""
    narration = f"Being Entry posted for Speedy / CFS / {narration_suffix}"
    
    # WH Tax Amount = 2% of WH Tax Taxable (base amount)
    wh_tax_amount = round(inv.base_amount * 0.02, 2)
    
    return {
        "Entry Date": today,
        "Posting Date": today,
        "Organization": "SPEEDY MULTIMODES LTD",
        "Organization Branch": "MUMBAI",
        "Vendor Inv No": inv.invoice_number,
        "Vendor Inv Date": inv.invoice_date,
        "Currency": "INR",
        "ExchRate": "1",
        "Narration": narration.strip(),
        "Due Date": "",
        "Charge or GL": "WAREHOUSE CHARGES _ GST",
        "Charge or GL Name": "WAREHOUSE CHARGES (E) (1) (PAYMENT)",
        "Charge or GL Amount": str(inv.base_amount),
        "DR or CR": "DR",
        "Cost Center": "CCL Export",
        "Branch": "HO",
        " Charge Narration": "",
        "TaxGroup": "GSTIN",
        "Tax Type": "Taxable",
        "SAC or HSN": inv.sac_code or "996711",
        "Taxcode1": "CGST",
        "Taxcode1 Amt": str(inv.cgst_amount),
        "Taxcode2": "SGST",
        "Taxcode2 Amt": str(inv.sgst_amount),
        "Taxcode3": "",
        "Taxcode3 Amt": "",
        "Taxcode4": "",
        "Taxcode4 Amt": "",
        "Avail Tax Credit": "100",
        "LOB": "CCL EXP",
        "Ref Type": "",
        "Ref No": job_no,
        "Amount ": str(inv.base_amount),
        "Start Date": "",
        "End Date": "",
        "WH Tax Code": "194C",
        "WH Tax Percentage": "2",
        "WH Tax Taxable": str(inv.base_amount),
        "WH Tax Amount": str(wh_tax_amount),
        "Round Off": "No",
        "CC Code": ""
    }

# ==========================================
# 3. GUI MODULE (Nagarkot Brand)
# ==========================================
BG_COLOR = "#F4F6F8"
CARD_BG = "#FFFFFF"
ACCENT = "#1F3F6E"
ACCENT_HOVER = "#2A528F"
TEXT_PRIMARY = "#1E1E1E"
TEXT_SECONDARY = "#6B7280"
BORDER_COLOR = "#E5E7EB"
ERROR_RED = "#D8232A"
SUCCESS_GREEN = "#1F3F6E"
LOG_BG = "#FAFBFC"
LOG_FG = "#1E1E1E"

class SpeedyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Speedy Invoice Extractor")
        try:
            self.root.state("zoomed")
        except Exception:
            self.root.attributes("-fullscreen", True)
        self.root.configure(bg=BG_COLOR)

        self.selected_files = []
        self.job_register_path = None
        self.job_df = None
        self._logo_image = None

        self.log_queue = queue.Queue()
        self._setup_styles()
        self._create_widgets()
        self._start_log_polling()

    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Card.TLabelframe", background=CARD_BG, borderwidth=1,
                         relief="solid")
        style.configure("Card.TLabelframe.Label", background=CARD_BG,
                         foreground=TEXT_PRIMARY, font=("Segoe UI", 10, "bold"))
        style.configure("Modern.TButton", font=("Segoe UI", 9), padding=(14, 6),
                         background=CARD_BG, borderwidth=1, relief="solid")
        style.map("Modern.TButton",
                   background=[("active", "#F5F5F5"), ("pressed", "#EEEEEE")])
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"),
                         padding=(20, 8), foreground="#FFFFFF", background=ACCENT,
                         borderwidth=0)
        style.map("Accent.TButton",
                   background=[("active", ACCENT_HOVER), ("pressed", ACCENT_HOVER),
                               ("disabled", "#90CAF9")],
                   foreground=[("disabled", "#FFFFFF")])

    def _create_widgets(self):
        main_frame = Frame(self.root, bg=BG_COLOR)
        main_frame.pack(fill=BOTH, expand=True)

        # ── HEADER (explicit height so title never clips) ──
        header_frame = Frame(main_frame, bg=CARD_BG, height=80)
        header_frame.pack(fill=X)
        header_frame.pack_propagate(False)
        Frame(main_frame, bg=BORDER_COLOR, height=1).pack(fill=X)

        # Logo — top-left, fixed 20px height
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(base_path, "logo.png")

        if HAS_PIL and os.path.isfile(logo_path):
            try:
                img = Image.open(logo_path)
                h = 20
                w = int(img.width * h / img.height)
                img = img.resize((w, h), Image.LANCZOS)
                self._logo_image = ImageTk.PhotoImage(img)
                Label(header_frame, image=self._logo_image, bg=CARD_BG).place(
                    x=24, rely=0.5, anchor="w")
            except Exception:
                Label(header_frame, text="NAGARKOT", font=("Segoe UI", 12, "bold"),
                      fg=ACCENT, bg=CARD_BG).place(x=24, rely=0.5, anchor="w")
        else:
            Label(header_frame, text="NAGARKOT", font=("Segoe UI", 12, "bold"),
                  fg=ACCENT, bg=CARD_BG).place(x=24, rely=0.5, anchor="w")

        # Title — true center
        Label(header_frame, text="Speedy Invoice Extractor",
              font=("Segoe UI", 16, "bold"), bg=CARD_BG,
              fg=TEXT_PRIMARY).place(relx=0.5, rely=0.35, anchor="center")
        Label(header_frame,
              text="Convert Speedy CFS invoices to Logisys Purchase CSV",
              font=("Segoe UI", 9), bg=CARD_BG,
              fg=TEXT_SECONDARY).place(relx=0.5, rely=0.7, anchor="center")

        # ── FOOTER (pack before body so it stays at bottom) ──
        footer_border = Frame(main_frame, bg=BORDER_COLOR, height=1)
        footer_border.pack(fill=X, side=BOTTOM)
        footer_frame = Frame(main_frame, bg=CARD_BG, padx=24, pady=10)
        footer_frame.pack(fill=X, side=BOTTOM)
        Label(footer_frame, text="Nagarkot Forwarders Pvt. Ltd. \u00A9",
              fg=TEXT_SECONDARY, bg=CARD_BG,
              font=("Segoe UI", 8)).pack(side=LEFT)
        ttk.Button(footer_frame, text="Exit", command=self.root.destroy,
                   style="Modern.TButton").pack(side=RIGHT)

        # ── BODY ──
        body = Frame(main_frame, bg=BG_COLOR, padx=40, pady=24)
        body.pack(fill=BOTH, expand=True)

        # Input Files Card
        file_card = ttk.LabelFrame(body, text="  Input Files  ",
                                    style="Card.TLabelframe", padding=20)
        file_card.pack(fill=X, pady=(0, 16))
        file_inner = Frame(file_card, bg=CARD_BG)
        file_inner.pack(fill=X)

        # Row 1 — Job Register (required, listed first)
        row1 = Frame(file_inner, bg=CARD_BG)
        row1.pack(fill=X, pady=(0, 8))
        ttk.Button(row1, text="Select Export Job Register",
                   command=self._select_excel,
                   style="Modern.TButton").pack(side=LEFT)
        self.excel_label = Label(row1, text="Not selected",
                                 fg=ERROR_RED, bg=CARD_BG,
                                 font=("Segoe UI", 9))
        self.excel_label.pack(side=LEFT, padx=(12, 0))

        # Row 2 — PDF Invoices
        row2 = Frame(file_inner, bg=CARD_BG)
        row2.pack(fill=X)
        ttk.Button(row2, text="Select PDF Invoices",
                   command=self._select_pdfs,
                   style="Modern.TButton").pack(side=LEFT)
        self.file_label = Label(row2, text="No files selected",
                                fg=TEXT_SECONDARY, bg=CARD_BG,
                                font=("Segoe UI", 9))
        self.file_label.pack(side=LEFT, padx=(12, 0))

        # Action Row
        action_frame = Frame(body, bg=BG_COLOR)
        action_frame.pack(fill=X, pady=(0, 16))
        self.process_btn = ttk.Button(action_frame,
                                       text="\u25B6  Process & Generate CSV",
                                       command=self._process,
                                       style="Accent.TButton")
        self.process_btn.pack(side=LEFT)
        self.status_lbl = Label(action_frame, text="Ready",
                                fg=TEXT_SECONDARY, bg=BG_COLOR,
                                font=("Segoe UI", 9))
        self.status_lbl.pack(side=LEFT, padx=20)

        # Processing Log Card
        log_card = ttk.LabelFrame(body, text="  Processing Log  ",
                                   style="Card.TLabelframe", padding=12)
        log_card.pack(fill=BOTH, expand=True)
        log_inner = Frame(log_card, bg=CARD_BG)
        log_inner.pack(fill=BOTH, expand=True)

        from tkinter import Scrollbar, VERTICAL
        self.log_txt = Text(log_inner, wrap=WORD, state="disabled",
                            bg=LOG_BG, fg=LOG_FG, font=("Consolas", 9),
                            relief="flat", padx=10, pady=8)
        log_scroll = Scrollbar(log_inner, orient=VERTICAL,
                               command=self.log_txt.yview)
        self.log_txt.configure(yscrollcommand=log_scroll.set)
        self.log_txt.pack(side=LEFT, fill=BOTH, expand=True)
        log_scroll.pack(side=RIGHT, fill=Y)

    # ── Helpers ──
    def _log(self, msg):
        self.log_queue.put(f"{msg}\n")

    def _start_log_polling(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_txt.config(state="normal")
                self.log_txt.insert(END, msg)
                self.log_txt.see(END)
                self.log_txt.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self._start_log_polling)

    # ── File Selection ──
    def _select_pdfs(self):
        files = filedialog.askopenfilenames(
            title="Select Speedy Invoice PDFs",
            filetypes=[("PDF files", "*.pdf")])
        if files:
            self.selected_files = list(files)
            self.file_label.config(
                text=f"{len(self.selected_files)} PDF(s) selected",
                fg=TEXT_PRIMARY)

    def _select_excel(self):
        file = filedialog.askopenfilename(
            title="Select Export Job Register",
            filetypes=[("Excel files", "*.xlsx")])
        if file:
            self.job_register_path = file
            try:
                self.job_df = pd.read_excel(self.job_register_path)
                self.excel_label.config(
                    text=f"\u2713  {os.path.basename(file)}",
                    fg=TEXT_PRIMARY)
            except Exception as e:
                self.excel_label.config(
                    text=f"Failed to load: {e}", fg=ERROR_RED)
                self.job_df = None

    # ── Processing ──
    def _process(self):
        if not self.job_register_path or self.job_df is None:
            messagebox.showerror("Missing Input",
                                 "Please select the Export Job Register first.")
            return
        if not self.selected_files:
            messagebox.showerror("Missing Input",
                                 "Please select at least one PDF invoice.")
            return

        # Clear previous log
        self.log_txt.config(state="normal")
        self.log_txt.delete("1.0", END)
        self.log_txt.config(state="disabled")

        self.process_btn.state(['disabled'])
        self.status_lbl.config(text="Processing...", fg=ACCENT)
        threading.Thread(target=self._process_thread, daemon=True).start()

    def _process_thread(self):
        import csv
        try:
            total = len(self.selected_files)
            self._log(f"Processing {total} invoice(s)...\n")

            base_dir = os.path.dirname(self.selected_files[0])
            out_dir = os.path.join(base_dir, "CSV Output")
            os.makedirs(out_dir, exist_ok=True)
            parsed_rows = []
            ok_count = 0
            fail_count = 0

            for i, pdf in enumerate(self.selected_files, 1):
                fname = os.path.basename(pdf)
                inv = parse_invoice(pdf)

                if not inv.invoice_number:
                    self._log(f"[{i}/{total}] {fname}  \u2717 FAILED — "
                              f"{', '.join(inv.extraction_errors)}")
                    fail_count += 1
                    continue

                # Job No lookup
                job_no = ""
                sb_info = ""
                if inv.shipping_bill_no:
                    job_no = get_job_no_from_df(inv.shipping_bill_no, self.job_df)
                    sb_info = f"SB:{inv.shipping_bill_no}"
                    if job_no:
                        sb_info += f" \u2192 {job_no}"
                    else:
                        sb_info += " \u2192 No match"

                parsed_rows.append(invoice_to_csv_row(inv, job_no))
                ok_count += 1
                self._log(f"[{i}/{total}] {inv.invoice_number}  |  "
                          f"\u20B9{inv.total_amount:,.2f}  |  {sb_info}")

            self._log("")  # blank line

            if parsed_rows:
                timestamp = datetime.now().strftime("%d%b%Y_%H%M").upper()
                filepath = os.path.join(
                    out_dir, f"Speedy_Purchases_{timestamp}.csv")
                with open(filepath, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
                    writer.writeheader()
                    writer.writerows(parsed_rows)
                self._log(f"\u2713 CSV saved: {os.path.basename(filepath)}")
                self._log(f"  {ok_count} OK, {fail_count} failed  |  {out_dir}")
                self.root.after(0, lambda fp=filepath: messagebox.showinfo(
                    "Success", f"CSV saved:\n{fp}"))
            else:
                self._log("No invoices were successfully parsed.")

        except Exception as e:
            self._log(f"\u2717 Error: {e}")
        finally:
            self.root.after(0, self._process_done)

    def _process_done(self):
        self.process_btn.state(['!disabled'])
        self.status_lbl.config(text="Complete", fg=SUCCESS_GREEN)


if __name__ == "__main__":
    app_root = Tk()
    SpeedyApp(app_root)
    app_root.mainloop()

