#!/usr/bin/env python3
"""
PCI Monthly Summary Generator - Simple Edition
Professional billing tool for Extended Care
"""

__version__ = "3.0"

import os
import sys
import json
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Tuple
import tkinter as tk
from tkinter import filedialog, messagebox

# Third-party imports
try:
    import customtkinter as ctk
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError as e:
    print(f"Missing required package: {e}")
    print("Run: pip install customtkinter pandas openpyxl")
    sys.exit(1)


# =============================================================================
# CONFIGURATION
# =============================================================================

class Config:
    """App settings - edit these values as needed."""
    HOURLY_RATE = 12.0
    STAFF_DISCOUNT = 0.50  # 50% off for staff children
    CATALOG_ITEM_FK = "2504"
    DESCRIPTION = "EC After/Before School Care Drop-In"
    
    # Keywords in Notes column that will skip a row
    IGNORE_KEYWORDS = ["skip", "ignore", "no charge", "waived", "monthly"]
    
    # Preferences file location
    PREFS_FILE = Path.home() / ".pci_billing_prefs.json"


# =============================================================================
# ERROR HANDLING
# =============================================================================

class BillingError(Exception):
    """Custom exception for billing-related errors."""
    pass


def get_user_friendly_error(error: Exception) -> str:
    """Convert technical errors to user-friendly messages."""
    error_str = str(error).lower()
    error_type = type(error).__name__
    
    if "no such file" in error_str or "filenotfound" in error_type.lower():
        return "File not found. Please check the file path and try again."
    
    if "permission denied" in error_str:
        return "File is locked. Please close it in Excel and try again."
    
    if "sheet" in error_str and "not found" in error_str:
        return f"Sheet not found. Please check the sheet name is correct."
    
    if error_type == "KeyError":
        col = str(error).strip("'\"")
        return f"Missing required column: '{col}'. Please check your file format."
    
    if "time spent" in error_str or "student" in error_str:
        return "Missing required columns. File must have 'Student' and 'Time Spent' columns."
    
    if "empty" in error_str:
        return "No data found in the file or sheet."
    
    # Generic fallback
    return f"Error: {str(error)}"


def log_error(error: Exception, context: str = ""):
    """Log error details for debugging."""
    log_path = Path.home() / "pci_billing_errors.log"
    try:
        with open(log_path, "a") as f:
            f.write(f"\n{'='*60}\n")
            f.write(f"Time: {datetime.now()}\n")
            f.write(f"Context: {context}\n")
            f.write(f"Error: {type(error).__name__}: {error}\n")
            f.write(f"Traceback:\n{traceback.format_exc()}\n")
    except:
        pass  # Don't fail if logging fails


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def normalize_id(val) -> Optional[str]:
    """Convert ID to string, handling floats like 109569.0 -> '109569'."""
    if pd.isna(val):
        return None
    if isinstance(val, float):
        return str(int(val))
    return str(val).strip()


def parse_currency(value) -> float:
    """Parse currency like $120.00 or (50.00) to float."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace('$', '').replace(',', '')
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    try:
        return float(s)
    except:
        return 0.0


def should_ignore_row(notes_value) -> Tuple[bool, str]:
    """Check if a row should be ignored based on Notes column."""
    if pd.isna(notes_value):
        return False, ""
    
    notes = str(notes_value).lower().strip()
    for keyword in Config.IGNORE_KEYWORDS:
        if keyword.lower() in notes:
            return True, keyword
    return False, ""


def load_prefs() -> dict:
    """Load saved preferences."""
    try:
        if Config.PREFS_FILE.exists():
            return json.loads(Config.PREFS_FILE.read_text())
    except:
        pass
    return {}


def save_prefs(prefs: dict):
    """Save preferences for next session."""
    try:
        Config.PREFS_FILE.write_text(json.dumps(prefs, indent=2))
    except:
        pass


# =============================================================================
# CORE BILLING LOGIC
# =============================================================================

def process_attendance_file(filepath: str, school_year: str, item_date: str,
                            staff_children: List[str] = None) -> Tuple[pd.DataFrame, pd.DataFrame, dict, pd.DataFrame]:
    """
    Process an attendance file and generate billing data.
    
    Returns: (summary_df, bookbill_df, stats_dict, skipped_df)
    """
    staff_children = staff_children or []
    staff_set = {s.lower().strip() for s in staff_children}
    
    # Load file
    try:
        if filepath.lower().endswith('.csv'):
            df = pd.read_csv(filepath)
        else:
            df = pd.read_excel(filepath)
    except Exception as e:
        log_error(e, f"Loading file: {filepath}")
        raise BillingError(f"Could not read file: {get_user_friendly_error(e)}")
    
    if df.empty:
        raise BillingError("The file appears to be empty.")
    
    # Find required columns
    student_col = None
    time_col = None
    person_id_col = None
    notes_col = None
    
    for col in df.columns:
        col_lower = str(col).lower()
        if 'student' in col_lower or 'person_name' in col_lower:
            student_col = col
        if 'time' in col_lower and 'spent' in col_lower:
            time_col = col
        if 'person_id' in col_lower:
            person_id_col = col
        if 'note' in col_lower:
            notes_col = col
    
    if not student_col:
        raise BillingError("Missing 'Student' or 'Person Name' column.")
    if not time_col:
        raise BillingError("Missing 'Time Spent' column.")
    
    # Process rows
    summary = {}
    skipped_rows = []
    
    for idx, row in df.iterrows():
        student = str(row[student_col]).strip()
        if not student or student.lower() == 'nan':
            continue
        
        # Check if should ignore
        if notes_col:
            should_skip, keyword = should_ignore_row(row.get(notes_col))
            if should_skip:
                skipped_rows.append({
                    'Row': idx + 2,  # Excel row number
                    'Student': student,
                    'Reason': f"Ignored: '{keyword}' in Notes"
                })
                continue
        
        # Get values
        try:
            hours = float(row[time_col]) if pd.notna(row[time_col]) else 0
        except (ValueError, TypeError):
            hours = 0
        
        person_id = normalize_id(row.get(person_id_col)) if person_id_col else None
        
        # Aggregate by student
        if student not in summary:
            summary[student] = {
                'hours': 0,
                'person_id': person_id,
                'is_staff': student.lower().strip() in staff_set
            }
        summary[student]['hours'] += hours
    
    if not summary:
        raise BillingError("No billable records found after filtering.")
    
    # Calculate charges
    results = []
    for student, data in summary.items():
        hours = data['hours']
        rate = Config.HOURLY_RATE
        if data['is_staff']:
            rate = rate * (1 - Config.STAFF_DISCOUNT)
        charge = round(hours * rate, 2)
        
        results.append({
            'Student': student,
            'Hours': round(hours, 2),
            'Rate': rate,
            'Charge': charge,
            'person_id': data['person_id'],
            'Staff': 'Yes' if data['is_staff'] else ''
        })
    
    summary_df = pd.DataFrame(results)
    summary_df = summary_df.sort_values('Student')
    
    # Build book bill
    bookbill_rows = []
    for _, row in summary_df.iterrows():
        if row['Charge'] > 0:
            bookbill_rows.append({
                'person_id': row['person_id'] or '',
                'person_name': row['Student'],
                'school_year': school_year,
                'item_date': item_date,
                'catalog_item_fk': Config.CATALOG_ITEM_FK,
                'description': Config.DESCRIPTION,
                'quantity': 1,
                'unit_price': row['Charge'],
                'item_amount': row['Charge']
            })
    
    bookbill_df = pd.DataFrame(bookbill_rows)
    skipped_df = pd.DataFrame(skipped_rows) if skipped_rows else pd.DataFrame()
    
    # Stats
    stats = {
        'total_rows': len(df),
        'students': len(summary_df),
        'total_hours': round(summary_df['Hours'].sum(), 2),
        'total_charges': round(summary_df['Charge'].sum(), 2),
        'staff_count': len([r for r in results if r['Staff']]),
        'skipped_count': len(skipped_rows)
    }
    
    return summary_df, bookbill_df, stats, skipped_df


def combine_files(folder_path: str, sheet_name: str, school_year: str,
                  item_date: str) -> Tuple[pd.DataFrame, int, List[Tuple[str, str]]]:
    """
    Combine data from multiple Excel files in a folder.
    
    Returns: (combined_df, files_processed, files_skipped)
    """
    if not os.path.isdir(folder_path):
        raise BillingError(f"Folder not found: {folder_path}")
    
    files = [f for f in os.listdir(folder_path)
             if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    
    if not files:
        raise BillingError("No Excel files found in the selected folder.")
    
    all_data = []
    processed = 0
    skipped = []
    
    for filename in files:
        filepath = os.path.join(folder_path, filename)
        try:
            df = pd.read_excel(filepath, sheet_name=sheet_name)
            
            # Filter blank rows (columns A and B must have values)
            if len(df.columns) >= 2:
                col_a, col_b = df.columns[0], df.columns[1]
                mask = df[col_a].notna() & df[col_b].notna()
                mask &= df[col_a].astype(str).str.strip() != ''
                mask &= df[col_b].astype(str).str.strip() != ''
                mask &= df[col_a].astype(str).str.strip() != '0'
                df = df[mask]
            
            if df.empty:
                skipped.append((filename, "No valid data rows"))
                continue
            
            # Update billing fields
            if 'school_year' in df.columns:
                df['school_year'] = school_year
            if 'item_date' in df.columns:
                df['item_date'] = item_date
            
            all_data.append(df)
            processed += 1
            
        except ValueError as e:
            if "sheet" in str(e).lower():
                skipped.append((filename, f"Sheet '{sheet_name}' not found"))
            else:
                skipped.append((filename, str(e)))
        except Exception as e:
            log_error(e, f"Processing file: {filename}")
            skipped.append((filename, get_user_friendly_error(e)))
    
    if not all_data:
        raise BillingError("Could not extract data from any files.")
    
    combined = pd.concat(all_data, ignore_index=True)
    return combined, processed, skipped


# =============================================================================
# UI COMPONENTS
# =============================================================================

class StatusBar(ctk.CTkFrame):
    """Status bar with icon indicator."""
    
    def __init__(self, parent):
        super().__init__(parent, height=30, fg_color="transparent")
        
        self.indicator = ctk.CTkLabel(self, text="●", width=20, text_color="gray")
        self.indicator.pack(side="left", padx=(10, 5))
        
        self.label = ctk.CTkLabel(self, text="Ready", font=ctk.CTkFont(size=12))
        self.label.pack(side="left")
    
    def set(self, message: str, status: str = "info"):
        """Update status. status: 'info', 'success', 'error', 'working'"""
        colors = {
            "info": "gray",
            "success": "#22c55e",
            "error": "#ef4444",
            "working": "#3b82f6"
        }
        self.indicator.configure(text_color=colors.get(status, "gray"))
        self.label.configure(text=message)
        self.update()


class Toast:
    """Non-blocking notification message."""
    
    @staticmethod
    def show(parent, message: str, msg_type: str = "info", duration: int = 3000):
        colors = {
            "success": "#22c55e",
            "error": "#ef4444",
            "warning": "#f59e0b",
            "info": "#3b82f6"
        }
        
        toast = ctk.CTkFrame(parent, fg_color=colors.get(msg_type, "#3b82f6"), corner_radius=6)
        toast.place(relx=0.5, rely=0.92, anchor="s")
        
        ctk.CTkLabel(
            toast, text=f"  {message}  ",
            font=ctk.CTkFont(size=13), text_color="white"
        ).pack(padx=15, pady=10)
        
        parent.after(duration, toast.destroy)


class FilePickerRow(ctk.CTkFrame):
    """File/folder selection row."""
    
    def __init__(self, parent, label: str, placeholder: str = "",
                 filetypes=None, is_folder: bool = False):
        super().__init__(parent, fg_color="transparent")
        
        self.is_folder = is_folder
        self.filetypes = filetypes or [("All files", "*.*")]
        
        # Label
        ctk.CTkLabel(
            self, text=label,
            font=ctk.CTkFont(size=13, weight="bold"),
            width=130, anchor="w"
        ).pack(side="left", padx=(0, 10))
        
        # Entry
        self.var = ctk.StringVar()
        self.entry = ctk.CTkEntry(
            self, textvariable=self.var,
            placeholder_text=placeholder,
            height=36
        )
        self.entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        
        # Browse button
        ctk.CTkButton(
            self, text="Browse", width=80, height=36,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray40"),
            command=self._browse
        ).pack(side="right")
    
    def _browse(self):
        if self.is_folder:
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename(filetypes=self.filetypes)
        if path:
            self.var.set(path)
    
    def get(self) -> str:
        return self.var.get().strip()
    
    def set(self, value: str):
        self.var.set(value)


class DateRow(ctk.CTkFrame):
    """Date entry with quick buttons."""
    
    def __init__(self, parent, label: str):
        super().__init__(parent, fg_color="transparent")
        
        # Label
        ctk.CTkLabel(
            self, text=label,
            font=ctk.CTkFont(size=13, weight="bold"),
            width=130, anchor="w"
        ).pack(side="left", padx=(0, 10))
        
        # Entry
        self.var = ctk.StringVar()
        self.entry = ctk.CTkEntry(
            self, textvariable=self.var,
            placeholder_text="MM/DD/YYYY",
            width=120, height=36
        )
        self.entry.pack(side="left", padx=(0, 8))
        
        # Quick buttons
        ctk.CTkButton(
            self, text="Today", width=60, height=36,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray40"),
            command=self._today
        ).pack(side="left", padx=(0, 4))
        
        ctk.CTkButton(
            self, text="1st", width=40, height=36,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray40"),
            command=self._first
        ).pack(side="left")
    
    def _today(self):
        self.var.set(datetime.now().strftime("%m/%d/%Y"))
    
    def _first(self):
        now = datetime.now()
        self.var.set(f"{now.month:02d}/01/{now.year}")
    
    def get(self) -> str:
        return self.var.get().strip()
    
    def set(self, value: str):
        self.var.set(value)


class InputRow(ctk.CTkFrame):
    """Simple labeled input row."""
    
    def __init__(self, parent, label: str, placeholder: str = "", width: int = 120):
        super().__init__(parent, fg_color="transparent")
        
        ctk.CTkLabel(
            self, text=label,
            font=ctk.CTkFont(size=13, weight="bold"),
            width=130, anchor="w"
        ).pack(side="left", padx=(0, 10))
        
        self.var = ctk.StringVar()
        self.entry = ctk.CTkEntry(
            self, textvariable=self.var,
            placeholder_text=placeholder,
            width=width, height=36
        )
        self.entry.pack(side="left")
    
    def get(self) -> str:
        return self.var.get().strip()
    
    def set(self, value: str):
        self.var.set(value)
    
    def insert(self, idx, value: str):
        self.entry.insert(idx, value)


class ResultsBox(ctk.CTkFrame):
    """Results display area."""
    
    def __init__(self, parent):
        super().__init__(parent, fg_color=("gray90", "gray20"), corner_radius=8)
        
        self.label = ctk.CTkLabel(
            self, text="",
            font=ctk.CTkFont(size=12),
            justify="left",
            anchor="nw"
        )
        self.label.pack(fill="both", expand=True, padx=15, pady=15)
    
    def set(self, text: str):
        self.label.configure(text=text)
    
    def clear(self):
        self.label.configure(text="")


# =============================================================================
# MAIN APPLICATION
# =============================================================================

class PCIBillingApp(ctk.CTk):
    """Professional billing tool for Extended Care."""
    
    def __init__(self):
        super().__init__()
        
        self.title("PCI Billing Tool")
        self.geometry("750x650")
        self.minsize(650, 550)
        
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        
        self.prefs = load_prefs()
        self._build_ui()
    
    def _build_ui(self):
        """Build the user interface."""
        
        # Header
        header = ctk.CTkFrame(self, fg_color=("gray85", "gray20"), corner_radius=0)
        header.pack(fill="x")
        
        header_inner = ctk.CTkFrame(header, fg_color="transparent")
        header_inner.pack(fill="x", padx=20, pady=15)
        
        ctk.CTkLabel(
            header_inner, text="PCI Billing Tool",
            font=ctk.CTkFont(size=22, weight="bold")
        ).pack(side="left")
        
        ctk.CTkLabel(
            header_inner, text=f"v{__version__}",
            font=ctk.CTkFont(size=11), text_color="gray"
        ).pack(side="left", padx=(10, 0), pady=(5, 0))
        
        # Tabs
        self.tabs = ctk.CTkTabview(self, corner_radius=8)
        self.tabs.pack(fill="both", expand=True, padx=20, pady=(10, 5))
        
        tab1 = self.tabs.add("Generate Bills")
        tab2 = self.tabs.add("Combine Files")
        
        self._build_generate_tab(tab1)
        self._build_combine_tab(tab2)
        
        # Status bar
        self.status = StatusBar(self)
        self.status.pack(fill="x", pady=(0, 10))
    
    def _build_generate_tab(self, parent):
        """Build the Generate Bills tab."""
        
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Description
        ctk.CTkLabel(
            container,
            text="Generate billing summary and import file from attendance data",
            font=ctk.CTkFont(size=13), text_color="gray"
        ).pack(anchor="w", pady=(0, 15))
        
        # File picker
        self.gen_file = FilePickerRow(
            container, "Attendance File:",
            placeholder="Select Excel or CSV file",
            filetypes=[("Excel/CSV", "*.xlsx;*.xls;*.csv")]
        )
        self.gen_file.pack(fill="x", pady=6)
        
        # School year
        self.gen_year = InputRow(container, "School Year:", "e.g., 2025")
        self.gen_year.pack(fill="x", pady=6)
        if self.prefs.get('school_year'):
            self.gen_year.insert(0, self.prefs['school_year'])
        
        # Item date
        self.gen_date = DateRow(container, "Billing Date:")
        self.gen_date.pack(fill="x", pady=6)
        if self.prefs.get('item_date'):
            self.gen_date.set(self.prefs['item_date'])
        
        # Ignore keywords info
        ignore_frame = ctk.CTkFrame(container, fg_color=("gray90", "gray25"), corner_radius=6)
        ignore_frame.pack(fill="x", pady=(15, 10))
        
        ignore_text = f"Rows with these keywords in Notes will be skipped: {', '.join(Config.IGNORE_KEYWORDS)}"
        ctk.CTkLabel(
            ignore_frame, text=ignore_text,
            font=ctk.CTkFont(size=11), text_color="gray"
        ).pack(padx=12, pady=8)
        
        # Generate button
        ctk.CTkButton(
            container, text="Generate Bills",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=45, corner_radius=8,
            fg_color="#2563eb", hover_color="#1d4ed8",
            command=self._generate_bills
        ).pack(fill="x", pady=(10, 15))
        
        # Results
        self.gen_results = ResultsBox(container)
        self.gen_results.pack(fill="both", expand=True)
    
    def _build_combine_tab(self, parent):
        """Build the Combine Files tab."""
        
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Description
        ctk.CTkLabel(
            container,
            text="Combine multiple Excel files into a single import-ready CSV",
            font=ctk.CTkFont(size=13), text_color="gray"
        ).pack(anchor="w", pady=(0, 15))
        
        # Folder picker
        self.combine_folder = FilePickerRow(
            container, "Folder:",
            placeholder="Select folder with Excel files",
            is_folder=True
        )
        self.combine_folder.pack(fill="x", pady=6)
        
        # Sheet name
        self.combine_sheet = InputRow(container, "Sheet Name:", "e.g., Sheet2", width=200)
        self.combine_sheet.pack(fill="x", pady=6)
        self.combine_sheet.insert(0, "Veracross Book Bill Import")
        
        # School year
        self.combine_year = InputRow(container, "School Year:", "e.g., 2025")
        self.combine_year.pack(fill="x", pady=6)
        if self.prefs.get('school_year'):
            self.combine_year.insert(0, self.prefs['school_year'])
        
        # Item date
        self.combine_date = DateRow(container, "Billing Date:")
        self.combine_date.pack(fill="x", pady=6)
        if self.prefs.get('item_date'):
            self.combine_date.set(self.prefs['item_date'])
        
        # Combine button
        ctk.CTkButton(
            container, text="Combine Files",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=45, corner_radius=8,
            fg_color="#2563eb", hover_color="#1d4ed8",
            command=self._combine_files
        ).pack(fill="x", pady=(20, 15))
        
        # Results
        self.combine_results = ResultsBox(container)
        self.combine_results.pack(fill="both", expand=True)
    
    def _generate_bills(self):
        """Generate billing files from attendance data."""
        # Get inputs
        filepath = self.gen_file.get()
        school_year = self.gen_year.get()
        item_date = self.gen_date.get()
        
        # Validate
        if not filepath:
            Toast.show(self, "Please select a file", "error")
            return
        if not os.path.exists(filepath):
            Toast.show(self, "File not found", "error")
            return
        if not school_year:
            Toast.show(self, "Please enter school year", "error")
            return
        if not item_date:
            Toast.show(self, "Please enter billing date", "error")
            return
        
        self.status.set("Processing...", "working")
        self.gen_results.clear()
        self.update()
        
        try:
            # Process file
            summary_df, bookbill_df, stats, skipped_df = process_attendance_file(
                filepath, school_year, item_date
            )
            
            # Save outputs
            base_dir = os.path.dirname(filepath)
            base_name = os.path.splitext(os.path.basename(filepath))[0]
            
            summary_path = os.path.join(base_dir, f"{base_name}_SUMMARY.xlsx")
            bookbill_path = os.path.join(base_dir, f"{base_name}_BOOKBILL.csv")
            
            summary_df.to_excel(summary_path, index=False)
            bookbill_df.to_csv(bookbill_path, index=False)
            
            files_created = [summary_path, bookbill_path]
            
            # Save skipped rows if any
            if not skipped_df.empty:
                skipped_path = os.path.join(base_dir, f"{base_name}_SKIPPED.csv")
                skipped_df.to_csv(skipped_path, index=False)
                files_created.append(skipped_path)
            
            # Save preferences
            self.prefs['school_year'] = school_year
            self.prefs['item_date'] = item_date
            save_prefs(self.prefs)
            
            # Show results
            result_text = (
                f"Completed Successfully\n\n"
                f"Students:        {stats['students']}\n"
                f"Total Hours:     {stats['total_hours']}\n"
                f"Total Charges:   ${stats['total_charges']:,.2f}\n"
                f"Staff Children:  {stats['staff_count']}\n"
                f"Rows Skipped:    {stats['skipped_count']}\n\n"
                f"Files Created:\n"
            )
            for f in files_created:
                result_text += f"  • {os.path.basename(f)}\n"
            
            self.gen_results.set(result_text)
            self.status.set(f"Saved to {base_dir}", "success")
            Toast.show(self, "Files generated successfully", "success")
            
        except BillingError as e:
            self.gen_results.set(f"Error:\n{str(e)}")
            self.status.set("Error", "error")
            Toast.show(self, str(e)[:60], "error")
            
        except Exception as e:
            log_error(e, "Generate bills")
            error_msg = get_user_friendly_error(e)
            self.gen_results.set(f"Error:\n{error_msg}")
            self.status.set("Error", "error")
            Toast.show(self, error_msg[:60], "error")
    
    def _combine_files(self):
        """Combine multiple Excel files."""
        # Get inputs
        folder = self.combine_folder.get()
        sheet_name = self.combine_sheet.get()
        school_year = self.combine_year.get()
        item_date = self.combine_date.get()
        
        # Validate
        if not folder:
            Toast.show(self, "Please select a folder", "error")
            return
        if not os.path.isdir(folder):
            Toast.show(self, "Folder not found", "error")
            return
        if not sheet_name:
            Toast.show(self, "Please enter sheet name", "error")
            return
        if not school_year:
            Toast.show(self, "Please enter school year", "error")
            return
        if not item_date:
            Toast.show(self, "Please enter billing date", "error")
            return
        
        self.status.set("Combining files...", "working")
        self.combine_results.clear()
        self.update()
        
        try:
            # Combine
            combined_df, processed, skipped = combine_files(
                folder, sheet_name, school_year, item_date
            )
            
            # Prompt for save location
            output_path = filedialog.asksaveasfilename(
                initialdir=folder,
                initialfile=f"Combined_{school_year}_{item_date.replace('/', '-')}.csv",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")]
            )
            
            if not output_path:
                self.status.set("Cancelled", "info")
                return
            
            # Save
            combined_df.to_csv(output_path, index=False)
            
            # Save prefs
            self.prefs['school_year'] = school_year
            self.prefs['item_date'] = item_date
            save_prefs(self.prefs)
            
            # Show results
            result_text = (
                f"Completed Successfully\n\n"
                f"Files Processed: {processed}\n"
                f"Total Rows:      {len(combined_df)}\n"
                f"Files Skipped:   {len(skipped)}\n\n"
                f"Output: {os.path.basename(output_path)}"
            )
            
            if skipped:
                result_text += "\n\nSkipped Files:"
                for filename, reason in skipped[:5]:
                    result_text += f"\n  • {filename}: {reason}"
                if len(skipped) > 5:
                    result_text += f"\n  ... and {len(skipped) - 5} more"
            
            self.combine_results.set(result_text)
            self.status.set("Files combined", "success")
            Toast.show(self, f"Combined {processed} files", "success")
            
        except BillingError as e:
            self.combine_results.set(f"Error:\n{str(e)}")
            self.status.set("Error", "error")
            Toast.show(self, str(e)[:60], "error")
            
        except Exception as e:
            log_error(e, "Combine files")
            error_msg = get_user_friendly_error(e)
            self.combine_results.set(f"Error:\n{error_msg}")
            self.status.set("Error", "error")
            Toast.show(self, error_msg[:60], "error")


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    app = PCIBillingApp()
    app.mainloop()


if __name__ == "__main__":
    main()
