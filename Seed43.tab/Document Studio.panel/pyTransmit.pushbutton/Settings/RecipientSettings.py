# -*- coding: utf-8 -*-
"""
Recipient Database Manager for pyTransmit
==========================================
This script provides a comprehensive database management interface for recipient data.

Features:
- Import from CSV/Excel with column mapping
- Export to CSV/Excel (all or selected records)
- Editable DataGrid with inline validation
- Drag & drop row reordering
- Multi-select with Ctrl/Shift
- Context menu (Edit, Delete, Duplicate, Copy)
- Duplicate detection with fuzzy matching
- Search/filter functionality
- Auto-save to SQLite database

Author: pyTransmit Suite
Version: 1.0
"""

import os
import sys
import json
import csv
from difflib import SequenceMatcher
from pyrevit import revit, forms, script, DB
from pyrevit.forms import WPFWindow
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")

from System.Windows import Application
from System.Windows.Controls import DataGrid
from System.Windows.Input import MouseButtonState
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows.Forms import OpenFileDialog, SaveFileDialog, DialogResult
from System.Windows.Forms import OpenFileDialog as WinFormsOpenFileDialog
from System import EventHandler

# xlsxwriter is bundled with pyRevit's IronPython environment — no install needed.

# For Excel import, we'll use COM automation (works in IronPython)
EXCEL_READ_SUPPORT = True  # COM is always available on Windows

# ═══════════════════════════════════════════════════════════════════════════
# DATABASE CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════

DB_FOLDER = os.path.dirname(os.path.abspath(__file__))
DB_FILE   = os.path.join(DB_FOLDER, 'recipients.json')
DIST_FILE = os.path.join(DB_FOLDER, 'distribution.json')

# ═══════════════════════════════════════════════════════════════════════════
# DATA MODEL
# ═══════════════════════════════════════════════════════════════════════════

class RecipientRecord(INotifyPropertyChanged):
    """Client List record: Company + Attention To."""
    def __init__(self, id=None, company="", attention_to="", recipient=""):
        self._id = id
        # Accept legacy 'recipient' kwarg as company for backwards compat
        self._company = company or recipient
        self._attention_to = attention_to
        self._status = "✓ Valid"
        self._status_color = "#00FF00"
        self._property_changed_handlers = []

    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)

    def _notify_property_changed(self, property_name):
        args = PropertyChangedEventArgs(property_name)
        for handler in self._property_changed_handlers:
            handler(self, args)

    @property
    def Id(self):
        return self._id

    @Id.setter
    def Id(self, value):
        if self._id != value:
            self._id = value
            self._notify_property_changed("Id")

    @property
    def Company(self):
        return self._company

    @Company.setter
    def Company(self, value):
        if self._company != value:
            self._company = value
            self._notify_property_changed("Company")
            self._validate()

    # Keep Recipient as alias so old code doesn't break
    @property
    def Recipient(self):
        return self._company

    @Recipient.setter
    def Recipient(self, value):
        self.Company = value

    @property
    def AttentionTo(self):
        return self._attention_to

    @AttentionTo.setter
    def AttentionTo(self, value):
        if self._attention_to != value:
            self._attention_to = value
            self._notify_property_changed("AttentionTo")
            self._validate()

    @property
    def Status(self):
        return self._status

    @Status.setter
    def Status(self, value):
        if self._status != value:
            self._status = value
            self._notify_property_changed("Status")

    @property
    def StatusColor(self):
        return self._status_color

    @StatusColor.setter
    def StatusColor(self, value):
        if self._status_color != value:
            self._status_color = value
            self._notify_property_changed("StatusColor")

    def _validate(self):
        if not self._company:
            self.Status = "⚠ Invalid"
            self.StatusColor = "#FF0000"
        else:
            self.Status = "✓ Valid"
            self.StatusColor = "#00FF00"


class DistributionRecord(INotifyPropertyChanged):
    """Distribution List record: a single role/label (e.g. Architect, Contractor)."""
    def __init__(self, id=None, distribution=""):
        self._id = id
        self._distribution = distribution
        self._property_changed_handlers = []

    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)

    def _notify(self, prop):
        args = PropertyChangedEventArgs(prop)
        for handler in self._property_changed_handlers:
            handler(self, args)

    @property
    def Id(self):
        return self._id

    @property
    def Distribution(self):
        return self._distribution

    @Distribution.setter
    def Distribution(self, value):
        if self._distribution != value:
            self._distribution = value
            self._notify("Distribution")

# ═══════════════════════════════════════════════════════════════════════════
# DATABASE OPERATIONS
# ═══════════════════════════════════════════════════════════════════════════

class RecipientDatabase:
    """JSON-based database for Client List (Company + Attention To)."""
    def __init__(self, db_path=DB_FILE):
        self.db_path = db_path
        if not os.path.exists(self.db_path):
            self._save_json([])

    def _load_json(self):
        try:
            with open(self.db_path, 'r') as f:
                return json.load(f)
        except (IOError, ValueError):
            return []

    def _save_json(self, data):
        with open(self.db_path, 'w') as f:
            json.dump(data, f, indent=2)

    def load_all(self):
        data = self._load_json()
        records = []
        for i, item in enumerate(data):
            # Support old key 'recipient' for backwards compat
            company = item.get('company', item.get('recipient', ''))
            records.append(RecipientRecord(
                id=item.get('id', i),
                company=company,
                attention_to=item.get('attention_to', '')
            ))
        return records

    def save_all(self, records):
        data = []
        for i, record in enumerate(records):
            data.append({
                'id': i,
                'company': record.Company,
                'attention_to': record.AttentionTo,
                'display_order': i
            })
        self._save_json(data)


class DistributionDatabase:
    """JSON-based database for Distribution List (role labels only)."""
    DEFAULT = ['Architect/Designer', 'Owner/Developer', 'Contractor', 'Local Authority']

    def __init__(self, db_path=DIST_FILE):
        self.db_path = db_path
        if not os.path.exists(self.db_path):
            self._save_json([{'id': i, 'distribution': v, 'display_order': i}
                             for i, v in enumerate(self.DEFAULT)])

    def _load_json(self):
        try:
            with open(self.db_path, 'r') as f:
                return json.load(f)
        except (IOError, ValueError):
            return []

    def _save_json(self, data):
        with open(self.db_path, 'w') as f:
            json.dump(data, f, indent=2)

    def load_all(self):
        return [DistributionRecord(id=item.get('id', i),
                                   distribution=item.get('distribution', ''))
                for i, item in enumerate(self._load_json())]

    def save_all(self, records):
        self._save_json([{'id': i, 'distribution': r.Distribution, 'display_order': i}
                         for i, r in enumerate(records)])

# ═══════════════════════════════════════════════════════════════════════════
# IMPORT/EXPORT UTILITIES
# ═══════════════════════════════════════════════════════════════════════════

def read_excel(file_path):
    """
    Read Excel file using COM automation (compatible with IronPython).
    Returns: (headers_list, rows_list)
    """
    from System import Type, Activator, Array
    from System.Runtime.InteropServices import Marshal
    import System.Reflection as Reflection
    
    excel = None
    workbook = None
    
    try:
        # Create Excel application instance using late binding
        excel_type = Type.GetTypeFromProgID("Excel.Application")
        excel = Activator.CreateInstance(excel_type)
        
        # Set properties using reflection (late binding)
        excel_type.InvokeMember("Visible", 
            Reflection.BindingFlags.SetProperty,
            None, excel, Array[object]([False]))
        
        excel_type.InvokeMember("DisplayAlerts",
            Reflection.BindingFlags.SetProperty,
            None, excel, Array[object]([False]))
        
        # Get workbooks collection
        workbooks = excel_type.InvokeMember("Workbooks",
            Reflection.BindingFlags.GetProperty,
            None, excel, None)
        
        # Open the workbook
        workbooks_type = workbooks.GetType()
        workbook = workbooks_type.InvokeMember("Open",
            Reflection.BindingFlags.InvokeMethod,
            None, workbooks, Array[object]([file_path]))
        
        # Get worksheets collection
        workbook_type = workbook.GetType()
        worksheets = workbook_type.InvokeMember("Worksheets",
            Reflection.BindingFlags.GetProperty,
            None, workbook, None)
        
        # Get first worksheet
        worksheets_type = worksheets.GetType()
        worksheet = worksheets_type.InvokeMember("Item",
            Reflection.BindingFlags.GetProperty,
            None, worksheets, Array[object]([1]))
        
        # Get used range
        worksheet_type = worksheet.GetType()
        used_range = worksheet_type.InvokeMember("UsedRange",
            Reflection.BindingFlags.GetProperty,
            None, worksheet, None)
        
        # Get row and column counts
        used_range_type = used_range.GetType()
        rows = used_range_type.InvokeMember("Rows",
            Reflection.BindingFlags.GetProperty,
            None, used_range, None)
        
        rows_type = rows.GetType()
        rows_count = rows_type.InvokeMember("Count",
            Reflection.BindingFlags.GetProperty,
            None, rows, None)
        
        columns = used_range_type.InvokeMember("Columns",
            Reflection.BindingFlags.GetProperty,
            None, used_range, None)
        
        columns_type = columns.GetType()
        cols_count = columns_type.InvokeMember("Count",
            Reflection.BindingFlags.GetProperty,
            None, columns, None)
        
        # Get cells collection
        cells = worksheet_type.InvokeMember("Cells",
            Reflection.BindingFlags.GetProperty,
            None, worksheet, None)
        cells_type = cells.GetType()
        
        # Read all data
        all_data = []
        for row in range(1, rows_count + 1):
            row_data = []
            for col in range(1, cols_count + 1):
                # Get cell
                cell = cells_type.InvokeMember("Item",
                    Reflection.BindingFlags.GetProperty,
                    None, cells, Array[object]([row, col]))
                
                # Get cell value
                cell_type = cell.GetType()
                cell_value = cell_type.InvokeMember("Value2",
                    Reflection.BindingFlags.GetProperty,
                    None, cell, None)
                
                # Convert to string, handle None
                if cell_value is None:
                    cell_value = ""
                else:
                    cell_value = str(cell_value)
                row_data.append(cell_value)
                
                # Release cell COM object
                Marshal.ReleaseComObject(cell)
            
            all_data.append(row_data)
        
        # Release COM objects
        Marshal.ReleaseComObject(cells)
        Marshal.ReleaseComObject(columns)
        Marshal.ReleaseComObject(rows)
        Marshal.ReleaseComObject(used_range)
        Marshal.ReleaseComObject(worksheet)
        Marshal.ReleaseComObject(worksheets)
        
        # Close workbook without saving
        workbook_type.InvokeMember("Close",
            Reflection.BindingFlags.InvokeMethod,
            None, workbook, Array[object]([False]))
        Marshal.ReleaseComObject(workbook)
        Marshal.ReleaseComObject(workbooks)
        
        # Quit Excel
        excel_type.InvokeMember("Quit",
            Reflection.BindingFlags.InvokeMethod,
            None, excel, None)
        Marshal.ReleaseComObject(excel)
        
        # Extract headers and data
        if not all_data:
            return [], []
        
        headers = all_data[0]
        
        # Filter out empty rows from data
        data_rows = []
        for row in all_data[1:]:
            if any(cell.strip() for cell in row):
                data_rows.append(row)
        
        return headers, data_rows
    
    except Exception as e:
        # Make sure to clean up COM objects on error
        try:
            if workbook:
                workbook_type = workbook.GetType()
                workbook_type.InvokeMember("Close",
                    Reflection.BindingFlags.InvokeMethod,
                    None, workbook, Array[object]([False]))
                Marshal.ReleaseComObject(workbook)
        except:
            pass
        
        try:
            if excel:
                excel_type = excel.GetType()
                excel_type.InvokeMember("Quit",
                    Reflection.BindingFlags.InvokeMethod,
                    None, excel, None)
                Marshal.ReleaseComObject(excel)
        except:
            pass
        
        raise Exception("Error reading Excel file: {}".format(str(e)))

def read_csv(file_path):
    """
    Read CSV file and return headers + data rows.
    Returns: (headers_list, rows_list)
    """
    try:
        with open(file_path, 'r') as f:
            reader = csv.reader(f)
            rows = list(reader)
        
        if not rows:
            return [], []
        
        headers = rows[0]
        data_rows = [row for row in rows[1:] if any(cell.strip() for cell in row)]
        
        return headers, data_rows
    
    except Exception as e:
        raise Exception("Error reading CSV file: {}".format(str(e)))

def write_csv(file_path, records):
    """Write records to CSV file"""
    with open(file_path, 'w') as f:
        writer = csv.writer(f)
        writer.writerow(['Recipient', 'Attention To'])
        
        for record in records:
            writer.writerow([record.Recipient, record.AttentionTo])

def write_excel(file_path, records):
    """Write records to Excel file using xlsxwriter"""
    import xlsxwriter
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet('Recipients')
    
    # Add header format
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#208A3C',
        'font_color': 'white',
        'border': 1
    })
    
    # Write headers
    worksheet.write(0, 0, 'Recipient', header_format)
    worksheet.write(0, 1, 'Attention To', header_format)
    
    # Write data
    for i, record in enumerate(records, start=1):
        worksheet.write(i, 0, record.Recipient)
        worksheet.write(i, 1, record.AttentionTo)
    
    # Set column widths
    worksheet.set_column(0, 0, 30)
    worksheet.set_column(1, 1, 30)
    
    workbook.close()

def calculate_similarity(str1, str2):
    """Calculate similarity ratio between two strings"""
    return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()

# ═══════════════════════════════════════════════════════════════════════════
# PREVIEW WINDOW
# ═══════════════════════════════════════════════════════════════════════════

class PreviewRecord(object):
    """Simple data holder for preview grid"""
    def __init__(self, recipient, attention_to):
        self.Recipient = recipient
        self.AttentionTo = attention_to

class ImportPreviewWindow(WPFWindow):
    """
    Preview window for import data showing a DataGrid.
    """
    def __init__(self, preview_data):
        # Load XAML
        xaml_file = os.path.join(os.path.dirname(__file__), 'ImportPreview.xaml')
        WPFWindow.__init__(self, xaml_file)
        
        # Create preview collection
        self.preview_records = ObservableCollection[PreviewRecord]()
        
        # Add data to collection
        for recipient, attention_to in preview_data:
            self.preview_records.Add(PreviewRecord(recipient, attention_to))
        
        # Bind to grid
        self.preview_grid.ItemsSource = self.preview_records
        
        # Update info text
        total_rows = len(preview_data)
        if total_rows > 100:
            self.preview_info_tb.Text = "Showing first 100 of {} rows to be imported".format(total_rows)
        else:
            self.preview_info_tb.Text = "Showing all {} rows to be imported".format(total_rows)
    
    def close_window(self, sender, args):
        """Close the preview window"""
        self.Close()

# ═══════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ═══════════════════════════════════════════════════════════════════════════

class RecipientManagerWindow(WPFWindow):
    """
    Main window for recipient database management.
    """
    def __init__(self):
        # Load XAML
        xaml_file = os.path.join(os.path.dirname(__file__), 'RecipientManager.xaml')
        WPFWindow.__init__(self, xaml_file)
        
        # Initialize database
        self.db = RecipientDatabase()
        
        # Initialize data collection
        self.recipients = ObservableCollection[RecipientRecord]()
        
        # Load existing recipients
        for record in self.db.load_all():
            self.recipients.Add(record)
        
        # Bind to grid
        self.recipients_grid.ItemsSource = self.recipients
        
        # Import state
        self.import_file_path = None
        self.import_headers = []
        self.import_data = []
        
        # Drag & drop state
        self.drag_source_index = None
        
        # Update record count
        self._update_record_count()
        
        # Auto-save on window closing
        self.Closing += self._on_closing
    
    def _on_closing(self, sender, args):
        """Auto-save database when window closes"""
        try:
            self.db.save_all(list(self.recipients))
        except Exception as e:
            pass  # Silently fail on save error during close
    
    def _update_record_count(self):
        """Update the record count display"""
        self.record_count_tb.Text = "{} records".format(len(self.recipients))
    
    # ═══════════════════════════════════════════════════════════════════════
    # IMPORT HANDLERS
    # ═══════════════════════════════════════════════════════════════════════
    
    def browse_file(self, sender, args):
        """Browse for import file (CSV or Excel)"""
        dialog = OpenFileDialog()
        dialog.Title = "Select Import File"
        
        # Set filter - Excel is always supported via COM
        dialog.Filter = "Supported Files (*.csv;*.xlsx;*.xls)|*.csv;*.xlsx;*.xls|CSV Files (*.csv)|*.csv|Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
        
        if dialog.ShowDialog() == DialogResult.OK:
            try:
                file_path = dialog.FileName
                self.import_file_path = file_path
                
                # Read file based on extension
                ext = os.path.splitext(file_path)[1].lower()
                
                if ext in ['.xlsx', '.xls']:
                    headers, data = read_excel(file_path)
                elif ext == '.csv':
                    headers, data = read_csv(file_path)
                else:
                    forms.alert("Unsupported file format: {}".format(ext), title="Import Error")
                    return
                
                # Store import data
                self.import_headers = headers
                self.import_data = data
                
                # Update UI
                self.file_path_tb.Text = file_path
                
                # Populate column mapping dropdowns
                self._populate_column_mappings(headers)
                
                # Enable import buttons
                self.preview_btn.IsEnabled = True
                self.import_btn.IsEnabled = True
            
            except Exception as e:
                forms.alert("Error reading file:\n{}".format(str(e)), title="Import Error")
                self.import_file_path = None
                self.import_headers = []
                self.import_data = []
    
    def _populate_column_mappings(self, headers):
        """Populate column mapping dropdowns with file headers"""
        # Clear existing items
        self.recipient_mapping_cb.Items.Clear()
        self.attention_mapping_cb.Items.Clear()
        
        # Add blank option
        self.recipient_mapping_cb.Items.Add("-- Select Column --")
        self.attention_mapping_cb.Items.Add("-- Select Column --")
        
        # Add headers
        for header in headers:
            self.recipient_mapping_cb.Items.Add(header)
            self.attention_mapping_cb.Items.Add(header)
        
        # Try to auto-detect mappings based on common column names
        recipient_keywords = ['recipient', 'role', 'company', 'organization', 'to']
        attention_keywords = ['attention', 'attn', 'person', 'contact', 'name']
        
        for i, header in enumerate(headers, start=1):  # start=1 because of blank option
            header_lower = header.lower()
            
            # Check for recipient column
            if any(keyword in header_lower for keyword in recipient_keywords):
                if self.recipient_mapping_cb.SelectedIndex == 0:  # Not yet selected
                    self.recipient_mapping_cb.SelectedIndex = i
            
            # Check for attention column
            if any(keyword in header_lower for keyword in attention_keywords):
                if self.attention_mapping_cb.SelectedIndex == 0:  # Not yet selected
                    self.attention_mapping_cb.SelectedIndex = i
        
        # If no auto-detection, select first option
        if self.recipient_mapping_cb.SelectedIndex == -1:
            self.recipient_mapping_cb.SelectedIndex = 0
        if self.attention_mapping_cb.SelectedIndex == -1:
            self.attention_mapping_cb.SelectedIndex = 0
    
    def preview_import(self, sender, args):
        """Preview import data before importing - shows DataGrid preview window"""
        if not self._validate_import_mapping():
            return
        
        recipient_col = self.recipient_mapping_cb.SelectedIndex - 1  # -1 for blank option
        attention_col = self.attention_mapping_cb.SelectedIndex - 1
        
        # Build preview data (limit to first 100 rows for performance)
        preview_data = []
        for i, row in enumerate(self.import_data[:100]):
            recipient = row[recipient_col].strip() if recipient_col < len(row) else ""
            attention = row[attention_col].strip() if attention_col < len(row) else ""
            
            # Skip completely empty rows
            if recipient or attention:
                preview_data.append((recipient, attention))
        
        # Show preview window
        try:
            preview_window = ImportPreviewWindow(preview_data)
            preview_window.ShowDialog()
        except Exception as e:
            forms.alert("Error showing preview:\n{}".format(str(e)), title="Preview Error")
    
    def import_data(self, sender, args):
        """Import data into the grid"""
        if not self._validate_import_mapping():
            return
        
        recipient_col = self.recipient_mapping_cb.SelectedIndex - 1
        attention_col = self.attention_mapping_cb.SelectedIndex - 1
        
        # Import records
        imported_count = 0
        duplicate_count = 0
        
        for row in self.import_data:
            recipient = row[recipient_col].strip() if recipient_col < len(row) else ""
            attention = row[attention_col].strip() if attention_col < len(row) else ""
            
            # Skip empty rows
            if not recipient and not attention:
                continue
            
            # Check for exact duplicates (case-insensitive)
            is_duplicate = False
            for existing in self.recipients:
                # Use case-insensitive exact matching
                if (recipient.lower() == existing.Recipient.lower() and 
                    attention.lower() == existing.AttentionTo.lower()):
                    is_duplicate = True
                    duplicate_count += 1
                    break
            
            if not is_duplicate:
                new_record = RecipientRecord(
                    recipient=recipient,
                    attention_to=attention
                )
                self.recipients.Add(new_record)
                imported_count += 1
        
        # Update UI
        self._update_record_count()
        
        # Show summary
        summary = "Import complete!\n\n"
        summary += "Imported: {} records\n".format(imported_count)
        if duplicate_count > 0:
            summary += "Skipped (duplicates): {} records".format(duplicate_count)
        
        forms.alert(summary, title="Import Complete")
        
        # Reset import UI
        self._reset_import_ui()
    
    def _validate_import_mapping(self):
        """Validate that column mappings are selected"""
        if self.recipient_mapping_cb.SelectedIndex <= 0:
            forms.alert("Please select a column for 'Recipient'", title="Mapping Required")
            return False
        
        if self.attention_mapping_cb.SelectedIndex <= 0:
            forms.alert("Please select a column for 'Attention To'", title="Mapping Required")
            return False
        
        return True
    
    def _reset_import_ui(self):
        """Reset import UI to initial state"""
        self.file_path_tb.Text = ""
        self.recipient_mapping_cb.Items.Clear()
        self.attention_mapping_cb.Items.Clear()
        self.preview_btn.IsEnabled = False
        self.import_btn.IsEnabled = False
        self.import_file_path = None
        self.import_headers = []
        self.import_data = []
    
    # ═══════════════════════════════════════════════════════════════════════
    # GRID OPERATIONS
    # ═══════════════════════════════════════════════════════════════════════
    
    def add_row(self, sender, args):
        """Add a new blank row"""
        new_record = RecipientRecord()
        self.recipients.Add(new_record)
        self._update_record_count()
        
        # Select the new row
        self.recipients_grid.SelectedItem = new_record
        self.recipients_grid.ScrollIntoView(new_record)
    
    def delete_rows(self, sender, args):
        """Delete selected rows"""
        grid = getattr(self, 'recipients_grid', None)
        to_remove = list(grid.SelectedItems) if (grid and grid.SelectedItems) else []
        
        if not to_remove:
            return
        
        if forms.alert(
            "Delete {} record(s)?".format(len(to_remove)),
            title="Confirm Delete",
            ok=False,
            yes=True,
            no=True
        ):
            for record in to_remove:
                self.recipients.Remove(record)
            
            self._update_record_count()
            self.delete_rows_btn.IsEnabled = False
    
    def export_data(self, sender, args):
        """Export data to CSV or Excel"""
        grid = getattr(self, 'recipients_grid', None)
        selected_records = list(grid.SelectedItems) if (grid and grid.SelectedItems) else []
        
        if selected_records:
            export_choice = forms.CommandSwitchWindow.show(
                ['Export All Records', 'Export Selected Records Only'],
                message='What would you like to export?'
            )
            if export_choice == 'Export Selected Records Only':
                records_to_export = selected_records
            else:
                records_to_export = list(self.recipients)
        else:
            records_to_export = list(self.recipients)
        
        if not records_to_export:
            forms.alert("No records to export.", title="Export")
            return
        
        # Ask for file type
        format_options = ['Excel (.xlsx)', 'CSV (.csv)']
        
        file_type = forms.CommandSwitchWindow.show(
            format_options,
            message='Select export format:'
        )
        
        if not file_type:
            return
        
        # Get save location
        dialog = SaveFileDialog()
        dialog.Title = "Export Recipients"
        
        if file_type.startswith('Excel'):
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            dialog.DefaultExt = "xlsx"
        else:
            dialog.Filter = "CSV Files (*.csv)|*.csv"
            dialog.DefaultExt = "csv"
        
        if dialog.ShowDialog() == DialogResult.OK:
            try:
                if file_type.startswith('Excel'):
                    write_excel(dialog.FileName, records_to_export)
                else:
                    write_csv(dialog.FileName, records_to_export)
                
                forms.alert(
                    "Exported {} record(s) successfully!".format(len(records_to_export)),
                    title="Export Complete"
                )
            except Exception as e:
                forms.alert("Error exporting:\n{}".format(str(e)), title="Export Error")
    
    def search_changed(self, sender, args):
        """Filter grid based on search text"""
        from System.Windows.Data import CollectionViewSource
        
        search_text = self.search_tb.Text.lower().strip()
        
        # Get the view from the grid's ItemsSource
        view = CollectionViewSource.GetDefaultView(self.recipients_grid.ItemsSource)
        
        if not search_text:
            # Clear filter - show all records
            view.Filter = None
        else:
            # Set filter to search in both Recipient and AttentionTo fields
            def filter_predicate(item):
                try:
                    recipient = item.Recipient.lower() if item.Recipient else ""
                    attention = item.AttentionTo.lower() if item.AttentionTo else ""
                    return search_text in recipient or search_text in attention
                except:
                    return False
            
            view.Filter = filter_predicate
        
        # Refresh the view
        view.Refresh()
    
    # ═══════════════════════════════════════════════════════════════════════
    # SELECTION HANDLERS
    # ═══════════════════════════════════════════════════════════════════════
    
    def grid_selection_changed(self, sender, args):
        """Enable/disable delete button based on grid highlight selection."""
        grid = getattr(self, 'recipients_grid', None)
        has_selection = grid is not None and grid.SelectedItems is not None and len(list(grid.SelectedItems)) > 0
        self.delete_rows_btn.IsEnabled = has_selection
        self._update_context_menu_select_items()

    def _update_context_menu_select_items(self):
        """Show 'Select All' when nothing/partial selected; show 'Select None' when anything selected."""
        try:
            grid = getattr(self, 'recipients_grid', None)
            has_selection = grid is not None and grid.SelectedItems is not None and len(list(grid.SelectedItems)) > 0
            total = len(list(self.recipients)) if self.recipients else 0
            all_selected = has_selection and len(list(grid.SelectedItems)) == total

            select_all_item  = getattr(self, 'rec_ctx_select_all', None)
            select_none_item = getattr(self, 'rec_ctx_select_none', None)
            from System.Windows import Visibility
            if select_all_item:
                select_all_item.Visibility = Visibility.Collapsed if all_selected else Visibility.Visible
            if select_none_item:
                select_none_item.Visibility = Visibility.Visible if has_selection else Visibility.Collapsed
        except: pass

    def context_select_all(self, sender, args):
        """Right-click Select All — highlights all rows in the grid."""
        try:
            grid = getattr(self, 'recipients_grid', None)
            if grid:
                grid.SelectAll()
        except: pass

    def context_select_none(self, sender, args):
        """Right-click Select None — clears all grid highlights."""
        try:
            grid = getattr(self, 'recipients_grid', None)
            if grid:
                grid.UnselectAll()
            self.delete_rows_btn.IsEnabled = False
            self._update_context_menu_select_items()
        except: pass
    
    # ═══════════════════════════════════════════════════════════════════════
    # DRAG & DROP — row reordering with green drop indicator
    # ═══════════════════════════════════════════════════════════════════════

    _drag_source_row  = None
    _drag_start_pos   = None
    _drop_indicator   = None
    _drop_target_idx  = -1

    def _setup_drag_drop(self):
        grid = getattr(self, 'recipients_grid', None)
        if grid is None:
            return
        grid.AllowDrop = True
        grid.PreviewMouseLeftButtonDown += self._drag_mouse_down
        grid.MouseMove                  += self._drag_mouse_move
        grid.DragOver                   += self._drag_over
        grid.Drop                       += self._drag_drop
        grid.DragLeave                  += self._drag_leave

    def _get_row_at_point(self, grid, point):
        try:
            from System.Windows.Media import VisualTreeHelper
            el = grid.InputHitTest(point)
            while el is not None:
                if hasattr(el, 'DataContext') and el.DataContext in grid.ItemsSource:
                    idx = list(grid.ItemsSource).index(el.DataContext)
                    return el.DataContext, idx
                try:   el = VisualTreeHelper.GetParent(el)
                except: break
        except: pass
        return None, -1

    def _show_drop_indicator(self, grid, idx):
        """Draw a 2px green line at the insertion point.

        Strategy: find the nearest Panel ancestor, then use
        TransformToAncestor in one call to get the grid's origin
        relative to that panel -- no incremental loop needed.
        """
        try:
            self._hide_drop_indicator()
            from System.Windows.Controls import Border, Panel
            from System.Windows.Media import SolidColorBrush, Color, VisualTreeHelper
            import System.Windows

            row_h  = grid.RowHeight          if grid.RowHeight          > 0 else 28
            col_hh = grid.ColumnHeaderHeight if grid.ColumnHeaderHeight > 0 else 36
            y_local = col_hh + idx * row_h

            # Walk up the visual tree to find the nearest Panel ancestor
            panel  = None
            cursor = VisualTreeHelper.GetParent(grid)
            while cursor is not None:
                if isinstance(cursor, Panel):
                    panel = cursor
                    break
                cursor = VisualTreeHelper.GetParent(cursor)

            if panel is None:
                return

            # Get the grid's top-left corner in panel coordinates (one transform)
            try:
                tf      = grid.TransformToAncestor(panel)
                origin  = tf.Transform(System.Windows.Point(0, 0))
                offset_y = origin.Y
            except:
                return

            indicator                     = Border()
            indicator.Height              = 2
            indicator.Background          = SolidColorBrush(Color.FromRgb(0x20, 0x8A, 0x3C))
            indicator.IsHitTestVisible    = False
            indicator.Width               = grid.ActualWidth
            indicator.VerticalAlignment   = System.Windows.VerticalAlignment.Top
            indicator.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
            indicator.Margin              = System.Windows.Thickness(
                                               origin.X,
                                               offset_y + y_local,
                                               0, 0)
            panel.Children.Add(indicator)
            self._drop_indicator = (indicator, panel)
        except: pass

    def _hide_drop_indicator(self):
        try:
            if self._drop_indicator:
                ind, parent = self._drop_indicator
                parent.Children.Remove(ind)
                self._drop_indicator = None
        except:
            self._drop_indicator = None

    def _drag_mouse_down(self, sender, args):
        try:
            self._drag_start_pos  = args.GetPosition(sender)
            row, _                = self._get_row_at_point(sender, self._drag_start_pos)
            self._drag_source_row = row
        except: pass

    def _drag_mouse_move(self, sender, args):
        if self._drag_source_row is None or self._drag_start_pos is None:
            return
        try:
            import System.Windows.Input as WI
            if args.LeftButton != WI.MouseButtonState.Pressed:
                self._drag_source_row = None
                return
            pos = args.GetPosition(sender)
            if (abs(pos.X - self._drag_start_pos.X) < 4 and
                    abs(pos.Y - self._drag_start_pos.Y) < 4):
                return
            from System.Windows import DragDrop, DataObject, DragDropEffects
            DragDrop.DoDragDrop(sender,
                                DataObject('RecRow', self._drag_source_row),
                                DragDropEffects.Move)
            self._drag_source_row = None
        except: pass

    def _drag_over(self, sender, args):
        try:
            from System.Windows import DragDropEffects
            args.Effects = DragDropEffects.Move
            args.Handled = True
            pos    = args.GetPosition(sender)
            _, idx = self._get_row_at_point(sender, pos)
            if idx < 0:
                idx = len(list(self.data))
            self._drop_target_idx = idx
            self._show_drop_indicator(sender, idx)
        except: pass

    def _drag_leave(self, sender, args):
        self._hide_drop_indicator()

    def _drag_drop(self, sender, args):
        self._hide_drop_indicator()
        try:
            src_row = args.Data.GetData('RecRow')
            if src_row is None:
                return
            items = list(self.data)
            if src_row not in items:
                return
            src_idx  = items.index(src_row)
            dest_idx = self._drop_target_idx
            if dest_idx < 0:
                dest_idx = len(items)
            if dest_idx == src_idx or dest_idx == src_idx + 1:
                return
            self.data.Remove(src_row)
            if src_idx < dest_idx:
                dest_idx -= 1
            if dest_idx >= len(list(self.data)):
                self.data.Add(src_row)
            else:
                self.data.Insert(dest_idx, src_row)
            self.save()
            args.Handled = True
        except: pass

    def grid_mouse_down(self, sender, args): pass   # kept for compat
    def grid_mouse_move(self, sender, args): pass
    def grid_drag_over(self, sender, args):  pass
    def grid_drop(self, sender, args):       pass
    def _find_ancestor(self, el, t):         return None

    # ── Column-header sort (saves order) ─────────────────────────────

    _last_sort = {}   # {'col': 'asc'|'desc'}

    def grid_sorting(self, sender, args):
        """
        Intercept DataGrid Sorting event, sort the ObservableCollection
        directly so order persists, toggle asc/desc on repeat clicks.
        """
        try:
            args.Handled   = True
            col_header     = str(args.Column.Header).lower()
            items          = list(self.data)
            if not items:
                return
            if col_header in ('recipient (role / company)', 'recipient'):
                key_fn = lambda r: (r.Recipient  or '').lower()
            elif col_header in ('attention to (person)', 'attention to'):
                key_fn = lambda r: (r.AttentionTo or '').lower()
            else:
                return   # unknown column — skip
            prev = self._last_sort.get(col_header)
            if prev == 'asc':
                items.sort(key=key_fn, reverse=True)
                self._last_sort[col_header] = 'desc'
            else:
                items.sort(key=key_fn)
                self._last_sort[col_header] = 'asc'
            self.data.Clear()
            for item in items:
                self.data.Add(item)
            self.save()
        except: pass
    
    # ═══════════════════════════════════════════════════════════════════════
    # CONTEXT MENU HANDLERS
    # ═══════════════════════════════════════════════════════════════════════
    
    def context_edit(self, sender, args):
        """Edit selected row"""
        # Grid is already editable, this could open a detail dialog if needed
        pass
    
    def context_duplicate(self, sender, args):
        """Duplicate selected row"""
        selected = self.recipients_grid.SelectedItem
        if selected:
            new_record = RecipientRecord(
                recipient=selected.Recipient,
                attention_to=selected.AttentionTo
            )
            self.recipients.Add(new_record)
            self._update_record_count()
    
    def context_delete(self, sender, args):
        """Delete selected row"""
        selected = self.recipients_grid.SelectedItem
        if selected:
            if forms.alert(
                "Delete this record?",
                title="Confirm Delete",
                ok=False,
                yes=True,
                no=True
            ):
                self.recipients.Remove(selected)
                self._update_record_count()
    
    def context_copy(self, sender, args):
        """Copy selected rows to clipboard"""
        grid     = getattr(self, 'recipients_grid', None)
        selected = list(grid.SelectedItems) if (grid and grid.SelectedItems) else []
        if not selected and grid and grid.SelectedItem:
            selected = [grid.SelectedItem]
        if selected:
            from System.Windows import Clipboard
            Clipboard.SetText("\n".join(
                "{}\t{}".format(r.Recipient, r.AttentionTo) for r in selected))


# ═══════════════════════════════════════════════════════════════════════════
# MAIN EXECUTION
# ═══════════════════════════════════════════════════════════════════════════

def main():
    try:
        window = RecipientManagerWindow()
        window.ShowDialog()
    except Exception as e:
        forms.alert("Error initializing window:\n{}".format(str(e)), exitscript=True)

if __name__ == "__main__":
    main()


# ═══════════════════════════════════════════════════════════════════════════
# PANEL CONTROLLER  — embedded use inside pyTransmit main window
# ═══════════════════════════════════════════════════════════════════════════

class RecipientSettingsController(object):
    """
    Drives the Recipient panel (embedded in pyTransmit) with two tabs:
      - Client List        : Company + Attention To  (recipients.json)
      - Distribution List  : Role label only          (distribution.json)
    """

    def __init__(self):
        # ── Client List ───────────────────────────────────────────────
        self.db   = RecipientDatabase()
        self.data = ObservableCollection[RecipientRecord]()
        for r in self.db.load_all():
            self.data.Add(r)

        # ── Distribution List ─────────────────────────────────────────
        self.dist_db   = DistributionDatabase()
        self.dist_data = ObservableCollection[DistributionRecord]()
        for r in self.dist_db.load_all():
            self.dist_data.Add(r)

        self._import_headers = []
        self._import_data    = []
        self._snapshot       = {}
        self._current_tab    = 'client'   # 'client' | 'dist'

    # ── Snapshot / discard ────────────────────────────────────────────

    def _take_snapshot(self):
        self._snapshot = {
            'client': [(r.Company, r.AttentionTo) for r in self.data],
            'dist':   [r.Distribution for r in self.dist_data],
        }

    def discard(self):
        snap = self._snapshot
        if not snap:
            return
        self.data.Clear()
        for company, attn in snap.get('client', []):
            self.data.Add(RecipientRecord(company=company, attention_to=attn))
        self.dist_data.Clear()
        for label in snap.get('dist', []):
            self.dist_data.Add(DistributionRecord(distribution=label))

    # ── Save ──────────────────────────────────────────────────────────

    def save(self):
        self.db.save_all(list(self.data))
        self.dist_db.save_all(list(self.dist_data))

    # ── Attach ────────────────────────────────────────────────────────

    def attach(self, root):
        self._walk(root)
        self._wire_events()
        try: self.recipients_grid.ItemsSource = self.data
        except: pass
        try: self.dist_grid.ItemsSource = self.dist_data
        except: pass
        self._show_tab('client')

    def _walk(self, element):
        if element is None:
            return
        try:
            n = element.Name
            if n:
                setattr(self, n, element)
        except: pass
        from System.Windows.Controls import Panel, ContentControl, Decorator
        try:
            if isinstance(element, Panel):
                for child in element.Children:
                    self._walk(child)
                return
        except: pass
        try:
            if isinstance(element, Decorator):
                self._walk(element.Child); return
        except: pass
        try:
            if isinstance(element, ContentControl):
                self._walk(element.Content)
        except: pass
        try:
            for child in element.Children:
                self._walk(child)
        except: pass
        try:
            cm = element.ContextMenu
            if cm:
                for item in cm.Items:
                    self._walk(item)
        except: pass

    def _wire_events(self):
        def bind(name, event, handler):
            el = getattr(self, name, None)
            if el is not None:
                try: getattr(el, event).__iadd__(handler)
                except: pass

        # Tab buttons
        bind('rec_client_tab_btn', 'Click', self.switch_to_client_tab)
        bind('rec_dist_tab_btn',   'Click', self.switch_to_dist_tab)

        # Shared add/delete — dispatch based on _current_tab
        bind('rec_add_btn',    'Click', self._add_row_dispatch)
        bind('rec_delete_btn', 'Click', self._delete_rows_dispatch)

        # Import controls
        bind('rec_browse_file_btn', 'Click', self.browse_file)
        bind('rec_import_btn',      'Click', self.import_data)

        # Client List grid
        bind('recipients_grid',    'SelectionChanged', self.grid_selection_changed)
        bind('recipients_grid',    'Sorting',          self.grid_sorting)
        bind('rec_ctx_select_all',  'Click', self.context_select_all)
        bind('rec_ctx_select_none', 'Click', self.context_select_none)
        bind('rec_ctx_duplicate',   'Click', self.context_duplicate)
        bind('rec_ctx_delete',      'Click', self.context_delete)
        bind('rec_ctx_copy',        'Click', self.context_copy)

        # Distribution List grid
        bind('dist_grid',                'SelectionChanged', self.dist_grid_selection_changed)
        bind('dist_grid',                'Sorting',          self.dist_grid_sorting)
        bind('rec_dist_ctx_select_all',  'Click', self.dist_context_select_all)
        bind('rec_dist_ctx_select_none', 'Click', self.dist_context_select_none)
        bind('rec_dist_ctx_duplicate',   'Click', self.dist_context_duplicate)
        bind('rec_dist_ctx_delete',      'Click', self.dist_context_delete)
        bind('rec_dist_ctx_copy',        'Click', self.dist_context_copy)

        self._setup_drag_drop()

    def _add_row_dispatch(self, sender, args):
        if self._current_tab == 'dist':
            self.dist_add_row(sender, args)
        else:
            self.add_row(sender, args)

    def _delete_rows_dispatch(self, sender, args):
        if self._current_tab == 'dist':
            self.dist_delete_rows(sender, args)
        else:
            self.delete_rows(sender, args)

    # ── Tab switching ─────────────────────────────────────────────────

    def _show_tab(self, tab_name):
        self._current_tab = tab_name
        import System.Windows.Media as M
        from System.Windows import Visibility as V
        green = M.SolidColorBrush(M.Color.FromRgb(0x20, 0x8A, 0x3C))
        grey  = M.SolidColorBrush(M.Color.FromRgb(0x40, 0x45, 0x53))

        # Switch which grid is visible
        for attr, active in [('recipients_grid', tab_name == 'client'),
                              ('dist_grid',       tab_name == 'dist')]:
            el = getattr(self, attr, None)
            if el: el.Visibility = V.Visible if active else V.Collapsed

        # Colour tab buttons
        for btn_attr, active_tab in [('rec_client_tab_btn', 'client'),
                                     ('rec_dist_tab_btn',   'dist')]:
            btn = getattr(self, btn_attr, None)
            if btn:
                try: btn.Background = green if tab_name == active_tab else grey
                except: pass

        # Re-label add/delete buttons to context and reset delete enabled state
        add_btn = getattr(self, 'rec_add_btn', None)
        del_btn = getattr(self, 'rec_delete_btn', None)
        if add_btn: add_btn.Content = "+ Add"
        if del_btn: del_btn.IsEnabled = False

    def switch_to_client_tab(self, sender, args): self._show_tab('client')
    def switch_to_dist_tab(self,   sender, args): self._show_tab('dist')

    # ── Drag-and-drop ─────────────────────────────────────────────────

    _drag_source_row  = None
    _drag_start_pos   = None
    _drag_source_grid = None
    _drop_target_idx  = -1
    _drop_popups      = {}

    def _setup_drag_drop(self):
        self._drop_popups = {}
        for grid_name in ['recipients_grid', 'dist_grid']:
            grid = getattr(self, grid_name, None)
            if grid is None:
                continue
            grid.AllowDrop = True
            grid.PreviewMouseLeftButtonDown += self._drag_mouse_down
            grid.PreviewMouseMove           += self._drag_mouse_move
            grid.DragOver                   += self._drag_over
            grid.Drop                       += self._drag_drop
            grid.DragLeave                  += self._drag_leave
            self._drop_popups[grid_name] = self._make_indicator_popup(grid)

    def _make_indicator_popup(self, grid):
        try:
            from System.Windows.Controls.Primitives import Popup
            from System.Windows.Controls import Border
            from System.Windows.Media import SolidColorBrush, Color
            import System.Windows
            bar = Border()
            bar.Height = 2; bar.Width = 400
            bar.Background = SolidColorBrush(Color.FromRgb(0x20, 0x8A, 0x3C))
            bar.IsHitTestVisible = False
            popup = Popup()
            popup.Child = bar
            popup.PlacementTarget = grid
            popup.Placement = System.Windows.Controls.Primitives.PlacementMode.Relative
            popup.AllowsTransparency = True
            popup.IsOpen = False
            return popup
        except: return None

    def _grid_name_of(self, grid):
        for name in ['recipients_grid', 'dist_grid']:
            if getattr(self, name, None) is grid:
                return name
        return None

    def _current_collection(self):
        return self.dist_data if self._current_tab == 'dist' else self.data

    def _get_row_at_point(self, grid, point):
        try:
            from System.Windows.Media import VisualTreeHelper
            el = grid.InputHitTest(point)
            while el is not None:
                if hasattr(el, 'DataContext') and el.DataContext in grid.ItemsSource:
                    idx = list(grid.ItemsSource).index(el.DataContext)
                    return el.DataContext, idx
                try:   el = VisualTreeHelper.GetParent(el)
                except: break
        except: pass
        return None, -1

    def _drag_mouse_down(self, sender, args):
        try:
            self._drag_start_pos   = args.GetPosition(sender)
            row, _                 = self._get_row_at_point(sender, self._drag_start_pos)
            self._drag_source_row  = row
            self._drag_source_grid = sender
        except: pass

    def _drag_mouse_move(self, sender, args):
        if self._drag_source_row is None or self._drag_start_pos is None: return
        try:
            import System.Windows.Input as WI
            if args.LeftButton != WI.MouseButtonState.Pressed:
                self._drag_source_row = None; return
            pos = args.GetPosition(sender)
            if abs(pos.X - self._drag_start_pos.X) < 4 and abs(pos.Y - self._drag_start_pos.Y) < 4: return
            from System.Windows import DragDrop, DataObject, DragDropEffects
            tag = 'DistRow' if self._current_tab == 'dist' else 'RecRow'
            DragDrop.DoDragDrop(sender, DataObject(tag, self._drag_source_row), DragDropEffects.Move)
            self._drag_source_row = None
        except: pass

    def _drag_over(self, sender, args):
        try:
            from System.Windows import DragDropEffects
            args.Effects = DragDropEffects.Move; args.Handled = True
            col = self._current_collection()
            _, idx = self._get_row_at_point(sender, args.GetPosition(sender))
            self._drop_target_idx = idx if idx >= 0 else len(list(col))
            gname = self._grid_name_of(sender)
            popup = self._drop_popups.get(gname)
            if popup:
                rh  = sender.RowHeight if sender.RowHeight > 0 else 28
                chh = sender.ColumnHeaderHeight if sender.ColumnHeaderHeight > 0 else 32
                popup.Child.Width = sender.ActualWidth
                popup.HorizontalOffset = 0
                popup.VerticalOffset = chh + self._drop_target_idx * rh
                popup.IsOpen = True
        except: pass

    def _drag_leave(self, sender, args):
        try:
            for p in self._drop_popups.values():
                if p: p.IsOpen = False
        except: pass

    def _drag_drop(self, sender, args):
        try:
            for p in self._drop_popups.values():
                if p: p.IsOpen = False
        except: pass
        try:
            tag = 'DistRow' if self._current_tab == 'dist' else 'RecRow'
            src_row = args.Data.GetData(tag)
            if src_row is None: return
            col = self._current_collection()
            items = list(col)
            if src_row not in items: return
            src_idx  = items.index(src_row)
            dest_idx = self._drop_target_idx
            if dest_idx < 0: dest_idx = len(items)
            if dest_idx == src_idx or dest_idx == src_idx + 1: return
            col.Remove(src_row)
            if src_idx < dest_idx: dest_idx -= 1
            if dest_idx >= len(list(col)): col.Add(src_row)
            else: col.Insert(dest_idx, src_row)
            args.Handled = True
        except: pass

    # ── Column sort ───────────────────────────────────────────────────

    _last_sort = {}

    def grid_sorting(self, sender, args):
        try:
            args.Handled = True
            col = str(args.Column.Header).lower()
            items = list(self.data)
            if not items: return
            if 'company' in col or 'recipient' in col:
                key_fn = lambda r: (r.Company or '').lower()
            elif 'attention' in col:
                key_fn = lambda r: (r.AttentionTo or '').lower()
            else: return
            prev = self._last_sort.get(col)
            items.sort(key=key_fn, reverse=(prev == 'asc'))
            self._last_sort[col] = 'desc' if prev == 'asc' else 'asc'
            self.data.Clear()
            for item in items: self.data.Add(item)
        except: pass

    def dist_grid_sorting(self, sender, args):
        try:
            args.Handled = True
            col = str(args.Column.Header).lower()
            items = list(self.dist_data)
            if not items: return
            key_fn = lambda r: (r.Distribution or '').lower()
            prev = self._last_sort.get('dist_' + col)
            items.sort(key=key_fn, reverse=(prev == 'asc'))
            self._last_sort['dist_' + col] = 'desc' if prev == 'asc' else 'asc'
            self.dist_data.Clear()
            for item in items: self.dist_data.Add(item)
        except: pass

    # ── Export ────────────────────────────────────────────────────────

    def export_data(self, sender, args):
        if self._current_tab == 'dist':
            records = list(self.dist_data)
        else:
            grid = getattr(self, 'recipients_grid', None)
            selected = list(grid.SelectedItems) if (grid and grid.SelectedItems) else []
            if selected:
                choice = forms.CommandSwitchWindow.show(
                    ['Export All Records', 'Export Selected Records Only'],
                    message='What would you like to export?')
                records = selected if choice == 'Export Selected Records Only' else list(self.data)
            else:
                records = list(self.data)

        if not records:
            forms.alert("No records to export.", title="Export"); return

        opts = ['Excel (.xlsx)', 'CSV (.csv)']
        fmt  = forms.CommandSwitchWindow.show(opts, message='Select export format:')
        if not fmt: return

        dialog = SaveFileDialog()
        dialog.Title = "Export"
        if fmt.startswith('Excel'):
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"; dialog.DefaultExt = "xlsx"
        else:
            dialog.Filter = "CSV Files (*.csv)|*.csv"; dialog.DefaultExt = "csv"

        if dialog.ShowDialog() == DialogResult.OK:
            try:
                if self._current_tab == 'dist':
                    with open(dialog.FileName, 'w') as f:
                        w = csv.writer(f)
                        w.writerow(['Distribution'])
                        for r in records: w.writerow([r.Distribution])
                else:
                    (write_excel if fmt.startswith('Excel') else write_csv)(dialog.FileName, records)
                forms.alert("Exported {} record(s).".format(len(records)), title="Export Complete")
            except Exception as e:
                forms.alert("Export error:\n{}".format(str(e)), title="Export Error")

    # ── Browse / Import (Client List only) ────────────────────────────

    def browse_file(self, sender, args):
        dialog = OpenFileDialog()
        dialog.Title  = "Select Import File"
        dialog.Filter = "Supported Files (*.csv;*.xlsx;*.xls)|*.csv;*.xlsx;*.xls|CSV Files (*.csv)|*.csv|Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
        if dialog.ShowDialog() == DialogResult.OK:
            try:
                fp  = dialog.FileName
                ext = os.path.splitext(fp)[1].lower()
                if ext in ('.xlsx', '.xls'):
                    headers, data = read_excel(fp)
                elif ext == '.csv':
                    headers, data = read_csv(fp)
                else:
                    forms.alert("Unsupported format: {}".format(ext), title="Import Error"); return
                self._import_headers = headers
                self._import_data    = data
                self.rec_file_path_tb.Text = fp
                self._populate_column_mappings(headers)
                try: self.rec_preview_btn.IsEnabled = True
                except: pass
                try: self.rec_import_btn.IsEnabled  = True
                except: pass
            except Exception as e:
                forms.alert("Error reading file:\n{}".format(str(e)), title="Import Error")
                self._import_headers = []
                self._import_data    = []

    def _populate_column_mappings(self, headers):
        for cb in [self.rec_recipient_mapping_cb, self.rec_attention_mapping_cb]:
            cb.Items.Clear()
            cb.Items.Add("-- Select Column --")
            for h in headers:
                cb.Items.Add(h)
        rec_kw = ['company', 'recipient', 'role', 'organization', 'to']
        att_kw = ['attention', 'attn', 'person', 'contact', 'name']
        for i, h in enumerate(headers, start=1):
            hl = h.lower()
            if any(k in hl for k in rec_kw) and self.rec_recipient_mapping_cb.SelectedIndex <= 0:
                self.rec_recipient_mapping_cb.SelectedIndex = i
            if any(k in hl for k in att_kw) and self.rec_attention_mapping_cb.SelectedIndex <= 0:
                self.rec_attention_mapping_cb.SelectedIndex = i
        if self.rec_recipient_mapping_cb.SelectedIndex < 0:
            self.rec_recipient_mapping_cb.SelectedIndex = 0
        if self.rec_attention_mapping_cb.SelectedIndex < 0:
            self.rec_attention_mapping_cb.SelectedIndex = 0

    def preview_import(self, sender, args):
        if self.rec_recipient_mapping_cb.SelectedIndex <= 0:
            forms.alert("Please select the Company column.", title="Mapping Required"); return
        if self.rec_attention_mapping_cb.SelectedIndex <= 0:
            forms.alert("Please select the Attention To column.", title="Mapping Required"); return
        r_idx = self.rec_recipient_mapping_cb.SelectedIndex - 1
        a_idx = self.rec_attention_mapping_cb.SelectedIndex - 1
        preview = []
        for row in self._import_data[:100]:
            r = row[r_idx].strip() if r_idx < len(row) else ""
            a = row[a_idx].strip() if a_idx < len(row) else ""
            if r or a: preview.append((r, a))
        try:
            ImportPreviewWindow(preview).ShowDialog()
        except Exception as e:
            forms.alert("Preview error:\n{}".format(str(e)), title="Preview Error")

    def import_data(self, sender, args):
        if self.rec_recipient_mapping_cb.SelectedIndex <= 0:
            forms.alert("Please select the Company column.", title="Mapping Required"); return
        if self.rec_attention_mapping_cb.SelectedIndex <= 0:
            forms.alert("Please select the Attention To column.", title="Mapping Required"); return
        r_idx = self.rec_recipient_mapping_cb.SelectedIndex - 1
        a_idx = self.rec_attention_mapping_cb.SelectedIndex - 1
        added = skipped = 0
        for row in self._import_data:
            r = row[r_idx].strip() if r_idx < len(row) else ""
            a = row[a_idx].strip() if a_idx < len(row) else ""
            if not r and not a: continue
            if any(x.Company.lower() == r.lower() and x.AttentionTo.lower() == a.lower()
                   for x in self.data):
                skipped += 1; continue
            self.data.Add(RecipientRecord(company=r, attention_to=a))
            added += 1
        msg = "Imported: {} records".format(added)
        if skipped: msg += "\nSkipped (duplicates): {}".format(skipped)
        forms.alert(msg, title="Import Complete")
        self.rec_file_path_tb.Text = ""
        self.rec_recipient_mapping_cb.Items.Clear()
        self.rec_attention_mapping_cb.Items.Clear()
        try: self.rec_preview_btn.IsEnabled = False
        except: pass
        try: self.rec_import_btn.IsEnabled  = False
        except: pass
        self._import_headers = []
        self._import_data    = []

    # ── Client List grid operations ───────────────────────────────────

    def add_row(self, sender, args):
        new = RecipientRecord()
        self.data.Add(new)
        try:
            self.recipients_grid.ScrollIntoView(new)
            self.recipients_grid.SelectedItem = new
        except: pass

    def _grid_selected_items(self):
        grid = getattr(self, 'recipients_grid', None)
        if grid and grid.SelectedItems: return list(grid.SelectedItems)
        if grid and grid.SelectedItem:  return [grid.SelectedItem]
        return []

    def delete_rows(self, sender, args):
        to_remove = self._grid_selected_items()
        if not to_remove: return
        if forms.alert("Delete {} record(s)?".format(len(to_remove)),
                       title="Confirm Delete", ok=False, yes=True, no=True):
            for r in to_remove: self.data.Remove(r)
            try: self.rec_delete_btn.IsEnabled = False
            except: pass

    def grid_selection_changed(self, sender, args):
        grid = getattr(self, 'recipients_grid', None)
        has = grid is not None and grid.SelectedItems is not None and len(list(grid.SelectedItems)) > 0
        try: self.rec_delete_btn.IsEnabled = has
        except: pass
        self._refresh_select_menu()
        try:
            grid  = getattr(self, 'recipients_grid', None)
            total = len(list(self.data))
            sel   = len(list(grid.SelectedItems)) if (grid and grid.SelectedItems) else 0
            from System.Windows import Visibility as V
            sa = getattr(self, 'rec_ctx_select_all',  None)
            sn = getattr(self, 'rec_ctx_select_none', None)
            if sa: sa.Visibility = V.Collapsed if (sel == total and total > 0) else V.Visible
            if sn: sn.Visibility = V.Visible   if sel > 0                      else V.Collapsed
        except: pass

    def context_select_all(self, sender, args):
        try:
            grid = getattr(self, 'recipients_grid', None)
            if grid: grid.SelectAll()
        except: pass

    def context_select_none(self, sender, args):
        try:
            grid = getattr(self, 'recipients_grid', None)
            if grid: grid.UnselectAll()
        except: pass

    def context_duplicate(self, sender, args):
        grid = getattr(self, 'recipients_grid', None)
        if grid and grid.SelectedItem:
            s = grid.SelectedItem
            self.data.Add(RecipientRecord(company=s.Company, attention_to=s.AttentionTo))

    def context_delete(self, sender, args):
        to_remove = self._grid_selected_items()
        if not to_remove: return
        if forms.alert("Delete {} record(s)?".format(len(to_remove)),
                       title="Confirm Delete", ok=False, yes=True, no=True):
            for r in to_remove: self.data.Remove(r)

    def context_copy(self, sender, args):
        try:
            from System.Windows import Clipboard
            items = self._grid_selected_items()
            if items:
                Clipboard.SetText("\n".join(
                    "{}\t{}".format(r.Company, r.AttentionTo) for r in items))
        except: pass

    def clear_selections(self):
        for gname in ['recipients_grid', 'dist_grid']:
            try:
                grid = getattr(self, gname, None)
                if grid: grid.UnselectAll()
            except: pass

    # ── Distribution List grid operations ────────────────────────────

    def dist_add_row(self, sender, args):
        new = DistributionRecord()
        self.dist_data.Add(new)
        try:
            self.dist_grid.ScrollIntoView(new)
            self.dist_grid.SelectedItem = new
        except: pass

    def _dist_selected_items(self):
        grid = getattr(self, 'dist_grid', None)
        if grid and grid.SelectedItems: return list(grid.SelectedItems)
        if grid and grid.SelectedItem:  return [grid.SelectedItem]
        return []

    def dist_delete_rows(self, sender, args):
        to_remove = self._dist_selected_items()
        if not to_remove: return
        if forms.alert("Delete {} record(s)?".format(len(to_remove)),
                       title="Confirm Delete", ok=False, yes=True, no=True):
            for r in to_remove: self.dist_data.Remove(r)
            try: self.rec_delete_btn.IsEnabled = False
            except: pass

    def dist_grid_selection_changed(self, sender, args):
        grid = getattr(self, 'dist_grid', None)
        has = grid is not None and grid.SelectedItems is not None and len(list(grid.SelectedItems)) > 0
        try: self.rec_delete_btn.IsEnabled = has
        except: pass
        self._dist_refresh_select_menu()

    def _dist_refresh_select_menu(self):
        try:
            grid  = getattr(self, 'dist_grid', None)
            total = len(list(self.dist_data))
            sel   = len(list(grid.SelectedItems)) if (grid and grid.SelectedItems) else 0
            from System.Windows import Visibility as V
            sa = getattr(self, 'rec_dist_ctx_select_all',  None)
            sn = getattr(self, 'rec_dist_ctx_select_none', None)
            if sa: sa.Visibility = V.Collapsed if (sel == total and total > 0) else V.Visible
            if sn: sn.Visibility = V.Visible   if sel > 0                      else V.Collapsed
        except: pass

    def dist_context_select_all(self, sender, args):
        try:
            grid = getattr(self, 'dist_grid', None)
            if grid: grid.SelectAll()
        except: pass

    def dist_context_select_none(self, sender, args):
        try:
            grid = getattr(self, 'dist_grid', None)
            if grid: grid.UnselectAll()
        except: pass

    def dist_context_duplicate(self, sender, args):
        grid = getattr(self, 'dist_grid', None)
        if grid and grid.SelectedItem:
            self.dist_data.Add(DistributionRecord(distribution=grid.SelectedItem.Distribution))

    def dist_context_delete(self, sender, args):
        to_remove = self._dist_selected_items()
        if not to_remove: return
        if forms.alert("Delete {} record(s)?".format(len(to_remove)),
                       title="Confirm Delete", ok=False, yes=True, no=True):
            for r in to_remove: self.dist_data.Remove(r)

    def dist_context_copy(self, sender, args):
        try:
            from System.Windows import Clipboard
            items = self._dist_selected_items()
            if items:
                Clipboard.SetText("\n".join(r.Distribution for r in items))
        except: pass
