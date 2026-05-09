# -*- coding: utf-8 -*-
"""
Settings Manager for pyTransmit
================================
Manage Reason for Issue, Method of Issue, Document Format, and Print Size settings.
Located in: Settings Manager/SettingsManager.py

Author: pyTransmit Suite
Version: 1.1
"""

import os
import json
from pyrevit import forms
from pyrevit.forms import WPFWindow
import clr

clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")

from System.Windows import Visibility
from System.Windows.Forms import OpenFileDialog, SaveFileDialog, DialogResult
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

# xlsxwriter is bundled with pyRevit's IronPython environment — no install needed.

# Database paths
# Databases saved next to OptionsManager.py
DB_FOLDER = os.path.dirname(os.path.abspath(__file__))
REASON_DB = os.path.join(DB_FOLDER, 'reason.json')
METHOD_DB = os.path.join(DB_FOLDER, 'method.json')
FORMAT_DB = os.path.join(DB_FOLDER, 'format.json')
PRINTSIZE_DB = os.path.join(DB_FOLDER, 'printsize.json')

# ═══════════════════════════════════════════════════════════════════════════
# DATA MODELS
# ═══════════════════════════════════════════════════════════════════════════

class CodedRecord(INotifyPropertyChanged):
    """Record with Code, Separator, Description (for Reason and Method)"""
    def __init__(self, code="", separator="=", description="", collection=None):
        self._code = code
        self._separator = separator
        self._description = description
        self._collection = collection  # Reference to parent collection
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
    def Code(self):
        return self._code
    
    @Code.setter
    def Code(self, value):
        if self._code != value:
            self._code = value
            self._notify("Code")
    
    @property
    def Separator(self):
        return self._separator
    
    @Separator.setter
    def Separator(self, value):
        if self._separator != value:
            self._separator = value
            # Update all other records in the collection
            if self._collection:
                for record in self._collection:
                    if record != self:  # Don't update self
                        record._separator = value
                        record._notify("Separator")
            self._notify("Separator")
    
    @property
    def Description(self):
        return self._description
    
    @Description.setter
    def Description(self, value):
        if self._description != value:
            self._description = value
            self._notify("Description")
    
class SimpleRecord(INotifyPropertyChanged):
    """Record with just Value (for Format and Print Size)"""
    def __init__(self, value=""):
        self._value = value
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
    def Value(self):
        return self._value
    
    @Value.setter
    def Value(self, value):
        if self._value != value:
            self._value = value
            self._notify("Value")

# ═══════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ═══════════════════════════════════════════════════════════════════════════

class SettingsManagerWindow(WPFWindow):
    def __init__(self):
        xaml_file = os.path.join(os.path.dirname(__file__), 'SettingsManager.xaml')
        WPFWindow.__init__(self, xaml_file)
        
        # Initialize collections
        self.reason_data = ObservableCollection[CodedRecord]()
        self.method_data = ObservableCollection[CodedRecord]()
        self.format_data = ObservableCollection[SimpleRecord]()
        self.printsize_data = ObservableCollection[SimpleRecord]()
        
        # Bind grids
        self.reason_grid.ItemsSource = self.reason_data
        self.method_grid.ItemsSource = self.method_data
        self.format_grid.ItemsSource = self.format_data
        self.printsize_grid.ItemsSource = self.printsize_data
        
        # Load data
        self.load_all_data()
        
        # Set initial tab
        self.current_tab = "reason"
        self._update_record_count()
    
    # ═══════════════════════════════════════════════════════════════════════
    # DATA LOADING/SAVING
    # ═══════════════════════════════════════════════════════════════════════
    
    def load_all_data(self):
        """Load all databases"""
        # Reason
        try:
            with open(REASON_DB, 'r') as f:
                data = json.load(f)
                for item in data:
                    self.reason_data.Add(CodedRecord(
                        item.get('code', ''),
                        item.get('separator', '='),
                        item.get('description', ''),
                        self.reason_data
                    ))
        except:
            pass
        
        # Method
        try:
            with open(METHOD_DB, 'r') as f:
                data = json.load(f)
                for item in data:
                    self.method_data.Add(CodedRecord(
                        item.get('code', ''),
                        item.get('separator', '='),
                        item.get('description', ''),
                        self.method_data
                    ))
        except:
            pass
        
        # Format
        try:
            with open(FORMAT_DB, 'r') as f:
                data = json.load(f)
                for item in data:
                    self.format_data.Add(SimpleRecord(item.get('value', '')))
        except:
            pass
        
        # Print Size
        try:
            with open(PRINTSIZE_DB, 'r') as f:
                data = json.load(f)
                for item in data:
                    self.printsize_data.Add(SimpleRecord(item.get('value', '')))
        except:
            pass
    
    def save_reason_data(self):
        """Save reason database"""
        data = []
        for record in self.reason_data:
            data.append({
                'code': record.Code,
                'separator': record.Separator,
                'description': record.Description
            })
        with open(REASON_DB, 'w') as f:
            json.dump(data, f, indent=2)
    
    def save_method_data(self):
        """Save method database"""
        data = []
        for record in self.method_data:
            data.append({
                'code': record.Code,
                'separator': record.Separator,
                'description': record.Description
            })
        with open(METHOD_DB, 'w') as f:
            json.dump(data, f, indent=2)
    
    def save_format_data(self):
        """Save format database"""
        data = []
        for record in self.format_data:
            data.append({'value': record.Value})
        with open(FORMAT_DB, 'w') as f:
            json.dump(data, f, indent=2)
    
    def save_printsize_data(self):
        """Save print size database"""
        data = []
        for record in self.printsize_data:
            data.append({'value': record.Value})
        with open(PRINTSIZE_DB, 'w') as f:
            json.dump(data, f, indent=2)
    
    def _update_record_count(self):
        """Update record count display"""
        if self.current_tab == "reason":
            count = len(self.reason_data)
        elif self.current_tab == "method":
            count = len(self.method_data)
        elif self.current_tab == "format":
            count = len(self.format_data)
        else:
            count = len(self.printsize_data)
        
        self.record_count_tb.Text = "{} records".format(count)
    
    # ═══════════════════════════════════════════════════════════════════════
    # TAB SWITCHING
    # ═══════════════════════════════════════════════════════════════════════
    
    def switch_to_reason_tab(self, sender, args):
        self.reason_grid.Visibility = Visibility.Visible
        self.method_grid.Visibility = Visibility.Collapsed
        self.format_grid.Visibility = Visibility.Collapsed
        self.printsize_grid.Visibility = Visibility.Collapsed
        
        self.reason_tab_btn.Style = self.FindResource("ActiveTabButtonStyle")
        self.method_tab_btn.Style = self.FindResource("TabButtonStyle")
        self.format_tab_btn.Style = self.FindResource("TabButtonStyle")
        self.printsize_tab_btn.Style = self.FindResource("TabButtonStyle")
        
        self.current_tab = "reason"
        self._update_record_count()
    
    def switch_to_method_tab(self, sender, args):
        self.reason_grid.Visibility = Visibility.Collapsed
        self.method_grid.Visibility = Visibility.Visible
        self.format_grid.Visibility = Visibility.Collapsed
        self.printsize_grid.Visibility = Visibility.Collapsed
        
        self.reason_tab_btn.Style = self.FindResource("TabButtonStyle")
        self.method_tab_btn.Style = self.FindResource("ActiveTabButtonStyle")
        self.format_tab_btn.Style = self.FindResource("TabButtonStyle")
        self.printsize_tab_btn.Style = self.FindResource("TabButtonStyle")
        
        self.current_tab = "method"
        self._update_record_count()
    
    def switch_to_format_tab(self, sender, args):
        self.reason_grid.Visibility = Visibility.Collapsed
        self.method_grid.Visibility = Visibility.Collapsed
        self.format_grid.Visibility = Visibility.Visible
        self.printsize_grid.Visibility = Visibility.Collapsed
        
        self.reason_tab_btn.Style = self.FindResource("TabButtonStyle")
        self.method_tab_btn.Style = self.FindResource("TabButtonStyle")
        self.format_tab_btn.Style = self.FindResource("ActiveTabButtonStyle")
        self.printsize_tab_btn.Style = self.FindResource("TabButtonStyle")
        
        self.current_tab = "format"
        self._update_record_count()
    
    def switch_to_printsize_tab(self, sender, args):
        self.reason_grid.Visibility = Visibility.Collapsed
        self.method_grid.Visibility = Visibility.Collapsed
        self.format_grid.Visibility = Visibility.Collapsed
        self.printsize_grid.Visibility = Visibility.Visible
        
        self.reason_tab_btn.Style = self.FindResource("TabButtonStyle")
        self.method_tab_btn.Style = self.FindResource("TabButtonStyle")
        self.format_tab_btn.Style = self.FindResource("TabButtonStyle")
        self.printsize_tab_btn.Style = self.FindResource("ActiveTabButtonStyle")
        
        self.current_tab = "printsize"
        self._update_record_count()
    
    # ═══════════════════════════════════════════════════════════════════════
    # ADD/DELETE
    # ═══════════════════════════════════════════════════════════════════════
    
    def add_row(self, sender, args):
        """Add new row to current tab"""
        if self.current_tab == "reason":
            self.reason_data.Add(CodedRecord("", "=", "", self.reason_data))
            self.save_reason_data()
        elif self.current_tab == "method":
            self.method_data.Add(CodedRecord("", "=", "", self.method_data))
            self.save_method_data()
        elif self.current_tab == "format":
            self.format_data.Add(SimpleRecord(""))
            self.save_format_data()
        else:
            self.printsize_data.Add(SimpleRecord(""))
            self.save_printsize_data()
        
        self._update_record_count()
    
    def delete_rows(self, sender, args):
        """Delete grid-highlighted rows from current tab"""
        grid_map = {
            'reason':    ('reason_grid',    self.reason_data,    self.save_reason_data),
            'method':    ('method_grid',    self.method_data,    self.save_method_data),
            'format':    ('format_grid',    self.format_data,    self.save_format_data),
            'printsize': ('printsize_grid', self.printsize_data, self.save_printsize_data),
        }
        grid_name, data, save_fn = grid_map[self.current_tab]
        grid = getattr(self, grid_name, None)
        to_remove = list(grid.SelectedItems) if (grid and grid.SelectedItems) else []
        if to_remove:
            for r in to_remove:
                data.Remove(r)
            save_fn()
        self._update_record_count()
    
    # ═══════════════════════════════════════════════════════════════════════
    # IMPORT/EXPORT
    # ═══════════════════════════════════════════════════════════════════════
    
    def browse_file(self, sender, args):
        """Browse for import file"""
        dialog = OpenFileDialog()
        dialog.Filter = "Excel/CSV Files (*.xlsx;*.csv)|*.xlsx;*.csv|Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv"
        dialog.Title = "Select Import File"
        if dialog.ShowDialog() == DialogResult.OK:
            self.file_path_tb.Text = dialog.FileName
            self.import_btn.IsEnabled = True
    
    def import_data(self, sender, args):
        """Import data from Excel or CSV file"""
        file_path = self.file_path_tb.Text
        if not os.path.exists(file_path):
            forms.alert("File not found!", title="Import Error")
            return
        
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext == '.csv':
            self._import_csv(file_path)
        elif ext in ['.xlsx', '.xls']:
            self._import_excel(file_path)
        else:
            forms.alert("Unsupported file format", title="Import Error")
    
    def _import_csv(self, file_path):
        """Import from single CSV file with sections"""
        try:
            import csv
            
            with open(file_path, 'r') as f:
                reader = csv.reader(f)
                rows = list(reader)
            
            current_section = None
            imported_counts = {}
            
            for i, row in enumerate(rows):
                if not row or not row[0]:
                    continue
                
                # Check for section markers
                if row[0].startswith('[') and row[0].endswith(']'):
                    section_name = row[0][1:-1]  # Remove [ ]
                    if section_name in ['Reason', 'Method', 'Format', 'Print Size']:
                        current_section = section_name
                        imported_counts[current_section] = 0
                    continue
                
                # Skip header rows (contain column names)
                if current_section and i > 0:
                    prev_row = rows[i-1] if i > 0 else []
                    if prev_row and prev_row[0].startswith('['):
                        # This is a header row, skip it
                        continue
                    
                    # Import data based on section
                    if current_section == 'Reason' and len(row) >= 3:
                        if row[0] or row[2]:  # Skip if both code and description are empty
                            self.reason_data.Add(CodedRecord(row[0], row[1] if len(row) > 1 else "=", row[2], self.reason_data))
                            imported_counts['Reason'] += 1
                    
                    elif current_section == 'Method' and len(row) >= 3:
                        if row[0] or row[2]:
                            self.method_data.Add(CodedRecord(row[0], row[1] if len(row) > 1 else "=", row[2], self.method_data))
                            imported_counts['Method'] += 1
                    
                    elif current_section == 'Format' and len(row) >= 1:
                        if row[0]:
                            self.format_data.Add(SimpleRecord(row[0]))
                            imported_counts['Format'] += 1
                    
                    elif current_section == 'Print Size' and len(row) >= 1:
                        if row[0]:
                            self.printsize_data.Add(SimpleRecord(row[0]))
                            imported_counts['Print Size'] += 1
            
            # Save all data
            self.save_reason_data()
            self.save_method_data()
            self.save_format_data()
            self.save_printsize_data()
            
            self._update_record_count()
            
        except Exception as e:
            forms.alert("Error importing CSV:\n{}".format(str(e)), title="Import Error")
    
    def _import_excel(self, file_path):
        """Import data from Excel file with multiple sheets"""
        try:
            # Read Excel using COM
            from System import Type, Activator, Array
            from System.Runtime.InteropServices import Marshal
            import System.Reflection as Reflection
            
            excel = None
            workbook = None
            
            try:
                excel_type = Type.GetTypeFromProgID("Excel.Application")
                excel = Activator.CreateInstance(excel_type)
                
                excel_type.InvokeMember("Visible", Reflection.BindingFlags.SetProperty, None, excel, Array[object]([False]))
                excel_type.InvokeMember("DisplayAlerts", Reflection.BindingFlags.SetProperty, None, excel, Array[object]([False]))
                
                workbooks = excel_type.InvokeMember("Workbooks", Reflection.BindingFlags.GetProperty, None, excel, None)
                workbooks_type = workbooks.GetType()
                workbook = workbooks_type.InvokeMember("Open", Reflection.BindingFlags.InvokeMethod, None, workbooks, Array[object]([file_path]))
                
                workbook_type = workbook.GetType()
                worksheets = workbook_type.InvokeMember("Worksheets", Reflection.BindingFlags.GetProperty, None, workbook, None)
                worksheets_type = worksheets.GetType()
                
                # Import each sheet
                imported_counts = {}
                
                # Try to find sheets by name
                for sheet_name in ["Reason", "Method", "Format", "Print Size"]:
                    try:
                        worksheet = worksheets_type.InvokeMember("Item", Reflection.BindingFlags.GetProperty, None, worksheets, Array[object]([sheet_name]))
                        count = self._import_sheet(worksheet, sheet_name)
                        if count > 0:
                            imported_counts[sheet_name] = count
                        Marshal.ReleaseComObject(worksheet)
                    except:
                        pass
                
                # Close workbook
                workbook_type.InvokeMember("Close", Reflection.BindingFlags.InvokeMethod, None, workbook, Array[object]([False]))
                Marshal.ReleaseComObject(worksheets)
                Marshal.ReleaseComObject(workbook)
                Marshal.ReleaseComObject(workbooks)
                
                excel_type.InvokeMember("Quit", Reflection.BindingFlags.InvokeMethod, None, excel, None)
                Marshal.ReleaseComObject(excel)
                
                self._update_record_count()
                
            except Exception as e:
                if workbook:
                    try:
                        workbook_type = workbook.GetType()
                        workbook_type.InvokeMember("Close", Reflection.BindingFlags.InvokeMethod, None, workbook, Array[object]([False]))
                        Marshal.ReleaseComObject(workbook)
                    except:
                        pass
                if excel:
                    try:
                        excel_type = excel.GetType()
                        excel_type.InvokeMember("Quit", Reflection.BindingFlags.InvokeMethod, None, excel, None)
                        Marshal.ReleaseComObject(excel)
                    except:
                        pass
                raise
                
        except Exception as e:
            forms.alert("Error importing Excel:\n{}".format(str(e)), title="Import Error")
    
    def _import_sheet(self, worksheet, sheet_name):
        """Import data from a worksheet"""
        from System import Type, Array
        from System.Runtime.InteropServices import Marshal
        import System.Reflection as Reflection
        
        worksheet_type = worksheet.GetType()
        used_range = worksheet_type.InvokeMember("UsedRange", Reflection.BindingFlags.GetProperty, None, worksheet, None)
        
        used_range_type = used_range.GetType()
        rows = used_range_type.InvokeMember("Rows", Reflection.BindingFlags.GetProperty, None, used_range, None)
        rows_type = rows.GetType()
        rows_count = rows_type.InvokeMember("Count", Reflection.BindingFlags.GetProperty, None, rows, None)
        
        columns = used_range_type.InvokeMember("Columns", Reflection.BindingFlags.GetProperty, None, used_range, None)
        columns_type = columns.GetType()
        cols_count = columns_type.InvokeMember("Count", Reflection.BindingFlags.GetProperty, None, columns, None)
        
        cells = worksheet_type.InvokeMember("Cells", Reflection.BindingFlags.GetProperty, None, worksheet, None)
        cells_type = cells.GetType()
        
        # Read all data
        all_data = []
        for row in range(1, rows_count + 1):
            row_data = []
            for col in range(1, cols_count + 1):
                cell = cells_type.InvokeMember("Item", Reflection.BindingFlags.GetProperty, None, cells, Array[object]([row, col]))
                cell_type = cell.GetType()
                cell_value = cell_type.InvokeMember("Value2", Reflection.BindingFlags.GetProperty, None, cell, None)
                row_data.append(str(cell_value) if cell_value else "")
                Marshal.ReleaseComObject(cell)
            all_data.append(row_data)
        
        Marshal.ReleaseComObject(cells)
        Marshal.ReleaseComObject(columns)
        Marshal.ReleaseComObject(rows)
        Marshal.ReleaseComObject(used_range)
        
        # Import based on sheet name (skip header row)
        count = 0
        if sheet_name == "Reason":
            for row in all_data[1:]:  # Skip header
                if len(row) >= 3 and (row[0] or row[2]):
                    self.reason_data.Add(CodedRecord(row[0], row[1] if len(row) > 1 else "=", row[2], self.reason_data))
                    count += 1
            self.save_reason_data()
        elif sheet_name == "Method":
            for row in all_data[1:]:
                if len(row) >= 3 and (row[0] or row[2]):
                    self.method_data.Add(CodedRecord(row[0], row[1] if len(row) > 1 else "=", row[2], self.method_data))
                    count += 1
            self.save_method_data()
        elif sheet_name == "Format":
            for row in all_data[1:]:
                if row and row[0]:
                    self.format_data.Add(SimpleRecord(row[0]))
                    count += 1
            self.save_format_data()
        elif sheet_name == "Print Size":
            for row in all_data[1:]:
                if row and row[0]:
                    self.printsize_data.Add(SimpleRecord(row[0]))
                    count += 1
            self.save_printsize_data()
        
        return count
    
    def export_data(self, sender, args):
        """Export all tabs to Excel file with multiple sheets or CSV file"""
        # Ask user for format
        from pyrevit import forms as pyrevit_forms
        
        options = ['Excel (.xlsx) - All tabs in one file', 'CSV - Single file with sections']
        choice = pyrevit_forms.CommandSwitchWindow.show(options, message='Select export format:')
        
        if not choice:
            return
        
        if choice.startswith('Excel'):
            self._export_excel()
        else:
            self._export_csv()
    
    def _export_excel(self):
        """Export all tabs to Excel with multiple sheets"""
        import xlsxwriter
        dialog = SaveFileDialog()
        dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dialog.FileName = "pyTransmit_Settings.xlsx"

        if dialog.ShowDialog() != DialogResult.OK:
            return

        try:
            workbook = xlsxwriter.Workbook(dialog.FileName)
            header_format = workbook.add_format({
                'bold': True, 'bg_color': '#208A3C',
                'font_color': 'white', 'border': 1
            })

            ws = workbook.add_worksheet('Reason')
            ws.write(0, 0, 'Code', header_format)
            ws.write(0, 1, 'Separator', header_format)
            ws.write(0, 2, 'Reason for Issue', header_format)
            for i, r in enumerate(self.reason_data, 1):
                ws.write(i, 0, r.Code); ws.write(i, 1, r.Separator); ws.write(i, 2, r.Description)

            ws = workbook.add_worksheet('Method')
            ws.write(0, 0, 'Code', header_format)
            ws.write(0, 1, 'Separator', header_format)
            ws.write(0, 2, 'Method of Issue', header_format)
            for i, r in enumerate(self.method_data, 1):
                ws.write(i, 0, r.Code); ws.write(i, 1, r.Separator); ws.write(i, 2, r.Description)

            ws = workbook.add_worksheet('Format')
            ws.write(0, 0, 'Document Format', header_format)
            for i, r in enumerate(self.format_data, 1):
                ws.write(i, 0, r.Value)

            ws = workbook.add_worksheet('Print Size')
            ws.write(0, 0, 'Print Size', header_format)
            for i, r in enumerate(self.printsize_data, 1):
                ws.write(i, 0, r.Value)

            workbook.close()
            forms.alert("Exported successfully!\n\n4 sheets: Reason, Method, Format, Print Size",
                        title="Export Success")
        except Exception as e:
            forms.alert("Error exporting:\n{}".format(str(e)), title="Export Error")
    
    def _export_csv(self):
        """Export all tabs as single CSV file with sections"""
        dialog = SaveFileDialog()
        dialog.Filter = "CSV Files (*.csv)|*.csv"
        dialog.FileName = "pyTransmit_Settings.csv"
        
        if dialog.ShowDialog() != DialogResult.OK:
            return
        
        try:
            import csv
            
            with open(dialog.FileName, 'w') as f:
                writer = csv.writer(f)
                
                # Reason section
                writer.writerow(['[Reason]'])
                writer.writerow(['Code', 'Separator', 'Reason for Issue'])
                for r in self.reason_data:
                    writer.writerow([r.Code, r.Separator, r.Description])
                writer.writerow([])  # Blank line
                
                # Method section
                writer.writerow(['[Method]'])
                writer.writerow(['Code', 'Separator', 'Method of Issue'])
                for r in self.method_data:
                    writer.writerow([r.Code, r.Separator, r.Description])
                writer.writerow([])  # Blank line
                
                # Format section
                writer.writerow(['[Format]'])
                writer.writerow(['Document Format'])
                for r in self.format_data:
                    writer.writerow([r.Value])
                writer.writerow([])  # Blank line
                
                # Print Size section
                writer.writerow(['[Print Size]'])
                writer.writerow(['Print Size'])
                for r in self.printsize_data:
                    writer.writerow([r.Value])
            
            forms.alert("Exported successfully to:\n{}".format(dialog.FileName), title="Export Success")
            
        except Exception as e:
            forms.alert("Error exporting:\n{}".format(str(e)), title="Export Error")

def main():
    try:
        window = SettingsManagerWindow()
        window.ShowDialog()
    except Exception as e:
        forms.alert("Error: {}".format(str(e)), exitscript=True)

if __name__ == "__main__":
    main()


# ═══════════════════════════════════════════════════════════════════════════
# PANEL CONTROLLER  — embedded use inside pyTransmit main window
# ═══════════════════════════════════════════════════════════════════════════

class OptionsSettingsController(object):
    """
    Drives the Options panel when embedded inside pyTransmit.

    script.py usage:
        import sys
        sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'Options Manager'))
        from OptionsManager import OptionsPanelController
        self.opt_ctrl = OptionsPanelController()
        self.opt_ctrl.attach(panel_root)   # called after XamlReader loads the panel
    """

    DEFAULT_REASON    = [("A","=","Issued for Approval"),("B","=","Issued for Construction"),
                         ("C","=","Issued for Coordination"),("D","=","Issued for Information"),
                         ("E","=","Issued for Review")]
    DEFAULT_METHOD    = [("E","=","Email"),("P","=","Post"),("H","=","By Hand"),("U","=","Uploaded to Portal")]
    DEFAULT_FORMAT    = ["PDF","DWG","RVT","IFC"]
    DEFAULT_PRINTSIZE = ["A0","A1","A2","A3","A4"]

    def __init__(self):
        self.current_tab    = "reason"
        self.reason_data    = ObservableCollection[CodedRecord]()
        self.method_data    = ObservableCollection[CodedRecord]()
        self.format_data    = ObservableCollection[SimpleRecord]()
        self.printsize_data = ObservableCollection[SimpleRecord]()
        self._snapshot      = {}   # set by _take_snapshot() when panel opens
        self._load_all()

    def _take_snapshot(self):
        """Capture current data so discard() can restore it."""
        self._snapshot = {
            'reason':    [(r.Code, r.Separator, r.Description) for r in self.reason_data],
            'method':    [(r.Code, r.Separator, r.Description) for r in self.method_data],
            'format':    [r.Value for r in self.format_data],
            'printsize': [r.Value for r in self.printsize_data],
        }

    def discard(self):
        """Restore all four collections to the snapshot captured at panel open."""
        snap = self._snapshot
        if not snap:
            return
        self.reason_data.Clear()
        for code, sep, desc in snap.get('reason', []):
            self.reason_data.Add(CodedRecord(code, sep, desc, self.reason_data))
        self.method_data.Clear()
        for code, sep, desc in snap.get('method', []):
            self.method_data.Add(CodedRecord(code, sep, desc, self.method_data))
        self.format_data.Clear()
        for val in snap.get('format', []):
            self.format_data.Add(SimpleRecord(val))
        self.printsize_data.Clear()
        for val in snap.get('printsize', []):
            self.printsize_data.Add(SimpleRecord(val))

    def _load_all(self):
        def load_coded(path, data, default):
            try:
                with open(path, 'r') as f:
                    rows = json.load(f)
                for item in rows:
                    data.Add(CodedRecord(item.get('code',''), item.get('separator','='),
                                         item.get('description',''), data))
            except:
                for c, s, d in default:
                    data.Add(CodedRecord(c, s, d, data))

        def load_simple(path, data, default):
            try:
                with open(path, 'r') as f:
                    rows = json.load(f)
                for item in rows:
                    data.Add(SimpleRecord(item.get('value', '')))
            except:
                for v in default:
                    data.Add(SimpleRecord(v))

        load_coded(REASON_DB,    self.reason_data,    self.DEFAULT_REASON)
        load_coded(METHOD_DB,    self.method_data,    self.DEFAULT_METHOD)
        load_simple(FORMAT_DB,   self.format_data,    self.DEFAULT_FORMAT)
        load_simple(PRINTSIZE_DB, self.printsize_data, self.DEFAULT_PRINTSIZE)

    # ── Attach ────────────────────────────────────────────────────────

    def attach(self, root):
        """Walk element tree, register named elements, wire events, bind grids."""
        self._walk(root)
        self._wire_events()
        self._bind_grids()
        self._show_tab("reason")

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
        try:                            # ContextMenu items
            cm = element.ContextMenu
            if cm:
                for item in cm.Items:
                    self._walk(item)
        except: pass

    def _wire_events(self):
        """Bind all button and grid events to handler methods."""
        def bind(name, event, handler):
            el = getattr(self, name, None)
            if el is not None:
                try: getattr(el, event).__iadd__(handler)
                except: pass

        bind('opt_browse_file_btn',   'Click',   self.browse_file)
        bind('opt_import_btn',        'Click',   self.import_data)
        bind('opt_reason_tab_btn',    'Click',   self.switch_to_reason_tab)
        bind('opt_method_tab_btn',    'Click',   self.switch_to_method_tab)
        bind('opt_format_tab_btn',    'Click',   self.switch_to_format_tab)
        bind('opt_printsize_tab_btn', 'Click',   self.switch_to_printsize_tab)
        bind('opt_add_btn',           'Click',   self.add_row)
        bind('opt_delete_btn',        'Click',   self.delete_rows)
        for grid_name in ['opt_reason_grid', 'opt_method_grid',
                          'opt_format_grid', 'opt_printsize_grid']:
            bind(grid_name, 'SelectionChanged',           self.grid_selection_changed)
            bind(grid_name, 'Sorting',                    self.grid_sorting)
            bind(grid_name, 'PreviewMouseLeftButtonDown', self._header_click_check)
        for prefix in ['opt_reason', 'opt_method', 'opt_format', 'opt_printsize']:
            bind(prefix + '_ctx_select_all',  'Click', self.context_select_all)
            bind(prefix + '_ctx_select_none', 'Click', self.context_select_none)
            bind(prefix + '_ctx_duplicate',   'Click', self.context_duplicate)
            bind(prefix + '_ctx_delete',      'Click', self.context_delete)
            bind(prefix + '_ctx_copy',        'Click', self.context_copy)
        self._setup_drag_drop()

    def _setup_drag_drop(self):
        """Enable drag-and-drop row reordering on all four grids."""
        self._drop_popups = {}
        for grid_name in ['opt_reason_grid', 'opt_method_grid',
                          'opt_format_grid', 'opt_printsize_grid']:
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

    _drag_source_row  = None
    _drag_start_pos   = None
    _drag_source_grid = None
    _drop_target_idx  = -1
    _drop_popups      = {}

    def _make_indicator_popup(self, grid):
        """Create a thin green Popup used as the drop-position line."""
        try:
            from System.Windows.Controls.Primitives import Popup
            from System.Windows.Controls import Border
            from System.Windows.Media import SolidColorBrush, Color
            import System.Windows
            bar = Border()
            bar.Height     = 2
            bar.Width      = 400
            bar.Background = SolidColorBrush(Color.FromRgb(0x20, 0x8A, 0x3C))
            bar.IsHitTestVisible = False
            popup = Popup()
            popup.Child              = bar
            popup.PlacementTarget    = grid
            popup.Placement          = System.Windows.Controls.Primitives.PlacementMode.Relative
            popup.AllowsTransparency = True
            popup.IsOpen             = False
            return popup
        except:
            return None

    def _show_indicator(self, grid, grid_name, idx):
        try:
            popup = self._drop_popups.get(grid_name)
            if popup is None:
                return
            rh  = grid.RowHeight if grid.RowHeight > 0 else 28
            chh = grid.ColumnHeaderHeight if grid.ColumnHeaderHeight > 0 else 32
            popup.Child.Width         = grid.ActualWidth
            popup.HorizontalOffset    = 0
            popup.VerticalOffset      = chh + idx * rh
            popup.IsOpen              = True
        except: pass

    def _hide_all_indicators(self):
        try:
            for popup in self._drop_popups.values():
                if popup is not None:
                    popup.IsOpen = False
        except: pass

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

    def _grid_name(self, grid):
        for name in ['opt_reason_grid', 'opt_method_grid',
                     'opt_format_grid', 'opt_printsize_grid']:
            if getattr(self, name, None) is grid:
                return name
        return None

    def _drag_mouse_down(self, sender, args):
        try:
            self._drag_start_pos   = args.GetPosition(sender)
            row, _                 = self._get_row_at_point(sender, self._drag_start_pos)
            self._drag_source_row  = row
            self._drag_source_grid = sender
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
            # Prevent DataGrid from also processing this mouse-move as a selection drag
            args.Handled = True
            from System.Windows import DragDrop, DataObject, DragDropEffects
            DragDrop.DoDragDrop(sender,
                                DataObject("OptRow", self._drag_source_row),
                                DragDropEffects.Move)
            self._drag_source_row = None
        except: pass

    def _drag_over(self, sender, args):
        try:
            from System.Windows import DragDropEffects
            args.Effects = DragDropEffects.Move
            args.Handled = True
            _, idx   = self._get_row_at_point(sender, args.GetPosition(sender))
            data     = self._current_data()
            self._drop_target_idx = idx if idx >= 0 else len(list(data))
            gname = self._grid_name(sender)
            self._show_indicator(sender, gname, self._drop_target_idx)
        except: pass

    def _drag_leave(self, sender, args):
        self._hide_all_indicators()

    def _drag_drop(self, sender, args):
        """Reorder the ObservableCollection and persist."""
        self._hide_all_indicators()
        try:
            src_row = args.Data.GetData("OptRow")
            if src_row is None:
                return
            data  = self._current_data()
            items = list(data)
            if src_row not in items:
                return
            src_idx  = items.index(src_row)
            dest_idx = self._drop_target_idx
            if dest_idx < 0:
                dest_idx = len(items)
            if dest_idx == src_idx or dest_idx == src_idx + 1:
                return
            data.Remove(src_row)
            if src_idx < dest_idx:
                dest_idx -= 1
            if dest_idx >= len(list(data)):
                data.Add(src_row)
            else:
                data.Insert(dest_idx, src_row)
            args.Handled = True
        except: pass

    # ── Bind grids ────────────────────────────────────────────────────

    def _bind_grids(self):
        for attr, data in [('opt_reason_grid',    self.reason_data),
                           ('opt_method_grid',    self.method_data),
                           ('opt_format_grid',    self.format_data),
                           ('opt_printsize_grid', self.printsize_data)]:
            el = getattr(self, attr, None)
            if el is not None:
                el.ItemsSource = data

    # ── Save ──────────────────────────────────────────────────────────

    def save_all(self):
        self._save_tab('reason')
        self._save_tab('method')
        self._save_tab('format')
        self._save_tab('printsize')

    def _save_tab(self, tab):
        try:
            if tab == 'reason':
                data = [{'code': r.Code, 'separator': r.Separator, 'description': r.Description}
                        for r in self.reason_data]
                with open(REASON_DB, 'w') as f: json.dump(data, f, indent=2)
            elif tab == 'method':
                data = [{'code': r.Code, 'separator': r.Separator, 'description': r.Description}
                        for r in self.method_data]
                with open(METHOD_DB, 'w') as f: json.dump(data, f, indent=2)
            elif tab == 'format':
                data = [{'value': r.Value} for r in self.format_data]
                with open(FORMAT_DB, 'w') as f: json.dump(data, f, indent=2)
            elif tab == 'printsize':
                data = [{'value': r.Value} for r in self.printsize_data]
                with open(PRINTSIZE_DB, 'w') as f: json.dump(data, f, indent=2)
        except: pass

    # ── Tab switching ─────────────────────────────────────────────────

    def _show_tab(self, tab_name):
        self.current_tab = tab_name
        import System.Windows.Media as M
        from System.Windows import Visibility
        green = M.SolidColorBrush(M.Color.FromRgb(0x20, 0x8A, 0x3C))
        grey  = M.SolidColorBrush(M.Color.FromRgb(0x40, 0x45, 0x53))
        tabs  = {'reason':    ('opt_reason_grid',    'opt_reason_tab_btn'),
                 'method':    ('opt_method_grid',    'opt_method_tab_btn'),
                 'format':    ('opt_format_grid',    'opt_format_tab_btn'),
                 'printsize': ('opt_printsize_grid', 'opt_printsize_tab_btn')}
        for name, (grid_attr, btn_attr) in tabs.items():
            grid = getattr(self, grid_attr, None)
            btn  = getattr(self, btn_attr,  None)
            active = (name == tab_name)
            if grid: grid.Visibility = Visibility.Visible if active else Visibility.Collapsed
            if btn:  btn.Background  = green if active else grey

    def switch_to_reason_tab(self,    sender, args): self._show_tab('reason')
    def switch_to_method_tab(self,    sender, args): self._show_tab('method')
    def switch_to_format_tab(self,    sender, args): self._show_tab('format')
    def switch_to_printsize_tab(self, sender, args): self._show_tab('printsize')

    # ── Current grid / data helpers ───────────────────────────────────

    def _current_data(self):
        return {'reason': self.reason_data, 'method': self.method_data,
                'format': self.format_data, 'printsize': self.printsize_data
               }[self.current_tab]

    def _current_grid(self):
        return getattr(self, {'reason':    'opt_reason_grid',
                               'method':   'opt_method_grid',
                               'format':   'opt_format_grid',
                               'printsize':'opt_printsize_grid'}[self.current_tab], None)

    # ── Add / Delete ──────────────────────────────────────────────────

    def add_row(self, sender, args):
        if self.current_tab in ('reason', 'method'):
            new = CodedRecord('', '=', '', self._current_data())
        else:
            new = SimpleRecord('')
        self._current_data().Add(new)
        grid = self._current_grid()
        if grid:
            grid.ScrollIntoView(new)
            grid.SelectedItem = new

    def delete_rows(self, sender, args):
        to_remove = self._grid_selected_items()
        if not to_remove:
            return
        if forms.alert('Delete {} record(s)?'.format(len(to_remove)),
                       title='Confirm Delete', ok=False, yes=True, no=True):
            data = self._current_data()
            for r in to_remove:
                data.Remove(r)

    # ── Grid selection & context-menu Select All / Select None ────────

    def _grid_selected_items(self):
        """Returns grid-highlighted rows from the current grid."""
        grid = self._current_grid()
        if grid and grid.SelectedItems:
            return list(grid.SelectedItems)
        if grid and grid.SelectedItem:
            return [grid.SelectedItem]
        return []

    def grid_selection_changed(self, sender, args):
        """Update Select All/None menu items on selection change."""
        self._refresh_select_menu()

    def _refresh_select_menu(self):
        """Show Select All unless all rows highlighted; show Select None when any highlighted."""
        try:
            grid  = self._current_grid()
            data  = self._current_data()
            total = len(list(data))
            sel   = len(list(grid.SelectedItems)) if (grid and grid.SelectedItems) else 0
            from System.Windows import Visibility as V
            tab_prefix = {
                'reason':    'opt_reason',
                'method':    'opt_method',
                'format':    'opt_format',
                'printsize': 'opt_printsize',
            }[self.current_tab]
            sa = getattr(self, tab_prefix + '_ctx_select_all',  None)
            sn = getattr(self, tab_prefix + '_ctx_select_none', None)
            if sa: sa.Visibility = V.Collapsed if (sel == total and total > 0) else V.Visible
            if sn: sn.Visibility = V.Visible   if sel > 0                      else V.Collapsed
        except: pass

    def context_select_all(self, sender, args):
        try:
            grid = self._current_grid()
            if grid: grid.SelectAll()
        except: pass

    def context_select_none(self, sender, args):
        try:
            grid = self._current_grid()
            if grid: grid.UnselectAll()
        except: pass

    def clear_selections(self):
        """Clear grid highlights on all four grids. Called by script.py on panel close."""
        for gname in ['opt_reason_grid', 'opt_method_grid',
                      'opt_format_grid', 'opt_printsize_grid']:
            try:
                grid = getattr(self, gname, None)
                if grid: grid.UnselectAll()
            except: pass

    # ── Context menu ──────────────────────────────────────────────────

    def context_duplicate(self, sender, args):
        grid = self._current_grid()
        if grid and grid.SelectedItem:
            s = grid.SelectedItem
            if self.current_tab in ('reason', 'method'):
                new = CodedRecord(s.Code, s.Separator, s.Description, self._current_data())
            else:
                new = SimpleRecord(s.Value)
            self._current_data().Add(new)

    def context_delete(self, sender, args):
        to_remove = self._grid_selected_items()
        if not to_remove:
            return
        if forms.alert('Delete {} record(s)?'.format(len(to_remove)),
                       title='Confirm Delete', ok=False, yes=True, no=True):
            data = self._current_data()
            for r in to_remove:
                data.Remove(r)

    def context_copy(self, sender, args):
        try:
            from System.Windows import Clipboard
            items = self._grid_selected_items()
            if items:
                lines = []
                for r in items:
                    if hasattr(r, 'Code'):
                        lines.append('{}{}{}'.format(r.Code, r.Separator, r.Description))
                    else:
                        lines.append(r.Value)
                Clipboard.SetText('\n'.join(lines))
        except: pass

    # ── Export ────────────────────────────────────────────────────────

    def export_data(self, sender, args):
        options = ['Excel (.xlsx) - All tabs in one file', 'CSV - Single file with sections']
        choice  = forms.CommandSwitchWindow.show(options, message='Select export format:')
        if not choice:
            return
        if choice.startswith('Excel'):
            self._export_excel()
        else:
            self._export_csv()

    def _export_excel(self):
        import xlsxwriter as _xlw
        dialog = SaveFileDialog()
        dialog.Filter   = 'Excel Files (*.xlsx)|*.xlsx'
        dialog.FileName = 'pyTransmit_Options.xlsx'
        if dialog.ShowDialog() != DialogResult.OK:
            return
        try:
            workbook = _xlw.Workbook(dialog.FileName)
            bold = workbook.add_format({'bold': True, 'bg_color': '#208A3C',
                                        'font_color': 'white', 'border': 1})
            sheets = [('Reason',    self.reason_data,    ['Code','Separator','Reason for Issue']),
                      ('Method',    self.method_data,    ['Code','Separator','Method of Issue']),
                      ('Format',    self.format_data,    ['Document Format']),
                      ('PrintSize', self.printsize_data, ['Print Size'])]
            for title, data, headers in sheets:
                ws = workbook.add_worksheet(title)
                for col, h in enumerate(headers):
                    ws.write(0, col, h, bold)
                for row, r in enumerate(data, start=1):
                    if hasattr(r, 'Code'):
                        ws.write(row, 0, r.Code)
                        ws.write(row, 1, r.Separator)
                        ws.write(row, 2, r.Description)
                    else:
                        ws.write(row, 0, r.Value)
            workbook.close()
            forms.alert('Exported to:\n{}'.format(dialog.FileName), title='Export Complete')
        except Exception as e:
            forms.alert('Export error:\n{}'.format(str(e)), title='Export Error')

    def _export_csv(self):
        import csv as _csv
        dialog = SaveFileDialog()
        dialog.Filter   = 'CSV Files (*.csv)|*.csv'
        dialog.FileName = 'pyTransmit_Options.csv'
        if dialog.ShowDialog() != DialogResult.OK:
            return
        try:
            with open(dialog.FileName, 'w') as f:
                w = _csv.writer(f)
                w.writerow(['[Reason]'])
                w.writerow(['Code', 'Separator', 'Reason for Issue'])
                for r in self.reason_data:  w.writerow([r.Code, r.Separator, r.Description])
                w.writerow([])
                w.writerow(['[Method]'])
                w.writerow(['Code', 'Separator', 'Method of Issue'])
                for r in self.method_data:  w.writerow([r.Code, r.Separator, r.Description])
                w.writerow([])
                w.writerow(['[Format]'])
                w.writerow(['Document Format'])
                for r in self.format_data:  w.writerow([r.Value])
                w.writerow([])
                w.writerow(['[Print Size]'])
                w.writerow(['Print Size'])
                for r in self.printsize_data: w.writerow([r.Value])
            forms.alert('Exported to:\n{}'.format(dialog.FileName), title='Export Complete')
        except Exception as e:
            forms.alert('Export error:\n{}'.format(str(e)), title='Export Error')

    # ── Column-header sort ────────────────────────────────────────────

    _last_sort = {}

    def _header_click_check(self, sender, args):
        """Fires on every mouse-down on the grid. If the click landed on a
        DataGridColumnHeader, read its text and sort — no visual-tree
        walking required at wire-up time."""
        try:
            from System.Windows.Media import VisualTreeHelper
            from System.Windows.Controls.Primitives import DataGridColumnHeader
            el = args.OriginalSource
            while el is not None:
                if isinstance(el, DataGridColumnHeader):
                    col_text = ''
                    try:
                        col_text = str(el.Content)
                    except: pass
                    if col_text:
                        self._do_sort(col_text.lower())
                    return
                try:
                    el = VisualTreeHelper.GetParent(el)
                except:
                    break
        except: pass

    def grid_sorting(self, sender, args):
        """Suppress WPF built-in sort arrow animation; actual sort done in _header_click_check."""
        try:
            args.Handled = True
        except: pass

    def _do_sort(self, col_header):
        """Sort _current_data() by col_header, toggle asc/desc."""
        try:
            data  = self._current_data()
            items = list(data)
            if not items:
                return

            col_map = {
                'code':             lambda r: (getattr(r, 'Code',        '') or '').lower(),
                'separator':        lambda r: (getattr(r, 'Separator',   '') or '').lower(),
                'reason for issue': lambda r: (getattr(r, 'Description', '') or '').lower(),
                'method of issue':  lambda r: (getattr(r, 'Description', '') or '').lower(),
                'document format':  lambda r: (getattr(r, 'Value',       '') or '').lower(),
                'print size':       lambda r: (getattr(r, 'Value',       '') or '').lower(),
            }
            key_fn = col_map.get(col_header)
            if key_fn is None:
                return  # unknown column — skip

            tab  = self.current_tab
            prev = self._last_sort.get(tab)
            if prev and prev[0] == col_header and prev[1] == 'asc':
                items.sort(key=key_fn, reverse=True)
                self._last_sort[tab] = (col_header, 'desc')
            else:
                items.sort(key=key_fn)
                self._last_sort[tab] = (col_header, 'asc')

            data.Clear()
            for item in items:
                data.Add(item)
        except: pass
    # ── Browse / Import ───────────────────────────────────────────────

    def browse_file(self, sender, args):
        dialog = OpenFileDialog()
        dialog.Filter = "Excel/CSV Files (*.xlsx;*.csv)|*.xlsx;*.csv|Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv"
        dialog.Title  = "Select Import File"
        if dialog.ShowDialog() == DialogResult.OK:
            self.opt_file_path_tb.Text    = dialog.FileName
            self.opt_import_btn.IsEnabled = True

    def import_data(self, sender, args):
        fp = getattr(getattr(self, 'opt_file_path_tb', None), 'Text', '')
        if not fp or not os.path.exists(fp):
            forms.alert("Please browse to a valid file first.", title="Import"); return
        ext = os.path.splitext(fp)[1].lower()
        try:
            if ext == '.csv':
                self._import_csv(fp)
            elif ext in ('.xlsx', '.xls'):
                self._import_excel(fp)
            else:
                forms.alert("Unsupported format.", title="Import"); return
            self.opt_file_path_tb.Text    = ""
            self.opt_import_btn.IsEnabled = False
        except Exception as e:
            forms.alert("Import error:\n{}".format(str(e)), title="Import")

    def _import_csv(self, file_path):
        """Import from sectioned CSV — same format as standalone export."""
        import csv as _csv
        with open(file_path, 'r') as f:
            rows = list(_csv.reader(f))
        current_section = None
        for i, row in enumerate(rows):
            if not row or not row[0]: continue
            if row[0].startswith('[') and row[0].endswith(']'):
                section_name = row[0][1:-1]
                if section_name in ['Reason', 'Method', 'Format', 'Print Size']:
                    current_section = section_name
                continue
            if current_section and i > 0:
                prev = rows[i-1] if i > 0 else []
                if prev and prev[0].startswith('['): continue   # skip header row
                if current_section == 'Reason' and len(row) >= 3 and (row[0] or row[2]):
                    self.reason_data.Add(CodedRecord(row[0], row[1] if len(row)>1 else "=", row[2], self.reason_data))
                elif current_section == 'Method' and len(row) >= 3 and (row[0] or row[2]):
                    self.method_data.Add(CodedRecord(row[0], row[1] if len(row)>1 else "=", row[2], self.method_data))
                elif current_section == 'Format' and row[0]:
                    self.format_data.Add(SimpleRecord(row[0]))
                elif current_section == 'Print Size' and row[0]:
                    self.printsize_data.Add(SimpleRecord(row[0]))

    def _import_excel(self, file_path):
        """Reuse the full Excel import from SettingsManagerWindow."""
        try:
            tmp = SettingsManagerWindow.__new__(SettingsManagerWindow)
            tmp.reason_data    = self.reason_data
            tmp.method_data    = self.method_data
            tmp.format_data    = self.format_data
            tmp.printsize_data = self.printsize_data
            tmp._import_excel(file_path)
        except Exception as e:
            raise Exception("Excel import error: {}".format(str(e)))
