# -*- coding: utf-8 -*-
__title__  = "Sheet Scale"
__author__ = "Seed43"
__doc__    = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Finds and replaces detail numbers on section viewports across one or
more selected sheets. Supports adding a prefix, suffix, or swapping
part of the existing number for something new.

Processes all section viewports on the selected sheets at once.
Automatically skips any change that would create a duplicate detail
number on the same sheet.
_____________________________________________________________________
How-to:
-> Select one or more sheets in the project browser or make a sheet
   the active view
-> Run the tool
-> In the window that opens, type what you want to find and replace
-> Add a prefix or suffix if needed
-> Click Run

The tool will update all section viewport detail numbers on the
selected sheets and tell you how many were changed.
_____________________________________________________________________
Notes:
- Only section viewports are affected, not plans or elevations
- If a resulting number already exists on that sheet, that viewport
  is skipped with a warning
- If a number is already taken, the tool will try up to 5 times
  by appending an asterisk before giving up
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

from Autodesk.Revit.DB import (
    ViewSheet, ViewType, Transaction, BuiltInParameter
)
from Autodesk.Revit.Exceptions import ArgumentException
from pyrevit import forms
from clr import AddReference
AddReference("System")
from System.Diagnostics.Process import Start
from System.Windows.Window import DragMove
from System.Windows.Input import MouseButtonState

# ── VARIABLES ────────────────────────────────────────────────────────────────

uidoc = __revit__.ActiveUIDocument
doc   = __revit__.ActiveUIDocument.Document


# ── GET SELECTED SHEETS ───────────────────────────────────────────────────────

selected_ids = uidoc.Selection.GetElementIds()
sheets       = []

if selected_ids:
    for elem_id in selected_ids:
        elem = doc.GetElement(elem_id)
        if isinstance(elem, ViewSheet):
            sheets.append(elem)

if not sheets and isinstance(doc.ActiveView, ViewSheet):
    sheets.append(doc.ActiveView)

if not sheets:
    forms.alert(
        "No sheets selected, and the active view is not a sheet view.",
        exitscript=True
    )


# ── GET SECTION VIEWPORTS ─────────────────────────────────────────────────────

selected_viewports = []

for sheet in sheets:
    viewport_ids = sheet.GetAllViewports()
    for vp_id in viewport_ids:
        vp   = doc.GetElement(vp_id)
        view = doc.GetElement(vp.ViewId)
        if view and view.ViewType == ViewType.Section:
            selected_viewports.append((vp, sheet))

if not selected_viewports:
    forms.alert(
        "No section viewports found on the selected sheet(s).",
        exitscript=True
    )


# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def update_project_browser():
    from Autodesk.Revit.UI import DockablePanes, DockablePane
    project_browser_id = DockablePanes.BuiltInDockablePanes.ProjectBrowser
    project_browser    = DockablePane(project_browser_id)
    project_browser.Hide()
    project_browser.Show()


# ── GUI ───────────────────────────────────────────────────────────────────────

class MyWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.main_title.Text = "Find and Replace Detail Numbers"

    def rename(self):
        t = Transaction(doc, "Find and Replace Detail Numbers")
        t.Start()
        modified_count = self.rename_detail_number()
        update_project_browser()
        t.Commit()
        forms.alert(
            "Successfully updated detail numbers for {} section viewports "
            "across {} sheet(s).".format(modified_count, len(sheets))
        )

    def rename_detail_number(self):
        modified_count = 0
        for vp, sheet in selected_viewports:
            existing_numbers = [
                doc.GetElement(vp_id)
                   .Parameter[BuiltInParameter.VIEWPORT_DETAIL_NUMBER]
                   .AsString()
                for vp_id in sheet.GetAllViewports()
            ]
            current_detail_num = (
                vp.Parameter[BuiltInParameter.VIEWPORT_DETAIL_NUMBER].AsString() or ""
            )
            detail_num_new = (
                self.detail_number_prefix
                + current_detail_num.replace(
                    self.detail_number_find, self.detail_number_replace)
                + self.detail_number_suffix
            )
            if detail_num_new in existing_numbers and detail_num_new != current_detail_num:
                forms.alert(
                    "Detail number '{}' already exists on sheet {}. "
                    "Skipping viewport.".format(detail_num_new, sheet.Title),
                    exitscript=False
                )
                continue
            fail_count = 0
            while fail_count < 5:
                fail_count += 1
                try:
                    if current_detail_num != detail_num_new:
                        vp.Parameter[BuiltInParameter.VIEWPORT_DETAIL_NUMBER].Set(
                            detail_num_new)
                        modified_count += 1
                        break
                except ArgumentException:
                    detail_num_new += "*"
                except Exception as e:
                    forms.alert(
                        "Failed to set detail number '{}' on sheet {}: {}".format(
                            detail_num_new, sheet.Title, str(e)),
                        exitscript=False
                    )
                    break
        return modified_count

    # ── GUI properties ────────────────────────────────────────────────────────

    def detail_number_find(self):
        return self.input_detail_number_find.Text
    detail_number_find = property(detail_number_find)

    def detail_number_replace(self):
        return self.input_detail_number_replace.Text or ""
    detail_number_replace = property(detail_number_replace)

    def detail_number_prefix(self):
        return self.input_detail_number_prefix.Text or ""
    detail_number_prefix = property(detail_number_prefix)

    def detail_number_suffix(self):
        return self.input_detail_number_suffix.Text or ""
    detail_number_suffix = property(detail_number_suffix)

    # ── GUI event handlers ────────────────────────────────────────────────────

    def button_close(self, sender, e):
        self.Close()

    def Hyperlink_RequestNavigate(self, sender, e):
        Start(e.Uri.AbsoluteUri)

    def header_drag(self, sender, e):
        if e.LeftButton == MouseButtonState.Pressed:
            DragMove(self)

    def button_run(self, sender, e):
        if not self.detail_number_find:
            forms.alert("Find field cannot be empty.", exitscript=True)
        self.rename()


# ── MAIN ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        MyWindow("Script.xaml").ShowDialog()
    except IOError as e:
        forms.alert(
            "Could not find Script.xaml: {}".format(str(e)),
            exitscript=True
        )
