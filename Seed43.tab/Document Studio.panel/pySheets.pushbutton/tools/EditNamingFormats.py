# -*- coding: utf-8 -*-
# EditNamingFormats.py
#pylint: disable=import-error,invalid-name,broad-except,superfluous-parens
import re
from collections import namedtuple
import os
import os.path as op
import json
from pyrevit import framework
from pyrevit import forms
from pyrevit import script, coreutils
from pyrevit.framework import Windows, Forms, ObjectModel

try:
    from Snippets import _dialogs as dlg
except Exception:
    dlg = None

from Snippets.seed43_theme import apply_seed43_palette
from Snippets import _userdata

def _alert(message):
    """Themed popup when available, falls back to forms.alert."""
    if dlg:
        dlg.message(message)
    else:
        forms.alert(message)

logger = script.get_logger()
config = script.get_config()

NamingFormatter = namedtuple('NamingFormatter', ['template', 'desc'])

class NamingFormat(forms.Reactive):
    """Print File Naming Format"""
    def __init__(self, name, template, builtin=False):
        self._name = name
        self._template = self.verify_template(template)
        self.builtin = builtin

    @staticmethod
    def verify_template(value):
        """Verify template is valid — extension is added by the main script per format."""
        # Strip any trailing export extension so the template stays format-agnostic
        for ext in ('.pdf', '.dwg', '.dgn', '.nwc', '.ifc',
                    '.png', '.jpg', '.jpeg', '.tif', '.tiff'):
            if value.lower().endswith(ext):
                value = value[:-len(ext)]
                break
        return value

    @forms.reactive
    def name(self):
        """Format name"""
        return self._name

    @name.setter
    def name(self, value):
        self._name = value

    @forms.reactive
    def template(self):
        """Format template string"""
        return self._template

    @template.setter
    def template(self, value):
        self._template = self.verify_template(value)

class EditNamingFormatsWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, start_with=None,
                 naming_dir=None, default_formats=None):
        try:
            forms.WPFWindow.__init__(self, xaml_file_name)
        except Exception as ex:
            _alert('Could not load the Naming Formats window.\n\n'
                  + self._error_chain(ex))
            raise
        apply_seed43_palette(self, op.dirname(__file__))
        self._drop_pos = 0
        self._starting_item = start_with
        self._saved = False
        # Independent storage: defaults to the shared per-sheet naming
        # folder, but a caller (e.g. Combined PDF Name) can point this
        # dialog at its own folder + its own builtin default(s).
        self._naming_dir = naming_dir or EditNamingFormatsWindow.NAMING_DIR
        self._default_formats = default_formats
        self.reset_naming_formats()
        self.reset_formatters()

    @staticmethod
    def _error_chain(ex):
        chain, e = [], ex
        while e is not None:
            chain.append(str(e))
            e = getattr(e, 'InnerException', None)
        return '\n-- caused by --\n'.join(chain)

    @staticmethod
    def get_default_formatters():
        return [
            NamingFormatter(
                template='{number}',
                desc="Sheet Number e.g. 'A1.00'"
            ),
            NamingFormatter(
                template='{name}',
                desc="Sheet Name e.g. '1ST FLOOR PLAN'"
            ),
            NamingFormatter(
                template='{name_dash}',
                desc="Sheet Name (with - for space) e.g. '1ST-FLOOR-PLAN'"
            ),
            NamingFormatter(
                template='{name_underline}',
                desc="Sheet Name (with _ for space) e.g. '1ST_FLOOR_PLAN'"
            ),
            NamingFormatter(
                template='{current_date}',
                desc="Today's Date e.g. '2019-10-12'"
            ),
            NamingFormatter(
                template='{issue_date}',
                desc="Sheet Issue Date e.g. '2019-10-12'"
            ),
            NamingFormatter(
                template='{rev_number}',
                desc="Revision Number e.g. '01'"
            ),
            NamingFormatter(
                template='{rev_desc}',
                desc="Revision Description e.g. 'ASI01'"
            ),
            NamingFormatter(
                template='{rev_date}',
                desc="Revision Date e.g. '2019-10-12'"
            ),
            NamingFormatter(
                template='{proj_name}',
                desc="Project Name e.g. 'MY_PROJECT'"
            ),
            NamingFormatter(
                template='{proj_number}',
                desc="Project Number e.g. 'PR2019.12'"
            ),
            NamingFormatter(
                template='{proj_building_name}',
                desc="Project Building Name e.g. 'BLDG01'"
            ),
            NamingFormatter(
                template='{proj_issue_date}',
                desc="Project Issue Date e.g. '2019-10-12'"
            ),
            NamingFormatter(
                template='{proj_org_name}',
                desc="Project Organization Name e.g. 'MYCOMP'"
            ),
            NamingFormatter(
                template='{proj_status}',
                desc="Project Status e.g. 'CD100'"
            ),
            NamingFormatter(
                template='{username}',
                desc="Active User e.g. 'eirannejad'"
            ),
            NamingFormatter(
                template='{revit_version}',
                desc="Active Revit Version e.g. '2019'"
            ),
            NamingFormatter(
                template='{sheet_param:PARAM_NAME}',
                desc="Value of Given Sheet Parameter e.g. Replace PARAM_NAME with target parameter name"
            ),
            NamingFormatter(
                template='{tblock_param:PARAM_NAME}',
                desc="Value of Given TitleBlock Parameter e.g. Replace PARAM_NAME with target parameter name"
            ),
            NamingFormatter(
                template='{proj_param:PARAM_NAME}',
                desc="Value of Given Project Information Parameter e.g. Replace PARAM_NAME with target parameter name"
            ),
            NamingFormatter(
                template='{glob_param:PARAM_NAME}',
                desc="Value of Given Global Parameter. Replace PARAM_NAME with target parameter name"
            ),
        ]

    @staticmethod
    def get_default_naming_formats():
        return [
            NamingFormat(
                name='PySheets Default',
                template='{number} {name}',
                builtin=True
            ),
        ]

    # Resolved through _userdata rather than relative to this file, so it
    # matches pySheets.py's USERDATA_DIR. pySheets.py owns the migration of
    # the old userdata/ tree; by the time this runs it has already happened.
    NAMING_DIR = _userdata.user_dir('pySheets', 'naming')

    @staticmethod
    def get_naming_formats(naming_dir=None, default_formats=None):
        naming_formats = list(default_formats) if default_formats is not None \
            else EditNamingFormatsWindow.get_default_naming_formats()
        d = naming_dir or EditNamingFormatsWindow.NAMING_DIR
        if op.isdir(d):
            for fn in sorted(os.listdir(d)):
                if not fn.lower().endswith('.json'):
                    continue
                try:
                    with open(op.join(d, fn), 'r') as f:
                        data = json.load(f)
                    naming_formats.append(
                        NamingFormat(name=data['name'], template=data['template']))
                except Exception:
                    pass
        return naming_formats

    @staticmethod
    def set_naming_formats(naming_formats, naming_dir=None):
        d = naming_dir or EditNamingFormatsWindow.NAMING_DIR
        if not op.isdir(d):
            os.makedirs(d)
        keep = set()
        for x in naming_formats:
            if x.builtin:
                continue
            fname = coreutils.cleanup_filename(x.name, windows_safe=True) + '.json'
            keep.add(fname)
            with open(op.join(d, fname), 'w') as f:
                json.dump({'name': x.name, 'template': x.template}, f, indent=2)
        for fn in os.listdir(d):
            if fn.lower().endswith('.json') and fn not in keep:
                os.remove(op.join(d, fn))

    @property
    def naming_formats(self):
        return self.formats_lb.ItemsSource

    @property
    def selected_naming_format(self):
        return self.formats_lb.SelectedItem

    @selected_naming_format.setter
    def selected_naming_format(self, value):
        self.formats_lb.SelectedItem = value
        self.namingformat_edit.DataContext = value

    def reset_formatters(self):
        self.formatters_wp.ItemsSource = \
            EditNamingFormatsWindow.get_default_formatters()

    def reset_naming_formats(self):
        self.formats_lb.ItemsSource = \
                ObjectModel.ObservableCollection[object](
                    EditNamingFormatsWindow.get_naming_formats(
                        self._naming_dir, self._default_formats)
                )
        if isinstance(self._starting_item, NamingFormat):
            for item in self.formats_lb.ItemsSource:
                if item.name == self._starting_item.name:
                    self.selected_naming_format = item
                    break

    def start_drag(self, sender, args):
        name_formatter = args.OriginalSource.DataContext
        Windows.DragDrop.DoDragDrop(
            self.formatters_wp,
            Windows.DataObject("name_formatter", name_formatter),
            Windows.DragDropEffects.Copy
            )

    def preview_drag(self, sender, args):
        point = args.GetPosition(self.template_tb)
        self._drop_pos = self.template_tb.GetCharacterIndexFromPoint(
            point, True)
        if self._drop_pos < 0:
            self._drop_pos = len(self.template_tb.Text or '')
        self.template_tb.CaretIndex = self._drop_pos
        args.Effects = Windows.DragDropEffects.Copy
        args.Handled = True

    def stop_drag(self, sender, args):
        name_formatter = args.Data.GetData("name_formatter")
        if name_formatter:
            new_template = \
                str(self.template_tb.Text)[:self._drop_pos] \
                + name_formatter.template \
                + str(self.template_tb.Text)[self._drop_pos:]
            self.template_tb.Text = new_template
            self.template_tb.Focus()

    SAMPLE_VALUES = {
        '{number}': 'A1.00', '{name}': '1ST FLOOR PLAN',
        '{name_dash}': '1ST-FLOOR-PLAN', '{name_underline}': '1ST_FLOOR_PLAN',
        '{index}': '0001', '{current_date}': '2026-07-07',
        '{issue_date}': '2026-07-07', '{rev_number}': '01',
        '{rev_desc}': 'ASI01', '{rev_date}': '2026-07-07',
        '{proj_name}': 'MY_PROJECT', '{proj_number}': 'PR2019.12',
        '{proj_building_name}': 'BLDG01', '{proj_issue_date}': '2026-07-07',
        '{proj_org_name}': 'MYCOMP', '{proj_status}': 'CD100',
        '{username}': 'jsmith', '{revit_version}': '2026',
    }

    def _update_preview(self):
        try:
            text = self.template_tb.Text or ''
            for token, val in self.SAMPLE_VALUES.items():
                text = text.replace(token, val)
            text = re.sub(r'\{(?:sheet_param|tblock_param|proj_param|'
                         r'glob_param):([^}]+)\}', r'<\1>', text)
            self.preview_tb.Text = text + '.pdf'
        except Exception:
            pass

    def template_changed(self, sender, args):
        self._update_preview()

    def namingformat_changed(self, sender, args):
        naming_format = self.selected_naming_format
        self.namingformat_edit.DataContext = naming_format
        self._update_preview()

    def duplicate_namingformat(self, sender, args):
        naming_format = self.selected_naming_format
        new_naming_format = NamingFormat(
            name='<unnamed>',
            template=naming_format.template
            )
        self.naming_formats.Add(new_naming_format)
        self.selected_naming_format = new_naming_format

    def delete_namingformat(self, sender, args):
        naming_format = self.selected_naming_format
        if naming_format is None:
            return
        item_index = self.naming_formats.IndexOf(naming_format)
        self.naming_formats.Remove(naming_format)
        if self.naming_formats.Count == 0:
            self.selected_naming_format = None
            return
        next_index = min([item_index, self.naming_formats.Count - 1])
        self.selected_naming_format = self.naming_formats[next_index]

    def save_formats(self, sender, args):
        EditNamingFormatsWindow.set_naming_formats(
            self.naming_formats, self._naming_dir)
        self._saved = True
        self.Close()

    def cancelled(self, sender, args):
        if not self._saved:
            self.reset_naming_formats()

    def win_close_clicked(self, sender, args):
        self.Close()

    def show_dialog(self):
        self.ShowDialog()

if __name__ == '__main__':
    EditNamingFormatsWindow('EditNamingFormats.xaml').show_dialog()