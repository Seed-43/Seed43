# -*- coding: utf-8 -*-
# ImportExportSettings.py
"""Export / Import Settings dialog.

One window class serves both directions (mode='export' / mode='import'),
since the two panels only differ in which side of the copy is the source
and whether the 'auto-update on startup' checkbox is shown. This lets one
user own the settings (naming presets, combined-naming presets, profiles)
and export them to a shared folder; everyone else imports from that
folder, optionally with auto-update so they always stay in sync.
"""
import os
import os.path as op
import json
import shutil

from pyrevit import forms
from pyrevit.framework import Windows

from Snippets import _dialogs as dlg


# (key, display label) — key is also the sub-folder name used on disk and
# the config-file suffix ('exp_<key>' / 'imp_<key>').
CATEGORIES = [
    ('naming',          'Sheet Naming'),
    ('naming_combined', 'Combined Sheet Naming'),
    ('profiles',        'Profiles'),
]

SETTINGS_SUBFOLDER = 'pySheets Settings'


def _read_sync(sync_path):
    try:
        with open(sync_path, 'r') as f:
            return json.load(f)
    except Exception:
        return {}


def _write_sync(sync_path, data):
    try:
        d = op.dirname(sync_path)
        if not op.isdir(d):
            os.makedirs(d)
        with open(sync_path, 'w') as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


def _copy_category(src_dir, dest_dir):
    """Copy every .json file from src_dir into dest_dir. Returns count copied."""
    if not op.isdir(src_dir):
        return 0
    files = [f for f in os.listdir(src_dir) if f.lower().endswith('.json')]
    if not files:
        return 0
    if not op.isdir(dest_dir):
        os.makedirs(dest_dir)
    for fn in files:
        shutil.copy2(op.join(src_dir, fn), op.join(dest_dir, fn))
    return len(files)


def run_export(folders, dest_root, selected):
    """Copy the selected categories from `folders` into <dest_root>/<label>/.
    Returns {label: count} for categories that copied at least one file."""
    copied = {}
    for key, label in CATEGORIES:
        if key not in selected:
            continue
        n = _copy_category(folders[key], op.join(dest_root, label))
        if n:
            copied[label] = n
    return copied


def run_import(folders, src_root, selected):
    """Copy the selected categories from <src_root>/<label>/ back into `folders`.
    Returns {label: count} for categories that copied at least one file."""
    copied = {}
    for key, label in CATEGORIES:
        if key not in selected:
            continue
        n = _copy_category(op.join(src_root, label), folders[key])
        if n:
            copied[label] = n
    return copied


def run_auto_import(folders, sync_path):
    """Silently re-import on startup if enabled + path is valid. No dialogs.
    Returns {label: count} of whatever actually got copied (may be empty)."""
    cfg = _read_sync(sync_path)
    if not cfg.get('import_auto'):
        return {}
    path = cfg.get('import_path', '')
    if not path or not op.isdir(path):
        return {}
    candidate = op.join(path, SETTINGS_SUBFOLDER)
    root = candidate if op.isdir(candidate) else path
    selected = set(key for key, _ in CATEGORIES if cfg.get('imp_' + key, True))
    if not selected:
        return {}
    try:
        return run_import(folders, root, selected)
    except Exception:
        return {}


class ImportExportSettingsWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, mode, folders, sync_path):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.mode      = mode            # 'export' or 'import'
        self._folders  = folders         # {key: folder_path}
        self._sync_path = sync_path
        self.changed   = False
        is_import = (mode == 'import')

        self.header_subtitle_tb.Text = 'Import Settings' if is_import else 'Export Settings'
        self.action_btn.Content      = 'Import' if is_import else 'Export'
        self.path_label_tb.Text      = ('Folder to import from' if is_import
                                        else 'Folder to export to')
        self.blurb_tb.Text = (
            'One person manages these settings and exports them here; everyone '
            'else imports from that shared folder.' if is_import else
            'Send these to a shared folder so other users can import them.')
        self.auto_cb.Visibility = (Windows.Visibility.Visible if is_import
                                   else Windows.Visibility.Collapsed)

        cfg = _read_sync(sync_path)
        prefix = 'imp_' if is_import else 'exp_'
        self.path_tb.Text = cfg.get('import_path' if is_import else 'export_path', '')
        self.naming_cb.IsChecked          = bool(cfg.get(prefix + 'naming', True))
        self.naming_combined_cb.IsChecked = bool(cfg.get(prefix + 'naming_combined', True))
        self.profiles_cb.IsChecked        = bool(cfg.get(prefix + 'profiles', True))
        if is_import:
            self.auto_cb.IsChecked = bool(cfg.get('import_auto', False))

        self.browse_btn.Click += self._browse_clicked
        self.action_btn.Click += self._action_clicked
        self.close_btn.Click  += lambda s, a: self.Close()
        self.win_close_btn.Click += lambda s, a: self.Close()

    def _browse_clicked(self, sender, args):
        from System.Windows.Forms import FolderBrowserDialog, DialogResult
        fbd = FolderBrowserDialog()
        fbd.Description = ('Choose the folder to import settings from'
                           if self.mode == 'import' else
                           'Choose a folder to export settings to')
        fbd.ShowNewFolderButton = (self.mode == 'export')
        if self.path_tb.Text and op.isdir(self.path_tb.Text):
            fbd.SelectedPath = self.path_tb.Text
        if fbd.ShowDialog() == DialogResult.OK:
            self.path_tb.Text = fbd.SelectedPath

    def _selected_keys(self):
        keys = set()
        if self.naming_cb.IsChecked:          keys.add('naming')
        if self.naming_combined_cb.IsChecked: keys.add('naming_combined')
        if self.profiles_cb.IsChecked:        keys.add('profiles')
        return keys

    def _save_config(self):
        cfg = _read_sync(self._sync_path)
        is_import = (self.mode == 'import')
        prefix = 'imp_' if is_import else 'exp_'
        cfg['import_path' if is_import else 'export_path'] = self.path_tb.Text
        cfg[prefix + 'naming']          = bool(self.naming_cb.IsChecked)
        cfg[prefix + 'naming_combined'] = bool(self.naming_combined_cb.IsChecked)
        cfg[prefix + 'profiles']        = bool(self.profiles_cb.IsChecked)
        if is_import:
            cfg['import_auto'] = bool(self.auto_cb.IsChecked)
        _write_sync(self._sync_path, cfg)

    def _action_clicked(self, sender, args):
        path = self.path_tb.Text
        if not path:
            dlg.message('Please choose a folder first.')
            return
        selected = self._selected_keys()
        if not selected:
            dlg.message('Please select at least one item.')
            return
        if not op.isdir(path):
            dlg.message('Folder does not exist:\n{}'.format(path))
            return
        try:
            if self.mode == 'export':
                dest_root = op.join(path, SETTINGS_SUBFOLDER)
                copied = run_export(self._folders, dest_root, selected)
                self._save_config()
                if not copied:
                    dlg.message('Nothing to export — no saved items found.')
                    return
                self.changed = True
                summary = '\n'.join('{}: {} file(s)'.format(k, v) for k, v in copied.items())
                dlg.message('Exported to:\n{}\n\n{}'.format(dest_root, summary))
            else:
                candidate = op.join(path, SETTINGS_SUBFOLDER)
                root = candidate if op.isdir(candidate) else path
                if not dlg.confirm(
                        'Import from:\n{}\n\n'
                        'Existing presets/profiles with the same name will be '
                        'overwritten.'.format(root)):
                    return
                copied = run_import(self._folders, root, selected)
                self._save_config()
                if not copied:
                    dlg.message('No matching settings found in that folder.')
                    return
                self.changed = True
                summary = '\n'.join('{}: {} file(s)'.format(k, v) for k, v in copied.items())
                dlg.message('Imported:\n\n{}'.format(summary))
        except Exception as ex:
            dlg.message('{} failed.\n\n{}'.format(self.mode.title(), ex))


def show_export(folders, sync_path):
    xaml_path = op.join(op.dirname(__file__), 'ImportExportSettings.xaml')
    win = ImportExportSettingsWindow(xaml_path, 'export', folders, sync_path)
    win.ShowDialog()
    return win.changed


def show_import(folders, sync_path):
    xaml_path = op.join(op.dirname(__file__), 'ImportExportSettings.xaml')
    win = ImportExportSettingsWindow(xaml_path, 'import', folders, sync_path)
    win.ShowDialog()
    return win.changed
