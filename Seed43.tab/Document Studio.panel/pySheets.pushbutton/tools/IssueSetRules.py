# -*- coding: utf-8 -*-
# IssueSetRules.py
"""Issue Set Rules window: pick the combined set order and confirm the
archive rules the Issue set file output follows.

The rules themselves are fixed (see issue_set.py); they are shown here so
whoever turns the mode on knows what it will move before it moves it.
"""
import os.path as op
from pyrevit import forms
from pyrevit.framework import Windows

from Snippets.seed43_theme import apply_seed43_palette

import issue_set


class IssueSetRulesWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, rules, naming_warning=None):
        forms.WPFWindow.__init__(self, xaml_file_name)
        apply_seed43_palette(self, op.dirname(xaml_file_name))
        self.result = None

        self._order_rbs = {issue_set.ORDER_NUMBER:  self.order_number_rb,
                           issue_set.ORDER_BROWSER: self.order_browser_rb,
                           issue_set.ORDER_LIST:    self.order_list_rb}
        order = (rules or {}).get('order', issue_set.ORDER_NUMBER)
        self._order_rbs.get(order, self.order_number_rb).IsChecked = True

        if naming_warning:
            self.naming_warning_tb.Text = naming_warning
            self.naming_warning.Visibility = Windows.Visibility.Visible

    def save_clicked(self, sender, args):
        order = next((k for k, rb in self._order_rbs.items() if rb.IsChecked),
                     issue_set.ORDER_NUMBER)
        self.result = {'order': order}
        self.Close()

    def cancel_clicked(self, sender, args):
        self.Close()

    def win_close_clicked(self, sender, args):
        self.Close()


def show_rules(rules, naming_warning=None):
    """Show the rules modally. Returns the rules dict, or None if cancelled."""
    xaml_path = op.join(op.dirname(__file__), 'IssueSetRules.xaml')
    win = IssueSetRulesWindow(xaml_path, rules, naming_warning)
    win.ShowDialog()
    return win.result
