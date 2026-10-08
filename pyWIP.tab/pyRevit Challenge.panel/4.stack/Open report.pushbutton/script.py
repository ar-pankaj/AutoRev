# -*- coding: utf-8 -*-
__title__ = 'Open report'
__author__ = 'Kamila Milewska'
__doc__ = """Version = 1.0
Date = 27.05.2026
________________________________________________________________
Description:
Opens a selected linkify report.
________________________________________________________________
How-To:
1. Pick path.
2. Enjoy your reports.
________________________________________________________________"""
#----------------------------------------------------------------------
from pyrevit import script, forms
picked_file = forms.pick_file(title='Select linkify report to open')
if not picked_file:
    forms.alert('Cancelled!', exitscript=True)
output = script.get_output()
output.open_page(picked_file)
#----------------------------------------------------------------------
