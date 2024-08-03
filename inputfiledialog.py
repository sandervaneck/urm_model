# -*- coding: utf-8 -*-
"""
Created on Mon Jun  3 12:33:42 2024

@author: M67B363
"""
from PyQt5.QtWidgets import QFileDialog

def choose_file(self, bestandsnaam):
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    filename, _ = QFileDialog.getOpenFileName(None, f"Selecteer {bestandsnaam}", "", "Excel Files (*.xlsx *.xlsm)", options=options)
    return filename