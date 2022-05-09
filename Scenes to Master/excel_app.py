import xlwings as xw
from PyQt6.QtWidgets import QMessageBox
import warning_msg

def check():
    app_count = xw.apps.count
    if app_count == 0:
        text = 'No Excel application opened.'
    elif app_count == 1:
        text = 'No Excel application opened.'
    else:
        text = str(app_count) + ' Excel applications opened.'
    # Show message box
    icon = QMessageBox.Icon.Information
    title = 'Excel Application'
    text2 = ''
    btns = QMessageBox.StandardButton.Ok
    ret = warning_msg.show_msg(icon, title, text, text2, btns)
    if ret == QMessageBox.StandardButton.Ok:
        return

def clear_all():
    app_count = xw.apps.count
    if app_count > 0:
        # Show message box
        icon = QMessageBox.Icon.Warning
        title = 'Excel Application'
        text = str(app_count) + ' Excel applications opened.'
        text2 = 'All Excel applications will be closed'
        btns = QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel
        ret = warning_msg.show_msg(icon, title, text, text2, btns)
        if ret == QMessageBox.StandardButton.Cancel:
            return
        for key in xw.apps.keys():
            xw.apps[key].kill()