from PyQt6.QtWidgets import QDialog, QMessageBox

def show_msg(icon, title, text, text2, btns):
    msg = QMessageBox()
    msg.setIcon(icon)
    msg.setWindowTitle(title)
    msg.setText(text)
    msg.setInformativeText(text2)
    msg.setStandardButtons(btns)
    return msg.exec()