#! python3

from PyQt6.QtWidgets import QDialog, QMessageBox
import convert_sc
import warning_msg

def convert_this_ep(self, s):
    self.total_sc = s.count(")\t")
    # Each episode reset number of scene done 
    self.doing_sc = 0
    # find initial searching point
    # search_start = re.search("Act 1", s, re.IGNORECASE).end()
    search_start = s.find('1)\t') -1
    if search_start < 0:
        icon = QMessageBox.Icon.Critical
        title = 'Abort'
        text = 'Cannot find Sc 1'
        text2 = 'Please abort and try re-number all the scenes.'
        btns = QMessageBox.StandardButton.Ok
        ret = warning_msg.show_msg(icon, title, text, text2, btns)
        if ret == QMessageBox.StandardButton.Ok:
            return
        
    # Find first scene number
    search_stop = s.find(")\t", search_start + 3) - 3
    # Do while the scene number can be found
    while s.find(")\t", search_start) > 0:
        if self.dlg_abort:
            break
        str_found = s[search_start:search_stop].strip()
        search_start = s.find(")\t", search_stop) - 3
        search_stop = s.find(")\t", search_stop + 6) - 3
        self.doing_sc = self.doing_sc + 1
        ret = convert_sc.breakdown(self, str_found)
        if ret == 'abort':
            # return to caller function
            return 'abort'
        # input done, now checking
        