#! python3
from PyQt6.QtWidgets import QDialog, QMessageBox
import warning_msg
import check
import update_master

# Called by main to break scene info to the class Scene
def breakdown(self, sc_str):
    # ========== Scene number =====================
    search_stop = sc_str.find(")\t")
    #self.sc.number = sc_str[search_stop - 2:search_stop].strip()
    self.sc.number = sc_str[0:search_stop].strip()
    # Update status bar
    self.m_ui.statusbar.showMessage("Converting: #" + self.sc.eps + "/" + self.sc.number)
    
    # Check if omitted
    if 'omitted' in sc_str.lower():
        self.sc.set = 'OMITTED'
        self.sc.set_area = ''
        self.sc.set_type = ""
        self.sc.set_area = ""
        self.sc.time_of_sc = ""
        self.sc.time_req = 0.0
        self.sc.cast_in_sc = []
        self.sc.cast_in_sc_i = []
        self.sc.descriptions = ""
        update_master.df_sc(self, self.df_ep, self.sc)
        return
    # ========== Set name & Set Area ===============
    # need to fix a disdance for the search,
    # otherwise it will continue to look for the separator not in first line
    search_start = sc_str.find(" ", search_stop + 2)
    search_max = sc_str.find("\n", search_start)
    search_mid = sc_str.find("/", search_start, search_max)
    # Find the end stopper sign, it can be hyphen, En dash or Em dash
    # Test for En dash
    search_stop = sc_str.find("–", search_start, search_max)
    if search_stop < 0:
        # Test for hyphen
        search_stop = sc_str.find("-", search_start, search_max)
        if search_stop < 0:
            # Test for Em dash
            search_stop = sc_str.find("—", search_start, search_max)
            if search_stop < 0:
                # Cannot find at all
                icon = QMessageBox.Icon.Critical
                title = 'Abort?'
                text = "Cannot find the separator '-'"
                text2 = 'in #' + self.sc.eps + '/' + self.sc.number
                btns = QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel
                ret = warning_msg.show_msg(icon, title, text, text2, btns)
                if ret == QMessageBox.StandardButton.Ok:
                    # return to caller function convert_sc.breakdown
                    return 'abort'

    if search_mid < 0:
        # mid devider "/" is not found
        self.sc.set = sc_str[search_start:search_stop].strip()
        # Clear set area
        self.sc.set_area = ""
    else:
        self.sc.set = sc_str[search_start:search_mid].strip()
        self.sc.set_area = sc_str[search_mid + 1 : search_stop].strip()
    # ============= Time of scene =====================
    self.sc.time_of_sc = sc_str[search_stop + 1 : search_max].strip()
    # ============= Cast Array ========================
    # Clear up old entry
    self.sc.cast_in_sc = []
    self.sc.cast_in_sc_i = []
    search_start = search_max
    search_stop = sc_str.find("\n", search_start + 3)
    search_last_ended = search_stop
    cast_str = sc_str[search_start:search_stop].strip()
    if len(cast_str) == 0:
        # no cast in list
        icon = QMessageBox.Icon.Information
        title = 'No cast'
        text = "No cast in #" + self.sc.eps + '/' + self.sc.number
        text2 = 'OK to continue, Cancel to abbort'
        btns = QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel
        ret = warning_msg.show_msg(icon, title, text, text2, btns)
        if ret == QMessageBox.StandardButton.Cancel:
            # return to caller function convert_sc.breakdown
            return 'abort'
    else:
        cast_num = cast_str.count("|") + 1
        search_start = 0
        search_stop = cast_str.find("|")
        if cast_num == 1:
            search_stop = len(cast_str)
        while cast_num > 0:
            if self.dlg_abort:
                break
            self.sc.cast_in_sc.append(cast_str[search_start:search_stop].strip())
            # Pad equal number of entries into the index array
            self.sc.cast_in_sc_i.append(0)
            cast_num = cast_num - 1
            search_start = search_stop + 1
            search_stop = cast_str.find("|", search_start)
            if cast_num == 1:
                search_stop = len(cast_str)

    # ============= Time required =======================
    cast_num = len(self.sc.cast_in_sc)
    if cast_num <= 9:
        self.sc.time_req = self.time_arr[cast_num]
    else:
        self.sc.time_req = self.time_arr[10]
    # ============= Descriptions ========================
    # Clear up old entry
    self.sc.descriptions = ""
    search_start = search_last_ended
    if sc_str.find("\n\n", search_start + 3) < 0:
        # ends on one paragraph
        search_stop = len(sc_str)
    else:
        # more than one paragraph or "Act 2", drop all those like "Intercut with" etc. 
        search_stop = sc_str.find("\n\n", search_start + 3)

    self.sc.descriptions = sc_str[search_start:search_stop].strip()

    # display %
    percent_done = int((1/self.total_eps)*(self.doing_sc/self.total_sc + (self.doing_eps - 1)) * 100)
    self.m_ui.progressBar.setValue(percent_done)
    
    check.check_set(self, self.sc, self.df_set)
    if self.dlg_abort:
        return
    if self.sc.set_area != "":
        check.check_area(self, self.sc, self.df_set)
        if self.dlg_abort:
            return
    if len(self.sc.cast_in_sc) >= 1:
        check.check_cast(self, self.sc, self.df_cast)
        if self.dlg_abort:
            return
    update_master.df_sc(self, self.df_ep, self.sc)
    