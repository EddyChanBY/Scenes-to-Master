#! python3
from PyQt6.QtWidgets import QDialog, QMessageBox
from docx2python import docx2python
from pathlib import Path
import os
import shutil
import re
import pandas as pd
import numpy as np
import xlwings as xw
import convert_ep
import warning_msg

def convert_start(self):
    self.m_ui.statusbar.showMessage("Conversion started")
    self.m_ui.statusbar.repaint()
    # Check all files in SxS_list
    # since all files are from the add button, presuming they are OK
    self.eps_numbers = []
    for each_ep in self.SxS_list:
        s = docx2python(each_ep).text
        # Find episode number
        search_start = re.search("Episode", s, re.IGNORECASE).end()
        search_stop = s.find("\n", search_start)
        this_ep_number = s[search_start:search_stop].strip()
        self.eps_numbers.append(this_ep_number)
    # For calculate %
    self.total_eps = len(self.eps_numbers)
    # Generate master file name only if it's blank or new file
    if self.m_ui.ledit_Master.text() == "" or self.m_ui.ledit_Master.text() == "New Master.xlsx":
        if len(self.eps_numbers) == 1:
            self.master_file = self.default_path + '/' + self.eps_numbers[0] + " master.xlsx"
            self.m_ui.ledit_Master.setText(os.path.basename(self.master_file))
        else:
            self.master_file = self.default_path + '/' + self.eps_numbers[0] + " to " + self.eps_numbers[len(self.eps_numbers)-1] + " master.xlsx"
            self.m_ui.ledit_Master.setText(os.path.basename(self.master_file))
        # Create a copy of template file with same prefix as the master file
        new_template = self.master_file[:self.master_file.find('master')] + "template.xlsx"
        if Path(self.template_file).is_file():
            shutil.copy(self.template_file, new_template)
        else:
            icon = QMessageBox.Icon.Warning
            title = 'Abort'
            text = "Cannot find the template file at:"
            text2 = self.template_file
            btns = QMessageBox.StandardButton.Ok
            ret = warning_msg.show_msg(icon, title, text, text2, btns)
            if ret == QMessageBox.StandardButton.Ok:
                return
        self.template_file = new_template
        self.m_ui.ledit_Template.setText(os.path.basename(self.template_file))
    # Setting up Excel work book
    # Open the Excel app if not there
    if self.app_pid == 0:
        self.m_ui.statusbar.showMessage("Opening Excel in the background")
        self.m_ui.statusbar.repaint()
        app = xw.App(visible = False)
        self.app_pid = app.pid
    else:
        app = xw.apps[self.app_pid]
    # Check if wb_master is definded
    try:
        wb_master = app.books[self.master_file]
    except:
        # If not, check if the master file exist
        path = Path(self.master_file)
        if path.is_file():
            # Just Open the file
            wb_master = app.books.open(self.master_file)
        else:
            wb_arr = []
            for book in app.books:
                wb_arr.append(book.name)

            if 'Book1' in wb_arr:
                wb_master = app.books[wb_arr.index('Book1')]
                wb_master.save(self.master_file)
            else:
                # in the unlikely case no new workbook is open by the 'app = xw.App(visible=False)' call
                wb_master = app.books.add()
                wb_master.save(self.master_file)

    # Check if wb_template is definded
    try:
        wb_template = app.books[os.path.basename(self.template_file)]
    except:
        # Check if the tamplate file exist
        path = Path(self.template_file)
        if path.is_file():
            # Just Open the file
            try:
                wb_template = app.books.open(self.template_file)
            except:
                icon = QMessageBox.Icon.Critical
                title = 'Abort'
                text = 'Cannot open the template file'
                text2 = 'Please abort.'
                btns = QMessageBox.StandardButton.Ok
                ret = warning_msg.show_msg(icon, title, text, text2, btns)
                if ret == QMessageBox.StandardButton.Ok:
                    return   
        else:
            icon = QMessageBox.Icon.Critical
            title = 'Abort'
            text = 'Cannot find the template file'
            text2 = 'Please abort.'
            btns = QMessageBox.StandardButton.Ok
            ret = warning_msg.show_msg(icon, title, text, text2, btns)
            if ret == QMessageBox.StandardButton.Ok:
                return
        
    # open template and master file as pandas
    # Get the template ready
    if self.df_set.empty:
        self.df_set = wb_template.sheets['Sets'].range('A1').options(pd.DataFrame, 
                    header=True,
                    index=False,
                    numbers=int,
                    empty=np.nan, 
                    expand='table').value
    if self.df_cast.empty:
        self.df_cast = wb_template.sheets['Casts'].range('A1').options(pd.DataFrame, 
                             header=True,
                             index=False,
                             numbers=int,
                             empty=np.nan, 
                             expand='table').value
    if self.df_ep.empty:
        self.df_ep = wb_template.sheets['Eps'].range('A1').options(pd.DataFrame, 
                             header=True,
                             index=False,
                             numbers=int,
                             empty=np.nan, 
                             expand='table').value
    self.time_arr = wb_template.sheets['Time'].range('B2:B12').value
    
    # Going in each episode
    for i in range(len(self.SxS_list)):
        if self.dlg_abort:
            self.dlg_abort = False
            break
        s = docx2python(self.SxS_list[i]).text
        # check if footnote or endnote at the end, 
        # this will cause error in finding next scene
        # because they appear as footnote1)/t and endnote1)/t
        # the extration will give both as scene e1
        if s.find('footnote') != -1:
            s = s[0:s.find('footnote')]
        if s.find('endnote') != -1:
            s = s[0:s.find('endnote')]
        self.sc.eps = self.eps_numbers[i]
        self.doing_eps = i + 1
        ret = convert_ep.convert_this_ep(self, s)
        if ret == 'abort':
            # return to caller function master
            return
        # Update the master dataframe column heading with new casts
        cast_arr = list(self.df_cast['Cast'])
        col_arr = list(self.df_ep)
        for name in cast_arr:
            col_arr[10 + cast_arr.index(name)] = name
        self.df_ep.columns = col_arr

        # Update Excel master file here
        ws_array = []
        for sheet in wb_master.sheets:
            ws_array.append(sheet.name)
        
        if self.sc.eps not in ws_array:
            # Create new sheet and name it as the episode number
            wb_template.sheets['Eps'].copy(after=wb_master.sheets[len(ws_array) - 1])
            wb_master.sheets['Eps'].name = self.sc.eps
        else:
            # else the sheet with current episode already exist, clear it for over write
            wb_master.sheets[self.sc.eps].clear_contents()
            
        # copy data over
        wb_master.sheets[self.sc.eps].range('A1').options(index=False).value = self.df_ep
        wb_master.save()
        # Now need to clear the df_ep for next episode
        self.df_ep = wb_template.sheets['Eps'].range('A1').options(pd.DataFrame, 
                             header=True,
                             index=False,
                             numbers=int,
                             empty=np.nan, 
                             expand='table').value
    
        # Update sets and cast in the template file, do this ep by ep just in case abort
        wb_template.sheets['Sets'].range('A1').options(index=False).value = self.df_set
        wb_template.sheets['Casts'].range('A1').options(index=False).value = self.df_cast
        wb_template.save()
    self.m_ui.statusbar.showMessage("All episodes converted.")
    self.m_ui.statusbar.repaint()
    
   
