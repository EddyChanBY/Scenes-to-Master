
import xlwings as xw
from pathlib import Path
import os
import pandas as pd
import numpy as np
from PyQt6.QtWidgets import QDialog, QMessageBox
import warning_msg
# to do reporting

def est_wb(self):
    # to open workbooks, access sheet and set up data frames
    if self.app_pid == 0:
        app = xw.App()
        self.app_pid = app.pid
    else:
        app = xw.apps[self.app_pid]
        
    # Check if work books are open
    wb_array = []
    for book in app.books:
        wb_array.append(book.name)
    
    # Master workbook
    if os.path.basename(self.master_file) in wb_array:
        wb_master = app.books[os.path.basename(self.master_file)]
    else:
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
    
    #Template workbook
    if os.path.basename(self.template_file) in wb_array:
        wb_temp = app.books[os.path.basename(self.template_file)]
    else:
        # Check if the tamplate file exist
        path = Path(self.template_file)
        if path.is_file():
            # Just Open the file
            try:
                wb_temp = app.books.open(self.template_file)
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
            text = "Cannot find the template file"
            text2 = 'Please abort.'
            btns = QMessageBox.StandardButton.Ok
            ret = warning_msg.show_msg(icon, title, text, text2, btns)
            if ret == QMessageBox.StandardButton.Ok:
                return
    return[wb_temp, wb_master]

def group_to_team(wb_temp, wb_master):
    df_team = wb_temp.sheets['Teams'].range('A1').options(pd.DataFrame, 
                             header=True,
                             index=False,
                             numbers=int,
                             empty=np.nan, 
                             expand='table').value

    ws_array = []
    for sheet in wb_master.sheets:
        ws_array.append(sheet.name)
    team_array = []
    all_eps = {}
    # tested each_ep is str 
    for each_ep in ws_array:
        if each_ep.isdigit():
            try:
                team_name = df_team.loc[(df_team['From'] <= int(each_ep)) 
                    & (df_team['To'] >= int(each_ep)), 'Team'].values[0]
            except:
                icon = QMessageBox.Icon.Warning
                title = 'Abort'
                text = 'It seems Episode ' + each_ep + ' is not in the team schedule.'
                text2 = 'Need to abort now!'
                btns = QMessageBox.StandardButton.Ok 
                ret = warning_msg.show_msg(icon, title, text, text2, btns)
                return 'Aborted'

            if team_name not in team_array:
                all_eps[team_name] = [each_ep]
                team_array.append(team_name)
            else:
                all_eps[team_name].append(each_ep)
    return all_eps

def prepare_team_master(self, wb_temp, wb_master, all_eps):
    ep_done = 0
    total_eps = 0
    # Update the master dataframe column heading with new casts
    df_cast = wb_temp.sheets['Casts'].range('A1').options(pd.DataFrame, 
        header=True, index=False, expand='table').value
    cast_arr = list(df_cast['Cast'])
    
    for eps in all_eps:
        total_eps = total_eps + len(all_eps[eps])
    for team in all_eps:
        team_eps = all_eps[team]
        df_team = pd.DataFrame()
        for ep in team_eps:
            if df_team.empty:
                # set first ep in df
                df_team = wb_master.sheets[ep].range('A1').options(pd.DataFrame, 
                             header=True,
                             index=False,
                             empty=np.nan, 
                             expand='table').value
                # print(df_team)
                # Update the cast column headings
                col_arr = list(df_team)
                for index, name in enumerate(cast_arr):
                    col_arr[10 + index] = name
                df_team.columns = col_arr
            else:
                # df exist and just append
                df_temp = wb_master.sheets[ep].range('A1').options(pd.DataFrame, 
                             header=True,
                             index=False,
                             empty=np.nan, 
                             expand='table').value
                # Update the cast column headings
                col_arr = list(df_temp)
                for index, name in enumerate(cast_arr):
                    col_arr[10 + index] = name
                df_temp.columns = col_arr
                df_team = pd.concat([df_team, df_temp])
            ep_done = ep_done + 1
        # all eps for this team added to df
        df_team = df_team.sort_values(['Type', 'Set'], ascending = [False, True])
        
        sheets = []
        for sheet in wb_master.sheets:
            sheets.append(sheet.name)
        team_name = 'Team ' + team
        if team_name in sheets:
            wb_master.sheets[team_name].clear_contents()
        else:
            wb_temp.sheets['Eps'].copy(after=wb_master.sheets[wb_master.sheets.count - 1])
            wb_master.sheets['Eps'].name = team_name
        # copy data over
        wb_master.sheets[team_name].range('A1').options(index=False).value = df_team
        percent_done = int(ep_done/total_eps * 80)
        self.m_ui.progressBar.setValue(percent_done)
        
def report_sets(self, wb_temp, wb_master, all_eps):
    df_set = wb_temp.sheets['Sets'].range('A1').options(pd.DataFrame, 
        header=True, index=False, expand='table').value
    
    set_list = list(df_set['Set'])
    team_list = []
    col_list = ['Type', 'Set']
    for team in all_eps.keys():
        team_list.append('Team ' + team)
        col_list.append('Team ' + team)
    col_list.append('Total')
    # prepare df_report for set
    df_report = pd.DataFrame(columns = col_list)
    df_report['Type'] = df_set['Type']
    df_report['Set'] = df_set['Set']
    # check data in team master
    for index,team in enumerate(team_list):
        col = index + 2
        df_team = wb_master.sheets[team].range('A1').options(pd.DataFrame, 
                             header=True,
                             index=False,
                             numbers=int,
                             empty=np.nan, 
                             expand='table').value
        
        for row, set_name in enumerate(set_list):
            df_report.iat[row, col] = (df_team.Set == set_name).sum()
    # Calculate total
    for row, set_name in enumerate(set_list):
        total_sc = 0
        col = 2
        for index, team in enumerate(team_list):
            total_sc = total_sc + df_report.iat[row, col]
            col = col + 1
        df_report.iat[row, col] = total_sc
    # Do scene total
    row = len(df_report.index)
    col = 2
    for team in team_list:
        df_report.loc[row, team] = df_report[team].sum()
    df_report.loc[row, 'Total'] = df_report['Total'].sum()
    # Do percentage
    total_ST = df_report.loc[df_report.Type == 'ST', 'Total'].sum()
    total_RC = df_report.loc[df_report.Type == 'RC', 'Total'].sum()
    total_OB = df_report.loc[df_report.Type == 'OB', 'Total'].sum()
    
    sheets = []
    for sheet in wb_master.sheets:
        sheets.append(sheet.name)
    if 'Set Report' in sheets:
        wb_master.sheets['Set Report'].clear()
    else:
        wb_master.sheets.add(name = 'Set Report', after = wb_master.sheets[wb_master.sheets.count - 1])
    wb_master.sheets['Set Report'].range('A1').options(index=False).value = df_report
    
    # Format sheet
    ws_set = wb_master.sheets['Set Report']
    set_maxr = ws_set.range('B3').end('down').row + 1
    set_maxc = ws_set.range('A1').end('right').column
    # Set the column width
    ws_set.range((1, 2),(set_maxr, 1)).autofit()
    ws_set.range((1, 3),(set_maxr,set_maxc)).column_width = 7
    # Set alignment
    ws_set.range((1,3),(set_maxr,set_maxc)).api.HorizontalAlignment = -4108
    # Bold the total values
    ws_set.range((1,set_maxc),(set_maxr,set_maxc)).api.Font.Bold = True
    # Draw border
    row = 2
    while row < set_maxr:
        ws_set.range((row,1),(row,set_maxc)).api.Borders(8).LineStyle = 1
        ws_set.range((row,1),(row,set_maxc)).api.Borders(8).Weight = 2
        row = row + 5
    ws_set.range((set_maxr,1),(set_maxr,set_maxc)).api.Borders(8).LineStyle = 1
    ws_set.range((set_maxr,1),(set_maxr,set_maxc)).api.Borders(8).Weight = 2
    ws_set.range((set_maxr,1),(set_maxr,set_maxc)).api.Borders(9).LineStyle = -4119
    ws_set.range((set_maxr,1),(set_maxr,set_maxc)).api.Borders(9).Weight = 2
    # display percentage
    row = set_maxr + 2
    total_set = total_ST + total_RC + total_OB
    ws_set.range((row,1)).value = 'Statistic:'
    ws_set.range((row + 1,2)).value = 'Descriptions'
    ws_set.range((row + 1,3)).value = 'Count'
    ws_set.range((row + 1,4)).value = 'Percentage'
    ws_set.range((row + 2,1)).value = 'ST'
    ws_set.range((row + 2,2)).value = 'Studio Sets'
    ws_set.range((row + 2,3)).value = total_ST
    ws_set.range((row + 2,4)).value = (total_ST / total_set)
    ws_set.range((row + 3,1)).value = 'RC'
    ws_set.range((row + 3,2)).value = 'Recurrent OB Sets'
    ws_set.range((row + 3,3)).value = total_RC
    ws_set.range((row + 3,4)).value = (total_RC / total_set)
    ws_set.range((row + 4,1)).value = 'OB'
    ws_set.range((row + 4,2)).value = 'Outside Broadcast'
    ws_set.range((row + 4,3)).value = total_OB
    ws_set.range((row + 4,4)).value = (total_OB / total_set)
    ws_set.range((row,4),(row + 4,4)).api.NumberFormat = "0.00%"


def report_cast(self, wb_temp, wb_master, all_eps):
    df_cast = wb_temp.sheets['Casts'].range('A1').options(pd.DataFrame, 
        header=True, index=False, expand='table').value
    cast_arr = list(df_cast['Cast'])
    # set up report data frame
    team_list = []
    for team in all_eps.keys():
        team_list.append('Team ' + team)
    # prepare df_report for set
    # Create multi header
    header = pd.MultiIndex.from_product([team_list +['Total'], ['#Sc', '#Set']], names=['Team','Count'])
    df_report = pd.DataFrame(columns = header)
    # Add the cast column with cast name
    df_report.insert(0, 'Cast', cast_arr, True)
    
    # check data in team master
    for index,team in enumerate(team_list):
        df_team = wb_master.sheets[team].range('A1').options(pd.DataFrame, 
                             header=True,
                             index=False,
                             numbers=int,
                             empty=np.nan, 
                             expand='table').value
        
        for row, cast_name in enumerate(cast_arr):
            col = index * 2 + 1
            df_report.iat[row, col] = (df_team.loc[df_team[cast_name] == 'X', ['Set']]).count().to_list()[0]
            col = col + 1
            df_report.iat[row, col] = (df_team.loc[df_team[cast_name] == 'X', ['Set']]).nunique().to_list()[0]
            col = col + 1

    # Calculate total
    for row, cast_name in enumerate(cast_arr):
        total_sc = 0
        total_set = 0
        col = 1
        for index, team in enumerate(team_list):
            total_sc = total_sc + df_report.iat[row, col]
            col = col + 1
            total_set = total_set + df_report.iat[row, col]
            col = col + 1
        df_report.iat[row, col] = total_sc
        col = col + 1
        df_report.iat[row, col] = total_set

    sheets = []
    for sheet in wb_master.sheets:
        sheets.append(sheet.name)
    if 'Cast Report' in sheets:
        wb_master.sheets['Cast Report'].clear()
    else:
        wb_master.sheets.add(name = 'Cast Report', after = wb_master.sheets[wb_master.sheets.count - 1])
    wb_master.sheets['Cast Report'].range('A1').options(index=False).value = df_report
    
    # Format sheet
    ws_cast = wb_master.sheets['Cast Report']
    cast_maxr = ws_cast.range('A3').end('down').row
    cast_maxc = ws_cast.range('A1').end('right').column
    # suppress suppress prompts and alert messages
    wb_master.app.display_alerts = False
    # Set the column width
    ws_cast.range((1, 1),(cast_maxr, 1)).autofit()
    ws_cast.range((3,2),(cast_maxr,cast_maxc)).column_width = 5
    # Set alignment
    ws_cast.range((3,2),(cast_maxr,cast_maxc)).api.HorizontalAlignment = -4108
    # Bold the total values
    ws_cast.range((3,cast_maxc - 1),(cast_maxr,cast_maxc)).api.Font.Bold = True
    # Greying the #Set columns
    col = 3
    for team in team_list:
        ws_cast.range((3,col),(cast_maxr,col)).color = (230, 230, 230)
        ws_cast.range((1,col - 1),(1,col)).merge()
        col = col + 2
    ws_cast.range((3,col),(cast_maxr,col)).color = (230, 230, 230)
    ws_cast.range((1,col - 1),(1,col)).merge()
    row = 3
    # Draw border
    while row < cast_maxr:
        ws_cast.range((row,1),(row,cast_maxc)).api.Borders(8).LineStyle = 1
        ws_cast.range((row,1),(row,cast_maxc)).api.Borders(8).Weight = 2
        row = row + 5
    # disable suppress suppress prompts and alert messages
    wb_master.app.display_alerts = True

