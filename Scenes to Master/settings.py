from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QMessageBox
import os
import os.path
from pathlib import Path
import shutil
import warning_msg


def get_setting(set_type):
    # get setting.txt file from local directory
    user_dir = os.path.expanduser("~")
    set_file = user_dir + '\s2m_app_setting.txt'
    if not os.path.exists(set_file):
        # setting file does not  exist, make one in currenyt directroy
        with open(set_file, 'w') as f:
            f.write('Settings:\n\n')
            f.close()
    f = open(set_file, 'a+')
    f.seek(0)
    settings_data = f.read()
    search_begin = settings_data.find('[' + set_type + ']')
    if search_begin == -1:
        # write new setting
        f.write('[' + set_type + ']\n\t')
        # Get input
        if set_type == 'Default Directory':
            folderpath = QtWidgets.QFileDialog.getExistingDirectory(None, 'Select Default Folder', user_dir)
            if folderpath == '':
                folderpath = user_dir + '\Desktop'
        # write in setting file
        f.write(folderpath + '\n')
    else:
        search_begin = search_begin + len(set_type) + 2
        search_end = settings_data[search_begin + 1:].find('[')
        if search_end == -1:
            folderpath = settings_data[search_begin:].strip()
        else:
            folderpath = settings_data[search_begin:search_end - 1].strip()
    f.close()
    # check and ensure template file copy over
    temp_file = folderpath + '/template.xlsx'
    original_temp = 'C:\\Program Files (x86)\\Long Form\\Scenes to Master\\template.xlsx'
    if not os.path.exists(temp_file):
        if os.path.exists(original_temp):
            shutil.copy(original_temp, temp_file)
        else:
            # Warning no base template file
            icon = QMessageBox.Icon.Critical
            title = 'No base template file'
            text = 'No base template file in default installation directory.'
            text2 = 'Get a copy from the Production Manger'
            btns = QMessageBox.StandardButton.Ok
            ret = warning_msg.show_msg(icon, title, text, text2, btns)
    return folderpath

def change_setting(set_type):
    # get setting.txt file from local directory
    user_dir = os.path.expanduser("~")
    set_file = user_dir + '\s2m_app_setting.txt'
    if not os.path.exists(set_file):
        # setting file does not  exist, make one in currenyt directroy
        with open(set_file, 'w') as f:
            f.write('Settings:\n\n')
            f.close()
    f = open(set_file, 'a+')
    f.seek(0)
    settings_data = f.read()
    search_begin = settings_data.find('[' + set_type + ']')
    if search_begin == -1:
        # write new setting
        f.write('[' + set_type + ']\n\t')
        old_folderpath = os.path.expanduser("~") + '\Desktop' 
    else:
        search_begin = search_begin + len(set_type) + 2
        search_end = settings_data[search_begin + 1:].find('[')
        if search_end == -1:
            old_folderpath = settings_data[search_begin:].strip()
        else:
            old_folderpath = settings_data[search_begin:search_end - 1].strip()
    if set_type == 'Default Directory':
        folderpath = QtWidgets.QFileDialog.getExistingDirectory(None, 'Select Default Folder', old_folderpath)
        if folderpath == '':
            folderpath = os.path.expanduser("~") + '\Desktop'
        # Replace the old with the new
        search_begin = search_begin + 2
        settings_data_1 = settings_data[0:search_begin]
        settings_data_2 = settings_data[search_begin + len(old_folderpath) :]
        settings_data_new = settings_data_1 + folderpath + settings_data_2
        f.close()
        f = open(set_file, 'w')
        # write in setting file
        f.write(settings_data_new)
        f.close()
    return folderpath