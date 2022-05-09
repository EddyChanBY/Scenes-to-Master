#! python3
from PyQt6.QtWidgets import QDialog, QMessageBox
import difflib
import pandas as pd
import numpy as np
import warning_msg

# Called by main to check input from the class Scene

# Check set
def check_set(self, sc, df):
    if self.dlg_abort:
        return
    # clean up both input and compared strings
    # change to all upper case
    name_input = sc.set.lower()
    name_input = name_input.strip()
    # change Closing Single Quote to Apostrophe
    name_input = name_input.replace("’", "'")
    set_arr = df["Set"].to_numpy()
    not_there = True
    for this_set in set_arr:
        # upper case
        name_compare = this_set.lower()
        name_compare = name_compare.strip()
        # change Closing Single Quote to Apostrophe
        name_compare = name_compare.replace("’", "'")
        if name_input == name_compare:
            sc.set = this_set
            index_in_set = df.index[df['Set']==this_set].tolist()
            sc.set_type = df.loc[index_in_set[0], "Type"]
            not_there = False
            break
    if not_there:
        # sugestions
        name_input = ' '.join(elem.capitalize() for elem in name_input.split())
        search_arr = difflib.get_close_matches(name_input, set_arr, len(set_arr), 0)
        dlg_title = "Found new set"
        dlg_where = "At #" + sc.eps + "/" + sc.number + " found new set name"
        dlg_instruction = "If this is not a new set, select from list:"
        ck_result = self.launch_ck_dialog_box(dlg_title, dlg_where, dlg_instruction, name_input, search_arr)
        if ck_result[0] == 1:
            # This is a new set
            # Prepare fill in string
            fill_in_arr = ["OB", ck_result[1]]
            for i in range(10):
                fill_in_arr.append(np.nan)
            # Update set list
            df.loc[len(df.index)] = fill_in_arr
            # Since additional cast added, update cast array
            set_arr = df["Set"].to_numpy()
            # Update set type
            sc.set_type = "OB"
        elif ck_result[0] == 0:
            # Not a new set, update set type
            index_in_set = df.index[df['Set']==ck_result[1]].tolist()
            sc.set_type = df.loc[index_in_set[0], "Type"]
        elif ck_result[0] == 3:
            # ck_result[0] == 3, user want to abort
            self.dlg_abort = True
            return 'Abort Dialog'
        sc.set = ck_result[1]
        
    
# Check area
def check_area(self, sc, df):
    if self.dlg_abort:
        return
    # clean up both input and compared strings
    # change to all upper case
    name_input = sc.set_area.lower()
    name_input = name_input.strip()
    in_set = sc.set
    # change Closing Single Quote to Apostrophe
    name_input = name_input.replace("’", "'")
    #look for the row in df
    df_rows = df.index[df['Set'] == in_set].tolist()
    #get the result  to form a list of areas
    area_list = list(df.iloc[df_rows[0], :].dropna())
    del area_list[0:2]
        
    area_arr = area_list
    not_there = True
    for this_area in area_arr:
        # upper case
        name_compare = this_area.lower()
        name_compare = name_compare.strip()
        # change Closing Single Quote to Apostrophe
        name_compare = name_compare.replace("’", "'")
        if name_input == name_compare:
            sc.set_area = this_area
            not_there = False
            break
    if not_there:
        # sugestions
        name_input = ' '.join(elem.capitalize() for elem in name_input.split())
        if len(area_arr) > 0:
            search_arr = difflib.get_close_matches(name_input, area_arr, len(area_arr), 0)
        else:
            search_arr = []
        dlg_title = "Found new area"
        dlg_where = "At #" + sc.eps + "/" + sc.number + " in " + sc.set + " found new area name"
        dlg_instruction = "If this is not a new area, select from list:"
        ck_result = self.launch_ck_dialog_box(dlg_title, dlg_where, dlg_instruction, name_input, search_arr)
        if ck_result[0] == 1:
            # This is a new area, update area list in template data frame
            if len(area_list) < 10:
                df.loc[df[df['Set'] == sc.set].index[0], df.columns[len(area_list) + 2]] = ck_result[1]
            else:
                icon = QMessageBox.Icon.Warning
                title = 'Skip Aera'
                text = "Too many aeras for this set,"
                text2 = 'droping this new area.'
                btns = QMessageBox.StandardButton.Ok
                ret = warning_msg.show_msg(icon, title, text, text2, btns)
        elif ck_result[0] == 3:
            # ck_result[0] == 3, user want to abort
            self.dlg_abort = True
            return 'Abort Dialog'
        sc.set_area = ck_result[1]
    
# Check cast
def check_cast(self, sc, df):
    if self.dlg_abort:
        return
    # Get the set array in df_cast from template file
    cast_arr = df["Cast"].to_numpy()
    # Make a copy of input cast list to delete extra
    new_casts_input = sc.cast_in_sc.copy()
    new_casts_input_i = sc.cast_in_sc_i.copy()
    for each_cast in sc.cast_in_sc:
        # change to all lower case
        name_input = each_cast.lower()
        name_input = name_input.strip()
        # change Closing Single Quote to Apostrophe
        name_input = name_input.replace("’", "'")
        not_there = True
        for this_cast in cast_arr:
            # lower case
            name_compare = this_cast.lower()
            name_compare = name_compare.strip()
            # change Closing Single Quote to Apostrophe
            name_compare = name_compare.replace("’", "'")
            if name_input == name_compare:
                # Found in template sheet
                # where is the cast in sc object
                index_in_sc = new_casts_input.index(each_cast)
                # where is the found cast name in template
                i_in_cast_arr = np.where(cast_arr == this_cast)
                # Update the name s in SxS to as in temple
                new_casts_input[index_in_sc] = cast_arr[i_in_cast_arr[0][0]]
                # Update row index as on temple 
                new_casts_input_i[index_in_sc] = i_in_cast_arr[0][0]
                
                not_there = False
                break
        if not_there:
            # sugestions
            name_input = ' '.join(elem.capitalize() for elem in name_input.split())
            search_arr = difflib.get_close_matches(name_input, cast_arr, len(cast_arr), 0)
            dlg_title = "Found new cast"
            dlg_where = "At #" + sc.eps + "/" + sc.number + " found new cast name"
            dlg_instruction = "If this is not a new cast, select from list:"
            ck_result = self.launch_ck_dialog_box(dlg_title, dlg_where, dlg_instruction, name_input, search_arr)
            if ck_result[0] == 1:
                # This is a new cast or part time
                # Update cast list in sc
                # where is the cast in sc object
                index_in_sc = new_casts_input.index(each_cast)
                # Replace that with the new 
                new_casts_input[index_in_sc] = ck_result[1]
                # Update cast list in template, old method don't work, the df=xxx create a copy, 
                # that does not update the calling df
                # df_new = pd.DataFrame([['New', ck_result[1]]], columns = ['Type', 'Cast'])
                #df = pd.concat([df, df_new], ignore_index=True)
                # Can just modify the df passed in
                df.loc[len(df.index)] = ["New", ck_result[1], np.nan]
                # Since additional cast added, update cast array
                cast_arr = df["Cast"].to_numpy()
                # Where is the cast index for just loaded
                i_in_index_list = df.index[df['Cast']== ck_result[1]].tolist()
                new_casts_input_i[index_in_sc] = i_in_index_list[0]
            elif ck_result[0] == 0:
                # This is a cast in list
                # where is the cast in sc object
                index_in_sc = new_casts_input.index(each_cast)
                # where is the found cast name in template
                i_in_cast_arr = np.where(cast_arr == ck_result[1])
                new_casts_input[index_in_sc] = ck_result[1]
                # Update row index as on temple 
                new_casts_input_i[index_in_sc] = i_in_cast_arr[0][0]
            elif ck_result[0] == 2:
                # ck_result[0] == 2, this is an extra
                if sc.extra_str == "":
                    sc.extra_str = ck_result[1]
                else:
                    sc.extra_str = sc.extra_str + ', ' + ck_result[1]
                # to delete the extra in new casts list
                index_in_sc = new_casts_input.index(each_cast)
                del new_casts_input[index_in_sc]
                del new_casts_input_i[index_in_sc]
            elif ck_result[0] == 3:
                # ck_result[0] == 3, user want to abort
                self.dlg_abort = True
                return 'Abort Dialog'
    sc.cast_in_sc = new_casts_input.copy()
    sc.cast_in_sc_i = new_casts_input_i.copy()
