import numpy as np

def df_sc(self, df, sc):
    # Prepare update string
    Update_arr = ['hh:mm - hh:mm', sc.eps, sc.number, sc.set, sc.set_area, sc.time_of_sc, sc.set_type, sc.time_req, "", sc.descriptions]
    for i in range(62+3):
        Update_arr.append(np.nan)
    last_row = len(df.index)
     # Update set list
    df.loc[last_row] = Update_arr
    # Now the cast
    for each_cast in sc.cast_in_sc_i:
        df.iat[last_row, each_cast + 10] = "X"
    # The extra if any
    if sc.extra_str != "":
        df.iat[last_row, 72] = sc.extra_str
        # Updated, clear for next scene
        sc.extra_str = ""
    
    
   