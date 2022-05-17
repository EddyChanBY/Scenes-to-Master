import xlwings as xw
import pandas as pd
import os

def clean_excel(self):
    if self.app_pid != 0:
        app = xw.apps[self.app_pid]
        try:
            # Check if wb_template is definded
            wb_template = app.books[os.path.basename(self.template_file)]
            wb_template.save()
            wb_template.close()
        except:
            pass
        try:
            wb_master = app.books[os.path.basename(self.master_file)]
            ws_array = []
            for sheet in wb_master.sheets:
                ws_array.append(sheet.name)
            if 'Sheet1' in ws_array and len(ws_array) > 1:
                wb_master.sheets['Sheet1'].delete()
            wb_master.save()
            wb_master.close()
        except:
            pass
        # clean up dataframe
        self.df_set = pd.DataFrame()
        self.df_cast = pd.DataFrame()
        self.df_ep = pd.DataFrame()
        try:
            app.kill()
            self.app_pid = 0
        except:
            pass