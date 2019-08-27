from tkinter import *
from tkinter import messagebox
from Global import grabobjs
from Global import CryptHandle

import smtplib
import pandas as pd
import os

# Global Variable declaration
curr_dir = os.path.dirname(os.path.abspath(__file__))
main_dir = os.path.dirname(curr_dir)
global_objs = grabobjs(main_dir, 'Grab_Distro')


class SettingsGUI:
    save_settings_button = None
    cat_txtbox = None
    distro_txtbox = None
    emp_txtbox = None
    sql_tables = pd.DataFrame()

    # Function that is executed upon creation of SettingsGUI class
    def __init__(self):
        self.header_text = 'Welcome to Grab Distro Settings!\nSettings can be changed below.\nPress save when finished'
        self.asql = global_objs['SQL']
        self.main = Tk()

        # GUI Variables
        self.server = StringVar()
        self.database = StringVar()
        self.cat_tbl = StringVar()
        self.distro_tbl = StringVar()
        self.emp_tbl = StringVar()

        self.main.bind('<Destroy>', self.gui_cleanup)

    # static function to add setting to Local_Settings shelf files
    @staticmethod
    def add_setting(setting_list, val, key, encrypt=True):
        assert (key and setting_list)

        global_objs[setting_list].del_item(key)

        if val:
            global_objs[setting_list].add_item(key=key, val=val, encrypt=encrypt)

    # Static function to fill textbox in GUI
    @staticmethod
    def fill_textbox(setting_list, val, key):
        assert (key and val and setting_list)
        item = global_objs[setting_list].grab_item(key)

        if isinstance(item, CryptHandle):
            val.set(item.decrypt_text())

    # Close SQL socket upon GUI destruction
    def gui_cleanup(self, event):
        self.asql.close()

    # Stores list of SQL tables
    def check_table(self, table):
        if table:
            param = table.split('.')
            if len(param) == 2:
                return not self.asql.query('''
                    select
                        1
        
                    from information_schema.tables
                    where
                        table_schema = '{0}'
                            and
                        table_name = '{1}'
                        '''.format(param[0], param[1])).empty
        return False

    # Function to build GUI for settings
    def build_gui(self, header=None):
        # Change to custom header title if specified
        if header:
            self.header_text = header

        # Set GUI Geometry and GUI Title
        self.main.geometry('605x155+500+50')
        self.main.title('Grab Distro Settings')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        network_frame = LabelFrame(self.main, text='Network Settings', width=508, height=70)
        table_frame = LabelFrame(self.main, text='SQL Tables', width=508, height=70)
        buttons_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        network_frame.pack(fill="both")
        table_frame.pack(fill="both")
        buttons_frame.pack(fill='both')

        # Apply Network Labels & Input boxes to the Network_Frame
        #     SQL Server Input Box
        server_label = Label(self.main, text='Server:', padx=15, pady=7)
        server_txtbox = Entry(self.main, textvariable=self.server, width=30)
        server_label.pack(in_=network_frame, side=LEFT)
        server_txtbox.pack(in_=network_frame, side=LEFT)
        server_txtbox.bind('<FocusOut>', self.check_network)

        #     Server Database Input Box
        database_label = Label(self.main, text='Database:')
        database_txtbox = Entry(self.main, textvariable=self.database, width=30)
        database_txtbox.pack(in_=network_frame, side=RIGHT, pady=7, padx=15)
        database_label.pack(in_=network_frame, side=RIGHT)
        database_txtbox.bind('<KeyRelease>', self.check_network)

        # Apply Textbox and Labels for SQL Table section
        #     Cat Tbl Input Box
        cat_label = Label(table_frame, text='CAT Emp:', padx=5, pady=7)
        self.cat_txtbox = Entry(table_frame, textvariable=self.cat_tbl)
        cat_label.grid(row=0, column=0, sticky='w', pady=5, padx=3)
        self.cat_txtbox.grid(row=0, column=1, sticky='e', pady=5, padx=3)

        #     Strat Distro Table Input Box
        distro_label = Label(table_frame, text='Strat Distro:')
        self.distro_txtbox = Entry(table_frame, textvariable=self.distro_tbl)
        distro_label.grid(row=0, column=2, sticky='w', pady=5, padx=3)
        self.distro_txtbox.grid(row=0, column=3, sticky='e', pady=5, padx=3)

        #     Emp Table Input Box
        emp_label = Label(table_frame, text='Emp Tbl:')
        self.emp_txtbox = Entry(table_frame, textvariable=self.emp_tbl)
        emp_label.grid(row=0, column=4, sticky='w', pady=5, padx=3)
        self.emp_txtbox.grid(row=0, column=5, sticky='e', pady=5, padx=3)

        # Apply Buttons to the Buttons Frame
        #     Save Button
        self.save_settings_button = Button(buttons_frame, text='Save Settings', width=20, command=self.save_settings)
        self.save_settings_button.grid(row=0, column=0, pady=6, padx=15)

        #     Cancel Button
        cancel_button = Button(buttons_frame, text='Cancel', width=20, command=self.cancel)
        cancel_button.grid(row=0, column=1, pady=6, padx=260)

        self.fill_gui()

        # Show dialog
        self.main.mainloop()

    # Function to fill GUI textbox fields
    def fill_gui(self):
        self.fill_textbox('Settings', self.server, 'Server')
        self.fill_textbox('Settings', self.database, 'Database')

        if not self.server.get() or not self.database.get() or not self.asql.test_conn('alch'):
            self.save_settings_button.configure(state=DISABLED)
            self.distro_txtbox.configure(state=DISABLED)
            self.cat_txtbox.configure(state=DISABLED)
            self.emp_txtbox.configure(state=DISABLED)
        else:
            self.asql.connect('alch')
            self.fill_textbox('Local_Settings', self.distro_tbl, 'Distro_Tbl')
            self.fill_textbox('Local_Settings', self.cat_tbl, 'Cat_Tbl')
            self.fill_textbox('Local_Settings', self.emp_tbl, 'Emp_Tbl')

    # Function to check network settings if populated
    def check_network(self, event):
        if self.server.get() and self.database.get() and \
                (global_objs['Settings'].grab_item('Server') != self.server.get() or
                 global_objs['Settings'].grab_item('Database') != self.database.get()):
            self.asql.change_config(server=self.server.get(), database=self.database.get())

            if self.asql.test_conn('alch'):
                self.save_settings_button.configure(state=NORMAL)
                self.cat_txtbox.configure(state=NORMAL)
                self.distro_txtbox.configure(state=NORMAL)
                self.emp_txtbox.configure(state=NORMAL)
                self.add_setting('Settings', self.server.get(), 'Server')
                self.add_setting('Settings', self.database.get(), 'Database')
                self.asql.connect('alch')
            else:
                self.save_settings_button.configure(state=DISABLED)
                self.cat_txtbox.configure(state=DISABLED)
                self.distro_txtbox.configure(state=DISABLED)
                self.emp_txtbox.configure(state=DISABLED)
        else:
            self.save_settings_button.configure(state=DISABLED)
            self.cat_txtbox.configure(state=DISABLED)
            self.distro_txtbox.configure(state=DISABLED)
            self.emp_txtbox.configure(state=DISABLED)

    # Function to save settings when the Save Settings button is pressed
    def save_settings(self):
        if not self.cat_tbl.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for CAT Employee Table field',
                                 parent=self.main)
        elif not self.distro_tbl.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Strategic Distro Table field',
                                 parent=self.main)
        elif not self.emp_tbl.get():
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Employee Table field',
                                 parent=self.main)
        elif not self.check_table(self.cat_tbl.get()):
            messagebox.showerror('Field Empty Error!', 'CAT Employee Table does not exist in SQL Server',
                                 parent=self.main)
        elif not self.check_table(self.distro_tbl.get()):
            messagebox.showerror('Field Empty Error!', 'Strategic Distro Table does not exist in SQL Server',
                                 parent=self.main)
        elif not self.check_table(self.emp_tbl.get()):
            messagebox.showerror('Field Empty Error!', 'Employee Table does not exist in SQL Server',
                                 parent=self.main)
        else:
            self.add_setting('Local_Settings', self.cat_tbl.get(), 'Cat_Tbl')
            self.add_setting('Local_Settings', self.distro_tbl.get(), 'Distro_Tbl')
            self.add_setting('Local_Settings', self.emp_tbl.get(), 'Emp_Tbl')
            self.main.destroy()

    def cancel(self):
        self.main.destroy()


# Main loop routine to create GUI Settings
if __name__ == '__main__':
    obj = SettingsGUI()
    obj.build_gui()
