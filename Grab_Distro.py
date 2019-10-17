from Global import grabobjs
from Grab_Distro_Settings import SettingsGUI
from win32com.client import *

import os
import sys
import pandas as pd
import traceback

if getattr(sys, 'frozen', False):
    application_path = sys.executable
else:
    application_path = __file__

curr_dir = os.path.dirname(os.path.abspath(application_path))
main_dir = os.path.dirname(curr_dir)
global_objs = grabobjs(main_dir, 'Grab_Distro')


class ReadEmail:
    outlook = None
    user_list = list()

    def __init__(self):
        self.asql = global_objs['SQL']
        self.cat_tbl = global_objs['Local_Settings'].grab_item('Cat_Tbl')
        self.distro_tbl = global_objs['Local_Settings'].grab_item('Distro_Tbl')
        self.emp_tbl = global_objs['Local_Settings'].grab_item('Emp_Tbl')

    def connect(self):
        self.outlook = gencache.EnsureDispatch("Outlook.Application")

    def grab_contacts(self):
        if self.outlook:
            global_objs['Event_Log'].write_log('Grabbing Outlook contacts for strategic and costassurance distros')

            address_lists = self.outlook.Session.GetGlobalAddressList()
            entries = address_lists.AddressEntries

            for entry in entries:
                if entry.Type == 'EX' and entry.Name == 'strategic':
                    for member in entry.Members:
                        user = member.GetExchangeUser()
                        self.user_list.append([0, user.PrimarySmtpAddress, user.FirstName, user.Name, user.LastName,
                                               user.BusinessTelephoneNumber])
                elif entry.Type == 'EX' and entry.Name == 'CostAssurance':
                    for member in entry.Members:
                        user = member.GetExchangeUser()
                        self.user_list.append([1, user.PrimarySmtpAddress, user.FirstName, user.Name, user.LastName,
                                               user.BusinessTelephoneNumber])
        else:
            global_objs['Event_Log'].write_log('Unable to connect to outlook IMAP server. Please connect to outlook',
                                               'error')

    def upload_contacts(self):
        if len(self.user_list) > 0:
            global_objs['Event_Log'].write_log('Uploading %s contacts to temp table' % len(self.user_list))
            self.asql.connect('alch')
            df = pd.DataFrame(self.user_list, columns=['Table_Type', 'Email', 'First_Name', 'Maiden_Name', 'Last_Name',
                                                       'Phone'])
            df['Maiden_Name'] = df['Maiden_Name'].map(
                lambda x: str(x).split(' ')[1].replace('(', '').replace(')', '') if len(str(x).split(' ')) == 3 else None)
            df['Email'] = df['Email'].map(lambda x: str(x).upper())
            df['Phone'] = df['Phone'].map(lambda x: str(x).replace('-', '').strip())

            try:
                self.asql.upload(df, 'contacts_tmp')

                try:
                    global_objs['Event_Log'].write_log('Inserting & trimming records in production tables')

                    self.asql.execute('''
                        INSERT INTO {0}
                        (
                            Active,
                            Email,
                            First_Name,
                            Maiden_Name,
                            Last_Name,
                            Phone
                        )
                        SELECT
                            1,
                            Email,
                            First_Name,
                            Maiden_Name,
                            Last_Name,
                            Phone
                        
                        FROM contacts_tmp C
                        
                        WHERE
                            Table_Type = 0
                                AND
                            NOT EXISTS
                            (
                                SELECT
                                    1
                        
                                FROM {0} SD
                        
                                WHERE
                                    (
                                        C.First_Name = SD.First_Name
                                            AND
                                        C.Last_Name = SD.Last_Name
                                    )
                                        OR
                                    (
                                        C.First_Name = SD.First_Name
                                            AND
                                        C.Maiden_Name = SD.Last_Name
                                    )
                                        OR
                                    C.Email = SD.Email
                            );
                    '''.format(self.distro_tbl.decrypt_text()))

                    self.asql.execute('''
                        UPDATE SD
                            SET
                                SD.Employee_ID = EMP.Employee_ID
                                
                        FROM {0} SD
                        INNER JOIN {1} EMP
                        ON
                            CASE
                                WHEN SD.Maiden_Name IS NOT NULL THEN CONCAT(First_Name, ' ', Maiden_Name)
                                ELSE CONCAT(First_Name, ' ', Last_Name)
                            END = EMP.Name
                    '''.format(self.distro_tbl.decrypt_text(), self.emp_tbl.decrypt_text()))

                    self.asql.execute('''
                        UPDATE SD
                            SET
                                SD.Maiden_Name = C.Maiden_Name,
                                SD.Last_Name = C.Last_Name

                        FROM {0} SD
                        INNER JOIN contacts_tmp C
                        ON
                            C.Table_Type = 0
                                AND
                            C.First_Name = SD.First_Name
                                AND
                            C.Maiden_Name = SD.Last_Name

                        WHERE
                            SD.Active = 1
                    '''.format(self.distro_tbl.decrypt_text()))

                    self.asql.execute('''
                        UPDATE SD
                            SET
                                SD.Active = 1

                        FROM {0} SD

                        WHERE
                            SD.Active = 0
                                AND
                            EXISTS
                            (
                                SELECT
                                    1

                                FROM contacts_tmp C

                                WHERE
                                    Table_Type = 0
                                        AND
                                    (
                                        C.First_Name = SD.First_Name
                                            AND
                                        C.Last_Name = SD.Last_Name
                                    )
                                        OR
                                    (
                                        C.First_Name = SD.First_Name
                                            AND
                                        C.Maiden_Name = SD.Last_Name
                                    )
                                        OR
                                    C.Email = SD.Email
                            );
                    '''.format(self.distro_tbl.decrypt_text()))

                    self.asql.execute('''
                        UPDATE SD
                            SET
                                SD.Active = 0
                        
                        FROM {0} SD
                        
                        WHERE
                            SD.Active = 1
                                AND
                            NOT EXISTS
                            (
                                SELECT
                                    1
                        
                                FROM contacts_tmp C
                        
                                WHERE
                                    Table_Type = 0
                                        AND
                                    (
                                        C.First_Name = SD.First_Name
                                            AND
                                        C.Last_Name = SD.Last_Name
                                    )
                                        OR
                                    (
                                        C.First_Name = SD.First_Name
                                            AND
                                        C.Maiden_Name = SD.Last_Name
                                    )
                                        OR
                                    C.Email = SD.Email
                            );
                    '''.format(self.distro_tbl.decrypt_text()))

                    self.asql.execute('''
                        INSERT INTO {0}
                        (
                            Team,
                            Position,
                            Full_Name,
                            Initials,
                            Status,
                            Phone,
                            Email
                        )
                        SELECT
                            'CAT',
                            'Member',
                            CONCAT(C.First_Name, ' ', C.Last_Name) Full_Name,
                            CONCAT(LEFT(C.First_Name, 1), LEFT(C.Last_Name, CASE WHEN CAT2.CAT_ID IS NULL THEN 1 ELSE 2 END)) Initials,
                            'Active',
                            C.Phone,
                            C.Email
                        
                        FROM contacts_tmp C
                        LEFT JOIN {0} CAT2
                        ON
                            CONCAT(LEFT(C.First_Name, 1), LEFT(C.Last_Name, 1)) = CAT2.Initials
                        
                        WHERE
                            Table_Type = 1
                                AND
                            NOT EXISTS
                            (
                                SELECT
                                    1
                        
                                FROM {0} CAT
                        
                                WHERE
                                    CAT.Email = C.Email
                                        OR
                                    CAT.Full_Name = CONCAT(C.First_Name, ' ', C.Last_Name)
                                        OR
                                    CAT.Full_Name = CONCAT(C.First_Name, ' ', C.Maiden_Name)
                            );
                    '''.format(self.cat_tbl.decrypt_text()))

                    self.asql.execute('''
                        UPDATE CAT
                            SET
                                CAT.Status = 'InActive'
                        
                        FROM {0} CAT
                        
                        WHERE
                            CAT.Status = 'Active'
                                AND
                            TEAM IN ('CDA', 'DART', 'DART/CARP', 'Audit', 'PFA', 'CARP', 'CAT')
                                AND
                            NOT EXISTS
                            (
                                SELECT
                                    1
                        
                                FROM contacts_tmp C
                        
                                WHERE
                                    Table_Type = 1
                                        AND
                                    (
                                        CAT.Email = C.Email
                                            OR
                                        CAT.Full_Name = CONCAT(C.First_Name, ' ', C.Last_Name)
                                            OR
                                        CAT.Full_Name = CONCAT(C.First_Name, ' ', C.Maiden_Name)
                                    )
                            );
                    '''.format(self.cat_tbl.decrypt_text()))

                    global_objs['Event_Log'].write_log('Finished operations. Closing program')
                finally:
                    self.asql.execute('DROP TABLE contacts_tmp')

            finally:
                self.asql.close()

    def close(self):
        if self.outlook:
            del self.outlook


def check_settings():
    my_return = False
    obj = SettingsGUI()

    if not global_objs['Settings'].grab_item('Server') \
            or not global_objs['Settings'].grab_item('Database') \
            or not global_objs['Local_Settings'].grab_item('Cat_Tbl') \
            or not global_objs['Local_Settings'].grab_item('Distro_Tbl') \
            or not global_objs['Local_Settings'].grab_item('Emp_Tbl'):
        header_text = 'Welcome to Grab Distro!\nSettings haven''t been established.\nPlease fill out all empty fields below:'
        obj.build_gui(header_text)
    else:
        my_return = True

    obj.cancel()
    del obj
    return my_return


if __name__ == '__main__':
    myobj = ReadEmail()

    try:
        if check_settings():
            myobj.connect()
            myobj.grab_contacts()
            myobj.upload_contacts()
    except:
        global_objs['Event_Log'].write_log(traceback.format_exc(), 'error')
    finally:
        myobj.close()
