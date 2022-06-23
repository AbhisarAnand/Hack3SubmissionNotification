import time

import gspread
import numpy as np
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from validate_email_address import validate_email

from EmailSender import EmailSender
from constants import *


class Main:
    def __init__(self):
        self.api_json = GOOGLE_API_JSON
        while True:
            print("Running Program Again")
            # Open and read the Google sheets
            self.dataframe = self.read_sheets(URL)
            self.replace_headers()
            self.rename_column()
            self.replace_missing_values()
            self.dataframe = self.get_all_projects()
            self.dataframe = self.get_projects_to_notify()
            self.dataframe.reset_index(inplace=True)
            self.rows = self.dataframe.shape[0]
            for row in range(0, self.rows):
                self.send_email(row)
            self.write(self.dataframe['Project Name'].tolist())
            time.sleep(10)

    @staticmethod
    def read_sheets(url) -> pd.DataFrame:
        return pd.read_csv(url)

    def replace_headers(self):
        new_header = self.dataframe.iloc[0]
        self.dataframe = self.dataframe[1:]
        self.dataframe.columns = new_header

    def replace_missing_values(self):
        self.dataframe.replace(np.nan, "Missing", inplace=True)
        self.dataframe.replace("", "Missing", inplace=True)

    def rename_column(self, column="Time Slots"):
        self.dataframe.rename(columns={np.nan: column}, inplace=True)

    def get_all_projects(self):
        return self.dataframe.loc[self.dataframe['Project Name'] != "Missing"]

    def get_projects_to_notify(self):
        return self.dataframe.loc[self.dataframe['Already Notified'] == "Missing"]

    def get_emails(self, row):
        email_addresses = [self.dataframe["Team Member 1 Email"][row], self.dataframe["Team Member 2 Email"][row],
                           self.dataframe["Team Member 3 Email"][row], self.dataframe["Team Member 4 Email"][row]]
        email_str = ""
        for email in email_addresses:
            if validate_email(email, verify=True):
                email_str += ", " + email
        email_str = email_str[2:]
        if email_str == "":
            self.failure("No valid emails", row)
        print("EMAIL STRING: ", email_str)
        print(validate_email(email_str, verify=True))
        return email_str

    def get_team_members(self, row):
        team_members = [self.dataframe["Team Member 1"][row], self.dataframe["Team Member 2"][row],
                        self.dataframe["Team Member 3"][row], self.dataframe["Team Member 4"][row]]
        name_str = ""
        for name in team_members:
            if name != "" and name != "Missing":
                name_str += ", " + name
        name_str = name_str[2:]
        return name_str

    def send_email(self, row):
        print(self.dataframe["Project Name"][row])
        email_sender = EmailSender(self.get_emails(row), self.dataframe['Time Slots'][row],
                                   self.dataframe['Judging Groups'][row], self.get_team_members(row))
        email_sender.send_email()
        if email_sender.status == "fail":
            self.failure("Email Failed", row)
        elif email_sender.status == "success":
            self.dataframe["Already Notified"][row] = "Done"

    def failure(self, reason, row):
        self.dataframe["Already Notified"][row] = "Failed: " + reason
        print("FAILED: ", reason)

    def write(self, project_names):
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.api_json, scopes)
        file = gspread.authorize(credentials)  # authenticate the JSON key with gspread
        sheet = file.open("Hack3: Submission Form 2022 (Responses)")  # open sheet
        sheet = sheet.get_worksheet_by_id(
            2071051198)  # replace sheet_name with the name that corresponds to yours, e.g, it can be sheet1
        for project in project_names:
            sheet.update_cell(sheet.find(project, in_column=4).row, sheet.find(project, in_column=4).col - 1,
                              self.dataframe.loc[self.dataframe["Project Name"] == project].iloc[0]['Already Notified'])


if __name__ == '__main__':
    main = Main()
