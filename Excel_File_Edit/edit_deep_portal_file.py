import os
import pandas as pd


class deep_portal:
    def __init__(self):
        self.final_directory = r'C:\GNA\Data Upload\Final_Files'
        pass

    def edit_deep_portal_file(self):
        deep_portal_input_file = os.path.join(self.final_directory, 'deep_portal.xlsx')
        df = pd.read_excel(deep_portal_input_file)

        # Convert column in time format
        df['Auction Initiation Time'] = pd.to_datetime(df['Auction Initiation Time']).dt.time
        df['Auction Result Time'] = pd.to_datetime(df['Auction Result Time']).dt.time
        df['Delivery Start Time'] = pd.to_datetime(df['Delivery Start Time']).dt.time
        df['Delivery End Time'] = pd.to_datetime(df['Delivery End Time']).dt.time

        print(df['Delivery Start Time'])


if __name__ == '__main__':
    deep_portal = deep_portal()
    deep_portal.edit_deep_portal_file()
    pass
