import os
import pandas as pd
import data_mapping
from datetime import datetime, timedelta

class deep_tam_for_particular_entity:

    def __init__(self):
        self.final_directory = r'C:\GNA\Data Upload'
        self.file_directory = os.path.join(self.final_directory,'deep_tam_data.xlsx')
        pass

    def make_changes_to_use_deep_tam_data(self):
        deep_file_path = os.path.join(self.file_directory)
        df = pd.read_excel(deep_file_path)

        df['Auction Initiation Date'] = pd.to_datetime(df['Auction Initiation Date']).dt.date
        df['Auction Result Date'] = pd.to_datetime(df['Auction Result Date']).dt.date
        df['Delivery Start Date'] = pd.to_datetime(df['Delivery Start Date']).dt.date
        df['Delivery End Date'] = pd.to_datetime(df['Delivery End Date']).dt.date

        # Convert time columns to datetime.time, fill NaT with default time
        df['Delivery Start Time'] = pd.to_datetime(df['Delivery Start Time'], errors='coerce').dt.time.fillna(
            datetime.min.time())
        df['Delivery End Time'] = pd.to_datetime(df['Delivery End Time'], errors='coerce').dt.time.fillna(
            datetime.min.time())

        df['Total Days'] = df['Delivery End Date'] - df['Delivery Start Date']
        df['Total Days'] = df['Total Days'].astype(str).str.split(' day').str[0]
        df['Total Days'] = df['Total Days'].str.replace('0:00:00', '0')
        df['Total Days'] = df['Total Days'].astype(int) + 1
        # print(df['Total Days'])

        # Calculate total hours by combining date and time into datetime
        df[['Start Time', 'End Time']] = df[['Delivery Start Time', 'Delivery End Time']]
        df[['Start Time', 'End Time']] = df[['Start Time', 'End Time']].astype(str)
        df[['Start Time', 'End Time']] = df[['Start Time', 'End Time']].astype(str).apply(
            lambda x: x.replace('23:59:59', '24:00:00'))
        df[['Start Time', 'End Time']] = df[['Start Time', 'End Time']].astype(str).apply(
            lambda x: x.str.split(':').str[0])
        df[['Start Time', 'End Time']] = df[['Start Time', 'End Time']].astype(int)
        df['Daily Hours'] = df['End Time'] - df['Start Time']
        # print(df['Daily Hours'])

        # Calculate Total Hours
        df['Total Hours'] = df['Total Days'] * df['Daily Hours']
        # print(df['Total Hours'])

        df['converted_to_mwh'] = df['Total Days'] * df['Total Hours']

        output_path = os.path.join(self.file_directory, 'deep_tam_data.xlsx')
        df.to_excel(output_path, index=False)
        print(f'Changes done file saved at {output_path}')

    def get_data_for_particular_buyer(self):
        deep_file_path = os.path.join(self.file_directory)
        final_file = pd.read_excel(deep_file_path)

        print(final_file['Buyer'].unique())
        buyer_name = ['Northern Railway-Haryana', 'Haryana Power Purchase Centre']
        # df = final_file[final_file['Buyer'].str.contains(buyer_name, na=False)]
        df = final_file[final_file['Buyer'].isin(buyer_name)]
        print(df)

        # Convert column in date format
        df['Auction Initiation Date'] = pd.to_datetime(df['Auction Initiation Date']).dt.date
        df['Auction Result Date'] = pd.to_datetime(df['Auction Result Date']).dt.date
        df['Delivery Start Date'] = pd.to_datetime(df['Delivery Start Date']).dt.date
        df['Delivery End Date'] = pd.to_datetime(df['Delivery End Date']).dt.date

        df['Requisition No'] = df['Requisition No'].astype(str).str.split(' ').str[0]
        df['Requisition No'] = df['Requisition No'].astype(str).str.split('\(').str[0]
        df['Requisition No'] = df['Requisition No'].str.replace('nan', '0')
        print(df['Requisition No'])
        df['Requisition No'] = df['Requisition No'].astype(int)

        df = df.sort_values(by=['Exchange Type', 'Auction Initiation Date', 'Auction No.', 'Buyer', 'Requisition No',
                                'Buy Total Quantity (in MW)'], ascending=[True, False, False, True, True, True])

        output_path = os.path.join(self.final_directory, f'haryana.xlsx')
        df.to_excel(output_path, index=False)
        print(f'Data is Usable Now {output_path}')

    def edit_merged_data(self):
        file_path = os.path.join(self.final_directory, 'haryana1.xlsx')
        df = pd.read_excel(file_path)

        df[['Minimum Accepted Price (in Rs./kWh)', 'Maximum Accepted Price (in Rs./kWh)']] = df[
            'Accepted Price (in Rs./kWh)'].astype(str).str.split('-', expand=True)
        df['Maximum Accepted Price (in Rs./kWh)'] = df['Maximum Accepted Price (in Rs./kWh)'].ffill(
            df['Accepted Price (in Rs./kWh)'])

        # df = df.groupby(
        # 	['Exchange Type',
        # 	 'Auction No.',
        # 	 'Auction Initiation Date',
        # 	 'Auction Initiation Time',
        # 	 'Auction Result Date',
        # 	 'Auction Result Time',
        # 	 'Buyer',
        # 	 'Requisition No',
        # 	 'Delivery Start Date',
        # 	 'Delivery End Date',
        # 	 'Delivery Start Time',
        # 	 'Delivery End Time',
        # 	 'Buy Total Quantity (in MW)']).agg({
        # 	'Booking Quantity (in MW)': 'sum',
        # 	'Allocated Quantity (in MW)': 'sum',
        # 	'Booking Accepted Price (in Rs./kWh)': self.price_range,
        # 	'Accepted Price (in Rs./kWh)': self.price_range}).reset_index()

        output_path = os.path.join(self.final_directory, f'haryana2.xlsx')
        df.to_excel(output_path, index=False)
        print(f'Data is Usable Now {output_path}')

    def make_graph_ready_format(self):
        file_path = os.path.join(self.final_directory, 'haryana1.xlsx')
        df = pd.read_excel(file_path)

        # Extract the month from 'Delivery Start Date'
        df['Delivery Month'] = df['Delivery Start Date'].dt.strftime('%b')
        df['Delivery Year'] = df['Delivery Start Date'].dt.strftime('%Y')
        df['Delivery Month-Year'] = df['Delivery Start Date'].dt.strftime('%m-%Y')

        df = df.groupby(['Exchange Type', 'Buyer', 'Delivery Month', 'Delivery Year', 'Delivery Month-Year']).agg(
            {'Buy Total Quantity (in MW)': 'mean', 'Booking Quantity (in MW)': 'mean',
                'Allocated Quantity (in MW)': 'mean', 'Unallocated Quantity (in MW)': 'mean',
                'Minimum Accepted Price (in Rs./kWh)': 'mean',
                'Maximum Accepted Price (in Rs./kWh)': 'mean'}).reset_index()

        df = df.sort_values(by='Delivery Month-Year')

        output_path = os.path.join(self.final_directory, f'haryana2.xlsx')
        df.to_excel(output_path, index=False)
        print(f'Data is Usable Now {output_path}')