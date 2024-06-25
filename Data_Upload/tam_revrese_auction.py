import pandas as pd
import os
from fuzzywuzzy import fuzz
import Data_Upload.data_mapping
import Data_Upload.iex_reverse_auction
import Data_Upload.pxil_reverse_auction
import Data_Upload.hpx_reverse_auction
from itertools import combinations


class tam_reverse_auction:
    def __init__(self):
        self.file_directory = r'C:\GNA\Data Upload'

    def merged_data(self):
        iex_file_path = os.path.join(self.file_directory, 'final_iex.xlsx')
        pxil_file_path = os.path.join(self.file_directory, 'final_pxil.xlsx')
        hpx_file_path = os.path.join(self.file_directory, 'final_hpx.xlsx')

        df1 = pd.read_excel(iex_file_path)
        df2 = pd.read_excel(pxil_file_path)
        df3 = pd.read_excel(hpx_file_path)

        df = pd.concat(([df3, df2, df1]))

        # Columns to Move
        column_to_move = ['Exchange Type', 'Auction No.', 'Auction Initiation Date', 'Auction Result Date', 'Buyer',
                          'Requisition No', 'Delivery Start Date', 'Delivery End Date', 'Delivery Start Time',
                          'Delivery End Time', 'Buy Total Quantity (in MW)', 'Buy Total Quantity (in MWH)',
                          'Buy Minimum Quantity (in MW)', 'Buy Minimum Quantity (in MWH)', 'Booking Quantity (in MW)',
                          'Minimum Booking Quantity (in MW)', 'Allocated Quantity (in MW)',
                          'Allocated Quantity (in MWH)', 'Accepted Price (in Rs./kWh)', 'Delivery Point',
                          'Total count of Delivery Days', ]
        remaining_column_order = [col for col in df.columns if col not in column_to_move]
        column_order = column_to_move + remaining_column_order
        df = df[column_order]

        df['Auction Initiation Date'] = df['Auction Initiation Date'].dt.date
        df['Auction Result Date'] = df['Auction Result Date'].dt.date
        df['Delivery Start Date'] = df['Delivery Start Date'].dt.date
        df['Delivery End Date'] = df['Delivery End Date'].dt.date

        output_path = os.path.join(self.file_directory, 'final_tam.xlsx')
        df.to_excel(output_path, index=False)
        print(f'File Merged {output_path}')

    def remove_duplicate_buyer_name(self):
        deep_file_path = os.path.join(self.file_directory, 'final_tam.xlsx')
        df = pd.read_excel(deep_file_path)

        unique_names = df['Buyer'].unique()

        for name in unique_names:
            for key, value in Data_Upload.data_mapping.tam_data_mapping.items():
                key_similarity_ratio = fuzz.ratio(name, key)
                value_similarity_ratio = fuzz.ratio(name, value)
                if key_similarity_ratio >= 91 or value_similarity_ratio >= 91:
                    df['Buyer'] = df['Buyer'].str.replace(name, key)

        # Convert column in time format
        df[['Delivery Start Time', 'Delivery End Time']] = df[['Delivery Start Time', 'Delivery End Time']].astype(
            str).replace('As per Seller Profile', '')
        df[['Delivery Start Time', 'Delivery End Time']] = df[['Delivery Start Time', 'Delivery End Time']].astype(
            str).replace('24:00', '23:59')
        df['Delivery Start Time'] = pd.to_datetime(df['Delivery Start Time'], format='%H:%M').dt.time
        df['Delivery End Time'] = pd.to_datetime(df['Delivery End Time'], format='%H:%M').dt.time
        df['Delivery End Time'] = df['Delivery End Time'].astype(str).replace('23:59:00', '23:59:59')
        df['Delivery End Time'] = pd.to_datetime(df['Delivery End Time'], format='%H:%M:%S').dt.time

        df['Allocated Quantity (in MW)'] = df['Allocated Quantity (in MW)'].astype(str).str.replace('TobeAllocated', '')
        df['Accepted Price (in Rs./kWh)'] = df['Accepted Price (in Rs./kWh)'].astype(str).str.replace('TobeAllocated',
                                                                                                      '')

        output_path = os.path.join(self.file_directory, 'final_tam.xlsx')
        df.to_excel(output_path, index=False)
        print(f'Duplicates removed from {deep_file_path}')

    def data_mapping(self):
        deep_file_path = os.path.join(self.file_directory, 'final_tam.xlsx')
        df = pd.read_excel(deep_file_path)

        for name, mane_to_change in Data_Upload.data_mapping.tam_data_mapping.items():
            df['Buyer'] = df['Buyer'].str.replace(name, mane_to_change)

        # Convert column in date format
        df['Auction Initiation Date'] = pd.to_datetime(df['Auction Initiation Date']).dt.date
        df['Auction Result Date'] = pd.to_datetime(df['Auction Result Date']).dt.date
        df['Delivery Start Date'] = pd.to_datetime(df['Delivery Start Date']).dt.date
        df['Delivery End Date'] = pd.to_datetime(df['Delivery End Date']).dt.date

        df['Exclusion Dates'] = df['Exclusion Dates'].astype(str)
        df['Accepted Price (in Rs./kWh)'] = df['Accepted Price (in Rs./kWh)'].astype(float)
        df['Allocated Quantity (in MW)'] = df['Allocated Quantity (in MW)'].astype(float)

        df['Exclusion Dates'] = df['Exclusion Dates'].astype(str)
        df['Auction No.'] = df['Auction No.'].astype(str)

        output_path = os.path.join(self.file_directory, 'final_tam.xlsx')
        df.to_excel(output_path, index=False)
        print(f'Data Mapping done file saved at {deep_file_path}')

    def get_data(self):
        Data_Upload.iex_reverse_auction.iex_reverse_auction().get_data()
        Data_Upload.pxil_reverse_auction.pxil_reverse_auction().get_data()
        Data_Upload.hpx_reverse_auction.hpx_reverse_auction().get_data()
        tam_reverse_auction().merged_data()
        tam_reverse_auction().remove_duplicate_buyer_name()
        tam_reverse_auction().data_mapping()


if __name__ == '__main__':
    tam = tam_reverse_auction()
    tam.get_data()
    pass
