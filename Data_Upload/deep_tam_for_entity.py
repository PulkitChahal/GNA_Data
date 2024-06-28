import os
import pandas as pd
import data_mapping
from datetime import datetime, timedelta


class deep_tam_for_particular_entity:

    def __init__(self):
        self.final_directory = r'C:\GNA\Data Upload'
        self.file_directory = os.path.join(self.final_directory, 'deep_tam_data.xlsx')
        pass

    def price_range(self, values):
        if max(values) == min(values):
            return f'{max(values)}'
        else:
            return f'{min(values)} - {max(values)}'

    def hpx_data_usable(self):
        input_file = os.path.join(self.final_directory, 'tam_auction_cassandra.xlsx')
        final_tam = pd.read_excel(input_file)
        df = pd.DataFrame(final_tam)
        df = df[df['exchange_type'].str.contains('HPX', na=False)]

        # Group and aggregate df
        df = df.groupby(
            ['exchange_type', 'auction_no', 'auction_result_date', 'buyer',
             'requisition_no', 'delivery_start_date', 'delivery_end_date', 'delivery_start_time',
             'delivery_end_time', 'buy_total_quantity_in_mw', 'buy_minimum_quantity_in_mw']).agg(
            {'booking_quantity_in_mw': 'sum',
             'allocated_quantity_in_mw': 'sum',
             'accepted_price_in_rs_per_kwh': (self.price_range)}).reset_index()
        print(df)

        # Calculate unallocated quantities
        df['unallocated_quantity_in_mw'] = df['buy_total_quantity_in_mw'] - df['allocated_quantity_in_mw']

        # Replace 'nan - nan' with 'NaN' in accepted prices
        df['accepted_price_in_rs_per_kwh'] = df['accepted_price_in_rs_per_kwh'].replace('nan - nan', 'NaN')

        output_path = os.path.join(self.final_directory, 'hpx_tam_usable.xlsx')
        df.to_excel(output_path, index=False)
        print(f'Data is Usable Now {output_path}')

    def tam_data_usable(self):
        input_file = os.path.join(self.final_directory, 'tam_auction_cassandra.xlsx')
        final_tam = pd.read_excel(input_file)

        df = pd.DataFrame(final_tam)

        # Copy the original DataFrame
        df1 = pd.DataFrame(df)

        # Group and aggregate df
        df = df.groupby(
            ['exchange_type', 'auction_no', 'auction_initiation_date', 'auction_result_date', 'buyer', 'requisition_no',
             'delivery_start_date', 'delivery_end_date', 'delivery_start_time', 'delivery_end_time',
             'buy_total_quantity_in_mw', 'buy_minimum_quantity_in_mw']).agg(
            {'allocated_quantity_in_mw': 'sum', 'accepted_price_in_rs_per_kwh': self.price_range}).reset_index()
        print(df)

        # # Ensure 'allocated_quantity_in_mwh' exists before attempting to group and aggregate
        # df1 = df1.groupby(['exchange_type', 'auction_no', 'auction_initiation_date', 'auction_result_date', 'buyer',
        #                    'requisition_no', 'delivery_start_date', 'delivery_end_date', 'delivery_start_time',
        #                    'delivery_end_time']).agg(
        #     {'allocated_quantity_in_mwh': 'sum', 'accepted_price_in_rs_per_kwh': self.price_range}).reset_index()
        # print(df1)
        #
        # # Calculate unallocated quantities
        # df['unallocated_quantity_in_mw'] = df['buy_total_quantity_in_mw'] - df['allocated_quantity_in_mw']
        # # df1['unallocated_quantity_in_mwh'] = df1['buy_total_quantity_in_mwh'] - df1['allocated_quantity_in_mwh']
        #
        # # Replace 'nan - nan' with 'NaN' in accepted prices
        # df['accepted_price_in_rs_per_kwh'] = df['accepted_price_in_rs_per_kwh'].replace('nan - nan', 'NaN')
        # df1['accepted_price_in_rs_per_kwh'] = df1['accepted_price_in_rs_per_kwh'].replace('nan - nan', 'NaN')
        #
        # # Merge df and df1 on the specified columns
        # common_columns = ['exchange_type', 'auction_no', 'auction_initiation_date', 'auction_result_date', 'buyer',
        #                   'delivery_start_date', 'delivery_end_date', 'delivery_start_time', 'delivery_end_time']
        #
        # # Merge on common columns
        # merged_df = pd.merge(df, df1, on=common_columns, how='outer', suffixes=('_mw', '_mwh'))
        #
        # merged_df['accepted_price_in_rs_per_kwh_mw'] = merged_df['accepted_price_in_rs_per_kwh_mw'].fillna(
        #     merged_df['accepted_price_in_rs_per_kwh_mwh'])
        #
        # merged_df.drop('accepted_price_in_rs_per_kwh_mwh', axis=1, inplace=True)
        #
        # merged_df.rename(columns={'accepted_price_in_rs_per_kwh_mw': 'accepted_price_in_rs_per_kwh'}, inplace=True)
        #
        # # Convert column in date format
        # merged_df['auction_initiation_date'] = pd.to_datetime(merged_df['auction_initiation_date']).dt.date
        # merged_df['auction_result_date'] = pd.to_datetime(merged_df['auction_result_date']).dt.date
        # merged_df['delivery_start_date'] = pd.to_datetime(merged_df['delivery_start_date']).dt.date
        # merged_df['delivery_end_date'] = pd.to_datetime(merged_df['delivery_end_date']).dt.date

        # Convert column in time format
        # merged_df['Delivery Start Time'] = pd.to_datetime(merged_df['Delivery Start Time'], format='%H:%M:%S').dt.time
        # merged_df['Delivery End Time'] = pd.to_datetime(merged_df['Delivery End Time'], format='%H:%M:%S').dt.time

        merged_df = merged_df.sort_values(by='auction_result_date', ascending=False)

        output_path = os.path.join(self.final_directory, 'tam_usable.xlsx')
        merged_df.to_excel(output_path, index=False)
        print(f'Data is Usable Now {output_path}')

    def merge_data(self):
        iex_file_path = os.path.join(self.final_directory, 'iex_usable.xlsx')
        hpx_file_path = os.path.join(self.final_directory, 'hpx_usable.xlsx')
        pxil_file_path = os.path.join(self.final_directory, 'pxil_usable.xlsx')
        deep_file_path = os.path.join(self.final_directory, 'deep_usable.xlsx')

        df1 = pd.read_excel(iex_file_path)
        df2 = pd.read_excel(hpx_file_path)
        df3 = pd.read_excel(pxil_file_path)
        df4 = pd.read_excel(deep_file_path)

        df = pd.concat([df4, df1, df2, df3])
        df['Auction No.'] = df['Auction No.'].astype(str)

        # Arrange Columns in Order
        column_to_move = ['Exchange Type', 'Auction No.', 'Auction Initiation Date', 'Auction Initiation Time',
                          'Auction Result Date', 'Auction Result Time', 'Buyer', 'Requisition No',
                          'Delivery Start Date', 'Delivery End Date', 'Delivery Start Time', 'Delivery End Time',
                          'Buy Total Quantity (in MW)', 'Buy Total Quantity (in MWH)', 'Buy Minimum Quantity (in MW)',
                          'Buy Minimum Quantity (in MWH)', 'Booking Quantity (in MW)', 'Allocated Quantity (in MW)',
                          'Allocated Quantity (in MWH)', 'Unallocated Quantity (in MW)',
                          'Unallocated Quantity (in MWH)', 'Booking Accepted Price (in Rs./kWh)',
                          'Accepted Price (in Rs./kWh)']
        remaining_column_order = [col for col in df.columns if col not in column_to_move]
        column_order = column_to_move + remaining_column_order
        df = df[column_order]

        df['Auction Initiation Date'] = df['Auction Initiation Date'].dt.date
        df['Auction Result Date'] = df['Auction Result Date'].dt.date
        df['Delivery Start Date'] = df['Delivery Start Date'].dt.date
        df['Delivery End Date'] = df['Delivery End Date'].dt.date

        df = df.sort_values(by=['Exchange Type', 'Auction Initiation Date', 'Auction No.', 'Buyer', 'Requisition No'],
                            ascending=[True, False, False, True, True])

        output_path = os.path.join(self.final_directory, 'deep_tam_data.xlsx')
        df.to_excel(output_path, index=False)
        print(f"Merged file saved to {output_path}")

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


if __name__ == '__main__':
    deep_tam_usable = deep_tam_for_particular_entity()
    deep_tam_usable.hpx_data_usable()
