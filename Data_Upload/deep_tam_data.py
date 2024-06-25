import os
import pandas as pd


class deep_tam_data:
    def __init__(self):
        self.final_directory = r'C:\GNA\Data Upload'
        pass

    def price_range(self, values):
        if max(values) == min(values):
            return f'{max(values)}'
        else:
            return f'{min(values)} - {max(values)}'

    def deep_portal_data_usable(self):
        file_path = os.path.join(self.final_directory, 'deep_portal.xlsx')
        df = pd.read_excel(file_path)
        print(df)

        df = df.groupby(['Exchange Type', 'Auction No.', 'Auction Initiation Date', 'Auction Initiation Time',
                         'Auction Result Date', 'Auction Result Time', 'Buyer', 'Requisition No', 'Delivery Start Date',
                         'Delivery End Date', 'Delivery Start Time', 'Delivery End Time',
                         'Buy Total Quantity (in MW)']).agg(
            {'Booking Quantity (in MW)': 'sum', 'Allocated Quantity (in MW)': 'sum',
             'Booking Accepted Price (in Rs./kWh)': self.price_range,
             'Accepted Price (in Rs./kWh)': self.price_range}).reset_index()

        df['Unallocated Quantity (in MW)'] = df['Buy Total Quantity (in MW)'] - df['Allocated Quantity (in MW)']
        df['Booking Accepted Price (in Rs./kWh)'] = df['Booking Accepted Price (in Rs./kWh)'].replace('nan - nan',
                                                                                                      'NaN')
        df['Accepted Price (in Rs./kWh)'] = df['Accepted Price (in Rs./kWh)'].replace('nan - nan', 'NaN')

        # Convert column in date format
        df['Auction Initiation Date'] = df['Auction Initiation Date'].dt.date
        df['Auction Result Date'] = df['Auction Result Date'].dt.date
        df['Delivery Start Date'] = df['Delivery Start Date'].dt.date
        df['Delivery End Date'] = df['Delivery End Date'].dt.date

        df = df.sort_values(by='Auction Initiation Date', ascending=False)

        output_path = os.path.join(self.final_directory, 'deep_usable.xlsx')
        df.to_excel(output_path, index=False)
        print(f'Data is Usable Now {output_path}')

    def hpx_tam_data_usable(self):
        file_path = os.path.join(self.final_directory, 'final_tam.xlsx')
        final_tam = pd.read_excel(file_path)
        df = final_tam[final_tam['Exchange Type'].str.contains('HPX', na=False)]
        print(df)

        df = df.groupby(
            ['Exchange Type', 'Auction No.', 'Auction Result Date', 'Buyer', 'Requisition No', 'Delivery Start Date',
             'Delivery End Date', 'Delivery Start Time', 'Delivery End Time', 'Buy Total Quantity (in MW)',
             'Buy Minimum Quantity (in MW)']).agg(
            {'Booking Quantity (in MW)': 'sum', 'Allocated Quantity (in MW)': 'sum',
             'Accepted Price (in Rs./kWh)': self.price_range}).reset_index()

        df['Unallocated Quantity (in MW)'] = df['Buy Total Quantity (in MW)'] - df['Allocated Quantity (in MW)']
        df['Accepted Price (in Rs./kWh)'] = df['Accepted Price (in Rs./kWh)'].replace('nan - nan', 'NaN')

        # Convert column in date format
        df['Auction Result Date'] = df['Auction Result Date'].dt.date
        df['Delivery Start Date'] = df['Delivery Start Date'].dt.date
        df['Delivery End Date'] = df['Delivery End Date'].dt.date

        # Convert column in time format
        df['Delivery Start Time'] = pd.to_datetime(df['Delivery Start Time'], format='%H:%M:%S').dt.time
        df['Delivery End Time'] = pd.to_datetime(df['Delivery End Time'], format='%H:%M:%S').dt.time

        df = df.sort_values(by='Auction Result Date', ascending=False)

        output_path = os.path.join(self.final_directory, 'hpx_usable.xlsx')
        df.to_excel(output_path, index=False)
        print(f'Data is Usable Now {output_path}')

    def pxil_tam_data_usable(self):
        file_path = os.path.join(self.final_directory, 'final_tam.xlsx')
        final_tam = pd.read_excel(file_path)
        df = final_tam[final_tam['Exchange Type'].str.contains('PXIL', na=False)]
        print(df)

        df = df.groupby(['Exchange Type', 'Auction No.', 'Auction Initiation Date', 'Auction Result Date', 'Buyer',
                         'Delivery Start Date', 'Delivery End Date', 'Delivery Start Time', 'Delivery End Time',
                         'Buy Total Quantity (in MW)', 'Buy Minimum Quantity (in MW)']).agg(
            {'Allocated Quantity (in MW)': 'sum', 'Accepted Price (in Rs./kWh)': self.price_range}).reset_index()

        df['Unallocated Quantity (in MW)'] = df['Buy Total Quantity (in MW)'] - df['Allocated Quantity (in MW)']
        df['Accepted Price (in Rs./kWh)'] = df['Accepted Price (in Rs./kWh)'].replace('nan - nan', 'NaN')

        # Convert column in date format
        df['Auction Initiation Date'] = df['Auction Initiation Date'].dt.date
        df['Auction Result Date'] = df['Auction Result Date'].dt.date
        df['Delivery Start Date'] = df['Delivery Start Date'].dt.date
        df['Delivery End Date'] = df['Delivery End Date'].dt.date

        # Convert column in time format
        df['Delivery Start Time'] = pd.to_datetime(df['Delivery Start Time'], format='%H:%M:%S').dt.time
        df['Delivery End Time'] = pd.to_datetime(df['Delivery End Time'], format='%H:%M:%S').dt.time

        df = df.sort_values(by='Auction Result Date', ascending=False)

        output_path = os.path.join(self.final_directory, 'pxil_usable.xlsx')
        df.to_excel(output_path, index=False)
        print(f'Data is Usable Now {output_path}')

    def iex_tam_data_usable(self):
        file_path = os.path.join(self.final_directory, 'final_tam.xlsx')
        final_tam = pd.read_excel(file_path)
        df = final_tam[final_tam['Exchange Type'].str.contains('IEX', na=False)]
        print(df)

        # Copy the original DataFrame
        df1 = df.copy()

        # Group and aggregate df
        df = df.groupby(['Exchange Type', 'Auction No.', 'Auction Initiation Date', 'Auction Result Date', 'Buyer',
                         'Delivery Start Date', 'Delivery End Date', 'Delivery Start Time', 'Delivery End Time',
                         'Buy Total Quantity (in MW)', 'Buy Minimum Quantity (in MW)']).agg(
            {'Allocated Quantity (in MW)': 'sum', 'Accepted Price (in Rs./kWh)': self.price_range}).reset_index()
        # print(df)

        # Group and aggregate df1
        df1 = df1.groupby(['Exchange Type', 'Auction No.', 'Auction Initiation Date', 'Auction Result Date', 'Buyer',
                           'Delivery Start Date', 'Delivery End Date', 'Delivery Start Time', 'Delivery End Time',
                           'Buy Total Quantity (in MWH)', 'Buy Minimum Quantity (in MWH)']).agg(
            {'Allocated Quantity (in MWH)': 'sum', 'Accepted Price (in Rs./kWh)': self.price_range}).reset_index()
        # print(df1)

        # Calculate unallocated quantities
        df['Unallocated Quantity (in MW)'] = df['Buy Total Quantity (in MW)'] - df['Allocated Quantity (in MW)']
        df1['Unallocated Quantity (in MWH)'] = df1['Buy Total Quantity (in MWH)'] - df1['Allocated Quantity (in MWH)']

        # Replace 'nan - nan' with 'NaN' in accepted prices
        df['Accepted Price (in Rs./kWh)'] = df['Accepted Price (in Rs./kWh)'].replace('nan - nan', 'NaN')
        df1['Accepted Price (in Rs./kWh)'] = df1['Accepted Price (in Rs./kWh)'].replace('nan - nan', 'NaN')

        # Merge df and df1 on the specified columns
        common_columns = ['Exchange Type', 'Auction No.', 'Auction Initiation Date', 'Auction Result Date', 'Buyer',
                          'Delivery Start Date', 'Delivery End Date', 'Delivery Start Time', 'Delivery End Time']

        # Merge on common columns
        merged_df = pd.merge(df, df1, on=common_columns, how='outer', suffixes=('_MW', '_MWH'))

        merged_df['Accepted Price (in Rs./kWh)_MW'] = merged_df['Accepted Price (in Rs./kWh)_MW'].fillna(
            merged_df['Accepted Price (in Rs./kWh)_MWH'])

        merged_df.drop('Accepted Price (in Rs./kWh)_MWH', axis=1, inplace=True)

        merged_df.rename(columns={'Accepted Price (in Rs./kWh)_MW': 'Accepted Price (in Rs./kWh)'}, inplace=True)

        # Convert column in date format
        merged_df['Auction Initiation Date'] = merged_df['Auction Initiation Date'].dt.date
        merged_df['Auction Result Date'] = merged_df['Auction Result Date'].dt.date
        merged_df['Delivery Start Date'] = merged_df['Delivery Start Date'].dt.date
        merged_df['Delivery End Date'] = merged_df['Delivery End Date'].dt.date

        # Convert column in time format
        merged_df['Delivery Start Time'] = pd.to_datetime(merged_df['Delivery Start Time'], format='%H:%M:%S').dt.time
        merged_df['Delivery End Time'] = pd.to_datetime(merged_df['Delivery End Time'], format='%H:%M:%S').dt.time

        merged_df = merged_df.sort_values(by='Auction Result Date', ascending=False)

        output_path = os.path.join(self.final_directory, 'iex_usable.xlsx')
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

    def get_data_for_particular_entity(self):
        deep_tam_file_path = os.path.join(self.final_directory, 'deep_tam_data.xlsx')
        final_file = pd.read_excel(deep_tam_file_path)

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

    def making_data_ready_for_ppt_report(self):
        deep_tam_file_path = os.path.join(self.final_directory,
                                          'Maharashtra State Electricity Distribution Co Limited (4).xlsx')
        df = pd.read_excel(deep_tam_file_path)
        print(df.columns)

        df['start_date'] = pd.to_datetime(df['Delivery Start Date']).dt.strftime('%d %b %y')
        df['end_date'] = pd.to_datetime(df['Delivery End Date']).dt.strftime('%d %b %y')
        print(df['start_date'])
        print(df['end_date'])

        df = df.sort_values(by='Delivery Start Date', ascending=True)

        df['delivery_date'] = df['start_date'].astype(str) + '-' + df['end_date'].astype(str)
        df['delivery_period'] = df['Delivery Start Time'].astype(str) + '-' + df['Delivery End Time']

        if 'Buy Total Quantity (in MW)' and 'Buy Total Quantity (in MWH)' in df.columns:
            columns_to_add = ['Exchange Type', 'delivery_date', 'delivery_period', 'Buy Total Quantity (in MW)',
                'Buy Total Quantity (in MWH)'
                'Allocated Quantity (in MW)', 'Allocated Quantity (in MWH)'
                                              'Accepted Price (in Rs./kWh)', ]
            df = df[columns_to_add]

        elif 'Buy Total Quantity (in MWH)' in df.columns:
            columns_to_add = ['Exchange Type', 'delivery_date', 'delivery_period', 'Buy Total Quantity (in MWH)'
                                                                                   'Allocated Quantity (in MWH)'
                                                                                   'Accepted Price (in Rs./kWh)', ]
            df = df[columns_to_add]

        else:
            columns_to_add = ['Exchange Type', 'delivery_date', 'delivery_period', 'Buy Total Quantity (in MW)',
                'Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)', ]
            df = df[columns_to_add]

        output_path = os.path.join(self.final_directory, 'maharashtra.xlsx')
        df.to_excel(output_path, index=False)
        print(f"Merged file saved to {output_path}")

    def get_data(self):
        # deep_tam_data().deep_portal_data_usable()
        # deep_tam_data().hpx_tam_data_usable()
        # deep_tam_data().pxil_tam_data_usable()
        # deep_tam_data().iex_tam_data_usable()
        # deep_tam_data().merge_data()
        # deep_tam_data().get_data_for_particular_entity()
        deep_tam_data().making_data_ready_for_ppt_report()
        pass


if __name__ == '__main__':
    deep_tam = deep_tam_data()
    deep_tam.get_data()
    pass
