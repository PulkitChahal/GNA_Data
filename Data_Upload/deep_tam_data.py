from itertools import combinations
import pandas as pd
import os
from fuzzywuzzy import fuzz
import data_mapping
from datetime import datetime, timedelta


class deep_tam_data:
	def __init__(self):
		self.file_directory = r'C:\GNA\Data Upload'
		pass
	
	def price_range(self, values):
		if max(values) == min(values):
			return f'{max(values)}'
		else:
			return f'{min(values)} - {max(values)}'
	
	def deep_portal_data_usable(self):
		file_path = os.path.join(self.file_directory, 'deep_portal.xlsx')
		df = pd.read_excel(file_path)
		print(df)
		
		df = df.groupby(
			['Exchange Type',
			 'Auction No.',
			 'Auction Initiation Date',
			 'Auction Initiation Time',
			 'Auction Result Date',
			 'Auction Result Time',
			 'Buyer',
			 'Requisition No',
			 'Delivery Start Date',
			 'Delivery End Date',
			 'Delivery Start Time',
			 'Delivery End Time',
			 'Buy Total Quantity (in MW)']).agg({
			'Booking Quantity (in MW)': 'sum',
			'Allocated Quantity (in MW)': 'sum',
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
		
		output_path = os.path.join(self.file_directory, 'deep_usable.xlsx')
		df.to_excel(output_path, index=False)
		print(f'Data is Usable Now {output_path}')
	
	def hpx_tam_data_usable(self):
		file_path = os.path.join(self.file_directory, 'final_tam.xlsx')
		final_tam = pd.read_excel(file_path)
		df = final_tam[final_tam['Exchange Type'].str.contains('HPX', na=False)]
		print(df)
		
		df = df.groupby(
			['Exchange Type',
			 'Auction No.',
			 'Auction Result Date',
			 'Buyer',
			 'Requisition No',
			 'Delivery Start Date',
			 'Delivery End Date',
			 'Delivery Start Time',
			 'Delivery End Time',
			 'Buy Total Quantity (in MW)',
			 'Buy Minimum Quantity (in MW)']).agg({
			'Booking Quantity (in MW)': 'sum',
			'Allocated Quantity (in MW)': 'sum',
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
		
		output_path = os.path.join(self.file_directory, 'hpx_usable.xlsx')
		df.to_excel(output_path, index=False)
		print(f'Data is Usable Now {output_path}')
	
	def pxil_tam_data_usable(self):
		file_path = os.path.join(self.file_directory, 'final_tam.xlsx')
		final_tam = pd.read_excel(file_path)
		df = final_tam[final_tam['Exchange Type'].str.contains('PXIL', na=False)]
		print(df)
		
		df = df.groupby(
			['Exchange Type',
			 'Auction No.',
			 'Auction Initiation Date',
			 'Auction Result Date',
			 'Buyer',
			 'Delivery Start Date',
			 'Delivery End Date',
			 'Delivery Start Time',
			 'Delivery End Time',
			 'Buy Total Quantity (in MW)',
			 'Buy Minimum Quantity (in MW)']).agg({
			'Allocated Quantity (in MW)': 'sum',
			'Accepted Price (in Rs./kWh)': self.price_range}).reset_index()
		
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
		
		output_path = os.path.join(self.file_directory, 'pxil_usable.xlsx')
		df.to_excel(output_path, index=False)
		print(f'Data is Usable Now {output_path}')
	
	def iex_tam_data_usable(self):
		file_path = os.path.join(self.file_directory, 'final_tam.xlsx')
		final_tam = pd.read_excel(file_path)
		df = final_tam[final_tam['Exchange Type'].str.contains('IEX', na=False)]
		print(df)
		
		# Copy the original DataFrame
		df1 = df.copy()
		
		# Group and aggregate df
		df = df.groupby([
			'Exchange Type',
			'Auction No.',
			'Auction Initiation Date',
			'Auction Result Date',
			'Buyer',
			'Delivery Start Date',
			'Delivery End Date',
			'Delivery Start Time',
			'Delivery End Time',
			'Buy Total Quantity (in MW)',
			'Buy Minimum Quantity (in MW)'
		]).agg({
			'Allocated Quantity (in MW)': 'sum',
			'Accepted Price (in Rs./kWh)': self.price_range
		}).reset_index()
		# print(df)
		
		# Group and aggregate df1
		df1 = df1.groupby([
			'Exchange Type',
			'Auction No.',
			'Auction Initiation Date',
			'Auction Result Date',
			'Buyer',
			'Delivery Start Date',
			'Delivery End Date',
			'Delivery Start Time',
			'Delivery End Time',
			'Buy Total Quantity (in MWH)',
			'Buy Minimum Quantity (in MWH)'
		]).agg({
			'Allocated Quantity (in MWH)': 'sum',
			'Accepted Price (in Rs./kWh)': self.price_range
		}).reset_index()
		# print(df1)
		
		# Calculate unallocated quantities
		df['Unallocated Quantity (in MW)'] = df['Buy Total Quantity (in MW)'] - df['Allocated Quantity (in MW)']
		df1['Unallocated Quantity (in MWH)'] = df1['Buy Total Quantity (in MWH)'] - df1['Allocated Quantity (in MWH)']
		
		# Replace 'nan - nan' with 'NaN' in accepted prices
		df['Accepted Price (in Rs./kWh)'] = df['Accepted Price (in Rs./kWh)'].replace('nan - nan', 'NaN')
		df1['Accepted Price (in Rs./kWh)'] = df1['Accepted Price (in Rs./kWh)'].replace('nan - nan', 'NaN')
		
		# Merge df and df1 on the specified columns
		common_columns = [
			'Exchange Type',
			'Auction No.',
			'Auction Initiation Date',
			'Auction Result Date',
			'Buyer',
			'Delivery Start Date',
			'Delivery End Date',
			'Delivery Start Time',
			'Delivery End Time'
		]
		
		# Merge on common columns
		merged_df = pd.merge(df, df1, on=common_columns, how='outer', suffixes=('_MW', '_MWH'))
		
		merged_df['Accepted Price (in Rs./kWh)_MW'] = merged_df['Accepted Price (in Rs./kWh)_MW'].fillna(
			merged_df['Accepted Price (in Rs./kWh)_MWH'])
		
		merged_df.drop('Accepted Price (in Rs./kWh)_MWH', axis=1, inplace=True)
		
		merged_df.rename(columns={
			'Accepted Price (in Rs./kWh)_MW': 'Accepted Price (in Rs./kWh)'
		}, inplace=True)
		
		# Convert column in date format
		merged_df['Auction Initiation Date'] = merged_df['Auction Initiation Date'].dt.date
		merged_df['Auction Result Date'] = merged_df['Auction Result Date'].dt.date
		merged_df['Delivery Start Date'] = merged_df['Delivery Start Date'].dt.date
		merged_df['Delivery End Date'] = merged_df['Delivery End Date'].dt.date
		
		# Convert column in time format
		merged_df['Delivery Start Time'] = pd.to_datetime(merged_df['Delivery Start Time'], format='%H:%M:%S').dt.time
		merged_df['Delivery End Time'] = pd.to_datetime(merged_df['Delivery End Time'], format='%H:%M:%S').dt.time
		
		merged_df = merged_df.sort_values(by='Auction Result Date', ascending=False)
		
		output_path = os.path.join(self.file_directory, 'iex_usable.xlsx')
		merged_df.to_excel(output_path, index=False)
		print(f'Data is Usable Now {output_path}')
	
	def merge_data(self):
		iex_file_path = os.path.join(self.file_directory, 'iex_usable.xlsx')
		hpx_file_path = os.path.join(self.file_directory, 'hpx_usable.xlsx')
		pxil_file_path = os.path.join(self.file_directory, 'pxil_usable.xlsx')
		deep_file_path = os.path.join(self.file_directory, 'deep_usable.xlsx')
		
		df1 = pd.read_excel(iex_file_path)
		df2 = pd.read_excel(hpx_file_path)
		df3 = pd.read_excel(pxil_file_path)
		df4 = pd.read_excel(deep_file_path)
		
		df = pd.concat([df4, df1, df2, df3])
		df['Auction No.'] = df['Auction No.'].astype(str)
		
		# Arrange Columns in Order
		column_to_move = [
			'Exchange Type',
			'Auction No.',
			'Auction Initiation Date',
			'Auction Initiation Time',
			'Auction Result Date',
			'Auction Result Time',
			'Buyer',
			'Requisition No',
			'Delivery Start Date',
			'Delivery End Date',
			'Delivery Start Time',
			'Delivery End Time',
			'Buy Total Quantity (in MW)',
			'Buy Total Quantity (in MWH)',
			'Buy Minimum Quantity (in MW)',
			'Buy Minimum Quantity (in MWH)',
			'Booking Quantity (in MW)',
			'Allocated Quantity (in MW)',
			'Allocated Quantity (in MWH)',
			'Unallocated Quantity (in MW)',
			'Unallocated Quantity (in MWH)',
			'Booking Accepted Price (in Rs./kWh)',
			'Accepted Price (in Rs./kWh)'
		]
		remaining_column_order = [col for col in df.columns if col not in column_to_move]
		column_order = column_to_move + remaining_column_order
		df = df[column_order]
		
		df['Auction Initiation Date'] = df['Auction Initiation Date'].dt.date
		df['Auction Result Date'] = df['Auction Result Date'].dt.date
		df['Delivery Start Date'] = df['Delivery Start Date'].dt.date
		df['Delivery End Date'] = df['Delivery End Date'].dt.date
		
		df = df.sort_values(by=['Exchange Type', 'Auction Initiation Date', 'Auction No.', 'Buyer', 'Requisition No'],
		                    ascending=[True, False, False, True, True])
		
		output_path = os.path.join(self.file_directory, 'deep_tam_data.xlsx')
		df.to_excel(output_path, index=False)
		print(f"Merged file saved to {output_path}")
	
	def make_changes_to_use_deep_tam_data(self):
		deep_file_path = os.path.join(self.file_directory, 'deep_tam_data.xlsx')
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
		deep_file_path = os.path.join(self.file_directory, 'deep_tam_data.xlsx')
		final_file = pd.read_excel(deep_file_path)
		
		print(final_file['Buyer'].unique())
		buyer_name = ['Northern Railway-Haryana','Haryana Power Purchase Centre']
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
		                        'Buy Total Quantity (in MW)'],
		                    ascending=[True, False, False, True, True, True])
		
		output_path = os.path.join(self.file_directory, f'haryana.xlsx')
		df.to_excel(output_path, index=False)
		print(f'Data is Usable Now {output_path}')
	
	def edit_merged_data(self):
		file_path = os.path.join(self.file_directory, 'haryana1.xlsx')
		df = pd.read_excel(file_path)
		
		df[['Minimum Accepted Price (in Rs./kWh)', 'Maximum Accepted Price (in Rs./kWh)']] = df['Accepted Price (in Rs./kWh)'].astype(str).str.split('-',expand=True)
		df['Maximum Accepted Price (in Rs./kWh)'] = df['Maximum Accepted Price (in Rs./kWh)'].ffill(df['Accepted Price (in Rs./kWh)'])
		
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
		
		output_path = os.path.join(self.file_directory, f'haryana2.xlsx')
		df.to_excel(output_path, index=False)
		print(f'Data is Usable Now {output_path}')
	
	def make_graph_ready_format(self):
		file_path = os.path.join(self.file_directory, 'haryana1.xlsx')
		df = pd.read_excel(file_path)
		
		# Extract the month from 'Delivery Start Date'
		df['Delivery Month'] = df['Delivery Start Date'].dt.strftime('%b')
		df['Delivery Year'] = df['Delivery Start Date'].dt.strftime('%Y')
		df['Delivery Month-Year'] = df['Delivery Start Date'].dt.strftime('%m-%Y')
		
		df = df.groupby(
			['Exchange Type',
			 'Buyer',
			 'Delivery Month',
			 'Delivery Year',
			 'Delivery Month-Year']).agg({
			'Buy Total Quantity (in MW)': 'mean',
			'Booking Quantity (in MW)': 'mean',
			'Allocated Quantity (in MW)': 'mean',
			'Unallocated Quantity (in MW)': 'mean',
			'Minimum Accepted Price (in Rs./kWh)':'mean',
			'Maximum Accepted Price (in Rs./kWh)':'mean'}).reset_index()
		
		df = df.sort_values(by='Delivery Month-Year')
		
		# Convert column in date format
		# df['Auction Initiation Date'] = pd.to_datetime(df['Auction Initiation Date']).dt.date
		# df['Auction Result Date'] = pd.to_datetime(df['Auction Result Date']).dt.date
		# df['Delivery Start Date'] = pd.to_datetime(df['Delivery Start Date']).dt.date
		# df['Delivery End Date'] = pd.to_datetime(df['Delivery End Date']).dt.date
		
		output_path = os.path.join(self.file_directory, f'haryana2.xlsx')
		df.to_excel(output_path, index=False)
		print(f'Data is Usable Now {output_path}')
		
		
	
	def get_data(self):
		# deep_tam_data.deep_portal_data_usable(self)
		# deep_tam_data.hpx_tam_data_usable(self)
		# deep_tam_data.pxil_tam_data_usable(self)
		# deep_tam_data.iex_tam_data_usable(self)
		# deep_tam_data.merge_data(self)
		# deep_tam_data.get_data_for_particular_buyer(self)
		# deep_tam_data().edit_merged_data()
		deep_tam_data().make_graph_ready_format()
		# deep_tam_data.make_changes_to_use_deep_tam_data(self)
		pass


if __name__ == '__main__':
	deep_tam = deep_tam_data()
	deep_tam.get_data()
	pass
