import os
import time
import numpy as np
import requests
import pandas as pd
import tabula
from datetime import datetime, timedelta
import openpyxl
from openpyxl.utils import column_index_from_string

file_directory = r'C:\GNA\Coding\Coal Data\Download Files'
output_directory = r'C:\GNA\Coding\Coal Data\edited_xlsx_files'
main_directory = r'C:\GNA\Coding\Coal Data'
error_log_file = r'C:\GNA\Coding\Coal Data\error_log.xlsx'


class coal_npp:
	def __init__(self):
		self.final_directory = r'C:\GNA\Data Upload'
		pass
	
	def coal_data_download(self):
		file_directory = r'C:\GNA\Coding\Coal Data\Download Files'
		if os.path.exists(file_directory):
			for file in os.listdir(file_directory):
				file_path_full = os.path.join(file_directory, file)
				if os.path.isfile(file_path_full):
					os.remove(file_path_full)
		else:
			os.makedirs(file_directory)
		
		base_date = datetime.now()
		# end_date = datetime(2019, 1, 1)
		end_date = base_date - timedelta(20)
		
		while base_date > end_date:
			start_date = base_date.strftime('%d-%m-%Y')
			start_date_2 = base_date.strftime('%Y-%m-%d')
			
			url = f'https://npp.gov.in/public-reports/cea/daily/fuel/{start_date}/dailyCoal1-{start_date_2}.xlsx'
			response = requests.get(url)
			
			if response.status_code == 200:
				with open(os.path.join(file_directory, f'Coal Data-{start_date_2}.xlsx'), 'wb') as f:
					f.write(response.content)
				
				print(f"Downloaded data for date: {start_date_2}")
			else:
				print(f"No data available for date: {start_date_2}")
			
			base_date -= timedelta(days=1)
		
		if base_date == end_date:
			print("Reached end date. Exiting loop.")
	
	def edit_xlsx_file_using_pandas(self):
		output_directory = r'C:\GNA\Coding\Coal Data\edited_xlsx_files'
		if os.path.exists(output_directory):
			for file in os.listdir(output_directory):
				file_path_full = os.path.join(output_directory, file)
				if os.path.isfile(file_path_full):
					os.remove(file_path_full)
		else:
			os.makedirs(output_directory)
			
		error_files = []
		xlsx_files = []
		for file in os.listdir(file_directory):
			if file.endswith('.xlsx') and file.__contains__('Coal Data'):
				xlsx_files.append(file)
		
		for xlsx_file in xlsx_files:
			file_path = os.path.join(file_directory, xlsx_file)
			output_file = os.path.join(output_directory, os.path.splitext(os.path.basename(xlsx_file))[0] + '.xlsx')
			print(xlsx_file)
			# df = openpyxl.load_workbook(file_path)
			df = pd.read_excel(file_path, engine='openpyxl')
			region_row_index = None
			region_column_index = None
			
			# Find the row and column containing 'Region' values
			for index, row in df.iterrows():
				for column in df.columns:
					if isinstance(row[column], str) and (
							row[column].startswith('Region') or row[column].startswith('Sl No.') or row[
						column].__contains__('State')):
						region_row_index = index
						region_column_index = df.columns.get_loc(column)
						break
				if region_row_index is not None:
					break
			# Delete rows and columns before 'Region'
			if region_row_index is not None and region_column_index is not None:
				df = df.iloc[region_row_index:]
				df = df.iloc[:, region_column_index:]
				print("Region data indexed successfully.")
			else:
				print("No 'Region' or 'Sl No.' values found.")
			
			print(df.iloc[1, 0])
			
			if df.iloc[1, 0] == 1:
				table_heading = df.iloc[0]
				df = df.iloc[1:]
				df.columns = table_heading
			else:
				combined_row = []
				for i in range(len(df.columns)):
					value1 = str(df.iloc[0, i])  # Convert to string
					value2 = str(df.iloc[1, i])  # Convert to string
					if value1 == 'nan' and value2 == 'nan':
						combined_row.append(f'{df.iloc[0, i - 1]} {df.iloc[1, i - 1]}')
					elif value1 == 'nan' and value2 != 'nan':
						combined_row.append(f'{df.iloc[0, i - 1]} {value2}')
					elif value2 == 'nan':
						combined_row.append(f'{value1}')
					else:
						combined_row.append(f'{value1} {value2}')
				
				df.iloc[0] = combined_row
				table_heading = df.iloc[0]
				df = df.iloc[1:]
				df.columns = table_heading
			
			df.to_excel(output_file, index=False)
			df = pd.read_excel(output_file)
			
			# Iterate over columns and rename them based on conditions
			for col in df.columns:
				if col.startswith('Receipt of the day'):
					df = df.rename(columns={col: "Receipt of the day of report ('000 Tonnes)"})
				elif col.startswith('Consumption of the day'):
					df = df.rename(columns={col: "Consumption of the day of report ('000 Tonnes)"})
				elif col.startswith('Critical'):
					df = df.rename(columns={col: "Critical (*) (if stock<25%)"})
				elif col.startswith('Normative Stock Required'):
					df = df.rename(columns={col: "Normative Stock Required (In '000 Tonnes')"})
				elif col.startswith('nan nan') or col.startswith('nan Total'):
					df = df.rename(columns={col: "Actual Stock In ('000 Tonnes') Total"})
				elif col.startswith("Actual Stock In (TT) Import"):
					df = df.rename(columns={col: "Actual Stock In ('000 Tonnes') Import"})
				elif col.startswith("Actual Stock In (TT) Indigenous"):
					df = df.rename(columns={col: "Actual Stock In ('000 Tonnes') Indigenous"})
				elif col.startswith('% of Actual Stock'):
					df = df.rename(columns={col: "Actual Stock as % of Normative Stock"})
				elif col.startswith('Actual Stock as % of Normative  Stock'):
					df = df.rename(columns={col: "Actual Stock as % of Normative Stock"})
				elif col.endswith('Sl No. / State'):
					df = df.rename(columns={col: "Region/State"})
				elif col.endswith('Region/ State'):
					df = df.rename(columns={col: "Region/State"})
				elif col.startswith('Name of Thermal Power Station'):
					df = df.rename(columns={col: "Name of Thermal Power Station/ Performance of Utility"})
				elif col.startswith("Actual Stock In '000 Tonnes'"):
					df = df.rename(columns={col: "Actual Stock In ('000 Tonnes') Indigenous"})
				elif col.startswith("Actual Stock as % of Normative  Stock"):
					df = df.rename(columns={col: "Actual Stock as % of Normative Stock"})
				elif col.startswith("Daily Requirement @85% PLF (TT)"):
					df = df.rename(columns={col: "Requirement for the day @85% PLF (In '000 Tonnes')"})
				elif col.startswith("Requirement for the day (In '000 Tonnes')"):
					df = df.rename(columns={col: "Requirement for the day @85% PLF (In '000 Tonnes')"})
				elif col.startswith("Reasons for critical coal stock/Remarks nan"):
					df = df.rename(columns={col: "Reasons for critical coal stock/Remarks"})
				elif col.startswith('Capacity   (MW)'):
					df = df.rename(columns={col: "Capacity (MW)"})
				elif col.startswith('Percentage (%) of Actual Stock Vis-à-vis Normative  Stock'):
					df = df.rename(columns={col: "Actual Stock as % of Normative Stock"})
				elif col.startswith("Normative stock Required(In '000 Tonnes')"):
					df = df.rename(columns={col: "Normative Stock Required (In '000 Tonnes')"})
				elif col.startswith("Normative Stock Required  (In '000 Tonnes')"):
					df = df.rename(columns={col: "Normative Stock Required (In '000 Tonnes')"})
				elif col.startswith("nan In Days"):
					df = df.rename(columns={col: "Actual Stock In Days"})
				elif col.startswith(" Actual Stock as % of Normative  Stock"):
					df = df.rename(columns={col: "Actual Stock as % of Normative Stock"})
				elif col.startswith("Indigenous/ Import  Stock (In '000 Tonnes')"):
					df = df.rename(columns={col: "Actual Stock (Indigenous/ Import) (In '000 Tonnes)"})
			
			df.to_excel(output_file, index=False)
			df = pd.read_excel(output_file)
			
			region_row_index = None
			region_column_index = None
			
			for index, row in df.iterrows():
				for column in df.columns:
					if isinstance(row[column], str) and (
							row[column].__contains__('All India total') or row[column].__contains__('All India') or row[
						column].__contains__('Region-wise') or
							row[column].__contains__('Grand') or row[column].startswith('As per') or
							row[column].startswith('Pithead') or row[column].startswith('NOTE')):
						region_row_index = index
						region_column_index = df.columns.get_loc(column)
						break
				if region_row_index is not None:
					break
			
			if region_row_index is not None and region_column_index is not None:
				df = df.iloc[:region_row_index]
				print("Rows and columns deleted successfully.")
			else:
				print("No 'NOTE:' or 'Pithead' values found.")
			
			# Convert values in column 0 to integers, and handle non-convertible values
			converted_values = pd.to_numeric(df.iloc[:, 0], errors='coerce')
			
			# Iterate over the DataFrame and replace integer values with the value from the cell above
			for index, value in enumerate(df.iloc[:, 0]):
				if isinstance(value, int):
					df.iloc[index, 0] = df.iloc[index - 1, 0] if index > 0 else value
			
			# Drop rows where the next column of the table is empty
			df = df.dropna(subset=df.columns[1:], how='all')
			
			file_date = xlsx_file.split('Data-')[1].split('.xlsx')[0]
			df.insert(0, 'Date', file_date)
			
			plf_count = 1  # Initialize a counter for 'PLF' columns
			# Iterate over the columns
			for column in df.columns:
				# Check if column heading contains 'PLF'
				if column.startswith('PLF'):
					# Concatenate all row values with column heading
					concatenated_value = df[column].astype(str) + ': ' + column
					# Update the values in the same column
					df[column] = concatenated_value
					# Rename the column heading containing 'PLF' to 'PLF %'
					new_column_name = f'PLF % {plf_count}'  # Generate new column name
					df.rename(columns={column: new_column_name}, inplace=True)
					plf_count += 1  # Increment the 'PLF' column counter
			
			for col in df.columns:
				if col.startswith('Tentative PLF current month upto date'):
					df = df.rename(columns={col: "PLF%/Tentative PLF current month upto date"})
				elif col.startswith('PLF % 1'):
					df = df.rename(columns={col: "PLF%/Tentative PLF current month upto date"})
				elif col.startswith("Actual Stock In ('000 Tonnes') Indigenous.1"):
					df = df.rename(columns={col: "Actual Stock In ('000 Tonnes') Import"})
				elif col.startswith("Reasons for critical coal stock/Remarks.1"):
					df = df.rename(columns={col: "Reasons for critical coal stock/Remarks"})
			
			# Columns to round
			columns_to_round = ['Receipt of the day of report (\'000 Tonnes)',
			                    'Consumption of the day of report (\'000 Tonnes)']
			
			# Iterate over the columns and round the values if the column exists
			for col in columns_to_round:
				if col in df.columns:
					df[col] = pd.to_numeric(df[col], errors='coerce')
					df[col] = df[col].round()
			
			# Columns to round to one decimal place
			columns_to_round_by_1 = ['Requirement for the day @85% PLF (In \'000 Tonnes\')',
			                         'Normative Stock Required (In \'000 Tonnes\')',
			                         'Actual Stock In (\'000 Tonnes\') Indigenous',
			                         'Actual Stock In (\'000 Tonnes\') Import',
			                         'Actual Stock In (\'000 Tonnes\') Total']
			
			# Iterate over the columns and round the values if the column exists
			for col in columns_to_round_by_1:
				if col in df.columns:
					df[col] = pd.to_numeric(df[col], errors='coerce')
					df[col] = df[col].round(decimals=2)
			
			if 'Region/State' in df.columns:
				df['Region/State'] = df['Region/State'].astype(str)
				df = df[~df['Region/State'].str.contains('1', na=False)]
				
				if 'Region/State' in df.columns:
					df['Region/State'] = df['Region/State'].apply(lambda x: x.split('/')[1] if '/' in x else x)
				
				df['Region/State'] = df['Region/State'].replace('nan', '')
				df['Region/State'] = df['Region/State'].replace('', np.nan)
				df['Region/State'] = df['Region/State'].fillna(method='ffill')
				
				columns_to_drop = ["Actual Stock In ('000 Tonnes') Total.3", "Critical (*) (if stock<25%).1",
				                   "Reasons for critical coal stock/Remarks.1"]
				df.drop(columns=[col for col in columns_to_drop if col in df.columns], inplace=True)
				
				df = df[df['Date'] != '']
				df.dropna(axis=1, how='all', inplace=True)
				df_cleaned = df.dropna(subset=['Region/State', 'Mode of Transport'])
				
				df_cleaned.to_excel(output_file, index=False)
				print(f'File Saved at {output_file}')
	
	def merge_xlsx_file(self):
		xlsx_files = []
		for file in os.listdir(output_directory):
			if file.endswith('.xlsx'):
				xlsx_files.append(os.path.join(output_directory, file))
		
		merged_df = pd.DataFrame()
		for file in xlsx_files:
			df = pd.read_excel(file)
			# df.drop(df.index[:0], inplace=True)
			merged_df = pd.concat([merged_df, df], ignore_index=True)
		
		merged_df['Date'] = pd.to_datetime(merged_df['Date'], format='%Y-%m-%d').dt.date
		merged_df['PLF%/Tentative PLF current month upto date'] = merged_df['PLF%/Tentative PLF current month upto date'].astype(str)
		merged_df.dropna(axis=1, how='all', inplace=True)
		
		merged_file_path = os.path.join(self.final_directory, 'coal_npp.xlsx')
		merged_df.to_excel(merged_file_path, index=False)
		print(f"Merged file saved to '{merged_file_path}'")
	
	def edit_merged_file(self):
		merged_file_path = os.path.join(self.final_directory, 'coal_npp.xlsx')
		df = pd.read_excel(merged_file_path)
		df['Date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d').dt.date
		
		df['PLF%/Tentative PLF current month upto date'] = df['PLF%/Tentative PLF current month upto date'].astype(str)
		
		output_path = os.path.join(self.final_directory, 'coal_npp.xlsx')
		df.to_excel(output_path, index=False)
		print(f"File Edited and Saved to '{merged_file_path}'")
	
	def get_data(self):
		coal_npp.coal_data_download(self)
		coal_npp.edit_xlsx_file_using_pandas(self)
		coal_npp.merge_xlsx_file(self)
		pass


if __name__ == '__main__':
	coal_data = coal_npp()
	coal_data.get_data()
	pass
