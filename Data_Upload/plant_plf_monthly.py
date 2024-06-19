import numpy as np
import pandas as pd
import requests
import os
from datetime import datetime, timedelta
import xlwings as xw

file_directory = r'C:\GNA\Coding\Plant PLF\PLF Monthly Report'
output_path = r'C:\GNA\Coding\Plant PLF\Monthly Report Edited'
if not os.path.exists(output_path):
	os.makedirs(output_path)
final_directory = r'C:\GNA\Data Upload'


class plant_plf:
	def __init__(self):
		pass
	
	def convert_xls_to_xlsx(self, file_path):
		xls_files = []
		for file in os.listdir(file_path):
			if file.endswith('.xls'):
				xls_files.append(os.path.join(file_path, file))
		
		for file in xls_files:
			output_path = os.path.join(file_path, os.path.splitext(os.path.basename(file))[0] + '.xlsx')
			
			try:
				app = xw.App(visible=False)  # Open Excel in the background
				workbook = app.books.open(file)
				workbook.save(output_path)
				workbook.close()
				app.quit()
				print(f"Conversion completed. File saved at: {output_path}")
			except Exception as e:
				print(f"Error: {e}")
	
	def plant_plf_report_monthly(self):
		file_directory = r'C:\GNA\Coding\Plant PLF\PLF Monthly Report'
		if not os.path.exists(file_directory):
			os.makedirs(file_directory)
		
		base_date = datetime.now()
		end_date = datetime(2024, 1, 1)
		
		while base_date > end_date:
			start_year = base_date.strftime('%Y')
			start_month = base_date.strftime('%b').upper()
			start_date_2 = base_date.strftime('%Y-%b').upper()
			
			url = f'https://npp.gov.in/public-reports/cea/monthly/generation/18%20col%20act/{start_year}/{start_month}//18%20col%20act-16_{start_date_2}.xls'
			response = requests.get(url)
			
			if response.status_code == 200:
				with open(f'{file_directory}/{start_date_2}.xls', 'wb') as f:
					f.write(response.content)
				
				print(f"Downloaded data for date: {start_date_2}")
			else:
				print(f"No data available for date: {start_date_2}")
			
			# Subtract one month
			base_date = base_date.replace(day=1) - timedelta(days=1)
			# Handling cases where the new month has fewer days
			if base_date.day > 28:
				base_date = base_date.replace(day=28)
		
		if base_date == end_date:
			print("Reached end date. Exiting loop.")
	
	def edit_xlsx_files(self):
		xlsx_files = []
		for file in os.listdir(file_directory):
			if file.endswith('.xlsx'):
				xlsx_files.append(os.path.join(file_directory, file))
		
		for xlsx_file in xlsx_files:
			output_file = os.path.join(file_directory, os.path.splitext(os.path.basename(xlsx_file))[0] + '.xlsx')
			# final_output_file = os.path.join(output_path, os.path.splitext(os.path.basename(xlsx_file))[0] + '.xlsx')
			df = pd.read_excel(xlsx_file)
			
			region_row_index = None
			# Find the row and column containing 'Region' values
			for index, row in df.iterrows():
				for column in df.columns:
					if isinstance(row[column], str) and (
							row[column].startswith('CATEGORY') or row[column].__contains__('CATEGORY/')):
						region_row_index = index
						break
				if region_row_index is not None:
					break
			
			# Drop all rows above this value
			if region_row_index is not None:
				df = df.iloc[region_row_index:]
				print("CATEGORY data indexed successfully.")
			else:
				print("No 'CATEGORY' values found.")
			
			# Drop rows starting with 'Total'
			df = df[~df.apply(lambda row: row.astype(str).str.startswith('TOTAL')).any(axis=1)]
			
			# Add new column 'Region' and delete value from column '0'
			mask = (df.iloc[:, 0].str.contains('REGION')) & (df.iloc[:, 1:].isnull().all(axis=1))
			df['Region Type'] = np.where(mask, df.iloc[:, 0], '')
			df.loc[mask, df.columns[0]] = ''
			
			df = df.dropna(axis=1, how='all')
			df = df.dropna(axis=0, how='all')
			
			df.to_excel(output_file, index=False)
			df = pd.read_excel(xlsx_file)
			
			# Add new column 'Sector Type' and delete value from column '0'
			mask = (df.iloc[:, 0].str.contains('SECTOR')) & (df.iloc[:, 1:].isnull().all(axis=1))
			df['Sector Type'] = np.where(mask, df.iloc[:, 0], '')
			df.loc[mask, df.columns[0]] = ''
			
			# Create new columns 'Station' and 'Station Type' and move values from column '0' and '1'
			mask = df.iloc[:, 0].notna() & df.iloc[:, 1].notna()
			df['Station'] = np.where(mask, df.iloc[:, 0], '')
			df['Station Type'] = np.where(mask, df.iloc[:, 1], '')
			df.loc[mask, [df.columns[0], df.columns[1]]] = ''
			
			df.to_excel(output_file, index=False)
			df = pd.read_excel(xlsx_file)
			
			# Check if all columns except the first one are empty
			mask = df.iloc[:, 1:].isnull().all(axis=1)
			# Extract values from the first column where all other columns are empty
			state_values = df.loc[mask, df.columns[0]]
			# Store these values into a new column named 'State'
			df['State'] = state_values
			# Replace extracted values with NaN in column 0
			df.iloc[mask.values, 0] = np.nan
			
			df[['Region Type', 'State', 'Sector Type']] = df[['Region Type', 'State', 'Sector Type']].ffill()
			df[['Station', 'Station Type']] = df[['Station', 'Station Type']].bfill()
			
			df.iloc[1] = df.iloc[1].ffill()
			df.insert(0, 'Region Type', df.pop('Region Type'))
			df.insert(1, 'State', df.pop('State'))
			df.insert(2, 'Sector Type', df.pop('Sector Type'))
			df.insert(3, 'Station', df.pop('Station'))
			df.insert(4, 'Station Type', df.pop('Station Type'))
			
			# Initialize an empty list to hold the new column headers
			new_headers = []
			# Iterate over each column
			for col in df.columns:
				combined_header = ""
				# Iterate over the first 5 rows to combine their values
				for i in range(5):
					value = df.iloc[i][col]
					if pd.notna(value):
						combined_header += str(value) + " "
				# Strip the trailing space and append to new_headers list
				new_headers.append(combined_header.strip())
			# Assign the combined headers to the DataFrame
			df.columns = new_headers
			# Drop the first 5 rows as they are now used for headers
			df = df.iloc[5:].reset_index(drop=True)
			
			# Drop rows where column 0 has NaN values
			df = df.dropna(subset=[df.columns[5]])
			
			df = df.dropna(axis=1, how='all')
			df = df.dropna(axis=0, how='all')
			
			df.columns = df.columns.str.replace('\n', ' ')
			
			# Renaming Columns
			df.columns.values[0] = 'region'
			df.columns.values[1] = 'state'
			df.columns.values[2] = 'sector_type'
			df.columns.values[3] = 'station'
			df.columns.values[4] = 'station_type'
			df.columns.values[5] = 'plant'
			df.columns.values[6] = 'monitored_capacity_mw'
			df.columns.values[7] = 'target_for_year'
			df.columns.values[8] = 'generation_gwh_program'
			df.columns.values[9] = 'generation_gwh_actual'
			df.columns.values[10] = 'generation_gwh_actual_same_month'
			df.columns.values[11] = 'generation_gwh_percentage_of_program'
			df.columns.values[12] = 'generation_gwh_percentage_of_last_year'
			df.columns.values[18] = 'plant_load_factor_percentage_program'
			df.columns.values[19] = 'plant_load_factor_percentage_actual'
			df.columns.values[20] = 'plant_load_factor_percentage_actual_same_month'
			
			df.drop(df.columns[13:18], axis=1, inplace=True)
			
			file_name = os.path.splitext(os.path.basename(xlsx_file))[0]
			df.insert(0, 'month', file_name)
			
			df.to_excel(output_file, index=False)
			print(f'File Edited {output_file}')
	
	def merge_xlsx_files_plant_plf(self):
		xlsx_files = []
		for file in os.listdir(file_directory):
			if file.endswith('.xlsx'):
				xlsx_files.append(os.path.join(file_directory, file))
		
		merged_df = pd.DataFrame()
		for file in xlsx_files:
			df = pd.read_excel(file)
			merged_df = pd.concat([merged_df, df], ignore_index=True)
		
		merged_df = merged_df.round(decimals=2)
		merged_df['month'] = pd.to_datetime(merged_df['month'], format='%Y-%b').dt.strftime('%Y-%b')
		
		merged_file_path = os.path.join(final_directory, 'plant_plf.xlsx')
		merged_df.to_excel(merged_file_path, index=False)
		print(f"Merged file saved to '{merged_file_path}'")
	
	def get_data(self):
		plant_plf.plant_plf_report_monthly(self)
		plant_plf.convert_xls_to_xlsx(self, file_directory)
		plant_plf.edit_xlsx_files(self)
		plant_plf.merge_xlsx_files_plant_plf(self)
		pass


if __name__ == '__main__':
	plant_plf_npp = plant_plf()
	plant_plf_npp.get_data()
	pass
