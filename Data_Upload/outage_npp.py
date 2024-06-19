import pandas as pd
import requests
import os
from datetime import datetime, timedelta
import xlwings as xw

main_directory = r'C:\GNA\Coding\Outage'

final_directory = r'C:\GNA\Data Upload'


def daily_outage_report_thermal_nuclear_units_only_for_500mw():
	file_path = r'C:\GNA\Coding\Outage\Thermal,Nuclear for 500MW'
	if os.path.exists(file_path):
		for file in os.listdir(file_path):
			file_path_full = os.path.join(file_path, file)
			if os.path.isfile(file_path_full):
				os.remove(file_path_full)
	else:
		os.makedirs(file_path)
	output_path = r'C:\GNA\Coding\Outage\Thermal,Nuclear for 500MW Edited'
	if os.path.exists(output_path):
		for file in os.listdir(output_path):
			file_path_full = os.path.join(output_path, file)
			if os.path.isfile(file_path_full):
				os.remove(file_path_full)
	else:
		os.makedirs(output_path)
	
	base_date = datetime.now()
	end_date = datetime.now() - timedelta(10)
	# end_date = datetime(2019, 1, 1)
	
	while base_date > end_date:
		start_date = base_date.strftime('%d-%m-%Y')
		start_date_2 = base_date.strftime('%Y-%m-%d')
		
		url = f'https://npp.gov.in/public-reports/cea/daily/dgr/{start_date}/dgr11-{start_date_2}.xls'
		response = requests.get(url)
		
		if response.status_code == 200:
			with open(f'{file_path}/{start_date}.xls', 'wb') as f:
				f.write(response.content)
			
			print(f"Downloaded data for date: {start_date_2}")
		else:
			print(f"No data available for date: {start_date_2}")
		
		base_date = base_date - timedelta(days=1)
	
	if base_date == end_date:
		print("Reached end date. Exiting loop.")
	
	xls_to_xlsx(file_path)
	edit_xlsx_files(file_path, output_path)
	merge_xlsx_files(output_path)


def daily_outage_report_coal_lignite_nuclear():
	file_path = r'C:\GNA\Coding\Outage\Coal,Lignite and Nuclear'
	if os.path.exists(file_path):
		for file in os.listdir(file_path):
			file_path_full = os.path.join(file_path, file)
			if os.path.isfile(file_path_full):
				os.remove(file_path_full)
	else:
		os.makedirs(file_path)
	output_path = r'C:\GNA\Coding\Outage\Coal,Lignite and Nuclear Edited'
	if os.path.exists(output_path):
		for file in os.listdir(output_path):
			file_path_full = os.path.join(output_path, file)
			if os.path.isfile(file_path_full):
				os.remove(file_path_full)
	else:
		os.makedirs(output_path)
	
	base_date = datetime.now()
	end_date = datetime.now() - timedelta(10)
	# end_date = datetime(2019, 1, 1)
	
	while base_date > end_date:
		start_date = base_date.strftime('%d-%m-%Y')
		start_date_2 = base_date.strftime('%Y-%m-%d')
		
		url = f'https://npp.gov.in/public-reports/cea/daily/dgr/{start_date}/dgr10-{start_date_2}.xls'
		response = requests.get(url)
		
		if response.status_code == 200:
			with open(f'{file_path}/{start_date}.xls', 'wb') as f:
				f.write(response.content)
			
			print(f"Downloaded data for date: {start_date_2}")
		else:
			print(f"No data available for date: {start_date_2}")
		
		base_date = base_date - timedelta(days=1)
	
	if base_date == end_date:
		print("Reached end date. Exiting loop.")
	
	xls_to_xlsx(file_path)
	edit_xlsx_files(file_path, output_path)
	merge_xlsx_files(output_path)


def daily_outage_report_hydro_units():
	file_path = r'C:\GNA\Coding\Outage\Hydro Units'
	if os.path.exists(file_path):
		for file in os.listdir(file_path):
			file_path_full = os.path.join(file_path, file)
			if os.path.isfile(file_path_full):
				os.remove(file_path_full)
	else:
		os.makedirs(file_path)
	output_path = r'C:\GNA\Coding\Outage\Hydro Units Edited'
	if os.path.exists(output_path):
		for file in os.listdir(output_path):
			file_path_full = os.path.join(output_path, file)
			if os.path.isfile(file_path_full):
				os.remove(file_path_full)
	else:
		os.makedirs(output_path)
	
	base_date = datetime.now()
	end_date = datetime.now() - timedelta(10)
	# end_date = datetime(2019, 1, 1)
	
	while base_date > end_date:
		start_date = base_date.strftime('%d-%m-%Y')
		start_date_2 = base_date.strftime('%Y-%m-%d')
		
		url = f'https://npp.gov.in/public-reports/cea/daily/dgr/{start_date}/dgr7-{start_date_2}.xls'
		response = requests.get(url)
		
		if response.status_code == 200:
			with open(f'{file_path}/{start_date}.xls', 'wb') as f:
				f.write(response.content)
			
			print(f"Downloaded data for date: {start_date_2}")
		else:
			print(f"No data available for date: {start_date_2}")
		
		base_date = base_date - timedelta(days=1)
	
	if base_date == end_date:
		print("Reached end date. Exiting loop.")
	
	xls_to_xlsx(file_path)
	edit_xlsx_files(file_path, output_path)
	merge_xlsx_files(output_path)


def xls_to_xlsx(file_path):
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


def edit_xlsx_files(file_path, output_directory):
	xlsx_files = []
	for file in os.listdir(file_path):
		if file.endswith('.xlsx'):
			xlsx_files.append(os.path.join(file_path, file))
	
	for xlsx_file in xlsx_files:
		output_file = os.path.join(output_directory, os.path.splitext(os.path.basename(xlsx_file))[0] + '.xlsx')
		
		df = pd.read_excel(xlsx_file)
		region_row_index = None
		
		# Find the row and column containing 'Region' values
		for index, row in df.iterrows():
			for column in df.columns:
				if isinstance(row[column], str) and (
						row[column].startswith('State') or row[column].__contains__('State/S')):
					region_row_index = index
					break
			if region_row_index is not None:
				break
		
		# Drop all rows above this value
		if region_row_index is not None:
			df = df.iloc[region_row_index:]
			print("State data indexed successfully.")
		else:
			print("No 'State' values found.")
		
		table_heading = df.iloc[0]
		df = df.iloc[1:]
		df.columns = table_heading
		
		# Drop rows where the next column of the table is empty
		df = df.dropna(subset=df.columns[1:], how='all')
		
		rows_to_drop = []  # Initialize a list to store indices of rows to drop
		for index, row in df.iterrows():
			for column in df.columns:
				if 'Total' in str(row[column]):  # Check if 'Total' is present in any cell of the row
					rows_to_drop.append(index)  # If 'Total' is found, add the index of the row to the list
		# Drop the rows containing 'Total'
		df.drop(index=rows_to_drop, inplace=True)
		
		# Assuming 'df' is your DataFrame
		df = df[~(df.iloc[:, 0].astype(str).str.contains('1'))]
		
		# Drop rows where the first column is empty
		df = df.dropna(subset=[df.columns[0]])
		
		# Split columns starting with 'Date' and 'Expected'
		for col in df.columns:
			if col.startswith('Date'):
				df['Time of Maintenance'] = df[col].str.split(' ').str[1]
				df[col] = df[col].str.split(' ').str[0]
		
		# Split columns starting with 'Date' and 'Expected'
		for col in df.columns:
			if col.startswith('Expected'):
				df['Expected Time of Return'] = df[col].str.split(' ').str[1]
				df[col] = df[col].str.split(' ').str[0]
		
		# Extract value between 'Data-' and '.xlsx' from the filename
		file_date = os.path.splitext(os.path.basename(xlsx_file))[0]
		df.insert(0, 'Date', file_date)
		
		plant_type = os.path.splitext(os.path.basename(file_path))[0]
		df.insert(1, 'Plant Type', plant_type)
		
		df['Date& Time of Maintenance'] = df['Date& Time of Maintenance'].astype(str).str.replace('/', '-')
		df['Expected/sync Date of Return'] = df['Expected/sync Date of Return'].astype(str).str.replace('/', '-')
		
		df['Date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y')
		df['Date& Time of Maintenance'] = pd.to_datetime(df['Date& Time of Maintenance'], format='%d-%m-%Y')
		df['Expected/sync Date of Return'] = pd.to_datetime(df['Expected/sync Date of Return'], format='%d-%m-%Y')
		
		df['Date'] = df['Date'].dt.date
		df['Date& Time of Maintenance'] = df['Date& Time of Maintenance'].dt.date
		df['Expected/sync Date of Return'] = df['Expected/sync Date of Return'].dt.date
		
		df.to_excel(output_file, index=False)
		print(f'File Edited Successfully {xlsx_file}')


def merge_xlsx_files(output_path):
	xlsx_files = []
	for file in os.listdir(output_path):
		if file.endswith('.xlsx'):
			xlsx_files.append(os.path.join(output_path, file))
	
	merged_df = pd.DataFrame()
	for file in xlsx_files:
		df = pd.read_excel(file)
		merged_df = pd.concat([merged_df, df], ignore_index=True)
	
	merged_file_path = os.path.join(main_directory, os.path.splitext(os.path.basename(output_path))[0] + '.xlsx')
	merged_df.to_excel(merged_file_path, index=False)
	print(f"Merged file saved to '{merged_file_path}'")


def merge_final_files():
	xlsx_files = []
	for file in os.listdir(main_directory):
		if file.endswith('.xlsx'):
			xlsx_files.append(os.path.join(main_directory, file))
	merged_df = pd.DataFrame()
	
	for file in xlsx_files:
		df = pd.read_excel(file)
		df.reset_index(drop=True, inplace=True)
		merged_df = pd.concat([merged_df, df], ignore_index=True)
	
	merged_df['State/System'] = merged_df['State/System'].astype(str).replace('Andhra Pradesh.','Andhra Pradesh')
	
	final_file = os.path.join(final_directory, 'outage_npp.xlsx')
	merged_df.to_excel(final_file, index=False)
	print(f"Final File Saved at '{final_file}'")


def get_data():
	daily_outage_report_thermal_nuclear_units_only_for_500mw()
	daily_outage_report_coal_lignite_nuclear()
	daily_outage_report_hydro_units()
	merge_final_files()
	pass


if __name__ == '__main__':
	get_data()
	pass
