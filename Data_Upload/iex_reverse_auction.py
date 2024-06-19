import re
import numpy as np
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import time
import tabula
import pandas as pd
from datetime import datetime, timedelta
from PyPDF2 import PdfReader

file_directory = r'C:\GNA\Coding\Reverse Auction\IEX Reverse Auction'
if not os.path.exists(file_directory):
	os.makedirs(file_directory)

output_directory_1 = r'C:\GNA\Coding\Reverse Auction\IEX Reverse Auction Table0'
if not os.path.exists(output_directory_1):
	os.makedirs(output_directory_1)

output_directory_2 = r'C:\GNA\Coding\Reverse Auction\IEX Reverse Auction Table1'
if not os.path.exists(output_directory_2):
	os.makedirs(output_directory_2)

output_directory_3 = r'C:\GNA\Coding\Reverse Auction\IEX Reverse Auction Table2'
if not os.path.exists(output_directory_3):
	os.makedirs(output_directory_3)

main_directory = r'C:\GNA\Coding\Reverse Auction'

final_directory = r'C:\GNA\Data Upload'

error_log_file_0 = r'C:\GNA\Coding\Reverse Auction\error_log_iex_table0.xlsx'
error_log_file_1 = r'C:\GNA\Coding\Reverse Auction\error_log_iex_table1.xlsx'
error_log_file_2 = r'C:\GNA\Coding\Reverse Auction\error_log_iex_table2.xlsx'

month_replacements = {
	'January': '01-', 'Jan': '01-',
	'February': '02-', 'Feb': '02-',
	'March': '03-', 'Mar': '03-',
	'April': '04-', 'Apr': '04-', 'Apil': '04-',
	'May': '05-',
	'June': '06-', 'Jun': '06-',
	'July': '07-', 'Jul': '07-',
	'August': '08-', 'Aug': '08-',
	'September': '09-', 'Sept': '09-', 'Sep': '09-',
	'October': '10-', 'Oct': '10-', 'OCT': '10-',
	'November': '11-', 'Nov': '11-',
	'December': '12-', 'Dec': '12-'
}


class iex_reverse_auction():
	def __init__(self):
		pass
	
	def download_pdf(self, url, destination):
		response = requests.get(url)
		with open(destination, 'wb') as f:
			f.write(response.content)
	
	def get_links_from_website(self):
		retry_count = 10  # Number of retry attempts
		wait_time = 1  # Time to wait between retry attempts (in seconds)
		
		for attempt in range(retry_count):
			try:
				options = Options()
				prefs = {'download.default_directory': r'C:\GNA\Coding\Reverse Auction\IEX Reverse Auction'}
				options.add_experimental_option('prefs', prefs)
				chromedriver_path = r'C:\Users\pulki\.cache\selenium\chromedriver\win64\125.0.6422.76\chromedriver.exe'
				driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
				# driver = webdriver.Chrome(options=options)
				driver.get('https://www.iexindia.com/TAM_Anyday.aspx?id=ET0WIElIYsE%3d&mid=Gy9kTd80D98%3d')
				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_InnerContent_grdreport")))
				
				url_link = []
				
				while True:
					table = driver.find_element(By.ID, 'ctl00_InnerContent_grdreport')
					rows = table.find_elements(By.TAG_NAME, 'tr')
					for row in rows[:-2]:
						links = row.find_elements(By.TAG_NAME, 'a')
						for link in links:
							href = link.get_attribute('href')
							url_link.append(href)
							if href.endswith('.pdf'):
								iex_reverse_auction.download_pdf(self, href, f'{file_directory}\\{href.split("/")[-1]}')
					
					last_row = rows[-1]
					span_tag = last_row.find_element(By.TAG_NAME, 'span')
					if span_tag.text == '3':
						break
					pagination_links = driver.find_elements(By.XPATH, '//a[contains(@href, "__doPostBack")]')
					next_pagination_link = None
					for link in pagination_links:
						if link.text.isdigit() and int(link.text) > int(span_tag.text):
							next_pagination_link = link
							break
					
					if next_pagination_link:
						next_pagination_link.click()
					else:
						ellipsis_links = driver.find_elements(By.XPATH, '//a[contains(text(), "...")]')
						if len(ellipsis_links) == 0:
							break
						else:
							ellipsis_links[-1].click()
				
				links_df = pd.DataFrame(url_link, columns=['Link'])
				link_file = os.path.join(main_directory, 'url_links.xlsx')
				links_df.to_excel(link_file, index=False)
				break  # Exit the loop if execution is successful
			
			except Exception as e:
				print(f"Attempt {attempt + 1}/{retry_count} failed:", e)
				if attempt < retry_count - 1:
					print(f"Retrying in {wait_time} seconds...")
					time.sleep(wait_time)
				else:
					print("Maximum retry attempts reached. Exiting...")
				continue
	
	def extract_text_from_pdf(self, file_path):
		with open(file_path, 'rb') as f:
			reader = PdfReader(f)
			auction_initiation_date = []
			auction_result_date = []
			auction_id = []
			for page in reader.pages:
				text = page.extract_text()
				# print(text)
				initiation_on = text.find(' ID ')
				initiation_end = text.find('\nBuyer ')
				done_on = text.find('\nBuyer ')
				done_on_end = text.find('\nAllocation')
				auction_no = text.find(' ID ')
				auction_no_end = text.find('ated ')
				
				if initiation_on != -1 and initiation_end != -1:
					initiation_date = text[initiation_on + len('\nReverse '):initiation_end]
					auction_initiation_date.append(initiation_date.strip())
				elif initiation_on != -1:
					initiation_date = text[initiation_on + len('\nReverse '):]
					auction_initiation_date.append(initiation_date.strip())
				
				if done_on != -1 and done_on_end != -1:
					result_date = text[done_on + len('\nResults '):done_on_end]
					auction_result_date.append(result_date.strip())
				elif done_on != -1:
					result_date = text[done_on + len('\nResults '):]
					auction_result_date.append(result_date.strip())
				
				if auction_no != -1 and auction_no_end != -1:
					auction_no_value = text[auction_no + len(' ID '):auction_no_end].strip()
					auction_id.append(auction_no_value)
				elif auction_no != -1:
					auction_no_value = text[auction_no + len(' ID '):].strip()
					auction_id.append(auction_no_value)
			# print(auction_initiation_date)
			# print(auction_result_date)
			# print(auction_id)
			return auction_initiation_date, auction_result_date, auction_id
	
	def pdf_to_xlsx_table0(self):
		pdf_files = []
		error_files = []
		for file in os.listdir(file_directory):
			if file.endswith(
					'.pdf') and 'fter' in file:
				pdf_files.append(file)
		
		for pdf_file in pdf_files:
			file_path = os.path.join(file_directory, pdf_file)
			output_path = os.path.join(output_directory_1, os.path.splitext(pdf_file)[0] + '.xlsx')
			try:
				auction_initiation_date, auction_result_date, auction_id = iex_reverse_auction.extract_text_from_pdf(
					self, file_path)
				tables = tabula.read_pdf(file_path, pages='all')
				modified_tables = []
				
				# Process the first and last tables
				for i, table in enumerate([tables[0]]):
					df = pd.DataFrame(table)
					df = df.fillna(method='ffill')
					if not df.empty and len(df.columns) > 1:  # Check if DataFrame is not empty and has columns
						# Store column headings in the first row of the table
						df.loc[-1] = df.columns
						df.index = df.index + 1
						df = df.sort_index()
						# Process the first table
						if i == 0 and df.iloc[0, 0].startswith('Buyer'):
							df.columns = range(len(df.columns))
						else:
							df.columns = range(len(df.columns))
							df = df.drop(0, axis=1)
							df.columns = range(len(df.columns))
						modified_tables.append(df)
				
				concatenated_df = pd.concat(modified_tables, ignore_index=True)
				# Replace specified values in the first and second columns
				for col in [0, 1]:
					concatenated_df.iloc[:, col] = concatenated_df.iloc[:, col].replace({
						'Buy - Total Quantity (in MwH)': 'Buy - Total Quantity (in MWH)',
						'Buy - Total Quan ty (in MWH)': 'Buy - Total Quantity (in MWH)',
						'Buy - Minimum Quan ty (in MWH)': 'Buy - Minimum Quantity (in MWH)',
						'Buy - Total Quantity (in MWh)': 'Buy - Total Quantity (in MWH)',
						'Buy - Minimum Quantity (in MWh)': 'Buy - Minimum Quantity (in MWH)',
						'Buy - Minimum Quan ty (in MW)': 'Buy - Minimum Quantity (in MW)',
						'Buy - Total Quan ty (in MW)': 'Buy - Total Quantity (in MW)',
						'Exclusion Date(s)': 'Exclusion Dates',
						'Exclusion Days': 'Exclusion Dates',
						'Exclusion Day': 'Exclusion Dates',
						'Exclusion Day(s)': 'Exclusion Dates',
						'Exclusion days': 'Exclusion Dates',
						'Exclusion day(s)': 'Exclusion Dates',
						'Exclusion Date': 'Exclusion Dates',
						'Energy type': 'Energy Type',
					})
				concatenated_df.dropna(axis=0, how='any', inplace=True)
				print(concatenated_df)
				
				transposed_df = concatenated_df.transpose()
				# Set column headings from values of row 1
				transposed_df.columns = transposed_df.iloc[0]
				# Drop the starting row of the transposed DataFrame
				transposed_df = transposed_df.iloc[1:]
				
				column_values = {col: [] for col in transposed_df.columns.unique()}
				# Iterate through each column of the transposed DataFrame
				for col_name, col_data in transposed_df.items():
					# Initialize variables to store values for the current column
					current_col = None
					current_value = None
					# Iterate through each value in the column
					for value in col_data:
						# If the value is NaN, store the current value in the dictionary and reset variables
						if pd.isna(value):
							if current_col:
								column_values[current_col].append(current_value)
							current_col = None
							current_value = None
						# If the value is not NaN, update the current column and value
						else:
							current_col = col_name
							current_value = value
					# Store the last value in the column
					if current_col:
						column_values[current_col].append(current_value)
				# Create a DataFrame from the dictionary
				new_table = pd.DataFrame(column_values)
				
				# Create a Series with the auction ID repeated for each row in the DataFrame
				auction_series = pd.Series([auction_id[0]] * len(new_table), name='Auction No.')
				auction_initiation = pd.Series([auction_initiation_date[0]] * len(new_table),
				                               name='Auction Initiation Date')
				auction_result = pd.Series([auction_result_date[0]] * len(new_table), name='Auction Result Date')
				
				# Concatenate the auction Series with the DataFrame
				new_table = pd.concat([auction_result, new_table], axis=1)
				new_table = pd.concat([auction_initiation, new_table], axis=1)
				new_table = pd.concat([auction_series, new_table], axis=1)
				new_table.dropna(axis=1, how='all', inplace=True)
				new_table['Auction No.'] = new_table['Auction No.'].astype(str)
				print(new_table['Auction No.'])
				print(new_table['Auction Initiation Date'])
				print(new_table['Auction Result Date'])
				
				new_table.to_excel(output_path, index=False)
				print(f'File Saved at {output_path}')
			except Exception as e:
				print(f"Error processing file '{pdf_file}': {e}")
				error_files.append(pdf_file)
		if error_files:
			try:
				pd.DataFrame({'Error Files': error_files}).to_excel(error_log_file_0, index=False)
				print(f"Error log saved to '{error_log_file_0}'")
			except Exception as e:
				print(f"Error saving error log: {e}")
	
	def pdf_to_xlsx_table1(self):
		pdf_files = []
		error_files = []
		for file in os.listdir(file_directory):
			if file.endswith('.pdf') and 'fter' in file:
				pdf_files.append(file)
		
		for pdf_file in pdf_files:
			file_path = os.path.join(file_directory, pdf_file)
			output_path = os.path.join(output_directory_2, os.path.splitext(pdf_file)[0] + '.xlsx')
			try:
				auction_initiation_date, auction_result_date, auction_id = iex_reverse_auction.extract_text_from_pdf(
					self, file_path)
				tables = tabula.read_pdf(file_path, pages='all')
				modified_tables = []
				
				for i, table in enumerate([tables[1]]):
					df = pd.DataFrame(table)
					if not df.empty and len(df.columns) > 1:
						df.loc[-1] = df.columns
						df.index = df.index + 1
						df = df.sort_index()
						if i == 0 and df.iloc[0, 0].startswith('Total'):
							df.columns = range(len(df.columns))
						else:
							df.columns = range(len(df.columns))
							df = df.drop(0, axis=1)
							df.columns = range(len(df.columns))
						modified_tables.append(df)
				
				concatenated_df = pd.concat(modified_tables, ignore_index=True)
				# Replace specified values in the first and second columns
				for col in [0, 1]:
					concatenated_df.iloc[:, col] = concatenated_df.iloc[:, col].replace({
						'L6 Price Discovered (in Rs. /kWh': 'L6 Price Discovered (in Rs. /kWh)',
						'L1 Price Discovered (in Rs./kWh)': 'L1 Price Discovered (in Rs. /kWh)',
						'L1 Price Discovered (in Rs /kWh)': 'L1 Price Discovered (in Rs. /kWh)',
						'L2 Price Discovered (in Rs /kWh)': 'L2 Price Discovered (in Rs. /kWh)',
						'L3 Price Discovered (in Rs /kWh)': 'L3 Price Discovered (in Rs. /kWh)',
						'Minimum Sell Quan ty (in MW) Offered by the L1 Seller': 'Minimum Sell Quantity (in MW) Offered by the L1 Seller',
						'Minimum Sell Quan ty (in MW) Offered by the L1': 'Minimum Sell Quantity (in MW) Offered by the L1 Seller',
						'Minimum Sell Quan ty (in MWH) Offered by the L1 Seller': 'Minimum Sell Quantity (in MWH) Offered by the L1 Seller',
						'Minimum Sell Quan ty (in MW) Offered by the L2': 'Minimum Sell Quantity (in MW) Offered by the L2 Seller',
						'Minimum Sell Quan ty (in MW) Offered by the L3': 'Minimum Sell Quantity (in MW) Offered by the L3 Seller',
						'Minimum Sell Quan ty (in MW)Offered by the L4 Seller': 'Minimum Sell Quantity (in MW) Offered by the L4 Seller',
						'Minimum Sell Quan ty (in MW) Offered by the L5': 'Minimum Sell Quantity (in MW) Offered by the L5 Seller',
						'Minimum Sell Quantity (in MWH) Offered by the L1': 'Minimum Sell Quantity (in MWH) Offered by the L1 Seller',
						'Total count of Sellers who par cipated in the auc': 'Total count of Sellers who participated in the auction',
						'Total Sell Quan ty (in MW) Offered by the L1 Seller': 'Total Sell Quantity (in MW) Offered by the L1 Seller',
						'Total Sell Quan ty (in MW) Offered by the L2 Seller': 'Total Sell Quantity (in MW) Offered by the L2 Seller',
						'Total Sell Quan ty (in MW) Offered by the L3 Seller': 'Total Sell Quantity (in MW) Offered by the L3 Seller',
						'Total Sell Quan ty (in MW) Offered by the L4 Seller': 'Total Sell Quantity (in MW) Offered by the L4 Seller',
						'Total Sell Quan ty (in MW) Offered by the L5 Seller': 'Total Sell Quantity (in MW) Offered by the L5 Seller',
						'Total Sell Quantity (in MWh) Offered by the L3 Seller': 'Total Sell Quantity (in MWH) Offered by the L3 Seller',
						'Total Sell Quantity (in MWh) Offered by the L4 Seller': 'Total Sell Quantity (in MWH) Offered by the L4 Seller',
						'Total Sell Quan ty (in MWH) Offered by the L1 Seller': 'Total Sell Quantity (in MWH) Offered by the L1 Seller',
						'Total Sell Quan ty (in MWH) Offered by the L2 Seller': 'Total Sell Quantity (in MWH) Offered by the L2 Seller'
					})
				if 3 in concatenated_df.columns:
					concatenated_df[3].fillna('', inplace=True)
					concatenated_df.loc[concatenated_df[3] != '', 1] = concatenated_df[3]
					concatenated_df.drop('2', axis=1, inplace=True)
					concatenated_df.drop('3', axis=1, inplace=True)
				print(concatenated_df)
				
				transposed_df = concatenated_df.transpose()
				# Set column headings from values of row 1
				transposed_df.columns = transposed_df.iloc[0]
				# Drop the starting row of the transposed DataFrame
				transposed_df = transposed_df.iloc[1:]
				
				transposed_df.insert(0, 'Auction No.', auction_id[0])
				transposed_df['Auction No.'] = transposed_df['Auction No.'].str.replace(' ', '')
				transposed_df['Auction No.'] = transposed_df['Auction No.'].str.split('\n').str[0]
				transposed_df.dropna(axis=1, how='all', inplace=True)
				transposed_df['Auction No.'] = transposed_df['Auction No.'].astype(str)
				print(transposed_df['Auction No.'])
				
				transposed_df.to_excel(output_path, index=False)
				print(f'File Saved at {output_path}')
			except Exception as e:
				print(f"Error processing file '{pdf_file}': {e}")
				error_files.append(pdf_file)
		if error_files:
			try:
				pd.DataFrame({'Error Files': error_files}).to_excel(error_log_file_1, index=False)
				print(f"Error log saved to '{error_log_file_1}'")
			except Exception as e:
				print(f"Error saving error log: {e}")
	
	def pdf_to_xlsx_table2(self):
		pdf_files = []
		error_files = []
		for file in os.listdir(file_directory):
			if file.endswith('.pdf') and 'fter' in file:
				pdf_files.append(file)
		
		for pdf_file in pdf_files:
			file_path = os.path.join(file_directory, pdf_file)
			output_path = os.path.join(output_directory_3, os.path.splitext(pdf_file)[0] + '.xlsx')
			try:
				auction_initiation_date, auction_result_date, auction_id = iex_reverse_auction.extract_text_from_pdf(
					self, file_path)
				tables = tabula.read_pdf(file_path, pages='all')
				modified_tables = []
				# Process the first and last tables
				for i, table in enumerate([tables[2]]):
					df = pd.DataFrame(table)
					if not df.empty and len(df.columns) > 1:  # Check if DataFrame is not empty and has columns
						# Store column headings in the first row of the table
						df.loc[-1] = df.columns
						df.index = df.index + 1
						df = df.sort_index()
						# Check if the first value of the first row starts with 'Allocated'
						if df.iloc[0, 0].startswith('Allocated'):
							df.columns = range(len(df.columns))
						else:
							df.columns = range(len(df.columns))
							df = df.drop(0, axis=1)
							df.columns = range(len(df.columns))
						modified_tables.append(df)
				
				concatenated_df = pd.concat(modified_tables, ignore_index=True)
				
				# Replace specified values in the first and second columns
				replacement_dict = {
					'Allocated Quan ty (in MW)': 'Allocated Quantity (in MW)',
					'Allocated Quanty (in MW)': 'Allocated Quantity (in MW)',
					'Accepted Price (in Rs. /kWh)': 'Accepted Price (in Rs./kWh)',
					'Allocated Quantity L2(in MW)': 'Allocated Quantity L2 (in MW)',
					'Allocated Quantity L3(in MW)': 'Allocated Quantity L3 (in MW)',
					'Allocated Quantity (in MWh)': 'Allocated Quantity (in MWH)',
					'Accepted Price ( in Rs./kWh)': 'Accepted Price (in Rs./kWh)',
					'Allocated Quantity L1 (in MW)': 'Allocated Quantity (in MW)',
					'Allocated Quantity L2 (in MW)': 'Allocated Quantity (in MW)',
					'Allocated Quantity L3 (in MW)': 'Allocated Quantity (in MW)',
					'Allocated Quantity L4 (in MW)': 'Allocated Quantity (in MW)',
					'Accepted Price L4 (in Rs./kWh)': 'Accepted Price (in Rs./kWh)',
					'Accepted Price L1 (in Rs./kWh)': 'Accepted Price (in Rs./kWh)',
					'Accepted Price L2 (in Rs./kWh)': 'Accepted Price (in Rs./kWh)',
					'Accepted Price L3 (in Rs./kWh)': 'Accepted Price (in Rs./kWh)',
					'Allocated  Quantity  L2(in MW)': 'Allocated Quantity (in MW)',
					'Allocated  Quantity  L3(in MW)': 'Allocated Quantity (in MW)',
				}
				for col in [0, 1]:
					concatenated_df.iloc[:, col] = concatenated_df.iloc[:, col].replace(replacement_dict)
				
				print(concatenated_df)
				
				transposed_df = concatenated_df.transpose()
				transposed_df.columns = transposed_df.iloc[0]
				
				# Drop the starting row of the transposed DataFrame
				transposed_df = transposed_df.iloc[1:]
				transposed_df.columns = transposed_df.columns.str.replace('L2 ', '').str.replace('L3 ', '').str.replace(
					'L1 ', '').str.replace('L4 ', '')
				
				column_values = {col: [] for col in transposed_df.columns.unique()}
				# Iterate through each column of the transposed DataFrame
				for col_name, col_data in transposed_df.items():
					# Initialize variables to store values for the current column
					current_col = None
					current_value = None
					# Iterate through each value in the column
					for value in col_data:
						# If the value is NaN, store the current value in the dictionary and reset variables
						if pd.isna(value):
							if current_col:
								column_values[current_col].append(current_value)
							current_col = None
							current_value = None
						# If the value is not NaN, update the current column and value
						else:
							current_col = col_name
							current_value = value
					# Store the last value in the column
					if current_col:
						column_values[current_col].append(current_value)
				# Create a DataFrame from the dictionary
				new_table = pd.DataFrame(column_values)
				
				# Create a Series with the auction ID repeated for each row in the DataFrame
				auction_series = pd.Series([auction_id[0]] * len(new_table), name='Auction No.')
				# Concatenate the auction Series with the DataFrame
				new_table = pd.concat([auction_series, new_table], axis=1)
				new_table.dropna(axis=1, how='all', inplace=True)
				new_table['Auction No.'] = new_table['Auction No.'].astype(str)
				# print(new_table['Auction No.'])
				print(new_table)
				new_table.to_excel(output_path, index=False)
				print(f'File Saved at {output_path}')
			except Exception as e:
				print(f"Error processing file '{pdf_file}': {e}")
				error_files.append(pdf_file)
		if error_files:
			try:
				pd.DataFrame({'Error Files': error_files}).to_excel(error_log_file_2, index=False)
				print(f"Error log saved to '{error_log_file_2}'")
			except Exception as e:
				print(f"Error saving error log: {e}")
	
	def merge_table0(self):
		xlsx_files = [file for file in os.listdir(output_directory_1) if file.endswith('.xlsx')]
		merged_df = pd.DataFrame()
		for xlsx_file in xlsx_files:
			file_path = os.path.join(output_directory_1, xlsx_file)
			try:
				df = pd.read_excel(file_path)
				merged_df = pd.concat([merged_df, df], ignore_index=True)
			except Exception as e:
				print(f"Error reading file '{xlsx_file}': {e}")
		
		if not merged_df.empty:
			try:
				merged_table0 = os.path.join(main_directory, 'merged_iex_table0.xlsx')
				merged_df.to_excel(merged_table0, index=False)
				print(f"Merged file saved to '{merged_table0}'")
			except Exception as e:
				print(f"Error saving merged file: {e}")
		else:
			print("No data to Merge")
	
	def merge_table1(self):
		xlsx_files = [file for file in os.listdir(output_directory_2) if file.endswith('.xlsx')]
		merged_df = pd.DataFrame()
		for xlsx_file in xlsx_files:
			file_path = os.path.join(output_directory_2, xlsx_file)
			try:
				df = pd.read_excel(file_path)
				merged_df = pd.concat([merged_df, df], ignore_index=True)
			except Exception as e:
				print(f"Error reading file '{xlsx_file}': {e}")
		
		if not merged_df.empty:
			try:
				merged_table1 = os.path.join(main_directory, 'merged_iex_table1.xlsx')
				merged_df.to_excel(merged_table1, index=False)
				print(f"Merged file saved to '{merged_table1}'")
			except Exception as e:
				print(f"Error saving merged file: {e}")
		else:
			print("No data to Merge")
	
	def merge_table2(self):
		xlsx_files = [file for file in os.listdir(output_directory_3) if file.endswith('.xlsx')]
		merged_df = pd.DataFrame()
		for xlsx_file in xlsx_files:
			file_path = os.path.join(output_directory_3, xlsx_file)
			try:
				df = pd.read_excel(file_path)
				merged_df = pd.concat([merged_df, df], ignore_index=True)
			except Exception as e:
				print(f"Error reading file '{xlsx_file}': {e}")
		
		if not merged_df.empty:
			try:
				merged_table2 = os.path.join(main_directory, 'merged_iex_table2.xlsx')
				merged_df.to_excel(merged_table2, index=False)
				print(f"Merged file saved to '{merged_table2}'")
			except Exception as e:
				print(f"Error saving merged file: {e}")
		else:
			print("No data to Merge")
	
	def edit_table0(self):
		file_path = os.path.join(main_directory, 'merged_iex_table0.xlsx')
		df = pd.read_excel(file_path)
		df['Buyer'] = df['Buyer'].astype(str).replace('_', ' ', regex=True)
		df = df[~df['Buyer'].str.startswith('Unnamed')]
		
		# Edit columns Auction No, Auction Initiation Date , Auction Result Date
		df['Auction No.'] = df['Auction No.'].astype(str).replace(' ', '', regex=True)
		df['Auction No.'] = df['Auction No.'].str.split(r'\n\n').str[0]
		df['Auction Initiation Date'] = df['Auction Initiation Date'].astype(str).replace(' ', '', regex=True)
		df['Auction Initiation Date'] = df['Auction Initiation Date'].str.split(r'edon').str[1]
		df['Auction Result Date'] = df['Auction Result Date'].astype(str).replace(' ', '', regex=True)
		df['Auction Result Date'] = df['Auction Result Date'].astype(str).replace(r'\n', '', regex=True)
		df['Auction Result Date'] = df['Auction Result Date'].str.split(r'neon').str[1]
		df['Auction Result Date'] = df['Auction Result Date'].str.split(r'Total').str[0]
		
		for month_name, month_num in month_replacements.items():
			df['Auction Initiation Date'] = df['Auction Initiation Date'].str.replace(month_name, month_num)
			df['Auction Result Date'] = df['Auction Result Date'].str.replace(month_name, month_num)
			df['Delivery Dates'] = df['Delivery Dates'].str.replace(month_name, month_num)
		
		df['Auction Initiation Date'] = df['Auction Initiation Date'].astype(str).apply(
			lambda x: re.sub(r'(?:st|nd|rd|th|TH|RD|ND|,)', '-', x))
		df['Auction Result Date'] = df['Auction Result Date'].astype(str).apply(
			lambda x: re.sub(r'(?:st|nd|rd|th|TH|RD|ND|,)', '-', x))
		
		df['Auction Initiation Date'] = df['Auction Initiation Date'].astype(str).apply(
			lambda x: re.sub(r'^(\d)-(\d{2}-\d{4})$', r'0\1-\2', x) if isinstance(x, str) else x)
		df['Auction Result Date'] = df['Auction Result Date'].astype(str).apply(
			lambda x: re.sub(r'^(\d)-(\d{2}-\d{4})$', r'0\1-\2', x) if isinstance(x, str) else x)
		
		df['Auction No.'] = df['Auction No.'].astype(str)
		
		# Edit Delivery Dates
		df['Delivery Dates'] = df['Delivery Dates'].astype(str)
		new_rows = []
		for index, row in df.iterrows():
			if '&' in row['Delivery Dates']:
				dates = row['Delivery Dates'].split('&')
				df.at[index, 'Delivery Dates'] = dates[0].strip()
				new_row = {'Delivery Dates': dates[1].strip(), 'Delivery Period': row['Delivery Period']}
				for col in df.columns:
					if col not in ['Delivery Dates', 'Delivery Period']:
						new_row[col] = row[col]
				new_rows.append(new_row)
		df = df._append(new_rows, ignore_index=True)
		df['Delivery Dates'] = df['Delivery Dates'].astype(str).replace(' ', '', regex=True)
		df['Delivery Dates'] = df['Delivery Dates'].astype(str).apply(
			lambda x: re.sub(r'(?:to|TO|,)', ' ', x))
		df['Delivery Dates'] = df['Delivery Dates'].astype(str).apply(
			lambda x: re.sub(r'(?:th|/|,)', '-', x))
		df[['Delivery Start Date', 'Delivery End Date']] = df['Delivery Dates'].astype(str).str.split(' ', expand=True)
		df.drop(['Delivery Dates'], axis=1, inplace=True)
		
		# Edit Delivery Period
		df['Delivery Period'] = df['Delivery Period'].astype(str).replace('&', ',', regex=True)
		df['Delivery Period'] = df['Delivery Period'].astype(str).replace(' ', '', regex=True)
		df['Delivery Period'] = df['Delivery Period'].astype(str).replace('–', 'to', regex=True)
		new_rows = []
		for index, row in df.iterrows():
			if ',' in row['Delivery Period']:
				times = row['Delivery Period'].split(',')
				df.at[index, 'Delivery Period'] = times[0].strip()
				new_row = {'Delivery Period': times[1].strip()}
				for col in df.columns:
					if col not in ['Delivery Period']:
						new_row[col] = row[col]
				new_rows.append(new_row)
		df = df._append(new_rows, ignore_index=True)
		df[['Delivery Start Time', 'Delivery End Time']] = df['Delivery Period'].astype(str).str.split('to',
		                                                                                               expand=True)
		df.drop(['Delivery Period'], axis=1, inplace=True)
		df[['Delivery Start Time', 'Delivery End Time']] = df[['Delivery Start Time', 'Delivery End Time']].astype(
			str).replace('hrs.', '', regex=True)
		df[['Delivery Start Time', 'Delivery End Time']] = df[['Delivery Start Time', 'Delivery End Time']].astype(
			str).replace('hrs', '', regex=True)
		df[['Delivery Start Time', 'Delivery End Time']] = df[['Delivery Start Time', 'Delivery End Time']].astype(
			str).replace('hr', '', regex=True)
		df[['Delivery Start Time', 'Delivery End Time']] = df[['Delivery Start Time', 'Delivery End Time']].astype(
			str).replace('Hrs.', '', regex=True)
		
		# Edit Buy Total and Buy Minimum
		df['Buy - Minimum Quantity (in MWH)'] = df['Buy - Minimum Quantity (in MWH)'].astype(str).str.split(r' \(').str[
			0]
		df['Buy - Minimum Quantity (in MWH)'] = \
			df['Buy - Minimum Quantity (in MWH)'].astype(str).str.split(r'\( Non Solar').str[0]
		df['Buy - Total Quantity (in MWH)'] = df['Buy - Total Quantity (in MWH)'].astype(str).str.split(r'\(').str[0]
		df['Buy - Total Quantity (in MW)'] = df['Buy - Total Quantity (in MW)'].astype(str).str.split(r' \(').str[0]
		df['Buy - Minimum Quantity (in MWH)'] = df['Buy - Minimum Quantity (in MWH)'].astype(float)
		df['Buy - Total Quantity (in MWH)'] = df['Buy - Total Quantity (in MWH)'].astype(float)
		# df['Buy - Minimum Quantity (in MWH)'] /= 24
		# df['Buy - Total Quantity (in MWH)'] /= 24
		#
		# df.loc[df['Buy - Total Quantity (in MW)'].isnull(), 'Buy - Total Quantity (in MW)'] = df[
		# 	'Buy - Total Quantity (in MWH)']
		# df.loc[df['Buy - Minimum Quantity (in MW)'].isnull(), 'Buy - Minimum Quantity (in MW)'] = df[
		# 	'Buy - Minimum Quantity (in MWH)']
		#
		# df.loc[df['Buy - Total Quantity (in MW)'] == 'nan', 'Buy - Total Quantity (in MW)'] = df[
		# 	'Buy - Total Quantity (in MWH)']
		# df.loc[df['Buy - Minimum Quantity (in MWH)'] == 'nan', 'Buy - Minimum Quantity (in MW)'] = df[
		# 	'Buy - Minimum Quantity (in MWH)']
		
		# df.drop(['Buy - Minimum Quantity (in MWH)', 'Buy - Total Quantity (in MWH)'], axis=1, inplace=True)
		# df[['Buy - Total Quantity (in MW)', 'Buy - Minimum Quantity (in MW)']] = df[
		# 	['Buy - Total Quantity (in MW)', 'Buy - Minimum Quantity (in MW)']].astype(float)
		
		df['Buyer'] = df['Buyer'].astype(str).replace('_', ' ', regex=True)
		
		# Set Column Indexes
		df.insert(4, 'Delivery Start Date', df.pop('Delivery Start Date'))
		df.insert(5, 'Delivery End Date', df.pop('Delivery End Date'))
		df.insert(6, 'Delivery Start Time', df.pop('Delivery Start Time'))
		df.insert(7, 'Delivery End Time', df.pop('Delivery End Time'))
		
		df['Delivery Start Date'] = pd.to_datetime(df['Delivery Start Date'], format='%d-%m-%Y').dt.date
		df['Delivery End Date'] = pd.to_datetime(df['Delivery End Date'], format='%d-%m-%Y', errors='coerce').dt.date
		df['Auction Initiation Date'] = pd.to_datetime(df['Auction Initiation Date'], format='%d-%m-%Y',
		                                               errors='coerce').dt.date
		df['Auction Result Date'] = pd.to_datetime(df['Auction Result Date'], format='%d-%m-%Y').dt.date
		
		df = df[~df['Delivery Start Time'].str.startswith('nan')]
		
		df['Delivery Start Time'] = df['Delivery Start Time'].astype(str).replace('AsperSellerprofile',
		                                                                          'As per Seller Profile', regex=True)
		df['Delivery End Time'] = df['Delivery End Time'].str.replace('None', 'As per Seller Profile')
		df['Delivery Start Time'] = df['Delivery Start Time'].astype(str).replace('AsPerSellerProfile',
		                                                                          'As per Seller Profile', regex=True)
		df['Delivery Start Time'] = df['Delivery Start Time'].astype(str).replace('AspersellerProfile',
		                                                                          'As per Seller Profile', regex=True)
		df['Delivery Start Time'] = df['Delivery Start Time'].astype(str).replace('Aspersellerprofile',
		                                                                          'As per Seller Profile', regex=True)
		df['Delivery Start Time'] = df['Delivery Start Time'].astype(str).replace('AsperSellerProfile',
		                                                                          'As per Seller Profile', regex=True)
		
		edited_table0 = os.path.join(main_directory, 'edited_iex_table0.xlsx')
		df.to_excel(edited_table0, index=False)
		print(f"Edited File Saved at {edited_table0}")
	
	def edit_table1(self):
		file_path = os.path.join(main_directory, 'merged_iex_table1.xlsx')
		df = pd.read_excel(file_path)
		# Edit Auction No. column
		df['Auction No.'] = df['Auction No.'].astype(str).replace(' ', '', regex=True)
		df['Auction No.'] = df['Auction No.'].astype(str).str.split('\n\n').str[0]
		
		df.dropna(axis=1, how='all', inplace=True)
		edited_table1 = os.path.join(main_directory, 'edited_iex_table1.xlsx')
		df.to_excel(edited_table1, index=False)
		print(f"Edited File Saved at {edited_table1}")
	
	def edit_table2(self):
		file_path = os.path.join(main_directory, 'merged_iex_table2.xlsx')
		df = pd.read_excel(file_path)
		# Edit Auction No. column
		df['Auction No.'] = df['Auction No.'].astype(str).replace(' ', '', regex=True)
		df['Auction No.'] = df['Auction No.'].astype(str).str.split('\n\n').str[0]
		
		# Edit Column Allocated Quantity (in MW), Accepted Price (in Rs./kWh)
		split_values = df['Allocated Quantity (in MW), Accepted Price (in Rs./kWh)'].astype(str).str.split('@',
		                                                                                                   expand=True)
		
		df['Allocated Quantity (in MW) New'] = split_values[0]
		df['Accepted Price (in Rs./kWh) New'] = split_values[1]
		
		df['Allocated Quantity (in MW)'] = df['Allocated Quantity (in MW)'].fillna(df['Allocated Quantity (in MW) New'])
		df['Accepted Price (in Rs./kWh)'] = df['Accepted Price (in Rs./kWh)'].fillna(
			df['Accepted Price (in Rs./kWh) New'])
		
		df.drop(['Allocated Quantity (in MW) New', 'Accepted Price (in Rs./kWh) New',
		         'Allocated Quantity (in MW), Accepted Price (in Rs./kWh)'], axis=1, inplace=True)
		
		# Edit Column Allocated Quantity (in MW), Accepted Price (in Rs./kWh)
		df[['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']] = df[
			['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']].astype(str).replace('-', 'nan', regex=True)
		df[['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']] = df[
			['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']].astype(str).replace('Mw', '', regex=True)
		df[['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']] = df[
			['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']].astype(str).replace('Na', 'nan', regex=True)
		df[['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']] = df[
			['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']].astype(str).replace('@', '', regex=True)
		df[['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']] = df[
			['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']].astype(str).replace('Unnamed: 1',
		                                                                                       'To be Allocated',
		                                                                                       regex=True)
		df[['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']] = df[
			['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']].astype(str).replace('To be allocate',
		                                                                                       'To be Allocated',
		                                                                                       regex=True)
		df[['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']] = df[
			['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']].astype(str).replace('To be allocated',
		                                                                                       'To be Allocated',
		                                                                                       regex=True)
		df[['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']] = df[
			['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']].astype(str).replace(' ', '', regex=True)
		
		for index, row in df.iterrows():
			allocated_quantity_count = row['Allocated Quantity (in MW)'].count(',') + row[
				'Allocated Quantity (in MW)'].count('&')
			accepted_price_count = row['Accepted Price (in Rs./kWh)'].count(',') + row[
				'Accepted Price (in Rs./kWh)'].count(
				'&')
			
			if allocated_quantity_count == accepted_price_count:
				df.at[index, 'Accepted Price (in Rs./kWh)'] = row['Accepted Price (in Rs./kWh)'].replace('&', ',')
				df.at[index, 'Allocated Quantity (in MW)'] = row['Allocated Quantity (in MW)'].replace('&', ',')
		
		new_rows = []
		for index, row in df.iterrows():
			if '&' in row['Allocated Quantity (in MW)']:
				quantities = row['Allocated Quantity (in MW)'].split('&')
				prices = row['Accepted Price (in Rs./kWh)'].split('&')
				for i in range(max(len(quantities), len(prices))):
					new_row = {}
					if i < len(quantities):
						new_row['Allocated Quantity (in MW)'] = quantities[i].strip()
					else:
						new_row['Allocated Quantity (in MW)'] = row['Allocated Quantity (in MW)']
					if i < len(prices):
						new_row['Accepted Price (in Rs./kWh)'] = prices[i].strip()
					else:
						new_row['Accepted Price (in Rs./kWh)'] = row['Accepted Price (in Rs./kWh)']
					for col in df.columns:
						if col not in ['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']:
							new_row[col] = row[col]
					new_rows.append(new_row)
			else:
				new_rows.append(row.to_dict())
		new_df = pd.DataFrame(new_rows)
		
		new_rows = []
		for index, row in new_df.iterrows():
			if ',' in row['Allocated Quantity (in MW)']:
				quantities = row['Allocated Quantity (in MW)'].split(',')
				prices = row['Accepted Price (in Rs./kWh)'].split(',')
				for i in range(max(len(quantities), len(prices))):
					new_row = {}
					if i < len(quantities):
						new_row['Allocated Quantity (in MW)'] = quantities[i].strip()
					else:
						new_row['Allocated Quantity (in MW)'] = row['Allocated Quantity (in MW)']
					if i < len(prices):
						new_row['Accepted Price (in Rs./kWh)'] = prices[i].strip()
					else:
						new_row['Accepted Price (in Rs./kWh)'] = row['Accepted Price (in Rs./kWh)']
					for col in df.columns:
						if col not in ['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']:
							new_row[col] = row[col]
					new_rows.append(new_row)
			else:
				new_rows.append(row.to_dict())
		new_df = pd.DataFrame(new_rows)
		
		edited_table2 = os.path.join(main_directory, 'edited_iex_table2.xlsx')
		new_df.to_excel(edited_table2, index=False)
		print(f"Edited File Saved at {edited_table2}")
	
	def merge_files(self):
		table0 = os.path.join(main_directory, 'edited_iex_table0.xlsx')
		# table1 = os.path.join(main_directory, 'edited_iex_table1.xlsx')
		table2 = os.path.join(main_directory, 'edited_iex_table2.xlsx')
		df1 = pd.read_excel(table0)
		# df2 = pd.read_excel(table1)
		df3 = pd.read_excel(table2)
		
		# Convert 'Auction No.' to string in all DataFrames
		df1['Auction No.'] = df1['Auction No.'].astype(str)
		# df2['Auction No.'] = df2['Auction No.'].astype(str)
		df3['Auction No.'] = df3['Auction No.'].astype(str)
		
		# merged_df = pd.merge(pd.merge(df1, df3, on='Auction No.', how="right"), df2, on='Auction No.', how='right')
		merged_df = pd.merge(df1, df3, on='Auction No.', how="right")
		
		merged_df['Auction No.'] = merged_df['Auction No.'].astype(str)
		merged_df.insert(0, 'Exchange Type', 'IEX')
		merged_df = merged_df.dropna(subset=['Buyer'])
		merged_df['Delivery Start Date'] = merged_df['Delivery Start Date'].dt.date
		merged_df['Delivery End Date'] = merged_df['Delivery End Date'].dt.date
		merged_df['Auction Initiation Date'] = merged_df['Auction Initiation Date'].dt.date
		merged_df['Auction Result Date'] = merged_df['Auction Result Date'].dt.date
		
		merged_df.rename(columns={
			'Buy - Total Quantity (in MW)': 'Buy Total Quantity (in MW)',
			'Buy - Minimum Quantity (in MW)': 'Buy Minimum Quantity (in MW)'
		}, inplace=True)
		merged_df['Allocated Quantity (in MW)'] = merged_df['Allocated Quantity (in MW)'].astype(str).replace('NaN', '',
		                                                                                                      regex=True)
		merged_df['Allocated Quantity (in MW)'] = merged_df['Allocated Quantity (in MW)'].astype(str).replace('nan', '',
		                                                                                                      regex=True)
		merged_df['Allocated Quantity (in MW)'] = merged_df['Allocated Quantity (in MW)'].astype(str).replace(
			'TobeAllocatedd', 'TobeAllocated',
			regex=True)
		merged_df['Accepted Price (in Rs./kWh)'] = merged_df['Accepted Price (in Rs./kWh)'].astype(str).replace(
			'TobeAllocatedd', 'TobeAllocated',
			regex=True)
		merged_df['Accepted Price (in Rs./kWh)'] = merged_df['Accepted Price (in Rs./kWh)'].astype(str).replace('nan', '',
		                                                                                                      regex=True)
		merged_df = merged_df.sort_values(by='Auction Result Date', ascending=False)
		
		# Rename Columns
		merged_df.rename(columns={
            'Buy - Total Quantity (in MWH)': 'Buy Total Quantity (in MWH)',
            'Buy - Minimum Quantity (in MWH)': 'Buy Minimum Quantity (in MWH)'
        }, inplace=True)
 
		# Columns to Move
		column_to_move = [
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
			'Buy Total Quantity (in MWH)',
			'Buy Minimum Quantity (in MW)',
			'Buy Minimum Quantity (in MWH)',
			'Delivery Point',
			'Total count of Delivery Days',
			'Energy Type',
			'Allocated Quantity (in MW)',
			'Allocated Quantity (in MWH)',
			'Accepted Price (in Rs./kWh)',
		]
		remaining_column_order = [col for col in merged_df.columns if col not in column_to_move]
		column_order = column_to_move + remaining_column_order
		merged_df = merged_df[column_order]
		
		final_file = os.path.join(final_directory, 'final_iex.xlsx')
		merged_df.to_excel(final_file, index=False)
		print(f"Final File Saved at {final_file}")
	
	
	def get_data(self):
		iex_reverse_auction.get_links_from_website(self)
		iex_reverse_auction.pdf_to_xlsx_table0(self)
		# # iex_reverse_auction.pdf_to_xlsx_table1(self)
		iex_reverse_auction.pdf_to_xlsx_table2(self)
		iex_reverse_auction.merge_table0(self)
		# # iex_reverse_auction.merge_table1(self)
		iex_reverse_auction.merge_table2(self)
		iex_reverse_auction.edit_table0(self)
		# # iex_reverse_auction.edit_table1(self)
		iex_reverse_auction.edit_table2(self)
		iex_reverse_auction.merge_files(self)
		pass


if __name__ == '__main__':
	tam_iex = iex_reverse_auction()
	tam_iex.get_data()
	pass

