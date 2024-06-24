import numpy as np
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import tabula
import pandas as pd
from PyPDF2 import PdfReader
import re

main_directory = r'C:\GNA\Data\Reverse Auction'

file_directory = r'C:\GNA\Data\Reverse Auction\HPX Reverse Auction'
if os.path.exists(file_directory):
	for file in os.listdir(file_directory):
		file_path_full = os.path.join(file_directory, file)
		if os.path.isfile(file_path_full):
			os.remove(file_path_full)
else:
	os.makedirs(file_directory)

output_directory = r'C:\GNA\Data\Reverse Auction\HPX Reverse Auction xlsx Files'
if os.path.exists(output_directory):
	for file in os.listdir(output_directory):
		file_path_full = os.path.join(output_directory, file)
		if os.path.isfile(file_path_full):
			os.remove(file_path_full)
else:
	os.makedirs(output_directory)

final_directory = r'C:\GNA\Data Upload'

error_log_file = r'C:\GNA\Data\Reverse Auction\hpx_pdf_not_converted.xlsx'

month_replacements = {
	'January': '01', 'Jan': '01',
	'February': '02', 'Feb': '02',
	'March': '03', 'Mar': '03',
	'April': '04', 'Apr': '04',
	'May': '05',
	'June': '06', 'Jun': '06',
	'July': '07', 'Jul': '07',
	'August': '08', 'Aug': '08',
	'September': '09', 'Sep': '09',
	'October': '10', 'Oct': '10',
	'November': '11', 'Nov': '11',
	'December': '12', 'Dec': '12'
}


class hpx_reverse_auction():
	def __init__(self):
		pass
	
	def download_url(self, url, destination):
		response = requests.get(url)
		with open(destination, 'wb') as f:
			f.write(response.content)
	
	def edit_table_for_file(self):
		file_name_file_path = os.path.join(main_directory, 'table_data.xlsx')
		df = pd.read_excel(file_name_file_path)
		df['Entity Name'] = df['Entity Name'].astype(str).str.split('HPX').str[1:]
		df['Entity Name'] = df['Entity Name'].astype(str).str.split('_').str[0]
		df['Entity Name'] = df['Entity Name'].astype(str).str.split(':').str[0]
		df['Entity Name'] = df['Entity Name'].str.replace('-', '/')
		df['Entity Name'] = df['Entity Name'].str.replace(',', '[')
		df['Entity Name'] = df['Entity Name'].str.replace("'", '.')
		df['Entity Name'] = df['Entity Name'].str.replace("[", 'H')
		df['Entity Name'] = df['Entity Name'].str.replace(".", 'P')
		df['Entity Name'] = df['Entity Name'].astype(str).str.split("P/").str[-1]
		df['Entity Name'] = df['Entity Name'].str.replace('/', '_')
		df['Entity Name'] = df['Entity Name'].str.replace(' ', '')
		
		# Remove ordinal suffixes ('th', 'st', 'rd', 'nd') from 'Auction Date' column
		df['Auction Date'] = df['Auction Date'].apply(lambda x: re.sub(r'(?:st|nd|rd|th|,)', '', x))
		
		for month_name, month_num in month_replacements.items():
			df['Auction Date'] = df['Auction Date'].str.replace(month_name, month_num)
		df['Auction Date'] = df['Auction Date'].str.replace(' ', '-')
		# Convert 'Auction Date' column to datetime format
		df['Auction Date'] = pd.to_datetime(df['Auction Date'], format='%d-%m-%Y').dt.date
		
		df.rename(columns={'Entity Name': 'Auction No.'}, inplace=True)
		df.drop(columns=['Auction Details'], inplace=True)
		df.to_excel(file_name_file_path, index=False)
	
	def get_pdf_files_url(self):
		options = Options()
		prefs = {'download.default_directory': r'C:\GNA\Data\Reverse Auction\HPX Reverse Auction'}
		if not os.path.exists(prefs['download.default_directory']):
			os.makedirs(prefs['download.default_directory'])
		options.add_experimental_option("prefs", prefs)
		chromedriver_path = r'C:\Users\pulki\.cache\selenium\chromedriver\win64\125.0.6422.76\chromedriver.exe'
		driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
		
		# driver = webdriver.Chrome(options=options)
		driver.get('https://www.hpxindia.com/auction_detail.html')
		
		table_find = WebDriverWait(driver, 10).until(
			EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/div/div/div/div/table/tbody/tr/td/table')))
		# Extract table data
		table = driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div/div/div/table/tbody/tr/td/table')
		table_html = table.get_attribute('outerHTML')
		df = pd.read_html(table_html)[0]  # Read HTML table into DataFrame
		
		# Store DataFrame into Excel file
		file_name_file_path = os.path.join(main_directory, 'table_data.xlsx')
		df.to_excel(file_name_file_path, index=False)
		links = table_find.find_elements(By.TAG_NAME, 'a')
		
		for i, link in enumerate(links[:10]):
		# for link in links:
			href = link.get_attribute('href')
			filename = href.split('/')[-1]
			print(filename)
			destination = f"{prefs['download.default_directory']}/{filename}"
			hpx_reverse_auction.download_url(self, href, destination)
		
		hpx_reverse_auction.edit_table_for_file(self)
	
	def extract_text_from_pdf(self, file_path):
		with open(file_path, 'rb') as file:
			reader = PdfReader(file)
			buyer_names = []
			auction_nos = []
			# Iterate through each page
			for page in reader.pages:
				# Extract text from the page
				text = page.extract_text()
				auction_no = text.find('HPX/')
				auction_no_end = text.find('Buyer')
				# Find the index of 'Buyer Name'
				index_buyer_name = text.find('Name')
				# Find the index of 'Requisition'
				index_requisition = text.find('Requisition')
				if index_buyer_name != -1 and index_requisition != -1:
					# Extract the text between 'Buyer Name' and 'Requisition'
					buyer_name = text[index_buyer_name + len('Name:'):index_requisition].strip() or text[
					                                                                                index_buyer_name + len(
						                                                                                'Name:'):].strip()
					buyer_names.append(buyer_name)
				elif index_buyer_name != -1:
					# Extract the text after 'Buyer Name' if 'Requisition' is not found
					buyer_name = text[index_buyer_name + len('Name:'):].strip()
					buyer_names.append(buyer_name)
				
				if auction_no != -1 and auction_no_end != -1:
					auction_no_value = text[auction_no + len('HPX/'):auction_no_end].strip()
					auction_nos.append(auction_no_value)
				elif auction_no != -1:
					auction_no_value = text[auction_no + len('HPX/'):].strip()
					auction_nos.append(auction_no_value)
		return buyer_names, auction_nos
	
	def pdf_to_xlsx(self):
		error_files = []
		pdf_files = [file for file in os.listdir(file_directory) if file.endswith('.pdf')]
		
		for file in pdf_files:
			file_path = os.path.join(file_directory, file)
			output_path = os.path.join(output_directory, os.path.splitext(file)[0] + '.xlsx')
			
			try:
				buyer_names, auction_nos = hpx_reverse_auction.extract_text_from_pdf(self, file_path)
				
				# Extract tables from PDF using tabula
				tables = tabula.read_pdf(file_path, pages='all', multiple_tables=True)
				if tables:
					# Ensure lengths of buyer names and auction numbers match the number of tables
					num_tables = len(tables)
					buyer_names = buyer_names[:num_tables]
					auction_nos = auction_nos[:num_tables]
					
					for i, table in enumerate(tables):
						table.insert(0, 'Auction No.', auction_nos[i])
						table.insert(1, 'Buyer', buyer_names[i])
					
					df = pd.concat(tables, ignore_index=True)
					
					df.replace(to_replace='_x000D_', value=' ', regex=True, inplace=True)
					
					df.to_excel(output_path, index=False)
					print(f"Successfully extracted tables from {file} and saved to {output_path}")
				else:
					print(f"No tables found in {file}")
			except Exception as e:
				print(f"Error processing file '{file}': {e}")
				error_files.append(file)
		if error_files:
			try:
				pd.DataFrame({'Error Files': error_files}).to_excel(error_log_file, index=False)
				print(f"Error log saved to '{error_log_file}'")
			except Exception as e:
				print(f"Error saving error log: {e}")
	
	def edit_xlsx_file(self):
		xlsx_files = []
		for file in os.listdir(output_directory):
			if file.endswith('.xlsx'):
				xlsx_files.append(file)
				
		for xlsx_file in xlsx_files:
			file_path = os.path.join(output_directory, xlsx_file)
			output_file = os.path.join(output_directory, os.path.splitext(xlsx_file)[0] + '.xlsx')
			
			print(xlsx_file)
			df = pd.read_excel(file_path)
			# Clean column names and replace values
			df.columns = df.columns.str.strip()
			
			df.replace(to_replace='_x000D_', value=' ', regex=True, inplace=True)
			df.replace(to_replace='\r', value=' ', regex=True, inplace=True)
			df.columns = df.columns.str.replace('_x000D_', ' ')
			df.columns = df.columns.str.replace('\r', ' ')
			
			# Rename specific column headings
			df.rename(columns={'Requisitioned Qty (Min. Bid Qty)': 'Requisitioned Qty (Min. Bid Qty) (In MW)',
			                   'Requisitioned Qty\r(Min. Bid Qty)\r(In MW)': 'Requisitioned Qty (Min. Bid Qty) (In MW)',
			                   'Allocated Qty to Seller': 'Allocated Qty to Seller (In MW)',
			                   'Accepted Price': 'Discovered/Accepted Price (Rs./MWh)',
			                   'Booking Qty (Min. Qty)': 'Booking Qty (Min. Qty) (In MW)',
			                   'Accepted Price (Rs/MWh)': 'Discovered/Accepted Price (Rs./MWh)',
			                   'Discovered Price (Rs./MWh)': 'Discovered/Accepted Price (Rs./MWh)',
			                   'Discovered Qty to Seller (In MW)': 'Allocated Qty to Seller (In MW)',
			                   'Requisition\rNo': 'Requisition No',
			                   'Requisitioned Qty\r(Min. Bid Qty)': 'Requisitioned Qty (Min. Bid Qty) (In MW)',
			                   'Booking Qty\r(Min. Qty)': 'Booking Qty (Min. Qty) (In MW)',
			                   'Allocated Qty\rto Seller': 'Allocated Qty to Seller (In MW)',
			                   'Requisition \nNo': 'Requisition No'
			                   }, inplace=True)
			
			for index, row in df.iterrows():
				if 'IPO' in str(row['Requisition No']) or 'RA' in str(row['Requisition No']):
					print(f"Row: {index + 2}, Column: {df.columns.get_loc('Requisition No') + 1}")
					df.at[index, 'Type'] = df.at[index, 'Requisition No']
					df.at[index, 'Requisition No'] = ''
					df.at[index, 'Booking Qty (Min. Qty) (In MW)'] = df.at[index, 'From Delivery Period']
					df.at[index, 'From Delivery Period'] = ''
					df.at[index, 'Allocated Qty to Seller (In MW)'] = df.at[index, 'To Delivery Period']
					df.at[index, 'To Delivery Period'] = ''
					if 'Auction Type' in df.columns:  # Check if 'Auction Type' column exists
						df.at[index, 'Discovered/Accepted Price (Rs./MWh)'] = df.at[index, 'Auction Type']
						df.at[index, 'Auction Type'] = ''
					else:
						df.at[index, 'Discovered/Accepted Price (Rs./MWh)'] = df.at[
							index, 'Requisitioned Qty (Min. Bid Qty) (In MW)']
						df.at[index, 'Requisitioned Qty (Min. Bid Qty) (In MW)'] = ''
						
				if 'IPO' in str(row['Requisitioned Qty (Min. Bid Qty) (In MW)']) or 'RA' in str(
						row['Requisitioned Qty (Min. Bid Qty) (In MW)']):
					df.at[index, 'Discovered/Accepted Price (Rs./MWh)'] = df.at[
						index, 'Allocated Qty to Seller (In MW)']
					df.at[index, 'Allocated Qty to Seller (In MW)'] = df.at[index, 'Booking Qty (Min. Qty) (In MW)']
					df.at[index, 'Booking Qty (Min. Qty) (In MW)'] = df.at[index, 'Type']
					df.at[index, 'Type'] = df.at[index, 'Requisitioned Qty (Min. Bid Qty) (In MW)']
					df.at[index, 'Requisitioned Qty (Min. Bid Qty) (In MW)'] = ''
			
			# Split columns and assign values
			df[['Delivery Start Time', 'Delivery Start Date']] = df['From Delivery Period'].astype(str).str.split(' ',
			                                                                                                      expand=True)
			df[['Delivery End Time', 'Delivery End Date']] = df['To Delivery Period'].astype(str).str.split(' ',
			                                                                                                expand=True)
			# Drop original columns
			df.drop(['From Delivery Period', 'To Delivery Period'], axis=1, inplace=True)
			columns_to_fill = ['Requisition No', 'Requisitioned Qty (Min. Bid Qty) (In MW)', 'Delivery Start Time',
			                   'Delivery Start Date', 'Delivery End Time', 'Delivery End Date']
			for col in columns_to_fill:
				df[col].replace('', np.nan, inplace=True)
			df[columns_to_fill] = df[columns_to_fill].ffill()
			
			if 'Auction Type' in df.columns:
				df['Auction Type'] = df['Auction Type'].ffill()
			else:
				pass
			print(df)
			
			df.to_excel(output_file, index=False)
			print(f"File Edited successfully {output_file}")
	
	def clean_date(self, date):
		if '(' in date:
			return date.split('(')[1].split(')')[0]
		else:
			return date
	
	def merge_xlsx_file(self):
		xlsx_files = [file for file in os.listdir(output_directory) if file.endswith('.xlsx')]
		merged_df = pd.DataFrame()
		
		for xlsx_file in xlsx_files:
			file_path = os.path.join(output_directory, xlsx_file)
			try:
				df = pd.read_excel(file_path)
				merged_df = pd.concat([merged_df, df], ignore_index=True)
			except Exception as e:
				print(f"Error reading file '{xlsx_file}': {e}")
		
		if not merged_df.empty:
			column_to_fill = ['Delivery Start Time', 'Delivery End Time']
			merged_df[column_to_fill] = merged_df[column_to_fill].ffill()
			merged_df['Delivery Start Date'] = merged_df['Delivery Start Date'].apply(self.clean_date)
			merged_df['Delivery End Date'] = merged_df['Delivery End Date'].apply(self.clean_date)
			merged_df['Auction No.'] = merged_df['Auction No.'].str.replace('/', '_')
			
			for month_name, month_num in month_replacements.items():
				merged_df['Delivery Start Date'] = merged_df['Delivery Start Date'].str.replace(month_name, month_num)
				merged_df['Delivery End Date'] = merged_df['Delivery End Date'].str.replace(month_name, month_num)
			merged_df['Delivery Start Date'] = pd.to_datetime(merged_df['Delivery Start Date'], format='%d-%m-%Y')
			merged_df['Delivery End Date'] = pd.to_datetime(merged_df['Delivery End Date'], format='%d-%m-%Y')
			
			merged_df['Discovered/Accepted Price (Rs./MWh)'] = \
				merged_df['Discovered/Accepted Price (Rs./MWh)'].astype(str).str.split('Rs ').str[1].str.split(
					' /').str[0]
			merged_df['Discovered/Accepted Price (Rs./MWh)'] = merged_df['Discovered/Accepted Price (Rs./MWh)'].astype(
				float)
			merged_df['Discovered/Accepted Price (Rs./MWh)'] /= 1000
			
			merged_df[['Booking Quantity (in MW)', 'Minimum Booking Quantity (in MW)']] = merged_df[
				'Booking Qty (Min. Qty) (In MW)'].astype(str).str.split(r' ', expand=True)
			merged_df[
				['Requisitioned/Requested Quantity (in MW)', 'Minimum Requisitioned/Requested Quantity (in MW)']] = \
				merged_df['Requisitioned Qty (Min. Bid Qty) (In MW)'].astype(str).str.split(r' ', expand=True)
			merged_df['Minimum Booking Quantity (in MW)'] = \
				merged_df['Minimum Booking Quantity (in MW)'].astype(str).str.split('(').str[1].str.split(')').str[0]
			merged_df['Minimum Requisitioned/Requested Quantity (in MW)'] = \
				merged_df['Minimum Requisitioned/Requested Quantity (in MW)'].astype(str).str.split('(').str[
					1].str.split(
					')').str[0]
			
			merged_df.drop(['Requisitioned Qty (Min. Bid Qty) (In MW)', 'Booking Qty (Min. Qty) (In MW)'], axis=1,
			               inplace=True)
			merged_df['Booking Quantity (in MW)'] = merged_df['Booking Quantity (in MW)'].str.replace('nan','')
			
			try:
				output_path = os.path.join(main_directory, 'merged_file_hpx.xlsx')
				merged_df.to_excel(output_path, index=False)
				print(f"Merged file saved to '{output_path}'")
			except Exception as e:
				print(f"Error saving merged file: {e}")
		else:
			print("No data to merge.")
	
	def merge_final_files(self):
		file_1 = r"C:\GNA\Data\Reverse Auction\merged_file_hpx.xlsx"
		file_2 = r"C:\GNA\Data\Reverse Auction\table_data.xlsx"
		
		df1 = pd.read_excel(file_1)
		df2 = pd.read_excel(file_2)
		
		df1['Auction No.'] = df1['Auction No.'].astype(str)
		df2['Auction No.'] = df2['Auction No.'].astype(str)
		
		merged_df = pd.merge(df1, df2, on="Auction No.", how="left")
		merged_df.insert(2, 'Exchange Type', 'HPX')
		
		merged_df['Auction Date'] = merged_df['Auction Date'].dt.date
		merged_df['Delivery Start Date'] = merged_df['Delivery Start Date'].dt.date
		merged_df['Delivery End Date'] = merged_df['Delivery End Date'].dt.date
		
		merged_df['Delivery Start Time'] = merged_df['Delivery Start Time'].astype(str).replace(' ', '', regex=True)
		# merged_df['Delivery Start Time'] = pd.to_datetime(merged_df['Delivery Start Time'], format='%H:%M').dt.time
		merged_df['Delivery End Time'] = merged_df['Delivery End Time'].astype(str).replace(' ', '', regex=True)
		
		column_to_move = [
			'Exchange Type',
			'Auction No.',
			'Auction Date',
			'Buyer',
			'Requisition No',
			'Delivery Start Date',
			'Delivery End Date',
			'Delivery Start Time',
			'Delivery End Time',
			'Requisitioned/Requested Quantity (in MW)',
			'Minimum Requisitioned/Requested Quantity (in MW)',
			'Type',
			'Booking Quantity (in MW)',
			'Minimum Booking Quantity (in MW)',
			'Allocated Qty to Seller (In MW)',
			'Discovered/Accepted Price (Rs./MWh)'
		]
		remaining_column_order = [col for col in merged_df.columns if col not in column_to_move]
		column_order = column_to_move + remaining_column_order
		merged_df = merged_df[column_order]
		
		merged_df.rename(columns={
			'Allocated Qty to Seller (In MW)': 'Allocated Quantity (in MW)',
			'Discovered/Accepted Price (Rs./MWh)': 'Accepted Price (in Rs./kWh)',
			'Requisitioned/Requested Quantity (in MW)': 'Buy Total Quantity (in MW)',
			'Minimum Requisitioned/Requested Quantity (in MW)': 'Buy Minimum Quantity (in MW)',
			'Auction Date': 'Auction Result Date'
		}, inplace=True)
		
		merged_df = merged_df.sort_values(by='Auction Result Date', ascending=False)
		output_path = os.path.join(final_directory, 'final_hpx.xlsx')
		merged_df.to_excel(output_path, index=False)
		print(f"Merged file saved to {output_path}")
	
	
	def get_data(self):
		hpx_reverse_auction.get_pdf_files_url(self)
		hpx_reverse_auction.pdf_to_xlsx(self)
		hpx_reverse_auction.edit_xlsx_file(self)
		hpx_reverse_auction.merge_xlsx_file(self)
		hpx_reverse_auction.merge_final_files(self)
		pass


if __name__ == '__main__':
	tam_hpx = hpx_reverse_auction()
	tam_hpx.get_data()
	pass
