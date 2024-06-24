import os
import time
import xlwings
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import tabula
from datetime import datetime
from fuzzywuzzy import fuzz
import data_mapping
from itertools import combinations


class deep_portal:
	def __init__(self):
		self.main_directory = r'C:\GNA\Data\Deep Portal'
		self.file_directory = r'C:\GNA\Data\Deep Portal\Downloaded Files'
		self.output_directory = r'C:\GNA\Data\Deep Portal\Edited xlsx Files'
		self.final_directory = r'C:\GNA\Data Upload'

		self.pdf_file_directory = r'C:\GNA\Data\Deep Portal\pdf Edited xlsx Files'
		if not os.path.exists(self.pdf_file_directory):
			os.makedirs(self.pdf_file_directory)

		self.output_pdf_directory = r'C:\GNA\Data\Deep Portal\pdf_file Edited xlsx Files'
		if not os.path.exists(self.output_pdf_directory):
			os.makedirs(self.output_pdf_directory)
		pass
	
	def deep_portal_file_download(self):
		if os.path.exists(self.file_directory):
			for file in os.listdir(self.file_directory):
				file_path_full = os.path.join(self.file_directory, file)
				if os.path.isfile(file_path_full):
					os.remove(file_path_full)
		else:
			os.makedirs(self.file_directory)
		
		options = Options()
		prefs = {'download.default_directory': self.file_directory}
		options.add_experimental_option('prefs', prefs)
		chromedriver_path = r'C:\Users\pulki\.cache\selenium\chromedriver\win64\125.0.6422.76\chromedriver.exe'
		driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
		# driver = webdriver.Chrome(options=options)
		
		driver.get(
			'https://www.mstcecommerce.com/auctionhome/container.jsp?title_id=Rev.%20Auction%20Result&linkid=0&main_link=y&sublink=n&main_link_name=285&portal=ppa&homepage=index&arcDate=31-10-2021')
		
		table_find = WebDriverWait(driver, 15).until(
			EC.presence_of_element_located((By.XPATH, '/html/body/div/div[3]/div[2]/table')))
		
		print('Links in Page:')
		links = table_find.find_elements(By.TAG_NAME, 'a')  # Corrected to find_elements
		
		# for i, link in enumerate(links):
		for i, link in enumerate(links[:5]):
			print(link.get_attribute('href'))
			link.click()
			time.sleep(1)
		driver.quit()
		
		deep_portal.xls_to_xlsx(self)
	
	def xls_to_xlsx(self):
		xls_files = []
		for filename in os.listdir(self.file_directory):
			if filename.endswith('.xls'):
				xls_files.append(filename)
		for xls_file in xls_files:
			file_path = os.path.join(self.file_directory, xls_file)
			output_path = os.path.join(self.file_directory, os.path.splitext(xls_file)[0] + '.xlsx')
			try:
				app = xlwings.App(visible=False)
				workbook = app.books.open(file_path)
				workbook.save(output_path)
				workbook.close()
				app.quit()
				print(f"Conversion completed. File saved: {xls_file}")
			except xlwings.XlwingsError as e:
				print(f"Error converting {xls_file}: {e}")
			except Exception as e:
				print(f"Error converting {xls_file}: {e}")
				continue  # Move to the next file
		
		deep_portal.edit_xlsx_file(self)
	
	def pdf_to_xlsx(self):
		pdf_files = []
		for filename in os.listdir(self.file_directory):
			if filename.endswith('.pdf'):
				pdf_files.append(filename)
		for pdf_file in pdf_files:
			file_path = os.path.join(self.file_directory, pdf_file)
			output_path = os.path.join(self.pdf_file_directory, os.path.splitext(pdf_file)[0] + '.xlsx')
			try:
				tables = tabula.read_pdf(file_path, pages='all', multiple_tables=True)
				with pd.ExcelWriter(output_path) as writer:
					for i, table in enumerate(tables):
						table.to_excel(writer, sheet_name=f'Table_{i + 1}', index=False)
						print((f'Table Excracted. File saved at: {output_path}'))
			except Exception as e:
				print(e)
	
	def edit_xlsx_file(self):
		if os.path.exists(self.output_directory):
			for file in os.listdir(self.output_directory):
				file_path_full = os.path.join(self.output_directory, file)
				if os.path.isfile(file_path_full):
					os.remove(file_path_full)
		else:
			os.makedirs(self.output_directory)
		
		xlsx_files = []
		for filename in os.listdir(self.file_directory):
			if filename.endswith('.xlsx'):
				xlsx_files.append(filename)
		
		for xlsx_file in xlsx_files:
			file_path = os.path.join(self.file_directory, xlsx_file)
			output_path = os.path.join(self.output_directory, os.path.splitext(xlsx_file)[0] + '.xlsx')
			
			print(xlsx_file)
			df = pd.read_excel(file_path)
			df = df.dropna(how='all')
			
			# Unmerge merged cells
			for index, merged_info in df.iterrows():
				for column in df.columns:
					if isinstance(merged_info[column], list):  # Check if cell is merged
						for cell in merged_info[column][1:]:  # Set other cells to None to 'unmerge'
							df.at[index, column] = None
			
			# Select the first four columns
			first_four_columns = df.columns[:4]
			df[first_four_columns] = df[first_four_columns].ffill()
			df.insert(2, 'Auction Date', df.iloc[0, 1])
			
			# Set the header of the second column as 'Tender Name'
			heading_2nd_col = df.columns[1]
			df.insert(0, "Tender Name", heading_2nd_col)
			
			# Copy 'Tender Name' and create a new column before it
			df.insert(1, "New Column", df['Tender Name'].apply(lambda x: x.split('/')[0]))
			# Rename columns
			df.columns = ['Tender Details', 'Tender Name', 'Sl.No.', 'Requisition Details', 'Auction Date',
			              'Quantity Requisitioned',
			              'Description',
			              'Booking Quantity', 'Booking Amount', 'Allotted Quantity', 'Accepted price', 'Type(IPO/RA)']
			
			# Drop unnecessary rows
			df = df.iloc[2:]
			
			# Split 'Description' column and concatenate new columns
			description_split = df['Description'].str.split(r'[:\s]+', expand=True)
			description_split.columns = [f'Description_{i + 1}' for i in range(description_split.shape[1])]
			df = pd.concat([df, description_split], axis=1)
			
			df['Description_3'] = df['Description_3'].str.replace('.', ':')
			df['Description_6'] = df['Description_6'].str.replace('.', ':')
			
			df[['Auction Initiation Date', 'Auction Initiation Time', 'To', 'Auction Result Date',
			    'Auction Result Time']] = \
				df['Auction Date'].astype(str).str.split(' ', expand=True)
			
			# Drop unnecessary columns
			df = df.drop(columns=['Description', 'Description_1', 'Description_4', 'Auction Date', 'To', 'Sl.No.'],
			             axis=1)
			
			df['Quantity Requisitioned'] = df['Quantity Requisitioned'].astype(str)
			df['Quantity Requisitioned'] = df['Quantity Requisitioned'].replace('MW', '', regex=True)
			df['Quantity Requisitioned'] = df['Quantity Requisitioned'].replace('\(', '', regex=True)
			df['Quantity Requisitioned'] = df['Quantity Requisitioned'].replace('\)', '', regex=True)
			df['Quantity Requisitioned'] = df['Quantity Requisitioned'].astype(float)
			
			# Replace '.' with '-' in 'Start Date' and 'End Date' and convert to datetime
			if 'Description_2' in df.columns and 'Description_5' in df.columns:
				df[['Description_2', 'Description_5']] = df[['Description_2', 'Description_5']].astype(str).apply(
					lambda x: x.str.replace('.', '-'))
				df['Description_2'] = pd.to_datetime(df['Description_2'], format='%d-%m-%Y', errors='coerce').dt.date
				df['Description_5'] = pd.to_datetime(df['Description_5'], format='%d-%m-%Y', errors='coerce').dt.date
			
			# Convert 'Auction Initiation Date' and 'Auction Result Date' in date format
			df[['Auction Initiation Date', 'Auction Result Date']] = df[
				['Auction Initiation Date', 'Auction Result Date']].astype(str).apply(lambda x: x.str.replace('/', '-'))
			df['Auction Initiation Date'] = pd.to_datetime(df['Auction Initiation Date'], format='%d-%m-%Y',
			                                               errors='coerce').dt.date
			df['Auction Result Date'] = pd.to_datetime(df['Auction Result Date'], format='%d-%m-%Y',
			                                           errors='coerce').dt.date
			
			df['Accepted price'] = pd.to_numeric(df['Accepted price'], errors='coerce')
			df['Accepted price'] = df['Accepted price'].astype(float)
			
			# Rename Columns
			df.rename(columns={
				'Tender Details': 'Auction No.',
				'Tender Name': 'Buyer',
				'Requisition Details': 'Requisition No',
				'Quantity Requisitioned': 'Buy Total Quantity (in MW)',
				'Description_2': 'Delivery Start Date',
				'Description_5': 'Delivery End Date',
				'Description_3': 'Delivery Start Time',
				'Description_6': 'Delivery End Time',
				'Type(IPO/RA)': 'Type',
				'Booking Quantity': 'Booking Quantity (in MW)',
				'Allotted Quantity': 'Allocated Quantity (in MW)',
				'Booking Amount': 'Booking Accepted Price (in Rs./kWh)',
				'Accepted price': 'Accepted Price (in Rs./kWh)'
			}, inplace=True)
			
			# Arrange Columns in Order
			column_to_move = [
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
				'Type',
				'Buy Total Quantity (in MW)',
				'Booking Quantity (in MW)',
				'Allocated Quantity (in MW)',
				'Booking Accepted Price (in Rs./kWh)',
				'Accepted Price (in Rs./kWh)'
			]
			remaining_column_order = [col for col in df.columns if col not in column_to_move]
			column_order = column_to_move + remaining_column_order
			df = df[column_order]
			
			df.to_excel(output_path, index=False)
			print(f'File Edited {xlsx_file}')
	
	def edit_pdf_files(self):
		xlsx_files = []
		for filename in os.listdir(self.pdf_file_directory):
			if filename.endswith('.xlsx'):
				xlsx_files.append(filename)
		
		for xlsx_file in xlsx_files:
			file_path = os.path.join(self.pdf_file_directory, xlsx_file)
			output_path = os.path.join(self.output_pdf_directory, os.path.splitext(xlsx_file)[0] + '.xlsx')
			
			print(xlsx_file)
			df = pd.read_excel(file_path)
			df = df.dropna(how='all')
			
			# Unmerge merged cells
			for index, merged_info in df.iterrows():
				for column in df.columns:
					if isinstance(merged_info[column], list):  # Check if cell is merged
						for cell in merged_info[column][1:]:  # Set other cells to None to 'unmerge'
							df.at[index, column] = None
			
			# Select the first four columns
			first_four_columns = df.columns[:4]
			df[first_four_columns] = df[first_four_columns].ffill()
			df.insert(2, 'Auction Date', df.iloc[0, 1])
			
			# Set the header of the second column as 'Tender Name'
			heading_2nd_col = df.columns[1]
			df.insert(0, "Tender Name", heading_2nd_col)
			
			# Copy 'Tender Name' and create a new column before it
			df.insert(1, "New Column", df['Tender Name'].apply(lambda x: x.split('/')[0]))
			
			# Rename columns
			df.columns = ['Tender Details', 'Tender Name', 'Sl.No.', 'Requisition Details', 'Auction Date',
			              'Quantity Requisitioned',
			              'Description',
			              'Booking Quantity', 'Booking Amount', 'Allotted Quantity', 'Accepted price', 'Type(IPO/RA)']
			
			# Drop unnecessary rows
			df = df.iloc[2:]
			
			# Split 'Description' column and concatenate new columns
			df['Description'] = df['Description'].str.replace('to_x000D_', 'to ')
			description_split = df['Description'].str.split(r'[:\s]+', expand=True)
			description_split.columns = [f'Description_{i + 1}' for i in range(description_split.shape[1])]
			df = pd.concat([df, description_split], axis=1)
			
			df['Description_3'] = df['Description_3'].str.replace('.', ':')
			df['Description_6'] = df['Description_6'].str.replace('.', ':')
			
			df[['Auction Initiation Date', 'Auction Initiation Time', 'To', 'Auction Result Date',
			    'Auction Result Time']] = \
				df['Auction Date'].astype(str).str.split(' ', expand=True)
			
			# Drop unnecessary columns
			df = df.drop(columns=['Description', 'Description_1', 'Description_4', 'Auction Date', 'To', 'Sl.No.'],
			             axis=1)
			
			df['Quantity Requisitioned'] = df['Quantity Requisitioned'].astype(str)
			df['Quantity Requisitioned'] = df['Quantity Requisitioned'].replace('MW', '', regex=True)
			df['Quantity Requisitioned'] = df['Quantity Requisitioned'].replace('\(', '', regex=True)
			df['Quantity Requisitioned'] = df['Quantity Requisitioned'].replace('\)', '', regex=True)
			df['Quantity Requisitioned'] = df['Quantity Requisitioned'].astype(float)
			
			# Replace '.' with '-' in 'Start Date' and 'End Date' and convert to datetime
			if 'Description_2' in df.columns and 'Description_5' in df.columns:
				df[['Description_2', 'Description_5']] = df[['Description_2', 'Description_5']].astype(str).apply(
					lambda x: x.str.replace('.', '-'))
				df['Description_2'] = pd.to_datetime(df['Description_2'], format='%d-%m-%Y', errors='coerce').dt.date
				df['Description_5'] = pd.to_datetime(df['Description_5'], format='%d-%m-%Y', errors='coerce').dt.date
			
			# Convert 'Auction Initiation Date' and 'Auction Result Date' in date format
			df[['Auction Initiation Date', 'Auction Result Date']] = df[
				['Auction Initiation Date', 'Auction Result Date']].astype(str).apply(lambda x: x.str.replace('/', '-'))
			df['Auction Initiation Date'] = pd.to_datetime(df['Auction Initiation Date'], format='%d-%m-%Y',
			                                               errors='coerce').dt.date
			df['Auction Result Date'] = pd.to_datetime(df['Auction Result Date'], format='%d-%m-%Y',
			                                           errors='coerce').dt.date
			
			df['Accepted price'] = pd.to_numeric(df['Accepted price'], errors='coerce')
			df['Accepted price'] = df['Accepted price'].astype(float)
			
			# Rename Columns
			df.rename(columns={
				'Tender Details': 'Auction No.',
				'Tender Name': 'Buyer',
				'Requisition Details': 'Requisition No',
				'Quantity Requisitioned': 'Buy Total Quantity (in MW)',
				'Description_2': 'Delivery Start Date',
				'Description_5': 'Delivery End Date',
				'Description_3': 'Delivery Start Time',
				'Description_6': 'Delivery End Time',
				'Type(IPO/RA)': 'Type',
				'Booking Quantity': 'Booking Quantity (in MW)',
				'Allotted Quantity': 'Allocated Quantity (in MW)',
				'Booking Amount': 'Booking Accepted Price (in Rs./kWh)',
				'Accepted price': 'Accepted Price (in Rs./kWh)'
			}, inplace=True)
			
			# Arrange Columns in Order
			column_to_move = [
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
				'Type',
				'Buy Total Quantity (in MW)',
				'Booking Quantity (in MW)',
				'Allocated Quantity (in MW)',
				'Booking Accepted Price (in Rs./kWh)',
				'Accepted Price (in Rs./kWh)'
			]
			remaining_column_order = [col for col in df.columns if col not in column_to_move]
			column_order = column_to_move + remaining_column_order
			df = df[column_order]
			
			df.to_excel(output_path, index=False)
			print(f'File Edited {xlsx_file}')
	
	def merge_file(self):
		xlsx_files = []
		for filename in os.listdir(self.output_directory):
			if filename.endswith('.xlsx'):
				xlsx_files.append(filename)
		
		merged_data = pd.DataFrame()
		
		for xlsx_file in xlsx_files:
			file_path = os.path.join(self.output_directory, xlsx_file)
			
			try:
				data = pd.read_excel(file_path)
				# Append data to the merged DataFrame if it's not empty
				if not data.empty:
					merged_data = pd.concat([merged_data, data], ignore_index=True)
				else:
					print(f"Data is empty in {xlsx_file}")
			except Exception as e:
				print(f"Failed to read {xlsx_file}: {e}")
		
		# Convert column in date format
		merged_data['Auction Initiation Date'] = pd.to_datetime(merged_data['Auction Initiation Date']).dt.date
		merged_data['Auction Result Date'] = pd.to_datetime(merged_data['Auction Result Date']).dt.date
		merged_data['Delivery Start Date'] = pd.to_datetime(merged_data['Delivery Start Date']).dt.date
		merged_data['Delivery End Date'] = pd.to_datetime(merged_data['Delivery End Date']).dt.date
		
		# Convert column in time format
		merged_data['Auction Initiation Time'] = pd.to_datetime(merged_data['Auction Initiation Time'],
		                                                        format='%H:%M:%S').dt.time
		merged_data['Auction Result Time'] = pd.to_datetime(merged_data['Auction Result Time'],
		                                                    format='%H:%M:%S').dt.time
		merged_data['Delivery Start Time'] = pd.to_datetime(merged_data['Delivery Start Time'], format='%H:%M').dt.time
		merged_data['Delivery End Time'] = merged_data['Delivery End Time'].str.replace('24:00', '23:59')
		merged_data['Delivery End Time'] = pd.to_datetime(merged_data['Delivery End Time'], format='%H:%M').dt.time
		merged_data['Delivery End Time'] = merged_data['Delivery End Time'].astype(str).str.replace('23:59:00',
		                                                                                            '23:59:59')
		merged_data['Delivery End Time'] = pd.to_datetime(merged_data['Delivery End Time'], format='%H:%M:%S').dt.time
		
		merged_data.insert(0, 'Exchange Type', 'DEEP')
		merged_data = merged_data.sort_values(by=['Auction Initiation Date', 'Requisition No'], ascending=[False, True])
		
		merged_output_path = os.path.join(self.final_directory, 'deep_portal.xlsx')
		merged_data.to_excel(merged_output_path, index=False)
		print(f'File Saved at {merged_output_path}')
	
	def remove_duplicate_buyer_name(self):
		deep_file_path = os.path.join(self.final_directory, 'deep_portal.xlsx')
		df = pd.read_excel(deep_file_path)
		
		unique_names = df['Buyer'].unique()
		
		for name in unique_names:
			for key, value in data_mapping.deep_data_mapping.items():
				key_similarity_ratio = fuzz.ratio(name, key)
				value_similarity_ratio = fuzz.ratio(name, value)
				if key_similarity_ratio >= 91 or value_similarity_ratio >= 91:
					df['Buyer'] = df['Buyer'].str.replace(name, key)
		
		output_path = os.path.join(self.final_directory, 'deep_portal.xlsx')
		df.to_excel(output_path, index=False)
		print(f'Duplicates removed from {deep_file_path}')
	
	def data_mapping(self):
		deep_file_path = os.path.join(self.final_directory, 'deep_portal.xlsx')
		df = pd.read_excel(deep_file_path)
		
		for name, mane_to_change in data_mapping.deep_data_mapping.items():
			df['Buyer'] = df['Buyer'].str.replace(name, mane_to_change)
		
		# Convert column in date format
		df['Auction Initiation Date'] = pd.to_datetime(df['Auction Initiation Date']).dt.date
		df['Auction Result Date'] = pd.to_datetime(df['Auction Result Date']).dt.date
		df['Delivery Start Date'] = pd.to_datetime(df['Delivery Start Date']).dt.date
		df['Delivery End Date'] = pd.to_datetime(df['Delivery End Date']).dt.date
		
		# Convert column in time format
		df['Auction Initiation Time'] = pd.to_datetime(df['Auction Initiation Time']).dt.time
		df['Auction Result Time'] = pd.to_datetime(df['Auction Result Time']).dt.time
		df['Delivery Start Time'] = pd.to_datetime(df['Delivery Start Time']).dt.time
		df['Delivery End Time'] = pd.to_datetime(df['Delivery End Time']).dt.time
		
		output_path = os.path.join(self.final_directory, 'deep_portal.xlsx')
		df.to_excel(output_path, index=False)
		print(f'Data Mapping done file saved at {deep_file_path}')
	
	def get_data(self):
		deep_portal.deep_portal_file_download(self)
		deep_portal.merge_file(self)
		deep_portal.remove_duplicate_buyer_name(self)
		deep_portal.data_mapping(self)
		pass


class deep_bulletin_board():
	def __init__(self):
		self.deep_bulletin_board_directory = r'C:\GNA\Data\Deep Portal\Deep-e-bidding'
		if not os.path.exists(self.deep_bulletin_board_directory):
			os.makedirs(self.deep_bulletin_board_directory)
			pass
	
	def save_to_excel(self, table_data):
		today_date = datetime.now().strftime('%d-%m-%Y')
		filename = f'Deep e-bidding_{today_date}.xlsx'
		output_path = os.path.join(self.deep_bulletin_board_directory, filename)
		if not table_data:
			print("No table data to save.")
			return
		df = pd.DataFrame(table_data, columns=['Event Name', 'Bid Start Date_1'])
		df[['Bid Start Date', 'Bid Start Time']] = df['Bid Start Date_1'].str.split(expand=True)
		df.drop(columns=['Bid Start Date_1'], inplace=True)
		df.to_excel(output_path, index=False)
		print(f"Table data saved to {output_path} successfully.")
	
	def bulletin_board(self):
		options = Options()
		options.add_argument("--headless")  # Run selenium under headless mode
		prefs = {'download.default_directory': self.deep_bulletin_board_directory}
		options.add_experimental_option('prefs', prefs)
		chromedriver_path = r'C:\Users\pulki\.cache\selenium\chromedriver\win64\125.0.6422.76\chromedriver.exe'
		driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
		# driver = webdriver.Chrome(options=options)
		
		driver.get('https://www.mstcecommerce.com/auctionhome/ppa/index.jsp')
		table = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, 'table-hover')))
		
		rows = table.find_elements(By.TAG_NAME, 'tr')
		table_data = []
		for row in rows:
			cells = row.find_elements(By.TAG_NAME, 'td')
			row_data = []
			for cell in cells:
				row_data.append(cell.text)
			table_data.append(row_data)
		
		deep_bulletin_board.save_to_excel(self, table_data)


if __name__ == '__main__':
	bulletin_board_deep = deep_bulletin_board()
	deep_data = deep_portal()
	
	bulletin_board_deep.bulletin_board()
	deep_data.get_data()
	pass
