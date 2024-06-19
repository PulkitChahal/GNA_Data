import requests
import os
import pandas as pd


class uttarakhand_generation:
	def __init__(self):
		file_directory = r'C:\GNA\Coding\Uttarakhand'
		if not os.path.exists(file_directory):
			os.mkdir(file_directory)
		self.file_directory = file_directory
		pass
	
	def get_generators_list(self):
		headers = {
			'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
			'Accept': '*/*',
			'Accept-Language': 'en-US,en;q=0.9',
			'Accept-Encoding': 'gzip, deflate, br, zstd',
			'Referer': 'https://uksldc.com/ViewReportSchedule/Index/GetNetSchedule',
			'Connection': 'keep-alive'
		}
		pay_load = {
			'customertype': 1
		}
		url = 'https://uksldc.com/ViewReportSchedule/GetCustomerlist'

		response = requests.post(url, json=pay_load, headers=headers)
		
		if response.status_code == 200:
			data = response.json()
			print(data)
			with open(f'{self.file_directory}\generator.json','w') as f:
				f.write(data)
		else:
			print('Response failed with status:', response.status_code)
		
	def json_to_xlsx(self):
		json_files = []
		for file in os.listdir(self.file_directory):
			if file.endswith('.json'):
				json_files.append(os.path.join(self.file_directory, file))
		
		for json_file in json_files:
			output_file = os.path.join(self.file_directory, os.path.splitext(json_file)[0]+'.xlsx')
			df = pd.read_json(json_file)
			df.to_excel(output_file, index=False)
			print(f"File saved to '{output_file}'")
	
	
	def find_duplicate_generator(self):
		pass
	
	def get_data(self):
		uttarakhand_generation().get_generators_list()
		uttarakhand_generation().json_to_xlsx()

if __name__ == '__main__':
	uk_generator = uttarakhand_generation()
	uk_generator.get_data()