import requests
import os
import pandas as pd
from datetime import datetime, timedelta
from Data_Upload.data_mapping import uttarakhand_uksldc_generator, uttarakhand_uksldc_oa
from concurrent.futures import ThreadPoolExecutor, as_completed


class uttarakhand_generation_uksldc:
    def __init__(self):
        self.main_directory = r'C:\GNA\Data\Uttarakhand'
        self.file_directory = r'C:\GNA\Data\Uttarakhand\download_data'
        self.clear_or_create_directory(self.file_directory)
        self.error_log_file = os.path.join(self.main_directory, 'error_log.xlsx')
        pass

    def clear_or_create_directory(self, directory):
        if os.path.exists(directory):
            for file in os.listdir(directory):
                file_path_full = os.path.join(directory, file)
                if os.path.isfile(file_path_full):
                    os.remove(file_path_full)
        else:
            os.makedirs(directory)

    def get_generators_list(self):
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
            'Accept': '*/*', 'Accept-Language': 'en-US,en;q=0.9', 'Accept-Encoding': 'gzip, deflate, br, zstd',
            'Referer': 'https://uksldc.com/ViewReportSchedule/Index/GetNetSchedule', 'Connection': 'keep-alive'}
        pay_load = {'customertype': 1}
        url = 'https://uksldc.com/ViewReportSchedule/GetCustomerlist'

        response = requests.post(url, json=pay_load, headers=headers)

        if response.status_code == 200:
            data = response.json()
            print(data)
            with open(f'{self.main_directory}\generator.json', 'w') as f:
                f.write(data)
        else:
            print('Response failed with status:', response.status_code)

    def json_to_xlsx(self):
        json_files = []
        for file in os.listdir(self.main_directory):
            if file.endswith('.json'):
                json_files.append(os.path.join(self.main_directory, file))

        for json_file in json_files:
            output_file = os.path.join(self.main_directory, os.path.splitext(json_file)[0] + '.xlsx')
            df = pd.read_json(json_file)
            df.to_excel(output_file, index=False)
            print(f"File saved to '{output_file}'")

    def get_entity_id_and_name(self):
        input_file = os.path.join(self.main_directory, 'generator.xlsx')
        df = pd.read_excel(input_file)

        df['id'] = df['id'].astype(str).str.split('.').str[0]

        uttarakhand_uksldc_generator = pd.Series(df['id'].values, index=df['OA']).to_dict()
        print(uttarakhand_uksldc_generator)

        df['id'] = df['id'].astype(str).str.split('.').str[0]

        uttarakhand_uksldc_oa = pd.Series(df['id'].values, index=df['OA']).to_dict()
        print(uttarakhand_uksldc_oa)

    def download_generator_data_(self):
        error_files = []
        base_date = datetime.now()
        end_date = datetime(2024, 1, 1)

        for key, generator_id in uttarakhand_uksldc_generator.items():
            print(f'{key}:{generator_id}')

            # Reset base_date for each generator
            current_date = base_date

            while current_date > end_date:
                start_date = current_date.strftime('%Y-%m-%d')
                date_format = datetime.strptime(start_date, '%Y-%m-%d')
                day = date_format.day
                month = date_format.month
                year = date_format.year

                url = f'https://uksldc.com/ViewReportSchedule/GetNetSchedule?fromDate={month}%2F{day}%2F{year}&sldcrevision=6&formate=M%2Fd%2Fyyyy&type=2&entityid={generator_id}&customertype=1&_=1719475093968'
                response = requests.get(url)

                if response.status_code == 200:
                    file_path = os.path.join(self.file_directory, f'{key}_{start_date}.json')
                    with open(file_path, 'wb') as f:
                        f.write(response.content)
                    print(f'Downloaded data for {start_date}')
                else:
                    print(f'No data available for {start_date}')
                    error_files.append(f'{key}_{start_date}')

                # Decrement the current_date
                current_date -= timedelta(days=1)

            print(f'Finished downloading for generator: {generator_id}')

        if error_files:
            pd.DataFrame({'Error_Files': error_files}).to_excel(self.error_log_file, index=False)
            print(f'Error log saved to {self.error_log_file}')
        print("Download process completed.")

    def download_data_for_date(self, generator_info):
        key, generator_id, current_date = generator_info
        start_date = current_date.strftime('%Y-%m-%d')
        date_format = datetime.strptime(start_date, '%Y-%m-%d')
        day = date_format.day
        month = date_format.month
        year = date_format.year

        url = f'https://uksldc.com/ViewReportSchedule/GetNetSchedule?fromDate={month}%2F{day}%2F{year}&sldcrevision=6&formate=M%2Fd%2Fyyyy&type=2&entityid={generator_id}&customertype=1&_=1719475093968'
        response = requests.get(url)

        if response.status_code == 200:
            file_path = os.path.join(self.file_directory, f'{key}_{start_date}.json')
            with open(file_path, 'wb') as f:
                f.write(response.content)
            print(file_path)
            return None
        else:
            return f'{key}_{start_date}'

    def download_generator_data(self):
        error_files = []
        base_date = datetime.now()
        end_date = datetime(2024, 1, 1)
        tasks = []

        with ThreadPoolExecutor(max_workers=10) as executor:
            for key, generator_id in uttarakhand_uksldc_generator.items():
                current_date = base_date
                while current_date > end_date:
                    tasks.append((key, generator_id, current_date))
                    current_date -= timedelta(days=1)

            future_to_task = {executor.submit(self.download_data_for_date, task): task for task in tasks}
            for future in as_completed(future_to_task):
                error = future.result()
                if error:
                    error_files.append(error)

        if error_files:
            pd.DataFrame({'Error_Files': error_files}).to_excel(self.error_log_file, index=False)
            print(f'Error log saved to {self.error_log_file}')
        print("Download process completed.")

    def get_data(self):
        # uttarakhand_generation_uksldc.get_generators_list(self)
        # uttarakhand_generation_uksldc.json_to_xlsx(self)
        # uttarakhand_generation_uksldc.get_entity_id_and_name(self)
        uttarakhand_generation_uksldc.download_generator_data_(self)


if __name__ == '__main__':
    uk_generator_uksldc = uttarakhand_generation_uksldc()
    uk_generator_uksldc.get_data()
    pass
