import re
import requests
import os
import time
import tabula
import pandas as pd
from datetime import datetime, timedelta
from PyPDF2 import PdfReader
import json

MONTHS = {
    'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
    'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
}

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

main_directory = r'C:\GNA\Data\Reverse Auction'

file_directory = r"C:\GNA\Data\Reverse Auction\PXIL Reverse Auction"

result_report_file = r"C:\GNA\Data\Reverse Auction\RAResultReport.xlsx"

output_directory = r'C:\GNA\Data\Reverse Auction\PXIL Reverse Auction xlsx Files'
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

output_directory1 = r'C:\GNA\Data\Reverse Auction\PXIL Reverse Auction xlsx Files1'
if not os.path.exists(output_directory1):
    os.makedirs(output_directory1)

error_log_file = r'C:\GNA\Data\IEX Reverse Auction\error_log_pxil.xlsx'

final_directory = r'C:\GNA\Data Upload'
if not os.path.exists(final_directory):
    os.makedirs(final_directory)


class pxil_reverse_auction():
    def __init__(self):
        pass

    def links_for_data(self):
        base_date = datetime.now()
        base_end_date = datetime.now() - timedelta(30)
        base_date = base_date.strftime('%Y-%m-%d')
        base_end_date = base_end_date.strftime('%Y-%m-%d')
        print(base_date)
        print(base_end_date)
        pay_load = {
            "from_date": base_end_date,
            "to_date": base_date,
            "result_ra_type": "ALL"
        }
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Accept-Language": "en-US,en;q=0.9",
            "Accept-EnData": "gzip, deflate, br, zstd",
            "Connection": "keep-alive",
        }
        url = 'https://powerexindia.in/get_result_data/'

        data = requests.post(url, data=pay_load, headers=headers)
        data = data.json()
        file_name = os.path.join(main_directory, 'links.josn')
        with open(file_name, 'w') as json_file:
            json.dump(data, json_file, indent=4)

        pxil_reverse_auction.json_to_xlsx(self)

    def json_to_xlsx(self):
        file_name = os.path.join(main_directory, 'links.josn')
        df = pd.read_json(file_name)
        df.to_excel(result_report_file, index=False)

        pxil_reverse_auction.edit_result_report(self)

    def edit_result_report(self):
        # file = r"C:\GNA\Data\Reverse Auction\RAResultReport.xlsx"
        df = pd.read_excel(result_report_file)
        df = df['data'].astype(str).str.split(r"':", expand=True)
        df = df.drop(columns=[0, 1, 6], axis=0)
        df = df.replace("'", '', regex=True)
        df[2] = df[2].astype(str).str.split(',').str[0]
        df[3] = df[3].astype(str).str.split(',').str[0]
        df[4] = df[4].astype(str).str.split(',').str[0]
        df[5] = df[5].astype(str).str.split(',').str[0]
        df = df[4].astype(str).str.split(' ', expand=True)
        df.to_excel(result_report_file, index=False)

    def download_file(self, url, destination):
        response = requests.get(url)
        with open(destination, 'wb') as f:
            f.write(response.content)

    def generate_url_links(self):
        file_path = r'C:\GNA\Data\Reverse Auction\PXIL Reverse Auction'
        if not os.path.exists(file_path):
            os.makedirs(file_path)

        excel_file = r'C:\GNA\Data\Reverse Auction\RAResultReport.xlsx'
        df = pd.read_excel(excel_file)
        date = df[1].tolist()
        month = df[2].tolist()
        year = df[3].tolist()
        auction_id = df[7].tolist()

        for date, month, year, auction_id in zip(date, month, year, auction_id):
            url = f'https://www.powerexindia.in/media/downloads/{date}%20{month}%20{year}%20Reverse%20Auction%20Result%20{auction_id}'
            response = requests.get(url)

            if response.status_code == 200:
                with open(f'{file_path}/{auction_id}.pdf', 'wb') as f:
                    f.write(response.content)
                print(f"File {auction_id}.pdf downloaded successfully.")

    def extract_text_from_pdf(self, file_path):
        with open(file_path, 'rb') as f:
            reader = PdfReader(f)
            auction_initiation_date = []
            auction_result_date = []
            accepted_quantity_text = []

            for page in reader.pages:
                text = page.extract_text()

                initiation_on = text.find('nitiated on ')
                initiation_end = text.find('Buyer')
                done_on = text.find('one on ')
                done_on_end = text.find('Total count of Selle')
                accepted_quantity_start = text.find('Acceptance/Rejection')
                accepted_quantity_end = text.find('This report')

                if initiation_on != -1 and initiation_end != -1:
                    initiation_date = text[initiation_on + len('nitiated on '):initiation_end]
                    auction_initiation_date.append(initiation_date.strip())
                elif initiation_on != -1:
                    initiation_date = text[initiation_on + len('nitiated on '):]
                    auction_initiation_date.append(initiation_date.strip())

                if done_on != -1 and done_on_end != -1:
                    result_date = text[done_on + len('one on '):done_on_end]
                    auction_result_date.append(result_date.strip())
                elif done_on != -1:
                    result_date = text[done_on + len('one on '):]
                    auction_result_date.append(result_date.strip())

                if accepted_quantity_start != -1 and accepted_quantity_end != -1:
                    accepted_quantity_section = text[accepted_quantity_start + len(
                        'Acceptance/Rejection'):accepted_quantity_end]
                    accepted_quantity_text.append(accepted_quantity_section.strip())

            print(auction_initiation_date)
            print(auction_result_date)
            print(accepted_quantity_text)
            return auction_initiation_date, auction_result_date, accepted_quantity_text

    def pdf_to_xlsx(self):
        pdf_files = []
        for pdf_file in os.listdir(file_directory):
            if pdf_file.endswith('.pdf'):
                pdf_files.append(pdf_file)

        for pdf_file in pdf_files:
            file_path = os.path.join(file_directory, pdf_file)
            file_name = pdf_file.split('.pdf')[0]  # Split file name by '.pdf'
            output_path = os.path.join(output_directory, file_name + '_except_last_file.xlsx')
            last_table_output_path = os.path.join(output_directory, file_name + '_last.xlsx')
            last_table = None

            auction_initiation_date, auction_result_date, accepted_quantity_text = pxil_reverse_auction.extract_text_from_pdf(
                self, file_path)
            tables = tabula.read_pdf(file_path, pages='all', multiple_tables=True)
            if tables:
                merged_df = pd.DataFrame()
                for i, table in enumerate(tables):
                    if 'Start Date' in table.columns:
                        last_table = pd.DataFrame(table)
                        last_table.insert(0, 'Auction No.', file_name)
                    else:
                        df = pd.DataFrame(table)
                        if i == 0 and not df.empty and len(df.columns) > 1:
                            df.loc[-1] = df.columns
                            df.index = df.index + 1
                            df = df.sort_index()
                            df.columns = range(len(df.columns))
                            merged_df = pd.concat([merged_df, df], ignore_index=True)
                        if i == 1 and not df.empty:
                            df.loc[-1] = df.columns
                            df.index = df.index + 1
                            df = df.sort_index()
                            df.columns = range(len(df.columns))
                            merged_df = pd.concat([merged_df, df], ignore_index=True)

                if last_table is not None and 'Start Date' in last_table.columns:
                    last_table_output_path = os.path.join(output_directory, file_name + '_last.xlsx')
                    last_table.to_excel(last_table_output_path, index=False)
                    print(
                        f"Successfully extracted and saved the last table from {pdf_file} to {last_table_output_path}")
                else:
                    print(
                        f"The last table from {pdf_file} does not contain 'Start Date'. Skipping saving the last table.")

                concatenated_df = pd.concat([merged_df], ignore_index=True)
                transposed_df = concatenated_df.transpose()
                transposed_df.columns = transposed_df.iloc[0]
                transposed_df = transposed_df.iloc[1:]
                transposed_df.columns = transposed_df.columns.str.replace(r'\.(\d+)$', '')
                # Rename columns
                transposed_df.rename(
                    columns={'Delivery Dates': 'Delivery Start Date',
                             '(Start Date to End Date)': 'Delivery End Date', },
                    inplace=True)
                transposed_df.insert(0, 'Auction No.', file_name)

                if len(auction_initiation_date) > 0:
                    transposed_df.insert(1, 'Auction Initiation Date', auction_initiation_date[0])
                else:
                    transposed_df.insert(1, 'Auction Initiation Date', '')
                if len(auction_result_date) > 0:
                    transposed_df.insert(2, 'Auction Result Date', auction_result_date[0])
                else:
                    transposed_df.insert(2, 'Auction Result Date', '')
                if len(accepted_quantity_text) > 0:
                    transposed_df.insert(3, 'Accepted Quantity', accepted_quantity_text[0])
                else:
                    transposed_df.insert(3, 'Accepted Quantity', '')

                transposed_df.to_excel(output_path, index=False)
                print(f"Successfully extracted tables from {pdf_file} and saved to {output_path}")

    def edit_last_table(self):
        xlsx_files = []
        for xlsx_file in os.listdir(output_directory):
            if xlsx_file.endswith('_last.xlsx'):
                xlsx_files.append(xlsx_file)

        for xlsx_file in xlsx_files:
            file_path = os.path.join(output_directory, xlsx_file)
            output_path = os.path.join(output_directory, os.path.splitext(xlsx_file)[0] + '.xlsx')
            print(xlsx_file)

            df = pd.read_excel(file_path)
            df['Start Date'] = pd.to_datetime(df['Start Date'], format='%d-%m-%Y')
            df['End Date'] = pd.to_datetime(df['End Date'], format='%d-%m-%Y')

            df.to_excel(output_path, index=False)
            print(f'File Edited Successfully {xlsx_file}')

    def edit_except_last_files(self):
        xlsx_files = []
        for xlsx_file in os.listdir(output_directory):
            if xlsx_file.endswith('_except_last_file.xlsx'):
                xlsx_files.append(xlsx_file)

        for xlsx_file in xlsx_files:
            file_path = os.path.join(output_directory, xlsx_file)
            output_path = os.path.join(output_directory, os.path.splitext(xlsx_file)[0] + '.xlsx')

            df = pd.read_excel(file_path)
            if 'Unnamed: 0' and 'Unnamed: 6' in df.columns:
                df['Buyer_2'] = df['Unnamed: 0'].astype(str) + ' ' + df['Unnamed: 6'].astype(str)
                df['Buyer'].fillna(df['Buyer_2'], inplace=True)
                df.drop(['Unnamed: 0', 'Unnamed: 6', 'Buyer_2'], axis=1, inplace=True)

            df['Delivery Start Date'] = df['Delivery Start Date'].astype(str).str.split(' ').str[0]
            df[['Delivery Start Date', 'Delivery End Date']] = df[['Delivery Start Date', 'Delivery End Date']].astype(
                str).replace('/', '-', regex=True)

            df['Delivery Period (From-To, in Hrs)'] = df['Delivery Period (From-To, in Hrs)'].astype(str)
            new_rows = []
            for index, row in df.iterrows():
                if '&' in row['Delivery Period (From-To, in Hrs)']:
                    times = row['Delivery Period (From-To, in Hrs)'].split('&')
                    df.at[index, 'Delivery Period (From-To, in Hrs)'] = times[0].strip()
                    new_row = {'Delivery Period (From-To, in Hrs)': times[1].strip()}
                    for col in df.columns:
                        if col not in ['Delivery Period (From-To, in Hrs)']:
                            new_row[col] = row[col]
                    new_rows.append(new_row)
            df = df._append(new_rows, ignore_index=True)

            df['Delivery Period (From-To, in Hrs)'] = df['Delivery Period (From-To, in Hrs)'].astype(str).replace('to',
                                                                                                                  '-',
                                                                                                                  regex=True)
            df[['Delivery Start Time', 'Delivery End Time']] = df['Delivery Period (From-To, in Hrs)'].str.split('-',
                                                                                                                 expand=True)

            df[['Delivery Start Time', 'Delivery End Time']] = df[['Delivery Start Time', 'Delivery End Time']].astype(
                str).replace('Hrs.', '', regex=True)
            df[['Delivery Start Time', 'Delivery End Time']] = df[['Delivery Start Time', 'Delivery End Time']].astype(
                str).replace('Hrs', '', regex=True)
            df[['Delivery Start Time', 'Delivery End Time']] = df[['Delivery Start Time', 'Delivery End Time']].astype(
                str).replace('Hr', '', regex=True)

            df['Auction Initiation Date'] = df['Auction Initiation Date'].apply(
                lambda x: re.sub(r'(?:st|nd|rd|th|,)', '', str(x)))
            df['Auction Result Date'] = df['Auction Result Date'].apply(
                lambda x: re.sub(r'(?:st|nd|rd|th|,)', '', str(x)))

            for month_name, month_num in month_replacements.items():
                df['Auction Initiation Date'] = df['Auction Initiation Date'].str.replace(month_name, month_num)
                df['Auction Result Date'] = df['Auction Result Date'].str.replace(month_name, month_num)

            df['Auction Initiation Date'] = df['Auction Initiation Date'].str.replace(' ', '-')
            df['Auction Result Date'] = df['Auction Result Date'].str.replace(' ', '-')

            df['Auction Initiation Date'] = pd.to_datetime(df['Auction Initiation Date'], format='%d-%m-%Y')
            df['Auction Result Date'] = pd.to_datetime(df['Auction Result Date'], format='%d-%m-%Y')
            df['Delivery Start Date'] = pd.to_datetime(df['Delivery Start Date'], format='%d-%m-%Y')
            df['Delivery End Date'] = pd.to_datetime(df['Delivery End Date'], format='%d-%m-%Y')

            df['Accepted Quantity'] = df['Accepted Quantity'].astype(str).replace(' ', '', regex=True)
            df['Accepted Quantity'] = df['Accepted Quantity'].astype(str).replace('\n', '', regex=True)

            df['Accepted Quantity'] = df['Accepted Quantity'].astype(str).replace(
                'ReverseAuctionResultsnotAvailableasnosellerparticipatedAcceptedPrice', '', regex=True)
            df['Accepted Quantity'] = df['Accepted Quantity'].astype(str).replace(
                'byBuyerDetailsAcceptedquantity', '', regex=True)
            df['Accepted Quantity'] = df['Accepted Quantity'].astype(str).replace(
                'ReverseAuctionResultsnotAcceptedbytheBuyerforAllocationAcceptedPrice', '', regex=True)

            df.drop(['Delivery Period (From-To, in Hrs)'], axis=1, inplace=True)

            df.to_excel(output_path, index=False)
            print(f'File Edited Successfully {xlsx_file}')

    def merge_last_files(self):
        xlsx_files = []
        for xlsx_file in os.listdir(output_directory):
            if xlsx_file.endswith('_last.xlsx'):
                xlsx_files.append(xlsx_file)
        merged_df = pd.DataFrame()
        for xlsx_file in xlsx_files:
            file_path = os.path.join(output_directory, xlsx_file)
            df = pd.read_excel(file_path)
            merged_df = pd.concat([merged_df, df], ignore_index=True)

        merged_file_path = os.path.join(main_directory, 'merged_last_pxil.xlsx')
        merged_df.to_excel(merged_file_path, index=False)
        print(f"Merged file Saved to '{merged_file_path}'")

    def merge_except_last_files(self):
        xlsx_files = []
        for xlsx_file in os.listdir(output_directory):
            if xlsx_file.endswith('_except_last_file.xlsx'):
                xlsx_files.append(xlsx_file)
        merged_df = pd.DataFrame()
        for xlsx_file in xlsx_files:
            file_path = os.path.join(output_directory, xlsx_file)
            df = pd.read_excel(file_path)
            merged_df = pd.concat([merged_df, df], ignore_index=True)

        merged_df['Accepted Quantity'] = merged_df['Accepted Quantity'].astype(str).replace(
            'ReverseAuctionResultsnotAvailableasnosellerparticipatedAcceptedPrice', '', regex=True)
        merged_df['Accepted Quantity'] = merged_df['Accepted Quantity'].astype(str).replace(
            'byBuyerDetailsAcceptedquantity', '', regex=True)
        merged_df['Accepted Quantity'] = merged_df['Accepted Quantity'].astype(str).replace(
            'ReverseAuctionResultsnotAcceptedbytheBuyerforAllocationAcceptedPrice', '', regex=True)

        new_rows = []
        for index, row in merged_df.iterrows():
            if '(inMW),AcceptedPrice(inRs./kWh)' in row['Accepted Quantity']:
                times = row['Accepted Quantity'].split('(inMW),AcceptedPrice(inRs./kWh)')
                merged_df.at[index, 'Accepted Quantity'] = times[0].strip()
                new_row = {'Accepted Quantity': times[1].strip()}
                for col in merged_df.columns:
                    if col not in ['Accepted Quantity']:
                        new_row[col] = row[col]
                new_rows.append(new_row)
        merged_df = merged_df._append(new_rows, ignore_index=True)

        merged_df = merged_df.dropna(subset=['Accepted Quantity'])
        merged_df = merged_df[merged_df['Accepted Quantity'] != '']

        merged_df['Accepted Quantity'] = merged_df['Accepted Quantity'].astype(str).replace(r'inMw', '', regex=True)
        merged_df['Accepted Quantity'] = merged_df['Accepted Quantity'].astype(str).replace(r'KWh', '', regex=True)
        merged_df['Accepted Quantity'] = merged_df['Accepted Quantity'].astype(str).replace(r'/', '', regex=True)
        merged_df['Accepted Quantity'] = merged_df['Accepted Quantity'].astype(str).replace(r'in', '', regex=True)
        merged_df['Accepted Quantity'] = merged_df['Accepted Quantity'].astype(str).replace(r'\(', '', regex=True)
        merged_df['Accepted Quantity'] = merged_df['Accepted Quantity'].astype(str).replace(r'\)', '', regex=True)
        merged_df['Accepted Quantity'] = merged_df['Accepted Quantity'].astype(str).replace('MW', '', regex=True)
        merged_df['Accepted Quantity'] = merged_df['Accepted Quantity'].astype(str).replace('Rs.', '', regex=True)
        merged_df['Accepted Quantity'] = merged_df['Accepted Quantity'].astype(str).replace('Acceptedquantity', '',
                                                                                            regex=True)

        merged_df[['Allocated Quantity (in MW)', 'Accepted Price (in Rs./kWh)']] = merged_df[
            'Accepted Quantity'].astype(
            str).str.split('@', expand=True)
        merged_df.drop(['Accepted Quantity'], axis=1, inplace=True)

        column_to_move = [
            'Auction No.',
            'Auction Initiation Date',
            'Auction Result Date',
            'Buyer',
            'Delivery Start Date',
            'Delivery End Date',
            'Delivery Start Time',
            'Delivery End Time',
            'Buy - Total Quantity (in MW)',
            'Buy - Minimum Quantity (in MW)',
            'Delivery Point',
            'Total count of Delivery Days',
            'Exclusion Days',
            'Allocated Quantity (in MW)',
            'Accepted Price (in Rs./kWh)',
        ]
        remaining_column_order = [col for col in merged_df.columns if col not in column_to_move]
        column_order = column_to_move + remaining_column_order
        merged_df = merged_df[column_order]

        merged_df['Allocated Quantity (in MW)'] = pd.to_numeric(merged_df['Allocated Quantity (in MW)'],
                                                                errors='coerce')
        merged_df['Allocated Quantity (in MW)'] = merged_df['Allocated Quantity (in MW)'].fillna('')
        merged_df['Accepted Price (in Rs./kWh)'] = pd.to_numeric(merged_df['Accepted Price (in Rs./kWh)'],
                                                                 errors='coerce')
        merged_df['Accepted Price (in Rs./kWh)'] = merged_df['Accepted Price (in Rs./kWh)'].fillna('')

        merged_df['Auction No.'] = merged_df['Auction No.'].astype(str).replace('-Anydy', 'Anyday', regex=True)
        merged_df['Auction No.'] = merged_df['Auction No.'].astype(str).replace('-Anyday', 'Anyday', regex=True)

        merged_file_path = os.path.join(main_directory, 'merged_except_last_file_pxil.xlsx')
        merged_df.to_excel(merged_file_path, index=False)
        print(f"Merged File Saved to '{merged_file_path}'")

    def merge_final_files(self):
        file_1 = r"C:\GNA\Data\Reverse Auction\merged_except_last_file_pxil.xlsx"
        file_2 = r"C:\GNA\Data\Reverse Auction\merged_last_pxil.xlsx"
        df1 = pd.read_excel(file_1)
        df2 = pd.read_excel(file_2)

        # Merge DataFrames based on 'Auction No.' column
        merged_df = pd.merge(df1, df2, on="Auction No.", how="left")
        column_to_drop = ['Qty',
                          'Min Qty', ]
        merged_df = merged_df.drop(columns=column_to_drop)

        merged_df.insert(0, 'Exchange Type', 'PXIL')

        merged_df['Start Date'] = pd.to_datetime(merged_df['Start Date'], format='%d-%m-%Y')
        merged_df['End Date'] = pd.to_datetime(merged_df['End Date'], format='%d-%m-%Y')
        merged_df['Allocated Quantity (in MW)'] = merged_df['Allocated Quantity (in MW)'].astype(float)
        merged_df['Accepted Price (in Rs./kWh)'] = merged_df['Accepted Price (in Rs./kWh)'].astype(float)

        merged_df['End Date'].fillna(merged_df['Delivery End Date'], inplace=True)
        merged_df['Start Date'].fillna(merged_df['Delivery Start Date'], inplace=True)
        merged_df['From Slot'].fillna(merged_df['Delivery Start Time'], inplace=True)
        merged_df['To Slot'].fillna(merged_df['Delivery End Time'], inplace=True)

        merged_df['From Slot'] = merged_df['From Slot'].astype(str).replace(' ', '', regex=True)
        merged_df['To Slot'] = merged_df['To Slot'].astype(str).replace(' ', '', regex=True)

        merged_df.drop(['Delivery Start Date', 'Delivery End Date', 'Delivery Start Time', 'Delivery End Time'], axis=1,
                       inplace=True)

        merged_df.rename(columns={
            'From Slot': 'Delivery Start Time',
            'To Slot': 'Delivery End Time',
            'Start Date': 'Delivery Start Date',
            'End Date': 'Delivery End Date'
        }, inplace=True)

        merged_df['Auction Result Date'] = merged_df['Auction Result Date'].dt.date
        merged_df['Auction Initiation Date'] = merged_df['Auction Initiation Date'].dt.date
        merged_df['Delivery Start Date'] = merged_df['Delivery Start Date'].dt.date
        merged_df['Delivery End Date'] = merged_df['Delivery End Date'].dt.date

        # merged_df['Delivery Start Time'] = merged_df['Delivery Start Time'].astype(str).replace(' ', '', regex=True)
        # merged_df['Delivery Start Time'] = pd.to_datetime(merged_df['Delivery Start Time'], format='%H:%M').dt.time
        # merged_df['Delivery End Time'] = merged_df['Delivery End Time'].astype(str).replace(' ', '', regex=True)
        # merged_df['Delivery End Time'] = merged_df['Delivery End Time'].replace('24:00', '24:00:00')
        # merged_df['Delivery End Time'] = pd.to_datetime(merged_df['Delivery End Time'], format='%H:%M').dt.time

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
            'Buy - Total Quantity (in MW)',
            'Buy - Minimum Quantity (in MW)',
            'Delivery Point',
            'Total count of Delivery Days',
            'Exclusion Days',
            'Allocated Quantity (in MW)',
            'Accepted Price (in Rs./kWh)',
            'Total count of Sellers who participated in the auction'
        ]
        # remaining_column_order = [col for col in merged_df.columns if col not in column_to_move]
        # column_order = column_to_move + remaining_column_order
        # merged_df = merged_df[column_order]
        merged_df = merged_df[column_to_move]

        merged_df.rename(columns={
            'Buy - Minimum Quantity (in MW)': 'Buy Minimum Quantity (in MW)',
            'Buy - Total Quantity (in MW)': 'Buy Total Quantity (in MW)',
            'Exclusion Days': 'Exclusion Dates'
        }, inplace=True)

        merged_df = merged_df.sort_values(by='Auction Result Date', ascending=False)
        output_path = os.path.join(final_directory, 'final_pxil.xlsx')
        merged_df.to_excel(output_path, index=False)
        print(f'Final File Saved at {output_path}')

    def get_data(self):
        pxil_reverse_auction.links_for_data(self)
        pxil_reverse_auction.generate_url_links(self)
        pxil_reverse_auction.pdf_to_xlsx(self)
        pxil_reverse_auction.edit_except_last_files(self)
        pxil_reverse_auction.merge_last_files(self)
        pxil_reverse_auction.merge_except_last_files(self)
        pxil_reverse_auction.merge_final_files(self)
        pass


if __name__ == '__main__':
    tam_pxil = pxil_reverse_auction()
    tam_pxil.get_data()
    pass
