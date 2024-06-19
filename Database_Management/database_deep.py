import sqlite3
import os
import pandas as pd


class deep_data_database():
	def __init__(self):
		self.file_directory = r'C:\GNA\Data Upload'
		self.db_name = os.path.join(self.file_directory, 'deep_portal_data.db')
		self.conn = sqlite3.connect(self.db_name)
		self.cur = self.conn.cursor()
		pass
	
	def create_table(self):
		self.cur.execute('''CREATE TABLE IF NOT EXISTS deep (
	        exchange_type varchar(20),
	        auction_no varchar(150),
	        auction_initiation_date DATE,
	        auction_initiation_time TIME,
	        auction_result_date DATE,
	        auction_result_time TIME,
	        buyer varchar(250),
	        requisition_no varchar(50),
	        delivery_start_date DATE,
	        delivery_end_date DATE,
	        delivery_start_time TIME,
	        delivery_end_time TIME,
	        type varchar(150),
	        buy_total_quantity_mw float,
	        booking_quantity_mw float,
	        allocated_quantity_mw float,
	        booking_accepted_price_kwh float,
	        accepted_price_kwh float,
	        PRIMARY KEY (
	            exchange_type,
	            auction_no,
	            requisition_no,
	            delivery_start_date,
	            type,
	            booking_quantity_mw,
	            allocated_quantity_mw,
	            booking_accepted_price_kwh,
	            accepted_price_kwh
	        )
	    )''')
	
	def insert_data(self):
		file_directory_for_data = os.path.join(self.file_directory, 'deep_portal.xlsx')
		df = pd.read_excel(file_directory_for_data)
		
		insert_query = '''INSERT OR IGNORE INTO deep (
	                exchange_type,
	                auction_no,
	                auction_initiation_date,
	                auction_initiation_time,
	                auction_result_date,
	                auction_result_time,
	                buyer,
	                requisition_no,
	                delivery_start_date,
	                delivery_end_date,
	                delivery_start_time,
	                delivery_end_time,
	                type,
	                buy_total_quantity_mw,
	                booking_quantity_mw,
	                allocated_quantity_mw,
	                booking_accepted_price_kwh,
	                accepted_price_kwh
	            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''
		
		# Iterate over DataFrame rows and insert them into the database
		for _, row in df.iterrows():
			# Print the row and types for debugging
			print(f"Row: {row}")
			print(f"Types: {[type(value) for value in row]}")

			self.cur.execute(insert_query, (
				str(row['Exchange Type']),
				str(row['Auction No.']),
				row['Auction Initiation Date'],
				row['Auction Initiation Time'],
				row['Auction Result Date'],
				row['Auction Result Time'],
				str(row['Buyer']),
				str(row['Requisition No']),
				row['Delivery Start Date'],
				row['Delivery End Date'],
				row['Delivery Start Time'],
				row['Delivery End Time'],
				str(row['Type']),
				float(row['Buy Total Quantity (in MW)']),
				float(row['Booking Quantity (in MW)']),
				float(row['Allocated Quantity (in MW)']),
				float(row['Booking Accepted Price (in Rs./kWh)']),
				float(row['Accepted Price (in Rs./kWh)'])
			))
		
		self.conn.commit()
		self.conn.close()
	
	def make_database(self):
		deep_data_database.create_table(self)
		deep_data_database.insert_data(self)


if __name__ == '__main__':
	deep_database = deep_data_database()
	deep_database.make_database()
	pass
