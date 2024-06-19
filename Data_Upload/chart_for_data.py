import pandas as pd
import matplotlib.pyplot as plt
import os


class deep_tam_chart:
	def __init__(self):
		self.file_directory = r'C:\GNA\Data Upload\Haryana'
		if not os.path.exists(self.file_directory):
			os.makedirs(self.file_directory)
		pass
	
	def chart_for_deep_tam(self):
		file_path = r"C:\GNA\Data Upload\haryana2.xlsx"
		df = pd.read_excel(file_path)
		
		years = df['Delivery Year'].unique()
		# Sort by Delivery Month (assuming 'Delivery Month' is in a datetime format)
		df['Delivery Month'] = pd.to_datetime(df['Delivery Month'], format='%b')
		df = df.sort_values(by='Delivery Month')
		
		years = df['Delivery Year'].unique()
		
		for year in years:
			yearly_data = df[df['Delivery Year'] == year]
			
			fig, ax1 = plt.subplots(figsize=(16, 6))
			
			# Plot Minimum Accepted Price
			color1 = 'tab:blue'
			color2 = 'tab:red'
			ax1.set_xlabel('Delivery Month')
			ax1.set_ylabel('Buy Total Quantity (in MW)', color=color1)
			ax1.bar(yearly_data['Delivery Month'].dt.strftime('%b'),
			        yearly_data['Buy Total Quantity (in MW)'],
			        color=color1, label='Buy Total Quantity')
			ax1.bar(yearly_data['Delivery Month'].dt.strftime('%b'),
			        yearly_data['Allocated Quantity (in MW)'],
			        color=color2, label='Allocated Quantity')
			ax1.tick_params(axis='y', labelcolor=color1)
			
			# Grid for the primary axis (ax1)
			ax1.grid(True)
			
			# Adding legend
			lines1, labels1 = ax1.get_legend_handles_labels()
			# lines2, labels2 = ax1.get_legend_handles_labels()
			ax1.legend(lines1, labels1, loc='best')
			
			# Add title and save plot
			plt.title(f'Buy Quantity for Year {year}')
			fig.tight_layout()
			plt.show()
			plt.savefig(os.path.join(self.file_directory, f'Buy_Quantity_for_Year_{year}.png'))
			plt.close(fig)  # Close the figure to free up memory
	
	def chart_for_deep_tam_final(self):
		file_path = r"C:\GNA\Data Upload\haryana2.xlsx"
		df = pd.read_excel(file_path)
		
		# Sort by Delivery Month (assuming 'Delivery Month' is in a datetime format)
		df['Delivery Month'] = pd.to_datetime(df['Delivery Month'], format='%b')
		df = df.sort_values(by='Delivery Month')
		
		years = df['Delivery Year'].unique()
		
		for year in years:
			yearly_data = df[df['Delivery Year'] == year]
			
			fig, ax1 = plt.subplots(figsize=(16, 6))
			
			# Plot Buy Total Quantity and Allocated Quantity as clustered bars
			color1 = 'tab:blue'
			color2 = 'tab:red'
			ax1.set_xlabel('Delivery Month')
			ax1.set_ylabel('Accepted Price (in Rs./kWh)')
			
			# Width of each bar cluster
			bar_width = 0.4
			
			# Calculate positions for bars
			positions1 = range(len(yearly_data))
			positions2 = [pos + bar_width for pos in positions1]
			
			# Plot Buy Total Quantity
			ax1.bar(positions1,
			        yearly_data['Minimum Accepted Price (in Rs./kWh)'],
			        color=color1, width=bar_width, label='Minimum Accepted Price')
			
			# Plot Allocated Quantity
			ax1.bar(positions2,
			        yearly_data['Maximum Accepted Price (in Rs./kWh)'],
			        color=color2, width=bar_width, label='Maximum Accepted Price')
			
			# Set x-axis ticks and labels
			ax1.set_xticks(positions1)
			ax1.set_xticklabels(yearly_data['Delivery Month'].dt.strftime('%b'))
			
			# Grid for the primary axis (ax1)
			ax1.grid(True)
			# Adding legend
			ax1.legend(loc='best')
			
			# Add title and save plot
			plt.title(f'Accepted Price for Year {year}')
			fig.tight_layout()
			plt.savefig(os.path.join(self.file_directory, f'Accepted_Price_for_Year_{year}.png'))
			# plt.show()
			plt.close(fig)


if __name__ == '__main__':
	deep_chart = deep_tam_chart()
	deep_chart.chart_for_deep_tam_final()
	pass
