import xlwings


class ExcelFileEdit():
	def __init__(self):
		pass
	
	def convert_xls_to_xlsx(self, file_path, output_path):
		self.file_path = file_path
		self.output_path = output_path
		try:
			app = xlwings.App(visible=False)
			workbook = app.books.open(file_path)
			workbook.save(output_path)
			workbook.close()
			app.quit()
			print(f"Conversion completed. File saved: {output_path}")
		except Exception as e:  # Catch all exceptions
			print(f"Error converting {file_path}: {e}")


if __name__ == "__main__":
	file_edit = ExcelFileEdit()
	file_path = r"C:\Users\pulki\Downloads\T-GNA Approved Transaction Report_May_24.xls"
	output_path = r"C:\Users\pulki\Downloads\T-GNA Approved Transaction Report_May_24.xlsx"
	file_edit.convert_xls_to_xlsx(file_path, output_path)
