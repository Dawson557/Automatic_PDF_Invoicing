import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile


class Sheet_Handler:
	def __init__(self, month, year):
		self.month = month
		self.year = year
		self.rent_total = 0.0
		self.rent_tax_total = 0.0
		self.premier_total = 0.0
		self.suivi_total = 0.0
		self.commission_tax_total = 0.0

	def individual_report(self, rent, therapist, filename, data_package):
		sheet = None
		workbook = pd.read_excel(filename, sheet_name=None, dtype='object', index_col=0)
		#Check that sheet exists in workbook and initialize if not.
		if str(self.year) in workbook:
			sheet = workbook[str(self.year)]
		else:
			sheet = initialize_new_sheet()
		
		#Update sheet with this month's values
		if rent:
			rent_amount, taxes = data_package.get_rent(therapist)
			sheet[self.month]['Rent'] = rent_amount
			sheet[self.month]['Rent Taxes'] = taxes
			self.update_rent_totals(rent_amount, taxes)
		else:
			premier, suivi, taxes = data_package.get_excel_values(therapist)
			sheet[self.month]['Premier'] = premier
			sheet[self.month]['Suivi'] = suivi
			sheet[self.month]['Commission Taxes'] = taxes
			self.update_commission_totals(premier, suivi, taxes)
		# #add in formulae
		# sheet = add_formulae(sheet)
		#update the sheet in the workbook
		workbook[str(self.year)] = sheet

		#Write each sheet back to workbook
		writer_object = pd.ExcelWriter(filename, engine='xlsxwriter')
		for key in workbook:
			df = pd.DataFrame.from_dict(workbook[key])
			df = add_formulae(df)
			df.to_excel(writer_object, sheet_name=key)

		writer_object.save()	
		

	def update_commission_totals(self, premier, suivi, taxes):
		self.premier_total += premier
		self.suivi_total += suivi	
		self.commission_tax_total += taxes

	def update_rent_totals(self, rent, tax):
		self.rent_total += rent
		self.rent_tax_total += tax

	def summary_report(self, filename, rent):
		sheet = None
		workbook = pd.read_excel(filename, sheet_name=None, dtype='object', index_col=0)
		#Check that sheet exists in workbook and initialize if not.
		if str(self.year) in workbook:
			sheet = workbook[str(self.year)]
		else:
			sheet = initialize_summary_sheet()

		if rent:
			sheet[self.month]['Rent'] = self.rent_total
			sheet[self.month]['Rent Taxes'] = self.rent_tax_total
		else:
			sheet[self.month]['Premier'] = self.premier_total
			sheet[self.month]['Suivi'] = self.suivi_total
			sheet[self.month]['Commission Taxes'] = self.commission_tax_total

		sheet = add_formulae(sheet)
		#update the sheet in the workbook
		workbook[str(self.year)] = sheet

		#Write each sheet back to workbook
		writer_object = pd.ExcelWriter(filename, engine='xlsxwriter')
		for key in workbook:
			df = pd.DataFrame.from_dict(workbook[key])
			df.to_excel(writer_object, sheet_name=key)

		writer_object.save()	

def initialize_new_sheet():
	z = [0,0,0,0,0,"SUM"]
	template = {'Janvier':z, 'Fevrier':z, 'Mars':z, 'Avril':z,'Mai':z, 'Juin':z, 'Juillet':z, 
	'Aout':z, 'Septembre':z, 'Octobre':z, 'Novembre':z, 'Decembre':z,  'Cumulative Year':z} 
	df = pd.DataFrame(template, index=['Rent', 'Premier', 'Suivi', 'Rent Taxes', 'Commission Taxes', 'Total'])
	return df

#This is currently the same as initialize_new_sheet but I'm keeping it seperate because I'll likely be changing it
def initialize_summary_sheet():
	z = [0,0,0,0,0,"SUM"]
	template = {'Janvier':z, 'Fevrier':z, 'Mars':z, 'Avril':z,'Mai':z, 'Juin':z, 'Juillet':z, 
	'Aout':z, 'Septembre':z, 'Octobre':z, 'Novembre':z, 'Decembre':z,  'Cumulative Year':z} 
	df = pd.DataFrame(template, index=['Rent', 'Premier', 'Suivi', 'Rent Taxes', 'Commission Taxes', 'Total'])
	return df

#Loading in the excel sheet it recognizes formulae as the number values shown, this converts them back to formulae for saving
def add_formulae(sheet):
	column = ['=SUM(B2:M2)', '=SUM(B3:M3)','=SUM(B4:M4)','=SUM(B5:M5)', '=SUM(B6:M6)', '=SUM(N2:N6)',  ]
	row = ['=SUM(B2:B6)','=SUM(C2:C6)', '=SUM(D2:D6)','=SUM(E2:E6)','=SUM(F2:F6)','=SUM(G2:G6)',
	'=SUM(H2:H6)','=SUM(I2:I6)','=SUM(J2:J6)','=SUM(K2:K6)','=SUM(L2:L6)','=SUM(M2:M6)', '=SUM(N2:N6)']
	index = -1
	for name, val in sheet.iteritems():
		index += 1
		if (name != 'Cumulative Year'):
			sheet[name][-1] = row[index]
		else:
			sheet[name] = column
	return sheet

def create_workbook(filename, year):
	df = initialize_new_sheet()
	df.to_excel(filename, str(year), index=True)

