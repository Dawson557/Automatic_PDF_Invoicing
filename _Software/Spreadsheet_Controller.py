import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import openpyxl as pyxl


class Sheet_Handler:
	def __init__(self, month, year):
		self.month = month
		self.year = year
		self.rent_total = 0.0
		self.premier_total = 0.0
		self.suivi_total = 0.0
		self.taxes_total = 0.0

	def individual_report(self, rent, therapist, filename, data_package):
		sheet = None
		workbook = pd.read_excel(filename, sheet_name=None)
		#Check that sheet exists in workbook and initialize if not.
		if str(self.year) in workbook:
			sheet = workbook[str(self.year)]
		else:
			sheet = initialize_new_sheet()
		
		#Update sheet with this month's values
		if rent:
			rent_amount = data_package.get_rent(therapist)
			sheet[self.month]['Rent'] = rent_amount
			self.rent_total += rent_amount
		else:
			premier, suivi, taxes = data_package.get_excel_values(therapist)
			sheet[self.month]['Premier'] = premier
			sheet[self.month]['Suivi'] = suivi
			sheet[self.month]['Taxes'] = taxes
			self.update_totals(premier, suivi, taxes)

		#update the sheet in the workbook
		workbook[str(self.year)] = sheet

		#Write each sheet back to workbook
		writer_object = pd.ExcelWriter(filename, engine='xlsxwriter')
		for key in workbook:
			df = pd.DataFrame.from_dict(workbook[key])
			df.to_excel(writer_object, sheet_name=key)

		writer_object.save()	
		self.update_totals(premier, suivi, taxes)

	def update_totals(self, premier, suivi, taxes):
		self.premier_total += premier
		self.suivi_total += suivi	
		self.taxes_total += taxes

	def summary_report(self, filename, rent):
		sheet = None
		workbook = pd.read_excel(filename, sheet_name=None)
		#Check that sheet exists in workbook and initialize if not.
		if str(self.year) in workbook:
			sheet = workbook[str(self.year)]
		else:
			sheet = initialize_summary_sheet()

		if rent:
			sheet[self.month]['Rent'] = self.rent_total
		else:
			sheet[self.month]['Premier'] = self.premier_total
			sheet[self.month]['Suivi'] = self.suivi_total
			sheet[self.month]['Taxes'] = self.taxes_total

		#update the sheet in the workbook
		workbook[str(self.year)] = sheet

		#Write each sheet back to workbook
		writer_object = pd.ExcelWriter(filename, engine='xlsxwriter')
		for key in workbook:
			df = pd.DataFrame.from_dict(workbook[key])
			df.to_excel(writer_object, sheet_name=key)

		writer_object.save()	



def initialize_individual_report(filename, year):
	cumulative = ['=SUM(B2:M2)', '=SUM(B3:M3)','=SUM(B4:M4)','=SUM(B5:M5)','=SUM(N2:N5)' ]
	template = {'Janvier':[0,0,0,0,'=SUM(B2:B5)'], 'Fevrier':[0,0,0,0,'=SUM(C2:C5)'], 'Mars':[0,0,0,0,'=SUM(D2:D5)'], 'Avril':[0,0,0,0,'=SUM(E2:E5)'],
	 'Mai':[0,0,0,0,'=SUM(F2:F5)'], 'Juin':[0,0,0,0,'=SUM(G2:G5)'], 'Juillet':[0,0,0,0,'=SUM(H2:H5)'], 'Aout':[0,0,0,0,'=SUM(I2:I5)'], 
	 'Septembre':[0,0,0,0,'=SUM(J2:J5)'], 'Octobre':[0,0,0,0,'=SUM(K2:K5)'], 'Novembre':[0,0,0,0,'=SUM(L2:L5)'], 'Decembre':[0,0,0,0,'=SUM(M2:M5)'],
	  'Cumulative Year':cumulative} 
	df = pd.DataFrame(template, index=['Rent', 'Premier', 'Suivi', 'Taxes', 'Total'])

	writer_object = pd.ExcelWriter(filename, engine='xlsxwriter')
	df.to_excel(writer_object, sheet_name=str(year))
	writer_object.save()

def initialize_new_sheet():
	cumulative = ['=SUM(B2:M2)', '=SUM(B3:M3)','=SUM(B4:M4)','=SUM(B5:M5)','=SUM(N2:N5)' ]
	template = {'Janvier':[0,0,0,0,'=SUM(B2:B5)'], 'Fevrier':[0,0,0,0,'=SUM(C2:C5)'], 'Mars':[0,0,0,0,'=SUM(D2:D5)'], 'Avril':[0,0,0,0,'=SUM(E2:E5)'],
	 'Mai':[0,0,0,0,'=SUM(F2:F5)'], 'Juin':[0,0,0,0,'=SUM(G2:G5)'], 'Juillet':[0,0,0,0,'=SUM(H2:H5)'], 'Aout':[0,0,0,0,'=SUM(I2:I5)'], 
	 'Septembre':[0,0,0,0,'=SUM(J2:J5)'], 'Octobre':[0,0,0,0,'=SUM(K2:K5)'], 'Novembre':[0,0,0,0,'=SUM(L2:L5)'], 'Decembre':[0,0,0,0,'=SUM(M2:M5)'],
	  'Cumulative Year':cumulative} 
	df = pd.DataFrame(template, index=['Rent', 'Premier', 'Suivi', 'Taxes', 'Total'])
	return df

#This is currently the same as initialize_new_sheet but I'm keeping it seperate because I'll likely be changing it
def initialize_summary_sheet():
	cumulative = ['=SUM(B2:M2)', '=SUM(B3:M3)','=SUM(B4:M4)','=SUM(B5:M5)','=SUM(N2:N5)' ]
	template = {'Janvier':[0,0,0,0,'=SUM(B2:B5)'], 'Fevrier':[0,0,0,0,'=SUM(C2:C5)'], 'Mars':[0,0,0,0,'=SUM(D2:D5)'], 'Avril':[0,0,0,0,'=SUM(E2:E5)'],
	 'Mai':[0,0,0,0,'=SUM(F2:F5)'], 'Juin':[0,0,0,0,'=SUM(G2:G5)'], 'Juillet':[0,0,0,0,'=SUM(H2:H5)'], 'Aout':[0,0,0,0,'=SUM(I2:I5)'], 
	 'Septembre':[0,0,0,0,'=SUM(J2:J5)'], 'Octobre':[0,0,0,0,'=SUM(K2:K5)'], 'Novembre':[0,0,0,0,'=SUM(L2:L5)'], 'Decembre':[0,0,0,0,'=SUM(M2:M5)'],
	  'Cumulative Year':cumulative} 
	df = pd.DataFrame(template, index=['Rent', 'Premier', 'Suivi', 'Taxes', 'Total'])
	return df

def create_workbook(filename, year):
	df = initialize_new_sheet()
	df.to_excel(filename, str(year), index=True)

