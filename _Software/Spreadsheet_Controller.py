import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile


class Sheet_Handler:
	def __init__(self, month, year):
		self.month = month
		self.year = year
		self.rent_total = 0.0
		self.premier_total = 0.0
		self.suivi_total = 0.0
		self.taxes_total = 0.0

	def individual_report(self,rent, filename, data, year):
		# df = pd.read_excel(filename, sheet_name=year)

		# writer_object = pd.ExcelWriter(filename, engine='xlsxwriter')
		# df.to_excel(writer_object, sheet_name=str(year))
		# writer_object.save()	
		# self.update_totals(rent, premier, suivi, taxes)
		pass

	def update_totals(self, rent, premier, suivi, taxes):
		self.rent_total += rent
		self.premier_total += premier
		self.suivi_total += suivi	
		self.taxes_total += taxes

	def summary_report():
		pass	



def initialize_report(filename, year):
	cumulative = ['=SUM(B2:M2)', '=SUM(B3:M3)','=SUM(B4:M4)','=SUM(B5:M5)','=SUM(N2:N6)' ]
	template = {'January':[0,0,0,0,'=SUM(B2:B5)'], 'February':[0,0,0,0,'=SUM(C2:C5)'], 'March':[0,0,0,0,'=SUM(D2:D5)'], 'April':[0,0,0,0,'=SUM(E2:E5)'],
	 'May':[0,0,0,0,'=SUM(F2:F5)'], 'June':[0,0,0,0,'=SUM(G2:G5)'], 'July':[0,0,0,0,'=SUM(H2:H5)'], 'August':[0,0,0,0,'=SUM(I2:I5)'], 
	 'September':[0,0,0,0,'=SUM(J2:J5)'], 'October':[0,0,0,0,'=SUM(K2:K5)'], 'November':[0,0,0,0,'=SUM(L2:L5)'], 'December':[0,0,0,0,'=SUM(M2:M5)'],
	  'Cumulative Year':cumulative} 
	df = pd.DataFrame(template, index=['Rent', 'Premier', 'Suivi', 'Taxes', 'Total'])

	writer_object = pd.ExcelWriter(filename, engine='xlsxwriter')
	df.to_excel(writer_object, sheet_name=str(year))
	writer_object.save()



	

