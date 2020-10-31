import os
from shutil import copyfile
from glob import glob
import argparse

#read/write excel
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
#Misc. utilities
from Misc_Utils import get_date
from Misc_Utils import create_filename
#creating canvas to draw pdf
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from Draw_Utils import draw_invoice_template, draw_service, draw_end, draw_rent
#email
from Emailer import email, build_email_service
#central spreadsheet tools
import Spreadsheet_Controller
#Information manager for therpasit data
import Data_Packager

parser = argparse.ArgumentParser()
parser.add_argument('--month', type=int, default=0)
parser.add_argument('--year', type=int, default=0)
parser.add_argument('--email', action='store_true')
parser.add_argument('--rent', action='store_true')

opt = parser.parse_args()

TPS_rate = 0.05
TVQ_rate = 0.09975

day, month, month_num, year = get_date(opt.rent, opt.month, opt.year)

#sender = "testa6390@gmail.com" #testing account 
sender = "Bianca@divanbleu.com" 
subject = "Facture de {}".format(month)

#Spreadsheet Controller object
sheet_handler = Spreadsheet_Controller.Sheet_Handler(month_num, year)

def main():
	#ClientRevenueReportList.xlsx is generated by GoRendezVous
	#OverviewSheet is maintained by Bianca, contains data about each therapists payment schemes
	month_data = pd.read_excel('ClientRevenueReportList.xlsx', sheet_name='Sheet1')
	therapist_data = pd.read_excel('OverviewSheet.xlsx', sheet_name="Sheet1")
	data_package = Data_Packager.Therapist_Data_Package(therapist_data, month_data)

	email_service = build_email_service()

	if not os.path.exists("_Recent"):
		os.makedirs("_Recent")
	files = glob('_Recent/*')
	for f in files:
		os.remove(f)

	#loading percent is used solely for visual queue
	loading_percent = 100.0/data_package.num_therapists
	percent_complete = 0.0


	for therapist in data_package.therapists:
		percent_complete += loading_percent

		#skip commission invoice for non-Divan blue members
		if (opt.rent == False) and not (data_package.isDB_team(therapist)):
			continue


		# therapist = str(therapist_data['Name'][i])
		# pays_tax = True if therapist_data['Pays Tax'][i] == "Y" else False
		# first_percentage = float(therapist_data['First Appointment Percentage'][i])
		# follow_percentage = float(therapist_data['Follow Up Percentage'][i])
		# rent = float(therapist_data['Rent'][i])

		file_directory, filename = create_filename(therapist, month, year, opt.rent)
		invoice = canvas.Canvas(file_directory + filename , pagesize=A4)
		position = draw_invoice_template(invoice, therapist, year, str(month_num).zfill(2), str(day).zfill(2))

		excel_fn = therapist + os.sep + therapist + "_totals.xlsx"
		services = data_package.therapists[therapist]['services']
		sheet_handler.individual_report(opt.rent, excel_fn, services)

		if (opt.rent):
			position = draw_rent(invoice, position, rent)
			partial = rent

		else:
			partial = 0.0
			for service in services:
				# draw_service(c, position, service, rate, revenue, commission, quantity)
				rate, revenue, commission, quantity = data_package.get_drawing_variables(therapist, service)
				if (rate > 0):
					position = draw_service(invoice, position, service, str(rate), revenue, commission, quantity)
					partial += revenue * rate
				else:
					position = draw_service(invoice, position, service, "-", revenue, 0, quantity)
				
				# if ("premier" in service.lower()):
				# 	service_type = 'first'
				# elif ("suivi" in service.lower()):
				# 	service_type = 'follow'
				# 	position = draw_service(invoice, position, service, str(first_percentage), revenue, revenue * first_percentage, quantity)
				# 	partial += revenue * first_percentage
				# elif ("suivi" in service.lower()):
				# 	ser
				# 	position = draw_service(invoice, position, service, str(follow_percentage), revenue, revenue * follow_percentage, quantity)
				# 	partial += revenue * follow_percentage
				# else:
				# 	position = draw_service(invoice, position, service, "-", revenue, 0, quantity)

		total = 0.0
		TPS_total = partial * TPS_rate
		TVQ_total = partial * TVQ_rate
		total = partial + TPS_total + TVQ_total

		draw_end(invoice, position, partial, TPS_total, TVQ_total, total)
		invoice.showPage()
		invoice.save()

		#Copy last created invoices into _Recent/ directory
		copyfile(file_directory + filename, "_Recent" + os.sep + filename)

		#Send emails
		if (opt.email) and (total != 0):
			email_address = str(therapist_data['E-mail'][i])
			if (email_address != "nan"):
				message_text = "Bonjour {},\n\
				\nVoici votre facture! Nous vous remercions pour votre paiement.\n\
				\nMerci de faire affaire avec nous!\nDivan bleu inc.".format(therapist.split()[0])
				email(email_service, sender, email_address, subject, message_text, file_directory, filename)

		
		print(" {}{} completed".format(int(percent_complete), "%"), end="\r")

	if (opt.email):
		print("")
		print("--- Invoices Created and Emailed ---")
	else:
		print("")
		print("--- Invoices Created ---")


if __name__ == '__main__':
    main()