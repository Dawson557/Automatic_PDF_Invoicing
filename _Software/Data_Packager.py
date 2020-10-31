import pandas as pd

class Therapist_Data_Package:

	def __init__(self, therapist_data, month_data):
		# self.therapist_data = therapist_data
		# self.month_data = month_data
		self.therapists = {}
		for i in therapist_data.index:
			therapist = str(therapist_data['Name'][i])
			first = float(therapist_data['First Appointment Percentage'][i])
			follow = float(therapist_data['Follow Up Percentage'][i])
			rent = float(therapist_data['Rent'][i])
			DB_team = True if therapist_data['DB Team'][i] == "Y" else False

			#load services into dictionary
			services = {}
			for j in month_data.index:
				if str(month_data['Professional'][j]) == therapist:
					service = month_data['Service'][j].strip()
					price = month_data['Price'][j]
					quantity = month_data['Quantity'][j]
					revenue = month_data['Revenue'][j]
					services[service] = {'price':price, 'quantity':quantity, 'revenue':revenue}



			self.therapists[therapist] = { 'first':first, 'follow':follow, 'rent':rent, 'DB':DB_team, 'services':services}

		self.num_therapists = len(self.therapists)

	def isDB_team(self, key):
		return self.therapists[key]['DB']

	def get_drawing_variables(self, key, service):
		if ("premier" in service.lower()):
			service_type = 'first'
			rate = self.therapists[key][service_type]
		elif ("suivi" in service.lower()):
			service_type = 'follow'
			rate = self.therapists[key][service_type]
		else:
			service_type = None
			rate = 0
		revenue = self.therapists[key]['services'][service]['revenue']
		commission = rate * revenue
		quantity = self.therapists[key]['services'][service]['quantity']
		return rate, revenue, commission, quantity


				


