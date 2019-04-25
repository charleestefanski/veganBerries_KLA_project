import pandas as pd

def outgrowthOrchard(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="A:K")
	data_dict['OutgrowthOrchard'] = {}
	for x in range(52):
		weekNum = 'Week ' + str(x+1)
		data_dict['OutgrowthOrchard'][weekNum] = {}
		data_dict['OutgrowthOrchard'][weekNum]['Fertilizer Added'] = data.iat[x+1,1]
		data_dict['OutgrowthOrchard'][weekNum]['Fertilizer'] = data.iat[x+1,2]
		data_dict['OutgrowthOrchard'][weekNum]['Mulch Added'] = data.iat[x+1,3]
		data_dict['OutgrowthOrchard'][weekNum]['Mulch'] = data.iat[x+1,4]
		data_dict['OutgrowthOrchard'][weekNum]['Rain'] = data.iat[x+1,5]
		data_dict['OutgrowthOrchard'][weekNum]['Watering'] = data.iat[x+1,6]
		data_dict['OutgrowthOrchard'][weekNum]['Nutrition Level'] = data.iat[x+1,7]
		data_dict['OutgrowthOrchard'][weekNum]['Moisture'] = data.iat[x+1,8]
		data_dict['OutgrowthOrchard'][weekNum]['Sun/Temp'] = data.iat[x+1,9]
		data_dict['OutgrowthOrchard'][weekNum]['Weekly Yield'] = data.iat[x+1,10]
	return data_dict

def organicGardens(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="M:W")
	data_dict['OrganicGardens'] = {}
	for x in range(52):
		weekNum = 'Week ' + str(x+1)
		data_dict['OrganicGardens'][weekNum] = {}
		data_dict['OrganicGardens'][weekNum]['Fertilizer Added'] = data.iat[x+1,1]
		data_dict['OrganicGardens'][weekNum]['Fertilizer'] = data.iat[x+1,2]
		data_dict['OrganicGardens'][weekNum]['Mulch Added'] = data.iat[x+1,3]
		data_dict['OrganicGardens'][weekNum]['Mulch'] = data.iat[x+1,4]
		data_dict['OrganicGardens'][weekNum]['Rain'] = data.iat[x+1,5]
		data_dict['OrganicGardens'][weekNum]['Watering'] = data.iat[x+1,6]
		data_dict['OrganicGardens'][weekNum]['Nutrition Level'] = data.iat[x+1,7]
		data_dict['OrganicGardens'][weekNum]['Moisture'] = data.iat[x+1,8]
		data_dict['OrganicGardens'][weekNum]['Sun/Temp'] = data.iat[x+1,9]
		data_dict['OrganicGardens'][weekNum]['Weekly Yield'] = data.iat[x+1,10]
	return data_dict

def frugalFarms(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="Y:AI")
	data_dict['FrugalFarms'] = {}
	for x in range(52):
		weekNum = 'Week ' + str(x+1)
		data_dict['FrugalFarms'][weekNum] = {}
		data_dict['FrugalFarms'][weekNum]['Fertilizer Added'] = data.iat[x+1,1]
		data_dict['FrugalFarms'][weekNum]['Fertilizer'] = data.iat[x+1,2]
		data_dict['FrugalFarms'][weekNum]['Mulch Added'] = data.iat[x+1,3]
		data_dict['FrugalFarms'][weekNum]['Mulch'] = data.iat[x+1,4]
		data_dict['FrugalFarms'][weekNum]['Rain'] = data.iat[x+1,5]
		data_dict['FrugalFarms'][weekNum]['Watering'] = data.iat[x+1,6]
		data_dict['FrugalFarms'][weekNum]['Nutrition Level'] = data.iat[x+1,7]
		data_dict['FrugalFarms'][weekNum]['Moisture'] = data.iat[x+1,8]
		data_dict['FrugalFarms'][weekNum]['Sun/Temp'] = data.iat[x+1,9]
		data_dict['FrugalFarms'][weekNum]['Weekly Yield'] = data.iat[x+1,10]
	return data_dict

def powerHousePlantation(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="AK:AU")
	data_dict['PowerHousePlantation'] = {}
	for x in range(52):
		weekNum = 'Week ' + str(x+1)
		data_dict['PowerHousePlantation'][weekNum] = {}
		data_dict['PowerHousePlantation'][weekNum]['Fertilizer Added'] = data.iat[x+1,1]
		data_dict['PowerHousePlantation'][weekNum]['Fertilizer'] = data.iat[x+1,2]
		data_dict['PowerHousePlantation'][weekNum]['Mulch Added'] = data.iat[x+1,3]
		data_dict['PowerHousePlantation'][weekNum]['Mulch'] = data.iat[x+1,4]
		data_dict['PowerHousePlantation'][weekNum]['Rain'] = data.iat[x+1,5]
		data_dict['PowerHousePlantation'][weekNum]['Watering'] = data.iat[x+1,6]
		data_dict['PowerHousePlantation'][weekNum]['Nutrition Level'] = data.iat[x+1,7]
		data_dict['PowerHousePlantation'][weekNum]['Moisture'] = data.iat[x+1,8]
		data_dict['PowerHousePlantation'][weekNum]['Sun/Temp'] = data.iat[x+1,9]
		data_dict['PowerHousePlantation'][weekNum]['Weekly Yield'] = data.iat[x+1,10]
	return data_dict

def gardantianGrove(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="AW:BG")
	data_dict['GardanationGrove'] = {}
	for x in range(52):
		weekNum = 'Week ' + str(x+1)
		data_dict['GardanationGrove'][weekNum] = {}
		data_dict['GardanationGrove'][weekNum]['Fertilizer Added'] = data.iat[x+1,1]
		data_dict['GardanationGrove'][weekNum]['Fertilizer'] = data.iat[x+1,2]
		data_dict['GardanationGrove'][weekNum]['Mulch Added'] = data.iat[x+1,3]
		data_dict['GardanationGrove'][weekNum]['Mulch'] = data.iat[x+1,4]
		data_dict['GardanationGrove'][weekNum]['Rain'] = data.iat[x+1,5]
		data_dict['GardanationGrove'][weekNum]['Watering'] = data.iat[x+1,6]
		data_dict['GardanationGrove'][weekNum]['Nutrition Level'] = data.iat[x+1,7]
		data_dict['GardanationGrove'][weekNum]['Moisture'] = data.iat[x+1,8]
		data_dict['GardanationGrove'][weekNum]['Sun/Temp'] = data.iat[x+1,9]
		data_dict['GardanationGrove'][weekNum]['Weekly Yield'] = data.iat[x+1,10]
	return data_dict

def berryBonds(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="BI:BS")
	data_dict['BerryBonds'] = {}
	for x in range(52):
		weekNum = 'Week ' + str(x+1)
		data_dict['BerryBonds'][weekNum] = {}
		data_dict['BerryBonds'][weekNum]['Fertilizer Added'] = data.iat[x+1,1]
		data_dict['BerryBonds'][weekNum]['Fertilizer'] = data.iat[x+1,2]
		data_dict['BerryBonds'][weekNum]['Mulch Added'] = data.iat[x+1,3]
		data_dict['BerryBonds'][weekNum]['Mulch'] = data.iat[x+1,4]
		data_dict['BerryBonds'][weekNum]['Rain'] = data.iat[x+1,5]
		data_dict['BerryBonds'][weekNum]['Watering'] = data.iat[x+1,6]
		data_dict['BerryBonds'][weekNum]['Nutrition Level'] = data.iat[x+1,7]
		data_dict['BerryBonds'][weekNum]['Moisture'] = data.iat[x+1,8]
		data_dict['BerryBonds'][weekNum]['Sun/Temp'] = data.iat[x+1,9]
		data_dict['BerryBonds'][weekNum]['Weekly Yield'] = data.iat[x+1,10]
	return data_dict

def friendlyFarmstead(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="BU:CE")
	data_dict['FriendlyFarmstead'] = {}
	for x in range(52):
		weekNum = 'Week ' + str(x+1)
		data_dict['FriendlyFarmstead'][weekNum] = {}
		data_dict['FriendlyFarmstead'][weekNum]['Fertilizer Added'] = data.iat[x+1,1]
		data_dict['FriendlyFarmstead'][weekNum]['Fertilizer'] = data.iat[x+1,2]
		data_dict['FriendlyFarmstead'][weekNum]['Mulch Added'] = data.iat[x+1,3]
		data_dict['FriendlyFarmstead'][weekNum]['Mulch'] = data.iat[x+1,4]
		data_dict['FriendlyFarmstead'][weekNum]['Rain'] = data.iat[x+1,5]
		data_dict['FriendlyFarmstead'][weekNum]['Watering'] = data.iat[x+1,6]
		data_dict['FriendlyFarmstead'][weekNum]['Nutrition Level'] = data.iat[x+1,7]
		data_dict['FriendlyFarmstead'][weekNum]['Moisture'] = data.iat[x+1,8]
		data_dict['FriendlyFarmstead'][weekNum]['Sun/Temp'] = data.iat[x+1,9]
		data_dict['FriendlyFarmstead'][weekNum]['Weekly Yield'] = data.iat[x+1,10]
	return data_dict

def nancysNursery(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="CG:CQ")
	data_dict['NancysNursery'] = {}
	for x in range(52):
		weekNum = 'Week ' + str(x+1)
		data_dict['NancysNursery'][weekNum] = {}
		data_dict['NancysNursery'][weekNum]['Fertilizer Added'] = data.iat[x+1,1]
		data_dict['NancysNursery'][weekNum]['Fertilizer'] = data.iat[x+1,2]
		data_dict['NancysNursery'][weekNum]['Mulch Added'] = data.iat[x+1,3]
		data_dict['NancysNursery'][weekNum]['Mulch'] = data.iat[x+1,4]
		data_dict['NancysNursery'][weekNum]['Rain'] = data.iat[x+1,5]
		data_dict['NancysNursery'][weekNum]['Watering'] = data.iat[x+1,6]
		data_dict['NancysNursery'][weekNum]['Nutrition Level'] = data.iat[x+1,7]
		data_dict['NancysNursery'][weekNum]['Moisture'] = data.iat[x+1,8]
		data_dict['NancysNursery'][weekNum]['Sun/Temp'] = data.iat[x+1,9]
		data_dict['NancysNursery'][weekNum]['Weekly Yield'] = data.iat[x+1,10]
	return data_dict


def easygoingEstates(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="CS:DC")
	data_dict['EasygoingEstates'] = {}
	for x in range(52):
		weekNum = 'Week ' + str(x+1)
		data_dict['EasygoingEstates'][weekNum] = {}
		data_dict['EasygoingEstates'][weekNum]['Fertilizer Added'] = data.iat[x+1,1]
		data_dict['EasygoingEstates'][weekNum]['Fertilizer'] = data.iat[x+1,2]
		data_dict['EasygoingEstates'][weekNum]['Mulch Added'] = data.iat[x+1,3]
		data_dict['EasygoingEstates'][weekNum]['Mulch'] = data.iat[x+1,4]
		data_dict['EasygoingEstates'][weekNum]['Rain'] = data.iat[x+1,5]
		data_dict['EasygoingEstates'][weekNum]['Watering'] = data.iat[x+1,6]
		data_dict['EasygoingEstates'][weekNum]['Nutrition Level'] = data.iat[x+1,7]
		data_dict['EasygoingEstates'][weekNum]['Moisture'] = data.iat[x+1,8]
		data_dict['EasygoingEstates'][weekNum]['Sun/Temp'] = data.iat[x+1,9]
		data_dict['EasygoingEstates'][weekNum]['Weekly Yield'] = data.iat[x+1,10]
	return data_dict

def makeData_Dict():
	veganBerrie_data_dict = {}
	outgrowthOrchard(veganBerrie_data_dict)
	organicGardens(veganBerrie_data_dict)
	frugalFarms(veganBerrie_data_dict)
	powerHousePlantation(veganBerrie_data_dict)
	gardantianGrove(veganBerrie_data_dict)
	berryBonds(veganBerrie_data_dict)
	friendlyFarmstead(veganBerrie_data_dict)
	nancysNursery(veganBerrie_data_dict)
	easygoingEstates(veganBerrie_data_dict)
	return veganBerrie_data_dict

def weatherPredictions():
	weather_predictions_dict = {}
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='National Weather Predictions', usecols="E,F")
	for x in range(52):
		weekNum = 'Week ' + str(x + 1)
		weather_predictions_dict[weekNum] = {}
		weather_predictions_dict[weekNum]['Rain Forecast'] = data.iat[x+1,0]
		weather_predictions_dict[weekNum]['Sun/Temp Forecast'] = data.iat[x+1,1]
	return weather_predictions_dict


def find_matching_weeks(farm_data_dict, future_weather_dict):
	matching_weeks = {}

	for future_week in future_weather_dict:
		matching_weeks[future_week] = []
		rain_amnt = future_weather_dict[future_week]['Rain Forecast']
		sun_temp = future_weather_dict[future_week]['Sun/Temp Forecast']
		for farm in farm_data_dict:
			for week in farm_data_dict[farm]:
				if (((farm_data_dict[farm][week]['Rain'] <= (rain_amnt + 0.01)) and (farm_data_dict[farm][week]['Rain'] >= (rain_amnt - 0.01))) and (farm_data_dict[farm][week]['Sun/Temp'] >= (sun_temp - 3)) and (farm_data_dict[farm][week]['Sun/Temp'] <= (sun_temp + 3))):
					matching_weeks[future_week] += [farm_data_dict[farm][week]]
	return matching_weeks

def decide_recommendations(matching_weeks_dict):
	reccomendation_dict = {}
	highest_yield_data = {}
	for week in matching_weeks_dict:
		highest_yield = 0
		best_match = {}
		for match in matching_weeks_dict[week]:
			if match['Weekly Yield'] > highest_yield:
				highest_yield = match['Weekly Yield']
				best_match = match
		highest_yield_data[week] = best_match
	print(highest_yield_data)
	
	for week in highest_yield_data:
		if highest_yield_data[week] != {}:
			print(week)
			reccomendation_dict[week] = {}
			fertilizer_amnt = 0
			fertilizer_amnt = highest_yield_data[week]['Fertilizer']
			if highest_yield_data[week]['Fertilizer Added'] != 'nan':
				fertilizer_amnt += highest_yield_data[week]['Fertilizer Added']
			mulch_amnt = 0
			mulch_amnt = highest_yield_data[week]['Mulch']
			if highest_yield_data[week]['Mulch Added'] != 'nan':
				mulch_amnt += highest_yield_data[week]['Mulch Added']
			watering_amnt = 0
			if highest_yield_data[week]['Watering'] != 'nan':
				watering_amnt += highest_yield_data[week]['Watering']
			reccomendation_dict[week]['Fertilizer Input'] = fertilizer_amnt
			reccomendation_dict[week]['Mulch Input'] = mulch_amnt
			reccomendation_dict[week]['Watering Input'] = watering_amnt
		else:
			reccomendation_dict[week] = 'empty week'
	print(reccomendation_dict)



#def write_reccomendations(matching_weeks_dict):
	#writer = pd.ExcelWriter('veganBerrieData.xlsx', engine='xlsxwriter')
	#df.to_excel(writer, sheet_name='National Weather Predictions')



data_from_farms = makeData_Dict()
national_weather_data = weatherPredictions()

#print(data_from_farms)
matching_dict = find_matching_weeks(data_from_farms, national_weather_data)

decide_recommendations(matching_dict)


























