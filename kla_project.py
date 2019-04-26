import pandas as pd
import numpy as np
import csv

"""Reads data from veganBerrieData.xlsx
	Input: dictionaty for accumulating data
	Returns the same dictionary with a new key for the farm containing the data from each week
"""
def outgrowthOrchard(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="A:K")
	#replace type nan cells with 0
	data = data.fillna(0)
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

"""Reads data from veganBerrieData.xlsx
	Input: dictionaty for accumulating data
	Returns the same dictionary with a new key for the farm containing the data from each week
"""
def organicGardens(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="M:W")
	data = data.fillna(0)
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

"""Reads data from veganBerrieData.xlsx
	Input: dictionaty for accumulating data
	Returns the same dictionary with a new key for the farm containing the data from each week
"""
def frugalFarms(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="Y:AI")
	data = data.fillna(0)
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

"""Reads data from veganBerrieData.xlsx
	Input: dictionaty for accumulating data
	Returns the same dictionary with a new key for the farm containing the data from each week
"""
def powerHousePlantation(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="AK:AU")
	data = data.fillna(0)
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

"""Reads data from veganBerrieData.xlsx
	Input: dictionaty for accumulating data
	Returns the same dictionary with a new key for the farm containing the data from each week
"""
def gardantianGrove(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="AW:BG")
	data = data.fillna(0)
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

"""Reads data from veganBerrieData.xlsx
	Input: dictionaty for accumulating data
	Returns the same dictionary with a new key for the farm containing the data from each week
"""
def berryBonds(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="BI:BS")
	data = data.fillna(0)
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

"""Reads data from veganBerrieData.xlsx
	Input: dictionaty for accumulating data
	Returns the same dictionary with a new key for the farm containing the data from each week
"""
def friendlyFarmstead(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="BU:CE")
	data = data.fillna(0)
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

"""Reads data from veganBerrieData.xlsx
	Input: dictionaty for accumulating data
	Returns the same dictionary with a new key for the farm containing the data from each week
"""
def nancysNursery(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="CG:CQ")
	data = data.fillna(0)
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

"""Reads data from veganBerrieData.xlsx
	Input: dictionaty for accumulating data
	Returns the same dictionary with a new key for the farm containing the data from each week
"""
def easygoingEstates(data_dict):
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='Data from Farmers', usecols="CS:DC")
	data = data.fillna(0)
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

"""Calls each of the functions for accumulating data from each of the farms
	Input: none
	Returns completed dictionary where the keys are each farm, with a dictionary of all relevant data from that farm
"""
def make_data_Dict():
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

"""Reads data from veganBerrieData.xlsx, National Weather Predictions sheet
	Input: none
	Returns a dictionary with a key for each future week, with the rain and sun/temp forecast for that week
"""
def weather_predictions():
	weather_predictions_dict = {}
	data = pd.read_excel('veganBerrieData.xlsx', sheet_name='National Weather Predictions', usecols="E,F")
	for x in range(52):
		weekNum = 'Week ' + str(x + 1)
		weather_predictions_dict[weekNum] = {}
		weather_predictions_dict[weekNum]['Rain Forecast'] = data.iat[x+1,0]
		weather_predictions_dict[weekNum]['Sun/Temp Forecast'] = data.iat[x+1,1]
	return weather_predictions_dict

"""Finds the ideal level of moisture, mulch, nutrition, and fertilizer by averaging the data from the weeks
	across all the farms with the highest yields
	Input: dictionaty of data from farms
	Returns a dictionary containing the ideal levels for each category
"""
def find_ideals(farm_data_dict):
	highest_yield = 0
	highest_yield_data = []
	#iterate through every week at each farm to find weeks with the highest yields
	for farm in farm_data_dict:
		for week in farm_data_dict[farm]:
			if farm_data_dict[farm][week]['Weekly Yield'] > highest_yield:
				highest_yield = farm_data_dict[farm][week]['Weekly Yield']
				highest_yield_data += [farm_data_dict[farm][week]]
	#organize the data from all of the weeks with the highest yields
	moisture = []
	nutrition_lvl = []
	mulch = []
	fertilizer = []
	for week in highest_yield_data:
		moisture += [week['Moisture']]
		nutrition_lvl += [week['Nutrition Level']]
		mulch += [week['Mulch']]
		fertilizer += [week['Fertilizer']]
	#find averages for each category
	moisture_avg = sum(moisture)/len(moisture)
	nutrition_lvl_avg = sum(nutrition_lvl)/len(nutrition_lvl)
	mulch_avg = sum(mulch)/len(mulch)
	fertilizer_avg = sum(fertilizer)/len(fertilizer)
	#create dictionary of ideal averages
	ideal_averages_dict = {'Moisture': moisture_avg, 'Nutrition Level': nutrition_lvl_avg, 'Mulch': mulch_avg, 'Fertilizer': fertilizer_avg}

	return ideal_averages_dict

"""Creates a dictionary of the difference in fertilizer, mulch, nutrition, and water in order to establish some sense of how much these fields
	change from week to week
	Input: dictionaty of data from farms
	Returns a dictionary with keys of consecutive weeks and the differences in said categories
"""
def find_week_differences(farm_data_dict):
	i = 1
	differences_dict = {}
	#iterate through data from each farm
	for farm in farm_data_dict:
		differences_dict[farm] = {}
		#create dictionary of differences starting with weeks 1 and 2 and ended with weeks 51 and 52
		i = 1
		while i < 52:
			weeks_str = 'Week' + str(i) + ' and ' + str(i + 1)
			rain_amnt = farm_data_dict[farm]['Week ' + str(i + 1)]['Rain']
			sun_and_temp = farm_data_dict[farm]['Week ' + str(i + 1)]['Sun/Temp']
			watering = farm_data_dict[farm]['Week ' + str(i)]['Watering'] + farm_data_dict[farm]['Week ' + str(i + 1)]['Watering']
			fertilizer_diff = farm_data_dict[farm]['Week ' + str(i + 1)]['Fertilizer'] - farm_data_dict[farm]['Week ' + str(i)]['Fertilizer']
			mulch_diff = farm_data_dict[farm]['Week ' + str(i + 1)]['Mulch'] - farm_data_dict[farm]['Week ' + str(i)]['Mulch']
			nutrition_diff = farm_data_dict[farm]['Week ' + str(i + 1)]['Nutrition Level'] - farm_data_dict[farm]['Week ' + str(i)]['Nutrition Level']
			moisture_diff = farm_data_dict[farm]['Week ' + str(i + 1)]['Moisture'] - farm_data_dict[farm]['Week ' + str(i)]['Moisture']

			differences_dict[farm][weeks_str] = {'Rain': rain_amnt, 'Sun/Temp': sun_and_temp, 'Watering': watering, 'Fertilizer Difference': fertilizer_diff, 'Mulch Difference': mulch_diff, 'Nutrition Difference': nutrition_diff, 'Moisture Difference': moisture_diff}
			i += 1
	return differences_dict

"""Finds the ideal amount of water for each week based on the amount of watering and rain recieved by plants on weeks
	with the highest yield
	Input: dictionaty of data from farms, dictionary of ideal levels
	Returns a dictionary containing the ideal levels for each category
"""
def find_ideal_water_amount(farm_data_dict, ideal_dict):
	ideal_moisture = ideal_dict['Moisture']
	ideal_moisture_weeks = []
	#iterate though each week at each farm to find weeks where moisture level was ideal
	for farm in data_from_farms:
		for week in farm_data_dict[farm]:
			if (farm_data_dict[farm][week]['Moisture'] >= (ideal_moisture - 1)) and (farm_data_dict[farm][week]['Moisture'] <= (ideal_moisture + 1)):
				ideal_moisture_weeks += [farm_data_dict[farm][week]]
	#add the amount of watering and rain from ideal weeks
	water_amnts = []
	for week in ideal_moisture_weeks:
		water = week['Watering'] + week['Rain']
		water_amnts += [water]
	#average the amount of water, as most farms water every other week, this averahe is meant to mitigate that issue
	ideal_water_amount = sum(water_amnts)/len(water_amnts)
	return ideal_water_amount

"""Finds similar weeks based on rain and sun/temp data
	Input: dictionary of differences between weeks as created in find_week_differences, temperature for matching, rain amount for matching
	Returns the data from the weeks that match the inputed rain and temp (if there is not match nothing is returned)
"""
def find_similar_week(differences_dict, temp, rain):
	for farm in differences_dict:
		for weeks in differences_dict[farm]:
			if (differences_dict[farm][weeks]['Rain'] == rain) and (differences_dict[farm][weeks]['Sun/Temp'] == temp):
				return differences_dict[farm][weeks]

"""Creates a dictionary of reccomendations of fertilizer, mulch, and watering
	Input: ideal amounts dictionary as created in find_ideals
		dictionary of differences in weeks as created in find_week_differences
		dictionary of national weather predictions as created in weather_predictions
		ideal water level as found in find_ideal_water_amount
	Returns a dictionary comtaining the reccomendations for each week
"""
def make_reccomendations(ideal_nums, week_diffs_dict, national_data_dict, water_ideal):
	recs_dict = {}
	#creates week 1 dictionary using the ideal values
	recs_dict['Week 1'] = {'Fertilizer to add': ideal_nums['Fertilizer'], 'Mulch to add': ideal_nums['Mulch'], 'Watering': (water_ideal - national_data_dict['Week 1']['Rain Forecast'])}
	#iterates through the future weeks and finds a similar week and the differences are used to find reccomended amounts to add
	i = 1
	while i < 52:
		diffs = find_similar_week(week_diffs_dict, national_data_dict['Week ' + str(i)]['Sun/Temp Forecast'], national_data_dict['Week ' + str(i)]['Rain Forecast'])
		#if a similar week cannot be found, depreciation in levels are not considered and the levels from the previous week is used
		if diffs is None:
			fertilizer = recs_dict['Week ' + str(i)]['Fertilizer to add']
			mulch = recs_dict['Week ' + str(i)]['Mulch to add']
		else:
			fertilizer = recs_dict['Week ' + str(i)]['Fertilizer to add'] + diffs['Fertilizer Difference']
			mulch = recs_dict['Week ' + str(i)]['Mulch to add'] + diffs['Mulch Difference']
		water = water_ideal - national_data_dict['Week ' + str(i + 1)]['Rain Forecast']
		recs_dict['Week ' + str(i + 1)] = {'Fertilizer to add': fertilizer, 'Mulch to add': mulch, 'Watering': water}
		i += 1
	return recs_dict
 


def createCSVReport(recs_dict):
	""" Writes csv file with reccomandations
	Input: dictionary created in make_reccomendations
	"""
	outfile = open("veganBerrieRecomendations.csv", "w")
	writer = csv.writer(outfile)
	myFields = ['Week', 'Fertilizer to add', 'Mulch to add', 'Watering']
	writer.writerow(myFields)
	for week in list(recs_dict.keys()):
		writer.writerow([week, recs_dict[week]['Fertilizer to add'], recs_dict[week]['Mulch to add'], recs_dict[week]['Watering']])
	outfile.close()




data_from_farms = make_data_Dict()
national_weather_data = weather_predictions()
ideals_dict = find_ideals(data_from_farms)
differences_dict = find_week_differences(data_from_farms)
ideal_water_amnt = find_ideal_water_amount(data_from_farms, ideals_dict)
reccomendations = make_reccomendations(ideals_dict, differences_dict, national_weather_data, ideal_water_amnt)
createCSVReport(reccomendations)




























