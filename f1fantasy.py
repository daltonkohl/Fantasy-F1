"""
Fantasy F1: This program is to be ran in conjuction with the F1 Fantatsy Spreadsheet to create the score updating 
portion while the excel spreadsheet handles roster creation and fronted visuals.
Author: Dalton Kohl
"""
#pandas handles reading the excel sheets to get the players and rosters.
import pandas as pd
#requests handles the response from the F1 api.
import requests
#unicode handles the special characters in the driver names.
import unidecode
#openpyxl handles the data insertion into the excel file.
import openpyxl


"""
description:
    The get_race_results(round) function uses the f1 api to get the points scored by each driver for the given round
    then returns a dictionary mapping driver to points scored that round.
params:
    round: the round number
returns:
    race_result_dict: a dictionary mapping the driver full name to the amount of points scored that round
"""
def get_race_results(round):
    #set api url and get the response for the given round
    url = f"https://ergast.com/api/f1/2024/{round}/results.json"
    response = requests.get(url).json()
    race_result_dict = {}
    json_result_list = response['MRData']['RaceTable']['Races'][0]['Results']
    #for each driver create the dictoinary entry from the api results
    for driver in json_result_list:
        driver_name = f"{unidecode.unidecode(driver['Driver']['givenName'])} {unidecode.unidecode(driver['Driver']['familyName'])}"
        points = int(driver['points'])
        race_result_dict[driver_name] = points
    return race_result_dict

"""
description:
    The get_owner_scores(race_result_dict) function gets the driver roster from the "F1 Fantasy.xlsx" excel file
    for each owner. It then goes through each driver in the owner's roster and gets the owners score for the round based
    on their four drivers round scores averaged. A dictionary is then returned mapping owner name to their round score.
params:
    race_result_dict: a dictionary mapping the drivers full name to their score
returns:
    owner_score_dict: a dictionary mapping the owner name to their point total
"""
def get_owner_scores(race_result_dict):
    owner_roster = {}
    owner_score_dict = {}
    #read in the excel sheet to get the roster dictionary mapping of owner to drivers
    roster_sheet = pd.read_excel(io='F1 Fantasy.xlsx', sheet_name = '2024 Draft', usecols='B:D').to_dict()
    #makes the values of the dictionary lists
    for key in roster_sheet.keys():
        owner_roster[key] = list(roster_sheet[key].values())[:-1]
    #loops through the owner's roster totaling the points of all their drivers then averaging those points
    for owner in owner_roster:
        point_count = 0
        for driver in owner_roster[owner]:
            try:
                point_count += race_result_dict[driver]
            except KeyError:
                print(f"Owner: {owner} Driver: {driver} did not participate in the race.")
        owner_score_dict[owner] = point_count / len(owner_roster[owner])
    return owner_score_dict

"""
description:
    The update_standings_by_round function takes the owner score dictionary and updates the excel sheet for the given round with their scores.
params:
    owner_score_dict: a dictionary mapping the owner's name to their score
    round: the round number
returns:
    None
"""
def update_standings_by_round(owner_score_dict, round):
    #opens the "F1 Fantasy.xlsx" excel workbook for editing
    score_sheet = openpyxl.load_workbook('F1 Fantasy.xlsx')
    #set the initial value for the correct column to edit
    col_count = 2
    #updates the correct cell with the owner's score updating the column count after each iteration
    for owner in owner_score_dict:
        score_sheet['2024 Standings'].cell(row = round + 1, column = col_count).value = owner_score_dict[owner]
        col_count += 1
    #saves the workbook so the updates are applied in excel
    score_sheet.save('F1 Fantasy.xlsx')


"""
description:
    The main(round) function is the main runner function that makes all subsequent function calls to get the correct owner scores and update the excel sheet.
params:
    round: the round number
returns:
    None
"""
def main(round):
    race_results = get_race_results(round)
    owner_scores = get_owner_scores(race_results)
    update_standings_by_round(owner_scores, round)


if(__name__ == "__main__"):
    main(3)

