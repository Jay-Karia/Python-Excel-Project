from openpyxl import Workbook
import json

def ReadJSONData():
    with open("CE_Analytics.json") as json_file:
        data = json.load(json_file)

ReadJSONData()