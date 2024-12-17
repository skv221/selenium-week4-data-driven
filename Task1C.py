import json

jsonFile = open("D:\Selenium Practices\Week 4\Test Form.json")

jsonData = json.load(jsonFile)

for data in jsonData:
    print(data['Name']," | ", data['Email']," | ", data['Expected Behaviour'])