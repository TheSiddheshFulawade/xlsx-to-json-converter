import openpyxl
import uuid

#Load the Workbook
workbook = openpyxl.load_workbook('service_list.xlsx')

# Select the first sheet
sheet = workbook.active

arrFields = []


# Iterate over each row in the sheet
for row in sheet.iter_rows(values_only=True):
    cell1 = row[0]
    cell2 = row[1]
    arrFields.append(row)


code = []
for value in arrFields[1:]:
    json_code = '''{
  "_id": {
    "$oid": " ''' + str(uuid.uuid4()) + ''' "
  },
  "user_id": {
    "$oid": " ''' + str(value[8]) + ''' "
  },
  "service_id": [
    {
      "$oid": " ''' + str(value[7]) + ''' "
    },
    {
      "$oid": " ''' + str(value[7]) + ''' "
    }
  ],
  "service_name": " ''' + value[1] + ''' ",
  "title": " ",
  "address": " ''' + value[2] + ''' ",
  "city": " ''' + value[0] + ''' ",
  "state": " ",
  "pincode": " ''' + str(value[4]) + ''' ",
  "contact_person": " ''' + str(value[3]) + ''' ",
  "contact_person_name": " ",
  "email_id": " ",
  "avaibility": {
    "avaibilityDays": {
      "Sunday": true,
      "Monday": true,
      "Wednesday": true,
      "Thursday": true,
      "Friday": true,
      "Saturday": true
    },
    "allDayAvailable": true,
    "availableTimeFrom": "10 Am",
    "availableTimeTo": "7 PM",
    "comment": "",
    "24by7avaibility": false
  },
  "address_details": {
    "lat": ''' + str(value[5]) + ''',
    "lng": ''' + str(value[6]) + ''',
    "Label": " ''' + value[2] + ''' ",
    "Municipality": " ",
    "Neighborhood": " ",
    "PostalCode": " ''' + str(value[4]) + ''' ",
    "Region": " ",
    "SubRegion": " "
  },
  "status": " ",
  "created_at": "",
  "updated_at": {
    "$date": " "
  },
  "created_by": "",
  "updated_by": {
    "$oid": " "
  },
  "location": {
    "type": " ",
    "coordinates": [''' + str(value[5]) + ''', ''' + str(value[6]) + ''']
  },
  "documents": [
    {
      "name": " ",
      "photo": " "
    }
  ],
  "registration_number": " "
}'''
    code.append(json_code)

code = "\n".join(code)
file_name = "Json.json"  
with open(file_name, 'w') as file:
    file.write(code)

