import xlsxwriter

data = [
    {
        'name': "Daniel Collins",
        'phone': "336-693-6860",
        'email': "danocogreen@gmail.com",
        'address': "Apt# 308 - 8209 University Ridge Dr",
        'city': "Charlotte",
        'state': "North Carolina",
        'zip_code': "28213",
        'country': "United States of America",
    },
    {
        'name': "Lauren Faw",
        'phone': "843-425-8313",
        'email': "danocogreen@gmail.com",
        'address': "Apt# 308 - 8209 University Ridge Dr",
        'city': "Charlotte",
        'state': "North Carolina",
        'zip_code': "28213",
        'country': "United States of America",
    },
    {
        'name': "Holly Collins",
        'phone': "336-509-7377",
        'email': "Hacollin03@yahoo.com",
        'address': "5275 Liberty Grove Rd",
        'city': "Liberty",
        'state': "North Carolina",
        'zip_code': "27298",
        'country': "United States of America",
    },
    {
        'name': "Emily Kennedy",
        'phone': "336-404-3176",
        'email': "collins.emily51@yahoo.com",
        'address': "1323 Browns Crossroads Rd",
        'city': "Staley",
        'state': "North Carolina",
        'zip_code': "27355",
        'country': "United States of America",
    }
]


workbook = xlsxwriter.Workbook("AllAboutPythonExcel.xlsx")
worksheet =  workbook.add_worksheet("First Sheet")

worksheet.write(0, 0, "#")
worksheet.write(0, 1, "Name")
worksheet.write(0, 2, "Phone")
worksheet.write(0, 3, "Email")
worksheet.write(0, 4, "Address")
worksheet.write(0, 5, "City")
worksheet.write(0, 6, "State")
worksheet.write(0, 7, "Zip Code")
worksheet.write(0, 8, "Country")

for index, entry in enumerate(data):
    worksheet.write(index+1, 0, str(index))
    worksheet.write(index+1, 1, entry["name"])
    worksheet.write(index+1, 2, entry["phone"])
    worksheet.write(index+1, 3, entry["email"])
    worksheet.write(index+1, 4, entry["address"])
    worksheet.write(index+1, 5, entry["city"])
    worksheet.write(index+1, 6, entry["state"])
    worksheet.write(index+1, 7, entry["zip_code"])
    worksheet.write(index+1, 8, entry["country"])

workbook.close()
