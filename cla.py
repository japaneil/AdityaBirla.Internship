from num2words import num2words
import survey
import pickle
from datetime import date
from spire.doc import *
from spire.doc.common import *

truck_num = input("Please enter truck number and press enter: ")

gross = int(input("Please enter gross weight and press enter: "))

tare = int(input("Please enter tare weight and press enter: "))

net = float(gross - tare)

value_rate = survey.routines.inquire('Add New Rate? ', default = False)
if value_rate == True:
    rate = input("Add new value: ")
    file_rate = open('rate.p', 'wb')
    pickle.dump(rate, file_rate)
    file_rate.close()
file_rate = open('rate.p', 'rb')
rate = pickle.load(file_rate)
file_rate.close()
amount = net * float(rate)
amount_num = num2words(amount, lang='en_IN' , to='currency', currency = 'INR') # check why conversion to inr gives error

trans_name = input("Please enter transporter's name and press enter: ")


file_challan = open('serial.txt', 'rb')
challan_num = int(pickle.load(file_challan))
file_challan.close()
challan_num += 1
file_challan = open('serial.txt', 'wb')
pickle.dump(str(challan_num), file_challan)
file_challan.close()

today = date.today()
date_today = today.strftime("%d/%m/%y")

mode = "Road"

destination_list = []
value = survey.routines.inquire('Add New Destination Value? ', default = False)

while value == True:
    destination_list.append(input("Add new value: "))
    value = survey.routines.inquire('Add New Destination Value? ', default = False)
    file = open('destination.p', 'wb')
    pickle.dump(destination_list, file)
    file.close()

value_rem = survey.routines.inquire('Remove Destination Value? ', default = False)

file = open('destination.p', 'rb')
destination_list = pickle.load(file)
file.close()

while value_rem == True:
    for i in range(len(destination_list)):
        print(str(i) + ". " + destination_list[i])
        i += 1
    remove = int(input("Please enter index of value you want to remove: "))
    destination_list.remove(destination_list[remove])
    file = open('destination.p', 'wb')
    pickle.dump(destination_list, file)
    file.close()
    value_rem = survey.routines.inquire('Remove Destination Value? ', default = False)

file = open('destination.p', 'rb')
destination_list = pickle.load(file)
file.close()

index = survey.routines.select('destination? ',  options = destination_list,  focus_mark = '-> ',  evade_color = survey.colors.basic('yellow'))

destination = destination_list[index]

# print(truck_num)
# print(gross)
# print(tare)
# print(amount)
# print(amount_num)
# print(trans_name)
# print(challan_num)
# print(date_today)
# print(mode)
# print(destination)

document = Document()
document.LoadFromFile("template1.docx")
document.Replace("Challan_num", str(challan_num), False, False)
document.Replace("trans_name", str(trans_name), False, False)
document.Replace("date_today", str(date_today), False, False)
document.Replace("mode_var", str(mode), False, False)
document.Replace("Destination_var", str(destination), False, False)
document.Replace("truck_num", str(truck_num), False, False)

document.SaveToFile(str(challan_num) + ".docx", FileFormat.Docx2016)
document.Close()