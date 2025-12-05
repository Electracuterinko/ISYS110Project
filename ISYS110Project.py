import pip
import sys #Needed for exiting program when user does not want to order flowers
import xlsxwriter #Needed when writing to excel files 
import pandas as pd #Needed for reading excel files 
from pydantic import BaseModel, EmailStr #Needed for BaseModel Class and email validation
import datetime #Neeed to access current and future date validation 

flower_menu = pd.read_excel('BoqPrices.xlsx') # Reading in and creating a variable for bouquet prices
accessory_menu = pd.read_excel('accessory.xlsx') # Reading in and creating a variable for accessory prices

#This function asks the user if they want to order flowers
#If they do the program returns to the main code and continues the program
#If not program exits
#reiterates through the function until valid input is put in 
def are_we_ordering():
    first_try = input("\nWould you you like to order flowers? Y/N ")
    if first_try == "N" or first_try =="n":
        print("\nThank you for visiting our shop!")
        sys.exit(0)
    elif first_try =="Y" or first_try =="y":
        return
    else:
        print("\nI'm sorry I don't understand.")
        are_we_ordering()


#This function asks the user which bouquet they want
#reiterates through function until valid input is put in 
#Displays chosen bouquet
#Returns chosen boquet name and partial excel index
def so_were_ordering():
    try: 
        boq = int(input("Just type in the menu number(0-9)\n "))
        if boq > 9 or boq <0:
            print("\nOops! That's not a valid menu number \n")
            return so_were_ordering()
        else:       
           cell_value = flower_menu.iat[boq, 0]
           print("Great choice! You chose the ", cell_value, "bouquet!")
           return cell_value, boq
        
    except ValueError:
        print("\nOops! Just valid numbers please!\n")
        return so_were_ordering()


#This function asks the user if they want to add accessories to their bouquet
#It is not required to add accessories 
#It reiterates through the function until valid input is entered 
#Returns a True or Falso to trigger or pass the next function 
def accessory_maybe():
    accessory_pls = input("\nWould you like to add an accessory to that bouquet? Y/N ")
    if accessory_pls == "N" or accessory_pls =="n":
        return False
    elif accessory_pls == "Y" or accessory_pls =="y":
        return True
    else:
        print("I'm sorry I don't understand. \n")
        return accessory_maybe()    

#This function asks the user how many accessories the user wants to add to their bouquet
#It gives them a max of 5 accessories 
#It reiterates through the function until valid input is entered
#Returns number of accessories wanted 
def accessory_num():
    try:
        acces_num = int(input("\nGreat! How many accessories do you want? (Max 5)\n "))
        if acces_num > 5 or acces_num < 0:
            print("Oops! That's not a valid number of accessories \n")
            return accessory_num()
        else:
           return acces_num
    except ValueError:
        print("Oops! Just valid number of accessories please\n")
        return accessory_num()

#We ask the user which accessories they want to add to their bouquet
#It reiterates through the function until valid input is entered
#Returns accessory type which is then added to a list 
def accessory_type():
    try:
        accessory = int(input("\nJust type in the accessory menu number(0-4)\n "))
        if accessory > 4 or accessory < 0:
            print("Oops! That's not a valid accessory menu number \n")
            return accessory_type()
        else:
           return accessory
    except ValueError:
        print("Oops! Just valid accessory numbers please!\n")
        return accessory_type()

#This class is used for Email validation
#This emailstr variable from Pydantic is great for validating emails 
class Email(BaseModel):
    email:EmailStr

#This function asks the user for their email 
#Uses Email Class for validation 
#Iterates through function until valid email is given 
#returns email 
def get_email():
    try:
        e = input("\nWhat is your email address? ")
        tele = Email( email = e)
        return e
        
    except ValueError:
        print("\nThat is not a valid email address")
        return get_email()
    except TypeError:
        print("\nThat is not a valid email address")
        return get_email()

#Asks user for their first and last names 
#Capitalizes names in case uer does not 
# Checks to make sure both are only letters
# Iterates through function until valid input is entered 
# returns first and last name  
def get_name():
    print("Please Enter")
    f = input("First Name: ")
    f = f.title()
    l = input("Last Name: ")
    l = l.title()
    if f.isalpha() == True and l.isalpha() ==True:
        return f, l 
    else:
        print("\nOops, some input was invalid, let's try again")
        return get_name()


#This function asks user for their number
#it checks that the input is a number and correct number of numbers
#reiterates through function until valid input is entered
#Returns number
def get_number():
    try:
        num = int(input("\nWhat is your phone number? \nPlease input with no dashes or spaces (ex. 2698675309) "))
        form_num = str(num)
        if len(form_num) == 10:
                return num
        else:
            print("\nThat phone number is invalid")
            return get_number()
    except ValueError:
        print("\nThat is not a valid phone number")
        return get_number()

#This function checks to see if there are any numbers in a string, if there 
# are it returns positive  
def has_numbers(inputString):
    return any(char.isdigit() for char in inputString)

#This function asks the user for their address
# It checks that the right information is numerical
#and the right information isnt
# Also capitlized road and city in case user doesn't
# address information is returned
# reiterates tthrough function until valid input is entered     
def get_address():
    try:
        print("Please enter")
        house_num = int(input("Street Number: "))
        road = input("Road Name: ")
        road = road.title()
        city = input("City: ")
        city = city.title()
        zip = int(input("Zipcode: "))

        if has_numbers(city) ==False and len(str(zip)) == 5:
            return house_num, road, city, zip
        else:
            print("\nThat is not a valid address")
            return get_address()
    except ValueError:
        print("\nThat is not a valid address, somebodu put letters in the wrong place!")
        return get_address()


#This function asks the user for the day they want the flowers delivered
#It uses the datetime module to create date instances and we compare them
#To make sure the date the user wants is far enough away from the order
# date and is a date that hasn't already passed
# reiterates through code until valid input is entered 
def get_deliver_date():
    try:
        print("Please Enter in number format (MM/DD/YEAR)")
        mm = int(input("Month: "))
        dd = int(input("Day: "))
        yr = int(input("Year: "))
        deliver_date = datetime.date(yr, mm, dd)
        order_date = datetime.date.today()
        barmin = order_date + datetime.timedelta(days=3)
       
        if deliver_date < order_date:
            print("\nDelivery date has passed")
            return get_deliver_date()
        elif deliver_date < barmin:
            print("\nNot enough days notice, 3 days at least please")
            return get_deliver_date()
        else:
            return deliver_date, order_date
    except ValueError:
        print("\nThose are not valid date options")
        return get_deliver_date()


#This function writes the report for the sale
#This could have probably been done without being quite so bulky but I was kind of rushing
#Near the end 
#We use the xlsxwriter module to create the file, name sheets, create formats so that the 
#information looks correct in the excel fil, even set column sizes 
#We have to loop through the order list to make sure we get all of the accessories
#"If there is any and get their prices
#In the file we add all of the accessories and bouquet together and add tax to
#find the total amount 
#We name the excel file as order report and the users customer ID
def write_to_excel(f,l,e,n,user_ID,h,r,c,z,rf,rl,rh,rr,rc,rz,d_date,o_date, cell, order, boq):
    filename = f"Order_Report_{user_ID}.xlsx"
    order_file = xlsxwriter.Workbook(filename)

    phone_format = order_file.add_format({'num_format': r'(000) ""000""-0000'})
    customer_info = order_file.add_worksheet('Customer Information')
    customer_info.set_column(0, 12, 20)
    customer_info.write(0,0, 'Customer First Name')
    customer_info.write(0,1, 'Customer Last Name')
    customer_info.write(0,2, 'User ID')
    customer_info.write(0,3, 'Email')
    customer_info.write(0,4, 'Phone Number')
    customer_info.write(0,5, 'House Number')
    customer_info.write(0,6, 'Road')
    customer_info.write(0,7, 'City')
    customer_info.write(0,8, 'Zip Code')
    customer_info.write(1,0, f)
    customer_info.write(1,1, l)
    customer_info.write(1,2, user_ID)
    customer_info.write(1,3, e)
    customer_info.write(1,4, n, phone_format)
    customer_info.write(1,5, h)
    customer_info.write(1,6, r)
    customer_info.write(1,7, c)
    customer_info.write(1,8, z)

    recipient_info = order_file.add_worksheet("Recipient Information")
    recipient_info.set_column(0, 12, 20)
    recipient_info.write(0,0, 'Recipient First Name')
    recipient_info.write(0,1, 'Recipient Last Name')
    recipient_info.write(0,2, 'House Number')
    recipient_info.write(0,3, 'Road')
    recipient_info.write(0,4, 'City')
    recipient_info.write(0,5, 'Zip Code')
    recipient_info.write(1,0, rf)
    recipient_info.write(1,1, rl)
    recipient_info.write(1,2, rh)
    recipient_info.write(1,3, rr)
    recipient_info.write(1,4, rc)
    recipient_info.write(1,5, rz)

    date_format = order_file.add_format({'num_format': 'dd/mm/yyyy'})
    order_info = order_file.add_worksheet("Order Information")
    currency_format = order_file.add_format({'num_format': '$#,##0.00'})
    order_info.set_column(0, 12, 20)
    order_info.write(0,0, "Date Ordered")
    order_info.write(0,1, "Delivery Date")
    order_info.write(0,2, "Bouquet Ordered")
    order_info.write(1,0, o_date, date_format)
    order_info.write(1,1, d_date, date_format)
    order_info.write(1,2, cell)
    order_info.write(2,2, flower_menu.iat[boq, 1], currency_format)
    for i in range (len(order)):
        order_info.write(0, 3+i, "Accessory")
    for i in range (len(order)):
        order_info.write(1, 3+i, accessory_menu.iat[order[i], 0])
    for i in range (len(order)):
        order_info.write(2, 3+i, accessory_menu.iat[order[i], 1], currency_format)
    order_info.write(1,8, "Total w/o Tax")
    order_info.write(1,9, "Tax")
    order_info.write_formula(2, 8, '=SUM(C3:H3)', currency_format)
    percent_format = order_file.add_format({'num_format': '0.00%'})
    order_info.write(2, 9, .06, percent_format)
    order_info.write(1,10, "Total w/ Tax")
    order_info.write_formula(2, 10, '=I3*1.06', currency_format)
    order_file.close()
    return filename
            
#Welcome Guests
print("Hello! Welcome to Bree's Flower Boutique!")
#Check to see if we're ordering, if not program exits, if so program continues 
are_we_ordering()
#If program continues we start asking for information on the one paying
#for the bouquet by calling function 
print("\nGreat! First we just need to get some information about you!.\n")
f, l = get_name()
e = get_email()
n = get_number()
#Create User ID from first and last name and phone number
user_ID = f + l + str(n)
#finds out users address
print("\nWhat is your address?")
h, r, c, z = get_address()
#Displayes flower menu
print("\nNow what bouquet would you like to add to your order? \n ")
print("This is our flower menu\n")
print(flower_menu)
print("")
#Creates list for accesories 
order=[]
#Find out which boquet user wants 
cell, boq = (so_were_ordering())
#Displays accessory menu 
print("\nThis is our accessory menu\n")
print(accessory_menu)

#If user doesn't want accesories we bypass the for loop 
if accessory_maybe() == False:
    pass
else:
    #If guess does want accesories we ask how many and then ask which ones
    #for as many as they told us they wanted 
    for i in range(accessory_num()):
        order.append(accessory_type())

#Now we ask for information on the recipient basically  the same way 
#we did with the user
print("What a beautiful bouquet you've created! \n")
print("Now we just need some information on the recipient ")
print("\nWhat is the name of the recipient? \n")
rf, rl = get_name()
print("\nWhere are we sending this bouquet?\n")
rh, rr, rc, rz = get_address()
print("\nWhen should we deliver this bouquet? ")
print("We do need at least 3 days notice\n")
#Find out desired deliver date and create current date 
d_date, o_date = get_deliver_date()
#We send all of our information the be written into a excel file 
filename = write_to_excel(f,l,e,n,user_ID,h,r,c,z,rf,rl,rh,rr,rc,rz,d_date,o_date, cell, order, boq)

#Then we display user information from file 
print("\nHere is your information\n")
your_info = pd.read_excel(filename, sheet_name = 0)
print(your_info)

#Then we display recipient information from file 
print("\nHere is the recipients information\n")
recip_info = pd.read_excel(filename, sheet_name = 1)
print(recip_info)

#Then we display order information from file  
print("And your receipt for your records\n")
receipt_info = pd.read_excel(filename, sheet_name = 2)
print(receipt_info)

#Thanks user for order
print("\nThank you so much for shopping with us! Come again!\n")


