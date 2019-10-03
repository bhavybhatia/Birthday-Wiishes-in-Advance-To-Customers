from datetime import date
from selenium import webdriver
import time
import xlrd

driver=webdriver.Chrome()#parenthesis contain the path of the wevdriver

driver.get("http://web.whatsapp.com")


def message(name,phone_number):
    phone_number=int(phone_number)
    time.sleep(10)
    link="https://web.whatsapp.com/send?phone=91"+str(phone_number)
    driver.get(link)
    time.sleep(5)
    text=driver.find_element_by_xpath("//div[@class='_3u328 copyable-text selectable-text']")
    text.send_keys("Hi *"+name+"*wish you a very *Happy Birthday* in advance, may your special day be full of happiness, fun and cheer!")
    print("Message to ",name," has been sent!")
    send_button=driver.find_element_by_xpath("//span[@data-icon='send']")
    send_button.click()

def leapYear(year):
    if((year%4)!=0):
        return 0
    else:
        if(year%100==0):
            if(year%400==0):
                return 1
            else:
                return 0
        else:
            return 1



loc=("D:\python_Workspace\Advance Birthday Wish To Customers\sample_file.xlsx")#parenthesis contain the path for the location of the xlsx file
wb=xlrd.open_workbook(loc)
sheet=wb.sheet_by_index(0)

a=list()
b=list()
for i in range(sheet.nrows):
    a.append(sheet.cell_value(i,1))

today=date.today()
f_date=date(1900,1,1)
final=((today - f_date).days+4)


y=date.today().year

for i in range(50):
    if leapYear(y):
        final-=366
    else:
        final-=365
    y-=1
    for j in range(sheet.nrows):
        repetition=0 #to check whether the same name is repeated earlier for the same date
        if a[j]==final:
            for k in range(len(b)):
                if(b[k]==sheet.cell_value(j,0)):
                    repetition=1
            if(repetition!=1):
                message(sheet.cell_value(j,0),sheet.cell_value(j,2))
                b.append(sheet.cell_value(j,0))
input()
