from openpyxl import *
from openpyxl.worksheet.datavalidation import DataValidation

def SkrivSkylt(c,desc,a,skylt):
    #Testatar Github
    Skyltlista.cell(row=skylt, column=10).value = c.upper()
    Skyltlista.cell(row=skylt, column=11).value = desc.upper()
    Skyltlista.cell(row=skylt, column=12).value = a.upper()

def SkrivMotor(c,desc,a, motor):
    Motor.cell(row=motor, column=1).value = c
    Motor.cell(row=motor, column=2).value = desc
    a = a.split('@')
    Motor.cell(row=motor, column=5).value = a[-1]

def SkrivEgenprovning(c,desc,a,egen):
    Egen.cell(row=egen, column=1).value = c
    Egen.cell(row=egen, column=2).value = desc
    #Egen.cell(row=egen, column=5).value = a

def systemName(rows):       #KLAR
    system = "Not defined"
    for i in range(7,rows):
        bold = AS1.cell(row=i, column=2).font
        beteckning = AS1.cell(row=i, column=2).value
        if bold.b is True:
            system = beteckning
            #print(system)
        elif beteckning is None:
            n=0
        elif system in beteckning:
            allreadydone = 1
        else:
            AS1.cell(row=i, column=2).value = system + "-" + beteckning
            #print(i)



skylt = 27  #Startrad for kyltlista
motor = 5   #Startrad for Motordata
egen = 4    #Startrad for Egenprovningen
wb = load_workbook('AS1.xlsx')  #Laddar dokument
AS1 = wb['AS1']                 #Laddar flik AS1
Skyltlista = wb['Skyltlista']   #Laddar flik Skyltlista
Motor = wb['Provning motorer']  #Laddar flik Provning motorer
Egen = wb['Egenkontroll']       #Laddar flik Egenkontroll

rows = AS1.max_row              #Kollar vilken sista raden ar
#rows = 40

systemName(rows)

# for i in range(7,rows):
    # t = AS1.cell(row=i, column=1).value
    # m = AS1.cell(row=i, column=6).value
    # b = AS1.cell(row=i, column=2).font
    # if b.b is True:
    #     egen +=1
    # for y in range(8,14):
    #     e = AS1.cell(row=i, column=y).value
    #     if e is None:
    #         n = 0
    #     else:
    #         c = AS1.cell(row=i, column=2).value
    #         desc = AS1.cell(row=i, column=3).value
    #         a = AS1.cell(row=5, column=2).value
    #         SkrivEgenprovning(c,desc,a,egen)
    #         egen += 1
    #         break
    #
    # if t is None:
    #     #print("None at row "+str(i))
    #     n = 0
    # elif "s" in t:
    #     c = AS1.cell(row=i, column=2).value
    #     desc = AS1.cell(row=i, column=3).value
    #     a = AS1.cell(row=5, column=2).value
    #     SkrivSkylt(c,desc,a,skylt)
    #     skylt += 1
    #
    # if m is None:
    #     # print("None at row "+str(i))
    #     n = 0
    # elif "@" in m:
    #     c = AS1.cell(row=i, column=2).value
    #     desc = AS1.cell(row=i, column=3).value
    #     a = AS1.cell(row=i, column=6).value
    #     SkrivMotor(c, desc, a, motor)
    #     motor += 1



wb.save('AS1_genererad.xlsx')
print("Done")