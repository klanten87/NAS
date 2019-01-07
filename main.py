from openpyxl import *


def SkrivSkylt(c,desc,a,skylt):
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


skylt = 27
motor = 5
egen = 4
wb = load_workbook('AS1.xlsx')
AS1 = wb['AS1']
Skyltlista = wb['Skyltlista']
Motor = wb['Provning motorer']
Egen = wb['Egenkontroll']

rows = AS1.max_row
#rows = 40
for i in range(7,rows):
    t = AS1.cell(row=i, column=1).value
    m = AS1.cell(row=i, column=6).value
    b = AS1.cell(row=i, column=2).font
    if b.b is True:
        egen +=1
    for y in range(8,14):
        e = AS1.cell(row=i, column=y).value
        if e is None:
            n = 0
        else:
            c = AS1.cell(row=i, column=2).value
            desc = AS1.cell(row=i, column=3).value
            a = AS1.cell(row=5, column=2).value
            SkrivEgenprovning(c,desc,a,egen)
            egen += 1
            break

    if t is None:
        #print("None at row "+str(i))
        n = 0
    elif "s" in t:
        c = AS1.cell(row=i, column=2).value
        desc = AS1.cell(row=i, column=3).value
        a = AS1.cell(row=5, column=2).value
        SkrivSkylt(c,desc,a,skylt)
        skylt += 1

    if m is None:
        # print("None at row "+str(i))
        n = 0
    elif "@" in m:
        c = AS1.cell(row=i, column=2).value
        desc = AS1.cell(row=i, column=3).value
        a = AS1.cell(row=i, column=6).value
        SkrivMotor(c, desc, a, motor)
        motor += 1



wb.save('AS1_genererad.xlsx')
print("Done")