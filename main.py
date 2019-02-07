from openpyxl import *

def systemName10(rows):       #KLAR
    system = "Not defined"
    for i in range(7,rows):
        bold = AS1.cell(row=i, column=2).font
        beteckning = AS1.cell(row=i, column=2).value
        if bold.b is True:
            system =  AS1.cell(row=i, column=2).value
            #print(system)
        elif beteckning is None:
            n=0
        elif system in beteckning:
            allreadydone = 1
        else:
            AS1.cell(row=i, column=2).value = system + "-" + beteckning


def SkrivSkylt10(rows): #KLAR
    skylt = 27  # Startrad for skyltlista
    nosign =['Elmätare','Tidkanal','Larm']
    for i in range(7,rows): #Går egenom alla rader från rad 7 till sista raden.
        keyword = AS1.cell(row=i, column=3).value
        A = AS1.cell(row=i, column=1).value  #Plockar fram om någon skrivit något i Optionkolumn.
        C = AS1.cell(row=i, column=3).value
        D = AS1.cell(row=i, column=4).value
        for col in range(9,18): #Kollar igenom kolumn I-Q om det är någon I/O-kopplad till komponenten.
            if keyword is not None and any(x in keyword for x in nosign):
                break
            V = AS1.cell(row=i, column=col).value
            if V is not None:
                writeSign(i,skylt,"200")
                skylt += 1
                break
            if C is not None and "Rökdetektor" in C:
                writeSign(i, skylt, "200")
                skylt += 1
                break
        if D is not None and "Siox" in D:
            B = AS1.cell(row=i, column=2).value
            split = B.split("-")
            del split[-1]
            S = "-".join(split) + "-BSC" + str(AS1.cell(row=i, column=8).value)
            Skyltlista.cell(row=skylt, column=2).value = "TYP 200"
            Skyltlista.cell(row=skylt, column=8).value = "1"
            Skyltlista.cell(row=skylt, column=10).value = S
            Skyltlista.cell(row=skylt, column=11).value = "BRANDSPJÄLLSCENTRAL"
            Skyltlista.cell(row=skylt, column=12).value = AS1.cell(row=5, column=2).value.upper()
            skylt += 1
        if A is not None and "s" in A:
            writeSign(i, skylt, "100")
            skylt += 1


def writeSign(i, skylt, typ):   #KLAR
    Skyltlista.cell(row=skylt, column=2).value = "TYP " + typ
    Skyltlista.cell(row=skylt, column=8).value = "1"
    Skyltlista.cell(row=skylt, column=10).value = AS1.cell(row=i, column=2).value.upper()
    Skyltlista.cell(row=skylt, column=11).value = AS1.cell(row=i, column=3).value.upper()
    if "200" in typ:
        Skyltlista.cell(row=skylt, column=12).value = AS1.cell(row=5, column=2).value.upper()






def SkrivMotor10(rows): #KLAR
    rad = 5  # Startrad for Motordata
    motor=['Frånluftsfläkt','Pump','Tilluftsfläkt']
    for i in range(7,rows):
        keyword = AS1.cell(row=i, column=3).value
        if keyword is not None and any(x in keyword for x in motor):
            F = AS1.cell(row=i, column=6).value
            Motor.cell(row=rad, column=1).value = AS1.cell(row=i, column=2).value
            Motor.cell(row=rad, column=2).value = AS1.cell(row=i, column=3).value
            F = F.split('@')
            Motor.cell(row=rad, column=5).value = F[-1]
            rad += 1

def SkrivEgenprovning(c,desc,a,egen):
    rad = 4 #Startrad
    Egen.cell(row=egen, column=1).value = c
    Egen.cell(row=egen, column=2).value = desc
    #Egen.cell(row=egen, column=5).value = a







egen = 4    #Startrad for Egenprovningen
wb = load_workbook('AS1.xlsx')  #Laddar dokument
AS1 = wb['AS1']                 #Laddar flik AS1
Skyltlista = wb['Skyltlista']   #Laddar flik Skyltlista
Motor = wb['Provning motorer']  #Laddar flik Provning motorer
Egen = wb['Egenkontroll']       #Laddar flik Egenkontroll

rows = AS1.max_row              #Kollar vilken sista raden ar
#rows = 40

systemName10(rows)
SkrivSkylt10(rows)
SkrivMotor10(rows)
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