from openpyxl import load_workbook
from openpyxl.styles import Font

wb = load_workbook('AS1.xlsx')      #Laddar dokument
AS1 = wb['AS1']                     #Laddar flik AS1
rows = AS1.max_row                  #Kollar vilken sista raden ar
E = AS1.cell(row=2, column=5).value #Laddar vilka skript som ska köras.
Skyltlista = wb['Skyltlista']       # Laddar flik Skyltlista
def systemName10(rows):       #KLAR
    system = "Not defined"
    for i in range(7,rows):
        bold = AS1.cell(row=i, column=2).font
        beteckning = AS1.cell(row=i, column=2).value
        if bold.b is True:
            system =  AS1.cell(row=i, column=2).value
        elif beteckning is not None and system not in beteckning:
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


def writeSign(i, skylt, typ):   #KLAR tillhör SkrivSkylt10
    Skyltlista.cell(row=skylt, column=2).value = "TYP " + typ
    Skyltlista.cell(row=skylt, column=8).value = "1"
    Skyltlista.cell(row=skylt, column=10).value = AS1.cell(row=i, column=2).value.upper()
    Skyltlista.cell(row=skylt, column=11).value = AS1.cell(row=i, column=3).value.upper()
    if "200" in typ:
        Skyltlista.cell(row=skylt, column=12).value = AS1.cell(row=5, column=2).value.upper()






def SkrivMotor10(rows): #KLAR
    rad = 5  # Startrad for Motordata
    Motor = wb['Provning motorer']  # Laddar flik Provning motorer
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

def SkrivEgenprovning10(rows):  #KLAR
    Egen = wb['Egenkontroll']  # Laddar flik Egenkontroll
    rad = 4 #Startrad
    switch =['Larm','Tidkanal','Elmätare']
    for i in range(7,rows):
        system = AS1.cell(row=i, column=2).font
        if system.b is True:
            rad += 1
            Egen.cell(row=rad, column=1).value = AS1.cell(row=i, column=2).value
            Egen.cell(row=rad, column=1).font = Font(bold=True)
            rad += 1
            continue
        B = AS1.cell(row=i, column=2).value
        C = AS1.cell(row=i, column=3).value
        if B is not None and C is not None and any(x in C for x in switch):
            E = AS1.cell(row=i, column=5).value
            if E is None:
                Egen.cell(row=rad, column=1).value = AS1.cell(row=i, column=2).value
                Egen.cell(row=rad, column=2).value = AS1.cell(row=i, column=3).value
                rad += 1
                continue
            Egen.cell(row=rad, column=1).value = AS1.cell(row=i, column=2).value
            Egen.cell(row=rad, column=2).value = AS1.cell(row=i, column=5).value
            rad += 1
        elif B is not None:
            Egen.cell(row=rad, column=1).value = AS1.cell(row=i, column=2).value
            Egen.cell(row=rad, column=2).value = AS1.cell(row=i, column=3).value
            rad += 1


# def keywordSort10(rows):
    #Tar nyckelordet och ser vad den ska generera fram.
    #Skickar sedan rad till Material,Bestallning, Larm och Installningsvarde

# def skrivBeställning10(row):
    #Hämtar beställningsunderlag från rad som fås från keywordSort och skriver till Beställningslistan
    #så kollar man om det finns någon som heter lika i fabrikat och typ, summerar och sedan lägger
    # sedan den i en array som man kollar
    #det första man gör i funktionen och breakar funktionen. Annan lev sorteras bort

# def skrivMaterial10(row):
    #




if "B" in E:
    systemName10(rows)
if "S" in E:
    SkrivSkylt10(rows)
if "M" in E:
    SkrivMotor10(rows)
if "E" in E:
    SkrivEgenprovning10(rows)


wb.save('AS1_genererad.xlsx')
print("Done")