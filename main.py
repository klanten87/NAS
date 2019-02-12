from openpyxl import load_workbook
from openpyxl.styles import Font

wb = load_workbook('AS1.xlsx')      #Laddar dokument
AS1 = wb['AS1']                     #Laddar flik AS1
rows = AS1.max_row                  #Kollar vilken sista raden ar
E = AS1.cell(row=2, column=5).value #Laddar vilka skript som ska köras.
Skyltlista = wb['Skyltlista']       # Laddar flik Skyltlista
Listor = wb['Listor']               # Laddar flik Listor
Avrop = wb['Avrop']                 # Laddar flik Avrop

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


def keywordSort10(rows):
    #Switch-funktion som kollar vad vi har och skickar sedan till rätt funktion.
    sort = {
        "Avluftsspjäll": Avluftsspjall,
        "Avluftstemperatur": Avluftstemperatur,
        "Brandspjäll": Brandspjall,
        "Differenstryckgivare": Differenstryckgivare,
        "Elmätare": Elmatare,
        "Expansionskärl": Expansionskarl,
        "Filtervakt": Filtervakt,
        "Flödesgivare": Flodesgivare,
        "Framledningstemperatur": Framledningstemperatur,
        "Frysskyddstemperatur": Frysskyddstemperatur,
        "Frånluftsfläkt": Franluftsflakt,
        "Frånluftsspjäll": Franluftsspjall,
        "Frånluftstemperatur": Franluftstemperatur,
        "Kallvattenmätare": Kallvattenmatare,
        "Kylmaskin": Kylmaskin,
        "Larm": Larm,
        "Ljusgivare": Ljusgivare,
        "Luftkvalitégivare": Luftkvalitegivare,
        "Pump": Pump,
        "Returledningstemperatur": Returledningstemperatur,
        "Rumstemperatur": Rumstemperatur,
        "Rökdetektor": Rokdetektor,
        "Spjällställdon": Spjallstalldon,
        "Tidkanal": Tidkanal,
        "Tilluftsfläkt": Tilluftsflakt,
        "Tilluftsspjäll": Tilluftsspjall,
        "Tilluftstemperatur": Tilluftstemperatur,
        "Tilluftstemperature.VVX": TilluftstemperaturVVX,
        "Tryckgivare": Tryckgivare,
        "Tryckvakt": Tryckvakt,
        "Uteluftsspjäll": Uteluftsspjall,
        "Uteluftskanalstemperatur": Uteluftskanalstemperatur,
        "Utomhustemperatur": Utomhustemperatur,
        "Varmvattenmätare": Varmvattenmatare,
        "VAV-spjäll": VAVspjall,
        "Ventilställdon": Ventilstalldon,
        "VVC-temperatur": VVCtemperatur,
        "Värmemängdsmätare": Varmemangdsmatare,
        "Värmeväxlare": Varmevaxlare,
        "Förlängdventilation": Forlangdventilation,
        "Serviceomkopplare": Serviceomkopplare,
        "Elbatteri": Elbatteri,
        "Fläktvakt": Flaktvakt,
    }
    material = 1
    best = 1
    larm = 1
    inst = 1
    for i in range(7,rows):
        system = AS1.cell(row=i, column=2).font
        if system.b is True:
            material += 1
            larm += 1
            inst += 1
            Listor.cell(row=material, column=7).value = AS1.cell(row=i, column=2).value
            Listor.cell(row=material, column=7).font = Font(bold=True)
            Listor.cell(row=larm, column=1).value = AS1.cell(row=i, column=2).value
            Listor.cell(row=larm, column=1).font = Font(bold=True)
            Listor.cell(row=inst, column=12).value = AS1.cell(row=i, column=2).value
            Listor.cell(row=inst, column=12).font = Font(bold=True)
            material += 1
            larm += 1
            inst += 1

        C = AS1.cell(row=i, column=3).value
        CBold = AS1.cell(row=i, column=3).font
        if C is not None and not "Reserv" in C and CBold.b is not True:
            if " " in C:
                C = C.split(" ")
                C = "".join(C)
            row_add = sort[C](i,material,larm,inst)
            row_add = list(row_add)
            material += int(row_add[0])
            larm += int(row_add[1])
            inst += int(row_add[2])




def Avluftsspjall(i,material,larm,inst):
    #Material
    #Beställning
    return("000")
def Avluftstemperatur(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    return("000")
def Brandspjall(i,material,larm,inst):
    #Material
    #Larm - Brandspjäll fel läge 2 m
    #Inställning - Motionering Söndag 23:55-23:59
    return("000")
def Differenstryckgivare(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    #Larm - Regleravvikelse +/-3KPa 15 m
    return("000")
def Elmatare(i,material,larm,inst):
    #Material
    return("000")
def Expansionskarl(i,material,larm,inst):
    #Larm - Låg tryck expansionkälr 2 m
    return("000")
def Filtervakt(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Smutsiga filter 15 m
    return("000")
def Flodesgivare(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    #Larm - Regleravvikelse 15 l/s 15 m
    #Inställning - Börvärde flöde Min/Max	15/30 l/s
    return("000")
def Framledningstemperatur(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    #Larm - Regleravvikelse +/-3°C 15 m
    #Inställning - Börvärde Utetemp X -20°C, -10°C, 0°C, 10°C, 20°C
    #Inställning - Börvärde framledning 65°C, 55°C, 47°, 32°C, 20°C
    return("000")
def Frysskyddstemperatur(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    #Larm - Frysvaktslarm 7°C A 3 s
    #Inställning - Varmhållning vid stillastående 20°C
    return("000")
def Franluftsflakt(i,material,larm,inst):
    #Material
    #Larm - Driftfel frånluftsfläkt 2 m
    return("000")
def Franluftsspjall(i,material,larm,inst):
    #Material
    #Beställning
    return("000")
def Franluftstemperatur(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    return("000")
def Kallvattenmatare(i,material,larm,inst):
    #Material
    #Beställning
    return("000")
def Kylmaskin(i,material,larm,inst):
    #Material
    #Larm - Summalarm Kylmaskin 30 s
    return("000")
def Larm(i,material,larm,inst):
    #Larm - %%Benämning%% 10 s
    return("000")
def Ljusgivare(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel
    #Inställning - Ljusnivå 80 Lux
    Listor.cell(row=material, column=7).value = AS1.cell(row=i,column=2).value
    Listor.cell(row=material, column=8).value = AS1.cell(row=i, column=3).value
    Listor.cell(row=material, column=9).value = AS1.cell(row=i, column=5).value
    Listor.cell(row=larm, column=1).value = AS1.cell(row=i, column=2).value
    Listor.cell(row=larm, column=2).value = "Givarfel"
    Listor.cell(row=larm, column=3).value = "B"
    Listor.cell(row=larm, column=5).value = "30s"
    Listor.cell(row=inst, column=12).value = AS1.cell(row=i, column=2).value
    Listor.cell(row=inst, column=13).value = "Ljusnivå gräns dag"
    Listor.cell(row=inst, column=14).value = "80 lux"
    return("111")


def Luftkvalitegivare(i,material,larm,inst):
    #Material
    #BEställning
    #Larm - Givarfel 30 s
    #Larm - Hög CO²-halt 900ppm 15 m
    #Inställning - Börvärde Luftkvalité 800 ppm
    return("000")
def Pump(i,material,larm,inst):
    #Material
    #Larm - Driftfel pump 2 m
    #Inställning - Pumpstart 7°C
    #Inställning - Pumpstopp 17°C
    return("000")
def Returledningstemperatur(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    return("000")
def Rumstemperatur(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    #Larm - Hög rumstemperatur 23°C 15m
    #Larm - Låg rumstemperatur 17°C 15m
    #Inställning - Börvärde rumstemperatur 21°C
    return("000")
def Rokdetektor(i,material,larm,inst):
    ##Material
    #Beställning
    #Larm - Utlöst rökdetektor 10 s
    #Larm - Servicelarm rökdetektor 10 s
    return("000")
def Spjallstalldon(i,material,larm,inst):
    #Material
    #Beställning
    return("000")
def Tidkanal(i,material,larm,inst):
    #Inställning - %%Benämning%% M-F 07:00-16:00
    return("000")
def Tilluftsflakt(i,material,larm,inst):
    #Material
    #Larm - Driftfel tilluftsfläkt 2 m
    return("000")
def Tilluftsspjall(i,material,larm,inst):
    #Material
    #Beställning
    return("000")
def Tilluftstemperatur(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    #Larm - Regleravvikelse +/-3°C 15m
    #Inställning - Börvärde tilluftstemp 19°C
    return("000")
def TilluftstemperaturVVX(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    return("000")
def Tryckgivare(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    #Larm - Regleravvikelse +/-10Pa 15m
    #Larm - Fläktvakt 50Pa 5m
    #Inställning - Börvärde tryck 150 Pa
    return("000")
def Tryckvakt(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Driftfel 30 s
    return("000")
def Flaktvakt(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Driftfel 30 s
    return("000")
def Uteluftsspjall(i,material,larm,inst):
    #Material
    #Beställning
    return("000")
def Uteluftskanalstemperatur(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    return("000")
def Utomhustemperatur(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    return("000")
def Varmvattenmatare(i,material,larm,inst):
    #Material
    #Beställning
    return("000")
def VAVspjall(i,material,larm,inst):
    #Material
    return("000")
def Ventilstalldon(i,material,larm,inst):
    #Material
    #Beställning
    return("000")
def VVCtemperatur(i,material,larm,inst):
    #Material
    #Beställning
    #Larm - Givarfel 30 s
    #Larm - Låg returtempertur 40°C 15 m
    return("000")
def Varmemangdsmatare(i,material,larm,inst):
    #Material
    #Beställning
    return("000")
def Varmevaxlare(i,material,larm,inst):
    #Larm - Summalarm 10s
    #Larm - Låg verkningsgrad 50% 30 m
    return("000")
def Forlangdventilation(i,material,larm,inst):
    #Material
    #Beställning
    #Instälning - Eftergångstid 1 timme
    return("000")
def Serviceomkopplare(i,material,larm,inst):
    #Larm - Serviceomkopplare i fel läge 10s
    return("000")
def Elbatteri(i,material,larm,inst):
    #Material
    #Larm - Överhettningslarm
    #Inställningsvärde - Eftergångstid 5m
    return("000")




# def skrivInstallning10(i,C,inst):
#     #Kollar nyckelord och skriver ut rätt fras och inställningsvärde
#     #Problem, vilken rad ska jag skriva på?
#
#     if "Frysskyddsgivare" in C:
#         Listor.cell(row=inst, column=11).value =  AS1.cell(row=i, column=2).value
#         Listor.cell(row=inst, column=12).value = "Varmhållning"


# def skrivBeställning10(row):
    #Hämtar beställningsunderlag från rad som fås från keywordSort och skriver till Beställningslistan
    #så kollar man om det finns någon som heter lika i fabrikat och typ, summerar och sedan lägger
    # sedan den i en array som man kollar
    #det första man gör i funktionen och breakar funktionen. Annan lev sorteras bort

# def skrivMaterial10(row):
    #
E = E.upper()
if "B" in E:
    systemName10(rows)
if "S" in E:
    SkrivSkylt10(rows)
if "M" in E:
    SkrivMotor10(rows)
if "E" in E:
    SkrivEgenprovning10(rows)
# if "L" in E or "A" in E:
keywordSort10(rows)


wb.save('AS1_genererad.xlsx')
print("Done")