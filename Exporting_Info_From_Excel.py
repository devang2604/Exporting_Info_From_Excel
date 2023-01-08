import openpyxl

book = openpyxl.load_workbook('Time_Table_2.xlsx')    # Accessing the required excel file

o2 = book.get_sheet_by_name("Table 68")    # Opening the required sheet from the excel file

#--------------------------------------------------------------------
def mon800n900():
    a = o2['B5']
    b = o2['D5']
    if a.value == 'O2':
      print((o2['B5']).value,end=" "+(o2['C5']).value+"\n")
      print((o2['B6']).value,end=" "+(o2['C6']).value+"\n")
    if b.value == 'O2':
      print((o2['D5']).value,end=" "+(o2['E5']).value+"\n")
      print((o2['D6']).value,end=" "+(o2['E6']).value+"\n")
    else:
      print("No class")    

def mon1015n1115():
    m = o2['G5']
    n = o2['I5']
    if m.value == 'O2':
      print((o2['G5']).value,end=" "+(o2['H5']).value+"\n")
      print((o2['G6']).value,end=" "+(o2['H6']).value+"\n")
    if n.value == 'O2':
      print((o2['I5']).value,end=" "+(o2['J5']).value+"\n")
      print((o2['I6']).value,end=" "+(o2['J6']).value+"\n")
    else:
      print("No class") 

def mon115n215():
    m = o2['L5']
    n = o2['N5']
    if m.value == 'O2':
      print((o2['G5']).value,end=" "+(o2['H5']).value+"\n")
      print((o2['G6']).value,end=" "+(o2['H6']).value+"\n")
    if n.value == 'O2':
      print((o2['I5']).value,end=" "+(o2['J5']).value+"\n")
      print((o2['I6']).value,end=" "+(o2['J6']).value+"\n")
    else:
      print("No class") 

def mon330n430n530():
    m = o2['Q5']
    n = o2['S5']
    o = o2['U5']
    if m.value == 'O2':
      print((o2['Q5']).value,end=" "+(o2['R5']).value+"\n")
      print((o2['Q6']).value,end=" "+(o2['R6']).value+"\n")
    if n.value == 'O2':
      print((o2['S5']).value,end=" "+(o2['T5']).value+"\n")
      print((o2['S6']).value,end=" "+(o2['T6']).value+"\n")
    if o.value == 'O2':
      print((o2['U5']).value,end=" "+(o2['V5']).value+"\n")
      print((o2['U6']).value,end=" "+(o2['V6']).value+"\n")  
    else:
      print("No class") 
#--------------------------------------------------------------------
def tues800n900():
    m = o2['B11']
    n = o2['D11']
    if m.value == 'O2':
      print((o2['B11']).value,end=" "+(o2['C11']).value+"\n")
      print((o2['B12']).value,end=" "+(o2['C12']).value+"\n")
    if n.value == 'O2':
      print((o2['D11']).value,end=" "+(o2['E11']).value+"\n")
      print((o2['D12']).value,end=" "+(o2['E12']).value+"\n")
    else:
      print("No class") 

def tues1015n1115():
    m = o2['G11']
    n = o2['I11']
    if m.value == 'O2':
      print((o2['G11']).value,end=" "+(o2['H11']).value+"\n")
      print((o2['G12']).value,end=" "+(o2['H12']).value+"\n")
    if n.value == 'O2':
      print((o2['I11']).value,end=" "+(o2['J11']).value+"\n")
      print((o2['I12']).value,end=" "+(o2['J12']).value+"\n")
    else:
      print("No class") 

def tues115n215():
    m = o2['L11']
    n = o2['N11']
    if m.value == 'O2':
      print((o2['L11']).value,end=" "+(o2['M11']).value+"\n")
      print((o2['L12']).value,end=" "+(o2['M12']).value+"\n")
    if n.value == 'O2':
      print((o2['N11']).value,end=" "+(o2['O11']).value+"\n")
      print((o2['N12']).value,end=" "+(o2['O12']).value+"\n")
    else:
      print("No class") 

def tues330n430n530():
    m = o2['Q11']
    n = o2['S11']
    o = o2['U11']
    if m.value == 'O2':
      print((o2['Q11']).value,end=" "+(o2['R11']).value+"\n")
      print((o2['Q12']).value,end=" "+(o2['R12']).value+"\n")
    if n.value == 'O2':
      print((o2['S11']).value,end=" "+(o2['T11']).value+"\n")
      print((o2['S12']).value,end=" "+(o2['T12']).value+"\n")
    if o.value == 'O2':
      print((o2['U11']).value,end=" "+(o2['V11']).value+"\n")
      print((o2['U12']).value,end=" "+(o2['V12']).value+"\n")  
    else:
      print("No class")
#--------------------------------------------------------------------
def wed800n900():
    p = o2['B17']
    q = o2['D17']
    if p.value == "O2":
      print((o2['B17']).value,end=" "+(o2['C17']).value+"\n")
      print((o2['B18']).value,end=" "+(o2['C18']).value+"\n")
    if q.value == "O2":
      print((o2['D17']).value,end=" "+(o2['E17']).value+"\n")
      print((o2['D18']).value,end=" "+(o2['E18']).value+"\n")
    else:
      print("No class") 

def wed1015n1115():
    p = o2['G17']
    q = o2['I17']
    if p.value == "O2":
      print((o2['G17']).value,end=" "+(o2['H17']).value+"\n")
      print((o2['G18']).value,end=" "+(o2['H18']).value+"\n")
    if q.value == "O2":
      print((o2['I17']).value,end=" "+(o2['J17']).value+"\n")
      print((o2['I18']).value,end=" "+(o2['J18']).value+"\n")
    else:
      print("No class")

def wed115n215():
    m = o2['L17']
    n = o2['N17']
    if m.value == 'O2':
      print((o2['L17']).value,end=" "+(o2['M17']).value+"\n")
      print((o2['L18']).value,end=" "+(o2['M18']).value+"\n")
    if n.value == 'O2':
      print((o2['N17']).value,end=" "+(o2['O17']).value+"\n")
      print((o2['N18']).value,end=" "+(o2['O18']).value+"\n")
    else:
      print("No class")

def wed330n430n530():
    m = o2['Q17']
    n = o2['S17']
    o = o2['U17']
    if m.value == 'O2':
      print((o2['Q17']).value,end=" "+(o2['R17']).value+"\n")
      print((o2['Q18']).value,end=" "+(o2['R18']).value+"\n")
    if n.value == 'O2':
      print((o2['S17']).value,end=" "+(o2['T17']).value+"\n")
      print((o2['S18']).value,end=" "+(o2['T18']).value+"\n")
    if o.value == 'O2':
      print((o2['U17']).value,end=" "+(o2['V17']).value+"\n")
      print((o2['U18']).value,end=" "+(o2['V18']).value+"\n")  
    else:
      print("No class")
#--------------------------------------------------------------------
def thurs800n900():
    a = o2['B23']
    b = o2['D23']
    if a.value == 'O2':
      print((o2['B23']).value,end=" "+(o2['C23']).value+"\n")
      print((o2['B24']).value,end=" "+(o2['C24']).value+"\n")
    if b.value == 'O2':
      print((o2['D23']).value,end=" "+(o2['E23']).value+"\n")
      print((o2['D24']).value,end=" "+(o2['E24']).value+"\n")
    else:
      print("No class")

def thurs1015n1115():
    p = o2['G23']
    q = o2['I23']
    if p.value == "O2":
      print((o2['G23']).value,end=" "+(o2['H23']).value+"\n")
      print((o2['G24']).value,end=" "+(o2['H24']).value+"\n")
    if q.value == "O2":
      print((o2['I23']).value,end=" "+(o2['J23']).value+"\n")
      print((o2['I24']).value,end=" "+(o2['J24']).value+"\n")
    else:
      print("No class")

def thurs115n215():
    m = o2['L23']
    n = o2['N23']
    if m.value == 'O2':
      print((o2['L23']).value,end=" "+(o2['M23']).value+"\n")
      print((o2['L24']).value,end=" "+(o2['M24']).value+"\n")
    if n.value == 'O2':
      print((o2['N23']).value,end=" "+(o2['O23']).value+"\n")
      print((o2['N24']).value,end=" "+(o2['O24']).value+"\n")
    else:
      print("No class")

def thurs330n430n530():
    m = o2['Q23']
    n = o2['S23']
    o = o2['U23']
    if m.value == 'O2':
      print((o2['Q23']).value,end=" "+(o2['R23']).value+"\n")
      print((o2['Q24']).value,end=" "+(o2['R24']).value+"\n")
    if n.value == 'O2':
      print((o2['S23']).value,end=" "+(o2['T23']).value+"\n")
      print((o2['S24']).value,end=" "+(o2['T24']).value+"\n")
    if o.value == 'O2':
      print((o2['U23']).value,end=" "+(o2['V23']).value+"\n")
      print((o2['U24']).value,end=" "+(o2['V24']).value+"\n")  
    else:
      print("No class")
#--------------------------------------------------------------------
def fri800n900():
    a = o2['B29']
    b = o2['D29']
    if a.value == 'O2':
      print((o2['B29']).value,end=" "+(o2['C29']).value+"\n")
      print((o2['B30']).value,end=" "+(o2['C30']).value+"\n")
    if b.value == 'O2':
      print((o2['D29']).value,end=" "+(o2['E29']).value+"\n")
      print((o2['D30']).value,end=" "+(o2['E30']).value+"\n")
    else:
      print("No class")

def fri1015n1115():
    p = o2['G29']
    q = o2['I29']
    if p.value == "O2":
      print((o2['G29']).value,end=" "+(o2['H29']).value+"\n")
      print((o2['G30']).value,end=" "+(o2['H30']).value+"\n")
    if q.value == "O2":
      print((o2['I29']).value,end=" "+(o2['J29']).value+"\n")
      print((o2['I30']).value,end=" "+(o2['J30']).value+"\n")
    else:
      print("No class")

def fri115n215():
    m = o2['L29']
    n = o2['N29']
    if m.value == 'O2':
      print((o2['L29']).value,end=" "+(o2['M29']).value+"\n")
      print((o2['L30']).value,end=" "+(o2['M30']).value+"\n")
    if n.value == 'O2':
      print((o2['N29']).value,end=" "+(o2['O29']).value+"\n")
      print((o2['N30']).value,end=" "+(o2['O30']).value+"\n")
    else:
      print("No class")

def fri330n430n530():
    m = o2['Q29']
    n = o2['S29']
    o = o2['U29']
    if m.value == 'O2':
      print((o2['Q29']).value,end=" "+(o2['R29']).value+"\n")
      print((o2['Q30']).value,end=" "+(o2['R30']).value+"\n")
    if n.value == 'O2':
      print((o2['S29']).value,end=" "+(o2['T29']).value+"\n")
      print((o2['S30']).value,end=" "+(o2['T30']).value+"\n")
    if o.value == 'O2':
      print((o2['U29']).value,end=" "+(o2['V29']).value+"\n")
      print((o2['U30']).value,end=" "+(o2['V30']).value+"\n")  
    else:
      print("No class")
#--------------------------------------------------------------------
def sat800n900():
    a = o2['B35']
    b = o2['D35']
    if a.value == 'O2':
      print((o2['B35']).value,end=" "+(o2['C35']).value+"\n")
      print((o2['B36']).value,end=" "+(o2['C36']).value+"\n")
    if b.value == 'O2':
      print((o2['D35']).value,end=" "+(o2['E35']).value+"\n")
      print((o2['D36']).value,end=" "+(o2['E36']).value+"\n")
    else:
      print("No class")

def sat1015n1115():
    p = o2['G35']
    q = o2['I35']
    if p.value == "O2":
      print((o2['G35']).value,end=" "+(o2['H35']).value+"\n")
      print((o2['G36']).value,end=" "+(o2['H36']).value+"\n")
    if q.value == "O2":
      print((o2['I35']).value,end=" "+(o2['J35']).value+"\n")
      print((o2['I36']).value,end=" "+(o2['J36']).value+"\n")
    else:
      print("No class")
  
def sat115n215():
    m = o2['L35']
    n = o2['N35']
    if m.value == 'O2':
      print((o2['L35']).value,end=" "+(o2['M35']).value+"\n")
      print((o2['L36']).value,end=" "+(o2['M36']).value+"\n")
    if n.value == 'O2':
      print((o2['N35']).value,end=" "+(o2['O35']).value+"\n")
      print((o2['N36']).value,end=" "+(o2['O36']).value+"\n")
    else:
      print("No class")
      
def sat330n430n530():
    m = o2['Q35']
    n = o2['S35']
    o = o2['U35']
    if m.value == 'O2':
      print((o2['Q35']).value,end=" "+(o2['R35']).value+"\n")
      print((o2['Q36']).value,end=" "+(o2['R36']).value+"\n")
    if n.value == 'O2':
      print((o2['S35']).value,end=" "+(o2['T35']).value+"\n")
      print((o2['S36']).value,end=" "+(o2['T36']).value+"\n")
    if o.value == 'O2':
      print((o2['U35']).value,end=" "+(o2['V35']).value+"\n")
      print((o2['U36']).value,end=" "+(o2['V36']).value+"\n")  
    else:
      print("No class")
#--------------------------------------------------------------------
