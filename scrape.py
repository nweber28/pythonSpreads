import pandas as pd
import datetime

wbName = 'phonePoolAudit.xlsx'
dfName = "SMH_DP"
df = pd.read_excel(wbName, sheet_name=dfName)
print(dfName)


# filter out ignores and cli serv phones ----------------------------------------------------------------
noIgnores = df[~df["Description"].str.contains("IGNORE", case=False)]
noIgnores = noIgnores[~noIgnores["Phone Type"].str.contains("Cisco Unified Client Services Framework", case=False)]

lastReg = noIgnores[~noIgnores["Last Registered"].str.contains("Never", na=False, case=False)] # filters out nevers
lastReg = lastReg[~lastReg["Last Registered"].str.contains("Now", na=False, case=False)] # filters out nows
valDatesOnly = lastReg[lastReg["Last Registered"] < datetime.datetime(2023,1,1)] # contains dates < 2023

nowsOnly = noIgnores[noIgnores["Last Registered"].str.contains("Now", na=False, case=False)] # contains nows only

# combine valid dates and now columns for accurate lastReg
lastReg = valDatesOnly._append(nowsOnly)

# Wireless models
colC = len(lastReg[lastReg["Phone Type"].str.contains("7925")]) + len(lastReg[lastReg["Phone Type"].str.contains("8821")])


# Conference models
colD = len(lastReg[lastReg["Phone Type"].str.contains("8831")]) + len(lastReg[lastReg["Phone Type"].str.contains("7937")])


# Fax and analog
faxAndAnalog = lastReg[lastReg["Phone Type"].str.contains("Analog Phone", case=False)]
fax = lastReg[lastReg["Description"].str.contains("Fax", case=False)]
analog = len(faxAndAnalog) - len(fax)


# IDF/MDF Phones
noAnalog = lastReg[~lastReg["Phone Type"].str.contains("Analog Phone", case=False)] # removes analogs
dfcount = len(noAnalog[noAnalog["Description"].str.contains("MDF", case=False)]) + len(noAnalog[noAnalog["Description"].str.contains("IDF", case=False)])


# Last Active - phone has not made 2023 call
lastActive = len(lastReg[lastReg["Last Active"] < '2023-01-01 00:00:00'])



# Print Statements
print("Wireless")
print(colC)

print("Conference")
print(colD)

print("MDF/IDF")
print(dfcount)

print("Analog")
print(analog)
print("Fax")
print(len(fax))

print("Last Active")
print(lastActive)

print("Last Registered")
print(len(lastReg))


# Combine all sheets into another sheet
import xlsxwriter

writer = pd.ExcelWriter("allCombined.xlsx", engine="xlsxwriter")
CombinedData = pd.DataFrame()

for sht in pd.ExcelFile(wbName).sheet_names:
    datFr = pd.read_excel(wbName, sheet_name= sht)
    CombinedData = CombinedData._append(datFr)

CombinedData.to_excel(writer, sheet_name= "AllData", index = False)
writer._save()

# print(pd.ExcelFile(wbName).sheet_names)

