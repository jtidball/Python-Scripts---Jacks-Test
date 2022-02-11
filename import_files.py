import os
import csv
from xml.dom.domreg import registered
import pandas as pd
from openpyxl import Workbook

# assign directory
directory = (r"V:\SFS Admin\Data Analysis - Usewicz\Operations\SFS\FCI at a Glance\Raw Data\202140")
 
# Create WorkBook 
workbook = Workbook()
sheet = workbook.active


#Save Workbook
workbook.save(filename="historical_ counts_by_day_202140.xlsx")
print("New File Created.")

# iterate over files in
# that directory
index = int(0)
for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    index += 1
    # checking if it is a file
    if os.path.isfile(f):
        print(f)
        
    # Read the File 
    data = pd.read_csv(f, low_memory=False)
    print("Reading Done")
    
    # Sort Values 
    # Remove out RES
    data = data[data.Z_CAMPUS != "RES"]
    
    # Filter by Term
    data.sort_values("EF_EARLIEST_SUBTERM", inplace = True)
    
    # Remove Duplicates keep first value
    data.drop_duplicates(subset=['A_ID'])
    print("Values Organized")
    
    # Count Registered
    registered_abj = sum(data.EF_EARLIEST_SUBTERM == "ABJ")
    registered_c = sum(data.EF_EARLIEST_SUBTERM == "C")
    registered_d = sum(data.EF_EARLIEST_SUBTERM == "D")
    
    # Count FCIs
    fci_abj = sum((data.EF_EARLIEST_SUBTERM == "ABJ") & (data.H_FCI_IND == 'Y'))
    fci_c = sum((data.EF_EARLIEST_SUBTERM == "C") & (data.H_FCI_IND == 'Y'))
    fci_d = sum((data.EF_EARLIEST_SUBTERM == "D") & (data.H_FCI_IND == 'Y'))

    # Count Count SAP
    sap_abj = sum(((data.EF_EARLIEST_SUBTERM == "ABJ") & (data.H_FCI_IND != 'Y')) & (data.U_SAP_IND == 'Y'))
    sap_c = sum(((data.EF_EARLIEST_SUBTERM == "C") & (data.H_FCI_IND != 'Y')) & (data.U_SAP_IND == 'Y'))
    sap_d = sum(((data.EF_EARLIEST_SUBTERM == "D") & (data.H_FCI_IND != 'Y')) & (data.U_SAP_IND == 'Y'))

    # Count Verifs Student 
    verif_abj = sum(((data.EF_EARLIEST_SUBTERM == "ABJ") & (data.H_FCI_IND != 'Y')) & (data.S_VERIF_INCOMPLETE == 'Y'))
    verif_c = sum(((data.EF_EARLIEST_SUBTERM == "C") & (data.H_FCI_IND != 'Y')) & (data.S_VERIF_INCOMPLETE == 'Y'))
    verif_d = sum(((data.EF_EARLIEST_SUBTERM == "D") & (data.H_FCI_IND != 'Y')) & (data.S_VERIF_INCOMPLETE == 'Y'))


    # Place Everything in it location
    print("Now Writing")
    reg1 = "B"+str(index+1)
    reg2 = "C"+str(index+1)
    reg3 = "D"+str(index+1)
    fci1 = "F"+str(index+1)
    fci2 = "G"+str(index+1)
    fci3 = "H"+str(index+1)
    sap1 = "J"+str(index+1)
    sap2 = "K"+str(index+1)
    sap3 = "L"+str(index+1)
    ver1 = "N"+str(index+1)
    ver2 = "O"+str(index+1)
    ver3 = "P"+str(index+1)

    sheet[reg1] = registered_abj
    sheet[reg2] = registered_c
    sheet[reg3] = registered_d
    sheet[fci1] = fci_abj
    sheet[fci2] = fci_c
    sheet[fci3] = fci_d
    sheet[sap1] = sap_abj
    sheet[sap2] = sap_c
    sheet[sap3] = sap_d
    sheet[ver1] = verif_abj
    sheet[ver2] = verif_c
    sheet[ver3] = verif_d

workbook.save(filename="historical_ counts_by_day_202140.xlsx")
