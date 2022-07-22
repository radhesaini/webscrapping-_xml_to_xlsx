import csv
import requests
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime         
import copy 
  
def parseXML(xmlfile):
  
    # create element tree object
    tree = ET.parse(xmlfile)
  
    # get root element
    root = tree.getroot()
  
    # create empty list for news items
    rowsitems = []
    fields = ["DATE", "VOUCHERNUMBER", "VOUCHERTYPENAME",
    #  "BASICBUYERNAME",
     "REFERENCEDATE"]
  
    # iterate news items
    for item in root.find("BODY").find("IMPORTDATA").find("REQUESTDATA").findall('TALLYMESSAGE'):
        # empty news dictionary
        rows = {}
        for child in item.findall("VOUCHER"):
            if not child.find("VOUCHERNUMBER").text.isdigit():
                continue
            for field in fields:
                elem = child.find(field)
                manualy = {
                    "VOUCHERNUMBER": "Vch No.", 
                    "DATE": "Date", 
                    "PARTYLEDGERNAME": "Debator", 
                    "VOUCHERTYPENAME": "Vch Type",
                    "REFERENCEDATE": "Ref Date"
                    }
                if elem is not None and elem.text is not None:
                    if field == "DATE" or field == "REFERENCEDATE":
                        rows[manualy[field]] = datetime.strptime(elem.text, '%Y%m%d').strftime('%d-%m-%Y')
                    else:
                        rows[manualy[field]] = elem.text

            for nested_item in child.findall("ALLLEDGERENTRIES.LIST"):
                rows["Particulars"] = nested_item.find("LEDGERNAME").text
                rows["Debtor"] = nested_item.find("LEDGERNAME").text
                c_rows = copy.deepcopy(rows)
                rows["Amount"] = nested_item.find("AMOUNT").text
                rows["Ref Amount"] = "NA"
                rows["Ref Type"] = "NA"
                rows["Ref Date"] = "NA"
                rows["Ref No."] = "NA"
                rows["Amount Verified"] = "NA"
                rows["Transaction Type"] = "Other" if "Bank" in rows["Particulars"] else "Parent"
                sum_of_child = 0
                for new_childs in nested_item.findall("BILLALLOCATIONS.LIST"):
                    if new_childs.find("AMOUNT") is None or new_childs.find("AMOUNT").text is None:
                        break
                    c_rows["Ref No."] = new_childs.find("NAME").text
                    c_rows["Ref Type"] = new_childs.find("BILLTYPE").text
                    c_rows["Ref Amount"] = new_childs.find("AMOUNT").text
                    c_rows["Amount"] = "NA"
                    c_rows["Transaction Type"] = "child"  
                    c_rows["Amount"] = "NA"
                    rowsitems.append(c_rows)
                    c_rows = copy.deepcopy(c_rows)
                    sum_of_child += eval(c_rows["Ref Amount"]) if c_rows["Ref Amount"] else 0
                rows["Amount Verified"] = "Yes" if sum_of_child == eval(rows["Amount"]) else "NA"
                rowsitems.append(rows)
                rows = copy.deepcopy(rows)
        # append rows dictionary to rows items list
    # return news items list
    return rowsitems
  
  
def savetoXLSX(rowsitems, filename):
  
    # specifying the fields for csv fil
    fields = ["Date", "Transaction Type", "Vch No.", "Debator", "Vch Type", "Ref Date", "Ref Amount", "Amount",
     "Ref Type", "Ref No.", "Particulars", "Amount Verified"]
    df = pd.DataFrame(rowsitems, columns=fields)
    # df.to_csv("./new.csv")
    df.to_excel ('./RadheOutput.xlsx', index = None, header=True)
  
      
def main():
    rowsitems = parseXML('./Input.xml')
    savetoXLSX(rowsitems, './topnews.csv')
      
      
if __name__ == "__main__":
    # calling main function
    main()