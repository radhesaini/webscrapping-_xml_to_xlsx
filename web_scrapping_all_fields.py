import csv
import requests
import xml.etree.ElementTree as ET
import pandas as pd
          
  
def parseXML(xmlfile):
  
    # create element tree object
    tree = ET.parse(xmlfile)
  
    # get root element
    root = tree.getroot()
  
    # create empty list for news items
    rowsitems = []
    df = pd.read_xml("./Input.xml", xpath="//VOUCHER/* | //TALLYMESSAGE")

    cols = df.dropna(how="all", axis=1).dropna(axis=0)
    fields = [f"{i}" for i in cols]
  
    # iterate news items
    for item in root.find("BODY").find("IMPORTDATA").find("REQUESTDATA").findall('TALLYMESSAGE'):
        # empty news dictionary
        rows = {}
        for child in item.findall("VOUCHER"):
            print(child.find("DATE").text)
            for field in fields:
                elem = child.find(field)
                if elem is not None and elem.text is not None:
                    rows[field] = elem.text
        # append rows dictionary to rows items list
        rowsitems.append(rows)
      
    # return news items list
    return rowsitems
  
  
def savetoXLSX(rowsitems, filename):
  
    # specifying the fields for csv file
    df = pd.read_xml("./Input.xml", xpath="//VOUCHER/* | //TALLYMESSAGE")
    cols = df.dropna(how="all", axis=1).dropna(axis=0)
    fields = [str(i) for i in cols]
    df = pd.DataFrame(rowsitems, columns=fields)
    df.to_excel ('./RadheOutput.xlsx', index = None, header=True)
  
      
def main():
    rowsitems = parseXML('./Input.xml')
    savetoXLSX(rowsitems, './topnews.csv')
      
      
if __name__ == "__main__":
    # calling main function
    main()