#!/usr/bin/python
#author: Dr. Jorge Merino
#organisation: University of Cambridge. Institute for Manufacturing.
#date: 2020.04.29
#license: GPL-3.0 (GNU General Public License version 3): https://opensource.org/licenses/GPL-3.0

import sys
import csv
import pandas as pd
import xlrd

def readcsvPandas(filename, cols, ids):
    data=dict()
    if cols!=None and ids==None:
        data=pd.read_csv( filename, usecols=cols, header=0 ).to_dict(orient='records')
    elif cols!=None and ids!=None:
        atts=list()
        atts.extend(cols)
        for id in ids:
            if not (id in atts):
                atts.append(id)
        data=pd.read_csv( filename, usecols=atts, header=0 ).to_dict(orient='records')
    else:
        data=pd.read_csv( filename, header=0 ).to_dict(orient='records')
    #print(data)

    index_id=0
    idMatching = dict()

    for key in data:
        if(ids!=None):
            key["id_new"] = getrowID(idMatching, key, ids)
            if(key["id_new"]==-1):
                key["id_new"] = index_id
                storerowID(idMatching,ids,key,index_id)
                index_id+=1
            for id in ids:
                key.pop(id, None)
        else:
            key["id_new"]=index_id
            index_id+=1
    return data


def getrowID(dict, key, ids):
    index=-1
    for entry in dict:
        thisIsIt=True
        for id in ids:
            if(dict[entry][id] != key[id]):
                thisIsIt = False
        if thisIsIt:
            index = entry

    return index

def storerowID(dict,ids,key,index):
    rowID = {}
    for id in ids:
        rowID[id]=key[id]
    dict[index]=rowID


def writecsvPandas(df, filename):
    df.to_csv(filename, index=None)

def readcsv(filename, cols, ids):
    with open(filename, 'r') as f:
        reader = csv.DictReader(f)
        d=[]

        index_id=0
        idMatching = dict()
        for key in reader:
            #print(row)
            if(ids!=None):
                key["id_new"] = getrowID(idMatching, key, ids)
                if(key["id_new"]==-1):
                    key["id_new"] = index_id
                    storerowID(idMatching,ids,key,index_id)
                    index_id+=1
                for id in ids:
                    key.pop(id, None)
            else:
                key["id_new"]=index_id
                index_id+=1
            if cols==None:
                d.append(key)
            else:
                newkey=dict()
                for col in cols:
                    if ((ids!=None and not col in ids) or ids==None):
                        newkey[col]=key[col]
                newkey["id_new"]=key["id_new"]
                d.append(newkey)

        return d


def writecsv(filename, content):
    with open(filename, 'w', newline="") as f:
        writer = csv.DictWriter(f, content[0].keys())
        writer.writeheader()
        writer.writerows(content)


def readxlsPandas(filename, cols, ids):
    data=dict()
    if cols!=None and ids==None:
        data=pd.read_excel( filename, usecols=cols, header=0 ).to_dict(orient='records')
    elif cols!=None and ids!=None:
        atts=list()
        atts.extend(cols)
        for id in ids:
            if not (id in atts):
                atts.append(id)
        data=pd.read_excel( filename, usecols=atts, header=0 ).to_dict(orient='records')
    else:
        data=pd.read_excel( filename, header=0 ).to_dict(orient='records')
    #print(data)

    index_id=0
    idMatching = dict()

    for key in data:
        if(ids!=None):
            key["id_new"] = getrowID(idMatching, key, ids)
            if(key["id_new"]==-1):
                key["id_new"] = index_id
                storerowID(idMatching,ids,key,index_id)
                index_id+=1
            for id in ids:
                key.pop(id, None)
        else:
            key["id_new"]=index_id
            index_id+=1
    return data

def excel2dict(excel_sheet):
    headers = []
    for col in range(excel_sheet.ncols):
        headers.append( excel_sheet.cell_value(0,col) )
    # transform the workbook to a list of dictionaries
    data =[]
    for row in range(1, excel_sheet.nrows):
        elm = {}
        for col in range(excel_sheet.ncols):
            elm[headers[col]]=excel_sheet.cell_value(row,col)
        data.append(elm)
    return data

def readxls(filename, cols, ids):
    wb = xlrd.open_workbook(filename)
    sheet = wb.sheet_by_index(0)
    data = excel2dict(sheet)
    #print(data)
    d=[]

    index_id=0
    idMatching = dict()
    for key in data:
        if(ids!=None):
            key["id_new"] = getrowID(idMatching, key, ids)
            if(key["id_new"]==-1):
                key["id_new"] = index_id
                storerowID(idMatching,ids,key,index_id)
                index_id+=1
            for id in ids:
                key.pop(id, None)
        else:
            key["id_new"]=index_id
            index_id+=1
        if cols==None:
            d.append(key)
        else:
            newkey=dict()
            for col in cols:
                if ((ids!=None and not col in ids) or ids==None):
                    newkey[col]=key[col]
            newkey["id_new"]=key["id_new"]
            d.append(newkey)

    return d

def getFileFormat(filename):
    parts=filename.split(".")
    format=parts[len(parts)-1]
    return format

def parseAttributes(argv):
    cols=None
    ids=None
    filename=None
    help=False
    #print(argv)
    i=0
    while i < len(argv):
        if argv[i]=="--ids":
            if len(argv)>(i+1):
                ids=argv[i+1].split(",")
                i+=1
            else:
                help=True
                i=len(argv)
        elif argv[i] == "--columns":
                if len(argv)>(i+1):
                    cols=argv[i+1].split(",")
                    i+=1
                else:
                    help=True
                    i=len(argv)
        elif argv[i] == "--help":
            help=True
            i=len(argv)
        else:
            #print("Here: " + argv[i])
            if(getFileFormat(argv[i]) == "xls" or getFileFormat(argv[i]) == "xlsx" or getFileFormat(argv[i]) == "csv"):
                filename=argv[i]
            else:
                help=True
                i=len(argv)
        i+=1

    if filename == None:
        help=True
    return filename, cols, ids, help

def printHelp():
    print("Error using the script")
    print("Correct use: python prepare_data.py [options] filename")
    print("Requirements: Python 3.7, xlrd library. Optional, but recommended: pandas library")
    print("Accepted formats: XLS(x) and CSV. In case the input is an Excel file(.xls, .xlsx), the script assumes it has only one sheet.")
    print("Options:")
    print("\t--ids <ids> \t\tNames of the columns that identify the staff in the dataset. IDs must be separated by commas \",\"")
    print("\t--columns <columns> \tNames of the columns to be shared from the entire dataset. Columns must be separated by commas \",\"")
    print("\t--help \t\t\tShows this help")
    print("Example: python prepare_data.py --ids staffid,name,surname --columns \"Grade Type\",Group,Reason,Start,End,\"Total Duration\" myfile.csv")

def main():
    filename, cols, ids, help = parseAttributes(sys.argv[1:])
    #print("Filename: " + filename)
    #print("Columns: ", cols)
    #print("IDs: ", ids)
    #print(help)
    if help:
        printHelp()
    else:
        try:
            print("Reading " + filename)
            print("Columns: ", cols)
            print("IDs: ", ids)

            if(getFileFormat(filename) == "xls" or getFileFormat(filename) == "xlsx"):
                data=readxlsPandas(filename, cols, ids)
            else:
                data=readcsvPandas(filename,cols, ids)
            writecsvPandas(pd.DataFrame.from_dict(data, orient='columns'), "output.csv")
            print("Output file: output.csv")
        except Exception as e:
            print ("Could not read data input: ", e)
            print ("Trying method 2")
            try:
                if(getFileFormat(filename) == "xls" or getFileFormat(filename) == "xlsx"):
                    data=readxls(filename, cols, ids)
                else:
                    data=readcsv(filename, cols, ids)
                #print(data)
                writecsv("output.csv", data)
                print("Output file: output.csv")
            except Exception as e2:
                print ("Could not read data input: ", e2)

if __name__ == "__main__":
    main()
