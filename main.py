import os
import shutil

import numpy as np
import openpyxl
import pandas as pd

from datetime import datetime
from os import path
from openpyxl import load_workbook
import re

from Utils import Utils


def companyNameLookUp(companyName):
    companyDict = {
        'ALL': 'All Seasons Realty Corp',
        'APL': 'Allianz-PNB Life Insurance, Inc. (APLII)',
        'ABI': 'Asia Brewery, Inc. (ABI), Subsidiaries',
        'BHC': 'Basic Holdings Corp.',
        'CPH': 'Century Park Hotel',
        # "EPP": "Eton Properties Philippines, Inc. (EPPI), Subsidiaries",
        "EPP": "Eton Properties Philippines, Inc. (EPPI), Subsidiaries",
        'FFI': 'Foremost Farms, Inc.',
        'FTC': 'Fortune Tobacco Corp.',
        'GDC': 'Grandspan Development Corp.',
        'HII': 'Himmel Industries, Inc.',
        'LRC': 'Landcom Realty Corp.',
        'LTG': 'LT Group, Inc. (Parent Company)',
        'DIR': 'LTGC Directors',
        'MAC': 'MacroAsia Corp., Subsidiaries and Affiliates',
        'PAL': 'Philippine Airlines, Inc. (PAL), Subsidiaries and Affiliates',
        'PNB': 'Philippine National Bank (PNB), Subsidiaries',
        'PMI': 'PMFTC Inc.',
        'RAP': 'Rapid Movers & Forwarders, Inc.',
        'TYK': 'Tan Yan Kee Foundation, Inc. (TYKFI)',
        'TDI': 'Tanduay Distillers, Inc. (TDI), Subsidiaries',
        'CHI': 'Charter House Inc.',
        'SPV': 'SPV-AMC Group',
        'TMC': 'Topkick Movers Corporation',
        'UNI': 'University of the East (UE)',
        'UER': 'University of the East Ramon Magsaysay Memorial Medical Center (UERMMMC)',
        'VMC': 'Victorias Milling Company, Inc. (VMC)',
        'ZHI': 'Zebra Holdings, Inc.',
        'STN': 'Sabre Travel Network Phils., Inc.',
    }
    company_Code = ""
    for key, value in companyDict.items():
        if companyName.strip() == value:
            company_Code = key

    return company_Code


def createDirectories(rp):
    inPath = os.path.join(rp, "in")
    outPath = os.path.join(rp, "out")
    excelSplitPath = os.path.join(rp, "excelSplit")

    if not path.exists(inPath):
        os.mkdir(inPath)

    if not path.exists(outPath):
        os.mkdir(outPath)

    if not path.exists(excelSplitPath):
        os.mkdir(excelSplitPath)

    pass


def checkCountEqualsQty(count, qty):
    if not len(count.split(",")) == int(qty):
        return False

    return True


def checkCtrlNumFormat(ctrlNumber, companyName):
    companyCode = companyNameLookUp(companyName)

    regex = "\\b" + companyCode + "_[A-Za-z0-9-]+[_]{1}[M|C]{1}[1-9]{1}[0-9]?$"
    regexPalex = "\\bPALEX_[A-Za-z0-9-]+[_]{1}[M|C]{1}[1-9]{1}[0-9]?$"

    if companyCode == 'PAL':
        if not re.match(regex, ctrlNumber):
            return False
    else:
        if companyCode == 'PALEX':
            if not re.match(regexPalex, ctrlNumber):
                return False
        else:
            if not re.match(regex, ctrlNumber):
                return False

    return True


def getEmpNumberFromCtrlNumber(x):

    print(x)

    pass


def getDataFromExcel(rp, masterFile):
    # concat rootpath with "in" folder dynamically
    inPath = os.path.join(rp, "in")
    excelSplitPath = os.path.join(rp, "excelSplit")
    masterFilePath = os.path.join(inPath, masterFile)

    # print(masterFilePath)
    df = pd.read_excel(masterFilePath,
                       header=1,
                       dtype={'ID': str,
                              'Control Number': str,
                              'Company Name': str,
                              'Age': str},
                       na_filter=False)

    # Validate Control Number Format
    df['Is Duplicate CtrlNumber'] = df.duplicated(subset="Control Number", keep='first')
    df['Is Valid CtrlNumber Format'] = df.apply(lambda x: checkCtrlNumFormat(x['Control Number'],
                                                                  x['Company Name']), axis=1)

    # Validate Email
    df['']
    df['Employee Number'] = df['Control Number'].str.split('_').str[1]

    df = df[['ID', 'Employee Number', 'Control Number', 'Company Name', 'Is Duplicate', 'Is Valid Format']]

    groups = df.groupby("Company Name")

    for i, company in groups:
        filename = os.path.join(excelSplitPath, i + "_EMPHH.xlsx")
        company.to_excel(filename)

    pass


def generateErrorLog(errMsg, companyCode, arg):
    util = Utils()
    outPath = os.path.join(rootPath, "out")

    # write
    if len(errMsg):
        util.createSubCompanyFolder(companyCode, outPath)
        f = open(
            outPath + "/" + companyCode + "/" + companyCode + "_" + arg + "_err_log_" + dateTime + ".txt",
            "w")
        for err in errMsg:
            f.writelines(err + "\n")

        errMsg.clear()
    pass


def getError_IsCtrlNumDuplicate(filename, FilePath):
    companyName = filename.split('_EMPHH')[0]

    print(companyName + " running IsCtrlNumDuplicate...")

    df = pd.read_excel(FilePath, dtype={'ID': str, 'Company Name': str,
                                        'Employee Number': str, 'Control Number': str}, na_filter=False)

    errMsg = []
    companyCode = companyNameLookUp(companyName)

    dupCtrlNumber = df.loc[df['Is Duplicate'] == True]

    # remove Duplicate to convert into List
    noDup = dupCtrlNumber.drop_duplicates(subset=['Control Number'])

    for ctrlNum in noDup['Control Number'].tolist():

        id = dupCtrlNumber.loc[dupCtrlNumber['Control Number'] == ctrlNum]

        # convert List to String
        new_list = [str(i) for i in id['ID'].tolist()]
        idStr = ', '.join(new_list)

        if not len(new_list) <= 1:
            errMsg.append("Error: ID[ " + idStr + " ] - " + ctrlNum + "")
            # print("Error: ID[ " + idStr + " ] - " + ctrlNum + "")

    generateErrorLog(errMsg, companyCode, "Duplicate_Control_Number")
    pass


def getError_IsCtrlNumFormatValid(filename, FilePath):
    companyName = filename.split('_EMPHH')[0]

    print(companyName + " running IsCtrlNumFormatValid...")

    df = pd.read_excel(FilePath, dtype={'ID': str, 'Company Name': str,
                                        'Employee Number': str, 'Control Number': str}, na_filter=False)

    errMsg = []
    companyCode = companyNameLookUp(companyName)

    df = df.loc[df['Is Valid Format'] == False]

    for j, row in df.iterrows():
        errMsg.append("Error: ID[ " + str(row['ID']) + " ] - " + row['Control Number'])
        # print("Error: ID[ " + str(row['ID']) + " ] - " + row['Control Number'])

    generateErrorLog(errMsg, companyCode, "Invalid_Control_Number")

    pass


def getErrLog(arrFilenames):
    for filename in arrFilenames:
        FilePath = os.path.join(excelLogPath, filename)
        if filename == "Asia Brewery, Inc. (ABI), Subsidiaries_EMPHH.xlsx":
            getError_IsCtrlNumDuplicate(filename, FilePath)
            getError_IsCtrlNumFormatValid(filename, FilePath)

        # if not filename == ".DS_Store":
        #     getError_IsCtrlNumDuplicate(filename, FilePath)

    pass


if __name__ == '__main__':
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', None)

    today = datetime.today()
    dateTime = today.strftime("%m_%d_%y_%H%M%S")

    rootPath = r"/Users/Ran/Documents/Vaccine/LTG_ControlNumber_Validation"
    excelLogPath = os.path.join(rootPath, "excelSplit")
    HHExcelFileName = 'HHLTGC_CEIRMasterlist.xlsx'

    createDirectories(rootPath)

    print("=============================================")
    print(" Running Control Number Validation Script... ")
    print("=============================================")

    getDataFromExcel(rootPath, HHExcelFileName)

    # Get all filenames in excelLogPath
    getErrLog(os.listdir(excelLogPath))
