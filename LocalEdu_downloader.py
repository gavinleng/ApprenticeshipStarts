__author__ = 'G'

import sys
sys.path.append('../harvesterlib')

import pandas as pd
import argparse
import json

import now
import openurl
import datasave as dsave


# url = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/382956/apprenticeships-starts-by-geography-level-and-age.xls"
# output_path = "tempAppStarts.csv"
# sheet = "Local Education Authority"
# required_indicators = ["2005/06", "2006/07", "2007/08", "2008/09", "2009/10", "2010/11", "2011/12", "2012/13", "2013/14", "2014/15"]


def download(url, sheet, reqFields, outPath, col, keyCol, digitCheckCol, noDigitRemoveFields):
    yearReq = reqFields
    dName = outPath

    # open url
    socket = openurl.openurl(url, logfile, errfile)

    # operate this excel file
    logfile.write(str(now.now()) + ' excel file loading\n')
    print('excel file loading------')
    xd = pd.ExcelFile(socket)
    df = xd.parse(sheet)

    # indicator checking
    logfile.write(str(now.now()) + ' indicator checking\n')
    print('indicator checking------')
    for i in range(df.shape[0]):
        yearCol = []
        for k in yearReq:
            k_asked = k
            for j in range(df.shape[1]):
                if str(k_asked) in str(df.iloc[i, j]):
                    yearCol.append(j)
                    restartIndex = i + 1

        if len(yearCol) == len(yearReq):
            break

    if len(yearCol) != len(yearReq):
        errfile.write(str(now.now()) + " Requested data " + str(yearReq).strip(
            '[]') + " don't match the excel file. Please check the file at: " + str(url) + " . End progress\n")
        logfile.write(str(now.now()) + ' error and end progress\n')
        sys.exit("Requested data " + str(yearReq).strip(
            '[]') + " don't match the excel file. Please check the file at: " + url)

    yearCol.append(df.shape[1])

    for i in range(restartIndex, df.shape[0]):
        kk = []
        k_asked = "All Apprenticeships"
        for k in range(len(yearCol)-1):
            for j in range(yearCol[k], yearCol[k+1]):
                if df.iloc[i, j] == k_asked:
                    kk.append(j)
                    restartIndex = i + 1
                    break

        if len(kk) == len(yearReq):
            break

    yearCol.pop()

    if len(kk) != len(yearReq):
        errfile.write(str(now.now()) + " Requested data " + str(yearReq).strip(
            '[]') + " in the field 'All Apprenticeships' don't match the excel file. Please check the file at: " + str(url) + " . End progress\n")
        logfile.write(str(now.now()) + ' error and end progress\n')
        sys.exit("Requested data " + str(yearReq).strip(
            '[]') + " in the field 'All Apprenticeships' don't match the excel file. Please check the file at: " + url)

    raw_data = {}
    for j in col:
        raw_data[j] = []

    # data reading
    logfile.write(str(now.now()) + ' data reading\n')
    print('data reading------')
    for i in range(restartIndex, df.shape[0]):
            ii = 0
            for k in kk:
                if (pd.notnull(df.iloc[i, 1])) and (pd.notnull(df.iloc[i, k])) and (df.iloc[i, 1] != "Total"):
                    ij = 0
                    for jj in ["Under 19", "19-24"]:
                        raw_data[col[0]].append(df.iloc[i, 1])
                        raw_data[col[1]].append(yearReq[ii])
                        raw_data[col[2]].append(jj)
                        raw_data[col[3]].append(df.iloc[i, k+ij])

                        ij += 1

                ii += 1
    logfile.write(str(now.now()) + ' data reading end\n')
    print('data reading end------')

    # save csv file
    dsave.save(raw_data, col, keyCol, digitCheckCol, noDigitRemoveFields, dName, logfile)


parser = argparse.ArgumentParser(description='Extract online Apprenticeship Starts Excel file Local Education Authority to .csv file.')
parser.add_argument("--generateConfig", "-g", help="generate a config file called config_AppStarts.json",
                    action="store_true")
parser.add_argument("--configFile", "-c", help="path for config file")
args = parser.parse_args()

if args.generateConfig:
    obj = {
        "url": "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/382956/apprenticeships-starts-by-geography-level-and-age.xls",
        "outPath": "tempAppStarts.csv",
        "sheet": "Local Education Authority",
        "reqFields": ["2005/06", "2006/07", "2007/08", "2008/09", "2009/10", "2010/11", "2011/12", "2012/13", "2013/14", "2014/15"],
        "colFields": ['name', 'year', 'age', 'value'],
        "primaryKeyCol": ['name', 'year', 'age'],#[0, 1, 2],
        "digitCheckCol": ['value'],#[3],
        "noDigitRemoveFields": []
    }

    logfile = open("log_tempAppStarts.log", "w")
    logfile.write(str(now.now()) + ' start\n')

    errfile = open("err_tempAppStarts.err", "w")

    with open("config_tempAppStarts.json", "w") as outfile:
        json.dump(obj, outfile, indent=4)
        logfile.write(str(now.now()) + ' config file generated and end\n')
        sys.exit("config file generated")

if args.configFile == None:
    args.configFile = "config_tempAppStarts.json"

with open(args.configFile) as json_file:
    oConfig = json.load(json_file)

    logfile = open('log_' + oConfig["outPath"].split('.')[0] + '.log', "w")
    logfile.write(str(now.now()) + ' start\n')

    errfile = open('err_' + oConfig["outPath"].split('.')[0] + '.err', "w")

    logfile.write(str(now.now()) + ' read config file\n')
    print("read config file")

download(oConfig["url"], oConfig["sheet"], oConfig["reqFields"], oConfig["outPath"], oConfig["colFields"], oConfig["primaryKeyCol"], oConfig["digitCheckCol"], oConfig["noDigitRemoveFields"])
