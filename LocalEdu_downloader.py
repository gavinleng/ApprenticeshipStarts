__author__ = 'G'


import sys
import urllib
import pandas as pd
import argparse
import json
import datetime
import hashlib

# url = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/382956/apprenticeships-starts-by-geography-level-and-age.xls"
# output_path = "tempAppStarts.csv"
# sheet = "Local Education Authority"
# required_indicators = ["2005/06", "2006/07", "2007/08", "2008/09", "2009/10", "2010/11", "2011/12", "2012/13", "2013/14", "2014/15"]


def download(url, sheet, reqFields, outPath):
    yearReq = reqFields
    dName = outPath

    col = ['name', 'year', 'age', 'value', 'pkey']

    # open url
    socket = openurl(url)

    # operate this excel file
    logfile.write(str(now()) + ' excel file loading\n')
    print('excel file loading------')
    xd = pd.ExcelFile(socket)
    df = xd.parse(sheet)

    # indicator checking
    logfile.write(str(now()) + ' indicator checking\n')
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
        errfile.write(str(now()) + " Requested data " + str(yearReq).strip(
            '[]') + " don't match the excel file. Please check the file at: " + str(url) + " . End progress\n")
        logfile.write(str(now()) + ' error and end progress\n')
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
        errfile.write(str(now()) + " Requested data " + str(yearReq).strip(
            '[]') + " in the field 'All Apprenticeships' don't match the excel file. Please check the file at: " + str(url) + " . End progress\n")
        logfile.write(str(now()) + ' error and end progress\n')
        sys.exit("Requested data " + str(yearReq).strip(
            '[]') + " in the field 'All Apprenticeships' don't match the excel file. Please check the file at: " + url)

    raw_data = {}
    for j in col:
        raw_data[j] = []

    # data reading
    logfile.write(str(now()) + ' data reading\n')
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
    logfile.write(str(now()) + ' data reading end\n')
    print('data reading end------')

    # create primary key by md5 for each row
    logfile.write(str(now()) + ' create primary key\n')
    print('create primary key------')
    keyCol = [0, 1, 2]
    raw_data[col[-1]] = fpkey(raw_data, col, keyCol)
    logfile.write(str(now()) + ' create primary key end\n')
    print('create primary key end------')

    # save csv file
    logfile.write(str(now()) + ' writing to file\n')
    print('writing to file ' + dName)
    dfw = pd.DataFrame(raw_data, columns=col)
    dfw.to_csv(dName, index=False)
    logfile.write(str(now()) + ' has been extracted and saved as ' + str(dName) + '\n')
    print('Requested data has been extracted and saved as ' + dName)
    logfile.write(str(now()) + ' finished\n')
    print("finished")

def openurl(url):
    try:
        socket = urllib.request.urlopen(url)
        return socket
    except urllib.error.HTTPError as e:
        errfile.write(str(now()) + ' file download HTTPError is ' + str(e.code) + ' . End progress\n')
        logfile.write(str(now()) + ' error and end progress\n')
        sys.exit('file download HTTPError = ' + str(e.code))
    except urllib.error.URLError as e:
        errfile.write(str(now()) + ' file download URLError is ' + str(e.args) + ' . End progress\n')
        logfile.write(str(now()) + ' error and end progress\n')
        sys.exit('file download URLError = ' + str(e.args))
    except Exception:
        print('file download error')
        import traceback
        errfile.write(str(now()) + ' generic exception: ' + str(traceback.format_exc()) + ' . End progress\n')
        logfile.write(str(now()) + ' error and end progress\n')
        sys.exit('generic exception: ' + traceback.format_exc())

def fpkey(data, col, keyCol):
    mystring = ''
    pkey = []
    for i in range(len(data[col[0]])):
        for j in keyCol:
            mystring += str(data[col[j]][i])
        mymd5 = hashlib.md5(mystring.encode()).hexdigest()
        pkey.append(mymd5)

    return pkey

def now():
    return datetime.datetime.now()


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
        "reqFields": ["2005/06", "2006/07", "2007/08", "2008/09", "2009/10", "2010/11", "2011/12", "2012/13", "2013/14", "2014/15"]
    }

    logfile = open("log_tempAppStarts.log", "w")
    logfile.write(str(now()) + ' start\n')

    errfile = open("err_tempAppStarts.err", "w")

    with open("config_tempAppStarts.json", "w") as outfile:
        json.dump(obj, outfile, indent=4)
        logfile.write(str(now()) + ' config file generated and end\n')
        sys.exit("config file generated")

if args.configFile == None:
    args.configFile = "config_tempAppStarts.json"

with open(args.configFile) as json_file:
    oConfig = json.load(json_file)

    logfile = open('log_' + oConfig["outPath"].split('.')[0] + '.log', "w")
    logfile.write(str(now()) + ' start\n')

    errfile = open('err_' + oConfig["outPath"].split('.')[0] + '.err', "w")

    logfile.write(str(now()) + ' read config file\n')
    print("read config file")

download(oConfig["url"], oConfig["sheet"], oConfig["reqFields"], oConfig["outPath"])
