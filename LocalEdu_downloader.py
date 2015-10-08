__author__ = 'G'

import sys
import urllib
import pandas as pd
import argparse
import json

# url = "https://www.gov.uk/government/uploads/system/uploads/attachment_data/file/382956/apprenticeships-starts-by-geography-level-and-age.xls"
# output_path = "tempAppStarts.csv"
# sheet = "Local Education Authority"
# required_indicators = ["2005/06", "2006/07", "2007/08", "2008/09", "2009/10", "2010/11", "2011/12", "2012/13", "2013/14", "2014/15"]


def download(url, sheet, reqFields, outPath):
    yearReq = reqFields
    dName = outPath

    col = ['name', 'year', 'age', 'value']

    try:
        socket = urllib.request.urlopen(url)
    except urllib.error.HTTPError as e:
        sys.exit('excel download HTTPError = ' + str(e.code))
    except urllib.error.URLError as e:
        sys.exit('excel download URLError = ' + str(e.args))
    except Exception:
        print('excel file download error')
        import traceback
        sys.exit('generic exception: ' + traceback.format_exc())

    # operate this excel file
    xd = pd.ExcelFile(socket)
    df = xd.parse(sheet)

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
        sys.exit("Requested data " + str(yearReq).strip(
            '[]') + " in the field 'All Apprenticeships' don't match the excel file. Please check the file at: " + url)

    raw_data = {}
    for j in col:
        raw_data[j] = []

    print('data reading------')
    for i in range(restartIndex, df.shape[0]):
            print('reading row ' + str(i))
            ii = 0
            for k in kk:
                if (not pd.isnull(df.iloc[i, 1])) and (not pd.isnull(df.iloc[i, k])) and (df.iloc[i, 1] != "Total"):
                    ij = 0
                    for jj in ["Under 19", "19-24"]:
                        raw_data[col[0]].append(df.iloc[i, 1])
                        raw_data[col[1]].append(yearReq[ii])
                        raw_data[col[2]].append(jj)
                        raw_data[col[3]].append(df.iloc[i, k+ij])

                        ij += 1

                ii += 1

    # save csv file
    print('writing to file ' + dName)
    dfw = pd.DataFrame(raw_data, columns=col)
    dfw.to_csv(dName, index=False)
    print('Requested data has been extracted and saved as ' + dName)
    print("finished")

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

    with open("config_AppStarts.json", "w") as outfile:
        json.dump(obj, outfile, indent=4)
        sys.exit("config file generated")

if args.configFile == None:
    args.configFile = "config_AppStarts.json"

with open(args.configFile) as json_file:
    oConfig = json.load(json_file)
    print("read config file")

download(oConfig["url"], oConfig["sheet"], oConfig["reqFields"], oConfig["outPath"])
