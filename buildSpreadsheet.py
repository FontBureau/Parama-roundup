#coding=utf-8
import sys
import os
from fontTools.designspaceLib import *
import csv
import math
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import tempfile
import shutil

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

basePath = os.path.split(__file__)[0]

def getCreds():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                os.path.join(basePath, 'credentials.json'), SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds


def getInputData(SPREADSHEET_ID):
    service = build('sheets', 'v4', credentials=getCreds())
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Index Refs").execute()
    values = result.get('values', [])
    return values


def writeOutputData(csvData, SPREADSHEET_ID, sheet_name):

    # write and then read a CSV from our data

    pathToCSV = tempfile.mkstemp()[1]
    with open(pathToCSV, 'w', encoding="utf8") as csvFile:
        csvw = csv.writer(csvFile)
        for row in csvData:
            csvw.writerow(row)
    with open(pathToCSV, 'r', encoding="utf8") as csvFile2:
        csvData = csvFile2.read()    

    # call up the api
    service = build('sheets', 'v4', credentials=getCreds())
    sheet = service.spreadsheets()
    sheet_metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    properties = sheet_metadata.get('sheets')
    
    # find the sheet name in the spreadsheet, if it exists
    sheet_id = None
    for item in properties:
        if item.get("properties").get('title') == sheet_name:
            sheet_id = (item.get("properties").get('sheetId'))
    
    # if there is no matching sheet, create one
    if sheet_id == None:
        sheet_id = createSheet(SPREADSHEET_ID, sheet_name)
    
    print('Writing to sheet name', sheet_name, sheet_id)
    
    body = {
        'requests': [{
            'pasteData': {
                "coordinate": {
                    "sheetId": sheet_id,
                },
                "data": csvData,
                "type": 'PASTE_NORMAL',
                "delimiter": ',',
            }
        },
        
        
    {
      "updateSheetProperties": {
        "properties": {
          "sheetId": sheet_id,
          "gridProperties": {
            "frozenRowCount": 1
          }
        },
        "fields": "gridProperties.frozenRowCount"
      }
    }
        
        ]
    }
    result = sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()

    os.remove(pathToCSV)    


def createSheet(SPREADSHEET_ID, title):
    print('creating new sheet', title)

    body = {
          "requests": [
            {
              "addSheet": {
                "properties": {
                  "title": title,
                  "gridProperties": {
                    "rowCount": 20,
                    "columnCount": 12
                  },
                  #"tabColor": {
                  #  "red": 1.0,
                  #  "green": 0.3,
                  #  "blue": 0.4
                  #}
                }
              }
            }
          ]
        }
    service = build('sheets', 'v4', credentials=getCreds())
    sheet = service.spreadsheets()
    result = sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()

    sheet_metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    properties = sheet_metadata.get('sheets')
    sheet_id = None
    for item in properties:
        if item.get("properties").get('title') == title:
            sheet_id = (item.get("properties").get('sheetId'))
    return(sheet_id)
    

# not using these functions yet, but I might
def excel_column_name(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name

def excel_column_number(name):
    """Excel-style column name to number, e.g., A = 1, Z = 26, AA = 27, AAA = 703."""
    n = 0
    for c in name:
        n = n * 26 + 1 + ord(c) - ord('A')
    return n

# how we round numbers for the spreadsheet
def doRounding(value):
    rounded = round(value, 1)
    if rounded == int(rounded):
        return int(rounded)
    if rounded % 2:
        return int(math.ceil(rounded))
    else:
        return int(math.floor(rounded))
    

# get point indexes from glyph
def getValueFromGlyphIndex(g, index):
    """
    Given a glyph and a point index, return that point.
    """
    index = int(index)
    i = 0
    for c in g:
        for p in c.points:
            if i == index:
                return (p.x, p.y)
            i += 1

    """If the entire set of contours for a glyph requires “n” points (i.e., contour points numbered from 0 to n-1), then the scaler will add points n, n+1, n+2, and n+3. Point “n” will be placed at the character origin, point “n+1” will be placed at the advance width point, point “n+2” will be placed at the top origin, and point “n+3” will be placed at the advance height point. For an illustration of how these phantom points are placed, see Figure 2-1 (points 17, 18, 19, 20) and Figure 2-2 (points 27, 28, 29, and 30).
    https://docs.microsoft.com/en-us/typography/opentype/spec/tt_instructing_glyphs#phantoms"""

    if index == i:
        return (0, 0)
    elif index == i+1:
        return (g.width, 0)
    elif index == i+2:
        return (0, g.getParent().info.ascender)
    elif index == i+3:
        return (g.width, g.getParent().info.ascender)



if __name__ == "__main__":

    writeAxes = True
    writeMeasurements = True
    writeWidths = True

    # if we are not in robofont, use fontparts to parse UFO
    try:
        test = OpenFont
    except:
        from fontParts.world import OpenFont

    import argparse

    parser = argparse.ArgumentParser(description='Make measurements in Google Doc')
    parser.add_argument('-d', '--designspace', type=str, dest='designspace',
                        help='the designspace file we will be analyzing')
    parser.add_argument('-i', '--input-spreadsheet', type=str,  dest='input',
                        help='the google spreadsheet we will read from')
    parser.add_argument('-o', '--output-spreadsheet', type=str, dest='output',
                        help='the google spreadsheet we will write to')

    args = parser.parse_args()

    designspacePath = args.designspace
    inputSpreadsheetID = args.input
    outputSpreadsheetID = args.output
    if not args.output:
        outputSpreadsheetID = inputSpreadsheetID

    print('designspacePath', designspacePath)
    print('inputSpreadsheetID', inputSpreadsheetID)
    print('outputSpreadsheetID', outputSpreadsheetID)

    # write stuff locally as well
    tempBasePath = tempfile.mkdtemp()
    axesPath = os.path.join(tempBasePath, 'axes.csv')
    measurementsPath = os.path.join(tempBasePath, 'measurements.csv')
    widthsPath = os.path.join(tempBasePath, 'widths.csv')


    # define the spreadsheet structure
    measurementsCols = ['Axis', 'Description', 'Direction', 'Note', 'Reference glyph 1', 'Point index 1', 'Reference glyph 2', 'Point index 2', 'Formula']
    axesCols = ['Tag', 'Label', 'Default', 'Min', 'Max']
    sourcesCols = ['Source', 'UPM']

    measurementNames = []
    axesRows = []
    sourcesRows = []

    widthsRows = []
    widthsCols= ['']

    # parse axis information from designspace file
    doc = DesignSpaceDocument()
    doc.read(designspacePath)
    designspaceFileName = os.path.split(designspacePath)[1]
    for axis in doc.axes:
        try:
            axesRows.append( [axis.tag, axis.labelNames['en'], axis.default, axis.minimum, axis.maximum] )
        except:
            axesRows.append( [axis.tag, '', axis.default, axis.minimum, axis.maximum] )

    if writeMeasurements or writeWidths:

        # get the measurements from the "Index Refs" sheet in the input spreadsheet
        # and collect them in a dictionary
        measurementInputs = getInputData(inputSpreadsheetID)
        measurements = {}
        for rowIndex, row in enumerate(measurementInputs):
            if rowIndex >= 1 and row:
                mDict = {}
                while len(row) < len(measurementsCols):
                    row.append('')
                for i, colName in enumerate(measurementsCols):
                    mDict[colName] = row[i] 
                # only analyze rows with a direction
                if mDict['Direction']:
                    measurementNames.append(row[0])
                    measurements[row[0]] = mDict

        # loop through designspace
        doc = DesignSpaceDocument()
        doc.read(designspacePath)
        designspaceFileName = os.path.split(designspacePath)[1]

        # gather glyph index for width counting
        for source in doc.sources:
            if source.copyInfo:
                f = OpenFont(source.path, showInterface=False)
                for gname in f.glyphOrder:
                    if gname in f:
                        widthsCols.append(gname)
                f.close()
        widthsRows.append(widthsCols)

        # process each source
        for source in doc.sources:
            if not os.path.exists(source.path):
                print('missing source', source.path)
                continue
            f = OpenFont(source.path, showInterface=False)
            charMap = f.getCharacterMapping()
    
            row = [os.path.split(source.path)[1], int(f.info.unitsPerEm)]
    
            # get the measurements
            for measurementName in measurementNames:
                value = None
                mDict = measurements[measurementName]

                # if x, use 0, if y, use 1
                direction = mDict['Direction']
                if direction == 'x':
                    directionIndex = 0
                elif direction == 'y':
                    directionIndex = 1
                
                # interpret the reference glyph and point indexes and obtain the measurement
                chars = [ 
                    ( mDict['Reference glyph 1'], mDict['Point index 1'] ), 
                    ( mDict['Reference glyph 2'], mDict['Point index 2'] ) 
                    ]
                values = [None, None]
                        
                for i, charAndPoint in enumerate(chars):
                    char, pointIndex = charAndPoint
                    # if the reference field is one character long, use the glyph with that character's unicode
                    # if not, assume that it’s the glyph name
                    gname = None
                    if char and len(char) == 1 and ord(char) in charMap:
                        gname = charMap[ord(char)][0]
                    elif char != '':
                            hexchar = int(char, 16)
                            if hexchar in charMap:
                                gname = charMap[hexchar][0]
                            else:
                                print('Could not find glyph for char', char)
                                continue
                    else:
                        char = gname
                    if gname and gname in f and pointIndex:
                        g = f[gname]
                        value = getValueFromGlyphIndex(g, pointIndex)
                        values[i] = value
            
                # if there's one point index, use one value
                # if there are multiple point indexes, subtract them
                if values[0] is not None and values[1] is not None:
                    value = abs( values[1][directionIndex]-values[0][directionIndex] )
                elif values[0] is not None:
                    value = values[0][directionIndex]
                else:
                    value = None

                # get normalized
                normalizedValue = None
                if value is not None:
                    normalizedValue = value/f.info.unitsPerEm * 1000
                  
                    value = doRounding(value)
                    normalizedValue = doRounding(normalizedValue)
        
                # write (UPM value and then permille)
                row.append(value)
                row.append(normalizedValue)
            
            sourcesRows.append(row)
            
            # process the widths
            widthsRow = [os.path.split(source.path)[1]]
            for gname in widthsCols[1:]:
                if gname in f:
                    widthsRow.append(f[gname].width)
                else:
                    widthsRow.append('')
    
            widthsRows.append(widthsRow)
    
            f.close()

    # write stuff to local CSVs and to the output spreadsheet
    if writeAxes:
        with open(os.path.join(os.path.split(designspacePath)[0], axesPath), 'w', encoding="utf8") as axisFile:
            csvw = csv.writer(axisFile)
            csvw.writerow(axesCols)
            for row in axesRows:
                csvw.writerow(row)
        writeOutputData([axesCols]+axesRows, outputSpreadsheetID, 'Axes')

    if writeMeasurements:
        with open(measurementsPath, 'w', encoding="utf8") as sourcesFile:
            csvw = csv.writer(sourcesFile)
            # add headers for upm and normalized values
            headers = sourcesCols
            for measurementName in measurementNames:
                headers.append(measurementName)
                headers.append(measurementName + ' ‰')
            csvw.writerow(headers)
            rows = []
            for row in sourcesRows:
                csvw.writerow(row)
                rows.append(row)
        writeOutputData([headers]+rows, outputSpreadsheetID, 'Measurements')

    if writeWidths:
        with open(os.path.join(os.path.split(designspacePath)[0], widthsPath), 'w', encoding="utf8") as widthsFile:
            csvw = csv.writer(widthsFile)
            for row in widthsRows:
                csvw.writerow(row)
            writeOutputData(widthsRows, outputSpreadsheetID, 'Widths')

    # get rid of the temp CSV folder
    shutil.rmtree(tempBasePath)
    print('done')
    