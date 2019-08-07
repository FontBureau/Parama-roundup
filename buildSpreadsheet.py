#coding=utf-8

"""

Make measurements in Google Doc
File > Download as > Comma-separated values (.csv)

Place your measurement file here:
    parama_roundup/import/measurements.csv

In terminal:
$ cd /path/to/my/repository/sources/
$ python3 /path/to/parama_roundup/buildSpreadsheet.py MyDesignspace.designspace
(where MyDesignspace.designspace is your designspace filename)


The output will go in:
    parama_roundup/export/Axes.csv
    parama_roundup/export/Measurements.csv
    
Return to google doc
File > Import (one CSV at a time)
"Create New Sheet"
    
"""
import sys
import os
from fontTools.designspaceLib import *
import csv
import math

# if we are not in robofont, use fontparts to parse UFO
try:
    test = OpenFont
except:
    from fontParts.world import OpenFont

# if we can get vanilla, select the designspace
# if not, use the commandline argument
try:
    from vanilla.dialogs import getFolder, getFile
    designspacePath = getFile('Get designspace file')[0]
except:
    designspacePath = sys.argv[1]

base = os.path.split(__file__)[0]
measurementsPath = os.path.join ( base, 'import/Inputs.csv' )
axesPath = os.path.join(base, 'export/Axes.csv')
sourcesPath = os.path.join(base, 'export/Measurements.csv')

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

# define the spreadsheet structure
measurementsCols = ['Measurement', 'Axis', 'Direction', 'Reference glyph 1', 'Point index 1', 'Reference glyph 2', 'Point index 2', ]
axesCols = ['Tag', 'Label', 'Default', 'Min', 'Max']
sourcesCols = ['Source', 'UPM']

measurementNames = []
axesRows = []
sourcesRows = []

# parse axis information from designspace file
doc = DesignSpaceDocument()
doc.read(designspacePath)
designspaceFileName = os.path.split(designspacePath)[1]
for axis in doc.axes:
    axesRows.append( [axis.tag, axis.labelNames['en'], axis.default, axis.minimum, axis.maximum] )

# write to axes.csv
with open(axesPath, 'w', encoding="utf8") as axisFile:
    csvw = csv.writer(axisFile)
    csvw.writerow(axesCols)
    for row in axesRows:
        csvw.writerow(row)


with open(measurementsPath, encoding="utf8") as measurementsFile:    
    measurementsReader = csv.reader(measurementsFile)
    measurements = {}
    
    # define the measurements and save them as a dictionary of their properties
    for rowIndex, row in enumerate(measurementsReader):
        if rowIndex >= 1:
            mDict = {}
            for i, colName in enumerate(measurementsCols[1:]):
                mDict[colName] = row[i+1]
            measurementNames.append(row[0])
            measurements[row[0]] = mDict

    # loop through designspaces
    doc = DesignSpaceDocument()
    doc.read(designspacePath)
    designspaceFileName = os.path.split(designspacePath)[1]
    # loop through sources within the designspace
    for source in doc.sources:
        #if 'Amstelvar-Roman.ufo' not in source.path:
        #    continue
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
                if len(char) > 1:
                    gname = char
                elif char and ord(char) in charMap:
                    gname = charMap[ord(char)][0]
                
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
            
            # write
            row.append(value)
            row.append(normalizedValue)
                
        sourcesRows.append(row)
        f.close()

# write sources.csv
with open(sourcesPath, 'w', encoding="utf8") as sourcesFile:
    csvw = csv.writer(sourcesFile)
    
    # add headers for upm and normalized values
    headers = sourcesCols
    for measurementName in measurementNames:
        headers.append(measurementName)
        headers.append(measurementName + ' ‰')
    
    csvw.writerow(headers)
    for row in sourcesRows:
        csvw.writerow(row)

print('done')