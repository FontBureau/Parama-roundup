#coding=utf-8

"""

Place your measurement file:
    parama_roundup/import/measurements.csv

In terminal:
    
$ cd /path/to/my/repository/sources/
$ python3 /path/to/parama_roundup/buildSpreadsheet.py

The Output will go in:
    parama_roundup/export/axes.csv
    parama_roundup/export/sources.csv
"""

import os
from fontTools.designspaceLib import *
from fontParts.world import OpenFont
import csv

# if we can get vanilla, select the folder that contains the designspaces
# if not, use the current working directory
try:
    from vanilla.dialogs import getFolder, getFile
    designspaceFolderPath = getFolder('Folder wih designspaces')[0]
except:
    designspaceFolderPath = os.getcwd()

# collect designspace files from designspace folder
designspacePaths = []
for root, dirs, files in os.walk(designspaceFolderPath):
    for filename in files:
        basePath, ext = os.path.splitext(filename)
        if ext in ['.designspace']:
            designspacePaths.append(os.path.join(root, filename))
print( 'Found %s designspace files in %s' %( str( len(designspacePaths) ), designspaceFolderPath ) )


base = os.path.split(__file__)[0]
measurementsPath = os.path.join ( base, 'import/measurements.csv' )
axesPath = os.path.join(base, 'export/axes.csv')
sourcesPath = os.path.join(base, 'export/sources.csv')

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


def getValueFromGlyphIndex(g, index):
    """
    Given a glyph and a list of point indexes, return a list of those points.
    """
    index = int(index)
    i = 0
    for c in g:
        for p in c.points:
            if i == index:
                return (p.x, p.y)
            i += 1

def getValuesFromGlyphIndexes(g, indexes):
    """
    Given a glyph and a list of point indexes, return a list of those points.
    """
    indexes = [int(ii) for ii in indexes]
    
    values = {}
    i = 0
    for c in g:
        for p in c.points:
            if i in indexes:
                #print('match', i, p.x, p.y)
                values[i] = (p.x, p.y)
            i += 1
    points = []
    for i in indexes:
        if i in values:
            points.append(values[i])
        else:
            print('\tError,', 'missing point:', g.name, os.path.split(g.getParent().path)[1], 'index', i)
    #print(indexes, points)
    return points

# define the spreadsheet structure
measurementsCols = ['Measurement', 'Axis', 'Direction', 'Reference glyph 1', 'Point index 1', 'Reference glyph 2', 'Point index 2', ]
axesCols = ['Tag', 'Designspace', 'Label', 'Default', 'Min', 'Max']
sourcesCols = ['Source', 'Designspace', 'UPM']

measurementNames = []
axesRows = []
sourcesRows = []

# parse axis information from designspace file
for designspacePath in designspacePaths:
    doc = DesignSpaceDocument()
    doc.read(designspacePath)
    designspaceFileName = os.path.split(designspacePath)[1]
    for axis in doc.axes:
        axesRows.append( [axis.tag, designspaceFileName, axis.labelNames['en'], axis.default, axis.minimum, axis.maximum] )

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
    for designspacePath in designspacePaths:
        doc = DesignSpaceDocument()
        doc.read(designspacePath)
        designspaceFileName = os.path.split(designspacePath)[1]
        # loop through sources within the designspace
        for source in doc.sources:
            #if 'Amstelvar-Roman.ufo' not in source.path:
            #    continue
            f = OpenFont(source.path, showInterface=False)
            charMap = f.getCharacterMapping()
            
            row = [os.path.split(source.path)[1], designspaceFileName, int(f.info.unitsPerEm)]
            
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
                    value = int(round(value))
                    normalizedValue = int(round(normalizedValue, 2))
                
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