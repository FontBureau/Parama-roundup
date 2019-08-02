#coding=utf-8
import os
from fontTools.designspaceLib import *
from fontParts.world import OpenFont
import csv

try:
    from vanilla.dialogs import getFolder, getFile
    designspaceFolderPath = getFolder('Measurements CSV')[0]
except:
    designspaceFolderPath = os.getcwd()

designspacePaths = []
for root, dirs, files in os.walk(designspaceFolderPath):
    for filename in files:
        basePath, ext = os.path.splitext(filename)
        if ext in ['.designspace']:
            designspacePaths.append(os.path.join(root, filename))

print( 'Found %s designspace files in %s' %( str( len(designspacePaths) ), designspaceFolderPath ) )

measurementsPath = os.path.join ( os.path.split(__file__)[0], 'measurements/measurements.csv' )

axesPath = '/Users/david/Documents/Font_Bureau/Misc/Parama-roundup/export/axes.csv'
sourcesPath = '/Users/david/Documents/Font_Bureau/Misc/Parama-roundup/export/sources.csv'


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

def getValuesFromGlyphIndexes(g, indexes):
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
            print('Error,', 'missing point:', g.name, os.path.split(g.getParent().path)[1], 'index', i)
    #print(indexes, points)
    return points


measurementsCols = ['Measurement', 'Axis', 'Direction', 'Reference glyph', 'Point indexes']
axesCols = ['Tag', 'Designspace', 'Label', 'Default', 'Min', 'Max']
sourcesCols = ['Source', 'Designspace', 'UPM']

measurementNames = []

axesRows = []
sourcesRows = []


for designspacePath in designspacePaths:
    doc = DesignSpaceDocument()
    doc.read(designspacePath)
    designspaceFileName = os.path.split(designspacePath)[1]
    for axis in doc.axes:
        axesRows.append( [axis.tag, designspaceFileName, axis.labelNames['en'], axis.default, axis.minimum, axis.maximum] )

with open(axesPath, 'w', encoding="utf8") as axisFile:
    csvw = csv.writer(axisFile)
    csvw.writerow(axesCols)
    for row in axesRows:
        csvw.writerow(row)



with open(measurementsPath, encoding="utf8") as measurementsFile:
    
    measurementsReader = csv.reader(measurementsFile)
    
    measurements = {}
        
    for rowIndex, row in enumerate(measurementsReader):
        if rowIndex >= 1:
            mDict = {}
            for i, colName in enumerate(measurementsCols[1:]):
                mDict[colName] = row[i+1]
            measurementNames.append(row[0])
            measurements[row[0]] = mDict

    for designspacePath in designspacePaths:
        doc = DesignSpaceDocument()
        doc.read(designspacePath)
        designspaceFileName = os.path.split(designspacePath)[1]
        for source in doc.sources:
            
            #if 'Amstelvar-Roman-opsz-min.ufo' not in source.path:
            #    continue
            
            f = OpenFont(source.path, showInterface=False)
            charMap = f.getCharacterMapping()
            row = [os.path.split(source.path)[1], designspaceFileName, int(f.info.unitsPerEm)]
            for measurementName in measurementNames:
                mDict = measurements[measurementName]
                direction = mDict['Direction']
                if direction == 'x':
                    directionIndex = 0
                elif direction == 'y':
                    directionIndex = 1

                pointIndexes = mDict['Point indexes'].split(' ')

                char = mDict['Reference glyph']
                gname = None
                if len(char) > 1:
                    gname = char
                if not gname and char and ord(char) in charMap:
                    gname = charMap[ord(char)][0]
                if gname and pointIndexes:
                    g = f[gname]
                    #print(f, g, pointIndexes)
                    values = getValuesFromGlyphIndexes(g, pointIndexes)
                    if len(values) == 1:
                        value = values[0][directionIndex]
                    else:
                        value = values[1][directionIndex]-values[0][directionIndex]
                    
                    normalizedValue = value/f.info.unitsPerEm * 1000
                    
                    value = int(round(value))
                    normalizedValue = int(round(normalizedValue))
                    
                    row.append(value)
                    row.append(normalizedValue)
                    
            sourcesRows.append(row)
            f.close()
        
with open(sourcesPath, 'w', encoding="utf8") as sourcesFile:

    csvw = csv.writer(sourcesFile)
    
    headers = sourcesCols
    for measurementName in measurementNames:
        headers.append(measurementName)
        headers.append(measurementName + ' ‰')
    
    csvw.writerow(headers)
    for row in sourcesRows:
        csvw.writerow(row)