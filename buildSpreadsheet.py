#coding=utf-8
import os
from fontTools.designspaceLib import *
from fontParts.world import OpenFont
from vanilla.dialogs import getFolder, getFile
import csv

#repositoryPath = getFile('Measurements CSV', allowsMultipleSelection=True)
#measurementsPath = getFile('Measurements CSV')

designspacePaths = [
    '/Users/david/Desktop/workspace/FB/fb-Amstelvar/sources/Amstelvar-NewCharset/Amstelvar-Roman-008.designspace', 
    '/Users/david/Desktop/workspace/FB/fb-Amstelvar/sources/Amstelvar-NewCharset/Amstelvar-Italic-002.designspace']

measurementsPath = 'measurements/Parametric measurements sample spreadsheet - Measurements.csv'



def getValuesFromGlyphIndexes(g, indexes):
    i = 0
    values = {}
    for c in g:
        for s in c:
            for p in s:
                if i in indexes:
                    #print('match', i, p.x, p.y)
                    values[i] = (p.x, p.y)
                i += 1
    points = []
    for i in indexes:
        points.append(values[i])
    return points


measurementsCols = ['Measurement', 'Axis', 'Direction', 'Reference glyph', 'Point indexes']
axesCols = ['Tag', 'Designspace', 'Label', 'Default', 'Min', 'Max']
sourcesCols = ['Master', 'Designspace', 'UPM']

measurementNames = []

axisRows = []
sourcesRows = []

sources = []


for designspacePath in designspacePaths:
    doc = DesignSpaceDocument()
    doc.read(designspacePath)
    designspaceFileName = os.path.split(designspacePath)[1]
    for axis in doc.axes:
        axisRows.append( [axis.tag, axis.labelNames['en'], axis.default, axis.minimum, axis.maximum] )





with open(measurementsPath, encoding="utf8") as measurementsFile:
    measurementsReader = csv.reader(measurementsFile)
    for row in measurementsReader:

        mDict = {}
        for i, colName in enumerate(measurementsCols):
            mDict[colName] = row[i]
    
        measurementNames.append(mDict['Measurement'])
        
        
