from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from fontTools.designspaceLib import *
import sys

KEY = 'com.typenetwork.paramaroundup'

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1-LcFQ2xjnI8GemKExw2xQ5ex8M2Xh-sD6TxMV96EGrY'
RANGE_NAME = 'Inputs!A2:H'

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
    


doc = DesignSpaceDocument()
doc.read(designspacePath)
designspaceFileName = os.path.split(designspacePath)[1]
for source in doc.sources:
    if source.copyInfo:
        f = OpenFont(source.path)
        measurements = {}
        for gname in f.glyphOrder:
            g = f[gname]
            
            getIndices:
                
            
            pointLabels = g.lib.get(com.typemytype.robofont.pointLabels)
            for pointLabel in pointLabels:
                            
                    