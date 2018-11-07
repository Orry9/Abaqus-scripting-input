# Import abaqus odb work related modules
from odbAccess import *
from abaqusConstants import *
from odbMaterial import *
from odbSection import *

from openpyxl import Workbook

 



import logging
logger = logging.getLogger(__name__)
loggerHandler = logging.StreamHandler(sys.__stdout__)
loggerHandler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s:%(message)s'))
loggerHandler.setLevel(logging.DEBUG)
logger.addHandler(loggerHandler)
logger.setLevel(logging.DEBUG)

logger.info('readOdb started')



workbook = Workbook()

# Set the active worksheet to be the Summary
worksheetSummary = workbook.active
worksheetSummary.title = 'Summary Experiment strong-axis'
worksheetSummary['A1'] = 'Experiment Number'
worksheetSummary['B1'] = 'Max. Reaction Force'
worksheetSummary['C1'] = 'Max. Reaction Alert'

# The empty dictionary for the collection of the max. Reaction Force
rf3MaxDict = {}
rf3Alert = {}

#for jobNr in range(1,168):
for jobNr in range(1,469):

    logger.debug('Creating job worksheet : Job-{0:0>4}'.format(jobNr))
    worksheetJob = workbook.create_sheet(title='Job-{0:0>4}'.format(jobNr))

    # Open odb file
    odb = openOdb(path='Z://strong-axis//Job-S-{0:0>4}/{0:0>4}-Job-2.odb'.format(jobNr, jobNr))

    # Get the NL-buckle step
    step = odb.steps['NL-buckle']

    # Select the history regions
    historyRegions = step.historyRegions

    # Create an empty jobData python dictionary object to capture
    # the data of interest
    jobData = {}

    # Iterate over the history regions keys
    for historyRegionKey in historyRegions.keys():
        
        # We are solely interested in history region keys
        # that contain the string 'Node I'
        if (historyRegionKey.find('Node I') != -1):
            # Logging the history region key iterating progress
            logger.debug('historyRegionKey : {}'.format(historyRegionKey))
            
            # Iterate over the history output keys
            for historyOutputKey in historyRegions[historyRegionKey].historyOutputs.keys():
                # Logging the history output key iterating progress
                logger.debug('historyOutputKey : {}'.format(historyOutputKey))
                
                # Capture the data in our dictionary
                jobData[historyOutputKey] = historyRegions[historyRegionKey].historyOutputs[historyOutputKey].data


    
    
    # Insertion of the 'Reaction Force' data in the Job sheet
    rowIndex = 1
    rf3DataColumnName = 'A'
    cellIndex ='{}{}'.format(rf3DataColumnName, rowIndex)
    worksheetJob[cellIndex] = 'Reaction Force'
    rowIndex = rowIndex + 1
    nrOfReactionForceElements = len(jobData['RF3'])
    reactionForceIndex = 0
    indexOfMaxReactionForce = -1
    rf3Max = jobData['RF3'][0][1]
    
    for rf3Data in jobData['RF3']:
        cellIndex ='{}{}'.format(rf3DataColumnName, rowIndex)
        worksheetJob[cellIndex] = rf3Data[1]
        
        if (rf3Max < rf3Data[1]):
            rf3Max = rf3Data[1]
            indexOfMaxReactionForce = reactionForceIndex     
        
        reactionForceIndex = reactionForceIndex + 1
        rowIndex = rowIndex + 1    

    
    rf3MaxDict[jobNr] = rf3Max
    if (indexOfMaxReactionForce >= (nrOfReactionForceElements - 5)):
        rf3Alert[jobNr] = 'X'
    else :
        rf3Alert[jobNr] = '' 


    # Insertion of the 'Lateral Displacement' data in the Job sheet
    rowIndex = 1
    u2DataColumnName = 'B'
    cellIndex ='{}{}'.format(u2DataColumnName, rowIndex)
    worksheetJob[cellIndex] = 'Lateral Displacement'
    rowIndex = rowIndex + 1
    for u2Data in jobData['U2']:
        cellIndex ='{}{}'.format(u2DataColumnName, rowIndex)
        worksheetJob[cellIndex] = u2Data[1]     
        rowIndex = rowIndex + 1    

    # Insertion of the 'Axial Displacement' data in the Job sheet
    rowIndex = 1
    u3DataColumnName = 'C'
    cellIndex ='{}{}'.format(u3DataColumnName, rowIndex)
    worksheetJob[cellIndex] = 'Axial Displacement'
    rowIndex = rowIndex + 1
    for u3Data in jobData['U3']:
        cellIndex ='{}{}'.format(u3DataColumnName, rowIndex)
        worksheetJob[cellIndex] = u3Data[1]     
        rowIndex = rowIndex + 1    




jobColumnName = 'A'
maxReactionForceColumnName = 'B'
alertColumnName = 'C'

for key in rf3MaxDict.keys():
    jobNr = key
    maxReactionForce = rf3MaxDict[key]
    maxReactionAlert = rf3Alert[key]
    
    cellIndex ='{}{}'.format(jobColumnName, (jobNr + 1)) 
    worksheetSummary[cellIndex] = jobNr
    
    cellIndex ='{}{}'.format(maxReactionForceColumnName, (jobNr + 1)) 
    worksheetSummary[cellIndex] = maxReactionForce
    
    cellIndex ='{}{}'.format(alertColumnName, (jobNr + 1)) 
    worksheetSummary[cellIndex] = maxReactionAlert
    

# Save the file
workbook.save("strong-axis-summary.xlsx")
