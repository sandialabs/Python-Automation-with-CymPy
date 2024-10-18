# -*- coding: utf-8 -*-
"""
BSD 3-Clause License

Copyright 2024 National Technology & Engineering Solutions of Sandia, LLC (NTESS). Under the terms of Contract DE-NA0003525 with NTESS, the U.S. Government retains certain rights in this software.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

* Redistributions of source code must retain the above copyright notice, this
  list of conditions and the following disclaimer.

* Redistributions in binary form must reproduce the above copyright notice,
  this list of conditions and the following disclaimer in the documentation
  and/or other materials provided with the distribution.

* Neither the name of the copyright holder nor the names of its
  contributors may be used to endorse or promote products derived from
  this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

@author: jazzoli & lblakel
This script authored by Joseph Azzolini, Logan Blakely, and Matthew J. Reno
at Sandia National Laboratories as part of DOE i2x funding

Contact:  Joseph Azzolini - jazzoli@sandia.gov
Logan Blakely - lblakel@sandia.gov
Matthew J. Reno - mjreno@sandia.gov
"""

###               Set Switch States Manually and Run EPRI DRIVE Script             ###


# This script reads in a set of switching device settings from a CSV file, sets
#   the switches in the study to match that configuration, and then runs EPRI
#   DRIVE and reports the hosting capacity for that switching configuration
#
#  Includes Switches, Reclosers, and Breakers

# The workflow is:
    #  1.  Load .sxst model using CymPy library
    #  2.  Load CSV files with switching device settings
    #  3.  Manually set the devices in the study to match CSV settings
    #  4.  Run EPRI Drive again to determine hosting capacity
    
# The CSV should have these columns:
#   Switch ID - the id of the switching device which must match the naming in the study file
#   Status - this should be either Close or Open
#   Type - this should be Switch, Recloser, or Breaker (the scripts could be expanded to include whatever switching devices were required)

#%% Python Library Imports
import numpy as np
import pandas as pd
import cympy
import cympy.rm
import locale
#import xlrd

###############################################################################

#%%  Set directory paths and filenames

# studyFolderPath = r'C:\<Path>\<To>\<Study>\<Folder>'
# studyFilename = r'\studyFile.sxst'

studyFolderPath = r'C:\Program Files\CYME\CYME\tutorial\How-to'
studyFilename = r'\NetwConfOptimiz.sxst'

# Folder to save .xlrd and .csv results
saveResultsFolder = r'C:\<Path>\<To>\<Save\<Results>'



# Location and name of .csv with switch states
switchStatesFolder = r'C:\<Path>\<To>\<Switch>\<CSV>'
switchStatesFilename = '\SwitchingDeviceStates_Manual_NCO.csv'



###############################################################################



#%% Open CYME Study and Verify that it loaded correctly

print('Opening CYME Study')
print('')

locale.setlocale(locale.LC_NUMERIC, '')
cympy.app.ActivateRefresh(True)

studyFilePath = studyFolderPath + studyFilename
cympy.study.Open(studyFilePath)
cympy.study.ActivateModifications(False)

# Check to see if the study loaded
spot_loads = cympy.study.ListDevices(cympy.enums.DeviceType.SpotLoad)
networks_df = pd.DataFrame(cympy.study.ListNetworks())

networks = cympy.study.ListNetworks()  # This gives all 'networks' which may include transmission lines

feeders = cympy.study.ListNetworks(cympy.enums.NetworkType.Feeder)  # This gives only 'networks' which are feeders which is what Drive requires
    

# Get load models and load characteristics for EPRI DRIVE settings
# In this example, the default load model was used for peak and light load conditions
# This load model was hard-coded in this example, since it is a required input for EPRI DRIVE
# However, this code shows how you would access the various load models available in a CYME study 
# Typically, DRIVE would need a separate load model for light loading and peak loading, 
# but this example circuit only contains a single default load model.
LoadModels = cympy.study.ListLoadModels()
LoadModelNames = []
LoadModelIDs = np.zeros((len(LoadModels),1),dtype=int)
for ii in range(len(LoadModels)):
    LoadModelIDs[ii]=LoadModels[ii].ID
    LoadModelNames.append(LoadModels[ii].Name)
peakLoadName = LoadModelNames[0]
peakLoadID = LoadModelIDs[0,0]
lightLoadName = LoadModelNames[0]
lightLoadID = LoadModelIDs[0,0]

feeders = cympy.study.ListNetworks(cympy.enums.NetworkType.Feeder)

load_flow = cympy.sim.LoadFlow()
load_flow.Run()


###############################################################################


#%%  Read and Parse CSV file with manual switch settings


switchStatesFilePath = switchStatesFolder + switchStatesFilename
manSwitchStatesDF = pd.read_csv(switchStatesFilePath)

manDeviceIDs = manSwitchStatesDF['Switch ID']
manDeviceStates = manSwitchStatesDF['Status']
manDeviceTypes = manSwitchStatesDF['Type']


# Get list of switches from the study
switchList = cympy.study.ListDevices(cympy.enums.DeviceType.Switch)
breakerList = cympy.study.ListDevices(cympy.enums.DeviceType.Breaker)
recloserList = cympy.study.ListDevices(cympy.enums.DeviceType.Recloser)

switchTest=switchList[0]
switchTest.GetValue('ClosedPhase')


# Using the list of ids, get the current/active switch state
allOrgStates = []
switchIDs = []
for switchCtr in range(0,len(switchList)):
    currSwitch = switchList[switchCtr]
    currID = currSwitch.GetValue('DeviceNumber')
    switchIDs.append(currID)
    allOrgStates.append(currSwitch.GetValue('ClosedPhase'))

# From breaker list, get a list of all breaker ids
breakerIDs = []
for breakerCtr in range(0,len(breakerList)):
    currBreaker = breakerList[breakerCtr]
    currID = currBreaker.GetValue('DeviceNumber')  
    breakerIDs.append(currID)
    allOrgStates.append(currBreaker.GetValue('ClosedPhase'))

# From recloser list, get a list of all recloser ids and states
recloserIDs = []
recloserStatus = []
for recloserCtr in range(0,len(recloserList)):
    currRecloser = recloserList[recloserCtr]
    currID = currRecloser.GetValue('DeviceNumber')
    recloserIDs.append(currID)
    allOrgStates.append(currRecloser.GetValue('ClosedPhase'))
 

# This section splits the devices by type - new types would need to be added here (for example fuses)
# This section also does error checking to make sure the device IDs in the CSV match device IDs in the study
#   devices which do not match are excluded, but the script will continue
manRecloserIDs = []
manRecloserStatus = []
manSwitchIDs = []
manSwitchStatus = []
manBreakerIDs = []
manBreakerStatus = []
missingDevices = []
unknownTypes = []

for deviceCtr in range(0,len(manDeviceIDs)):
    currID = manDeviceIDs[deviceCtr]
    currState = manDeviceStates[deviceCtr]
    currType = manDeviceTypes[deviceCtr]
    
    if currType == 'Switch':
        if currID not in set(switchIDs):
            missingDevices.append(currID)
        else:
            manSwitchIDs.append(currID)
            manSwitchStatus.append(currState)
    elif currType == 'Recloser':
        if currID not in set(recloserIDs):
            missingDevices.append(currID)
        else:
            manRecloserIDs.append(currID)
            manRecloserStatus.append(currState)
    elif currType == 'Breaker':
        if currID not in set(breakerIDs):
            missingDevices.append(currID)
        else:
            manBreakerIDs.append(currID)
            manBreakerStatus.append(currState)
    else:
        unknownTypes.append(currID)
    
    if len(unknownTypes) != 0:
        print('There are unknown device types in the CSV list.  You may need to add those device types to the script.  For this run, those devices have been excluded.  ')
    if len(missingDevices) != 0:
        print('There are device IDs in the CSV list which do not match device IDs in the study.  For this run those devices have been excluded. ')
    


# Notes:  
#   There are two different ways in the script of referencing the devices.  
#       The switchList variable has the Switch objects from cympy and these are
#       what is in the cyme study
print('CymPy Switch object:')
print(switchList[0])
print('')
#       and that has the fields DeviceNumber and DeviceType
#       In most of the script code I have also created lists of device ids, 
#       states, and types to work with
print('Device Lists in the script:')
print(switchIDs[0])
print('or')
print(manSwitchIDs[0])
print(manSwitchStatus[0])


###############################################################################

#%%  Manually set the switching device states

# This script doesn't check to see if the new states are the same as the old states
#       it just uses the values in the CSV to set everything
#       if you wanted to check in advance if the settings were the same or different,
#           or just compare the csv version to the study version
#           you could use originalValue = SwitchList[switchIndex].GetValue('ClosedPhase')
#           for each switching device

# Set Switch states
newSwitchStates = []
for switchCtr in range(0,len(manSwitchIDs)):
    currID = manSwitchIDs[switchCtr]
    currState = manSwitchStatus[switchCtr]
    # Get the location of the device in the study which corresponds to the current device from the CSV
    switchIndex = np.where(np.array(switchIDs) == currID)[0][0]
    if currState == 'Open':
        switchList[switchIndex].SetValue('None','ClosedPhase')
        newSwitchStates.append('None')
    else:
        # If the switch is closed we also need the phase information to correctly set the switch state
        currSection = cympy.study.GetSection(switchList[switchIndex].SectionID)
        currPhase = currSection.GetValue('Phase')
        switchList[switchIndex].SetValue(currPhase,'ClosedPhase')
        newSwitchStates.append(currPhase)


# Set Recloser states
for recloserCtr in range(0,len(manRecloserIDs)):
    currID = manRecloserIDs[recloserCtr]
    currState = manRecloserStatus[recloserCtr]
    # Get the location of the device in the study which corresponds to the current device from the CSV
    recloserIndex = np.where(np.array(recloserIDs) == currID)[0][0]
    if currState == 'Open':
        recloserList[recloserIndex].SetValue('None','ClosedPhase')
        newSwitchStates.append('None')

    else:
        # If the device is closed we also need the phase information to correctly set the switch state
        currSection = cympy.study.GetSection(recloserList[recloserIndex].SectionID)
        currPhase = currSection.GetValue('Phase')
        recloserList[recloserIndex].SetValue(currPhase,'ClosedPhase')
        newSwitchStates.append(currPhase)
               

# Set Breaker states
for breakerCtr in range(0,len(manBreakerIDs)):
    currID = manBreakerIDs[breakerCtr]
    currState = manBreakerStatus[breakerCtr]
    # Get the location of the device in the study which corresponds to the current device from the CSV
    breakerIndex = np.where(np.array(breakerIDs) == currID)[0][0]
    if currState == 'Open':
        breakerList[breakerIndex].SetValue('None','ClosedPhase')
        newSwitchStates.append('None')

    else:
        # If the switch is closed we also need the phase information to correctly set the switch state
        currSection = cympy.study.GetSection(breakerList[breakerIndex].SectionID)
        currPhase = currSection.GetValue('Phase')
        breakerList[breakerIndex].SetValue(currPhase,'ClosedPhase')
        newSwitchStates.append(currPhase)


#%% Set Up EPRI DRIVE Parameters 
DRIVE = cympy.sim.EPRIDrive()  # Assigns the specific simulation a variable name

#cympy.sim.

cympy.Describe('EPRIDriveParameters') # prints the settable parameters for this tool/simulation

# Use DRIVE.SetValue and DRIVE.GetValue to set a new value and get the current value respectively
#Parameters/Future Resource Settings
DRIVE.SetValue(True,'IncludeExistingDER')
DRIVE.SetValue('Photovoltaic','DERType')
DRIVE.SetValue(True,'InverterBasedInterface')
DRIVE.SetValue(False,'ExcludeExistingViolations')
DRIVE.SetValue(10.0,'MinTolerance')
DRIVE.SetValue('WholeNetwork','LargeDERDistribution')
DRIVE.SetValue(False,'UniformDERDistribution')
DRIVE.SetValue('Remove','BadImpedanceAction')
DRIVE.SetValue(100.0,'PowerFactor')
DRIVE.SetValue(120.0,'FaultContribution')
DRIVE.SetValue(60.0,'MaxDEROutputChange')
DRIVE.SetValue(100.0,'MaxDEROutputChange_AbnormalVoltages')
DRIVE.SetValue(True,'UseLoadModels')
DRIVE.SetValue(True,'SimulatePeakConditions')
DRIVE.SetValue(True,'SimulateOffPeakConditions')
DRIVE.SetValue(int(peakLoadID),'PeakLoadModelID')
DRIVE.SetValue(int(lightLoadID),'MinLoadModelID')

# Verifications
DRIVE.SetValue(False,'VerifyPrimaryOverVoltageLoad')
DRIVE.SetValue(True,'VerifyPrimaryVoltageDeviationGen')
DRIVE.SetValue(False,'VerifyPrimaryUnderVoltageLoad')
DRIVE.SetValue(True,'VerifyPrimaryUnderVoltageGen')
DRIVE.SetValue(105.0,'OverVoltageLimit')
DRIVE.SetValue(95.0,'UnderVoltageLimit')
DRIVE.SetValue(False,'VerifyPrimaryVoltageDeviationLoad')
DRIVE.SetValue(True,'VerifyPrimaryVoltageDeviationGen')
DRIVE.SetValue(False,'VerifyRegulatorVoltageDeviation')
DRIVE.SetValue(3.0,'MaxVoltageDeviation')
# DRIVE.SetValue(50.0,'MaxRegulatorVoltageDeviation')
DRIVE.SetValue(10.0,'AllowableViolation_Thermal_Deviation')
DRIVE.SetValue(1.0,'AllowableViolation_AbnormalVoltage')
DRIVE.SetValue(False,'VerifyThermalLoadingLoad')
DRIVE.SetValue(True,'VerifyThermalLoadingGen')
DRIVE.SetValue(100.0,'ThermalLoadingDischarging')
DRIVE.SetValue(1.0,'ThermalLoadingMinRating')

#If not including protection
DRIVE.SetValue(False,'VerifyAdditionalElementFaultCurrent')
DRIVE.SetValue(False,'VerifySympatheticTrip')
DRIVE.SetValue(False,'VerifyProtectionReach')
DRIVE.SetValue(False,'VerifyUnintentionalIslanding')
DRIVE.SetValue(False,'VerifyReverseFlow')
DRIVE.SetValue(False,'VerifyOperationalFlexibility')
DRIVE.SetValue(False,'VerifyFlicker')

# # If including protection
DRIVE.SetValue(False,'VerifyAdditionalElementFaultCurrent')
# DRIVE.SetValue(10.0,'MaxFaultCurrentDeviation')
# DRIVE.SetValue(False,'VerifySympatheticTrip')
# DRIVE.SetValue(150.0,'BreakerZeroSequenceCurrent')
DRIVE.SetValue(False,'VerifyProtectionReach')
# DRIVE.SetValue(10.0,'MaxBreakerFaultCurrentDeviation')
DRIVE.SetValue(False,'VerifyUnintentionalIslanding')
DRIVE.SetValue(False,'VerifyReverseFlow')
DRIVE.SetValue(False,'VerifyOperationalFlexibility')
DRIVE.SetValue(False,'VerifyFlicker')
# DRIVE.SetValue(0.35,'FlickerPst')
# DRIVE.SetValue(0.75,'FlickerPowerChange')
# DRIVE.SetValue(0.2,'FlickerShapeFactor')
# DRIVE.SetValue(0.0256,'FlickerCurveValue')

#Options
DRIVE.SetValue(False,'VerifyVoltageUnbalanceLoad')
DRIVE.SetValue(False,'VerifyVoltageUnbalanceGen')
DRIVE.SetValue(False,'GlobalEditDER')
DRIVE.SetValue(True,'ConsiderGenForGenImpacts')
DRIVE.SetValue(False,'ConsiderGenForLoadImpacts')
DRIVE.SetValue(False,'ConsiderStorageChargingForGenImpacts')
DRIVE.SetValue(False,'ConsiderStorageChargingForLoadImpacts')
DRIVE.SetValue(False,'DeleteIntermediateFilesAfterCalculation')

# Reset the Maximum Large DER Penetration to 20MW for 23kV feeders
print('Initial value - MaxLargeDERPenetrationLowVoltage: ' + DRIVE.GetValue('MaxLargeDERPenetrationLowVoltage'))
print('')
DRIVE.SetValue(10000,'MaxLargeDERPenetrationLowVoltage')

###############################################################################

#%% Run DRIVE with new switch settings
print('Starting EPRI DRIVE Run')
print('')

DRIVE.Run(feeders)

# Save EPRI DRIVE Report as .xlsx file
# Specify the name of the desired report - this will need to be in the list cympy.rm.ListReports()
report_name = 'Hosting Capacity Summary Report (Powered by EPRI DRIVE™)'  
# Specify the type of report as MSExcel to produce a .xlsx file
report_type_save = cympy.enums.ReportModeType.MSExcel
# Note that the path here must include the filename as well as the folder path

hc1Filename = r'\HCReport_Initial.xlsx'
savePathHC = saveResultsFolder + hc1Filename
cympy.rm.Save(report_name, feeders, report_type_save,savePathHC)


# Load report and calculate average HC across all feeders
hcData = pd.read_excel(savePathHC,header=None)

maxDERValues_Dist = []
maxDERValues_Cent = []
for rowCtr in range(0,hcData.shape[0]):
    currRow = hcData.loc[rowCtr,:]
    index1 = np.where(np.array(currRow)=='Hosting Capacity')[0]
    if len(index1) > 0:
        valueDist = currRow.loc[3]
        valueCent = currRow.loc[5]
        maxDERValues_Dist.append(valueDist)
        maxDERValues_Cent.append(valueCent)
maxDistAvg = np.round(np.mean(maxDERValues_Dist),decimals=2)
maxCentAvg = np.round(np.mean(maxDERValues_Cent),decimals=2)

print('Calulating the HC results from the output file of the intial run of EPRI DRIVE')
print('The Average Max Distributed DER Before Running Optimizer is ' + str(maxDistAvg))
print('The Average Max Centralized DER Before Running Optimizer is ' + str(maxCentAvg))
print('')


###############################################################################


# To save out a new study after making changes
newStudyFilename = '\\newStudy2.sxst'

studyFilePath = saveResultsFolder + newStudyFilename
cympy.study.Save(studyFilePath,True,True,True)


