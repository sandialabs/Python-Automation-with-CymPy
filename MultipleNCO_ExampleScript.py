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

###               Multiple Optimization Study Script             ###


# This script runs a loop with different network configuration optimization 
#    objective functions and reports the hosting capacity before and after
#    each optimization

# The workflow is:
    #  1.  Load .sxst model using CymPy library
    #  2.  Run EPRI DRIVE to determine intial hosting capacity
    #  3.  Loop through Network Configuration Optimization Tool objectives
    #  4.  Run EPRI DRIVE again after each NCO run to determine hosting 
    #         capacity after the changes suggested by the NCO are applied
    
    
# The results of steps 2-4 are saved as .xlrd or .csv files


#%% Python Library Imports
import numpy as np
import pandas as pd
import cympy
import cympy.rm
import locale
#import xlrd

###############################################################################

#%%  Set directory paths and filenames

# Location and name of .sxst file
studyFolderPath = r'C:\<Path>\<To>\<Study>\<Folder>'
studyFilename = r'\studyFile.sxst'

# Folder to save .xlrd and .csv results
saveResultsFolder = r'C:\<Path>\<To>\<Save\<Results>'



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
# In this example, custom load models were created for peak and light load conditions
# Those load models (index 4 and 5, respectively) were hard-coded in this example, since they are required inputs for EPRI DRIVE
# However, this code shows how you would access the various load models available in a CYME study 
LoadModels = cympy.study.ListLoadModels()
LoadModelNames = []
LoadModelIDs = np.zeros((len(LoadModels),1),dtype=int)
for ii in range(len(LoadModels)):
    LoadModelIDs[ii]=LoadModels[ii].ID
    LoadModelNames.append(LoadModels[ii].Name)
peakLoadName = LoadModelNames[4]
peakLoadID = LoadModelIDs[4,0]
lightLoadName = LoadModelNames[5]
lightLoadID = LoadModelIDs[5,0]

feeders = cympy.study.ListNetworks(cympy.enums.NetworkType.Feeder)

###############################################################################


#%%  Get Initial Switch State List



# The below includes Switch, Recloser, and Breaker


switchList = cympy.study.ListDevices(cympy.enums.DeviceType.Switch)

breakerList = cympy.study.ListDevices(cympy.enums.DeviceType.Breaker)
recloserList = cympy.study.ListDevices(cympy.enums.DeviceType.Recloser)
fuses = cympy.study.ListDevices(cympy.enums.DeviceType.Fuse)

switchingDeviceTypes = []
# From switch list, get a list of all switch ids
switchIDs = []
switchStatus = []
for switchCtr in range(0,len(switchList)):
    currSwitch = switchList[switchCtr]
    currID = currSwitch.GetValue('DeviceNumber')
    #currState = cympy.study.QueryInfoDevice("EqState",currID,cympy.enums.DeviceType.Switch)
    currState = currSwitch.GetValue('ClosedPhase')
    if currState == 'None':
        currState = 'Open'
    else:
        currState = 'Close'
    switchIDs.append(currID)
    switchStatus.append(currState)
    switchingDeviceTypes.append('Switch')


# From breaker list, get a list of all breaker ids
breakerIDs = []
breakerStatus = []
for breakerCtr in range(0,len(breakerList)):
    currBreaker = breakerList[breakerCtr]
    currID = currBreaker.GetValue('DeviceNumber')
    #currState = cympy.study.QueryInfoDevice("EqState",currID,cympy.enums.DeviceType.Breaker)
    currState = currSwitch.GetValue('ClosedPhase')
    if currState == 'None':
        currState = 'Open'
    else:
        currState = 'Close'    
    breakerIDs.append(currID)
    breakerStatus.append(currState)
    switchingDeviceTypes.append('Breaker')


# From recloser list, get a list of all recloser ids and states
recloserIDs = []
recloserStatus = []
for recloserCtr in range(0,len(recloserList)):
    currRecloser = recloserList[recloserCtr]
    currID = currRecloser.GetValue('DeviceNumber')
    #currState = cympy.study.QueryInfoDevice("EqState",currID,cympy.enums.DeviceType.Recloser)
    currState = currSwitch.GetValue('ClosedPhase')
    if currState == 'None':
        currState = 'Open'
    else:
        currState = 'Close'
    recloserIDs.append(currID)
    recloserStatus.append(currState)
    switchingDeviceTypes.append('Recloser')
    
    
# Note:  We chose to omit fuses from consideration, but those could be added with the same syntax,
#           simply subsitute the keyword Fuse for Recloser


allSwitchingDeviceIDs = switchIDs + breakerIDs + recloserIDs
allSwitchingStates = switchStatus + breakerStatus + recloserStatus


# Write Switch states to csv
df = pd.DataFrame()
df['Switch ID'] = np.array(allSwitchingDeviceIDs)
df['Status'] = allSwitchingStates
df['Type'] = switchingDeviceTypes
filename = '\SwitchingDevicesStates_Initial.csv'
filePath = saveResultsFolder + filename
df.to_csv(filePath)
print('')



# Notes:
#   Both of these methods will get the current status of a switching device
#    currState = cympy.study.QueryInfoDevice("EqState",currID,cympy.enums.DeviceType.Switch)
#    currState = currSwitch.GetValue('ClosedPhase')

#   currSwitch.GetValue('ClosedPhase') will return 'None' if the switch is open and 
#       'ABC', 'A', 'B', 'C', 'AB', 'AC', 'BC' if closed, depending on the phases which
#       The switching device is connected to.  We converted this to open/close because
#       it's a little easier to manually create a list of switch statuses without knowing
#       all of the phase connections - for use in the SetSwitchesRunDrive_Script - since
#       the phase information can be extracted from the device information.  Directly 
#       saving the phase information may useful though



#breakers = cympy.study.ListDevices(cympy.enums.DeviceType.Breaker)
#reclosers = cympy.study.ListDevices(cympy.enums.DeviceType.Recloser)
#fuses = cympy.study.ListDevices(cympy.enums.DeviceType.Fuse)
###############################################################################


#%% Set Up EPRI DRIVE Parameters 

DRIVE = cympy.sim.EPRIDrive()  # Assigns the specific simulation a variable name
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
DRIVE.SetValue(False,'VerifyPrimaryUnderVoltageGen')
DRIVE.SetValue(105.0,'OverVoltageLimit')
# DRIVE.SetValue(95.0,'UnderVoltageLimit')
DRIVE.SetValue(False,'VerifyPrimaryVoltageDeviationLoad')
DRIVE.SetValue(True,'VerifyPrimaryVoltageDeviationGen')
DRIVE.SetValue(True,'VerifyRegulatorVoltageDeviation')
DRIVE.SetValue(3.0,'MaxVoltageDeviation')
DRIVE.SetValue(50.0,'MaxRegulatorVoltageDeviation')
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
# DRIVE.SetValue(True,'VerifyAdditionalElementFaultCurrent')
# DRIVE.SetValue(10.0,'MaxFaultCurrentDeviation')
# DRIVE.SetValue(True,'VerifySympatheticTrip')
# DRIVE.SetValue(150.0,'BreakerZeroSequenceCurrent')
# DRIVE.SetValue(True,'VerifyProtectionReach')
# DRIVE.SetValue(10.0,'MaxBreakerFaultCurrentDeviation')
# DRIVE.SetValue(False,'VerifyUnintentionalIslanding')
# DRIVE.SetValue(False,'VerifyReverseFlow')
# DRIVE.SetValue(False,'VerifyOperationalFlexibility')
# DRIVE.SetValue(True,'VerifyFlicker')
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
DRIVE.SetValue(20000,'MaxLargeDERPenetrationLowVoltage')

###############################################################################

#%% Run DRIVE with intial switch settings
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



#%% Run Network Configuration Optimization tool

nco = cympy.sim.NetworkConfigurationOptimization()
nco.GetObjType()  # command give the string to use for this module in the Describe function 
cympy.Describe('SOMParameters') # The Describe function provides the list of settable parameters for the specified module

# Print intial parameter values
print('Initial value - Objective: ' + nco.GetValue('Objective'))
print('Initial value - Method: ' + nco.GetValue('Method'))
print('Initial value - ObjectiveLosses: ' + nco.GetValue('ObjectiveLosses'))
print('Initial value - AcceleratedSearch: ' + nco.GetValue('AcceleratedSearch'))
print('Initial value - InstallNewSwitch: ' + nco.GetValue('InstallNewSwitch'))
print('Initial value - LoadFlowParamConfigID: ' + nco.GetValue('LoadFlowParamConfigID'))
print('Initial value - ExcludedDeviceType: ' + nco.GetValue('ExcludedDeviceType'))
print('Initial value - OperateRemotelyControlled: ' + nco.GetValue('OperateRemotelyControlled'))
print('Initial value - ExcludedDevices: ' + nco.GetValue('ExcludedDevices'))
print('Initial value - AllowInitialViolation: ' + nco.GetValue('AllowInitialViolation'))
print('Initial value - IgnoreTieInSameTopo: ' + nco.GetValue('IgnoreTieInSameTopo'))
print('Initial value - ObjectiveOperations: ' + nco.GetValue('ObjectiveOperations'))
print('Initial value - ObjectiveLoadBalancing: ' + nco.GetValue('ObjectiveLoadBalancing'))
print('Initial value - ObjectiveDistance: ' + nco.GetValue('ObjectiveDistance'))
print('Initial value - ObjectiveVoltageExceptions: ' + nco.GetValue('ObjectiveVoltageExceptions'))
print('Initial value - ObjectiveLosses: ' + nco.GetValue('ObjectiveLosses'))
print('Initial value - ObjectiveOverload: ' + nco.GetValue('ObjectiveOverload'))
print('Initial value - ObjectiveWeightOperations: ' + nco.GetValue('ObjectiveWeightOperations'))
print('Initial value - ObjectiveWeightLoadBalancing: ' + nco.GetValue('ObjectiveWeightLoadBalancing'))
print('Initial value - ObjectiveWeightDistance: ' + nco.GetValue('ObjectiveWeightDistance'))
print('Initial value - ObjectiveWeightLosses: ' + nco.GetValue('ObjectiveWeightLosses'))
print('Initial value - ObjectiveWeightVoltageExceptions: ' + nco.GetValue('ObjectiveWeightVoltageExceptions'))
print('Initial value - ObjectiveWeightOverload: ' + nco.GetValue('ObjectiveWeightOverload'))
print('Initial value - EnableMinimumLoss: ' + nco.GetValue('EnableMinimumLoss'))
print('Initial value - MinimumLoss: ' + nco.GetValue('MinimumLoss'))
print('Initial value - EnableMinimumLoadingUnbalance: ' + nco.GetValue('EnableMinimumLoadingUnbalance'))
print('Initial value - MinimumLoadingUnbalance: ' + nco.GetValue('MinimumLoadingUnbalance'))
print('Initial value - EnableMinimumLengthUnbalance: ' + nco.GetValue('EnableMinimumLengthUnbalance'))
print('Initial value - MinimumLengthUnbalance: ' + nco.GetValue('MinimumLengthUnbalance'))
print('Initial value - EnableMinDistanceBetweenNewSwitch: ' + nco.GetValue('EnableMinDistanceBetweenNewSwitch'))
print('Initial value - MinDistanceBetweenNewSwitch: ' + nco.GetValue('MinDistanceBetweenNewSwitch'))
print('Initial value - EnableMaximumNumberSwitchingOperations: ' + nco.GetValue('EnableMaximumNumberSwitchingOperations'))
print('Initial value - MaximumNumberSwitchingOperations: ' + nco.GetValue('MaximumNumberSwitchingOperations'))

objectiveList = ['MinimizeLosses', 'MinimizeVoltageExceptions','MinimizeOverloadExceptions','BalanceLoad']
methodList = ['HeuristicLocal','HeuristicZones','HeuristicZones','HeuristicLocal']

noOpt = []
distHC = []
centHC = []

for objCtr in range(0,len(objectiveList)):
    currObj = objectiveList[objCtr]
    currMethod = methodList[objCtr]
    print('Starting ' + str(currObj) + ' Objective run')
    noOptFlag = False    
                     
    nco.SetValue(currObj, 'Objective')
    nco.SetValue(currMethod,'Method')
    
    
    # This tool takes the full list of networks including transmission lines
    #print('Run Network Configuration Optimization Tool')
    #print('')
    
    try:
        nco.Run(networks)
    except cympy.err.CymError as e:
        print('NCO message below (' + str(currObj) + ') :')
        print(e.GetMessage())
        noOptFlag = True
        noOpt.append(currObj)
    
    
    if not noOptFlag:
        # Save .xlsx report with NCO tool results
        report_name = 'Network Configuration Optimization - Summary'
        report_type_save = cympy.enums.ReportModeType.MSExcel
        report_type_show = cympy.enums.ReportModeType.CYMESpreadsheet
        filenameOpt = r'OptReport_' + str(currObj) + '_' + str(currMethod) + '.xlsx'
        savePathOpt = saveResultsFolder + filenameOpt
        cympy.rm.Save(report_name, networks, report_type_save,savePathOpt)
        
        
        # Using the list of ids, get the current/active switch state
               
        switchStatusAfter = []
        
        
        for switchCtr in range(0,len(switchList)):
            currSwitch = switchList[switchCtr]
            currID = currSwitch.GetValue('DeviceNumber')
            currState = cympy.study.QueryInfoDevice("EqState",currID,cympy.enums.DeviceType.Switch)
            switchStatusAfter.append(currState)
        
        
        # From breaker list, get a list of all breaker ids
        for breakerCtr in range(0,len(breakerList)):
            currBreaker = breakerList[breakerCtr]
            currID = currBreaker.GetValue('DeviceNumber')
            currState = cympy.study.QueryInfoDevice("EqState",currID,cympy.enums.DeviceType.Breaker)
            switchStatusAfter.append(currState)
        
        
        # From recloser list, get a list of all recloser ids and states
        recloserIDs = []
        recloserStatus = []
        for recloserCtr in range(0,len(recloserList)):
            currRecloser = recloserList[recloserCtr]
            currID = currRecloser.GetValue('DeviceNumber')
            currState = cympy.study.QueryInfoDevice("EqState",currID,cympy.enums.DeviceType.Recloser)
            switchStatusAfter.append(currState)
        
        

        
        # Write Switch states to csv
        df = pd.DataFrame()
        df['Switch ID'] = np.array(allSwitchingDeviceIDs)
        df['Status'] = switchStatusAfter
        df['Type'] = switchingDeviceTypes
        filenameSwitchAfter = '\SwitchDevicesAfter_' + str(currObj) + '_' + str(currMethod) + '.csv'
        filePathSwitchAfter = saveResultsFolder + filenameSwitchAfter
        df.to_csv(filePathSwitchAfter)
    
        
        print('Starting EPRI DRIVE Run')
        print('')
        
        DRIVE.Run(feeders)
        
        # Save EPRI DRIVE Report as .xlsx file
        # Specify the name of the desired report - this will need to be in the list cympy.rm.ListReports()
        report_name = 'Hosting Capacity Summary Report (Powered by EPRI DRIVE™)'  
        # Specify the type of report as MSExcel to produce a .xlsx file
        report_type_save = cympy.enums.ReportModeType.MSExcel
        # Note that the path here must include the filename as well as the folder path
        
        hc1Filename = r'\HCReport_' + str(currObj) + '_' + str(currMethod) + '.xlsx'
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
        
        distHC.append(maxDistAvg)
        centHC.append(maxCentAvg)
    # End of noOptFlag condition

# End of objCtr for loop


###############################################################################












