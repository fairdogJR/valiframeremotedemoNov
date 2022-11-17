import sys
import platform
import string
import platform

#import pcievaliframeTemplate

print(platform.python_version())
print(platform.architecture())

################################################## Constants ##################################################

# the directory where the ValiFrame DLLs are; set to None to use the working directory
ValiFrameDllDirectory = r'C:\Program Files\BitifEye\ValiFrameK1\PCIe\TestAutomation'

# set to False to let the user choose an application (if there is more than one available); otherwise, put app name here
ForceApplication = False

# set to True to ask the user to change application settings before the application is configured
AskUserToChangeApplicationPropertiesBeforeConfig = True

# set to True to ask the user to change application settings after the application is configured
AskUserToChangeApplicationPropertiesAfterConfig = False

# set to True to ask the user to change procedure settings before running a procedure
AskUserToChangeProcedureProperties = False


# set to True to always show the "Configure DUT" dialog, False to not show it, or None to let the user decide
ShowConfigDialogPreference = False

# a list of procedure IDs to run; set to None to let the user decide
ProcedureIdsToAutoExecute = None
#ProcedureIdsToAutoExecute = [438410] #- 438410: "32G LEQ Rx Compliance Test"


# set to True to confirm all dialogs automatically, set to False to ask the user each time
AutoConfirmAllDialogs = True

# set to True to always show the XML result, False to not show it, or None to let the user decide
# This includes the output results
ShowXmlPreference = True

# set to True if you want to show script events in the console
ScriptLogToConsole = True

# set to True to show ValiFrame log entries in the console
ValiFrameLogToConsole = True

# set to True to also show ValiFrame internal log messages in the console
ValiFrameLogInternalToConsole = False

# the ValiFrame log file (all messages); set to None to disable logging
ValiFrameLogFile = None

# the ValiFrame XML result file (older ones are overwritten); set to None to disable saving
ValiFrameXmlResultFile = "resultsfile.xml"

# set to True to automatically close the script when everything was done; otherwise, wait for some user input
AutoCloseScript = True

############################################ Import .NET Namespaces ###########################################

import clr

if ValiFrameDllDirectory != None:
    sys.path.append(ValiFrameDllDirectory)
clr.AddReference(r'ValiFrameRemote')
clr.AddReference(r'VFBase')
#clr.AddReference(r'System.Windows.Forms')
from BitifEye.ValiFrame.ValiFrameRemote import *
from BitifEye.ValiFrame.Logging import *
from BitifEye.ValiFrame.Base import *
from System import *

################################################## Functions ##################################################

def ScriptLog(line):
    if ScriptLogToConsole:
        print(line)
    return


def IsIronPython():
    return platform.python_implementation() == 'IronPython'


def StartValiFrame():
    ScriptLog('Creating ValiFrame instance...')
    valiFrame = ValiFrameRemote(ProductGroupE.ValiFrameK1)
    return valiFrame


def UserBoolQuery(trueAnswer):
    userInput = input()
    return (userInput == trueAnswer)


def LogEntryChangedHandler(logEntry):
    message = logEntry.Text
    severity = None
    isInternal = False
    if logEntry.Severity == VFLogSeverityE.Internal:
        severity = 'Internal'
        isInternal = True
    elif logEntry.Severity == VFLogSeverityE.Info:
        severity = 'Info'
    elif logEntry.Severity == VFLogSeverityE.Progress:
        severity = 'Progress'
    elif logEntry.Severity == VFLogSeverityE.Warning:
        severity = 'Warning'
    elif logEntry.Severity == VFLogSeverityE.Critical:
        severity = 'Critical'
    elif logEntry.Severity == VFLogSeverityE.Exception:
        severity = 'Exception'
    else:
        severity = 'Unknown severity'

    if ValiFrameLogToConsole:
        show = True
        if isInternal:
            show = ValiFrameLogInternalToConsole
        if show:
            print('Log (%s): %s' % (severity, message))

    if ValiFrameLogFile != None:
        with open(ValiFrameLogFile, 'a') as file:
            file.write('%s: %s\n' % (severity, message))
    return


def StatusChangedHandler(sender, description):
    print('Status changed: %s' % description)
    return


def ProcedureCompletedHandler(procedureId, xmlResult):
    print('Procedure %d is complete' % procedureId)
    ScriptLog('Saving XML result to file...')
    if ValiFrameXmlResultFile != None:
        with open(ValiFrameXmlResultFile, 'w') as file:
            file.write(xmlResult)
    showXml = ShowXmlPreference
    if showXml == None:
        print('Do you want to see the XML result (y/n)?')
        showXml = UserBoolQuery('y')
    if showXml:
        print(xmlResult)
    return


def DialogPopUpHandler(sender, dialogInformation):
    print('Popup: %s' % dialogInformation.DialogText)
    abort = False
    if not AutoConfirmAllDialogs:
        print('Continue (y/n)?')
        abort = UserBoolQuery('n')
    if abort:
        dialogInformation.Dialog.DialogResult = System.Windows.Forms.DialogResult.Cancel;
    return


def RegisterEventHandlers(valiFrame):
    valiFrame.LogEntryChanged += LogEntryChangedEventHandler(LogEntryChangedHandler)
    valiFrame.StatusChanged += StatusChangedEventHandler(StatusChangedHandler)
    valiFrame.ProcedureCompleted += ProcedureCompletedEventHandler(ProcedureCompletedHandler)
    valiFrame.DialogPopUp += DialogShowEventHandler(DialogPopUpHandler)
    return


def UnregisterEventHandlers(valiFrame):
    try:
        valiFrame.LogEntryChanged -= LogEntryChangedEventHandler(LogEntryChangedHandler)
        valiFrame.StatusChanged -= StatusChangedEventHandler(StatusChangedHandler)
        valiFrame.ProcedureCompleted -= ProcedureCompletedEventHandler(ProcedureCompletedHandler)
        valiFrame.DialogPopUp -= DialogShowEventHandler(DialogPopUpHandler)
    except:
        return  # ignore errors
    return


def SelectApplication(valiFrame):
    if ForceApplication != False:
        return ForceApplication

    ScriptLog('Getting list of available applications...')
    applicationNames = valiFrame.GetApplications()

    if len(applicationNames) < 1:
        raise RuntimeError('No applications found')
    elif len(applicationNames) == 1:
        ScriptLog('Selecting application %s (no others available)' % applicationNames[0])
        return applicationNames[0]
    else:
        while True:
            print('The following %d applications were found:' % len(applicationNames))
            for applicationName in applicationNames:
                print('- "%s"' % applicationName)
            print('Please select one (leave blank to selecte the first one): ')
            selectedApplicationName = input()
            if selectedApplicationName == '':
                return applicationNames[0]
            for applicationName in applicationNames:
                if applicationName == selectedApplicationName:
                    return applicationName
            print('Invalid selection')


def InitApplication(valiFrame, applicationName):
    ScriptLog('Initializing Application...')
    valiFrame.InitApplication(applicationName)
    return


def ConfigureApplication(valiFrame):
    showDialog = ShowConfigDialogPreference
    if showDialog == None:
        print('Do you want to show the GUI dialog to configure the application (y/n)?')
        showDialog = not UserBoolQuery('n')
    if showDialog:
        ScriptLog('Configuring application with GUI dialog...')
        valiFrame.ConfigureProduct()
    else:
        ScriptLog('Configuring application automatically...')
        valiFrame.ConfigureProductNoDialog()
    return


def GetAvailableApplicationProperties(valiFrame):
    ScriptLog('Getting a list of available application properties...')
    if IsIronPython():
        propertiesClr = valiFrame.GetApplicationPropertiesList()
        properties = dict(propertiesClr)
    else:
        propertiesClr = valiFrame.GetApplicationPropertiesList()
        properties = {}
        for prop in propertiesClr:
            print("{} : {}".format(prop.Key, prop.Value))
            #            keyValuePair = prop.split(',')
            #            key = keyValuePair[0][1:].strip()
            #            value = keyValuePair[1][0:len(keyValuePair[1])-1].strip()
            properties[prop.Key] = prop.Value
    return properties



def LetUserChangeApplicationProperties(valiFrame):
    properties = GetAvailableApplicationProperties(valiFrame)
    print('Available application properties:')
    for propertyKey in properties:
        print('- %s: %s' % (propertyKey, properties[propertyKey]))
    print('Do you want to change an application property (y/n)?')
    change = UserBoolQuery('y')
    while change:
        print('Please enter the name of the application property:')
        requestedPropertyName = input()
        okay = False
        for propertyKey in properties:
            if requestedPropertyName == propertyKey:
                print('Please enter the new value:')
                newValue = input()
                valiFrame.SetApplicationProperty(propertyKey, newValue)
                okay = True
        if okay:
            print('Do you want to change another application property (y/n)?')
            change = UserBoolQuery('y')
        else:
            print('Invalid input')
    return


def ChangePropertiesBeforeConfiguration(valiFrame):
    if AskUserToChangeApplicationPropertiesBeforeConfig:
        LetUserChangeApplicationProperties(valiFrame)
    return


def ChangePropertiesAfterConfiguration(valiFrame):
    if AskUserToChangeApplicationPropertiesAfterConfig:
        LetUserChangeApplicationProperties(valiFrame)
    return


def GetAvailableProcedures(valiFrame):
    ScriptLog('\nGetting list of available procedures...')
    if IsIronPython():
        procedureIds, procedureNames = valiFrame.GetProcedures()
    else:
        dummyIntArray = (-1, 1)  # a dummy array of signed ints
        dummyStrArray = ('dummy')  # a dummy array of strings
        dummyOut, procedureIds, procedureNames = valiFrame.GetProcedures(dummyIntArray, dummyStrArray)
    return procedureIds, procedureNames


def GetAvailableProcedureProperties(valiFrame, procedureId):
    ScriptLog('\nGetting a list of available procedure properties for procedure %d...' % procedureId)
    properties = valiFrame.GetProcedureProperties(procedureId)

    # convert to generic hash
    flatProcedureProperties = {}
    for obj in properties:
        propertyName = str(obj.Name)
        propertyValue = str(obj)
        flatProcedureProperties[propertyName] = propertyValue

    return flatProcedureProperties


def GetAvailableRelatedProperties(valiFrame, procedureId):
    ScriptLog('\nGetting a list of available RELATED procedure properties for procedure %d...' % procedureId)
    if IsIronPython():
        propertiesClr = valiFrame.GetProcedureRelatedProperties(procedureId)
        properties = dict(propertiesClr)
    else:
        propertiesClr = valiFrame.GetProcedureRelatedProperties(procedureId)
        properties = {}
        for prop in propertiesClr:
            print("{} : {}".format(prop.Name, prop.Value))
            properties[prop.Name] = prop.Value
    return properties


def ChangeProcedureProperties(valiFrame, procedureId):
    if AskUserToChangeProcedureProperties:
        properties = GetAvailableProcedureProperties(valiFrame, procedureId)
        print('Available properties for procedure %d:' % procedureId)
        for propertyKey in properties:
            print('- %s: %s' % (propertyKey, properties[propertyKey]))
        print('Do you want to change a procedure property (y/n)?')
        change = UserBoolQuery('y')
        while change:
            print('Please enter the name of the procedure property:')
            requestedPropertyName = input()
            okay = False
            for propertyKey in properties:
                if requestedPropertyName == propertyKey:
                    print('Please enter the new value:')
                    newValue = input()
                    valiFrame.SetProcedureProperty(procedureId, propertyKey, newValue)
                    okay = True
            if okay:
                print('Do you want to change another procedure property (y/n)?')
                change = UserBoolQuery('y')
            else:
                print('Invalid input')
    return


def ChangeRelatedProperties(valiFrame, procedureId):
    if AskUserToChangeProcedureProperties:
        properties = GetAvailableRelatedProperties(valiFrame, procedureId)
        print('\nAvailable RELATED properties for procedure %d:' % procedureId)
        for propertyKey in properties:
            print('- %s: %s' % (propertyKey, properties[propertyKey]))
        print('Do you want to change a RELATED procedure property (y/n)?')
        change = UserBoolQuery('y')
        while change:
            print('Please enter the name of the RELATED procedure property:')
            requestedPropertyName = input()
            okay = False
            for propertyKey in properties:
                if requestedPropertyName == propertyKey:
                    print('Please enter the new value:')
                    newValue = input()
                    print("Entered Property Value: {}".format(newValue))
                    valiFrame.SetProcedureProperty(procedureId, propertyKey, newValue)
                    okay = True
            if okay:
                print('Do you want to change another RELATED procedure property (y/n)?')
                change = UserBoolQuery('y')
            else:
                print('Invalid input')
    return


def SelectProcedure(valiFrame):
    procedureIds, procedureNames = GetAvailableProcedures(valiFrame)
    if len(procedureIds) < 1:
        raise RuntimeError('No procedures found')
    elif len(procedureIds) == 1:
        ScriptLog('Selecting procedure %d (no others available)' % procedureIds[0])
        return procedureIds[0]
    else:
        while True:
            print('The following %d procedures were found:' % len(procedureIds))
            i = 0
            while i < len(procedureIds):

                print('- %d: "%s"' % (procedureIds[i], procedureNames[i]))
                i += 1
            print('Please select an ID: ')
            selectedProcedureIdStr = input()
            for procedureId in procedureIds:
                if str(procedureId) == selectedProcedureIdStr:
                    return procedureId
            print('Invalid selection')


def RunProcedure(valiFrame, procedureId):
    valiFrame.RunProcedure(procedureId)
    return


def RunProcedures(valiFrame):
    if ProcedureIdsToAutoExecute != None:
        for procedureId in ProcedureIdsToAutoExecute:
            ChangeProcedureProperties(valiFrame, procedureId)
            RunProcedure(valiFrame, procedureId)
    else:
        continueTesting = True
        while continueTesting:
            procedureId = SelectProcedure(valiFrame)
            ChangeProcedureProperties(valiFrame, procedureId)
            ChangeRelatedProperties(valiFrame, procedureId)
            RunProcedure(valiFrame, procedureId)
            print('Run another procedure (y/n)?')
            continueTesting = UserBoolQuery('y')
    return


def FinishScript():
    if not AutoCloseScript:
        print('Press any key to exit')

        x = input()
#--------------------------
def dump_propertieslist():
    print("--------->Demo: Accessing Application Properties list (Configure Dut)")
    #breakpoint()  # break here to try out
    propertiesClr = valiFrame.GetApplicationPropertiesList()
    properties = {}
    for prop in propertiesClr:
        print("{} : {}".format(prop.Key, prop.Value))
        properties[prop.Key] = prop.Value
    print("--------->Finished Dumping Application Properties list (Configure Dut)")
    #breakpoint()
    #--------------------------

################################################ Main Script ################################################

try:

    valiFrame = StartValiFrame()
    RegisterEventHandlers(valiFrame)
    applicationName = SelectApplication(valiFrame)
    InitApplication(valiFrame, applicationName)

    dump_propertieslist()

    #Method 2 configuration, load a pre-saved Dut configuration
    #loading all speeds checked  addin  asic expert mode (to expose more tests)
    valiFrame.LoadProject(r'C:\valiframe_projects\pcie5expert32g.vfp')
    #breakpoint()
    #valiFrame.LoadProject(r'C:\valiframe_projects\pcie5System.vfp')

    #demo Lets change an application property before configuring
    propertyKey='Lane3'
    newValue='True'
    valiFrame.SetApplicationProperty(propertyKey, newValue)

    dump_propertieslist()
    #breakpoint()

    #ChangePropertiesBeforeConfiguration(valiFrame)
    #ConfigureApplication(valiFrame)
    #ChangePropertiesAfterConfiguration(valiFrame)

    valiFrame.ConfigureProductNoDialog()


    #-------This lists available tests and calibrations-------------------
    print("--------->Demo: Accessing Procedure name list")
    #breakpiont()
    procedureIds, procedureNames = GetAvailableProcedures(valiFrame)
    i=0
    while i < len(procedureIds):
        print('- %d: "%s"' % (procedureIds[i], procedureNames[i]))
        i += 1
    print("--------->Finished Accessing Procedure name list")
    #--------------------------

    #breakpoint()# break here to try out
    procId=[] #make a list of procedures
    #pick a procedure
    procId.append(438410) #- 438410: "32G LEQ Rx Compliance Test"
    #read procedure properties
    propsLeqRXCompliance=GetAvailableProcedureProperties(valiFrame, procId[0])
    print(propsLeqRXCompliance)
    print("------->End of - 438410: 32G LEQ Rx Compliance Test properties")

    #lets look at another test in this same hierarchy with different properties
    procId.append(438411) # - 438411: "32G LEQ Rx Jitter Tolerance Test"
    # read procedure properties
    propsJtol = GetAvailableProcedureProperties(valiFrame, procId[1])
    print(propsJtol)
    print ("------->End of - 438411: 32G LEQ Rx Jitter Tolerance Test properties")

    #lets look at yet another test in this same hierarchy with different properties
    procId.append(438414)  # - 438414: "32G LEQ Rx Sensitivity Test"
    # read procedure properties
    propsLeqRXSens = GetAvailableProcedureProperties(valiFrame, procId[2])
    print(propsLeqRXSens)
    print ("------->End of - 438414: 32G LEQ Rx Sensitivity Test properties")

    #Change related property
    propertyKey='Repetitions'
    newValue='5'
    valiFrame.SetProcedureProperty(procId[1], propertyKey, newValue)
    #this is confusing, there is a global repetitions and a test local  repetitions, this seems to set the local repetitions
    #how to access the global repetitions?

    #what about use ssc

    relatedProps=[]
    for x in range (len (procId)):
            relatedProps.append (GetAvailableRelatedProperties(valiFrame, procId[x]))
    #break here to see that I had set repetitions in the procId[1] was set but the other related properties show repetitions of 0
    #lesson to learn Always check for the specific test
    #breakpoint()

    # what about use ssc? this is a related parameter at the top leve of the hierarchy, can I set it and is it seen as global for the whole hierarchy?
    #Change related property
    propertyKey='32 GT/s Use SSC'
    newValue='True'
    valiFrame.SetProcedureProperty(procId[0], propertyKey, newValue)

    relatedProps = []
    for x in range(len(procId)):
        relatedProps.append(GetAvailableRelatedProperties(valiFrame, procId[x]))
    # break here to if use ssc sticks in the procId[0] was set but the other related properties show repetitions of 0
    #in addition a dictionary can only have a unique Key so the global repeated name wont be used i assume?
    #YES it sets it and sets it globally
    #breakpoint()


    # TODO add precedure properties and related properties
    #RunProcedures(valiFrame)
    valiFrame.RunProcedure(procId[0])

    #breakpoint()

    UnregisterEventHandlers(valiFrame)


#todo cleanup code

except Exception as e:
    print('EXCEPTION: %s' % str(e))

FinishScript()

