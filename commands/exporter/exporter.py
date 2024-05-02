#Author-Boopathi
#Description-Exports DMT file 
import subprocess
import time
import adsk.core, adsk.fusion, adsk.cam, traceback, os, datetime
import winreg

DETACHED_PROCESS = 0x00000008
# Global list to keep all event handlers in scope.
# This is only needed with Python.
handlers = []

def start(context):
    global ui
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        # Get the CommandDefinitions collection.
        cmdDefs = ui.commandDefinitions

        # CNC Export Toolbar codes
        dgkExportDef = cmdDefs.addButtonDefinition('ExportDGK','DGK Export','Export Fusion file into DGK file.','')
        
        clickDGKExport = dgkExportCreateHandler()
        dgkExportDef.commandCreated.add(clickDGKExport)
        handlers.append(clickDGKExport)

        toolbars_ = ui.toolbars
        toolbarQAT_ = toolbars_.itemById('QAT')
        toolbarControls_ = toolbarQAT_.controls
        fileDropDown = toolbarControls_.itemById('FileSubMenuCommand')   
        FiletoolbarControls_ = fileDropDown.controls

        FiletoolbarControls_.addCommand(dgkExportDef)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def getReg(path_,value_):
    try:
        if str(path_).startswith("HKEY_LOCAL_MACHINE"):
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, str(path_).split("HKEY_LOCAL_MACHINE\\")[1], 0, winreg.KEY_READ)

        exe = (winreg.QueryValueEx(key, value_)[0])
        winreg.CloseKey(key)
        return exe
    except:
        return False

def getUtilityPath():
    curVersion = datetime.datetime.today().year + 1
    defaultPath = []
    for i in range(3):
        defaultPath.append("HKEY_LOCAL_MACHINE\\SOFTWARE\\Autodesk\Manufacturing Data Exchange Utility\\{}".format(str(curVersion - i)))
        
    registryPaths = defaultPath + [
        "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{7881FDE3-C0A4-4981-A045-76E261AAAB72}",
        "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{C4BEB98E-7E1B-437E-918A-0FC681489C36}",
        "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{89B27082-4DB1-43CB-8D2F-2CDF3F019BBB}",
        "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{97597B65-9B7B-4144-9C14-131AF757B0D6}"
    ]

    found = False
    for i in registryPaths:
        exeDir = getReg(i, "InstallLocation")
        if exeDir and os.path.isdir(exeDir):
            exePath = os.path.join(exeDir, 'sys', 'exec64', 'sdx.exe')
            if exePath and os.path.isfile(exePath):
                found = True
                break

    if not found:
        criticalError("Autodesk Manufacturing Data Exchange Utility is not installed.", "Export Error")

    return exePath
        
def criticalError(message , title):
    ui.messageBox(message,title,
        adsk.core.MessageBoxButtonTypes.OKButtonType,
        adsk.core.MessageBoxIconTypes.CriticalIconType)
    return 

def sucessMessage(message , title):
    ui.messageBox(message,title,
        adsk.core.MessageBoxButtonTypes.OKButtonType,
        adsk.core.MessageBoxIconTypes.InformationIconType)
    return 

class dgkExportCreateHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            global  opFolder, utilityPath
            app = adsk.core.Application.get()
            ui  = app.userInterface
            eventArgs = adsk.core.CommandCreatedEventArgs.cast(args)
            cmd = eventArgs.command
            if (not ui.activeWorkspace.id == "FusionSolidEnvironment"):
                modelWS = ui.workspaces.itemById("FusionSolidEnvironment")
                modelWS.activate()    
            

            utilityPath = getUtilityPath()

            fldrDlg = ui.createFolderDialog()
            fldrDlg.title = "Select the Output folder"

            fldrDlgRes = fldrDlg.showDialog()
            if fldrDlgRes == adsk.core.DialogResults.DialogOK:
                opFolder =fldrDlg.folder
            else:
                return
            
            onExecute = dgkExportExecuteHandler()
            cmd.execute.add(onExecute)
            handlers.append(onExecute)

            #onInputChanged = InputChangedHandler()
            #cmd.inputChanged.add(onInputChanged)
            #handlers.append(onInputChanged)
 
        except:
         if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def formatString(txt):
    temp = str(txt).replace(" ", "_")
    temp = temp.split(":")[0]
    return temp

class dgkExportExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            command = args.firingEvent.sender
            app = adsk.core.Application.get()

            design: adsk.fusion.Design = app.activeDocument.products.itemByProductType("DesignProductType")
            exportMgr = design.exportManager

            rootComp = design.rootComponent

            processDialog = ui.createProgressDialog()
            processDialog.show('DGK File export','Exporting files...',0 , 10)
            processDialog.isCancelButtonShown = False

            compName = app.activeDocument.name 
            tempStep =  os.path.join(opFolder, compName + '.step')

            stepOptions = exportMgr.createSTEPExportOptions(tempStep, rootComp)
            res = exportMgr.execute(stepOptions)

            exportFile = os.path.join(opFolder , compName+'.dgk')

            if res:
                a = ' "{}" -tol 0.05 -t dgk -o "{}"'.format(tempStep, exportFile)
                proc = subprocess.Popen(utilityPath + a, stdout=subprocess.PIPE,text=True, creationflags=DETACHED_PROCESS)
                
                while proc.poll() is None:
                    time.sleep(1)
                    if processDialog.wasCancelled:
                        proc.kill()
                        processDialog.hide()
                        return

                    if (processDialog.progressValue < processDialog.maximumValue):
                        processDialog.progressValue += 1
                    else:
                        processDialog.progressValue = processDialog.minimumValue
                    
                proc.terminate()

            processDialog.progressValue = processDialog.maximumValue
            os.remove(tempStep)  # Remove Temporary step Files
            processDialog.hide()
            sucessMessage("Exported DGK file successfully at {}".format(exportFile), "Export Success")

        except:
         if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def stop(context):
    try:
        
        # Clean up the UI.
        cmdDef1 = ui.commandDefinitions.itemById('ExportDGK')
        if cmdDef1:
            cmdDef1.deleteMe()
        
        toolbars_ = ui.toolbars
        toolbarQAT_ = toolbars_.itemById('QAT')
        toolbarControls_ = toolbarQAT_.controls
        fileDropDown = toolbarControls_.itemById('FileSubMenuCommand')   
        FiletoolbarControls_ = fileDropDown.controls

        cntrl1 = FiletoolbarControls_.itemById('ExportDGK')
        if cntrl1:
            cntrl1.deleteMe()

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))	