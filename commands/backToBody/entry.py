import adsk.core
import os

import adsk.fusion
import adsk.fusion
from ...lib import fusionAddInUtils as futil
from ... import config
app = adsk.core.Application.get()
ui = app.userInterface

CMD_ID = f'{config.ADDIN_NAME}_cmdDialog'
CMD_NAME = 'Back to Body'
CMD_Description = 'Moves all the child bodies of the selected component to the root level.'
IS_PROMOTED = True
WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'fusionToolKitPanel'
TAB_ID = 'SolidTab'
TAB_NAME =""
PANEL_NAME = 'Tools'
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')
COMMAND_BESIDE_ID = ''

local_handlers = []

def start():
    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)

    futil.add_handler(cmd_def.commandCreated, command_created)

    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    tab:adsk.core.ToolbarTab = workspace.toolbarTabs.itemById(TAB_ID)

    if not tab:
        tab = workspace.toolbarTabs.add(TAB_ID, TAB_NAME)

    panel = tab.toolbarPanels.itemById(PANEL_ID)
    if not panel:
        panel = tab.toolbarPanels.add(PANEL_ID, PANEL_NAME, '', False)

    control = panel.controls.addCommand(cmd_def, COMMAND_BESIDE_ID, False)
    control.isPromoted = IS_PROMOTED

def stop():
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    tab = workspace.toolbarTabs.itemById(TAB_ID)
    panel = tab.toolbarPanels.itemById(PANEL_ID)
    command_control = panel.controls.itemById(CMD_ID)
    command_definition = ui.commandDefinitions.itemById(CMD_ID)

    # Delete the button command control
    if command_control:
        command_control.deleteMe()

    # Delete the command definition
    if command_definition:
        command_definition.deleteMe()

    if panel:
        panel.deleteMe()

def command_created(args: adsk.core.CommandCreatedEventArgs):

    commandInputs = args.command.commandInputs

    selectionInput = commandInputs.addSelectionInput('component', 'Component', 'Select the component to move the bodies from.')
    selectionInput.addSelectionFilter('Occurrences')
    selectionInput.setSelectionLimits(1, 1)

    futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(args.command.inputChanged, command_input_changed, local_handlers=local_handlers)
    futil.add_handler(args.command.executePreview, command_preview, local_handlers=local_handlers)
    futil.add_handler(args.command.validateInputs, command_validate_input, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)


def command_execute(args: adsk.core.CommandEventArgs):
    commandInput = args.command.commandInputs

    product = app.activeDocument.products.itemByProductType('DesignProductType')
    design = adsk.fusion.Design.cast(product)
    rootComp = design.rootComponent

    # Get the selected component
    selections:adsk.core.SelectionCommandInput = commandInput.itemById('component')
    selectedOcc = selections.selection(0).entity

    occs = []
    if selectedOcc.classType() == adsk.fusion.Occurrence.classType():
        getAllOccs(selectedOcc.childOccurrences, occs)
    else:
        getAllOccs(selectedOcc.occurrences, occs)

    occ: adsk.fusion.Occurrence
    for occ in occs:
        for body in occ.bRepBodies:
             body.copyToComponent(selectedOcc)
    
    for occ in reversed(occs):
        if design.designType == adsk.fusion.DesignTypes.DirectDesignType:
            occ.deleteMe()
        else:
            features_ = rootComp.features.removeFeatures
            features_.add(occ)

def command_preview(args: adsk.core.CommandEventArgs):
    inputs = args.command.commandInputs

def command_input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = args.input

def command_validate_input(args: adsk.core.ValidateInputsEventArgs):
    inputs = args.inputs
    
def command_destroy(args: adsk.core.CommandEventArgs):
    global local_handlers
    local_handlers = []

def getAllOccs(occurances: adsk.fusion.Occurrences, occs: list):
    for occ_ in occurances:
        occs.append(occ_)
        getAllOccs(occ_.childOccurrences, occs)
