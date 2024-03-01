#Author-Samuel Stephenson
#Description-Create an Excel File Bill of Materials with part thumbnails

import adsk.core, adsk.fusion, traceback
import subprocess, os, platform
from shutil import copyfile
import time
import csv
import re
from datetime import date
from .Modules import xlsxwriter

# Globals
app = adsk.core.Application.get()
if app:
    ui = app.userInterface

addin_path = os.path.dirname(os.path.realpath(__file__)) 

defaultExportStep = False
defaultProjectName = ''
defaultProductName = ''
defaultOwner = ''
defaultDesigner = ''

cameraBackup = app.activeViewport.camera
gridVisibilityBackup = False

# global set of event handlers to keep them referenced for the duration of the command
handlers = []

def timing(f):
    def wrap(*args, **kwargs):
        time1 = time.time()
        ret = f(*args, **kwargs)
        time2 = time.time()
        print('{:s} function took {:.3f} ms'.format(f.__name__, (time2-time1)*1000.0))

        return ret
    return wrap

def run(context):
    try:
        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        if not design:
            ui.messageBox('It is not supported in current workspace, please change to MODEL workspace and try again.')
            return

        commandDefinitions = ui.commandDefinitions
        #check the command exists or not
        cmdDef = commandDefinitions.itemById('BOMshot')
        if not cmdDef:
            cmdDef = commandDefinitions.addButtonDefinition('BOMshot',
                    'BOMshot',
                    'Create an Excel File Bill of Materials with part thumbnails',
                    './resources') # relative resource file path is specified

        onCommandCreated = BOMCommandCreatedHandler()
        cmdDef.commandCreated.add(onCommandCreated)
        # keep the handler referenced beyond this function
        handlers.append(onCommandCreated)
        inputs = adsk.core.NamedValues.create()
        cmdDef.execute(inputs)

        # prevent this module from being terminate when the script returns, because we are waiting for event handlers to fire
        adsk.autoTerminate(False)
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

class BoltCommandDestroyHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            # when the command is done, terminate the script
            # this will release all globals which will remove all event handlers
            adsk.terminate()
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

class BOMCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):    
    def __init__(self):
        super().__init__()        
    def notify(self, args):
        try:
            cmd = args.command
            cmd.isRepeatable = False
            onExecute = BOMCommandExecuteHandler()
            cmd.execute.add(onExecute)
            #onExecutePreview = BOMCommandExecuteHandler()
            #cmd.executePreview.add(onExecutePreview)
            onDestroy = BoltCommandDestroyHandler()
            cmd.destroy.add(onDestroy)
            # keep the handler referenced beyond this function
            handlers.append(onExecute)
            #handlers.append(onExecutePreview)
            handlers.append(onDestroy)

            #define the inputs
            inputs = cmd.commandInputs

            inputs.addImageCommandInput('image', '', 'resources/Icon_128.png')
            inputs.addBoolValueInput('exportStep', 'Export STEP', True, "", defaultExportStep)
            inputs.addStringValueInput('projectName', 'Project Name', defaultProjectName)
            inputs.addStringValueInput('productName', 'Product Name', defaultProductName)
            inputs.addStringValueInput('owner', 'Owner', defaultOwner)
            inputs.addStringValueInput('designer', 'Designer', defaultDesigner)

        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

class BOMCommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            #cameraBackup = app.activeViewport.camera
            gridVisibilityBackup = isGridDisplayOn()
            
            command = args.firingEvent.sender
            inputs = command.commandInputs

            bom = BOM()
            for input in inputs:
                if input.id == 'exportStep':
                    bom.exportStep = input.value
                if input.id == 'projectName':
                    bom.projectName = input.value
                if input.id == 'productName':
                    bom.productName = input.value
                if input.id == 'owner':
                    bom.owner = input.value
                if input.id == 'designer':
                    bom.designer = input.value

            bom.extractBOM()
            args.isValidResult = True

            app.activeViewport.camera = cameraBackup
            setGridDisplay(gridVisibilityBackup)
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

        # Force the termination of the command.
        adsk.terminate()   

class BOM:
    def __init__(self):
        self._exportStep = defaultExportStep
        self._projectName = defaultProjectName
        self._productName = defaultProductName
        self._owner = defaultOwner
        self._designer = defaultDesigner

    #properties
    @property
    def exportStep(self):
        return self._exportStep
    @exportStep.setter
    def exportStep(self, value):
        self._exportStep = value

    @property
    def projectName(self):
        return self._projectName
    @projectName.setter
    def projectName(self, value):
        self._projectName = value

    @property
    def productName(self):
        return self._productName
    @productName.setter
    def productName(self, value):
        self._productName = value

    @property
    def owner(self):
        return self._owner
    @owner.setter
    def owner(self, value):
        self._owner = value
        
    @property
    def designer(self):
        return self._designer
    @designer.setter
    def designer(self, value):
        self._designer = value

    def addComponentToList(self, list, component, path):
        # Gather any BOM worthy values from the component
        
        if component.material is not None:
            info = {
                'component': component,
                'thumbnail': path + '/images/' + name(component.name)  + '.png',
                'name': name(component.name),
                'instances': 1,
                'material': component.material.name,
            }

            list.append(info)

    def collectInstance(self, list, occ, path):
        component = occ.component
        instanceExistInList = False
        for listI in list:
            if listI['component'] == component:
                # Increment the instance count of the existing row.
                listI['instances'] += 1
                instanceExistInList = True
                break

        if not instanceExistInList:
            self.addComponentToList(list, component, path)
            takePhoto(occ, path)
            if self._exportStep:
                write_component(path, component)

    def extractBOM(self):

        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        title = 'Extract BOM'
        if not design:
            ui.messageBox('No active design', title)
            return
        
        root = product.rootComponent
        design.activateRootComponent()

        # Build project info object
        projectInfo = {
            'projectName': self._projectName,
            'productName': self._productName,
            'owner': self._owner,
            'designer': self._designer,
            'rootImage': '',
            'logoImage': ''
        }
        
        fileDialog = ui.createFileDialog()
        fileDialog.isMultiSelectEnabled = False
        fileDialog.title = " filename"
        fileDialog.filter = 'xlsx (*.xlsx)'
        fileDialog.initialFilename =  product.rootComponent.name
        fileDialog.filterIndex = 0
        dialogResult = fileDialog.showSave()
        if dialogResult == adsk.core.DialogResults.DialogOK:
            filename = fileDialog.filename
            dst_directory = os.path.splitext(filename)[0] + '_files'
        else:
            return
        # Gather information about each unique component
        bom = []

        # The path where thumbnails will be saved, updated to use a dynamic base path
        base_path = dst_directory + '/' + name(design.activeComponent.name)

        takeRootPhoto(root, dst_directory)

        def processComponent(occurrences, path):
                for occ in occurrences:
                    self.collectInstance(bom, occ, path)
                    # Recursively process subcomponents if they exist
                    if occ.childOccurrences:
                        processComponent(occ.childOccurrences, path + '/' + name(occ.component.name))
    
        # Start processing from the root component
        processComponent(root.occurrences, base_path)

        if len(bom) == 0:
            ui.messageBox('No components found', 'BOMshot')
            return
        
        projectInfo['rootImage'] = dst_directory + '/root.png'
        projectInfo['logoImage'] = dst_directory + '/logo.png'
        
        Unisolate(root.occurrences)
        
        buildXLSX(bom, os.path.splitext(filename)[0], projectInfo)
        
        dialogResult = ui.messageBox('BOM Extracted. Open file?', 'BOMshot', adsk.core.MessageBoxButtonTypes.OKCancelButtonType, adsk.core.MessageBoxIconTypes.InformationIconType)
        if dialogResult == adsk.core.DialogResults.DialogOK:
            openWithDefaultApplication(filename)

def openWithDefaultApplication(filename):
    if platform.system() == 'Darwin':       # macOS
        subprocess.call(('open', filename))
    elif platform.system() == 'Windows':    # Windows
        os.startfile(filename)
    else:                                   # linux variants
        subprocess.call(('xdg-open', filename))

def Unisolate(occs):
    for occ in occs:
        occ.isLightBulbOn = True

def HideAll(occs):
    for occ in occs:
        comp = occ.component
        comp.isBodiesFolderLightBulbOn = False

def ShowAll(occs):
    for occ in occs:
        comp = occ.component
        comp.isBodiesFolderLightBulbOn = True

def takeRootPhoto(component, path):
    cameraTarget = False
    occurrence = component.occurrences[0]

    cameraTarget = adsk.core.Point3D.create(occurrence.transform.translation.x, occurrence.transform.translation.y, occurrence.transform.translation.z)

    setGridDisplay(False)

    viewport = app.activeViewport
    camera = viewport.camera

    camera.target = cameraTarget
    camera.isFitView = True
    camera.isSmoothTransition = False
    camera.eye = adsk.core.Point3D.create(100 + cameraTarget.x, -100 + cameraTarget.y, 100 + cameraTarget.z)

    app.activeViewport.camera = camera
    
    app.activeViewport.refresh()
    adsk.doEvents()

    success = app.activeViewport.saveAsImageFile(path + '/root.png', 373, 709)
    if not success:
        ui.messageBox('Failed saving viewport image.')

def takePhoto(occ, base_path):
    currentComp = occ.component
    os.makedirs(base_path, exist_ok=True)
    filePath = os.path.join(base_path, "images")

    # Isolate the component
    occ.isIsolated = True

    setGridDisplay(False)

    viewport = app.activeViewport
    camera = viewport.camera
    
    cameraTarget = adsk.core.Point3D.create(occ.transform.translation.x, occ.transform.translation.y, occ.transform.translation.z)
        
    camera.target = cameraTarget
    camera.isFitView = True
    camera.isSmoothTransition = False
    camera.eye = adsk.core.Point3D.create(100 + cameraTarget.x, -100 + cameraTarget.y, 100 + cameraTarget.z)

    app.activeViewport.camera = camera
    
    app.activeViewport.refresh()
    adsk.doEvents()

    success = app.activeViewport.saveAsImageFile(filePath + '/' + name(currentComp.name)  + '.png', 300, 300)
    if not success:
        ui.messageBox('Failed saving viewport image.')

    occ.isIsolated = False

def write_component(component_base_path, component: adsk.fusion.Component):
    os.makedirs(component_base_path, exist_ok=True)
    write_step(component_base_path, component)

def write_step(output_path, component: adsk.fusion.Component):
    file_path = output_path + '/' + name(component.name) + ".stp"
    if os.path.exists(file_path):
      return
    export_manager = component.parentDesign.exportManager

    options = export_manager.createSTEPExportOptions(file_path, component)
    export_manager.execute(options)

def take(*path):
    out_path = os.path.join(*path)
    os.makedirs(out_path, exist_ok=True)
    return out_path

def name(name):
    name = re.sub('[^a-zA-Z0-9 \n\.]', '', name).strip()

    if name.endswith('.stp') or name.endswith('.stl') or name.endswith('.igs'):
      name = name[0: -4] + "_" + name[-3:]

    return name

def buildXLSX(bom, fileName, projectInfo):
    copyfile(addin_path + '/resources/logo.png', projectInfo['logoImage'])

    workbook = xlsxwriter.Workbook(fileName + '.xlsx')

    #define common formatting
    #title formatting
    title_format = workbook.add_format()
    title_format.set_bold()
    title_format.set_align('center')
    title_format.set_align('vcenter')

    #header formatting
    header_format = workbook.add_format()
    header_format.set_bg_color('#5599ff')
    header_format.set_bold()
    header_format.set_size(12)
    header_format.set_align('center')
    header_format.set_align('vcenter')

    #bom formatting
    bom_format = workbook.add_format()
    bom_format.set_align('center')
    bom_format.set_align('vcenter')
        
    #Build Project Summary Sheet
    project_worksheet = workbook.add_worksheet('Project')
    
    projectKey_format = workbook.add_format()
    projectKey_format.set_bold()
    projectKey_format.set_border(1)

    projectValue_format = workbook.add_format()
    projectValue_format.set_border(1)
    projectValue_format.set_align('center')
    projectValue_format.set_align('vcenter')

    notes_format = workbook.add_format()
    notes_format.set_border(1)

    #define edge column widths
    project_worksheet.set_column_pixels(0, 1, 20)
    project_worksheet.set_column_pixels(21, 21, 20)

    #project info
    
    project_worksheet.merge_range("C3:D3",'Project Name', projectKey_format)
    project_worksheet.merge_range("E3:J3", projectInfo['projectName'], projectValue_format)
    
    project_worksheet.merge_range("C4:D4",'Product Name', projectKey_format)
    project_worksheet.merge_range("E4:J4", projectInfo['productName'], projectValue_format)
    
    project_worksheet.merge_range("C5:D5",'Owner', projectKey_format)    
    project_worksheet.merge_range("E5:J5", projectInfo['owner'], projectValue_format)

    project_worksheet.merge_range("C6:D6",'Designer', projectKey_format)
    project_worksheet.merge_range("E6:J6", projectInfo['designer'], projectValue_format)

    project_worksheet.merge_range("C7:D7",'Completion Date', projectKey_format)    
    project_worksheet.merge_range("E7:J7", date.today().strftime("%B %d, %Y"), projectValue_format)

    #quote
    project_worksheet.merge_range("C9:D9", 'Quote', projectKey_format)    
    project_worksheet.merge_range("E9:J9",'', projectValue_format)

    #notes
    project_worksheet.merge_range("C11:J11",'Notes', projectKey_format)
    project_worksheet.merge_range("C12:J22", '', notes_format)

    #progress tracker
    project_worksheet.merge_range("C24:N24",'Progress Tracker', projectKey_format)
    
    project_worksheet.merge_range("C25:F25",'Design', projectKey_format)    
    project_worksheet.merge_range("G25:N25",'Shop', projectKey_format)

    project_worksheet.merge_range("C26:D26",'Model', projectValue_format)
    project_worksheet.merge_range("E26:F26",'BOM', projectValue_format)    
    project_worksheet.merge_range("G26:H26",'Quote', projectValue_format)    
    project_worksheet.merge_range("I26:J26",'Sample', projectValue_format)    
    project_worksheet.merge_range("K26:L26",'Adjustments', projectValue_format)    
    project_worksheet.merge_range("M26:N26",'Production', projectValue_format)
    
    # Conditional Formatting for Percentage Completion

    # Add a format. Light red fill with dark red text.
    lt50_format = workbook.add_format({
        "bg_color": "#FFC7CE",
        "font_color": "#9C0006",
        "num_format": 9,
        "border": 1
    })

    # Add a format. Green fill with dark green text.
    gt50_format = workbook.add_format({
        "bg_color": "#FFEB9C", 
        "font_color": "#9C5700",
        "num_format": 9,
        "border": 1
    })

    # Add a format. Green fill with dark green text.
    at100_format = workbook.add_format({
        "bg_color": "#C6EFCE", 
        "font_color": "#006100",
        "num_format": 9,
        "border": 1
    })
    
    # Less than 50%
    project_worksheet.conditional_format(
        "C27:N27", 
        {
            "type": 
            "cell", 
            "criteria": "<", 
            "value": 0.5, 
            "format": lt50_format
        }
    )

    # Between 50% and 99%
    project_worksheet.conditional_format(
        "C27:N27", 
        {
            "type": "cell", 
            "criteria": "between", 
            "minimum": 0.5, 
            "maximum": 0.99, 
            "format": gt50_format
        }
    )

    # 100%
    project_worksheet.conditional_format(
        "C27:N27", 
        {
            "type": "cell", 
            "criteria": "=", 
            "value": 1, 
            "format": at100_format
        }
    )

    project_worksheet.merge_range("C27:D27", 1)
    project_worksheet.merge_range("E27:F27", 0.9)    
    project_worksheet.merge_range("G27:H27", 0)    
    project_worksheet.merge_range("I27:J27", 0)    
    project_worksheet.merge_range("K27:L27", 0)    
    project_worksheet.merge_range("M27:N27", 0)

    #logo    
    project_worksheet.merge_range("C29:N38",'', projectKey_format)
    project_worksheet.insert_image("C29", projectInfo['logoImage'], {'x_offset': 15, 'y_offset': 15, 'x_scale': 3, 'y_scale': 3})

    #assembly image
    project_worksheet.merge_range("P3:U38",'', projectKey_format)
    project_worksheet.insert_image("P3", projectInfo['rootImage'], {'x_offset': 5, 'y_offset': 5})

    #Build BOM Sheet
    headerColumns = [
        'Part Number',
        'Thumbnail',
        'Part Name',
        'Quantity',
        'Material',
        'Finish',
        'Notes'
    ]

    bom_worksheet = workbook.add_worksheet('BOM')
    bom_worksheet.merge_range(0,0,0,6,'Primary Bill of Materials', title_format)
    bom_worksheet.set_row_pixels(0, 40)
    hcol = 0
    for header in headerColumns:
        bom_worksheet.set_column_pixels(hcol, hcol, 100)
        bom_worksheet.write(1, hcol, header, header_format)
        hcol += 1
    bom_worksheet.set_row_pixels(1, 40)
    row = 2
    for item in bom:
        #write part number as row number minus one (header)
        bom_worksheet.write(row, 0, row-1, header_format)
        col = 1
        for prop in item:
            if prop is 'thumbnail':
                bom_worksheet.set_column_pixels(col, col, 156)
                bom_worksheet.set_row_pixels(row, 156)
                bom_worksheet.insert_image(row, col, item[prop], {"x_offset": 3, "y_offset": 3, 'x_scale': 0.5, 'y_scale': 0.5})
                col += 1
            elif isinstance(item[prop], str) or isinstance(item[prop], int):
                bom_worksheet.write(row, col, item[prop], bom_format)
                col += 1
        row += 1
    bom_worksheet.autofit()

    workbook.close()

def isGridDisplayOn():
    app = adsk.core.Application.get()
    ui  = app.userInterface

    cmdDef = ui.commandDefinitions.itemById('ViewLayoutGridCommand')
    listCntrlDef = adsk.core.ListControlDefinition.cast(cmdDef.controlDefinition)
    layoutGridItem = listCntrlDef.listItems.item(0)
    
    if layoutGridItem.isSelected:
        return True
    else:
        return False

def setGridDisplay(turnOn):
    app = adsk.core.Application.get()
    ui  = app.userInterface

    cmdDef = ui.commandDefinitions.itemById('ViewLayoutGridCommand')
    listCntrlDef = adsk.core.ListControlDefinition.cast(cmdDef.controlDefinition)
    layoutGridItem = listCntrlDef.listItems.item(0)
    
    if turnOn:
        layoutGridItem.isSelected = True
    else:
        layoutGridItem.isSelected = False       