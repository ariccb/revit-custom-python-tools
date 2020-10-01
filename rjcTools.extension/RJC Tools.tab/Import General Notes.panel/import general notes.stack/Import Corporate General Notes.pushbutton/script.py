# Set sys.path
import sys
import os
import fnmatch

import pyrevit
from pyrevit.forms import alert
sys.path.insert(0,'C:\Program Files (x86)\IronPython 2.7\Lib')

import traceback


from pyrevit.framework import List
from pyrevit import revit, DB
from pyrevit import script
from pyrevit import forms
from pyrevit.revit import query
from Autodesk.Revit.DB import Element as DBElement

__doc__= '''Imports General Notes from the Corporate General Notes Revit File
            saved in the Resource Folder. It opens the Corporate file in the
            background as Detached, and loads them into the current Revit file.'''

logger = script.get_logger()
output = script.get_output()
options = DB.OpenOptions()
options.DetachFromCentralOption = options.DetachFromCentralOption.DetachAndPreserveWorksets


app = __revit__.Application
active_doc = __revit__.ActiveUIDocument.Document

VIEW_TOS_PARAM = DB.BuiltInParameter.VIEW_DESCRIPTION

class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes


class Option(forms.TemplateListItem):
    def __init__(self, op_name, default_state=False):
        super(Option, self).__init__(op_name)
        self.state = default_state

class OptionSet:
    def __init__(self):
        self.op_copy_vports = Option('Copy Viewports', True)
        self.op_copy_schedules = Option('Copy Schedules', True)
        self.op_copy_titleblock = Option('Copy Sheet Titleblock', True)
        self.op_copy_revisions = Option('Copy and Set Sheet Revisions', False)
        self.op_copy_placeholders_as_sheets = \
            Option('Copy Placeholders as Sheets', True)
        self.op_copy_guides = Option('Copy Guide Grids', True)
        self.op_update_exist_view_contents = \
            Option('Update Existing View Contents')


class MetricOrImperialOptionSet:
    def __init__(self):
        self.op_metric_gn = Option('Metric Corporate General Notes', True)
        self.op_imperial_gn = Option('Imperial Corporate General Notes', False)


class SheetOrViewOptionSet:
    def __init__(self):
        self.op_load_full_sheets = Option('Import Full Sheets', False)
        self.op_load_full_sheets = Option('Import Individual Views', True)


class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    """Handle copy and paste errors."""

    def OnDuplicateTypeNamesFound(self, args):
        """Use destination model types if duplicate."""
        return DB.DuplicateTypeAction.UseDestinationTypes

corp_gn_path = 'R:\Technical Resources\STR\General Notes and Details\Revit'
# cal_gn_path = 'R:\Office Services\CAL\STR\Cal Project Resources\_Production Resources\3 DD\RJC Calgary General Notes\RJC Calgary General Notes - Metric.rvt'


# Ask if you want imperial or metric notes, load correct path
def load_metric_or_imperial_gn_in_background():
    selected_units = ''
    op_set = MetricOrImperialOptionSet()
    metric_or_imperial = forms.SelectFromList.show(
        [getattr(op_set, x) for x in dir(op_set) if x.startswith('op_')], 
        title='Select Corporate General Note To Import',
        multiselect=False,
        height=225
        )
    # closure function for checking the file paths work (the scope is only 
    # inside metric_or_impeial_options() function)
    def check_GN_file_paths(file_path, units):
        # check filename with wildcards to allow changing of date and revit version in template filename
        for file in os.listdir(file_path):
            if fnmatch.fnmatch(file,'STR-STD-00?-*'+units+' Notes - Revit 20??.rvt'):
                file_path = file_path + '\\' + file
                logger.debug(file_path) # shows the file path of Corp. GN file after passing the fnmatch command
                if not os.path.isfile(file_path):                    
                    alert('The path to the {} Revit file needs to be updated, please select the correct file'.format(metric_or_imperial))
                    new_file_path = pyrevit.forms.pick_file(file_ext='rvt', \
                                                            init_dir=corp_gn_path, \
                                                            multi_file=False)

                    print('the new selected file to be loaded in background is: ' + new_file_path)
                    return new_file_path
                else:
                    output.print_md('**Loading... ** {}'.format(metric_or_imperial))
                    output.print_md('**From: ** {}'.format(file_path))
                    return file_path
        
    if not metric_or_imperial:
        sys.exit(0)

    # load metric file from corporate file path if Metric is chosen in the metric_or_imperial form
    if metric_or_imperial == 'Metric Corporate General Notes': 
        selected_units = 'Metric'
        
        

    # load imperial revit file from corporate file path if Imperial is chosen in the metric_or_imperial form
    if metric_or_imperial == 'Imperial Corporate General Notes': 
        selected_units = 'Imperial'
    
    logger.debug('Attempting to load the {0} Revit file'.format(metric_or_imperial))
    final_gn_file_path = check_GN_file_paths(corp_gn_path,selected_units)
    # converts into the file path that Revit can read
    modelPath = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(final_gn_file_path)
    backgroundDoc = app.OpenDocumentFile(modelPath, options)     

    print('The ' + metric_or_imperial + ' Revit file has been loaded in the background')
    logger.debug('Background document object: {}'.format(backgroundDoc))

    return backgroundDoc, metric_or_imperial


def get_source_views():
    selected_source_views = forms.select_views(title='Select Views To Import From {}'.format(gn_title),
                                                doc=source_doc,
                                                width=700)
                                                #This is loading "ViewSheet Propterties....don't know why.
                                                #Check out forms.select_views return values?
                                                #I think i just need to remove the Titleblock optionset out!!
    if not selected_source_views:
        sys.exit(0)
    else:
        return selected_source_views


def get_user_options():
    op_set = OptionSet()
    return_options = \
        forms.SelectFromList.show(
            [getattr(op_set, x) for x in dir(op_set) if x.startswith('op_')],
            title='Select Copy Options',
            button_name='Copy Now',
            multiselect=True
            )

    if not return_options:
        sys.exit(0)

    return op_set

### DON'T NEED THIS BECAUSE OUR DESTINATION IS THE CURRENT DOC.
# def get_dest_docs():
#     # find open documents other than the active doc
#     selected_dest_docs = \
#         forms.select_open_docs(title='Select Destination Documents')
#     if not selected_dest_docs:
#         sys.exit(0)
#     else:
#         return selected_dest_docs


### DON'T NEED THIS BECAUSE WE'RE LOADING THE VIEWS FROM A DIFFERENT DOCUMENT (FUNCTION DEF BELOW)
# def get_source_views():
#     view_elements = forms.select_sheets(button_name='Copy Sheets',
#                                          use_selection=True)

#     if not view_elements:
#         sys.exit(0)

#     return view_elements


def get_default_type(source_doc, type_group):
    return source_doc.GetDefaultElementTypeId(type_group)


def find_matching_view(dest_doc, source_view):
    for v in DB.FilteredElementCollector(dest_doc).OfClass(DB.View):
        if v.ViewType == source_view.ViewType \
                and query.get_name(v) == query.get_name(source_view):
            if source_view.ViewType == DB.ViewType.DrawingSheet:
                if v.SheetNumber == source_view.SheetNumber:
                    return v
            else:
                return v


def find_guide(guide_name, source_doc):
    # collect guides in dest_doc
    guide_elements = \
        DB.FilteredElementCollector(source_doc)\
            .OfCategory(DB.BuiltInCategory.OST_GuideGrid)\
            .WhereElementIsNotElementType()\
            .ToElements()
    
    # find guide with same name
    for guide in guide_elements:
        if str(guide.Name).lower() == guide_name.lower():
            return guide


def get_view_contents(dest_doc, source_view):
    view_elements = DB.FilteredElementCollector(dest_doc, source_view.Id)\
                      .WhereElementIsNotElementType()\
                      .ToElements()

    elements_ids = []
    for element in view_elements:
        if (element.Category and element.Category.Name == 'Title Blocks') \
                and not OPTION_SET.op_copy_titleblock:
            continue
        elif isinstance(element, DB.ScheduleSheetInstance) \
                and not OPTION_SET.op_copy_schedules:
            continue
        elif isinstance(element, DB.Viewport) \
                or 'ExtentElem' in query.get_name(element):
            continue
        elif isinstance(element, DB.Element) \
                and element.Category \
                and 'guide' in str(element.Category.Name).lower():
            continue
        elif isinstance(element, DB.Element) \
                and element.Category \
                and 'views' == str(element.Category.Name).lower():
            continue
        else:
            elements_ids.append(element.Id)
    return elements_ids


def ensure_dest_revision(src_rev, all_dest_revs, dest_doc):
    # check to see if revision exists
    for rev in all_dest_revs:
        if query.compare_revisions(rev, src_rev):
            return rev

    # if no matching revisions found, create a new revision and return
    logger.warning('Revision could not be found in destination model.\n'
                   'Revision Date: {}\n'
                   'Revision Description: {}\n'
                   'Creating a new revision. Please review revisions '
                   'after copying process is finished.'
                   .format(src_rev.RevisionDate, src_rev.Description))
    return revit.create.create_revision(description=src_rev.Description,
                                        by=src_rev.IssuedBy,
                                        to=src_rev.IssuedTo,
                                        date=src_rev.RevisionDate,
                                        doc=dest_doc)


def clear_view_contents(dest_doc, dest_view):
    logger.debug('Removing view contents: {}'.format(dest_view.Name))
    elements_ids = get_view_contents(dest_doc, dest_view)

    with revit.Transaction('Delete View Contents', doc=dest_doc):
        for el_id in elements_ids:
            try:
                dest_doc.Delete(el_id)
            except Exception as err:
                continue

    return True


def copy_view_contents(sourceDoc, source_view, dest_doc, dest_view,
                       clear_contents=False):
    logger.debug('Copying view contents: {} : {}'
                 .format(source_view.Name, source_view.ViewType))

    elements_ids = get_view_contents(sourceDoc, source_view)

    if clear_contents:
        if not clear_view_contents(dest_doc, dest_view):
            return False

    cp_options = DB.CopyPasteOptions()
    cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())

    if elements_ids:
        with revit.Transaction('Copy View Contents',
                               doc=dest_doc,
                               swallow_errors=True):
            DB.ElementTransformUtils.CopyElements(
                source_view,
                List[DB.ElementId](elements_ids),
                dest_view, None, cp_options
                )

    return True


def copy_view_props(source_view, dest_view):
    dest_view.Scale = source_view.Scale
    dest_view.Parameter[VIEW_TOS_PARAM].Set(
        source_view.Parameter[VIEW_TOS_PARAM].AsString()
    )


def copy_view(sourceDoc, source_view, dest_doc):
    matching_view = find_matching_view(dest_doc, source_view)
    if matching_view:
        print('\t\t\tView/Sheet already exists in {}'.format(source_view.Name))
        if OPTION_SET.op_update_exist_view_contents:
            if not copy_view_contents(sourceDoc,
                                      source_view,
                                      dest_doc,
                                      matching_view,
                                      clear_contents=True):
                logger.error('Could not copy view contents: {}'
                             .format(source_view.Name))
            else:
                print('\t\t\t\t\t\tExisting view contents replaced with the latest content from {}'.format(gn_title))

        return matching_view

    logger.debug('Copying view: {}'.format(source_view.Name))
    new_view = None

    if source_view.ViewType == DB.ViewType.DrawingSheet:
        try:
            logger.debug('Source view is a sheet. '
                         'Creating destination sheet.')

            with revit.Transaction('Create Sheet', doc=dest_doc):
                if not source_view.IsPlaceholder \
                        or (source_view.IsPlaceholder
                                and OPTION_SET.op_copy_placeholders_as_sheets):
                    new_view = \
                        DB.ViewSheet.Create(
                            dest_doc,
                            DB.ElementId.InvalidElementId
                            )
                else:
                    new_view = DB.ViewSheet.CreatePlaceholder(dest_doc)

                revit.update.set_name(new_view,
                                      revit.query.get_name(source_view))
                new_view.SheetNumber = source_view.SheetNumber
        except Exception as sheet_err:
            logger.error('Error creating sheet. | {}'.format(sheet_err))
    elif source_view.ViewType == DB.ViewType.DraftingView:
        try:
            logger.debug('Source view is a drafting. '
                         'Creating destination drafting view.')

            with revit.Transaction('Create Drafting View', doc=dest_doc):
                new_view = DB.ViewDrafting.Create(
                    dest_doc,
                    get_default_type(dest_doc,
                                     DB.ElementTypeGroup.ViewTypeDrafting)
                )
                revit.update.set_name(new_view,
                                      revit.query.get_name(source_view))
                copy_view_props(source_view, new_view)
        except Exception as sheet_err:
            logger.error('Error creating drafting view. | {}'
                         .format(sheet_err))
    elif source_view.ViewType == DB.ViewType.Legend:
        try:
            logger.debug('Source view is a legend. '
                         'Creating destination legend view.')

            first_legend = query.find_first_legend(dest_doc)
            if first_legend:
                with revit.Transaction('Create Legend View', doc=dest_doc):
                    new_view = \
                        dest_doc.GetElement(
                            first_legend.Duplicate(
                                DB.ViewDuplicateOption.Duplicate
                                )
                            )
                    revit.update.set_name(new_view,
                                        revit.query.get_name(source_view))
                    copy_view_props(source_view, new_view)
            else:
                logger.error('Destination document must have at least one '
                             'Legend view. Skipping legend.')
        except Exception as sheet_err:
            logger.error('Error creating drafting view. | {}'
                         .format(sheet_err))

    if new_view:
        copy_view_contents(sourceDoc, source_view, dest_doc, new_view)

    return new_view


def copy_viewport_types(sourceDoc, vport_type, vport_typename,
                        dest_doc, newvport):
    dest_vport_typenames = [DBElement.Name.GetValue(dest_doc.GetElement(x))
                            for x in newvport.GetValidTypes()]

    cp_options = DB.CopyPasteOptions()
    cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())

    if vport_typename not in dest_vport_typenames:
        with revit.Transaction('Copy Viewport Types',
                               doc=dest_doc,
                               swallow_errors=True):
            DB.ElementTransformUtils.CopyElements(
                sourceDoc,
                List[DB.ElementId]([vport_type.Id]),
                dest_doc,
                None,
                cp_options,
                )


def apply_viewport_type(sourceDoc, vport_id, dest_doc, newvport_id):
    with revit.Transaction('Apply Viewport Type', doc=dest_doc):
        vport = sourceDoc.GetElement(vport_id)
        vport_type = sourceDoc.GetElement(vport.GetTypeId())
        vport_typename = DBElement.Name.GetValue(vport_type)

        newvport = dest_doc.GetElement(newvport_id)

        copy_viewport_types(sourceDoc, vport_type, vport_typename,
                            dest_doc, newvport)

        for vtype_id in newvport.GetValidTypes():
            vtype = dest_doc.GetElement(vtype_id)
            if DBElement.Name.GetValue(vtype) == vport_typename:
                newvport.ChangeTypeId(vtype_id)


def copy_sheet_viewports(sourceDoc, source_view, dest_doc, dest_sheet):
    existing_views = [dest_doc.GetElement(x).ViewId
                      for x in dest_sheet.GetAllViewports()]

    for vport_id in source_view.GetAllViewports():
        vport = sourceDoc.GetElement(vport_id)
        vport_view = sourceDoc.GetElement(vport.ViewId)

        print('\t\tCopying/updating view: {}'
              .format(revit.query.get_name(vport_view)))
        new_view = copy_view(sourceDoc, vport_view, dest_doc)

        if new_view:
            ref_info = revit.query.get_view_sheetrefinfo(new_view)
            if ref_info and ref_info.sheet_num != dest_sheet.SheetNumber:
                logger.error('View is already placed on sheet \"%s - %s\"',
                             ref_info.sheet_num, ref_info.sheet_name)
                continue

            if new_view.Id not in existing_views:
                print('\t\t\tPlacing copied view on sheet.')
                with revit.Transaction('Place View on Sheet', doc=dest_doc):
                    nvport = DB.Viewport.Create(dest_doc,
                                                dest_sheet.Id,
                                                new_view.Id,
                                                vport.GetBoxCenter())
                if nvport:
                    apply_viewport_type(sourceDoc, vport_id,
                                        dest_doc, nvport.Id)
            else:
                print('\t\t\tView already exists on the sheet.')


def copy_sheet_revisions(sourceDoc, source_view, dest_doc, dest_sheet):
    all_src_revs = query.get_revisions(doc=sourceDoc)
    all_dest_revs = query.get_revisions(doc=dest_doc)
    revisions_to_set = []

    with revit.Transaction('Copy and Set Revisions', doc=dest_doc):
        for src_revid in source_view.GetAdditionalRevisionIds():
            set_rev = ensure_dest_revision(sourceDoc.GetElement(src_revid),
                                           all_dest_revs,
                                           dest_doc)
            revisions_to_set.append(set_rev)

        if revisions_to_set:
            revit.update.update_sheet_revisions(revisions_to_set,
                                                [dest_sheet],
                                                state=True,
                                                doc=dest_doc)


def copy_sheet_guides(sourceDoc, source_view, dest_doc, dest_sheet):
    # sheet guide
    source_view_guide_param = \
        source_view.Parameter[DB.BuiltInParameter.SHEET_GUIDE_GRID]
    source_view_guide_element = \
        sourceDoc.GetElement(source_view_guide_param.AsElementId())
    
    if source_view_guide_element:
        if not find_guide(source_view_guide_element.Name, dest_doc):
            # copy guides to dest_doc
            cp_options = DB.CopyPasteOptions()
            cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())

            with revit.Transaction('Copy Sheet Guide', doc=dest_doc):
                DB.ElementTransformUtils.CopyElements(
                    sourceDoc,
                    List[DB.ElementId]([source_view_guide_element.Id]),
                    dest_doc, None, cp_options
                    )

        dest_guide = find_guide(source_view_guide_element.Name, dest_doc)
        if dest_guide:
            # set the guide
            with revit.Transaction('Set Sheet Guide', doc=dest_doc):
                dest_sheet_guide_param = \
                    dest_sheet.Parameter[DB.BuiltInParameter.SHEET_GUIDE_GRID]
                dest_sheet_guide_param.Set(dest_guide.Id)
        else:
            logger.error('Error copying and setting sheet guide for sheet {}'
                         .format(source_view.Name))


def copy_sheet(sourceDoc, source_view, dest_doc):
    logger.debug('Copying sheet {} to document {}'
                 .format(source_view.Name,
                         dest_doc.Title))
    print('\tCopying/updating Sheet: {}'.format(source_view.Name))
    with revit.TransactionGroup('Import Sheet', doc=dest_doc):
        logger.debug('Creating destination sheet...')
        new_sheet = copy_view(sourceDoc, source_view, dest_doc)

        if new_sheet:
            if not new_sheet.IsPlaceholder:
                if OPTION_SET.op_copy_vports:
                    logger.debug('Copying sheet viewports...')
                    copy_sheet_viewports(sourceDoc, source_view,
                                         dest_doc, new_sheet)
                else:
                    print('Skipping viewports...')

                if OPTION_SET.op_copy_guides:
                    logger.debug('Copying sheet guide grids...')
                    copy_sheet_guides(sourceDoc, source_view,
                                      dest_doc, new_sheet)
                else:
                    print('Skipping sheet guides...')

            if OPTION_SET.op_copy_revisions:
                logger.debug('Copying sheet revisions...')
                copy_sheet_revisions(sourceDoc, source_view,
                                     dest_doc, new_sheet)
            else:
                print('Skipping revisions...')

        else:
            logger.error('Failed copying sheet: {}'.format(source_view.Name))

#copied code ^






# start of running code, above is definitions of functions
source_doc, gn_title = load_metric_or_imperial_gn_in_background()
OPTION_SET = get_user_options()

dest_doc = active_doc
logger.debug('destination doc: {}'.format(dest_doc))
logger.debug('source doc (loaded in background): {}'.format(source_doc))

#create list to add views from draftviews_collector that have a matching parameter value for the following for loop
matchedViews = []
#similar list to spit out names for checking/clarity
matchedViews_Names = []

source_views = get_source_views()
# print(source_views) # uncomment to see what the source view objects look like
view_count = len(source_views)
output.print_md('Loading **{}** views from {}'.format(view_count, gn_title))

#other_try_background_views = get_background_session_views()

total_work = view_count #* doc_count --> would use this if we were copying to more than one document, but we're not
work_counter = 0

#get_dest_docs() needs to be replaced with "current document"
#get source stuff needs to be replaced with "backgroundDoc"

output.print_md('**Copying View(s) to Document:** {0}'
                .format(dest_doc.Title))

for source_view in source_views:
    print('Copying View: {0}'.format(source_view.Name))
    copy_view(source_doc, source_view, dest_doc)
    # copy_sheet(revit.doc, source_view, dest_doc) --> function from 'pyrevit copy sheets to open documents
    work_counter += 1
    output.update_progress(work_counter, total_work)

output.print_md('**Copied {} views to {}.**'
                .format(view_count, dest_doc.Title))
output.print_md('**GENERAL NOTE IMPORT COMPLETE**')