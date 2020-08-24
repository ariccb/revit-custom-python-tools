# Set sys.path
import sys
sys.path.insert(0,'C:\Program Files (x86)\IronPython 2.7\Lib')

from pyrevit.framework import List
from pyrevit import revit, DB
from pyrevit import script
from pyrevit import forms
from pyrevit.revit import query


app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document

class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes

corp_metric_gn_path = 'R:\Technical Resources\STR\General Notes and Details\Revit\STR-STD-001-20200812-RCK-Metric Notes - Revit 2018.rvt'
corp_imp_gn_path = 'R:\Technical Resources\STR\General Notes and Details\Revit\STR-STD-000-20200812-RCK-Imperial Notes - Revit 2018.rvt'
cal_gn_path = 'R:\Office Services\CAL\STR\Cal Project Resources\_Production Resources\3 DD\RJC Calgary General Notes\RJC Calgary General Notes - Metric.rvt'
backgroundDoc = ''

logger = script.get_logger()
options = DB.OpenOptions()
options.DetachFromCentralOption = options.DetachFromCentralOption.DetachAndPreserveWorksets

# Ask if you want imperial or metric notes, load correct path
metric_or_imperial = forms.ask_for_one_item(
    ['Metric General Notes','Imperial General Notes'], 
    title='Import Corporate General Notes', 
    prompt='Load Metric or Imperial Notes?',
    default='Metric General Notes',
    )
if metric_or_imperial == 'Metric General Notes':
    print('metric')
    backgroundDoc = app.OpenDocumentFile(corp_metric_gn_path)
    print(backgroundDoc)

if metric_or_imperial == 'Imperial General Notes': 
    print('imperial')
    backgroundDoc = app.OpenDocumentFile(corp_imp_gn_path)
    print(backgroundDoc)

# get all views and collect names of current doc
all_graphviews = revit.query.get_all_views(doc=revit.doc)
gn_views = revit.query.get_elements_by_parameter(param_name='RJC Standard View ID',
                              doc=backgroundDoc, partial=False)
print(gn_views)

all_draftview_names = [revit.query.get_name(x)
                       for x in all_graphviews
                       if x.ViewType == DB.ViewType.DraftingView]
                  #    if x.LookupParameter('RJC Standard View ID') != None
                  #    if len(x.LookupParameter('RJC Standard View ID').ToString()) >= 3]
selection = revit

drafting_views_to_copy = forms.select_views(
    title='General Notes to Import',
    doc=backgroundDoc,
    filterfunc=lambda x: x.ViewType == DB.ViewType.DraftingView)

if drafting_views_to_copy:
    # iterate over interfacetypes drafting views
    for src_drafting in drafting_views_to_copy:
        logger.debug('Copying %s', revit.query.get_name(src_drafting))

        # get drafting view elements and exclude non-copyable elements


draftview_elements_collector = DB.FilteredElementCollector(backgroundDoc).OfClass(ViewDrafting).ToElements()
#draftview_ids_collector = FilteredElementCollector(backgroundDoc).OfClass(ViewDrafting).ToElementIds()
draftviewName = []
for i in draftview_elements_collector: #creates a loop and iterates the following code through each instance in the list 'draftview_elements_collector' using draftviewName as the variable
        param = i.LookupParameter('View Name').AsString()
        draftviewName.append(param)

# print(draftview_elements_collector)
# print(draftview_ids_collector)


# list_of_notes = []
# for i in draftview_elements_collector:
#     print(i)

#     name = draftview_elements_collector[i].Name
#     list_of_notes.append(name)
    
#     # list_of_notes.append(generalNote(draftview_elements_collector[i].Name,draftview_ids_collector[i],draftview_elements_collector[i]))
# print(list_of_notes)




# for i in draftview_elements_collector:
#     i = generalNote()
#     i.draft


forms.SelectFromList.show(
    list_of_notes,
    title='Select General Notes To Import',
    button_name='Select Views',
    name_attr='viewName')     
                     

'''           #collect all drafting views in project
            draftview_collector = FilteredElementCollector(doc).OfClass(ViewDrafting).ToElements()

            #uncomment the next two lines if you want to check the element id's for the views in the filter
            #draftview_ids_collector = FilteredElementCollector(doc).OfClass(ViewDrafting).ToElementIds()
            #print draftview_ids_collector

            #create list to add views from draftview_collector that have a matching parameter value for the following for loop
            matchedViews = []
            #similar list to spit out names for checking/clarity
            matchedViews_Names = []

            input('What General Note View ID Should I try to load?')
            #lookup the parameter 'RJC Standard View ID', using .AsString to get the value instead of just the Parameter, and add views to list if it matches any values in 'list_generalnote_ids'
            try: 
                for draftview in draftview_collector: #creates a loop and iterates the following code through each instance in the list 'draftview_collector' using draftview as the variable
                    param = draftview.LookupParameter('RJC Standard View ID').AsString()
                    if param in list_generalnote_ids:
                        matchedViews.append(draftview)         # saves the matching views to matchedViews array (the original element types)
                        matchedViews_Names.append(draftview.Name)    # saves the view name to the matchedViews_Names array
                print "Views in project matching the selection form are: ", matchedViews_Names   # prints the seelctedViews_Names array with added view names from ^ for loop (this array was empty prior to the for loop)

'''

    