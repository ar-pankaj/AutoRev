# -*- coding: utf-8 -*-
__title__ = "List Hidden Elements \n Per View"
__doc__ = """Version = 1.0
Date    = 15.07.2024
_____________________________________________________________________
Description:
Modify Toposolid by selection Model Lines which are mode 
at specific location
_____________________________________________________________________
How-to:
-> Click on the button
-> ...
_____________________________________________________________________
Last update:
- [16.07.2024] - 1.1 Fixed an issue...
- [15.07.2024] - 1.0 RELEASE
_____________________________________________________________________
To-Do:
- Describe Next Features
_____________________________________________________________________
Author: Pankaj Prabhakar"""
__highlight__ = 'new'

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
# Regular + Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.UI import *

# pyRevit
from pyrevit import revit, forms, script

# .NET Imports (You often need List import)
import clr
clr.AddReference("System")
from System.Collections.Generic import List

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application
selection = uidoc.Selection                     #type: Selection

# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝
#==================================================

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
#==================================================

# Function to get all hidden elements per view and category
def get_hidden_elements_per_view(doc):
    # Dictionary to hold hidden elements by view and category
    hidden_elements = {}

    # Get all views in the project
    views = FilteredElementCollector(doc).OfClass(View).ToElements()

    # Iterate through each view
    for view in views:
        # Initialize a list for hidden elements in this view
        hidden_elements[view.Name] = {}

        # Get all elements in the document
        elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

        # Check each element's visibility in the current view
        for elem in elements:
            # Check if the element is hidden in the current view
            if elem.IsHidden(view):
                # Get the category of the element
                category = elem.Category.Name if elem.Category is not None else "Uncategorized"

                # Add element to the hidden elements dictionary
                if category not in hidden_elements[view.Name]:
                    hidden_elements[view.Name][category] = List[ElementId]()
                
                hidden_elements[view.Name][category].Add(elem.Id)

    return hidden_elements


# Get the current document
doc = __revit__.ActiveUIDocument.Document

# Retrieve hidden elements
hidden_elements = get_hidden_elements_per_view(doc)

# Prepare a message to display the results
result_message = ""

# Loop through the collected data to construct the result message
for view_name, categories in hidden_elements.items():
    result_message += "View: {}\n".format(view_name)
    for category, ids in categories.items():
        result_message += "  Category: {}, Hidden Elements: {}\n".format(category, len(ids))

# Display the results in a TaskDialog
td = TaskDialog("Hidden Elements Summary")
td.TitleAutoPrefix = False
td.MainInstruction = result_message
td.MainIcon = TaskDialogIcon.TaskDialogIconShield
td.Show()