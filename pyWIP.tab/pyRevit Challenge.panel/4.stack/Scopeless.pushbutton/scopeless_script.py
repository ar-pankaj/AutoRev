# -*- coding: utf-8 -*-
__title__ = "Scopeless"
__doc__ = """Version = 1.1
Date    = 26.06.2026
_____________________________________________________________________
Description:
Quita la caja de referencia y desactiva la caja de sección en la vista activa.
_____________________________________________________________________
Author: César G. Ferrer (cgofer@globalomnium.com)"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ 
#==================================================
from Autodesk.Revit.DB import BuiltInParameter, ElementId, View3D
from pyrevit import revit,forms

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
doc = revit.doc
view = doc.ActiveView

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ 
#==================================================
with revit.Transaction("Scopless"):
    
    # 1️⃣ Remove the Scope Box
    scope_box = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
    if scope_box and not scope_box.IsReadOnly:
        scope_box.Set(ElementId.InvalidElementId)
    else:
        forms.alert('No se pudo modificar la caja de referencia, puede que esté bloqueada por la plantilla de vista', exitscript=True)

    # 2️⃣ Uncheck the Section Box
    if isinstance(view, View3D):
        if view.IsSectionBoxActive:
            view.IsSectionBoxActive = False
