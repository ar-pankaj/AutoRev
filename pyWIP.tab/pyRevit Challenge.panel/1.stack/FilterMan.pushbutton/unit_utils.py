# -*- coding: utf-8 -*-
"""
Unit conversion helpers that work across Revit 2022-2025+.
Always use these functions instead of manual division by 304.8.
"""
from __future__ import unicode_literals


def mm_to_feet(mm_value):
    """Convert millimeters to Revit internal units (decimal feet)."""
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
        return UnitUtils.ConvertToInternalUnits(
            float(mm_value), UnitTypeId.Millimeters
        )
    except (ImportError, AttributeError):
        from Autodesk.Revit.DB import UnitUtils, DisplayUnitType
        return UnitUtils.ConvertToInternalUnits(
            float(mm_value), DisplayUnitType.DUT_MILLIMETERS
        )


def feet_to_mm(feet_value):
    """Convert Revit internal units (decimal feet) to millimeters."""
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
        return UnitUtils.ConvertFromInternalUnits(
            float(feet_value), UnitTypeId.Millimeters
        )
    except (ImportError, AttributeError):
        from Autodesk.Revit.DB import UnitUtils, DisplayUnitType
        return UnitUtils.ConvertFromInternalUnits(
            float(feet_value), DisplayUnitType.DUT_MILLIMETERS
        )


def m_to_feet(m_value):
    """Convert meters to Revit internal units (decimal feet)."""
    return mm_to_feet(float(m_value) * 1000.0)
