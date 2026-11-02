"""
excel_interface — Excel Data Interface for Urban Renewal AI
===========================================================
Public entry point: ExcelDataInterface (interface.py)

Usage:
    from excel_interface import ExcelDataInterface

    iface = ExcelDataInterface()
    iface.load("urban_renewal_template.xlsx")
    iface.write({"SITE_001": "Tel Aviv", "PARCEL_002": "6660"}, source_name="XPlan")
    iface.save("output.xlsx")
    print(iface.report())
"""

from .interface import ExcelDataInterface
from .models import FieldLocation, WriteOperation, DataStatus

__all__ = ["ExcelDataInterface", "FieldLocation", "WriteOperation", "DataStatus"]
