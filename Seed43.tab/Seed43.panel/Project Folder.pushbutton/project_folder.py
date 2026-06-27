import subprocess
import os

__doc__ = "Opens the project folder in Windows Explorer."

doc = __revit__.ActiveUIDocument.Document
path = doc.PathName

cloud_prefixes = ("BIM 360:", "ACC:", "Autodesk Docs:")

# For workshared models, use the central model path instead of the local path
if doc.IsWorkshared:
    from Autodesk.Revit.DB import ModelPathUtils
    central_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(
        doc.GetWorksharingCentralModelPath()
    )
    if central_path:
        path = central_path

if not path:
    from pyrevit import forms
    forms.alert("Project has not been saved yet.")
elif path.startswith(cloud_prefixes):
    from pyrevit import forms
    forms.alert("This is a cloud model. No local folder to open.")
else:
    folder = os.path.dirname(path)
    if os.path.exists(folder):
        subprocess.Popen(['explorer', folder])
    else:
        from pyrevit import forms
        forms.alert("Folder not found: {}".format(folder))