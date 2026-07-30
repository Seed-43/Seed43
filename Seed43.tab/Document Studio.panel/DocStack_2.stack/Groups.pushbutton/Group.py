# -*- coding: utf-8 -*-
"""
pyTable diagnostic - run this once via pyRevit (e.g. pyRevit > Python
Console, or as a temporary one-off script/button) to find the exact
ElementId of the "<Lines>" style in THIS project.

Prints every GraphicsStyle under the "Lines" category, plus a couple
of alternate name-reading methods per entry, so we can see exactly
what's really there instead of guessing.
"""
from pyrevit import revit, DB, script

doc = revit.doc
output = script.get_output()

def safe_name(el):
    try:
        p = el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            return p.AsString()
    except Exception:
        pass
    return None

def safe_name2(el):
    try:
        return DB.Element.Name.GetValue(el)
    except Exception:
        return None

def safe_name3(el):
    try:
        return el.Name
    except Exception:
        return None

print("=== Lines category subcategories ===")
lines_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
for sub in lines_cat.SubCategories:
    try:
        gs = sub.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
        gs_id = gs.Id if gs else None
    except Exception as ex:
        gs_id = 'ERROR: {}'.format(ex)
    print("Subcategory raw obj -> SYMBOL_NAME_PARAM={!r}  Element.Name.GetValue={!r}  .Name={!r}  GraphicsStyleId={}".format(
        safe_name(sub), safe_name2(sub), safe_name3(sub), gs_id
    ))

print("")
print("=== Lines category's own top-level default GraphicsStyle ===")
try:
    gs = lines_cat.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
    print("Id={}  SYMBOL_NAME_PARAM={!r}  Element.Name.GetValue={!r}  .Name={!r}".format(
        gs.Id if gs else None, safe_name(gs), safe_name2(gs), safe_name3(gs)
    ))
except Exception as ex:
    print("ERROR: {}".format(ex))

print("")
print("=== ALL GraphicsStyle elements in document whose category is 'Lines' ===")
for gs in DB.FilteredElementCollector(doc).OfClass(DB.GraphicsStyle):
    try:
        cat = gs.GraphicsStyleCategory
        cat_name = safe_name3(cat) if cat else None
    except Exception:
        cat_name = 'ERROR'
    if cat_name == 'Lines' or (cat_name and 'ine' in str(cat_name)):
        print("Id={}  category_name={!r}  SYMBOL_NAME_PARAM={!r}  Element.Name.GetValue={!r}  .Name={!r}".format(
            gs.Id, cat_name, safe_name(gs), safe_name2(gs), safe_name3(gs)
        ))

print("")
print("=== EVERY GraphicsStyle in the document (id + all 3 name reads) ===")
for gs in DB.FilteredElementCollector(doc).OfClass(DB.GraphicsStyle):
    n1, n2, n3 = safe_name(gs), safe_name2(gs), safe_name3(gs)
    if n1 == '<Lines>' or n2 == '<Lines>' or n3 == '<Lines>' or \
       n1 == 'Lines' or n2 == 'Lines' or n3 == 'Lines':
        print("*** MATCH *** Id={}  SYMBOL_NAME_PARAM={!r}  Element.Name.GetValue={!r}  .Name={!r}".format(
            gs.Id, n1, n2, n3
        ))
