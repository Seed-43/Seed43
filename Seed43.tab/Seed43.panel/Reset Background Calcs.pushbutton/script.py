# -*- coding: utf-8 -*-
#Reset Background Calcs

import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

uiapp = __revit__
year = str(uiapp.Application.VersionNumber)[:4]
rfa = os.path.join(SCRIPT_DIR, "CLOSE THIS FAMILY {}.rfa".format(year))

if os.path.isfile(rfa):
    uiapp.OpenAndActivateDocument(rfa)
