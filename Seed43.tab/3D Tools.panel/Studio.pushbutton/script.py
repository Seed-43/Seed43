# -*- coding: utf-8 -*-
# Studio.pushbutton
#
# Opens Seed43 Studio in the default browser. There is no Revit work here -
# Studio is a browser tool that reads StruCad .bsw models and converts them
# to IFC (among other formats), which is how a .bsw model gets into Revit.

__title__  = "Studio"
__author__ = "SEED43"

import webbrowser

from pyrevit import script

STUDIO_URL = "https://seed43.org/studio/"


def main():
    try:
        webbrowser.open(STUDIO_URL)
    except Exception as ex:
        # No default browser, or the shell refused the call. Show the address
        # rather than failing silently, so it can still be opened by hand.
        script.get_output().print_md(
            "Could not open the browser: `{}`\n\n{}".format(ex, STUDIO_URL))


if __name__ == "__main__":
    main()
