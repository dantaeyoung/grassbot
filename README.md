grassbot
========

 -- GRASSBOT -- Rhino/Grasshopper EXECUTER --
------------------------------------------------------------------------
     Created by Dan Taeyoung Lee, November 2013, dan@taeyounglee.com
------------------------------------------------------------------------

Usage: GRASSBOT.exe /rhfile rhino-file /ghfile grasshopper-file [/bakename] [/tempdir] [/bake] [/application]

/rhfile rhino-file           Full Path to Rhino file to be opened.
/ghfile grasshopper-file     Full Path to Grasshopper file to be run.
/screenshotlocation path     Full path to screenshot save location.
/bakename name               Name of the Grasshopper node to bake
                                 - defaults to 'BAKE'
/tempdir path                Path to temporary directory (no trailing slash)
/bake (true/false)           If Grasshopper should bake geometry
/param name                  TBD - Changes CATIA 'name' param based on 'name' gh node.
/application (true/false)    If Grassbot should load Rhino as Application or Interface.
                                 - Application reopens Rhino; Interface reuses same Rhino instance.

Example: GRASSBOT.exe /rhfile D:/GRASSBOT/test.3dm /ghfile D:\WILDCAT\MakeCircle
s.gh /bakename BAKENODE /tempdir C:\TEMP

