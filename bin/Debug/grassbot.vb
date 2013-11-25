Imports System
Imports System.Threading
Imports Grasshopper.Kernel
Imports Grasshopper.Kernel.Types
Imports Rhino
'Imports CATIA
'Imports INFITF
'Imports MECMOD
Imports Rhino.Geometry

Imports Rhino.Commands



Module GRASSBOT

    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Declare Function GetCommandLineA Lib "Kernel32" () As String


    Dim BakeNodeName As String
    'Dim iPartDoc As Document
    Dim RhinoFile As String
    Dim GrasshopperFile As String
    Dim GH
    Dim CatTempExportFile As String
    Dim CatTempImportFile As String
    Dim TempDir As String
    Dim CommandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
    Dim BakeBool As Boolean
    Dim ReimportBool As Boolean
    Dim ReimportNewOnly As Boolean
    Dim RunMacro As Boolean
    Dim BakeActivateNodeName As String
    Dim OutputFileNodeName As String
    Dim OutputFile As String
    Dim ParamName As String
    Dim RunID As String
    Dim RhinoLaunchStringSuffix As String
    Dim ScreenshotLocation As String

    Dim objRhinoApp As Object
    Dim objRhinoScript As Object

    'Dim CATIA As Object
    'Dim catDoc As MECMOD.PartDocument
    'Dim catPart As MECMOD.Part
    'Dim catParams As Object




    Dim WildcatVersion As String = "v0.5.1_131112"


    <STAThread()>
    Sub Main()


        PrintGnuLicense()
        PrintWildcatLogo()

        SetDefaultParameters()
        If (ProcessCommandLineArgs() = True) Then
            If (GrasshopperFile <> "NOTSET") And (RhinoFile <> "NOTSET") Then
                DisplayParameters()
                RunWildcat()
            Else
                PrintUsage()
            End If
        End If


    End Sub



    Function RunWildcat()


        Console.WriteLine("----------- Opening Rhino as " & RhinoLaunchStringSuffix & " .... -----------")



        On Error Resume Next
        objRhinoApp = CreateObject("Rhino5x64." & RhinoLaunchStringSuffix)
        If (Err.Number <> 0) Then

            Console.WriteLine("Failed to open Rhino5, trying to open Rhino4")

            On Error Resume Next
            objRhinoApp = CreateObject("Rhino4. " & RhinoLaunchStringSuffix)
            If (Err.Number <> 0) Then
                Console.WriteLine("Failed to open Rhino4")
                Exit Function
            End If

        End If



        objRhinoApp.Visible = True
        objRhinoScript = objRhinoApp.GetScriptObject()

        ' Make attempts to get RhinoScript, sleep between each attempt.
        Dim nCount As Integer
        nCount = 0
        Do While (nCount < 10)
            On Error Resume Next
            objRhinoScript = objRhinoApp.GetScriptObject()
            If Err.Number <> 0 Then
                Err.Clear()
                Sleep(500)
                nCount = nCount + 1
            Else
                Exit Do
            End If
        Loop

        ' Display an error if needed.
        If (objRhinoScript Is Nothing) Then
            MsgBox("Failed to get RhinoScript")
        End If

        Console.WriteLine("----------- Opening Rhino document ... -----------")
        'suppress the "save?" dialog
        objRhinoScript.DocumentModified(False)
        Call objRhinoScript.Command("_-Open " & RhinoFile, 0)

        Console.WriteLine("----------- Opening Grasshopper ... -----------")
        Call objRhinoScript().Command("Grasshopper", 0)
        GH = objRhinoScript.GetPluginObject("Grasshopper")

        Console.WriteLine("----------- Opening Grasshopper definition " & GrasshopperFile & " ... -----------")
        GH.OpenDocument(GrasshopperFile)

        Console.WriteLine("----------- Assigning Grasshopper data ... -----------")

        GH.AssignDataToParameter(BakeActivateNodeName, "True")
        'GH.AssignDataToParameter(BakeActivateNodeName, "0")

        GH.AssignDataToParameter("RUNID", RunID)

        Console.WriteLine("----------- Executing Grasshopper definition ... -----------")

        'GH.ShowEditor()

        GH.RunSolver(True)

        Console.WriteLine("----------- Saving screenshot ... -----------")

        Call objRhinoScript.Command("-ViewCapturetoFile " & ScreenshotLocation & " Width=2000 Height=2000 DrawGrid=No DrawWorldAxes=No DrawCPlaneAxes=No _Enter")


        If (BakeBool = True) Then
            Console.WriteLine("----------- Baking geometry to Rhino ... -----------")
            'GH.AssignDataToParameter("OUTPUTFILE", OutputFile)
            GH.BakeDataInObject(BakeNodeName)
        End If

        If (ReimportBool = True) Then

            Console.WriteLine("----------- Exporting Rhino Document ... -----------")

            If (ReimportNewOnly = True) Then
                Console.WriteLine("----------- Exporting New Geometry from 'BAKELAYER' only ... -----------")
                Call objRhinoScript.Command("-_SelLayer BAKELAYER")
                Call objRhinoScript.Command("_Hide _enter _HideSwap _SelAll")
            Else
                Call objRhinoScript.Command("_SelAll")
            End If


            Call objRhinoScript.Command("_-Export " & CatTempImportFile & " _enter")
            'Call objRhinoScript.Command("_-SaveAs " & CatTempImportFile & " _enter")


            ''RunMacro = True
            ''temporary force
            'If RunMacro = True Then

            '    RunCatiaMacro()
            'End If

        End If

        If (String.Compare(RhinoLaunchStringSuffix, "Application") = 0) Then
            ' only exit when we're in application mode.

            'Console.WriteLine("----------- Closing Grasshopper document ... -----------")
            '' Close all grasshopper documents so we don't get the multi-save menu.
            'GH.CloseDocuments()


            'Console.WriteLine("----------- Set modified = false ... -----------")
            '' Set the document's modified flag to false so we are not
            '' prompted to save.
            'Call objRhinoScript.DocumentModified(False)

            'Console.WriteLine("----------- Exit Rhino ... -----------")
            '' Exit Rhino
            'Call objRhinoScript.Command("_-Exit", 0)

            Dim newThread As New Thread(New ThreadStart(AddressOf ThreadMethod))
            newThread.SetApartmentState(ApartmentState.STA)
            newThread.Start()
            newThread.Join()



        End If


        Return True

    End Function

    <STAThread()>
    Function ThreadMethod()


        Console.WriteLine("----------- Closing Grasshopper documents ... -----------")
        'Save and close all grasshopper documents so we don't get the multi-save menu.
        GH.SaveDocument()
        GH.CloseDocument()


        Console.WriteLine("----------- Set modified = false ... -----------")
        ' Set the document's modified flag to false so we are not
        ' prompted to save.
        Call objRhinoScript.DocumentModified(False)

        Console.WriteLine("----------- Exit Rhino ... -----------")
        ' Exit Rhino
        Call objRhinoScript.Command("_-Exit", 0)
    End Function

    Function FileExists(ByVal FileToTest As String) As Boolean
        FileExists = (Dir(FileToTest) <> "")
    End Function

    Sub DeleteFile(ByVal FileToDelete As String)
        If FileExists(FileToDelete) Then 'See above
            SetAttr(FileToDelete, vbNormal)
            Kill(FileToDelete)
        End If
    End Sub

    Sub ReadCmdLine()

        Dim strCmdLine As String ' Command line string   
        ' Read command line parameters into the string
        strCmdLine = GetCommandLineA

    End Sub

    Function ProcessCommandLineArgs()

        'PARAMETERS TO SET
        'GrasshopperFile
        'BakeNodeName
        'TempDir

        CommandLineArgs = My.Application.CommandLineArgs

        If Not IsNothing(CommandLineArgs) Then


            Dim flag As Boolean = False
            Dim name As String = ""
            Dim number As Integer = 0

            Dim i As Integer
            For i = 0 To CommandLineArgs.Count - 1


                Select Case CommandLineArgs(i)
                    Case "/bakename"
                        BakeNodeName = CommandLineArgs(i + 1)

                    Case "/param"
                        ParamName = CommandLineArgs(i + 1)

                    Case "/screenshotlocation"
                        ScreenshotLocation = CommandLineArgs(i + 1)

                    Case "/tempdir"
                        TempDir = CommandLineArgs(i + 1)

                    Case "/rhfile"
                        RhinoFile = CommandLineArgs(i + 1)

                    Case "/ghfile"
                        GrasshopperFile = CommandLineArgs(i + 1)

                    Case "/runid"
                        RunID = CommandLineArgs(i + 1)

                    Case "/bake"

                        If (Not String.Compare(CommandLineArgs(i + 1), "false", False)) Then
                            BakeBool = False
                        Else
                            BakeBool = True
                        End If

                    Case "/application"

                        If (Not String.Compare(CommandLineArgs(i + 1), "false", False)) Then

                            RhinoLaunchStringSuffix = "Application"
                        Else
                            RhinoLaunchStringSuffix = "Interface"
                        End If

                    Case "/newonly"
                        If (Not String.Compare(CommandLineArgs(i + 1), "false", False)) Then
                            ReimportNewOnly = False
                        Else
                            ReimportNewOnly = True
                        End If

                    Case "/?"
                        PrintUsage()
                        Return False
                    Case Else
                        PrintUsage()
                        Return False
                End Select
                i += 1
            Next

            If ReimportBool = False Then
                BakeBool = False
            End If

        End If

        CatTempExportFile = TempDir & "\_WILDCAT_TEMP_ORIG_IGS.igs"
        CatTempImportFile = TempDir & "\_WILDCAT_TEMP_MODIFIED_IGS.igs"
        OutputFile = TempDir & "\GrasshopperOutput.txt"



        Return True

    End Function



    Function PrintUsage()


        Console.WriteLine("Usage: GRASSBOT.exe /rhfile rhino-file /ghfile grasshopper-file [/bakename] [/tempdir] [/bake] [/application]" & vbNewLine & _
                          "" & vbNewLine & _
                          "/rhfile rhino-file           Full Path to Rhino file to be opened." & vbNewLine & _
                          "/ghfile grasshopper-file     Full Path to Grasshopper file to be run." & vbNewLine & _
                          "/screenshotlocation path     Full path to screenshot save location." & vbNewLine & _
                          "/bakename name               Name of the Grasshopper node to bake " & vbNewLine & _
                          "                                 - defaults to 'BAKE'" & vbNewLine & _
                          "/tempdir path                Path to temporary directory (no trailing slash)" & vbNewLine & _
                          "/bake (true/false)           If Grasshopper should bake geometry" & vbNewLine & _
                          "/param name                  TBD - Changes CATIA 'name' param based on 'name' gh node." & vbNewLine & _
                          "/application (true/false)    If Grassbot should load Rhino as Application or Interface. " & vbNewLine & _
                          "                                 - Application reopens Rhino; Interface reuses same Rhino instance." & vbNewLine)
        Console.WriteLine("Example: GRASSBOT.exe /rhfile D:/GRASSBOT/test.3dm /ghfile D:\WILDCAT\MakeCircles.gh /bakename BAKENODE /tempdir C:\TEMP")

    End Function

    Function SetDefaultParameters()

        'Set Defaults
        RhinoFile = "NOTSET"
        GrasshopperFile = "NOTSET"
        BakeNodeName = "BAKE"
        BakeActivateNodeName = "BAKEACTIVATE"
        OutputFileNodeName = "OUTPUTFILE"
        TempDir = System.IO.Path.GetTempPath()
        CatTempExportFile = TempDir & "\_WILDCAT_TEMP_ORIG_IGS.igs"
        CatTempImportFile = TempDir & "\_WILDCAT_TEMP_MODIFIED_IGS.igs"
        OutputFile = TempDir & "\GrasshopperOutput.txt"
        RunID = ""
        BakeBool = False
        ReimportBool = False
        ReimportNewOnly = False
        RunMacro = False
        ParamName = ""
        RhinoLaunchStringSuffix = "Interface"
        ScreenshotLocation = "c:\grassbot_screenshot.jpg"

        Return True

    End Function



    Function DisplayParameters()
        Console.WriteLine("PARAMETERS:" & vbNewLine)
        Console.WriteLine("RhinoFile: " & RhinoFile & vbNewLine)
        Console.WriteLine("Grasshopper File: " & GrasshopperFile & vbNewLine)
        Console.WriteLine("Grasshopper Bake Node Name: " & BakeNodeName & vbNewLine)
        Console.WriteLine("Temp Directory: " & TempDir & vbNewLine)
    End Function

    'Function RunCatiaMacro()

    '    Dim CATIA As Object

    '    CATIA = GetObject(, "CATIA.Application")

    '    Console.WriteLine("heyheyhey")



    '    Dim EmptyPar()
    '    Dim ScPath

    '    ScPath = "D:\WILDCAT"
    '    'MsgBox("Save Operation will be executed on the active document!")
    '    Sleep(3000)


    '    Dim CatSysServ As Object
    '    CatSysServ = CATIA.SystemService

    '    Dim ReturnCode = CatSysServ.ExecuteProcessus("D:\WILDCAT\test.CATScript")

    '    Console.WriteLine(ReturnCode & "fDFd")

    '    Call CatSysServ.ExecuteScript(ScPath, CatScriptLibraryType.catScriptLibraryTypeDirectory, "test.CATScript", "CATMain", EmptyPar)

    '    Dim sFilePath As String
    '    Dim sFileName As String
    '    Dim sModule As String
    '    Dim sProcedure As String
    '    Dim sFilePathAndName As String

    '    Dim Params() As Object
    '    Dim vRetVal As Object

    '    'Everything here is Case-Sensitive
    '    sFilePath = "D:\WILDCAT\"
    '    sFileName = "test.catscript"
    '    sModule = "WILDCAT"
    '    sProcedure = "CATMain" 'CatMain is only allowable Choice

    '    'Concate File Path and Name
    '    sFilePathAndName = sFilePath & "" & sFileName


    '    vRetVal = CatSysServ.ExecuteScript(sFilePathAndName, CatScriptLibraryType.catScriptLibraryTypeVBAProject, sModule, "CATMain", Params)
    '    'vRetVal = CatSysServ.ExecuteScript(sFilePath, CatScriptLibraryType.catScriptLibraryTypeVBAProject, sFileName, "CATMain", Params)
    '    'vRetVal only gets a value *if* the called macro *is* as Function,
    '    'otherwise it's 'Empty'.

    '    Console.WriteLine(vRetVal & "---")


    '    'Dim Shell As Object : Shell = CreateObject("WScript.Shell")

    '    'Dim ID As Object : ID = Shell.Run(sFilePathAndName)

    '    'Console.WriteLine(ID & "---3-3-3-3")


    'End Function

    'Function ChangeCatiaParameters(ByVal guid As Guid)

    '    ' here's what we should do
    '    ' because there is no way to get data from grasshopper
    '    ' create a line with length of the param value
    '    ' GH.BakeDataInObject() returns guid
    '    ' use guid to get length

    '    '  ChangeSingleCatiaParameter(ParamName, 1234567.0)
    '    Dim doc As Rhino.RhinoDoc
    '    Console.WriteLine("1")
    '    Dim obj As Rhino.DocObjects.RhinoObject = doc.Objects.Find(guid)
    '    Console.WriteLine("f2")
    '    Dim ret As Object
    '    ret = GH.BakeDataInObject("PARAMPARAM")

    '    Console.WriteLine("finisehd")

    'End Function

    'Function ChangeSingleCatiaParameter(ByVal paramname As String, ByVal paramval As Double)


    '    'catDoc = CATIA.ActiveDocument
    '    'catPart = catDoc.Part
    '    'catParams = catPart.Parameters
    '    Dim realParam1 As Object


    '    realParam1 = catParams.Item(paramname)
    '    realParam1.Value = paramval



    'End Function


    Sub PrintWildcatLogo()

        Console.WriteLine(vbNewLine)

        Console.WriteLine(" -- GRASSBOT -- Rhino/Grasshopper EXECUTER --" & vbNewLine & _
                            "------------------------------------------------------------------------" & vbNewLine & _
                            "     Created by Dan Taeyoung Lee, November 2013, dan@taeyounglee.com" & vbNewLine & _
                            "------------------------------------------------------------------------" & vbNewLine)


    End Sub


    Sub PrintGnuLicense()

        Console.WriteLine("-------------------------------------------------------------------------" & vbNewLine & _
                          "This program is free software: you can redistribute it and/or modify" & vbNewLine & _
    "it under the terms of the GNU General Public License as published by" & vbNewLine & _
    "the Free Software Foundation, either version 3 of the License, or" & vbNewLine & _
    "(at your option) any later version." & vbNewLine & vbNewLine & _
    "This program is distributed in the hope that it will be useful," & vbNewLine & _
    "but WITHOUT ANY WARRANTY; without even the implied warranty of" & vbNewLine & _
    "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the" & vbNewLine & _
    "GNU General Public License for more details." & vbNewLine & vbNewLine & _
    "You should have received a copy of the GNU General Public License" & vbNewLine & _
    "along with this program.  If not, see <http://www.gnu.org/licenses/>.")

    End Sub

End Module
