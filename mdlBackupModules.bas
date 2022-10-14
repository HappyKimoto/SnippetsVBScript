Attribute VB_Name = "mdlBackupModules"
Option Explicit

'
' Objective:
'   - Want to manage VBA codes in text files, instead of *.xlsm (Macro enabled excel book).
'
' Features:
' (1) Export modules
'   - Export:
'       - Modules
'       - Class Modules
'       - Forms
'   - Do not export:
'       - Sheet#
'       - ThisWorkbook
'   - Export the VBA script files to the specified folder
'
' (2) Export
'   - Create a list of enabled referenes and export it in a text file.
'
' Preparation:
'   - Enable [Trust access to the VBA project object model].
'       - Go to: File > Options
'       - Go to: Excel Optoins > Trust Center > Microsoft Excel Trust Center > Trust Center Settings...
'       - Go to: Trust Center > Macro Settings > Developer Macro Settings
'       - Turn on: Trust access to the VBA project object model
'
' How to use:
'   - Open Excel Workbook (*.xlsm) where the VBA modules are stored.
'   - Open VBA Editor (Alt + F4).
'   - Drag [mdlBackupModules.bas] from file explorer and
'   - Drop it onto the target [VBAProject] in VBA Project explorer.
'       - Ctrl + R to open Project Explorer (if not opened yet).
'   - Run procedure "Public Sub BackupModules()".
'       - On VBA Editor, Set the cursor within BackupModules() and press F5.
'       - On Excel sheet, press [Alt + F8] to open Macro and run [BackupModules].
'   - Select output folder.
'   - Press [OK] to close the confirmation message.
'          +-------------------------------+
'          | Backup Modules            [X] |
'          +-------------------------------+
'          | Completed.                    |
'          | 4 scripts have been exported. |
'          |                        [OK]   |
'          +-------------------------------+
'   - Confirm that expected files are exported to the the output folder.
'       - The contents of the output folder may look like below:
'        +----------------------------------+
'        | \---Test.xlsm(20221014121741)    | Folder naming convention is "FileName(yyyymmddhhmmss)"
'        |     |   Tools-References.txt     | Currently enabled references are listed here.
'        |     |                            |
'        |     \---Modules                  | Scripts are stored under this folder.
'        |             Class1.cls           |
'        |             mdlBackupModules.bas |
'        |             Module1.bas          |
'        |             UserForm1.frm        |
'        |             UserForm1.frx        |
'        +----------------------------------+

' Assumption:
'   - Do not want to store data or user interface on Worksheets because:
'       - File size can grow exponentially as you input data on Worksheets.
'       - It is hard to track changes on sheet because there are many hidden properties on Worksheets.
'
'   - Want to input data on Worksheets dynamically by creating new sheets if needed.

' ---------------------------------------------------------
' This module was developed by referencing to the following
' ---------------------------------------------------------
'
' https://social.msdn.microsoft.com/Forums/en-US/82bc5825-d74f-4e21-b7ac-376578f18d7f/export-modules-and-class-modules?forum=exceldev
'
' Adapt the following, particularly the path
'
' Sub ExportModules()
' Dim sPath As String
' Dim sFile As String
' Dim vbp As Object ' VBProject
' Dim comp As Object ' VBComponent
'
'     sPath = Application.DefaultFilePath & "\Test\"
'     On Error Resume Next
'     MkDir sPath
'     On Error GoTo 0
'
'     Set vbp = Workbooks("Book1").VBProject
'     For Each comp In vbp.VBComponents
'         sFile = ""
'         Select Case comp.Type
'         Case 1 ' vbext_ct_StdModule
'             sFile = comp.Name & ".mod"
'         Case 2 ' vbext_ct_ClassModule
'             sFile = comp.Name & ".cls"
'         Case 3 ' vbext_ct_MSForm
'             sFile = comp.Name & ".frm"
'             ' the frx will automatically be exported
'         Case 100 ' vbext_ct_Document
'             ' thisworkbook or sheet module
'             ' this will re-import as a class module
'             ' copy code to relevant object module then remove
'             If comp.CodeModule.CountOfLines Then
'                 sFile = comp.Name & ".cls"
'             End If
'         End Select
'         If Len(sFile) Then
'             comp.Export sPath & sFile
'         End If
'     Next
'
' End SubNote security settings must allow
' "trust access to the VBA project object model" in Trust-Center, Macro settings

' ---------------------------------------------------------
' Syntax of MsgBox
' ---------------------------------------------------------
'
' MsgBox( prompt [, buttons ] [, title ] [, helpfile, context ] )
'
' | prompt     | This is a required argument.                                                                       |
' | ---------- | -------------------------------------------------------------------------------------------------- |
' | [buttons]  | It determines what buttons and icons are displayed in the MsgBox.                                  |
' | [title]    | Here you can specify what caption you want in the message dialog box.                              |
' | [helpfile] | You can specify a help file that can be accessed when a user clicks on the Help button.            |
' | [context]  | It is a numeric expression that is the Help context number assigned to the appropriate Help topic. |
'
' | Button Constant    | Description                                |
' | ------------------ | ------------------------------------------ |
' | vbOKOnly           | Shows only the OK button                   |
' | vbOKCancel         | Shows the OK and Cancel buttons            |
' | vbAbortRetryIgnore | Shows the Abort, Retry, and Ignore buttons |
' | vbYesNo            | Shows the Yes and No buttons               |
' | vbYesNoCancel      | Shows the Yes, No, and Cancel buttons      |
' | vbRetryCancel      | Shows the Retry and Cancel buttons         |
'
' | Icon Constant | Description                     |
' | ------------- | ------------------------------- |
' | vbCritical    | Shows the critical message icon |
' | vbQuestion    | Shows the question icon         |
' | vbExclamation | Shows the warning message icon  |
' | vbInformation | Shows the information icon      |
'

' Purpose: Get a folder path from the user
' Return: Folder path
' Assumption: The user chooses a folder. If the user did not select a folder, empty string will be returned.
Private Function strFolderPath(ByVal strTitle As String) As String
    Dim strReturn As String: strReturn = ""
    Dim objDialog As Object: Set objDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With objDialog
        .Title = strTitle
        .AllowMultiSelect = False
        If .Show = -1 Then
            strReturn = .SelectedItems(1)
        End If
    End With
    Set objDialog = Nothing
    strFolderPath = strReturn
End Function

' Purpose:
'   Check if Developer Macro Setting is set correctly.
' Action Required:
'   If you encounter runtime error 1004:
'       - Go to: File > Options
'       - Go to: Excel Optoins > Trust Center > Microsoft Excel Trust Center > Trust Center Settings...
'       - Go to: Trust Center > Macro Settings > Developer Macro Settings
'       - Turn on: Trust access to the VBA project object model
' Assumption:
'   Runtime error 1004 may looke like as below
'       Microsoft Visual Basic
'       Run-time error '1004':
'       Method 'VBProject' of object '_Workbook' failed.
'       [Continue] [End] [Debug] [Help]
Private Sub TestError1004()
    Debug.Print "TypeName(Application.ThisWorkbook.VBProject):" & TypeName(Application.ThisWorkbook.VBProject)
    Debug.Print "TypeName(ActiveWorkbook.VBProject.VBComponents)" & TypeName(ActiveWorkbook.VBProject.VBComponents)
End Sub

Private Function strModuleFileNameTxt(ByRef objVBComponent As Object) As String
    ' [Forms], [Class Modules], and [Modules] are exported.
    ' File extensions are all set in txt.
    ' Regardless of file extension, VBA Editor will figure out the type by analyzing the header section,
    ' as you drag and drop those text files from file explorer to VBA Project explorer.
    Dim strModuleFileName As String: strModuleFileName = ""
    Select Case objVBComponent.Type
        Case 1  ' Modules
            strModuleFileName = objVBComponent.Name & ".txt"
        Case 2 ' Class Modules
            strModuleFileName = objVBComponent.Name & ".txt"
        Case 3 ' Forms
            strModuleFileName = objVBComponent.Name & ".txt"
            ' In case of Forms, a file named like objVBComponent.Name + .frx will be also exported.
            ' To re-use the form, you will want to
        Case 100 ' ThisWorkbook, Sheet1, Sheet2, Sheet3, ...
            ' Do not export modules for book or sheets
            ' strModuleFileName = objVBComponent.Name & ".txt"
    End Select
    strModuleFileNameTxt = strModuleFileName
End Function

Private Function strModuleFileNameStd(ByRef objVBComponent As Object) As String
    ' [Forms], [Class Modules], and [Modules] are exported.
    ' File extensions are the same for manual export.
    Dim strModuleFileName As String: strModuleFileName = ""
    Select Case objVBComponent.Type
        Case 1  ' Modules
            strModuleFileName = objVBComponent.Name & ".bas"
        Case 2 ' Class Modules
            strModuleFileName = objVBComponent.Name & ".cls"
        Case 3 ' Forms
            strModuleFileName = objVBComponent.Name & ".frm"
            ' In case of Forms, a file named like objVBComponent.Name + .frx will be also exported.
            ' To re-use the form, you will want to
        Case 100 ' ThisWorkbook, Sheet1, Sheet2, Sheet3, ...
            ' Do not export modules for book or sheets
            ' It appears ThisWorkbook, Sheet1, and Sheet2 are exported with file extension cls.
            ' strModuleFileName = objVBComponent.Name & ".cls"
    End Select
    strModuleFileNameStd = strModuleFileName
End Function

' Purpose: Get Tools/References
' Return: String of Reference descriptions
Private Function strToolsReferences()
    Dim strRef As String: strRef = ""
    Dim objRef As Object
    For Each objRef In Application.VBE.ActiveVBProject.References
        Debug.Print "Name: " & objRef.Name
        Debug.Print "Description: " & objRef.Description
        strRef = strRef + objRef.Description + vbCrLf
    Next
    strToolsReferences = strRef
End Function

' Purpose: Write text to file
' Inputs:
'   - ByVal strFp as absolute file path
'   - ByVal strTxt as text contents
' Effects:
'   - strTxt will be written to strFp
Private Sub WriteFile(ByVal strFp, ByVal strTxt)
    Dim objFSO As Object: Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFile As Object: Set objFile = objFSO.CreateTextFile(strFp)
    objFile.WriteLine strTxt
    objFile.Close
    Set objFile = Nothing
    Set objFSO = Nothing
End Sub

' Purpose: Export modules and references to text files.
Public Sub BackupModules()
    ' Make sure [Trust access to the VBA project object model] is enabled
    TestError1004
     
    ' Let the user selcect an output.
    ' A new folder is created based on excel file name and current date/time.
    ' Current date/time helps the output folder name be unique.
    Dim strParentDir As String: strParentDir = strFolderPath("Select output folder")
    Dim strFileName As String: strFileName = Application.ThisWorkbook.Name
    Dim strDateTime As String: strDateTime = "(" & Format(Now, "yyyyMMddhhmmss") & ")"
    Dim strOutDir As String: strOutDir = strParentDir & "\" & strFileName & strDateTime
    Debug.Print "strOutDir: " & strOutDir
    Call MkDir(strOutDir)
    
    ' Get checked items from [Tools] > [References...]
    ' If you are using the exported modules on a new workbook,
    ' you may want to add those references on the target workbook.
    Dim strRefList As String: strRefList = strToolsReferences()
    Dim strRefFilePath As String: strRefFilePath = strOutDir & "\Tools-References.txt"
    WriteFile strRefFilePath, strRefList
        
    ' Export modules
    ' Create a sub folder named "Modules" where the modules are to be saved
    Dim strModulesDir As String
    strModulesDir = strOutDir & "\Modules"
    Debug.Print "strModulesDir: " & strModulesDir
    Call MkDir(strModulesDir)
    ' loop through VBComponents
    Dim objVBComponent As Object
    Dim strModuleFileName As String
    Dim intExportCount As Integer: intExportCount = 0
    For Each objVBComponent In Application.ThisWorkbook.VBProject.VBComponents
    
        ' create output file name
        Debug.Print "VBComponent Name: " & objVBComponent.Name & "; Type: " & objVBComponent.Type
        strModuleFileName = strModuleFileNameStd(objVBComponent)
        ' strModuleFileName = strModuleFileNameTxt(objVBComponent)
        
        ' Syntax
        ' object.Export (filename)
        ' The Export syntax has these parts:
        ' | Part     | Description                                                                                  |
        ' | -------- | -------------------------------------------------------------------------------------------- |
        ' | object   | Required. An object expression that evaluates to an object in the Applies To list.           |
        ' | filename | Required. A String specifying the name of the file that you want to export the component to. |
        '
        ' Export if VBComponent type is the expected type.
        If Len(strModuleFileName) > 0 Then
            objVBComponent.Export strModulesDir & "\" & strModuleFileName
            Debug.Print Now & " Exported " & strModuleFileName
            intExportCount = intExportCount + 1
        End If
    Next objVBComponent
    
    ' Notify the user about how many many files have been exported.
    MsgBox "Completed." & vbCrLf & intExportCount & " scripts have been exported.", Title:="Backup Modules"
    
End Sub

