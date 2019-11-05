Attribute VB_Name = "m020_ImportExportVBACode"
Option Explicit

'------------------------------------------------------------------------------------------------------------------------
'   Code inspired by Ron De Bruin:
'   https://www.rondebruin.nl/win/s9/win002.htm
'
'   Requires references
'    - Microsoft Visual Basic For Applications Extensibility 5.3
'    - Microsoft Scripting Runtime
'
'   Requires Trust Access to VBA module:
'   In Excel 2007-2013, click the Developer tab and then click the Macro Security item. In that dialog,
'   choose Macro Settings and check the Trust access to the VBA project object model
'
'   Be aware that above may trigger action from antivirus software
'
'------------------------------------------------------------------------------------------------------------------------

Public Sub ExportVBAModules()
'Saves active workbook and exports file to VBA_Code subfolder in path of active workbook
' *****IMPORTANT NOTE****
' Any existing files in this subfolder will be deleted

    Dim sExportPath As String
    Dim sExportFileName As String
    Dim bExport As Boolean
    Dim sFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    
    ActiveWorkbook.Save
    sExportPath = ThisWorkbook.Path & Application.PathSeparator & "VBA_Code"
    On Error Resume Next
        MkDir sExportPath
        Kill sExportPath & "\*.*"
    On Error GoTo 0

    If ActiveWorkbook.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to export the code"
        Exit Sub
    End If
    
    For Each cmpComponent In ActiveWorkbook.VBProject.VBComponents
        
        bExport = True
        sFileName = cmpComponent.Name

        'Set filename
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                sFileName = cmpComponent.Name & ".cls"
            Case vbext_ct_MSForm
                sFileName = cmpComponent.Name & ".frm"
            Case vbext_ct_StdModule
                sFileName = cmpComponent.Name & ".bas"
            Case vbext_ct_Document
                ' This is a worksheet or workbook object - don't export.
                bExport = False
        End Select
        
        If bExport Then
            sExportFileName = sExportPath & Application.PathSeparator & sFileName
            cmpComponent.Export sExportFileName
        End If
   
    Next cmpComponent

    MsgBox "Code export complete"
End Sub


Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
'    If FolderWithVBAProjectFiles = "Error" Then
'        MsgBox "Import Folder not exist"
'        Exit Sub
'    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = "C:\Users\charl\Dropbox\Dropbox_Charl\Computer_Technical\Programming_GitHub\SpreadsheetBI\VBA_Code\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    'Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
        
    Next objFile
    
    MsgBox "Import is ready"
End Sub

'Function FolderWithVBAProjectFiles() As String
'    Dim WshShell As Object
'    Dim FSO As Object
'    Dim SpecialPath As String
'
'    Set WshShell = CreateObject("WScript.Shell")
'    Set FSO = CreateObject("scripting.filesystemobject")
'
'    SpecialPath = WshShell.SpecialFolders("MyDocuments")
'
'    If Right(SpecialPath, 1) <> "\" Then
'        SpecialPath = SpecialPath & "\"
'    End If
'
'    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
'        On Error Resume Next
'        MkDir SpecialPath & "VBAProjectFiles"
'        On Error GoTo 0
'    End If
'
'    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
'        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
'    Else
'        FolderWithVBAProjectFiles = "Error"
'    End If
'
'End Function

'Function DeleteVBAModulesAndUserForms()
'        Dim VBProj As VBIDE.VBProject
'        Dim VBComp As VBIDE.VBComponent
'
'        Set VBProj = ActiveWorkbook.VBProject
'
'        For Each VBComp In VBProj.VBComponents
'            If VBComp.Type = vbext_ct_Document Then
'                'Thisworkbook or worksheet module
'                'We do nothing
'            Else
'                VBProj.VBComponents.Remove VBComp
'            End If
'        Next VBComp
'End Function



