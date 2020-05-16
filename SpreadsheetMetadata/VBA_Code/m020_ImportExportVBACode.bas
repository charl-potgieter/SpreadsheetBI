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

Public Sub ExportVBAModules(ByRef wkb As Workbook, ByVal sFolderPath As String)
'Saves active workbook and exports file to sFolderPath
' *****IMPORTANT NOTE****
' Any existing files will be overwritten

    Dim sExportFileName As String
    Dim bExport As Boolean
    Dim sFileName As String
    Dim cmpComponent As VBIDE.VBComponent


    If wkb.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to export the code"
        Exit Sub
    End If
    
    For Each cmpComponent In wkb.VBProject.VBComponents
        
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
            sExportFileName = sFolderPath & Application.PathSeparator & sFileName
            cmpComponent.Export sExportFileName
        End If
   
    Next cmpComponent


End Sub


Public Sub ImportModules(ByRef wkb As Workbook, ByVal sFolder As String)
'Imports VBA code sFolder


    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim sTargetWorkbook As String
    Dim sImportPath As String
    Dim zFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If wkb.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If
    
    If wkb.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(sFolder).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If


    Set cmpComponents = wkb.VBProject.VBComponents
    
    For Each objFile In objFSO.GetFolder(sFolder).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
        
    Next objFile
    
    MsgBox "Import Complete"
End Sub


