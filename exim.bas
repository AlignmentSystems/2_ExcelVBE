Attribute VB_Name = "exim"
Option Explicit
Option Compare Text
Option Base 0
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       24th March 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Private Function EntryPointModulesDelete() As Boolean
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       24th March 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const cstrMethodName As String = "exim.EntryPointModulesDelete "
'Here we delete the modules we


End Function


Private Function EntryPointModulesExport() As Boolean
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       24th March 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const cstrMethodName As String = "exim.EntryPointModulesExport "
'Variables
Dim lngNumberOfExportableModules As Long
Dim lngExportedModuleCount As Long
Dim oWorkbook As Excel.Workbook
Dim oCom As VBIDE.VBComponent
Dim strTargetFileName As String

Set oWorkbook = ThisWorkbook

lngNumberOfExportableModules = GetExportableComponentCount(oWorkbook)

'Why?  Simple, so we can later on replace this method with no parameters with an updated version that
'allows the user to set the workbook on which to perform this operation...
For Each oCom In oWorkbook.VBProject.VBComponents
    strTargetFileName = ""
    If GetValidSaveAsNameAndPath(oWorkbook, oCom, strTargetFileName) Then
        oCom.Export strTargetFileName
        Debug.Print (cstrMethodName & "exported " & oCom.Name & " to " & strTargetFileName)
        lngExportedModuleCount = lngExportedModuleCount + 1
    End If
Next

Debug.Print (cstrMethodName & "exportable=" & lngNumberOfExportableModules & " exported=" & lngExportedModuleCount)

End Function

Private Function GetExportableComponentCount(TargetWorkbook As Excel.Workbook) As Long
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       24th March 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const cstrMethodName As String = "exim.GetExportableComponentCount "
'Variables
Dim lngIncrementalCount As Long
Dim lngTotal As Long
Dim oCom As VBIDE.VBComponent
Dim blnTest As Boolean
Dim strExtension As String

lngTotal = 0
For Each oCom In TargetWorkbook.VBProject.VBComponents
    strExtension = ""
    CanExportDelete oCom, strExtension, False
    If Len(strExtension) > 0 Then
        lngTotal = lngTotal + 1
    End If
Next

GetExportableComponentCount = lngTotal

End Function

Private Function CanExportDelete(TargetComponent As VBIDE.VBComponent, ExportExtension As String, CanDelete As Boolean) As Boolean
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       24th March 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const cstrMethodName As String = "exim.CanExportDelete "

ExportExtension = ""
CanDelete = False
CanExportDelete = False


Select Case TargetComponent.Type
    Case VBIDE.vbext_ComponentType.vbext_ct_ActiveXDesigner
        ExportExtension = ""
    Case VBIDE.vbext_ComponentType.vbext_ct_Document
        ExportExtension = ""
    Case VBIDE.vbext_ComponentType.vbext_ct_ClassModule
        ExportExtension = "cls"
    Case VBIDE.vbext_ComponentType.vbext_ct_MSForm
        ExportExtension = "frm"
    Case VBIDE.vbext_ComponentType.vbext_ct_StdModule
        ExportExtension = "bas"
    Case Else
        ExportExtension = ""
End Select

End Function

Private Function GetValidSaveAsNameAndPath(TargetWorkbook As Excel.Workbook, TargetComponent As VBIDE.VBComponent, ReturnString As String) As Boolean
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       24th March 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const cstrMethodName As String = "exim.GetValidSaveAsNameAndPath "
Dim strExtension As String

CanExportDelete TargetComponent, strExtension, False

If Len(strExtension) = 0 Then
    GetValidSaveAsNameAndPath = False
Else
'   By default we will use the location of this workbook...
    ReturnString = TargetWorkbook.Path & "\" & TargetComponent.Name & "." & strExtension
    GetValidSaveAsNameAndPath = True
End If

End Function
