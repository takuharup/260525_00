' import_vba.vbs
' Usage: cscript import_vba.vbs <full path to xlsm file>
' After import, run Module1.CreateUserForm() once from Excel to create UserForm1.

Option Explicit

Dim xlApp, wb, vbProj, xlsmPath, scriptDir

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript import_vba.vbs <xlsm-file-path>"
    WScript.Quit 1
End If

xlsmPath = WScript.Arguments(0)
scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Set wb = xlApp.Workbooks.Open(xlsmPath)
Set vbProj = wb.VBProject

On Error Resume Next
vbProj.VBComponents.Remove vbProj.VBComponents("Module1")
Err.Clear
On Error GoTo 0

vbProj.VBComponents.Import scriptDir & "Module1.bas"

wb.Save
wb.Close
xlApp.Quit

WScript.Echo "Done. Next: open the xlsm and run Module1.CreateUserForm()"
