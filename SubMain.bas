Attribute VB_Name = "Sub_Main_Public_Declares"
Option Explicit
'*************** API Functions
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


'App-wide variables:
Public Const q = """"
'***************XRen Options**********************
Public oUndoBatch As Boolean
Public oUndoVBScript As Boolean
Public oFlatBtns As Boolean
'***************Rename Error Log
Public sErrors As String
'***************UnRen.bat/vbs contents
Public UndoStr As String
Public vbsUndoStr As String



Public gv_Args As Variant

Function GetCmdLineArgs() As Variant

GetCmdLineArgs = Split(Command$, "*")

End Function


'Program Execution Starts Here
Sub Main()

RegisterShellEx

gv_Args = GetCmdLineArgs
    

'Load frmShellNotify      '///////turned off ONLY for now
 
 frmXRenMain.Show


End Sub


Sub RegisterShellEx()
Dim sVal As String
Dim ap As String

Dim sTemp As String
sTemp = GetSubVal(HKEY_LOCAL_MACHINE, "software\XRen", "AutoRegisterShellExt")
If sTemp = "0" Then
    Exit Sub
End If

ap = App.Path
If Right(ap, 1) <> "\" Then ap = ap & "\"


 sVal = GetSubVal(HKEY_CLASSES_ROOT, "*\shellex\ContextMenuHandlers\XRenContext", "")

If sVal = "" Then       'do the registration stuff
    'Register DLL
    Shell "RegSvr32.exe" & " " & q & ap & "XRenShx.dll" & q & " /s"
    'Create MenuHandler
    SetSubVal HKEY_CLASSES_ROOT, "*\shellex\ContextMenuHandlers\XRenContext", "", "{23FCFE69-A54B-11D4-8AD0-484C000107C0}"
    MsgBox "XRen Shell Extension Installation Succeeded." & vbCrLf & "Now you can right-click files in Explorer and select XRen", vbInformation, "XRen"
Else
            'all ok, do nothing

End If

End Sub


