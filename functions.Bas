Attribute VB_Name = "basFunctions"
Option Explicit

Function CreateUNRENVBS(sPath As String)

    Dim iFileNum As Integer
    
    On Error Resume Next
    iFileNum = FreeFile
    
    Open sPath & "UNREN.VBS" For Output As #iFileNum
            Print #iFileNum, "On Error Resume Next"
            Print #iFileNum, "Set fso = CreateObject(""Scripting.FileSystemObject"")"
            Print #iFileNum, vbsUndoStr
            Print #iFileNum, "Set a = fso.GetFile(""UNREN.VBS"")"
            Print #iFileNum, "a.Delete"
            
            Print #iFileNum, "If Err Then"
            Print #iFileNum, "         MsgBox ""Errors Found, some or all files UnRenamed!"""
            Print #iFileNum, "Else"
            Print #iFileNum, "         MsgBox  ""Original Filenames Restored!"""
            Print #iFileNum, "End If"
 
    Close #iFileNum
    vbsUndoStr = ""
    
    If Err Then
            CreateUNRENVBS = False
            Err = 0
    Else
            CreateUNRENVBS = True
    End If


End Function


Function RenameTextfile(sFileName As String) As String

Dim iFF As Integer
Dim sTemp As String
        
        iFF = FreeFile
         sTemp = ""
        
        Open sFileName For Input Access Read As #iFF
                 
                 Do While (Not EOF(iFF) And sTemp = "")  'ignore empty lines, OR names like .TXT might be created!!!
                        Line Input #iFF, sTemp
                        sTemp = Trim$(ReplaceChars(sTemp, "", "\/:*?<>|" + Chr$(34)))
                 Loop
        
        Close #iFF
        

RenameTextfile = sTemp

End Function

Function CreateUNRENBAT(sPath As String) As Boolean

    Dim iFileNum As Integer
    
    On Error Resume Next
    iFileNum = FreeFile
    
    Open sPath & "UNREN.BAT" For Output As #iFileNum
            Print #iFileNum, UndoStr
            Print #iFileNum, "rem IGNORE THE FOLLOWING MESSAGE"
            Print #iFileNum, "del unren.bat"
    Close #iFileNum
    
    UndoStr = ""
    
    If Err Then
            CreateUNRENBAT = False
            Err = 0
    Else
            CreateUNRENBAT = True
    End If

End Function

Function SplitFileNameExt(FName As String) As String

Dim ThePos As Integer

ThePos = InStrRev(FName, ".")
If ThePos = 0 Then
    SplitFileNameExt = ""
Else
    SplitFileNameExt = Right(FName, Len(FName) - ThePos)
End If
    

End Function

Function SplitFileNameName(FName As String) As String

Dim ThePos As Integer
ThePos = InStrRev(FName, ".")
If ThePos = 0 Then
    SplitFileNameName = FName
Else
    SplitFileNameName = Left$(FName, ThePos - 1)
End If

End Function


Function UpEachFirst(str As String) As String
'Dim OutStr As String
'Dim i As Integer
'Dim ch As String * 1
'Dim Char As String * 1
'
'OutStr = UCase(Left(str, 1))
'
'For i = 1 To Len(str) - 1
'    ch = Mid(str, i, 1)
'        If ch = " " Or ch = "-" Or ch = "_" Or ch = "." Then
'             Char = UCase(Mid(str, i + 1, 1))
'        Else
'             Char = LCase(Mid(str, i + 1, 1))
'        End If
'
'    OutStr = OutStr + Char
'
'Next i
'UpEachFirst = OutStr

'''''''''''''''''''SPACES ONLY:
UpEachFirst = StrConv(str, vbProperCase)

End Function

Function UpFirst(str As String) As String


        Dim FirstLetter As String, OtherLetters As String
        FirstLetter = UCase$(Left(str, 1))
        OtherLetters = LCase(Right(str, Len(str) - 1))
        UpFirst = FirstLetter + OtherLetters

End Function

Function Extract_HTML_Title(sFileName As String) As String
On Error GoTo ERROR_Title_Not_Found
Dim iFileNum As Integer
Dim sLine As String
Dim sAllLines As String
Dim pos1 As Integer, pos2 As Integer
Dim sTemp As String

iFileNum = FreeFile
sLine = ""
sAllLines = ""

Open sFileName For Input Access Read Shared As #iFileNum
Do While Not EOF(iFileNum)
    Line Input #iFileNum, sLine
    sAllLines = sAllLines + sLine
    If InStr(LCase(sLine), "</title>") <> 0 Then Exit Do
    If Seek(iFileNum) > (10240) Then Exit Do      'if not found in first 10 KBytes, Abort
Loop
Close #iFileNum
pos1 = InStr(LCase(sAllLines), "<title>")
pos2 = InStr(LCase(sAllLines), "</title>")
sTemp = Mid(sAllLines, pos1 + 7, pos2 - pos1 - 7)
sTemp = Trim(sTemp)

Extract_HTML_Title = sTemp
Exit Function

'<error handler>
ERROR_Title_Not_Found:
    Extract_HTML_Title = ""
    Exit Function
'</error handler>

End Function


Public Function DelLeft(sStr As String, iCount As Integer) As String

If (iCount >= Len(sStr) Or iCount <= 0) Then
            DelLeft = sStr
Else
            DelLeft = Right(sStr, Len(sStr) - iCount)
End If

End Function


Public Function DelLeftTo(sMainStr As String, sSubStr As String, boolMatchCase As Boolean, boolInclusive As Boolean) As String

Dim iPos As Integer
If boolMatchCase = True Then
        iPos = InStr(1, sMainStr, sSubStr, vbBinaryCompare)
Else
        iPos = InStr(1, sMainStr, sSubStr, vbTextCompare)
End If


If (iPos = 0) Then
            DelLeftTo = sMainStr
Else
            If boolInclusive = True Then
                        DelLeftTo = Right(sMainStr, Len(sMainStr) - iPos - Len(sSubStr) + 1)
            Else
                        DelLeftTo = Right(sMainStr, Len(sMainStr) - iPos + 1)
            End If
End If

End Function

Public Function DelRightTo(sMainStr As String, sSubStr As String, boolMatchCase As Boolean, boolInclusive As Boolean) As String

Dim iPos As Integer
If boolMatchCase = True Then
        iPos = InStrRev(sMainStr, sSubStr, -1, vbBinaryCompare)
Else
        iPos = InStrRev(sMainStr, sSubStr, -1, vbTextCompare)
End If


If (iPos = 0) Then
            DelRightTo = sMainStr
Else
            If boolInclusive = True Then
                        DelRightTo = Left(sMainStr, iPos - 1)
            Else
                        DelRightTo = Left(sMainStr, iPos + Len(sSubStr) - 1)
            End If
End If

End Function

Public Function DelRight(sStr As String, iCount As Integer) As String

If (iCount >= Len(sStr) Or iCount <= 0) Then
            DelRight = sStr
Else
            DelRight = Left(sStr, Len(sStr) - iCount)
End If

End Function



Public Function XRename(sSrcFile As String, sTgtFile As String) As String
Dim idx As Integer
Dim sNewTgt As String, sExt As String

sNewTgt = sTgtFile

If sSrcFile = sTgtFile Then
            Exit Function
End If

On Error Resume Next

Name sSrcFile As sTgtFile

If Err = 58 Then 'File Already Exists
    
        idx = 0
        Do
                Err = 0
                idx = idx + 1
                sExt = ExtractFileExtension(sTgtFile)
                sNewTgt = Left(sTgtFile, Len(sTgtFile) - Len(sExt) - 1) + "-" + Format$(idx) + "." + sExt
                Name sSrcFile As sNewTgt
            
        Loop While Err = 58
    
End If
    
    
If Err <> 0 Then      '--------------Other errors occured
            sErrors = sErrors & "Error:  " & Error & "      FileName:  " & sSrcFile & vbCrLf
            Err = 0
            Exit Function       ' to skip adding to UndoStr and vbsUndoStr
End If

UndoStr = UndoStr + "Ren " + q + ExtractFileName(sNewTgt) + q + " " + q + ExtractFileName(sSrcFile) + q + vbCrLf
vbsUndoStr = vbsUndoStr + "Set a = fso.GetFile(" + q + ExtractFileName(sNewTgt) + q + ")" + vbCrLf
vbsUndoStr = vbsUndoStr + "a.Name=" + q + ExtractFileName(sSrcFile) + q + vbCrLf

'Return the real new name
XRename = sNewTgt

End Function

Function GetWinDir() As String
    
    Dim WinDir As String
    Dim File As String
    Dim Res As Long
    WinDir = Space$(20)
    Res = GetWindowsDirectory(WinDir, 20)
    File = Left$(WinDir, InStr(1, WinDir, Chr$(0)) - 1)
    GetWinDir = Trim$(File) & "\"
    
End Function


Public Sub LoadXRenSettings()
Dim tmps As String
tmps = GetSubVal(HKEY_LOCAL_MACHINE, "software\XRen", "")

If Len(tmps) = 0 Then             'Key does NOT exist, so...
            LoadXRenDefaults
            Exit Sub
End If

' Load Settings:

' Upper Tab:
tmps = GetSubVal(HKEY_LOCAL_MACHINE, "software\XRen", "tab")
frmXRenMain.SSTab1.Tab = CInt(tmps)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Create Undo Batch File?
tmps = GetSubVal(HKEY_LOCAL_MACHINE, "software\XRen", "UndoBatch")
oUndoBatch = CBool(tmps)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Create Undo VBScript File?
tmps = GetSubVal(HKEY_LOCAL_MACHINE, "software\XRen", "UndoVBScript")
oUndoVBScript = CBool(tmps)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Use Flat Btns?
tmps = GetSubVal(HKEY_LOCAL_MACHINE, "software\XRen", "FlatButtons")
oFlatBtns = CBool(tmps)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Selection Style
tmps = GetSubVal(HKEY_LOCAL_MACHINE, "software\XRen", "SelStyle")
frmXRenMain.File1.MultiSelect = CInt(tmps)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub

Public Sub LoadXRenDefaults()


' Load Default Settings:

' Upper Tab:
frmXRenMain.SSTab1.Tab = 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Create Undo Batch File?
oUndoBatch = True
oUndoVBScript = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Use Flat Btns?
oFlatBtns = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Selection Style
frmXRenMain.File1.MultiSelect = xfSelectSimple
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub

Function ReplaceStr(ByVal sMainStr As String, ByVal sLookFor As String, ByVal sReplaceWith As String, ByVal tfMatchCase As Boolean) As String
Dim pos As Integer
Dim sLeft As String, sRight As String


pos = 0
Do
Select Case tfMatchCase
    Case True
    pos = InStr(pos + 1, sMainStr, sLookFor, vbBinaryCompare)
    Case Else
    pos = InStr(pos + 1, sMainStr, sLookFor, vbTextCompare)
End Select


If pos <> 0 Then
        sLeft = Left(sMainStr, pos - 1)
        sRight = Right(sMainStr, Len(sMainStr) - pos - Len(sLookFor) + 1)  ': MsgBox sRight
        sMainStr = sLeft & sReplaceWith & sRight
End If

Loop Until pos = 0

ReplaceStr = sMainStr

End Function

Public Sub SaveXRenSettings()
Dim tmps As String
'Dim strMasks As String
tmps = GetSubVal(HKEY_LOCAL_MACHINE, "software\XRen", "")

If Len(tmps) = 0 Then             'Key does not exist,  so create it and set default value
                     CreateKey HKEY_LOCAL_MACHINE, "software\XRen"
                     SetSubVal HKEY_LOCAL_MACHINE, "software\XRen", "", "XRen Settings"
                     SetSubVal HKEY_LOCAL_MACHINE, "software\XRen", "Version", "3.3"
                     'strMasks = "*.htm;*.html|*.jpg;*.jpe;*.jpeg;*.jif|*.mp3|*.txt"
                     'SetSubVal HKEY_LOCAL_MACHINE, "software\XRen", "Masks", strMasks
End If

' Save current settings:
'================
' Upper Tab:
tmps = frmXRenMain.SSTab1.Tab
SetSubVal HKEY_LOCAL_MACHINE, "software\XRen", "Tab", tmps
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Create Undo Batch File?
tmps = CStr(oUndoBatch)
SetSubVal HKEY_LOCAL_MACHINE, "software\XRen", "UndoBatch", tmps
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Create Undo VBScript File?
tmps = CStr(oUndoVBScript)
SetSubVal HKEY_LOCAL_MACHINE, "software\XRen", "UndoVBScript", tmps
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Use Flat Buttons?
tmps = CStr(oFlatBtns)
SetSubVal HKEY_LOCAL_MACHINE, "software\XRen", "FlatButtons", tmps
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Select Style:
tmps = frmXRenMain.File1.MultiSelect
SetSubVal HKEY_LOCAL_MACHINE, "software\XRen", "SelStyle", tmps
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub
Function ReplaceChars(ByVal astr As String, ByVal ReplaceWith As String, ByVal UnwantedChars As String) As String
Dim tmpStr As String
Dim ch As String
Dim i As Integer
Dim bReplaceExclamation As Boolean
Dim bReplaceLeftBracket As Boolean
Dim bReplaceRightBracket As Boolean

tmpStr = ""

For i = 1 To Len(UnwantedChars)
    ch = Mid$(UnwantedChars, i, 1)
    '  "!" and "[" and "]" have special meaning to LIKE, they will be handeled manullay
    If ch = "!" Then ch = "": bReplaceExclamation = True
    If ch = "[" Then ch = "": bReplaceLeftBracket = True
    If ch = "]" Then ch = "": bReplaceRightBracket = True
    tmpStr = tmpStr + ch
Next i
UnwantedChars = tmpStr

tmpStr = ""
ch = ""

'If Left(UnwantedChars, 1) <> "[" Then UnwantedChars = "[" + UnwantedChars
'If Right(UnwantedChars, 1) <> "]" Then UnwantedChars = UnwantedChars + "]"

UnwantedChars = "[" & UnwantedChars & "]"

For i = 1 To Len(astr)
    ch = Mid$(astr, i, 1)
    '  "!" and "[" and "]" have special meaning to LIKE
    If (ch = "!" And bReplaceExclamation) Then ch = ReplaceWith
    If (ch = "[" And bReplaceLeftBracket) Then ch = ReplaceWith
    If (ch = "]" And bReplaceRightBracket) Then ch = ReplaceWith
    If ch Like UnwantedChars Then
        ch = ReplaceWith
        If Right$(tmpStr, 1) = ReplaceWith Then ch = ""
    End If
    
    tmpStr = tmpStr + ch
Next i
ReplaceChars = tmpStr

End Function

Function CenterFormUp(frmForm As Form)  'as void
frmForm.Left = (Screen.Width - frmForm.Width) / 2
frmForm.Top = (Screen.Height - frmForm.Height) / 3
End Function

Function ExtractDirName(FileName As String) As String

'Extract the Directory name from a full file name
    Dim tmp$
    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, "\")
    If pos = 0 Then
        ExtractDirName = ""
        Exit Function
    End If
    
    Do While pos <> 0
        PrevPos = pos
        pos = InStr(pos + 1, FileName, "\")
    Loop

    tmp = Left(FileName, PrevPos)
    If Right(tmp, 1) = "\" Then tmp = Left(tmp, Len(tmp) - 1)
    tmp = tmp & "\" 'COOL?
    ExtractDirName = tmp
    
End Function
Function ExtractFileName(FileName As String) As String
    
'Extract the File title from a full file name


    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, "\")
    If pos = 0 Then
            ExtractFileName = FileName
            Exit Function
    End If
    
    Do While pos <> 0
            PrevPos = pos
            pos = InStr(pos + 1, FileName, "\")
    Loop

    ExtractFileName = Right(FileName, Len(FileName) - PrevPos)

End Function



Function RevRGB(hexRGB As String) As String
Dim Var1 As String
Dim Var2 As String
Dim Var3 As String

Var1 = Left(hexRGB, 2)
Var2 = Mid(hexRGB, 3, 2)
Var3 = Right(hexRGB, 2)

RevRGB = Var3 & Var2 & Var1

End Function


Function SaveFile(FileName As String, FileContent As String) As Boolean
On Error GoTo Save_Error
Dim FileNum As Integer

FileNum = FreeFile

Open FileName For Output As #FileNum

Print #FileNum, FileContent

Close FileNum
SaveFile = True
Exit Function

Save_Error:
SaveFile = False
Exit Function
End Function


Function LoadFile(FileName As String) As String
'Loads the contents of a file into a string variable

On Error GoTo LoadFile_Error
Dim ff As Integer
Dim FileContents As String

ff = FreeFile
Open FileName For Input As #ff
FileContents = Input(LOF(ff), ff)
Close #ff
LoadFile = FileContents
Exit Function
LoadFile_Error:
    LoadFile = "#ERROR#"
    Exit Function

End Function

Function AddBR(sText As String) As String
Dim sTemp As String
Dim idx As Long
ReDim blines(1 To 1) As String
sTemp = sText
Text2Lines sTemp, blines()

sTemp = ""
For idx = LBound(blines) To UBound(blines)
    sTemp = sTemp + blines(idx) + "<BR>" + vbCrLf
Next idx


AddBR = sTemp




End Function


Function SelectAllText(txtBox As TextBox) 'as void

txtBox.SelStart = 0
txtBox.SelLength = Len(txtBox.Text)

End Function

Sub Text2Lines(Text As String, Lines() As String)
Dim ch As String * 1
Dim Cntr As Long
Dim Index As Integer
Dim MaxIndex As Integer
Dim NewLine As String * 2

NewLine = Chr(13) + Chr(10)

ReDim Lines(1 To 9000)

Index = 1
For Cntr = 1 To Len(Text)
    ch = Mid$(Text, Cntr, 1)
    Select Case Asc(ch)
        Case 13
            'do nothing
        Case 10     'always after the 13
            Index = Index + 1
        Case Else

            Lines(Index) = Lines(Index) + ch
    End Select
Next Cntr

MaxIndex = Index

ReDim Preserve Lines(1 To MaxIndex)
End Sub

Function ChangeFileExtension(FileName As String, NewExtension As String) As String
Dim OldExt As String
OldExt = ExtractFileExtension(FileName)
ChangeFileExtension = Left$(FileName, Len(FileName) - Len(OldExt)) & NewExtension

End Function

Public Function ExtractFileExtension(FileName As Variant) As String

    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, ".")
    If pos = 0 Then
    ExtractFileExtension = ""
    Exit Function
    End If
    
    Do While pos <> 0
    PrevPos = pos
    pos = InStr(pos + 1, FileName, ".")
    Loop

    ExtractFileExtension = Right(FileName, Len(FileName) - PrevPos)

End Function

