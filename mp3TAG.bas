Attribute VB_Name = "MP3_ID_TAG"

Function Read_ID3Tag(sFileName As String) As Variant
'On Error GoTo ERROR_UNKNOWN
Dim iFileNum As Integer
Dim Return_Array(1 To 3) As String
Dim sTag As String * 127
Dim Title$, Artist$, Album$

iFileNum = FreeFile

Open sFileName For Binary Access Read Shared As #iFileNum

fl = LOF(1)
Seek #1, (fl - 128) + 1  'VB OFFSET!!!
Get #1, , sTag
Close #1

If Left(sTag, 3) <> "TAG" Then
    Read_ID3Tag = ""
Else
    Title = Trim(Mid(sTag, 4, 30))
    Artist = Trim(Mid(sTag, 34, 30))
    Album = Trim(Mid(sTag, 64, 30))
    Return_Array(1) = RTrimNulls(Title)
    Return_Array(2) = RTrimNulls(Artist)
    Return_Array(3) = RTrimNulls(Album)
    Read_ID3Tag = Return_Array
End If

End Function
Function RTrimNulls(sStr As String) As String
Dim idx As Integer
Dim ch As String * 1
For idx = Len(sStr) To 1 Step -1
    If Asc(Mid(sStr, idx, 1)) <> 0 Then Exit For
Next idx

RTrimNulls = Left(sStr, idx)
End Function


