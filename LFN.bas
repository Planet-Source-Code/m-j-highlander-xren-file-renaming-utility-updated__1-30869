Attribute VB_Name = "Sh2LFN"
Option Explicit

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

' Return the short file name for a long file name.
Public Function ShortFileName(ByVal long_name As String) As String
Dim length As Long
Dim short_name As String

    short_name = Space$(1024)
    length = GetShortPathName( _
        long_name, short_name, _
        Len(short_name))
    ShortFileName = Left$(short_name, length)
End Function


' Return the long file name for a short file name.
Public Function LongFileName(ByVal short_name As String) As String
Dim pos As Integer
Dim result As String
Dim long_name As String

If Left(short_name, 1) = Chr$(34) Then
          short_name = Right(short_name, Len(short_name) - 1)
End If

If Right(short_name, 1) = Chr$(34) Then
          short_name = Left(short_name, Len(short_name) - 1)
End If
          
    ' Start after the drive letter if any.
    If Mid$(short_name, 2, 1) = ":" Then
        result = Left$(short_name, 2)
        pos = 3
    Else
        result = ""
        pos = 1
    End If

    ' Consider each section in the file name.
    Do While pos > 0
        ' Find the next \.
        pos = InStr(pos + 1, short_name, "\")

        ' Get the next piece of the path.
        If pos = 0 Then
            long_name = Dir$(short_name, vbNormal + vbHidden + vbSystem + vbDirectory)
        Else
            long_name = Dir$(Left$(short_name, pos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
        End If
        result = result & "\" & long_name
    Loop

    LongFileName = result
End Function

