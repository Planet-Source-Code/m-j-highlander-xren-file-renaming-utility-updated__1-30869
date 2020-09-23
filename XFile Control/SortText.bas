Attribute VB_Name = "SortArray"
Option Explicit
Public Sub ShellSortExtDescEx(SortArray As Variant)


Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim Swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            If (LCase(SortArray(Row)) < LCase(SortArray(Row + Offset)) And LCase(ExtractFileExtension(CStr(SortArray(Row)))) = LCase(ExtractFileExtension(CStr(SortArray(Row + Offset))))) Then
               Swap SortArray(Row), SortArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
End Sub

Public Sub ShellSortExtAscEx(SortArray As Variant)


Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim Swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            If (LCase(SortArray(Row)) > LCase(SortArray(Row + Offset)) And LCase(ExtractFileExtension(SortArray(Row))) = LCase(ExtractFileExtension(SortArray(Row + Offset)))) Then
               Swap SortArray(Row), SortArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
End Sub



Public Sub ShellSortDateAsc(SortArray As Variant)
'The fastets sort algorithm!

Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim Swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            If FileDateTime(SortArray(Row)) > FileDateTime(SortArray(Row + Offset)) Then
               Swap SortArray(Row), SortArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop

    
Exit Sub

End Sub


Public Sub ShellSortDateDesc(SortArray As Variant)
'The fastets sort algorithm!

Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim Swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            If FileDateTime(SortArray(Row)) < FileDateTime(SortArray(Row + Offset)) Then
               Swap SortArray(Row), SortArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
End Sub

Public Sub ShellSortSizeAsc(SortArray As Variant)
'The fastets sort algorithm!

Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim Swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            If FileLen(SortArray(Row)) > FileLen(SortArray(Row + Offset)) Then
               Swap SortArray(Row), SortArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop

    
Exit Sub

End Sub


Public Sub ShellSortSizeDesc(SortArray As Variant)
'The fastets sort algorithm!

Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim Swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            If FileLen(SortArray(Row)) < FileLen(SortArray(Row + Offset)) Then
               Swap SortArray(Row), SortArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
End Sub


'Public Function ExtractFileExtension(ByVal FileName As String) As String
'
'    Dim pos As Integer
'    Dim PrevPos As Integer
'
'    pos = InStr(FileName, ".")
'    If pos = 0 Then
'    ExtractFileExtension = ""
'    Exit Function
'    End If
'
'    Do While pos <> 0
'    PrevPos = pos
'    pos = InStr(pos + 1, FileName, ".")
'    Loop
'
'    ExtractFileExtension = Right(FileName, Len(FileName) - PrevPos)
'
'End Function


Public Sub ShellSortExtAsc(SortArray As Variant)
'The fastets sort algorithm!

'same ext should be sorted by name!
'Calling "ShellSortAsc SortArray" didn't work :-(


Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim Swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            
    If LCase(ExtractFileExtension(SortArray(Row))) > LCase(ExtractFileExtension(SortArray(Row + Offset))) Then
               Swap SortArray(Row), SortArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
End Sub


Public Sub ShellSortExtDesc(SortArray As Variant)
'The fastets sort algorithm!

Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim Swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            
    If LCase(ExtractFileExtension(SortArray(Row))) < LCase(ExtractFileExtension(SortArray(Row + Offset))) Then
               Swap SortArray(Row), SortArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
   

End Sub


Public Sub ShellSortAsc(SortArray As Variant)
'The fastets sort algorithm!

Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim Swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            If LCase(SortArray(Row)) > LCase(SortArray(Row + Offset)) Then
               Swap SortArray(Row), SortArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
End Sub


Public Sub ShellSortDesc(SortArray As Variant)
'The fastets sort algorithm!

Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim Swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            If LCase(SortArray(Row)) < LCase(SortArray(Row + Offset)) Then
               Swap SortArray(Row), SortArray(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
End Sub



Sub Swap(Var1, Var2)

Dim tmp As Variant
    tmp = Var1
    Var1 = Var2
    Var2 = tmp

End Sub


