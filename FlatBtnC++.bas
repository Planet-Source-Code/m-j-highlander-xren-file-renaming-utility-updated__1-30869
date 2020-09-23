Attribute VB_Name = "Flat_Btns"
'Flat Buttons

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'The button style BS_FLAT used to change a button to a Flat one
Public Const BS_FLAT = &H8000&
'GWL_Style is the attribute we will use for changing the style of the button
Public Const GWL_STYLE = (-16)
'To set the button as a child window and not as a self dependent window
Public Const WS_CHILD = &H40000000


Sub UnFlatAllBtns(frmX As Form)

Dim btnX As Control
For Each btnX In frmX.Controls
    If Left(btnX.Name, 3) = "cmd" Then
            UnbtnFlat btnX
    End If
    
Next btnX

End Sub
Public Function btnFlat(cmdFlat As CommandButton)
    SetWindowLong cmdFlat.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    cmdFlat.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function

Public Function UnbtnFlat(cmdFlat As CommandButton)
    SetWindowLong cmdFlat.hwnd, GWL_STYLE, WS_CHILD
    cmdFlat.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function

Sub FlatAllBtns(frmX As Form)

Dim btnX As Control
For Each btnX In frmX.Controls
    If Left(btnX.Name, 3) = "cmd" Then
            btnFlat btnX
    End If
    
Next btnX

End Sub


