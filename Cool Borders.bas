Attribute VB_Name = "CoolBorders"
Option Explicit

'Getting the Office Look
'
'Fancy giving your application that Office 2000 feel?
'
'Get with it by giving all your controls a groovy new Office 2000-like border style, all thanks to this neat copy-and-paste code snippet.
'
'This code uses the API to alter the border of a control, giving it a very faint and rather cool-looking 3D border.
'
'For best results, change the Appearance and BorderStyle properties of your control(s) to Flat and None, respectively.
'
'To run this code, simply call the AddOfficeBorder method, passing it the hWnd property of your control.
'



Private Declare Function GetWindowLong Lib "user32" Alias _
        "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
        "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4


Sub AddBorderToAllTextBoxes(frmX As Form)

Dim X As Control

On Error Resume Next
For Each X In frmX.Controls
        If TypeOf X Is TextBox Then
                X.Appearance = vbFlat
                X.BorderStyle = 0
                AddOfficeBorder X
        End If
Next

End Sub


Public Function AddOfficeBorder(ctlX As Control)
    
    Dim lngRetVal As Long
    
    'Retrieve the current border style
    lngRetVal = GetWindowLong(ctlX.hWnd, GWL_EXSTYLE)
    
    'Calculate border style to use
    lngRetVal = lngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    
    'Apply the changes
    SetWindowLong ctlX.hWnd, GWL_EXSTYLE, lngRetVal
    SetWindowPos ctlX.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
End Function

