Attribute VB_Name = "ListBox_Scroll"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    
Public Const LB_SETHORIZONTALEXTENT = &H194

'Add Horizontal Scrollbar to List box


Public Sub AddScroll(List As Control)
    Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long
    'Find Longest Text in Listbox


    For i = 0 To List.ListCount - 1


        If Len(List.List(i)) > Len(List.List(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i
    'Get Twips
    lngGreatestWidth = List.Parent.TextWidth(List.List(intGreatestLen) + Space(1))
    'Space(1) is used to prevent the last Ch
    '     aracter from being cut off
    'Convert to Pixels
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    'Use api to add scrollbar

    SendMessage List.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
    
End Sub
