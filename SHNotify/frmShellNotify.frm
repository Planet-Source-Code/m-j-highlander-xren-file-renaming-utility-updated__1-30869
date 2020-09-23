VERSION 5.00
Begin VB.Form frmShellNotify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shell Change Notification (HIDDEN FORM)"
   ClientHeight    =   960
   ClientLeft      =   1440
   ClientTop       =   615
   ClientWidth     =   5070
   Icon            =   "frmShellNotify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrFlashMe 
      Left            =   360
      Top             =   240
   End
End
Attribute VB_Name = "frmShellNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    If SubClass(hwnd) Then
 

        Call SHNotify_Register(hwnd)

    Else

            'Uh..., it's supposed to work... :-)

    End If

    'HIDE the window!
    Me.Move -Width, -Height

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SHNotify_Unregister
    Call UnSubClass(hwnd)
    Set frmShellNotify = Nothing

End Sub

Public Sub NotificationReceipt(wParam As Long, lParam As Long)

    Dim sOut As String
    Dim shns As SHNOTIFYSTRUCT

    sOut = SHNotify_GetEventStr(lParam) & vbCrLf

    'Fill the SHNOTIFYSTRUCT from it's pointer.
    MoveMemory shns, ByVal wParam, Len(shns)

    'lParam is the ID of the notification event,
    'one of the SHCN_EventIDs.
    Select Case lParam

            '----------------------------------------------------
            'For the SHCNE_FREESPACE event, dwItem1 points
            'to what looks like a 10 byte struct. The first
            'two bytes are the size of the struct, and the
            'next two members equate to SHChangeNotify's
            'dwItem1 and dwItem2 params.

            'The dwItem1 member is a bitfield indicating which
            'drive(s) had it's (their) free space changed.
            'The bitfield is identical to the bitfield returned
            'from a GetLogicalDrives call, i.e, bit 0 = A:\, bit
            '1 = B:\, 2, = C:\, etc. Since VB does DWORD alignment
            'when MoveMemory'ing to a struct, we'll extract the
            'bitfield directly from it's memory location.

        Case SHCNE_FREESPACE

            Dim dwDriveBits As Long
            Dim wHighBit As Integer
            Dim wBit As Integer

            MoveMemory dwDriveBits, ByVal shns.dwItem1 + 2, 4

            'Get the zero based position of the highest
            'bit set in the bitmask (essentially determining
            'the value's highest complete power of 2).
            'Use floating point division (we want the exact
            'values from the Logs) and remove the fractional
            'value (the fraction indicates the value of
            'the last incomplete power of 2, which means the
            'bit isn't set).

            wHighBit = Int(Log(dwDriveBits) / Log(2))

            For wBit = 0 To wHighBit

                'If the bit is set...
                If (2 ^ wBit) And dwDriveBits Then

                    '... get it's drive string
                    sOut = sOut & Chr$(vbKeyA + wBit) & ":\" & vbCrLf

                End If
            Next

            '----------------------------------------------------
            'shns.dwItem1 also points to a 10 byte struct. The
            'struct's second member (after the struct's first
            'WORD size member) points to the system imagelist
            'index of the image that was updated.
        Case SHCNE_UPDATEIMAGE

            Dim iImage As Long

            MoveMemory iImage, ByVal shns.dwItem1 + 2, 4
            sOut = sOut & "Index of image in system imagelist: " & iImage & vbCrLf

            '----------------------------------------------------
            'Everything else except SHCNE_ATTRIBUTES is the
            'pidl(s) of the changed item(s). For SHCNE_ATTRIBUTES,
            'neither item is used. See the description of the
            'values for the wEventId parameter of the
            'SHChangeNotify API function for more info.
        Case Else
            Dim sDisplayname As String

            If shns.dwItem1 Then

                sDisplayname = GetDisplayNameFromPIDL(shns.dwItem1)

                If Len(sDisplayname) Then
                    sOut = sOut & "first item displayname: " & sDisplayname & vbCrLf
                    sOut = sOut & "first item path: " & GetPathFromPIDL(shns.dwItem1) & vbCrLf
                Else
                    sOut = sOut & "first item is invalid" & vbCrLf
                End If

            End If

            If shns.dwItem2 Then

                sDisplayname = GetDisplayNameFromPIDL(shns.dwItem2)

                If Len(sDisplayname) Then
                    sOut = sOut & "second item displayname: " & sDisplayname & vbCrLf
                    sOut = sOut & "second item path: " & GetPathFromPIDL(shns.dwItem2) & vbCrLf
                Else
                    sOut = sOut & "second item is invalid" & vbCrLf
                End If
            End If

    End Select

    'update the text window and flash
    'the window title
    'Text1 = Text1 & sOut & vbCrLf
    'Text1.SelStart = Len(Text1)
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    frmXRenMain.File1.Refresh
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    tmrFlashMe = True

End Sub

Private Sub tmrFlashMe_Timer()

    'initial settings: Interval = 1, Enabled = False

    Static nCount As Integer

    If nCount = 0 Then tmrFlashMe.Interval = 200

    nCount = nCount + 1
    Call FlashWindow(hwnd, True)

    'Reset everything after 3 flash cycles
    If nCount = 6 Then
        nCount = 0
        tmrFlashMe.Interval = 1
        tmrFlashMe = False
    End If

End Sub

