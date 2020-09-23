VERSION 5.00
Begin VB.Form frmErrorLog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Error Log"
   ClientHeight    =   3525
   ClientLeft      =   1380
   ClientTop       =   1620
   ClientWidth     =   5835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtErrorLog 
      Height          =   3030
      Left            =   -15
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   30
      Width           =   5700
   End
End
Attribute VB_Name = "frmErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function ShowErrorLog()

If sErrors <> "" Then
        Me.txtErrorLog.Text = sErrors        'public Var
        Me.Show vbModal
        sErrors = ""
Else
        'NO Errors
End If


End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case 27, 13 'Esc , Enter
        Unload Me
End Select

End Sub

Private Sub Form_Load()
txtErrorLog.Left = 0
txtErrorLog.Top = 0
txtErrorLog.Width = Me.ScaleWidth
txtErrorLog.Height = Me.ScaleHeight

End Sub



