VERSION 5.00
Begin VB.Form frmManualRename 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rename"
   ClientHeight    =   1410
   ClientLeft      =   1650
   ClientTop       =   1935
   ClientWidth     =   6480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAbort 
      Caption         =   "&Abort"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4418
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   855
      Width           =   1620
   End
   Begin VB.CommandButton cmdSkip 
      Cancel          =   -1  'True
      Caption         =   "&Skip"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   855
      Width           =   1620
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   443
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   870
      Width           =   1620
   End
   Begin VB.TextBox txtNewName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   375
      Width           =   6255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Type new filename"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1335
   End
End
Attribute VB_Name = "frmManualRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private bAbortFlag As Boolean
Private sOldFileName As String
Private Sub cmdAbort_Click()

Me.Aborted() = True
Me.Hide

End Sub


Private Sub cmdOk_Click()

Dim sNewFullName As String
Dim sOldDir As String

If Trim(txtNewName.Text) = "" Then
        Beep
        Exit Sub
End If

sOldDir = ExtractDirName(Me.OldFileName)
sNewFullName = sOldDir & Trim(txtNewName.Text)

If Me.OldFileName <> sNewFullName Then
            SHRename Me.OldFileName, sNewFullName
End If

Me.Hide
End Sub

Private Sub cmdSkip_Click()
txtNewName.Text = ""
Me.Hide
End Sub


Property Get Aborted() As Variant
Aborted = bAbortFlag
End Property

Property Let Aborted(ByVal vNewValue As Variant)
bAbortFlag = vNewValue
End Property

Private Sub Form_Activate()
txtNewName.Text = ExtractFileName(Me.OldFileName)
SelectAllText txtNewName
txtNewName.SetFocus

End Sub

Private Sub Form_Load()
Dim X As Control
On Error Resume Next
        For Each X In Me.Controls
                X.Font.Name = "tahoma"
                X.Font.Size = 8
        Next X
On Error GoTo 0

Me.Aborted = False

End Sub



Public Property Get OldFileName() As Variant
OldFileName = sOldFileName
End Property

Public Property Let OldFileName(ByVal vNewValue As Variant)
sOldFileName = vNewValue
End Property

Private Sub txtNewName_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
    Case "\", "/", ":", "*", "?", "<", ">", "|", Chr$(34)
        'illegal chars
        Beep
        KeyAscii = 0
    Case Else
    'do nothing
End Select

End Sub


