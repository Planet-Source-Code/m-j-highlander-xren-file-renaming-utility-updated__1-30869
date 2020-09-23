VERSION 5.00
Begin VB.Form frmAboutXRen 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1815
   ClientLeft      =   1410
   ClientTop       =   1545
   ClientWidth     =   5355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "XRen 3.2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   30
      TabIndex        =   1
      Top             =   1425
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mdsy 2001"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   4425
      TabIndex        =   0
      Top             =   1545
      Width           =   780
   End
   Begin VB.Image Image1 
      Height          =   1380
      Left            =   45
      Picture         =   "frmCoolAbout.frx":0000
      Top             =   30
      Width           =   5250
   End
End
Attribute VB_Name = "frmAboutXRen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me


End Sub



Private Sub Form_Click()
Unload Me
End Sub

Private Sub Image1_Click()
Unload Me
End Sub


Private Sub Label1_Click()
Unload Me
End Sub


Private Sub Label2_Click()
Unload Me
End Sub


