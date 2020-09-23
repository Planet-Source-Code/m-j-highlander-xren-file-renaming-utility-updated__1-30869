VERSION 5.00
Begin VB.Form frmMasks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit File Masks (Filters)"
   ClientHeight    =   3570
   ClientLeft      =   2640
   ClientTop       =   1320
   ClientWidth     =   3945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Apply && Save"
      Height          =   375
      Left            =   2450
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2450
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2450
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2450
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox lstMasks 
      Height          =   3375
      Left            =   50
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub


Private Sub cmdAdd_Click()
Dim sMask As String

sMask = InputBox("Enter File Mask using wildcards * and ?, examples:" & vbCrLf _
& "*.txt" & vbCrLf _
& "*.htm;*.html" & vbCrLf _
& "c*.ht?" & vbCrLf _
, "Add New Mask", "")

If sMask <> "" Then
    lstMasks.AddItem sMask
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub Command2_Click()

End Sub


Private Sub cmdRemove_Click()
On Error Resume Next 'in case nothing is selected

lstMasks.RemoveItem lstMasks.ListIndex
lstMasks.SetFocus

End Sub


Private Sub cmdSave_Click()

Dim idx As Integer
Dim sAllMasks As String
Dim MaskArray As String
Dim sMasks As String
Dim AstDotAstFound As Boolean
Dim MasksArray As Variant

If lstMasks.ListCount = 0 Then
    MsgBox "No Entries to save", vbCritical, "Oops!"
    Exit Sub
End If



'  1- Check for duplicates: //TODO?

'  2- Combine Masks:
    MaskArray = lstMasks.List(0)
    For idx = 1 To lstMasks.ListCount - 1
        MaskArray = MaskArray & "|" & lstMasks.List(idx)
    Next idx
    ''MsgBox MaskArray

'  3- Save into Registry:
    SetSubVal HKEY_LOCAL_MACHINE, "software\XRen", "Masks", MaskArray

'***''''''''''''Fill FileList Filter Combo

    sMasks = MaskArray
    
    If sMasks <> "" Then
                AstDotAstFound = False
                MasksArray = Split(sMasks, "|")
                frmXRenMain.cboFileMask.Clear
                For idx = LBound(MasksArray) To UBound(MasksArray)
                     frmXRenMain.cboFileMask.AddItem MasksArray(idx)
                     If MasksArray(idx) = "*.*" Then AstDotAstFound = True
                Next idx
                If AstDotAstFound = False Then frmXRenMain.cboFileMask.AddItem "*.*"
                frmXRenMain.cboFileMask.ListIndex = 0
    End If
'***''''''''''''''''''''''''''''''''''''''





'  4- Close Window
    Unload Me


End Sub


Private Sub Form_Load()

Dim sMasks As String
Dim MasksArray As Variant
Dim idx As Integer

If oFlatBtns = True Then
            FlatAllBtns Me
Else

End If

'''''''''''''Fill List

    sMasks = GetSubVal(HKEY_LOCAL_MACHINE, "software\XRen", "Masks")
    
    If sMasks <> "" Then
                MasksArray = Split(sMasks, "|")
                For idx = LBound(MasksArray) To UBound(MasksArray)
                     lstMasks.AddItem MasksArray(idx)
                Next idx
    End If
'''''''''''''''''''''''''''''''''''''''

End Sub


