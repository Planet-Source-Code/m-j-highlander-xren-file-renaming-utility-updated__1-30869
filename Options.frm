VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XRen Options"
   ClientHeight    =   5055
   ClientLeft      =   1725
   ClientTop       =   900
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   " Shell Extension "
      Height          =   1545
      Left            =   135
      TabIndex        =   10
      Top             =   1935
      Width           =   5685
      Begin VB.CommandButton cmdUnRegister 
         Caption         =   "Unregister"
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1035
         Width           =   1425
      End
      Begin VB.CommandButton cmdRegShellExt 
         Caption         =   "Register XRen as a Shell Extension"
         Height          =   375
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1035
         Width           =   3630
      End
      Begin VB.Label lblShell 
         Caption         =   "xxx"
         Height          =   555
         Left            =   135
         TabIndex        =   12
         Top             =   315
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdEditMask 
      Caption         =   "Edit File Mask List ..."
      Height          =   375
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3690
      Width           =   2385
   End
   Begin VB.Frame Frame1 
      Caption         =   " Undo File "
      Height          =   720
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5700
      Begin VB.CheckBox chkUndoBatch 
         Caption         =   "Create Undo Batch File"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Create 'UNREN.BAT' which can be used to restore original filenames"
         Top             =   360
         Value           =   1  'Checked
         Width           =   2100
      End
      Begin VB.CheckBox chkUndoVBScript 
         Caption         =   "Create Undo VBScript File"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2445
         TabIndex        =   7
         ToolTipText     =   "Create 'UNREN.VBS' which can be used to restore original filenames"
         Top             =   360
         Width           =   2340
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   750
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4530
      Width           =   1515
   End
   Begin VB.Frame frmSel 
      Caption         =   " Selection Style "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      TabIndex        =   3
      Top             =   990
      Width           =   5685
      Begin VB.OptionButton optSelExtended 
         Caption         =   "Extended: Use CTRL and/or SHIFT to select multiple files"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   585
         Width           =   4575
      End
      Begin VB.OptionButton optSelSimple 
         Caption         =   "Simple: Click every file you want to select"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   285
         Value           =   -1  'True
         Width           =   3900
      End
   End
   Begin VB.CheckBox chkFlatButtons 
      Caption         =   "Use Flat Buttons"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   540
      TabIndex        =   2
      Top             =   3735
      Value           =   1  'Checked
      Width           =   2070
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3465
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4545
      Width           =   1515
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   90
      X2              =   5735
      Y1              =   4365
      Y2              =   4365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   90
      X2              =   5735
      Y1              =   4350
      Y2              =   4350
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkUndoBatch_Click()

If chkUndoBatch.Value = vbChecked Then
            chkUndoVBScript.Value = vbUnchecked
End If

End Sub


Private Sub chkUndoVBScript_Click()

If chkUndoVBScript.Value = vbChecked Then
            chkUndoBatch.Value = vbUnchecked
End If

End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub



Private Sub cmdEditMask_Click()

frmMasks.Show vbModal

End Sub

Private Sub cmdFixVBS_Click()

Dim Msg As String
Dim sVal As String
Dim ap As String
Dim msgReturn As Integer

Msg = "XRen can generate Undo VBScript files "".VBS"" " _
      & "which can be executed by double-clicking them. " & vbCrLf _
      & "On some system the "".VBS"" extension might be associated with an " _
      & "MPEG Video Program causing Windows not to be able to run the script." & vbCrLf _
      & "(To run VBScript files you need Windows 98 or newer or Windows 95 with Windows Scripting Host installed.)" _
      & vbCrLf & vbCrLf & "Click OK to register "".VBS"" files as Windows Scripting Files"
msgReturn = MsgBox(Msg, vbOKCancel + vbInformation, "XRen")

If msgReturn = vbOK Then
        'Create Registry Entry
         SetSubVal HKEY_CLASSES_ROOT, ".vbs", "", "VBSFile"
End If

End Sub

Private Sub cmdOk_Click()
'Create Undo VBScript File
If chkUndoVBScript.Value = vbChecked Then
            oUndoVBScript = True
Else
            oUndoVBScript = False
End If
'---------------------------------------------------

'Create Undo Batch File
If chkUndoBatch.Value = vbChecked Then
            oUndoBatch = True
Else
            oUndoBatch = False
End If
'---------------------------------------------------
'Use Flat Btns
If chkFlatButtons.Value = vbChecked Then
            oFlatBtns = True
            FlatAllBtns frmXRenMain
Else
            oFlatBtns = False
            UnFlatAllBtns frmXRenMain
            btnFlat frmXRenMain.cmdStretch

End If
'---------------------------------------------------
'Select Style:
If optSelExtended.Value = True Then
            
            frmXRenMain.File1.MultiSelect = xfSelectExtended
Else
           
            frmXRenMain.File1.MultiSelect = xfSelectSimple
End If
'---------------------------------------------------


Unload Me

End Sub


Private Sub cmdRegShellExt_Click()

Dim Msg As String
Dim sVal As String
Dim ap As String

'================================================================

ap = App.Path
If Right(ap, 1) <> "\" Then ap = ap & "\"

'Register DLL
Shell "RegSvr32.exe" & " " & q & ap & "XRenShx.dll" & q & " /s"

'Create Registry Entry
SetSubVal HKEY_CLASSES_ROOT, "*\shellex\ContextMenuHandlers\XRenContext", "", "{23FCFE69-A54B-11D4-8AD0-484C000107C0}"
'MsgBox "XRen Shell Extension installation succeeded." & vbCrLf & "Now you can right-click files in Explorer and select XRen", vbInformation, "XRen"

SetSubVal HKEY_LOCAL_MACHINE, "software\XRen", "AutoRegisterShellExt", "1"
End Sub

Private Sub cmdUnRegister_Click()
Dim Msg As String
Dim sVal As String
Dim ap As String

'================================================================

ap = App.Path
If Right(ap, 1) <> "\" Then ap = ap & "\"

'UnRegister DLL
Shell "RegSvr32.exe" & " " & q & ap & "XRenShx.dll" & q & " /u /s"

'Delete Registry Entry
DeleteKey HKEY_CLASSES_ROOT, "*\shellex\ContextMenuHandlers\XRenContext"
SetSubVal HKEY_LOCAL_MACHINE, "software\XRen", "AutoRegisterShellExt", "0"

End Sub


Private Sub Form_Load()

lblShell.Caption = "Registering XRen as a Shell Extension adds an ""XRen"" menu item " _
      & "to the context menu in Windows Explorer."

Dim X As Control
On Error Resume Next
        For Each X In Me.Controls
                X.Font.Name = "tahoma"
                X.Font.Size = 8
        Next X
On Error GoTo 0


'Create Undo Batch File
If oUndoBatch = True Then
            chkUndoBatch.Value = vbChecked
Else
            chkUndoBatch.Value = vbUnchecked
End If
'---------------------------------------------------
'Create Undo VBScript File
If oUndoVBScript = True Then
            chkUndoVBScript.Value = vbChecked
Else
            chkUndoVBScript.Value = vbUnchecked
End If
'---------------------------------------------------
'Use Flat Btns
If oFlatBtns = True Then
            chkFlatButtons.Value = vbChecked
            FlatAllBtns Me
Else
           chkFlatButtons.Value = vbUnchecked

End If
'---------------------------------------------------
'Select Style:
If frmXRenMain.File1.MultiSelect = xfSelectExtended Then
            
            optSelExtended.Value = True
Else
            optSelExtended.Value = False
End If
'---------------------------------------------------

End Sub


