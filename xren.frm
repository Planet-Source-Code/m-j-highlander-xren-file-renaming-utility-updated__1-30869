VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmXRenMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XRen 3.3"
   ClientHeight    =   5610
   ClientLeft      =   330
   ClientTop       =   870
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "xren.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5610
   ScaleWidth      =   9165
   Begin VB.CommandButton cmdStretch 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "Stretch/Shrink FileList (Ctrl+> / Ctrl+<)"
      Top             =   4870
      Width           =   250
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   4840
      Begin VB.ComboBox cboFileMask 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "*.*"
         ToolTipText     =   "click ""Options | Edit File Mask List""  to add or remove masks (filters)"
         Top             =   240
         Width           =   2350
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2400
         ScaleHeight     =   195
         ScaleWidth      =   2295
         TabIndex        =   78
         Top             =   4800
         Width           =   2295
         Begin VB.Label lblSelFiles 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Selected=0 Files"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   79
            ToolTipText     =   "number of selected files / total number of displayed files"
            Top             =   25
            Width           =   1185
         End
      End
      Begin XRen.XFileListBox File1 
         Height          =   4245
         Left            =   2400
         TabIndex        =   5
         Top             =   630
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   7488
      End
      Begin VB.CommandButton cmdSelectNone 
         Appearance      =   0  'Flat
         Caption         =   "Select None"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Ctrl + N"
         Top             =   4680
         Width           =   1095
      End
      Begin VB.TextBox txtPath 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   90
         TabIndex        =   1
         ToolTipText     =   "Type a path and it will be selected in the Directory listbox, type DESKTOP to select the Windows Desktop folder"
         Top             =   225
         Width           =   2190
      End
      Begin VB.CommandButton cmdSelectAll 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ctrl + A"
         Top             =   4680
         Width           =   1080
      End
      Begin VB.CommandButton cmdInvertSelection 
         Appearance      =   0  'Flat
         Caption         =   "Invert Selection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ctrl + I"
         Top             =   4320
         Width           =   2190
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3240
         Left            =   90
         TabIndex        =   3
         ToolTipText     =   "Right-click to Open the selected folder or to Refresh"
         Top             =   990
         Width           =   2205
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   615
         Width           =   2200
      End
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7125
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   4575
      Width           =   1770
   End
   Begin VB.CommandButton cmdOptions 
      Appearance      =   0  'Flat
      Caption         =   "Options..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Ctrl + O"
      Top             =   4575
      Width           =   1770
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   1740
      TabIndex        =   32
      Top             =   5280
      Visible         =   0   'False
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   180
      Picture         =   "xren.frx":1042
      ScaleHeight     =   450
      ScaleWidth      =   1200
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5175
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4110
      Left            =   4905
      TabIndex        =   9
      Top             =   210
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   7250
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Serial Rename"
      TabPicture(0)   =   "xren.frx":1DE4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblExt"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtNum"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDigit"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtExt"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdSerial"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "UpDown1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkMaintainExt"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Change Case"
      TabPicture(1)   =   "xren.frx":1E00
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "cmdChangeCase"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Add / Replace"
      TabPicture(2)   =   "xren.frx":1E1C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblBefore"
      Tab(2).Control(1)=   "lblAfter"
      Tab(2).Control(2)=   "Line1"
      Tab(2).Control(3)=   "Line2"
      Tab(2).Control(4)=   "Label5"
      Tab(2).Control(5)=   "Label6"
      Tab(2).Control(6)=   "txtBefore"
      Tab(2).Control(7)=   "txtAfter"
      Tab(2).Control(8)=   "cmdAdd"
      Tab(2).Control(9)=   "cmdReplace"
      Tab(2).Control(10)=   "txtWith"
      Tab(2).Control(11)=   "txtReplace"
      Tab(2).Control(12)=   "optReplaceChars"
      Tab(2).Control(13)=   "optReplaceString"
      Tab(2).Control(14)=   "chkMatchCase"
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "Delete"
      TabPicture(3)   =   "xren.frx":1E38
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label9"
      Tab(3).Control(1)=   "Label10"
      Tab(3).Control(2)=   "Label11"
      Tab(3).Control(3)=   "Label12"
      Tab(3).Control(4)=   "Line3"
      Tab(3).Control(5)=   "Line4"
      Tab(3).Control(6)=   "upDelFirst"
      Tab(3).Control(7)=   "txtDelFirst"
      Tab(3).Control(8)=   "optDelEnd"
      Tab(3).Control(9)=   "optDelStart"
      Tab(3).Control(10)=   "txtDelTo"
      Tab(3).Control(11)=   "txtDelLast"
      Tab(3).Control(12)=   "upDelLast"
      Tab(3).Control(13)=   "cmdDelChars"
      Tab(3).Control(14)=   "cmdDelTo"
      Tab(3).Control(15)=   "chkInclusive"
      Tab(3).Control(16)=   "chkDelMatchCase"
      Tab(3).ControlCount=   17
      TabCaption(4)   =   "Text / HTML / Change Ext"
      TabPicture(4)   =   "xren.frx":1E54
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label13"
      Tab(4).Control(1)=   "Line5"
      Tab(4).Control(2)=   "Line6"
      Tab(4).Control(3)=   "Label14"
      Tab(4).Control(4)=   "cmdHTMLFilesRename"
      Tab(4).Control(5)=   "cmdTextFilesRename"
      Tab(4).Control(6)=   "txtNewExt"
      Tab(4).Control(7)=   "cmdChangeExt"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "MP3"
      TabPicture(5)   =   "xren.frx":1E70
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label8"
      Tab(5).Control(1)=   "Label7"
      Tab(5).Control(2)=   "cmdMP3"
      Tab(5).Control(3)=   "txtSeperator"
      Tab(5).Control(4)=   "cboTAG(2)"
      Tab(5).Control(5)=   "cboTAG(1)"
      Tab(5).Control(6)=   "cboTAG(0)"
      Tab(5).ControlCount=   7
      Begin VB.CheckBox chkDelMatchCase 
         Caption         =   "Match Case"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -72780
         TabIndex        =   77
         Top             =   2670
         Width           =   1215
      End
      Begin VB.CheckBox chkInclusive 
         Caption         =   "Inclusive"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -72780
         TabIndex        =   76
         Top             =   2415
         Width           =   1215
      End
      Begin VB.CommandButton cmdChangeExt 
         Appearance      =   0  'Flat
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73695
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   3480
         Width           =   1605
      End
      Begin VB.TextBox txtNewExt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   -73770
         TabIndex        =   72
         ToolTipText     =   "Leave empty to remove extension"
         Top             =   2940
         Width           =   1680
      End
      Begin VB.CommandButton cmdDelTo 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -73605
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   3480
         Width           =   1620
      End
      Begin VB.CommandButton cmdDelChars 
         Appearance      =   0  'Flat
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73650
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1665
         Width           =   1605
      End
      Begin MSComCtl2.UpDown upDelLast 
         Height          =   330
         Left            =   -72495
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   1185
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         OrigLeft        =   3270
         OrigTop         =   2250
         OrigRight       =   3510
         OrigBottom      =   2580
         Max             =   255
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtDelLast 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   -73500
         TabIndex        =   66
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtDelTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   -74370
         TabIndex        =   62
         Top             =   2985
         Width           =   2805
      End
      Begin VB.OptionButton optDelStart 
         Caption         =   "Delete from Start to"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74820
         TabIndex        =   64
         Top             =   2400
         Width           =   2145
      End
      Begin VB.OptionButton optDelEnd 
         Caption         =   "Delete from End to"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74805
         TabIndex        =   63
         Top             =   2700
         Width           =   2655
      End
      Begin VB.TextBox txtDelFirst 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   -73515
         TabIndex        =   60
         Top             =   780
         Width           =   915
      End
      Begin VB.CommandButton cmdTextFilesRename 
         Caption         =   "Rename using first line of text file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -74565
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Uses the first line of the text file for filename, ignores illegal chars and empty lines"
         Top             =   975
         Width           =   3000
      End
      Begin VB.CommandButton cmdHTMLFilesRename 
         Caption         =   "Rename HTML files using <TITLE> tag"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -74565
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Uses the <TITLE>...</TITLE> Tag for filename, ignores illegal charachters"
         Top             =   1650
         Width           =   3000
      End
      Begin VB.ComboBox cboTAG 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -74775
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   1560
         Width           =   1035
      End
      Begin VB.ComboBox cboTAG 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   1560
         Width           =   1065
      End
      Begin VB.ComboBox cboTAG 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -72285
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   1560
         Width           =   1155
      End
      Begin VB.TextBox txtSeperator 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   -73755
         TabIndex        =   51
         Text            =   " - "
         Top             =   2175
         Width           =   795
      End
      Begin VB.CommandButton cmdMP3 
         Appearance      =   0  'Flat
         Caption         =   "Rename"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73725
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3165
         Width           =   1605
      End
      Begin VB.CheckBox chkMatchCase 
         Caption         =   "Match Case"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -72270
         TabIndex        =   47
         Top             =   2955
         Width           =   1215
      End
      Begin VB.OptionButton optReplaceString 
         Caption         =   "Replace String"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -72675
         TabIndex        =   46
         Top             =   2700
         Width           =   1605
      End
      Begin VB.OptionButton optReplaceChars 
         Caption         =   "Replace Char's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -72675
         TabIndex        =   45
         Top             =   2355
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.TextBox txtReplace 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   -74235
         TabIndex        =   42
         Top             =   2280
         Width           =   1380
      End
      Begin VB.TextBox txtWith 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   -74235
         TabIndex        =   41
         ToolTipText     =   "One charachter only"
         Top             =   2685
         Width           =   1380
      End
      Begin VB.CommandButton cmdReplace 
         Appearance      =   0  'Flat
         Caption         =   "Replace"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -74745
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -74865
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1620
         Width           =   1335
      End
      Begin VB.TextBox txtAfter 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   -73320
         TabIndex        =   36
         Top             =   1170
         Width           =   1800
      End
      Begin VB.TextBox txtBefore 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   -73335
         TabIndex        =   34
         Top             =   765
         Width           =   1815
      End
      Begin VB.CheckBox chkMaintainExt 
         Caption         =   "Maintain Original Extension"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1260
         TabIndex        =   14
         ToolTipText     =   "Keeps the original extension unchanged, this can also be done by typing * in the Extension TextBox"
         Top             =   1815
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   3090
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2505
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1
         OrigLeft        =   3270
         OrigTop         =   2250
         OrigRight       =   3510
         OrigBottom      =   2580
         Max             =   9
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdChangeCase 
         Appearance      =   0  'Flat
         Caption         =   "Change Case"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73935
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3465
         Width           =   1770
      End
      Begin VB.CommandButton cmdSerial 
         Appearance      =   0  'Flat
         Caption         =   "Rename"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1005
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3165
         Width           =   1770
      End
      Begin VB.Frame Frame3 
         Caption         =   " Extension "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   -74910
         TabIndex        =   26
         Top             =   2160
         Width           =   3435
         Begin VB.OptionButton optExtUpFirst 
            Caption         =   "First letter upper case"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   315
            TabIndex        =   29
            Top             =   750
            Width           =   2700
         End
         Begin VB.OptionButton optExtLow 
            Caption         =   "lower case"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   315
            TabIndex        =   28
            Top             =   465
            Value           =   -1  'True
            Width           =   2700
         End
         Begin VB.OptionButton optExtUp 
            Caption         =   "UPPER CASE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   315
            TabIndex        =   27
            Top             =   195
            Width           =   2700
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Name "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   -74895
         TabIndex        =   21
         Top             =   630
         Width           =   3420
         Begin VB.OptionButton optUpEach 
            Caption         =   "Each First Letter Upper Case"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   315
            TabIndex        =   25
            Top             =   1125
            Value           =   -1  'True
            Width           =   2835
         End
         Begin VB.OptionButton optUpFirst 
            Caption         =   "First letter upper case"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   315
            TabIndex        =   24
            Top             =   825
            Width           =   2700
         End
         Begin VB.OptionButton optLower 
            Caption         =   "lower case"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   315
            TabIndex        =   23
            Top             =   525
            Width           =   2700
         End
         Begin VB.OptionButton optUpper 
            Caption         =   "UPPER CASE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   315
            TabIndex        =   22
            Top             =   240
            Width           =   2700
         End
      End
      Begin VB.TextBox txtExt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1020
         TabIndex        =   13
         ToolTipText     =   "Leave empty for no extension, type * to keep original extension"
         Top             =   1410
         Width           =   2505
      End
      Begin VB.TextBox txtDigit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2115
         MaxLength       =   1
         TabIndex        =   18
         Text            =   "1"
         ToolTipText     =   "Valid numbers: 1 to 9"
         Top             =   2505
         Width           =   915
      End
      Begin VB.TextBox txtNum 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2115
         MaxLength       =   9
         TabIndex        =   16
         Text            =   "1"
         Top             =   2145
         Width           =   1200
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1005
         MultiLine       =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Leave empty for a number-only filename"
         Top             =   720
         Width           =   2520
      End
      Begin MSComCtl2.UpDown upDelFirst 
         Height          =   330
         Left            =   -72510
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   750
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         OrigLeft        =   3270
         OrigTop         =   2250
         OrigRight       =   3510
         OrigBottom      =   2580
         Max             =   255
         Enabled         =   -1  'True
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Ext."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74535
         TabIndex        =   74
         Top             =   2985
         Width           =   660
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000005&
         X1              =   -71485
         X2              =   -73410
         Y1              =   2625
         Y2              =   2625
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         X1              =   -71485
         X2              =   -73410
         Y1              =   2610
         Y2              =   2610
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Extension"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74805
         TabIndex        =   73
         Top             =   2505
         Width           =   1305
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000003&
         X1              =   -71100
         X2              =   -74910
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         X1              =   -71100
         X2              =   -74910
         Y1              =   2175
         Y2              =   2175
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Last"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74670
         TabIndex        =   69
         Top             =   1260
         Width           =   810
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Chars"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -72060
         TabIndex        =   68
         Top             =   1260
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete First"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74685
         TabIndex        =   65
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Chars"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -72075
         TabIndex        =   61
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rename files using ID3 Tag in this order:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74775
         TabIndex        =   56
         Top             =   1080
         Width           =   2910
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Seperator"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74685
         TabIndex        =   55
         Top             =   2220
         Width           =   720
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Replace"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74880
         TabIndex        =   44
         Top             =   2295
         Width           =   570
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "With"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74715
         TabIndex        =   43
         Top             =   2730
         Width           =   330
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   -71100
         X2              =   -74910
         Y1              =   2175
         Y2              =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   -71100
         X2              =   -74910
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label lblAfter 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Add After Filename"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74865
         TabIndex        =   37
         Top             =   1230
         Width           =   1380
      End
      Begin VB.Label lblBefore 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Add Before Filename"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74880
         TabIndex        =   35
         Top             =   825
         Width           =   1485
      End
      Begin VB.Label lblExt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Extension"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   1470
         Width           =   705
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Digits"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   2580
         Width           =   1185
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Numbering From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   2190
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Base Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   750
         Width           =   795
      End
   End
   Begin VB.Label lblMailTo 
      AutoSize        =   -1  'True
      Caption         =   "mdsy@excite.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6555
      MouseIcon       =   "xren.frx":1E8C
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Top             =   5280
      Width           =   1305
   End
   Begin VB.Label lblButtomLine 
      AutoSize        =   -1  'True
      Caption         =   "XRen the filename changer utility.  For more info e-mail:  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2370
      TabIndex        =   33
      Top             =   5295
      Width           =   4095
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "mnuHidden"
      Visible         =   0   'False
      Begin VB.Menu mnuManualRename 
         Caption         =   "Rename Selected Files"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Selected Folder"
      End
      Begin VB.Menu mnuHyph 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortBy 
         Caption         =   "Sort By"
         Begin VB.Menu mnuSortName 
            Caption         =   "Name"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSortExtension 
            Caption         =   "Extension"
         End
         Begin VB.Menu mnuSortSize 
            Caption         =   "Size"
         End
         Begin VB.Menu mnuSortDate 
            Caption         =   "Date"
         End
      End
      Begin VB.Menu mnuSortOrder 
         Caption         =   "Sort Order"
         Begin VB.Menu mnuSortOrderAscending 
            Caption         =   "Ascending"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSortOrderDescending 
            Caption         =   "Descending"
         End
      End
      Begin VB.Menu hyphz 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmXRenMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ap As String
Dim fcolor









Private Sub cboFileMask_Change()
'On Error Resume Next

File1.Pattern = cboFileMask.Text
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

End Sub

Private Sub cboFileMask_Click()
'On Error Resume Next

File1.Pattern = cboFileMask.Text
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount


End Sub


Private Sub chkMaintainExt_Click()
If chkMaintainExt.Value = 1 Then
    lblExt.Enabled = False
    txtExt.Enabled = False
    txtExt.Text = ""
Else
    lblExt.Enabled = True
    txtExt.Enabled = True
End If
End Sub



Private Sub cmdAbout_Click()
       


End Sub

Private Sub cmdAdd_Click()
Dim sBefore As String, sAfter As String
Dim FName As String, fExt As String, frmt As String
Dim OriginalExt As String, srcFile As String, tgtFile As String
Dim result As Integer, fStart As Integer, Cntr As Integer, i As Integer

sBefore = txtBefore.Text
sAfter = txtAfter.Text
If (sBefore = "" And sAfter = "") Then
            Beep
            Exit Sub
End If

UndoStr = ""
vbsUndoStr = ""   ' Publuic Var
sErrors = ""
ap = Dir1.Path
If Len(ap) > 3 Then ap = ap + "\"

'^^^ Progress Bar code
Dim SelCount As Integer
Dim CurrFile As Integer
ProgressBar1.Visible = True
SelCount = File1.SelCount

If SelCount = 0 Then
                Beep
                Exit Sub
End If


For i = 0 To File1.ListCount - 1
If File1.Selected(i) = True Then
        srcFile = ap + File1.List(i)
        fExt = ExtractFileExtension(File1.List(i))
        FName = SplitFileNameName(File1.List(i)) 'Left(File1.List(i), Len(File1.List(i)) - Len(fExt) - 1)
        tgtFile = ap & sBefore & FName & sAfter & "." & fExt
        

    XRename srcFile, tgtFile

    CurrFile = CurrFile + 1
    ProgressBar1.Value = Int(CurrFile / SelCount * 100)
    

'        sErrors = sErrors & vbCrLf & "FileName: " & FName & "   Error:" & Error
End If
Next i


If oUndoBatch = True And UndoStr <> "" Then
        CreateUNRENBAT ap
End If
If oUndoVBScript = True And UndoStr <> "" Then
        CreateUNRENVBS ap
End If

File1.Refresh
ProgressBar1.Visible = False
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

If sErrors <> "" Then
    frmErrorLog.ShowErrorLog
End If


End Sub

Private Sub cmdChangeCase_Click()
Dim SelCount As Integer
Dim CurrFile As Integer
Dim srcFile As String
Dim ap As String
Dim TheName As String
Dim TheExt As String
Dim FName As String
Dim i As Integer
Dim tgtFile As String

sErrors = ""          ' Publuic Var
UndoStr = ""        ' Publuic Var
'vbsUndoStr = ""   'Not used here !!!

ap = Dir1.Path
If Len(ap) > 3 Then ap = ap + "\"


SelCount = File1.SelCount
If SelCount = 0 Then
            Beep
            Exit Sub
End If

ProgressBar1.Visible = True

For i = 0 To File1.ListCount - 1
    If File1.Selected(i) = True Then
                FName = File1.List(i)
                srcFile = ap + FName
                TheName = SplitFileNameName(FName)
                TheExt = SplitFileNameExt(FName)
                
                If optUpper Then TheName = UCase$(TheName)
                If optLower Then TheName = LCase$(TheName)
                If optUpFirst Then TheName = UpFirst(TheName)
                If optUpEach Then TheName = UpEachFirst(TheName)
                
                If optExtUp Then TheExt = UCase$(TheExt)
                If optExtLow Then TheExt = LCase$(TheExt)
                If optExtUpFirst Then TheExt = UpFirst(TheExt)
                tgtFile = ap + TheName + "." + TheExt
        
                XRename srcFile, tgtFile
                
                CurrFile = CurrFile + 1
                ProgressBar1.Value = Int(CurrFile / SelCount * 100)
        
    End If
    
Next i

If oUndoBatch = True And UndoStr <> "" Then
        CreateUNRENBAT ap
End If
If oUndoVBScript = True And UndoStr <> "" Then
        CreateUNRENBAT ap               ' VBScripts couldn't do it
End If

File1.Refresh
ProgressBar1.Visible = False
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

If sErrors <> "" Then
    frmErrorLog.ShowErrorLog
End If

End Sub

Private Sub cmdChangeExt_Click()
Dim CurrFile As Integer
Dim NewExt As String, TheName As String, TheExt As String
Dim FName As String, fExt As String, frmt As String
Dim OriginalExt As String, srcFile As String, tgtFile As String
Dim result As Integer, fStart As Integer, Cntr As Integer, i As Integer
Dim SelCount As Integer

SelCount = File1.SelCount
If SelCount = 0 Then
                Beep
                Exit Sub
End If

NewExt = Trim(txtNewExt.Text)
If Left(NewExt, 1) = "." Then
            NewExt = Right(NewExt, Len(NewExt) - 1)
End If
If NewExt = "." Then
            NewExt = ""
End If

sErrors = ""
UndoStr = ""
vbsUndoStr = ""   ' Publuic Var

ap = Dir1.Path
If Len(ap) > 3 Then ap = ap + "\"

ProgressBar1.Visible = True

For i = 0 To File1.ListCount - 1
    If File1.Selected(i) = True Then
                    FName = ap + File1.List(i)
                    TheName = SplitFileNameName(FName)
                    TheExt = ExtractFileExtension(FName)
                    srcFile = TheName + "." + TheExt
                    tgtFile = TheName + "." + NewExt
                    XRename srcFile, tgtFile
                    
                    CurrFile = CurrFile + 1
                    ProgressBar1.Value = Int(CurrFile / SelCount * 100)
                    Cntr = Cntr + 1
    End If
Next i


If oUndoBatch = True And UndoStr <> "" Then
        CreateUNRENBAT ap
End If
If oUndoVBScript = True And UndoStr <> "" Then
        CreateUNRENVBS ap
End If

File1.Refresh
ProgressBar1.Visible = False

If sErrors <> "" Then
    frmErrorLog.ShowErrorLog
End If

End Sub

Private Sub cmdDelChars_Click()
Dim CurrFile As Integer
Dim NewExt As String, TheName As String, TheExt As String
Dim FName As String, fExt As String, frmt As String
Dim OriginalExt As String, srcFile As String, tgtFile As String
Dim result As Integer, fStart As Integer, Cntr As Integer, i As Integer
Dim SelCount As Integer
Dim iDelLeft As Integer
Dim iDelRight As Integer

SelCount = File1.SelCount
If SelCount = 0 Then
                Beep
                Exit Sub
End If


sErrors = ""
UndoStr = ""
vbsUndoStr = ""   ' Publuic Var
ap = Dir1.Path
If Len(ap) > 3 Then ap = ap + "\"

iDelLeft = Val(Me.txtDelFirst.Text)
iDelRight = Val(Me.txtDelLast.Text)

If (iDelLeft + iDelRight = 0) Then      ' both are ZER0
             Beep
              Exit Sub
End If

ProgressBar1.Visible = True

For i = 0 To File1.ListCount - 1
                    If File1.Selected(i) = True Then
                                FName = File1.List(i)                  'file name w/o path
                                TheName = SplitFileNameName(FName)
                                TheExt = ExtractFileExtension(FName)
                                srcFile = ap + TheName + "." + TheExt
                                
                                TheName = DelRight(DelLeft(TheName, iDelLeft), iDelRight)
                                tgtFile = ap + TheName + "." + TheExt
                                
                                XRename srcFile, tgtFile
                                
                                CurrFile = CurrFile + 1
                                ProgressBar1.Value = Int(CurrFile / SelCount * 100)
                                Cntr = Cntr + 1
                    End If
Next i


If oUndoBatch = True And UndoStr <> "" Then
        CreateUNRENBAT ap
End If
If oUndoVBScript = True And UndoStr <> "" Then
        CreateUNRENVBS ap
End If

File1.Refresh
ProgressBar1.Visible = False
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

If sErrors <> "" Then
    frmErrorLog.ShowErrorLog
End If

End Sub

Private Sub cmdDelTo_Click()
Dim CurrFile As Integer
Dim NewExt As String, TheName As String, TheExt As String
Dim FName As String, fExt As String, frmt As String
Dim OriginalExt As String, srcFile As String, tgtFile As String
Dim result As Integer, fStart As Integer, Cntr As Integer, i As Integer
Dim SelCount As Integer
Dim boolInclusive As Boolean
Dim boolMatchCase As Boolean
Dim sLookFor As String


If Me.chkDelMatchCase.Value = vbChecked Then
            boolMatchCase = True
Else
            boolMatchCase = False
End If

If Me.chkInclusive.Value = vbChecked Then
            boolInclusive = True
Else
            boolInclusive = False
End If

sLookFor = Me.txtDelTo.Text

SelCount = File1.SelCount
If (SelCount = 0 Or sLookFor = "") Then
                Beep
                Exit Sub
End If


sErrors = ""
UndoStr = ""
vbsUndoStr = ""   ' Publuic Var
ap = Dir1.Path
If Len(ap) > 3 Then ap = ap + "\"


ProgressBar1.Visible = True



For i = 0 To File1.ListCount - 1
    If File1.Selected(i) = True Then
                        FName = File1.List(i)                  'file name w/o path
                        TheName = SplitFileNameName(FName)
                        TheExt = ExtractFileExtension(FName)
                        srcFile = ap + TheName + "." + TheExt
                        
                        If Me.optDelStart.Value = True Then
                                  TheName = DelLeftTo(TheName, sLookFor, boolMatchCase, boolInclusive)
                        Else
                                  TheName = DelRightTo(TheName, sLookFor, boolMatchCase, boolInclusive)
                        End If
                        
                        tgtFile = ap + TheName + "." + TheExt
                    
                        XRename srcFile, tgtFile

                        CurrFile = CurrFile + 1
                        ProgressBar1.Value = Int(CurrFile / SelCount * 100)
                        Cntr = Cntr + 1
    End If
Next i


If oUndoBatch = True And UndoStr <> "" Then
        CreateUNRENBAT ap
End If
If oUndoVBScript = True And UndoStr <> "" Then
        CreateUNRENVBS ap
End If

File1.Refresh
ProgressBar1.Visible = False
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

If sErrors <> "" Then
    frmErrorLog.ShowErrorLog
End If

End Sub


Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHTMLFilesRename_Click()

Dim SelCount As Integer
Dim CurrFile As Integer
Dim ff As Integer
Dim a$
Dim fExt As String, frmt As String
Dim OriginalExt As String, srcFile As String, tgtFile As String
Dim result As Integer, fStart As Integer, Cntr As Integer, i As Integer


UndoStr = ""
vbsUndoStr = ""   ' Publuic Var
sErrors = ""
ap = Dir1.Path
If Len(ap) > 3 Then ap = ap + "\"


SelCount = File1.SelCount
If SelCount = 0 Then
            Beep
            Exit Sub
End If

ProgressBar1.Visible = True
For i = 0 To File1.ListCount - 1
If File1.Selected(i) = True Then
                srcFile = ap + File1.List(i)
                a$ = Extract_HTML_Title(srcFile)
                a$ = Trim$(ReplaceChars(a$, "", "\/:*?<>|" + Chr$(34)))
                fExt = ExtractFileExtension(File1.List(i))
                tgtFile = ap & a$ & "." & fExt
                If a$ = "" Then tgtFile = srcFile
                
                XRename srcFile, tgtFile
                
                CurrFile = CurrFile + 1
                ProgressBar1.Value = Int(CurrFile / SelCount * 100)
   
End If
Next i


If oUndoBatch = True And UndoStr <> "" Then
    CreateUNRENBAT ap
End If
If oUndoBatch = True And UndoStr <> "" Then
        CreateUNRENBAT ap
End If

File1.Refresh
ProgressBar1.Visible = False
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

If sErrors <> "" Then
    frmErrorLog.ShowErrorLog
End If

End Sub

Private Sub cmdInvertSelection_Click()
Dim i As Integer

For i = 0 To File1.ListCount - 1
        
    File1.Selected(i) = Not (File1.Selected(i))
    
Next i
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

End Sub

Private Sub cmdMP3_Click()
Dim idx As Integer
Dim ff As Integer
Dim a$
Dim FName As String, fExt As String, frmt As String
Dim OriginalExt As String, srcFile As String, tgtFile As String
Dim result As Integer, fStart As Integer, Cntr As Integer, i As Integer
Dim vTag As Variant
Dim sFirst$, sSecond$, sThird$
Dim sTag As String
Dim sSep As String
Dim SelCount As Integer
Dim CurrFile As Integer


UndoStr = ""
sErrors = ""
vbsUndoStr = ""   ' Publuic Var

ap = Dir1.Path
If Len(ap) > 3 Then ap = ap + "\"

sSep = txtSeperator.Text

SelCount = File1.SelCount
If SelCount = 0 Then
    Beep
    Exit Sub
End If

ProgressBar1.Visible = True

For i = 0 To File1.ListCount - 1
            If File1.Selected(i) = True Then
            
                    srcFile = ap + File1.List(i)
                    fExt = ExtractFileExtension(File1.List(i))
                    FName = SplitFileNameName(File1.List(i))
                    ''''''''''''''''''''''''''MP3 TAG CODE
                    sTag = ""
                    vTag = Read_ID3Tag(srcFile)
                    If IsArray(vTag) Then  'TAG info AVAILABLE in MP3 File
                            If cboTAG(0).ListIndex = 0 Then sFirst = "" Else sFirst = vTag(cboTAG(0).ListIndex)
                            If cboTAG(1).ListIndex = 0 Then sSecond = "" Else sSecond = vTag(cboTAG(1).ListIndex)
                            If cboTAG(2).ListIndex = 0 Then sThird = "" Else sThird = vTag(cboTAG(2).ListIndex)
                            vTag(1) = sFirst
                            vTag(2) = sSecond
                            vTag(3) = sThird
                            For idx = 1 To 3
                                    If vTag(idx) <> "" Then
                                        sTag = sTag + vTag(idx) + sSep
                                    End If
                             Next idx
                                        
                               If Right(sTag, Len(sSep)) = sSep Then
                                        sTag = Left(sTag, Len(sTag) - Len(sSep))
                               End If
                    End If
                    ''''''''''''''''''TILL HERE
                    If Not (IsArray(vTag)) Then sTag = "" 'TAG info *NOT* AVAILABLE in MP3 File
                    a$ = sTag
                    tgtFile = ap & a$ & "." & fExt
                    If a$ = "" Then tgtFile = srcFile
                    
                    XRename srcFile, tgtFile
                    
                    
                    CurrFile = CurrFile + 1
                    ProgressBar1.Value = Int(CurrFile / SelCount * 100)
            End If
Next i

If oUndoBatch = True And UndoStr <> "" Then
        CreateUNRENBAT ap
End If
If oUndoVBScript = True And UndoStr <> "" Then
        CreateUNRENVBS ap
End If

File1.Refresh
ProgressBar1.Visible = False
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

If sErrors <> "" Then
    frmErrorLog.ShowErrorLog
End If

End Sub

Private Sub cmdOptions_Click()
frmOptions.Show vbModal
End Sub

Private Sub cmdReplace_Click()
Dim SelCount As Integer
Dim CurrFile As Integer
Dim tfMatchCase As Boolean
Dim ff As Integer
Dim a$
Dim FName As String, fExt As String, frmt As String
Dim OriginalExt As String, srcFile As String, tgtFile As String
Dim result As Integer, fStart As Integer, Cntr As Integer, i As Integer

If txtReplace.Text = "" Then
    Beep
    Exit Sub
End If

UndoStr = ""
vbsUndoStr = ""   ' Publuic Var
sErrors = ""

ap = Dir1.Path
If Len(ap) > 3 Then ap = ap + "\"

SelCount = File1.SelCount
If SelCount = 0 Then
                Beep
                Exit Sub
End If

ProgressBar1.Visible = True
For i = 0 To File1.ListCount - 1
If File1.Selected(i) = True Then
        srcFile = ap + File1.List(i)
        fExt = ExtractFileExtension(File1.List(i))
        FName = SplitFileNameName(File1.List(i))

If optReplaceChars.Value = True Then     ' ***Replace Chars
        a$ = Trim$(ReplaceChars(FName, txtWith.Text, txtReplace.Text))
Else      '*********Replace String
        If chkMatchCase.Value = 0 Then tfMatchCase = False Else tfMatchCase = True
        a$ = Trim$(ReplaceStr(FName, txtReplace.Text, txtWith.Text, tfMatchCase))

End If

        tgtFile = ap & a$ & "." & fExt
        If a$ = "" Then tgtFile = srcFile

        XRename srcFile, tgtFile

    CurrFile = CurrFile + 1
    ProgressBar1.Value = Int(CurrFile / SelCount * 100)


End If
Next i


If oUndoBatch = True And UndoStr <> "" Then
        CreateUNRENBAT ap
End If
If oUndoVBScript = True And UndoStr <> "" Then
        CreateUNRENVBS ap
End If

File1.Refresh
ProgressBar1.Visible = False
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount


If sErrors <> "" Then
    frmErrorLog.ShowErrorLog
End If

End Sub


Private Sub cmdSelectAll_Click()
Dim i As Integer


For i = 0 To File1.ListCount - 1
        
    File1.Selected(i) = True

Next i


lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

End Sub

Private Sub cmdSelectNone_Click()

    File1.Refresh
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

End Sub


Private Sub cmdSerial_Click()

Dim FName As String, fExt As String, frmt As String
Dim OriginalExt As String, srcFile As String, tgtFile As String
Dim result As Integer, fStart As Integer, Cntr As Integer, i As Integer
Dim SelCount As Integer
Dim CurrFile As Integer

SelCount = File1.SelCount
If SelCount = 0 Then
            Beep
            ProgressBar1.Visible = False
            Exit Sub
End If


sErrors = ""          ' Publuic Var
UndoStr = ""        ' Publuic Var
vbsUndoStr = ""   ' Publuic Var

ap = Dir1.Path
If Len(ap) > 3 Then ap = ap + "\"
FName = txtName.Text
If chkMaintainExt.Value = vbChecked Then
    'leave Extension unchainged
    fExt = "*"
Else
    fExt = txtExt.Text
End If

If Left$(fExt, 1) = "." Then fExt = Right$(fExt, Len(fExt) - 1)
If Trim$(fExt) = "" Then
    result = MsgBox("The renamed files will have No Extension," + vbCrLf + "Are you sure?", vbQuestion + vbYesNo, "Warning")
    If result = vbNo Then Exit Sub
End If

fStart = Val(txtNum)
frmt = String(Val(txtDigit.Text), "0")

Cntr = fStart
OriginalExt = fExt

'^^^ Progress Bar code
ProgressBar1.Visible = True


For i = 0 To File1.ListCount - 1
            If File1.Selected(i) = True Then
                            srcFile = ap + File1.List(i)
                            If OriginalExt = "*" Then fExt = ExtractFileExtension(File1.List(i))
                            tgtFile = ap + FName + Format(Cntr, frmt) + "." + fExt
                            XRename srcFile, tgtFile
                            CurrFile = CurrFile + 1
                            ProgressBar1.Value = Int(CurrFile / SelCount * 100)
                            Cntr = Cntr + 1
            End If
Next i


If oUndoBatch = True And UndoStr <> "" Then
        CreateUNRENBAT ap
End If
If oUndoVBScript = True And UndoStr <> "" Then
        CreateUNRENVBS ap
End If

File1.Refresh
ProgressBar1.Visible = False
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

If sErrors <> "" Then
    frmErrorLog.ShowErrorLog
End If

End Sub

Private Sub cmdStretch_Click()
If Frame1.Width < 8000 Then
    Frame1.Width = 9100
    File1.Width = 6600
    cmdStretch.Caption = "<"
    cmdStretch.Left = 8780
    cmdOptions.Visible = False
    cmdExit.Visible = False
    
Else
    Frame1.Width = 4840
    File1.Width = 2370
    cmdStretch.Caption = ">"
    cmdStretch.Left = 4560
    cmdOptions.Visible = True
    cmdExit.Visible = True
    
End If

File1.SetFocus
End Sub

Private Sub cmdTextFilesRename_Click()
Dim SelCount As Integer
Dim CurrFile As Integer
Dim ff As Integer
Dim a$
Dim FName As String, fExt As String, frmt As String
Dim OriginalExt As String, srcFile As String, tgtFile As String
Dim result As Integer, fStart As Integer, Cntr As Integer, i As Integer


UndoStr = ""
vbsUndoStr = ""   ' Publuic Var
sErrors = ""
ap = Dir1.Path
If Len(ap) > 3 Then ap = ap + "\"

SelCount = File1.SelCount
If SelCount = 0 Then
            Beep
            Exit Sub
End If

ProgressBar1.Visible = True
For i = 0 To File1.ListCount - 1
                    If File1.Selected(i) = True Then
                            srcFile = ap + File1.List(i)
                            a$ = RenameTextfile(srcFile)
                            fExt = ExtractFileExtension(File1.List(i))
                            tgtFile = ap & a$ & "." & fExt
                            If a$ = "" Then tgtFile = srcFile
                    
                            XRename srcFile, tgtFile
                    
                            CurrFile = CurrFile + 1
                            ProgressBar1.Value = Int(CurrFile / SelCount * 100)
                    
                    End If
Next i


If oUndoBatch = True And UndoStr <> "" Then
        CreateUNRENBAT ap
End If
If oUndoVBScript = True And UndoStr <> "" Then
        CreateUNRENVBS ap
End If

File1.Refresh
ProgressBar1.Visible = False
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

If sErrors <> "" Then
    frmErrorLog.ShowErrorLog
End If

End Sub

Private Sub Command1_Click()
    SetWindowLong UpDown1.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
     UpDown1.Visible = True
End Sub

Private Sub Dir1_Change()

File1.Path = Dir1.Path
txtPath.Text = Dir1.Path
txtPath.SelStart = Len(txtPath.Text)
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 32 Then
    Dir1.Path = Dir1.List(Dir1.ListIndex)
End If

End Sub


Private Sub Dir1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
   PopupMenu mnuHidden, 2
End If

End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
If Err Then
    Beep
    Drive1.Drive = Dir1.Path
End If
End Sub

Private Sub file1_Click()

lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

End Sub

Private Sub File1_DblClick()

    RunFile File1.FileName, File1.Path, SW_SHOWNORMAL


End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnuHidden, 2
End If

End Sub


Private Sub File1_PathChange()

lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

AddScroll File1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 1  ' ^A
        cmdSelectAll.Value = True
        KeyAscii = 0
    Case 9  ' ^I
        cmdInvertSelection.Value = True
        KeyAscii = 0
    Case 14 ' ^N
        cmdSelectNone.Value = True
        KeyAscii = 0
    Case 15 ' ^O
        cmdOptions.Value = True
        KeyAscii = 0
    
    Case Else
        ' DO NOTHING
End Select
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'MsgBox KeyCode

If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case 188, 190 ' "<" or  ">"
                cmdStretch.Value = True
                KeyCode = 0
            Case Else
                ' DO NOTHING
        End Select
End If

End Sub

Private Sub Form_Load()
Dim sMasks As String
Dim MasksArray As Variant
Dim idx As Integer
Dim AstDotAstFound As Boolean

File1.FontName = "tahoma"
File1.FontSize = "8"
Dim X As Control
On Error Resume Next
        For Each X In Me.Controls
                X.Font.Name = "tahoma"
                X.Font.Size = 8
        Next X
btnFlat cmdStretch
On Error GoTo 0

AddBorderToAllTextBoxes Me

LoadXRenSettings
If oFlatBtns Then
        FlatAllBtns Me
End If


    AddScroll File1
    fcolor = txtPath.ForeColor
    txtPath.Text = Dir1.Path
    CenterFormUp Me
    
''''''''''''''Fill MP3 TAG Combos
    Dim Cntr%
    For Cntr = 0 To 2
        cboTAG(Cntr).AddItem "-none-"
        cboTAG(Cntr).AddItem "Title"
        cboTAG(Cntr).AddItem "Artist"
        cboTAG(Cntr).AddItem "Album"
    Next Cntr
    
    cboTAG(0).ListIndex = 1
    cboTAG(1).ListIndex = 2
    cboTAG(2).ListIndex = 3
'''''''''''''''''''''''''''''''''''''''

'''''''''''''Fill FileList Filter Combo

    sMasks = GetSubVal(HKEY_LOCAL_MACHINE, "software\XRen", "Masks")
    
    If sMasks <> "" Then
                AstDotAstFound = False   ' *.*
                MasksArray = Split(sMasks, "|")
                For idx = LBound(MasksArray) To UBound(MasksArray)
                     cboFileMask.AddItem MasksArray(idx)
                     If MasksArray(idx) = "*.*" Then AstDotAstFound = True
                Next idx
                If AstDotAstFound = False Then cboFileMask.AddItem "*.*"
    Else  'No Entries found
                 cboFileMask.AddItem "*.*"
    End If
'''''''''''''''''''''''''''''''''''''''

'|>>>>>>>>> NEW CODE
'Dim idx As Integer
Dim inner As Integer
Dim sArgPath As String
Dim sArgFile As String

'<ERROR_CAPTURE>
On Error GoTo Err_NoArgs

sArgPath = ExtractDirName(LongFileName(CStr(gv_Args(LBound(gv_Args)))))
frmXRenMain.Dir1.Path = sArgPath

           For idx = LBound(gv_Args) To UBound(gv_Args)
                          gv_Args(idx) = ExtractFileName(LongFileName(gv_Args(idx)))
                          'MsgBox gv_Args(idx)
           Next idx
           
           For idx = LBound(gv_Args) To UBound(gv_Args)
                    For inner = 0 To File1.ListCount - 1
                            If gv_Args(idx) = File1.List(inner) Then File1.Selected(inner) = True
                    Next inner
            Next idx
Exit Sub

'<ERROR HANDLER>
Err_NoArgs:
                Err = 0
                Exit Sub
'</ERROR HANDLER>
End Sub

Private Sub Form_Unload(Cancel As Integer)

SaveXRenSettings

Unload frmShellNotify

End Sub

Private Sub Label4_Click()
End Sub


Private Sub lblMailTo_Click()

'Shell "rundll32 url.dll,FileProtocolHandler  mailto:mdsy@excite.com"
HyperLink "mailto:mdsy@excite.com"

End Sub

Private Sub mnuCopyFileName_Click()
MsgBox File1.FileName
End Sub

Private Sub mnuManualRename_Click()
Dim SelCount As Integer
Dim srcFile As String
Dim i As Integer

ap = Dir1.Path
If Len(ap) > 3 Then ap = ap + "\"


SelCount = 0
For i = 0 To File1.ListCount - 1
  If File1.Selected(i) = True Then SelCount = SelCount + 1
Next i
If SelCount = 0 Then
    Beep
    Exit Sub
End If



For i = 0 To File1.ListCount - 1
If File1.Selected(i) = True Then

        srcFile = ap + File1.List(i)
        frmManualRename.OldFileName = srcFile
        frmManualRename.Show vbModal
        If frmManualRename.Aborted = True Then
                   Exit For
        End If
        
End If
Next i

Unload frmManualRename
File1.Refresh
End Sub

Private Sub mnuOpen_Click()
    Shell "explorer.exe " & Dir1.Path, vbNormalFocus
End Sub

Private Sub mnuRefresh_Click()

Dir1.Refresh
File1.Refresh
lblSelFiles.Caption = File1.SelCount & " files selected" & " / " & File1.ListCount

End Sub


Private Sub mnuSortDate_Click()
mnuSortName.Checked = False
mnuSortExtension.Checked = False
mnuSortDate.Checked = True
mnuSortSize.Checked = False

File1.SortMethod = xfByDate
File1.Sort

End Sub

Private Sub mnuSortExtension_Click()
mnuSortName.Checked = False
mnuSortExtension.Checked = True
mnuSortDate.Checked = False
mnuSortSize.Checked = False

File1.SortMethod = xfbyextension
File1.Sort

End Sub


Private Sub mnuSortName_Click()
mnuSortName.Checked = True
mnuSortExtension.Checked = False
mnuSortDate.Checked = False
mnuSortSize.Checked = False

File1.SortMethod = xfByName
File1.Sort
End Sub

Private Sub mnuSortOrderAscending_Click()

mnuSortOrderAscending.Checked = True
mnuSortOrderDescending.Checked = False


        File1.SortOrder = xfSortAscending
        File1.Sort
        
End Sub

Private Sub mnuSortOrderDescending_Click()

mnuSortOrderAscending.Checked = False
mnuSortOrderDescending.Checked = True


        File1.SortOrder = xfSortDescending
        File1.Sort
        

End Sub


Private Sub mnuSortSize_Click()
mnuSortName.Checked = False
mnuSortExtension.Checked = False
mnuSortDate.Checked = False
mnuSortSize.Checked = True

File1.SortMethod = xfBySize
File1.Sort

End Sub


Private Sub picLogo_DblClick()
frmAboutXRen.Show vbModal
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then txtName.SetFocus
If SSTab1.Tab = 2 Then txtBefore.SetFocus
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_GotFocus()
'Text1.SelStart = 2
'Text1.SelLength = Len(Text1.Text) - 2

End Sub

Private Sub txtAfter_KeyPress(KeyAscii As Integer)

Select Case Chr(KeyAscii)
    Case "\", "/", ":", "*", "?", "<", ">", "|", Chr$(34)
        'illegal chars
        Beep
        KeyAscii = 0
    Case Else
    'do nothing
End Select

End Sub

Private Sub txtBefore_KeyPress(KeyAscii As Integer)

Select Case Chr(KeyAscii)
    Case "\", "/", ":", "*", "?", "<", ">", "|", Chr$(34)
        'illegal chars
        Beep
        KeyAscii = 0
    Case Else
    'do nothing
End Select

End Sub


Private Sub txtDelFirst_Change()
Dim v As Byte


v = Val(txtDelFirst.Text)
If v = 0 Then
    v = 1
    txtDelFirst = "1"
End If

upDelFirst.Value = v

End Sub

Private Sub txtDelFirst_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 48 To 57, 8
    'key is 0 to 9 Or {DEL} ok.. do nothing
    Case Else
    Beep
    KeyAscii = 0
End Select

End Sub


Private Sub txtDelLast_Change()
Dim v As Byte


v = Val(txtDelLast.Text)
If v = 0 Then
    v = 1
    txtDelLast = "1"
End If

upDelLast.Value = v

End Sub

Private Sub txtDelLast_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 48 To 57, 8
    'key is 0 to 9 Or {DEL} ok.. do nothing
    Case Else
    Beep
    KeyAscii = 0
End Select

End Sub


Private Sub txtDelTo_KeyPress(KeyAscii As Integer)


Select Case Chr(KeyAscii)
    Case "\", "/", ":", "*", "?", "<", ">", "|", Chr$(34)
        'illegal chars
        Beep
        KeyAscii = 0
    Case Else
    'do nothing
End Select

End Sub


Private Sub txtDigit_Change()
Dim v As Byte
Dim frmt As String
If txtDigit = "" Then txtDigit = "1"
frmt = String(Val(txtDigit.Text), "0")
txtNum.Text = Format(txtNum.Text, frmt)
v = Val(txtDigit.Text)
If v = 0 Then v = 1
UpDown1.Value = v

End Sub

Private Sub txtDigit_GotFocus()

'txtDigit.SelStart = 0
'txtDigit.SelLength = Len(txtDigit.Text)

End Sub

Private Sub txtDigit_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 49 To 57, 8
    'key is 0 to 5 Or {DEL} ok.. do nothing
    Case Else
    Beep
    KeyAscii = 0
End Select

End Sub

Private Sub txtExt_GotFocus()

'txtExt.SelStart = 0
'txtExt.SelLength = Len(txtExt.Text)


End Sub

Private Sub txtExt_KeyPress(KeyAscii As Integer)

Select Case Chr(KeyAscii)
    
    Case "\", "/", ":", "*", "?", "<", ">", "|", Chr$(34), "."
        ' illegal chars
        Beep
        KeyAscii = 0
    
    Case Else
    ' do nothing

End Select

End Sub


Private Sub txtName_GotFocus()

'txtName.SelStart = 0
'txtName.SelLength = Len(txtName.Text)

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

'since the textBox is multi-line...
If KeyAscii = 13 Or KeyAscii = 10 Then
    Beep
    KeyAscii = 0
End If


Select Case Chr(KeyAscii)
    Case "\", "/", ":", "*", "?", "<", ">", "|", Chr$(34)
        'illegal chars
        Beep
        KeyAscii = 0
    Case Else
    'do nothing
End Select

End Sub

Private Sub txtNewExt_KeyPress(KeyAscii As Integer)

Select Case Chr(KeyAscii)
    Case "\", "/", ":", "*", "?", "<", ">", "|", Chr(34), "."
        'illegal chars
        Beep
        KeyAscii = 0
    Case Else
    'do nothing
End Select

End Sub


Private Sub txtNum_GotFocus()

'txtNum.SelStart = 0
'txtNum.SelLength = Len(txtNum.Text)

End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 48 To 57, 8
    'key is 0 to 9 Or {DEL} ok.. do nothing
    Case Else
    Beep
    KeyAscii = 0
End Select
End Sub

Private Sub txtPath_Change()
On Error Resume Next
    Dir1.Path = txtPath.Text
If Err Then
   txtPath.ForeColor = vbRed
Else
    txtPath.ForeColor = fcolor
End If

If LCase$(txtPath.Text) = "desktop" Then
    Dir1.Path = GetWinDir() & "Desktop"
    txtPath.Text = Dir1.Path
    txtPath.SelStart = Len(txtPath.Text)
End If
End Sub

Private Sub txtPath_GotFocus()
SelectAllText txtPath
End Sub


Private Sub txtReplace_Change()

If txtReplace.Text = "0..9" Then
    txtReplace.Text = "0123456789"
    txtReplace.SelStart = Len(txtReplace.Text)
End If
End Sub

Private Sub txtReplace_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
    Case "\", "/", ":", "*", "?", "<", ">", "|", Chr$(34)
        'illegal chars
        Beep
        KeyAscii = 0
    Case Else
    'do nothing
End Select

End Sub


Private Sub txtSeperator_KeyPress(KeyAscii As Integer)

Select Case Chr(KeyAscii)
    Case "\", "/", ":", "*", "?", "<", ">", "|", Chr$(34)
        'illegal chars
        Beep
        KeyAscii = 0
    Case Else
    'do nothing
End Select
    

End Sub


Private Sub txtWith_KeyPress(KeyAscii As Integer)

If (txtWith.SelLength = 1 And optReplaceChars.Value = True) Then
    txtWith.Text = ""
End If
Select Case Chr(KeyAscii)
    Case "\", "/", ":", "*", "?", "<", ">", "|", Chr$(34)
        'illegal chars
        Beep
        KeyAscii = 0
    Case Else
    'do nothing
End Select
    
If Len(txtWith.Text) >= 1 Then
    Select Case KeyAscii
        Case 8
        '{BK_SPC} OK
        Case Else
        If optReplaceChars.Value = True Then
                Beep
                KeyAscii = 0
        End If
    End Select
End If

End Sub


Private Sub upDelFirst_Change()

txtDelFirst.Text = upDelFirst.Value


End Sub

Private Sub upDelLast_Change()

txtDelLast.Text = upDelLast.Value

End Sub


Private Sub UpDown1_Change()
txtDigit.Text = UpDown1.Value

End Sub



