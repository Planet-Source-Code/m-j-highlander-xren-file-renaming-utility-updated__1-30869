VERSION 5.00
Begin VB.UserControl XFileListBox 
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   ScaleHeight     =   4680
   ScaleWidth      =   3780
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   315
      TabIndex        =   2
      Top             =   2685
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.ListBox lstExtended 
      BackColor       =   &H00FEFDED&
      Height          =   1230
      Left            =   330
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   1365
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.ListBox lstSimple 
      BackColor       =   &H00F4FFFF&
      Height          =   1035
      Left            =   285
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   2085
   End
End
Attribute VB_Name = "XFileListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_MultiSelect As MULTISELECTSTYLE
Private theListBox As ListBox

Public Enum MULTISELECTSTYLE
            xfSelectSimple = 1
            xfSelectExtended = 2
End Enum

Public Event PathChange()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DblClick()
Public Event Click()

Public Enum enumSORTORDER
                    xfSortAscending = 1
                    xfSortDescending = 2
End Enum

Public Enum enumSORTMETHOD
                    
                    xfByName = 1
                    xfbyextension = 2
                    xfBySize = 3
                    xfByDate = 4
End Enum

Private p_SortOrder As enumSORTORDER
Private p_SortMethod As enumSORTMETHOD
Private p_SelCount As Integer
Public Property Let SelCount(ByVal bNewValue As Integer)

'IS IT NEEDED ???

p_SelCount = bNewValue

End Property

Public Property Get SortMethod() As enumSORTMETHOD

SortMethod = p_SortMethod

End Property

Public Property Let SortMethod(ByVal vNewSortMethod As enumSORTMETHOD)
p_SortMethod = vNewSortMethod
End Property

Public Property Get SelCount() As Integer

'Dim i As Integer
'Dim itmp As Integer
'
'itmp = 0
'
'For i = 0 To theListBox.ListCount - 1
'    If theListBox.Selected(i) Then
'        itmp = itmp + 1
'    End If
'Next i
'
'p_SelCount = itmp
'SelCount = p_SelCount

SelCount = theListBox.SelCount

End Property




Public Property Get ListCount() As Integer

ListCount = theListBox.ListCount

End Property

Public Property Get List(i As Integer) As String

List = theListBox.List(i)

End Property


Public Property Get FileName() As String

FileName = theListBox.List(theListBox.ListIndex)

End Property

Public Property Get MultiSelect() As MULTISELECTSTYLE

MultiSelect = p_MultiSelect
    
End Property

Public Property Let MultiSelect(ByVal state As MULTISELECTSTYLE)
p_MultiSelect = state


Select Case MultiSelect
    
    Case xfSelectExtended
            lstExtended.Visible = True
            lstSimple.Visible = False
            Set theListBox = lstExtended
            
    Case xfSelectSimple
            lstSimple.Visible = True
            lstExtended.Visible = False
            Set theListBox = lstSimple

End Select


theListBox.Top = 0
theListBox.Left = 0
theListBox.Height = Height
theListBox.Width = Width

File1_PathChange

End Property


Public Sub Sort()

If File1.ListCount = 0 Then
            theListBox.Clear
            Exit Sub
End If


ReDim tmpArray(0 To File1.ListCount - 1) As String
Dim vArray As Variant
Dim i As Integer
Dim MaxItem As Integer

For i = 0 To File1.ListCount - 1
        tmpArray(i) = File1.List(i)
Next i

vArray = tmpArray


Select Case SortMethod
            
           
            Case xfByName
     
                   SortByName vArray

            Case xfbyextension
                   SortByExt vArray
                   
            Case xfBySize
                    SortBySize vArray
                    
            Case xfByDate
                    SortByDate vArray
End Select

MaxItem = File1.ListCount - 1
theListBox.Clear

theListBox.Visible = False
For i = 0 To MaxItem
        'theListBox.List(i) = vArray(i)  'same as below
        theListBox.AddItem vArray(i)
Next i
theListBox.Visible = True
End Sub

Private Sub SortByName(sSortArray As Variant)
    
    If SortOrder = xfSortAscending Then
            ShellSortAsc sSortArray
    Else
            ShellSortDesc sSortArray
    End If
    
        
End Sub

Private Function SortByExt(sSortArray As Variant) As Variant

    
    If SortOrder = xfSortAscending Then
            ShellSortExtAsc sSortArray  'sort by extension
            ShellSortExtAscEx sSortArray
    Else
            ShellSortExtDesc sSortArray
            ShellSortExtDescEx sSortArray  'sort similar extensiobns by name (second key sort)
    End If
    


End Function

Private Function SortByDate(sSortArray As Variant) As Variant
ChDir File1.Path
ChDrive File1.Path

    If SortOrder = xfSortAscending Then
            ShellSortDateAsc sSortArray
    Else
            ShellSortDateDesc sSortArray
    End If

End Function

Private Function SortBySize(sSortArray As Variant) As Variant
ChDir File1.Path
ChDrive File1.Path

    If SortOrder = xfSortAscending Then
            ShellSortSizeAsc sSortArray
    Else
            ShellSortSizeDesc sSortArray
    End If

End Function


Private Sub File1_PathChange()
Dim i As Integer

'theListBox.Clear
'For i = 0 To File1.ListCount - 1
'        theListBox.AddItem File1.List(i)
'Next

Sort
SelCount = 0
RaiseEvent PathChange
End Sub


Private Sub lstExtended_Click()
RaiseEvent Click
End Sub

Private Sub lstExtended_DblClick()
RaiseEvent DblClick
End Sub

Private Sub lstExtended_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub


Private Sub lstSimple_Click()
RaiseEvent Click
End Sub

Private Sub lstSimple_DblClick()
RaiseEvent DblClick
End Sub

Private Sub lstSimple_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub


Private Sub UserControl_Initialize()
Dim i As Integer

MultiSelect = xfSelectSimple


lstSimple.Clear
lstExtended.Clear

For i = 0 To File1.ListCount - 1
        lstSimple.AddItem File1.List(i)
        lstExtended.AddItem File1.List(i)
Next

SortOrder = xfSortAscending
SortMethod = xfByName

Sort
End Sub


Private Sub UserControl_Resize()

theListBox.Top = 0
theListBox.Left = 0
theListBox.Height = Height
theListBox.Width = Width

End Sub



Public Property Get Path() As String

Path = File1.Path

End Property

Public Property Let Path(ByVal sPath As String)

File1.Path = sPath

End Property

Public Property Get hwnd() As Long

hwnd = theListBox.hwnd

End Property


Public Property Get Selected(ByVal i As Integer) As Boolean

Selected = theListBox.Selected(i)

End Property

Public Property Let Selected(ByVal i As Integer, ByVal bNewValue As Boolean)

theListBox.Selected(i) = bNewValue

End Property

Public Sub Refresh()

Dim i As Integer

'For i = 0 To theListBox.ListCount - 1
'    If theListBox.Selected(i) = True Then
'            theListBox.Selected(i) = False
'    End If
    
'Next i

'theListBox.Refresh


'lstSimple.Clear
'lstExtended.Clear

File1.Refresh
'For i = 0 To File1.ListCount - 1
'        lstSimple.AddItem File1.List(i)
'        lstExtended.AddItem File1.List(i)
'Next

Sort

'SelCount = 0

End Sub
Public Property Let FontName(ByVal sFont As String)
Attribute FontName.VB_Description = "Returns a Font object."
Attribute FontName.VB_UserMemId = -512
    
   lstSimple.FontName = sFont
   lstExtended.FontName = sFont
   
End Property



Public Property Let FontSize(ByVal iSize As Integer)
    
   lstSimple.FontSize = iSize
   lstExtended.FontSize = iSize
   
End Property


Public Property Get SortOrder() As enumSORTORDER

SortOrder = p_SortOrder

End Property

Public Property Let SortOrder(ByVal vNewSortOrder As enumSORTORDER)
p_SortOrder = vNewSortOrder
End Property

Public Property Get Pattern() As String
    
    Pattern = File1.Pattern
    
End Property

Public Property Let Pattern(ByVal sNewPattern As String)
Dim i As Integer
    
File1.Pattern = sNewPattern
    'Reload ListBoxes

lstSimple.Clear
lstExtended.Clear

For i = 0 To File1.ListCount - 1
        lstSimple.AddItem File1.List(i)
        lstExtended.AddItem File1.List(i)
Next

Sort
    
End Property

Public Property Get Enabled() As Boolean

Enabled = theListBox.Enabled

End Property

Public Property Let Enabled(ByVal bEnabled As Boolean)

    theListBox.Enabled = bEnabled

End Property
