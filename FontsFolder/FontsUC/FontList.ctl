VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl FontList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF80FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   MaskColor       =   &H00FF80FF&
   ScaleHeight     =   4770
   ScaleWidth      =   6150
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   5835
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1500
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1845
      Left            =   345
      ScaleHeight     =   1845
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   180
      Width           =   5145
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   1905
         Left            =   -30
         TabIndex        =   2
         Top             =   -30
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   3360
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   14995124
         ForeColorSel    =   14995124
         BackColorBkg    =   14995124
         GridColorFixed  =   14995124
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   60
      Picture         =   "FontList.ctx":0000
      Top             =   3015
      Width           =   5745
   End
End
Attribute VB_Name = "FontList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*************************************
'      Project: FontList
'      Version: 1.1.0
'   Programmer: Ken Foster
'         Date: April 2008
'*************************************

Private Const m_def_Selected = "Arial"

Dim m_Selected As String

Dim LastSel$    'The Selection that is colored
Event Click()

Private Sub Grid1_Click()
   Selected = GetFont
End Sub

Private Sub UserControl_Initialize()
Dim x As Integer

    UserControl.Picture = Image1.Picture
    UserControl.MaskPicture = UserControl.Image
    Picture1.BackColor = &HE4CEB4
    Grid1.BackColor = &HE4CEB4
    
    FillList
End Sub

Private Sub UserControl_Resize()
   UserControl.Width = Image1.Width
   UserControl.Height = Image1.Height
End Sub

Public Property Get Selected() As String
   Let Selected = m_Selected
End Property

Public Property Let Selected(ByVal NewSelected As String)
   Let m_Selected = NewSelected
   
   PropertyChanged "Selected"
End Property

Private Sub UserControl_InitProperties()
   Let Selected = m_def_Selected
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      Let Selected = .ReadProperty("Selected", m_def_Selected)
   End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "Selected", m_Selected, m_def_Selected
   End With
End Sub

Public Function FindFont(FontName$) As Boolean      'hilites the font, assures it's visible, or returns false
Dim a1%
Dim s1$
Dim nRows%  'The number of rows displayed -- not obviously available

    'Calculate nRows
    nRows = Grid1.Height \ Grid1.RowHeight(0)
    
    'Get a clean name
    s1 = UCase(Trim(FontName))
    
    'Find the sucker
    For a1 = 0 To List1.ListCount - 1
        If Trim(UCase(List1.List(a1))) = s1 Then
            'This is it
            Grid1.RowSel = a1
            a1 = a1 - (nRows / 2) + 1   'about the middle
            If a1 < 0 Then a1 = 0
            Grid1.TopRow = a1
            FindFont = True
            Exit Function
        End If
    Next
    
    'No such
    FindFont = False
    Exit Function
    
End Function

Public Function GetFont$()  'Returns currently selected FontName or -1
    With Grid1
        If .RowSel = -1 Then
            GetFont = ""
        Else
            .Row = .RowSel
            GetFont = .Text
            Exit Function
        End If
    End With
    
End Function

Public Sub FillList()
Dim a1%

    'Get sorted list of fonts
    List1.Clear
    For a1 = 0 To Screen.FontCount - 1
        List1.AddItem Screen.Fonts(a1)
    Next

    'Set no item as selected
    LastSel = -1
    
    'Set up the grid
    With Grid1
        .ColWidth(0) = .Width
        .Rows = List1.ListCount
        For a1 = 0 To List1.ListCount - 1
            .Row = a1
            .Text = List1.List(a1)
            .CellFontName = List1.List(a1)
            
        Next
    End With
            
End Sub

Private Sub Grid1_SelChange()   'This creates the blue hilite
Static Inhibit As Boolean
Dim NewSel%

    If Inhibit Then Exit Sub
    Inhibit = True

    With Grid1
    
        'Unlite any last selection
        NewSel = .Row
        If LastSel > -1 Then
            .Row = LastSel
            .CellBackColor = &HE4CEB4
            .CellForeColor = RGB(0, 0, 0)
        End If
            
        'Lite this selection
        .Row = NewSel
        .CellForeColor = RGB(255, 255, 255)
        .CellBackColor = RGB(0, 0, 255)
        
        'Remember
        LastSel = NewSel
    End With
    
    Inhibit = False
    
    Selected = GetFont
    RaiseEvent Click
End Sub

