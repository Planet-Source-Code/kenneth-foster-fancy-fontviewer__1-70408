VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   405
      TabIndex        =   9
      Top             =   5250
      Width           =   5400
   End
   Begin Project1.ucScrollbar Vs1 
      Height          =   1845
      Left            =   5415
      Top             =   345
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   3254
      Max             =   32767
      SmallChange     =   50
      LargeChange     =   150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Fontsize"
      Height          =   630
      Left            =   5910
      TabIndex        =   8
      Top             =   2880
      Width           =   1035
   End
   Begin VB.TextBox txtFS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   6195
      TabIndex        =   7
      Text            =   "14"
      Top             =   2535
      Width           =   420
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   405
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   4125
      Width           =   5400
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   7980
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1605
      Width           =   2205
   End
   Begin VB.PictureBox picHidden 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   240
      Picture         =   "Form1.frx":000E
      ScaleHeight     =   2235
      ScaleWidth      =   5775
      TabIndex        =   0
      Top             =   150
      Width           =   5775
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4CEB4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1830
         Left            =   255
         ScaleHeight     =   1830
         ScaleWidth      =   4905
         TabIndex        =   1
         Top             =   195
         Width           =   4905
         Begin VB.Frame fFrame 
            BackColor       =   &H00E4CEB4&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1890
            Left            =   0
            TabIndex        =   2
            Top             =   -60
            Width           =   4905
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Label1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Index           =   0
               Left            =   90
               TabIndex        =   6
               Top             =   105
               Width           =   1635
            End
         End
      End
   End
   Begin VB.Label lblExample 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AaBbCcDdEeFfGgHhIiJj KkLlMmNnOoPpQqRr SsTtUuVvWwXxYyZz 0123456789"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   405
      TabIndex        =   5
      Top             =   2430
      Width           =   5400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
 End Type
    Private Const SB_VERT = 1
    Private Const EM_GETFIRSTVISIBLELINE = &HCE
    Private Const EM_GETLINECOUNT = &HBA
    Private Const EM_GETRECT = &HB2
'This project uses 2 pictureboxes, 1 frame, 1 label (array), and 1 scrollbar
Dim sStore As Integer

Private Sub Load_List()
    On Error GoTo finish:
    Dim i As Integer
    Dim wWidth As Double
    
    Label1(0).FontName = List1.LIST(0)
    Label1(0).Left = cCenter(fFrame.Width, Label1(0).Width)
    Label1(0).ToolTipText = List1.LIST(0)
    Label1(0).Caption = List1.LIST(i)
    
    wWidth = Label1(0).Height + Label1(0).top
    For i = 1 To List1.ListCount - 1
        Load Label1(i)
        Label1(i).top = wWidth
        Label1(i).AutoSize = True
        Label1(i).Font.Name = List1.LIST(i)
        Label1(i).Caption = List1.LIST(i)
        Label1(i).ToolTipText = List1.LIST(i)
        Label1(i).Left = cCenter(fFrame.Width, Label1(i).Width)
        wWidth = Label1(i).Height + wWidth + 15
        Label1(i).Visible = True
      
    Next i
    fFrame.Height = wWidth + Label1(i - 1).Height
    Vs1.Min = 0
    Vs1.Max = (fFrame.Height - Picture1.Height) / 10

    Exit Sub
finish:
i = MsgBox(Err.Description, vbCritical)
End Sub

Private Function cCenter(LSide As Long, RSide As Long) As Long
    cCenter = (LSide - RSide) / 2
End Function

Private Sub Command1_Click()
   lblExample.FontSize = Int(txtFS.Text)
   Text1.FontSize = Int(txtFS.Text)
End Sub

Private Sub Form_Load()
Dim x As Integer

    fFrame.BackColor = &HE4CEB4
    
    For x = 0 To Screen.FontCount - 1
       List1.AddItem Screen.Fonts(x)
    Next x
    
    sStore = 0
    lblExample.BackColor = vbWhite
    Load_List
End Sub

Private Sub Label1_Click(Index As Integer)
    Label1(sStore).ForeColor = vbBlack
    Label1(sStore).FontBold = False
    Label1(Index).ForeColor = 8914708
    Label1(Index).Font.Bold = True
    sStore = Index
    lblExample.Font = Label1(Index).Font
    Text1.Text = Label1(Index).FontName
    Text2.Text = Label1(Index).FontName
    Text1.Font = Label1(Index).Font
End Sub

Private Sub Vs1_Change()
    Vs1_Scroll
End Sub

Private Sub Vs1_Scroll()
    fFrame.top = -(Vs1.Value) * 10
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   ShowScrollBars Text1
End Sub

Private Sub ShowScrollBars(theTextbox As TextBox)
    Dim firstVisibleLine As Long
    Dim r As RECT
    Dim numberOfLines As Long
    Dim numberOfVisibleLines As Long
    Dim rectHeight As Long
    Dim lineHeight As Long
    Dim hWnd As Long

    hWnd = theTextbox.hWnd

    firstVisibleLine = SendMessage(hWnd, EM_GETFIRSTVISIBLELINE, 0, 0)

    If firstVisibleLine <> 0 Then
        ShowScrollBar hWnd, SB_VERT, 1
    Else
        numberOfLines = SendMessage(hWnd, EM_GETLINECOUNT, 0, 0)
        SendMessage hWnd, EM_GETRECT, 0, r
        rectHeight = r.Bottom - r.top
        lineHeight = theTextbox.Parent.TextHeight("W") / Screen.TwipsPerPixelY
        numberOfVisibleLines = rectHeight / lineHeight

        If numberOfVisibleLines < numberOfLines Then
            ShowScrollBar hWnd, SB_VERT, 1
        Else
            ShowScrollBar hWnd, SB_VERT, 0
        End If
    End If
End Sub
