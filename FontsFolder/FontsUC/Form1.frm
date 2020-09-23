VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin Project1.FontList FontList1 
      Height          =   2250
      Left            =   540
      TabIndex        =   2
      Top             =   105
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   3969
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   735
      TabIndex        =   1
      Top             =   2505
      Width           =   4665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This is a sample of the selected font."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   420
      TabIndex        =   0
      Top             =   3045
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FontList1_Click()
    Text1.Text = FontList1.Selected
    Label1.FontName = FontList1.Selected
End Sub

