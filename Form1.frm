VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00D8E9EC&
   Caption         =   "XP_ArrowButton Test Form"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   240
      Width           =   3015
   End
   Begin XP_ArrowButton.XPArrowButton XPArrowButton1 
      Height          =   360
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   360
      _extentx        =   635
      _extenty        =   635
      buttontype      =   6
   End
   Begin XP_ArrowButton.XPArrowButton XPArrowButton1 
      Height          =   360
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   360
      _extentx        =   635
      _extenty        =   635
      buttontype      =   3
   End
   Begin XP_ArrowButton.XPArrowButton XPArrowButton1 
      Height          =   360
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   360
      _extentx        =   635
      _extenty        =   635
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Next..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Next..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub XPArrowButton1_Click(Index As Integer)

Text1.Text = "Button" & Index & "_Click" & vbCrLf & Text1.Text

End Sub

Private Sub XPArrowButton1_MouseDown(Index As Integer)

Text1.Text = "Button" & Index & "_MouseDown" & vbCrLf & Text1.Text

End Sub

Private Sub XPArrowButton1_MouseOut(Index As Integer)

Text1.Text = "Button" & Index & "_MouseOut" & vbCrLf & Text1.Text

End Sub

Private Sub XPArrowButton1_MouseIn(Index As Integer)

Text1.Text = "Button" & Index & "_MouseIn" & vbCrLf & Text1.Text

End Sub
