VERSION 5.00
Begin VB.Form existing 
   BackColor       =   &H00C0E0FF&
   Caption         =   "EXISTING USER"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080C0FF&
      Caption         =   "APPLICATION FORM DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8400
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080C0FF&
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8400
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9480
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "RENEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "BLOCKING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "CANCELLATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "REISSUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Desired Option:-"
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   12975
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   6000
      Picture         =   "Form1.frx":0000
      Top             =   3840
      Width           =   4875
   End
End
Attribute VB_Name = "existing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
st = "2"
Unload Me
Card.Show
End Sub

Private Sub Command2_Click()
st = "3"
Unload Me
Card.Show
End Sub

Private Sub Command3_Click()
st = "4"
Unload Me
Card.Show
End Sub

Private Sub Command4_Click()
st = "1"
Unload Me
Card.Show
End Sub

Private Sub Command5_Click()
Unload Me
main.Show
End Sub

Private Sub Command6_Click()
st = "5"
Unload Me
Card.Show
End Sub

Private Sub Command7_Click()
st = "6"
Unload Me
APPLICATION_FORM.Show
End Sub


