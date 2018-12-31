VERSION 5.00
Begin VB.Form help 
   BackColor       =   &H00C0E0FF&
   Caption         =   "HELP TOPICS"
   ClientHeight    =   3060
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9480
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0080C0FF&
      Caption         =   "CHANGING THE PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   4815
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080C0FF&
      Caption         =   "FILLING OF APPLICATION FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   4815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "RENEW / REISSUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   4815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "CANCELLATION/BLOCKING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "ENTRY FOR APPLICATION FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "SEING THE RECORDS OF CUSTOMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   6660
      Left            =   5640
      Picture         =   "help.frx":0000
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   4290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HELP TOPICS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3960
      TabIndex        =   0
      Top             =   600
      Width           =   9615
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
st = "1"
Unload Me
helping.Show
End Sub

Private Sub Command2_Click()
st = "2"
Unload Me
helping.Show
End Sub

Private Sub Command3_Click()
st = "3"
Unload Me
helping.Show
End Sub

Private Sub Command4_Click()
st = "4"
Unload Me
helping.Show
End Sub

Private Sub Command5_Click()
Unload Me
main.Show
End Sub

Private Sub Command7_Click()
st = "5"
Unload Me
helping.Show
End Sub

Private Sub Command8_Click()
st = "6"
Unload Me
helping.Show
End Sub
