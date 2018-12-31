VERSION 5.00
Begin VB.Form mains 
   BackColor       =   &H00C0E0FF&
   Caption         =   "ATM card management"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8400
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "HELP  AND SUPPORT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4680
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "FOR NEW USER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "FOR AN EXISTING USER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4680
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   1485
      Left            =   120
      Picture         =   "main.frx":0000
      Top             =   840
      Width           =   15180
   End
   Begin VB.Image Image2 
      Height          =   4920
      Left            =   9960
      Picture         =   "main.frx":5029
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4410
   End
   Begin VB.Image Image1 
      Height          =   4995
      Left            =   240
      Picture         =   "main.frx":16C29
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   3735
   End
End
Attribute VB_Name = "mains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
existing.Show
Unload Me
End Sub

Private Sub Command2_Click()
APPLICATION.Show
Unload Me
End Sub

Private Sub Command3_Click()
help.Show
Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
thank.Show
End Sub
