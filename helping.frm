VERSION 5.00
Begin VB.Form helping 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton back 
      BackColor       =   &H0080C0FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9120
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "helping.frx":0000
      Top             =   5280
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "helping.frx":0065
      Top             =   4800
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "helping.frx":00E5
      Top             =   4440
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "helping.frx":016D
      Top             =   4680
      Visible         =   0   'False
      Width           =   10215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "helping.frx":01FD
      Top             =   5040
      Visible         =   0   'False
      Width           =   10215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "helping.frx":0244
      Top             =   4680
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.Image Image2 
      Height          =   2625
      Left            =   10800
      Picture         =   "helping.frx":0296
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3900
   End
   Begin VB.Image Image1 
      Height          =   9840
      Left            =   120
      Picture         =   "helping.frx":9CAF
      Stretch         =   -1  'True
      Top             =   600
      Width           =   4815
   End
End
Attribute VB_Name = "helping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back_Click()
Unload Me
help.Show
End Sub

Private Sub Form_Load()
If st = "1" Then
Text1.Visible = True
ElseIf st = "2" Then
Text5.Visible = True
ElseIf st = "3" Then
Text3.Visible = True
ElseIf st = "4" Then
Text4.Visible = True
ElseIf st = "5" Then
Text2.Visible = True
ElseIf st = "6" Then
Text6.Visible = True
End If
End Sub


