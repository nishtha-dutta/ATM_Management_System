VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2850
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1683.874
   ScaleMode       =   0  'User
   ScaleWidth      =   5197.065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   1800
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "enter user ID"
      Top             =   480
      Width           =   3045
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080C0FF&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080C0FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "enter your password"
      Top             =   1200
      Width           =   3045
   End
   Begin VB.Label C 
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Line Line2 
      X1              =   4732.287
      X2              =   5182.98
      Y1              =   1417.999
      Y2              =   1417.999
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ATM Card Management system"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4732.287
      Y1              =   1417.999
      Y2              =   1417.999
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   1440
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   1320
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub C_Click()
Unload Me
input_form.Show
End Sub

Private Sub cmdCancel_Click()
Unload Me
thank.Show
End Sub

Private Sub cmdOK_Click()
If txtPassword = rs.Fields(0) And txtUserName = rs.Fields(1) Then
    mains.Show
    Unload Me
Else
    MsgBox "Invalid Password or user name, try again!", , "Login"
    txtPassword.Text = ""
    txtUserName.Text = ""
End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("atm.mdb")
Set rs = db.OpenRecordset("login")
End Sub

Private Sub Timer1_Timer()
If Label1.Left = 0 Then Label1.Left = 2880
Label1.Left = Label1.Left - 60
End Sub
