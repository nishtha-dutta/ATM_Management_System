VERSION 5.00
Begin VB.Form input_form 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CHANGE PASSWORD"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   465
      Left            =   3480
      TabIndex        =   9
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   465
      Left            =   3480
      TabIndex        =   8
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Left            =   3480
      TabIndex        =   7
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton cancel 
      BackColor       =   &H0080C0FF&
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton ok 
      BackColor       =   &H0080C0FF&
      Caption         =   "OK"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox input_text 
      Height          =   465
      Left            =   3480
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Confirm password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "New password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Old password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "input_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub cancel_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("atm.mdb")
Set rs = db.OpenRecordset("login")
End Sub

Private Sub ok_Click()
If rs.Fields(0) <> Text2.Text And rs.Fields(1) <> Text3.Text Then
MsgBox "Invalid username or password", vbCritical
Exit Sub
ElseIf Text1.Text <> input_text Then
MsgBox "Your new password and confirm password doesnot match", vbCritical
Exit Sub
ElseIf Text1.Text = "" And input_text = "" Then
MsgBox "Please enter the new password and cofirm password", vbCritical
Exit Sub
Else
rs.Edit
rs.Fields(0) = Text1.Text
rs.Update
MsgBox "Your password has successfully changed", vbInformation
End If
End Sub
