VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form APPLICATION 
   BackColor       =   &H00C0E0FF&
   Caption         =   "APPLICATION FORM"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14700
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   14700
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   4680
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   10200
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   9600
      TabIndex        =   25
      Top             =   7320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20709377
      CurrentDate     =   40264
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   11
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      MaxLength       =   14
      TabIndex        =   10
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      MaxLength       =   14
      TabIndex        =   9
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   5280
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   7
      Top             =   9360
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   6
      Top             =   6120
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "MALE"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "FEMALE"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   4
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      MaxLength       =   10
      TabIndex        =   3
      Top             =   7920
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   2
      Top             =   8640
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10200
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10200
      Width           =   2295
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "LAST NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   27
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Please Fill Following Details :-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   -240
      TabIndex        =   24
      Top             =   1800
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " A/C TYPE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   23
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "A/C NUMBER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   22
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FIRST NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   21
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   20
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "E-MAIL ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      Top             =   9360
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "CITY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   18
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "GENDER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5760
      TabIndex        =   17
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF BIRTH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   16
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PHONE NO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   15
      Top             =   8040
      Width           =   2895
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CARD NUMBER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   14
      Top             =   8760
      Width           =   2895
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   13
      Top             =   8640
      Width           =   2295
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "APPLICATION FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   4320
      TabIndex        =   12
      Top             =   240
      Width           =   9135
   End
End
Attribute VB_Name = "APPLICATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Private Sub Command1_Click()
Dim n As String
If Option1.Value = True Then
n = "M"
Else
n = "F"
End If
Set db = OpenDatabase("atm.mdb")
Set rs = db.OpenRecordset("customer")
Set rs1 = db.OpenRecordset("card_details")
Set rs2 = db.OpenRecordset("application_form")
rs.Edit
rs.Fields(0).Value = Text9.Text
rs.Fields(1).Value = Text2.Text
rs.Fields(2).Value = Text3.Text
rs.Fields(3).Value = Text4.Text
rs.Fields(4).Value = Text1.Text
rs.Fields(5).Value = Text5.Text
rs.Fields(6).Value = Text7.Text
rs.Fields(7).Value = n
rs.Fields(8).Value = Text6.Text
rs.Fields(9).Value = Text8.Text
rs.Fields(10).Value = DTPicker1.Value
rs.update
MsgBox "Form accepted", vbInformation
rs1.Edit
rs1.Fields(4).Value = VBA.Date
rs1.Fields(6).Value = VBA.Date + 31
rs1.Fields(3).Value = VBA.Date + 7336
rs1.Fields(5) = Text9.Text
rs1.update
rs2.Edit
rs.Fields(2).Value = VBA.Date
rs.Fields(0).Value = "sold"
rs2.Fields(3).Value = Text2.Text
rs2.update
Unload Me
main.Show
End Sub

Private Sub Command2_Click()
Unload Me
main.Show
End Sub

Private Sub Command3_Click()
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text6.Text = " "
Text8.Text = " "
Text9.Text = " "
Text5.Text = " "
Text7.Text = " "
Option1.Value = False
Option2.Value = False
End Sub

Private Sub Text1_KeyPress(keyascii As Integer)
If keyascii <> 8 And keyascii <> 32 And (keyascii < 65 Or keyascii > 90) And (keyascii < 97 Or keyascii > 122) Then
MsgBox "PLZ ENTER ALPHABETS", vbCritical
keyascii = 0
End If
End Sub

Private Sub Text2_KeyPress(keyascii As Integer)
If keyascii <> 8 And (keyascii < 48 Or keyascii > 57) Then
MsgBox "PLZ ENTER DIGIT 0-9", vbCritical
keyascii = 0
End If
End Sub

Private Sub Text3_KeyPress(keyascii As Integer)
If keyascii <> 8 And keyascii <> 32 And (keyascii < 65 Or keyascii > 90) And (keyascii < 97 Or keyascii > 122) Then
MsgBox "PLZ ENTER ALPHABETS", vbCritical
keyascii = 0
End If
End Sub

Private Sub Text4_KeyPress(keyascii As Integer)
If keyascii <> 8 And keyascii <> 32 And (keyascii < 65 Or keyascii > 90) And (keyascii < 97 Or keyascii > 122) Then
MsgBox "PLZ ENTER ALPHABETS", vbCritical
keyascii = 0
End If
End Sub

Private Sub Text7_KeyPress(keyascii As Integer)
If keyascii <> 8 And keyascii <> 32 And (keyascii < 65 Or keyascii > 90) And (keyascii < 97 Or keyascii > 122) Then
MsgBox "PLZ ENTER ALPHABETS", vbCritical
keyascii = 0
End If
End Sub

Private Sub Text8_KeyPress(keyascii As Integer)
If keyascii <> 8 And (keyascii < 48 Or keyascii > 57) Then
MsgBox "PLZ ENTER DIGIT 0-9", vbCritical
keyascii = 0
End If
End Sub

Private Sub Text9_Change()
If keyascii <> 8 And (keyascii < 48 Or keyascii > 57) Then
MsgBox "PLZ ENTER DIGIT 0-9", vbCritical
keyascii = 0
End If
End Sub
