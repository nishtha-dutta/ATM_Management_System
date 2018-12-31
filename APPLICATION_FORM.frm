VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form APPLICATION_FORM 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton back 
      BackColor       =   &H0080C0FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton delete 
      BackColor       =   &H0080C0FF&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton previous 
      BackColor       =   &H0080C0FF&
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton update 
      BackColor       =   &H0080C0FF&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton first 
      BackColor       =   &H0080C0FF&
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton last 
      BackColor       =   &H0080C0FF&
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton next 
      BackColor       =   &H0080C0FF&
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton add 
      BackColor       =   &H0080C0FF&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   9000
      TabIndex        =   6
      Top             =   3120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483638
      CalendarTitleBackColor=   -2147483636
      CalendarTrailingForeColor=   -2147483637
      Format          =   20709377
      CurrentDate     =   40268
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   9000
      TabIndex        =   5
      Top             =   3960
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   9000
      TabIndex        =   4
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOUNT NUMBER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   3960
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "SOLD DATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "APPLICATION FORM DETAILS "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   0
      Top             =   600
      Width           =   10575
   End
End
Attribute VB_Name = "APPLICATION_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset

Private Sub add_Click()
rs.Edit
rs.Fields(0) = Text2.Text
rs.Fields(3) = Text4.Text
rs.Fields(2) = DTPicker1.Value
rs.update
End Sub

Private Sub Delete_Click()
w = MsgBox("Are you sure to delete the record ??", vbQuestion + vbYesNo)
If w = vbYes Then
rs.delete
MsgBox "Record deleted successfully !!", vbInformation
Text2.Text = ""
Text4.Text = ""
DTPicker1.Value = 6 / 9 / 1990
End If
End Sub

Private Sub first_Click()
rs.MoveFirst
Text2.Text = rs.Fields(0)
Text4.Text = rs.Fields(3)
DTPicker1.Value = rs.Fields(2)
End Sub

Private Sub next_Click()
If rs.EOF = True Then
rs.MoveLast
l:
MsgBox "No more records !!", vbCritical
Else
rs.MoveNext
If rs.EOF = True Then
GoTo l
End If
Text2.Text = rs.Fields(0)
Text4.Text = rs.Fields(3)
DTPicker1.Value = rs.Fields(2)
End If
End Sub

Private Sub last_Click()
rs.MoveLast
Text2.Text = rs.Fields(0)
Text4.Text = rs.Fields(3)
DTPicker1.Value = rs.Fields(2)
End Sub

Private Sub previous_Click()
If rs.BOF = True Then
rs.MoveFirst
l:
MsgBox "No more records !!", vbCritical
Else
rs.MovePrevious
If rs.BOF = True Then
GoTo l
End If
Text2.Text = rs.Fields(0)
Text4.Text = rs.Fields(3)
DTPicker1.Value = rs.Fields(2)
End If
End Sub

Private Sub Command7_Click()
 rs.Edit
   rs.delete
 rs.update
End Sub

Private Sub back_Click()
Unload Me
existing.Show
End Sub



Private Sub Text1_KeyPress(keyascii As Integer)
If keyascii <> 8 And (keyascii < 48 Or keyascii > 57) Then
MsgBox "PLZ ENTER DIGIT 0-9", vbCritical
keyascii = 0
End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("atm.mdb")
Set rs = db.OpenRecordset("application_form")
End Sub

Private Sub Text2_KeyPress(keyascii As Integer)
If keyascii <> 8 And keyascii <> 32 And (keyascii < 65 Or keyascii > 90) And (keyascii < 97 Or keyascii > 122) Then
MsgBox "PLZ ENTER ALPHABETS", vbCritical
keyascii = 0
End If
End Sub



Private Sub Text4_KeyPress(keyascii As Integer)
If keyascii <> 8 And (keyascii < 48 Or keyascii > 57) Then
MsgBox "PLZ ENTER DIGIT 0-9", vbCritical
keyascii = 0
End If
End Sub

Private Sub Text5_keypress(keyascii As Integer)
If keyascii <> 8 And (keyascii < 48 Or keyascii > 57) Then
MsgBox "PLZ ENTER DIGIT 0-9", vbCritical
keyascii = 0
End If
End Sub

Private Sub update_Click()
rs.MoveFirst
Do While Not rs.EOF
If Text4.Text = rs.Fields(3) Then
rs.Edit
rs.Fields(0) = Text2.Text
rs.Fields(2) = DTPicker1.Value
rs.update
Exit Do
Else
rs.MoveNext
End If
Loop
End Sub
