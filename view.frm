VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form view 
   BackColor       =   &H00C0E0FF&
   Caption         =   "CUSTOMER DETAILS"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton ADD 
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Delete record"
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton First 
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "First Record"
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton Prev 
      BackColor       =   &H0080C0FF&
      Caption         =   "PREV"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Previous Record"
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton Update 
      BackColor       =   &H0080C0FF&
      Caption         =   "UPADTE"
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Update changes"
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton Delete 
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Delete record"
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton next_button 
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Next Record"
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton Last 
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
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Last Record"
      Top             =   8640
      Width           =   1935
   End
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Back"
      Top             =   10200
      Width           =   1935
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
      Left            =   10320
      TabIndex        =   10
      Top             =   6480
      Width           =   3015
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
      Left            =   10320
      MaxLength       =   10
      TabIndex        =   9
      Top             =   5760
      Width           =   3015
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "FEMALE"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11640
      TabIndex        =   8
      Top             =   4680
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "MALE"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10320
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
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
      Left            =   10320
      TabIndex        =   6
      Top             =   4080
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
      Left            =   10320
      TabIndex        =   5
      Top             =   7920
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
      Left            =   10320
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3360
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
      Left            =   10320
      MaxLength       =   14
      TabIndex        =   3
      Top             =   2640
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
      Left            =   10320
      MaxLength       =   14
      TabIndex        =   2
      Top             =   1920
      Width           =   3015
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
      Left            =   10320
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   10320
      TabIndex        =   0
      Top             =   5160
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
      Format          =   20643841
      CurrentDate     =   40264
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   10320
      TabIndex        =   23
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
      Format          =   20643841
      CurrentDate     =   40264
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "EXPIRY DATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6720
      TabIndex        =   22
      Top             =   7080
      Width           =   3255
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7080
      TabIndex        =   21
      Top             =   6360
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6720
      TabIndex        =   20
      Top             =   5640
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7200
      TabIndex        =   19
      Top             =   5040
      Width           =   2895
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   6600
      TabIndex        =   18
      Top             =   4440
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7200
      TabIndex        =   17
      Top             =   3840
      Width           =   1095
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6720
      TabIndex        =   16
      Top             =   7800
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6600
      TabIndex        =   15
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6360
      TabIndex        =   14
      Top             =   2520
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6960
      TabIndex        =   13
      Top             =   1920
      Width           =   2895
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6600
      TabIndex        =   12
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   4200
      TabIndex        =   11
      Top             =   -120
      Width           =   11775
   End
End
Attribute VB_Name = "view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Private Sub add_Click()
rs.MoveFirst
rs.Edit
rs.Fields(1) = Text2.Text
rs.Fields(3) = Text4.Text
rs.Fields(5) = Text5.Text
rs.Fields(6) = Text7.Text
rs.Fields(9) = Text8.Text
rs.Fields(0) = Text9.Text
rs.Fields(2) = Text3.Text
rs.Fields(8) = Text6.Text
If Option1.Value = True Then
rs.Fields(7) = "M"
Else
rs.Fields(7) = "F"
End If
rs.update
rs1.Edit
If rs.Fields(0) = rs1.Fields(5) Then
rs1.Fields(3) = DTPicker2.Value
End If
End Sub

Private Sub back_Click()
Unload Me
existing.Show
End Sub

Private Sub Delete_Click()
w = MsgBox("Are you sure to delete the record ??", vbQuestion + vbYesNo)
If w = vbYes Then
rs.delete
MsgBox "Record deleted successfully !!", vbInformation
Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text3.Text = ""
Text6.Text = ""
Option1.Value = False
Option2.Value = False
DTPicker2.Value = 6 / 9 / 1990
End If
End Sub

Private Sub first_Click()
rs.MoveFirst
Text2.Text = rs.Fields(1)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(5)
Text7.Text = rs.Fields(6)
Text8.Text = rs.Fields(9)
Text9.Text = rs.Fields(0)
Text3.Text = rs.Fields(2)
Text6.Text = rs.Fields(8)
If rs.Fields(7) = "M" Then
Option1.Value = True
Else
Option2.Value = True
End If
rs1.MoveFirst
While Not rs1.EOF
If rs.Fields(0) = rs1.Fields(5) Then
DTPicker2.Value = rs1.Fields(3)
End If
rs1.MoveNext
Wend
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("atm.mdb")
Set rs = db.OpenRecordset("customer")
Set rs1 = db.OpenRecordset("card_details")
If rs.RecordCount = 0 Then
MsgBox " Table is empty", vbCritical
Exit Sub
ElseIf rs.EOF Then
MsgBox "No such record", vbCritical
Exit Sub
Else
rs.MoveFirst
While Not rs.EOF
If rs.Fields(1) = ac Then
Text2.Text = rs.Fields(1)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(5)
Text7.Text = rs.Fields(6)
Text8.Text = rs.Fields(9)
Text9.Text = rs.Fields(0)
Text3.Text = rs.Fields(2)
Text6.Text = rs.Fields(8)
If rs.Fields(7) = "M" Then
Option1.Value = True
Else
Option2.Value = True
End If
rs1.MoveFirst
While Not rs1.EOF
If rs.Fields(0) = rs1.Fields(5) Then
DTPicker2.Value = rs1.Fields(3)
End If
rs1.MoveNext
Wend
Exit Sub
End If
rs.MoveNext
Wend
End If
End Sub

Private Sub last_Click()
rs.MoveLast
Text2.Text = rs.Fields(1)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(5)
Text7.Text = rs.Fields(6)
Text8.Text = rs.Fields(9)
Text9.Text = rs.Fields(0)
Text3.Text = rs.Fields(2)
Text6.Text = rs.Fields(8)
If rs.Fields(7) = "M" Then
Option1.Value = True
Else
Option2.Value = True
End If
While Not rs1.EOF
If rs.Fields(0) = rs1.Fields(5) Then
DTPicker2.Value = rs1.Fields(3)
End If
rs1.MoveNext
Wend
End Sub

Private Sub next_button_Click()
If rs.EOF = True Then
rs.MoveLast
l:
MsgBox "No more records !!", vbCritical
Else
rs.MoveNext
If rs.EOF = True Then
GoTo l
End If
Text2.Text = rs.Fields(1)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(5)
Text7.Text = rs.Fields(6)
Text8.Text = rs.Fields(9)
Text9.Text = rs.Fields(0)
Text3.Text = rs.Fields(2)
Text6.Text = rs.Fields(8)
If rs.Fields(7) = "M" Then
Option1.Value = True
Else
Option2.Value = True
End If
rs1.MoveFirst
While Not rs1.EOF
If rs.Fields(0) = rs1.Fields(5) Then
DTPicker2.Value = rs1.Fields(3)
End If
rs1.MoveNext
Wend
End If
End Sub

Private Sub Prev_Click()
If rs.BOF = True Then
rs.MoveFirst
l:
MsgBox "No more records !!", vbCritical
Else
rs.MovePrevious
If rs.BOF = True Then
GoTo l
End If
Text2.Text = rs.Fields(1)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(5)
Text7.Text = rs.Fields(6)
Text8.Text = rs.Fields(9)
Text9.Text = rs.Fields(0)
Text3.Text = rs.Fields(2)
Text6.Text = rs.Fields(8)
If rs.Fields(7) = "M" Then
Option1.Value = True
Else
Option2.Value = True
End If
rs1.MoveFirst
While Not rs1.EOF
If rs.Fields(0) = rs1.Fields(5) Then
DTPicker2.Value = rs1.Fields(3)
End If
rs1.MoveNext
Wend
End If
End Sub

Private Sub update_Click()
rs.MoveFirst
Do While Not rs.EOF
If Text2.Text = rs.Fields(3) Then
rs.Edit
rs.Fields(1) = Text2.Text
rs.Fields(3) = Text4.Text
rs.Fields(5) = Text5.Text
rs.Fields(6) = Text7.Text
rs.Fields(9) = Text8.Text
rs.Fields(0) = Text9.Text
rs.Fields(2) = Text3.Text
rs.Fields(8) = Text6.Text
If Option1.Value = True Then
rs.Fields(7) = "M"
Else
rs.Fields(7) = "F"
End If
rs.update
Exit Do
Else
rs.MoveNext
End If
Loop
End Sub
