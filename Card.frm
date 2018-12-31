VERSION 5.00
Begin VB.Form Card 
   BackColor       =   &H00C0E0FF&
   Caption         =   "CARD MANAGEMENT"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10110
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter Following details :-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   1
      Top             =   2880
      Width           =   10215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      TabIndex        =   0
      Top             =   1200
      Width           =   6855
   End
End
Attribute VB_Name = "Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim s As String

Private Sub Command1_Click()
Dim f As Boolean
If st = "1" Then
rs1.MoveFirst
While Not rs1.EOF
If rs1.Fields(1) = Text1.Text Then
s = rs1.Fields(0)
End If
rs1.MoveNext
Wend
rs.MoveFirst
While Not rs.EOF
If s = rs.Fields(5) Then
If rs.Fields(3) - VBA.Date = 0 Then
rs.Edit
rs.Fields(3) = rs.Fields(3) + 7300
rs.update
MsgBox "Record updated", vbInformation
Else
MsgBox "Your validity has not expired", vbCritical
End If
End If
rs.MoveNext
Wend
ElseIf st = "5" Then
ac = Text1.Text
view.Show
Unload Me
ElseIf st = "2" Then
f = False
rs1.MoveFirst
While Not rs1.EOF
If rs1.Fields(1).Value = Text1.Text Then
s = rs1.Fields(0).Value
End If
rs1.MoveNext
Wend
rs.MoveFirst
While Not rs.EOF
If s = rs.Fields(5) And rs.Fields(1) = "Y" Then
    rs.Edit
    rs.Fields(2) = VBA.Date
    rs.Fields(1) = "N"
    rs.update
    f = True
    MsgBox "Card reissued", vbInformation
    Exit Sub
End If
rs.MoveNext
Wend
If f = False Then
MsgBox "Your card cannot be reissued as it is not blocked", vbCritical
End If

ElseIf st = "3" Then
rs1.MoveFirst
While Not rs1.EOF
If rs1.Fields(1).Value = Text1.Text Then
s = rs1.Fields(0).Value
End If
rs1.MoveNext
Wend
f = False
rs.MoveFirst
Dim w As String
While Not rs.EOF
w = Format$(rs.Fields(0), "mm/dd/yyyy")
If s = rs.Fields(5).Value And w = "" Then
    rs.Edit
    rs.Fields(0).Value = VBA.Date
    rs.update
    MsgBox "Card cancelled successfully", vbInformation
    f = True
End If
rs.MoveNext
Wend
If f = False Then
MsgBox "card is already cancelled", vbCritical
Exit Sub
End If

ElseIf st = "4" Then
f = False
rs1.MoveFirst
Do While Not rs1.EOF
If rs1.Fields(1).Value = Text1.Text Then
s = rs1.Fields(0).Value
Exit Do
End If
rs1.MoveNext
Loop
rs.MoveFirst
Do While Not rs.EOF
If s = rs.Fields(5).Value Then
If rs.Fields(1).Value = "Y" Then
   f = True
   MsgBox "Card is already blocked", vbCritical
   Exit Sub
Else
rs.Edit
   rs.Fields(1).Value = "Y"
   rs.update
   MsgBox "Card blocked sucessfully", vbInformation
End If
End If
rs.MoveNext
Loop
 End If
End Sub

Private Sub Command2_Click()
Text1.Text = " "
End Sub

Private Sub Command3_Click()
st = ""
existing.Show
Unload Me
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("atm.mdb")
Set rs = db.OpenRecordset("card_details")
Set rs1 = db.OpenRecordset("customer")

If st = "1" Then
Label1.Caption = "RENEW CARD"
ElseIf st = "2" Then
Label1.Caption = "REISSUE CARD"
ElseIf st = "3" Then
Label1.Caption = "CARD CANCELLATION"
ElseIf st = "4" Then
Label1.Caption = "CARD BLOCKING"
ElseIf st = "5" Then
Label1.Caption = " CARD DETAILS"
End If
End Sub

Private Sub Text1_LostFocus()
Dim flag As Boolean
flag = False
rs1.MoveFirst
Do While Not rs1.EOF
If Text1.Text = rs1.Fields(1) Then
flag = True
Exit Do
End If
rs1.MoveNext
Loop
If flag = False Then
MsgBox "account number doesnot exist", vbCritical
End If
End Sub
