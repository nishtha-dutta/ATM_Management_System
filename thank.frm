VERSION 5.00
Begin VB.Form thank 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7485
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame thank 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000010&
      Height          =   7155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.Timer Timer1 
         Interval        =   80
         Left            =   4560
         Top             =   360
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Documentation"
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
         Left            =   5880
         TabIndex        =   10
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Documentation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Designing And Database"
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
         Left            =   6240
         TabIndex        =   8
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Designing And Coding"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "              MR. MUDASSIR ISLAM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   6
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Image Image2 
         Height          =   1425
         Left            =   4680
         Picture         =   "thank.frx":0000
         Top             =   5520
         Width           =   5220
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GUIDED BY:-    VD SIR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Index           =   8
         Left            =   120
         TabIndex        =   5
         Top             =   5880
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " MR. ARSHI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   4800
         TabIndex        =   4
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MR. MANISH RAI "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MISS NISHTHA DUTTA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEVELOPERS :-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "thank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Integer

Private Sub Timer1_Timer()
If Image2.Visible = True Then
Image2.Visible = False
Else
Image2.Visible = True
End If
If A = 50 Then
Unload Me
Else
A = A + 1
End If
End Sub
