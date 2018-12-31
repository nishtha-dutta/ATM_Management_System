VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7080
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   6795
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9585
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   4200
         Top             =   3000
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   6120
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image Image2 
         Height          =   2265
         Left            =   2760
         Picture         =   "frmSplash.frx":000C
         Top             =   360
         Width           =   6690
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading 0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8160
         TabIndex        =   7
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Connecting to database..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   6
         Top             =   5880
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   6375
         Left            =   0
         Picture         =   "frmSplash.frx":328A
         Top             =   120
         Width           =   2835
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "© Copyright 2010-13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   2
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Warning : No part of this product may be produced without permission. All rights reserved."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   480
         TabIndex        =   1
         Top             =   6480
         Width           =   9015
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   3
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Platform : Visual basic 6.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         TabIndex        =   4
         Top             =   4800
         Width           =   3900
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(keyascii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value > 99 Then
Unload Me
frmLogin.Show
Else
If Label1.Visible = True Then
Label1.Visible = False
Else
Label1.Visible = True
End If
Label2.Caption = "Loading" + Str(ProgressBar1.Value) + "%"
ProgressBar1.Value = ProgressBar1.Value + 1
End If
End Sub
