VERSION 5.00
Begin VB.Form frmSplash1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7365
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash1.frx":000C
   ScaleHeight     =   7365
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   480
      Top             =   6480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Easy Inventory Manager"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   975
      Left            =   1440
      TabIndex        =   7
      Top             =   1800
      Width           =   8655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   735
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   4080
      Width           =   15
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warning :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   840
      TabIndex        =   2
      Top             =   6240
      Width           =   690
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash1.frx":498D9
      ForeColor       =   &H8000000E&
      Height          =   555
      Left            =   1080
      TabIndex        =   1
      Top             =   6600
      Width           =   10305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait While Login Form Is Loading"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   735
      Left            =   2040
      TabIndex        =   0
      Top             =   5760
      Width           =   9495
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i%
Dim counter As Integer




Private Sub Timer1_Timer()
If i < 1000 Then
    i = i + 1
    Label12.Caption = "Loading " & i & " % "
    Label13.Width = i * 50
 Else
 Label12.Caption = "Loaded succsfully"
 Timer1.Interval = 0
 End If
 counter = counter + 1
If counter >= 100 Then
Timer1.Interval = 0
Unload Me
frmlogin.Show
End If
End Sub

