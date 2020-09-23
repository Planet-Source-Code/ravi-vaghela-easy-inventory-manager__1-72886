VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Developer"
   ClientHeight    =   10035
   ClientLeft      =   2535
   ClientTop       =   390
   ClientWidth     =   11115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   5
      Left            =   5400
      Top             =   8160
   End
   Begin VB.Timer Timer1 
      Interval        =   2800
      Left            =   4320
      Top             =   8160
   End
   Begin VB.Image Image3 
      Height          =   2580
      Left            =   6120
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2835
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   4920
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label contact 
      BackStyle       =   0  'Transparent
      Caption         =   " Contact: ravirnv@yahoo.com"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   10800
      TabIndex        =   12
      Top             =   5160
      Width           =   6255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "=>"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "=>"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "=>"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "=>"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "=>"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EasY InventorY ManageR"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   750
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   10515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Developer:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label ravi 
      BackStyle       =   0  'Transparent
      Caption         =   "Ravi Vaghela"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   10920
      TabIndex        =   4
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label reg 
      BackStyle       =   0  'Transparent
      Caption         =   "S-156105103"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   10920
      TabIndex        =   3
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label diploma 
      BackStyle       =   0  'Transparent
      Caption         =   " Diploma 5th Semester"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   10800
      TabIndex        =   2
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label aits 
      BackStyle       =   0  'Transparent
      Caption         =   "A.I.T.S.D.S. - Rajkot"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   10920
      TabIndex        =   1
      Top             =   4320
      Width           =   4935
   End
   Begin VB.Image Img 
      Height          =   3615
      Left            =   480
      Picture         =   "frmSplash.frx":1400A
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   3105
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "in Computer Engg."
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   11040
      TabIndex        =   0
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Image Image2 
      Height          =   11520
      Left            =   0
      Picture         =   "frmSplash.frx":1B6C5
      Top             =   -1080
      Width           =   15360
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
End
End Sub
Private Sub Timer2_Timer()
If ravi.Left >= 1050 Then
ravi.Move ravi.Left - 600
End If

If reg.Left >= 1050 Then
reg.Move reg.Left - 600
End If

If diploma.Left >= 1050 Then
diploma.Move diploma.Left - 600
End If

If Label3.Left >= 1500 Then
Label3.Move Label3.Left - 500
End If

If aits.Left >= 1050 Then
aits.Move aits.Left - 600
End If

If contact.Left >= 1050 Then
contact.Move contact.Left - 600
End If



End Sub


