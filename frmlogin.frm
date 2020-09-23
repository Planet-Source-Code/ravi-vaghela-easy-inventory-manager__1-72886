VERSION 5.00
Object = "{41BC0B27-5B30-4FD1-AE28-32EA831E20D4}#1.0#0"; "bse_button.ocx"
Begin VB.Form frmlogin 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "welcome"
   ClientHeight    =   6045
   ClientLeft      =   5115
   ClientTop       =   3120
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   Picture         =   "frmlogin.frx":0000
   ScaleHeight     =   6045
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BSE_Engine.BSE BSE1 
      Left            =   2280
      Top             =   5280
      _ExtentX        =   6588
      _ExtentY        =   1085
      SchemeStyle     =   14
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "CanceL"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4680
      TabIndex        =   8
      Top             =   4320
      Width           =   2220
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3360
      Width           =   2805
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3000
      TabIndex        =   2
      Top             =   4320
      Width           =   1380
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4560
      TabIndex        =   0
      Top             =   2760
      Width           =   2805
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "EasY InventorY ManageR"
      BeginProperty Font 
         Name            =   "Alexis Italic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   3
      Left            =   1680
      TabIndex        =   7
      Top             =   1920
      Width           =   6960
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   6
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   3360
      Width           =   2400
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   4
      Top             =   2760
      Width           =   2640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "WelcomE"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1005
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   4860
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset





Private Sub cmdOK_Click()
Dim u, p
u = txtUserName.Text
p = txtPassword.Text
        x = "Select *from log where user='" & u & "' and pass='" & p & "'"
        rs.Open x, con, adOpenDynamic, adLockOptimistic
        If rs.EOF = True Then
            MsgBox "Invalid User Name or Password"
        Else
            Unload Me
            frmmain.Show
        End If
        rs.Close
        Set rs = Nothing

End Sub

Private Sub Command1_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
BSE1.InitSubClassing
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Invdata.mdb;Persist Security Info=False"
End Sub



