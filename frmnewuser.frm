VERSION 5.00
Object = "{41BC0B27-5B30-4FD1-AE28-32EA831E20D4}#1.0#0"; "bse_button.ocx"
Begin VB.Form frmnewuser 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   3735
   ClientTop       =   2730
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   Picture         =   "frmnewuser.frx":0000
   ScaleHeight     =   8430
   ScaleWidth      =   12465
   ShowInTaskbar   =   0   'False
   Begin BSE_Engine.BSE BSE1 
      Left            =   0
      Top             =   1080
      _ExtentX        =   6588
      _ExtentY        =   1085
      SchemeStyle     =   14
   End
   Begin VB.TextBox txtconfpass 
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
      Left            =   6720
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4560
      Width           =   2805
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
      Left            =   6720
      TabIndex        =   0
      Top             =   3240
      Width           =   2805
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H000080FF&
      Caption         =   "OK"
      Default         =   -1  'True
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
      Left            =   4920
      TabIndex        =   3
      Top             =   5640
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000080FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   6600
      MaskColor       =   &H0080FFFF&
      TabIndex        =   4
      Top             =   5640
      Width           =   2100
   End
   Begin VB.TextBox txtPass 
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
      Left            =   6720
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3840
      Width           =   2805
   End
   Begin VB.Image ImgClose 
      Height          =   540
      Left            =   9600
      Picture         =   "frmnewuser.frx":51D9A
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   540
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   2640
      Picture         =   "frmnewuser.frx":528A6
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NEW USER ENTRY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   600
      Left            =   4080
      TabIndex        =   8
      Top             =   1920
      Width           =   4755
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm PassworD:"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   2400
      TabIndex        =   7
      Top             =   4560
      Width           =   4200
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "New UseR:"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Top             =   3240
      Width           =   2640
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "PassworD:"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   3840
      Width           =   2400
   End
End
Attribute VB_Name = "frmnewuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
 Dim x As String


Public Sub clear()
txtPass.Text = ""
txtconfpass.Text = ""
txtUserName.Text = ""
End Sub


Private Sub cmdCancel_Click()
frmmain.Caption = "Easy Inventory Manager - Main Form"
Unload Me
End Sub

Private Sub cmdOK_Click()
If UCase(txtPass.Text) = UCase(txtconfpass.Text) Then
rs1.Open "select* from log where user='" & txtUserName.Text & "'", con, adOpenDynamic, adLockOptimistic
        If rs1.EOF = True Or rs1.BOF = True Then
            rs1.AddNew
            rs1.Fields("user").Value = UCase(txtUserName.Text)
            rs1.Fields("Pass").Value = UCase(txtPass.Text)
            rs1.Update
            MsgBox "New user Added"
            
            Unload Me
            
        Else
            MsgBox "User Already exists"
            Call clear
            txtUserName.SetFocus
        End If
        rs1.Close
        Set rs1 = Nothing
Else
    MsgBox "Password and confirm password Does Not Match"
    Call clear
End If
End Sub

Private Sub Form_Load()
BSE1.InitSubClassing
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Invdata.mdb;Persist Security Info=False"

End Sub


Private Sub ImgClose_Click()
Unload Me
End Sub
