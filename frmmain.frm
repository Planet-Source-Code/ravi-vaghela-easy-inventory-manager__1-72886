VERSION 5.00
Object = "{41BC0B27-5B30-4FD1-AE28-32EA831E20D4}#1.0#0"; "bse_button.ocx"
Begin VB.Form frmmain 
   BackColor       =   &H80000009&
   Caption         =   "Main Form"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   -150
   ClientWidth     =   15240
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmmain.frx":030A
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   2760
      Top             =   1920
   End
   Begin BSE_Engine.BSE BSE1 
      Left            =   5880
      Top             =   5760
      _ExtentX        =   6588
      _ExtentY        =   1085
   End
   Begin VB.CommandButton Command8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      DisabledPicture =   "frmmain.frx":520A4
      DragIcon        =   "frmmain.frx":523AE
      Height          =   615
      Left            =   0
      Picture         =   "frmmain.frx":526B8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H80000013&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8160
      Width           =   2055
   End
   Begin VB.CommandButton cmdcstock 
      BackColor       =   &H80000013&
      Caption         =   "&Current Stock"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdempmst 
      BackColor       =   &H80000013&
      Caption         =   "&Employe Master"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdnewitems 
      BackColor       =   &H80000013&
      Caption         =   "&New Items"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdtran 
      BackColor       =   &H80000013&
      Caption         =   "&Transaction"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmditemre 
      BackColor       =   &H80000013&
      Caption         =   "&Item Report"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton cmdlog 
      BackColor       =   &H80000013&
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   1800
      Picture         =   "frmmain.frx":529C2
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   13545
   End
   Begin VB.Image Image3 
      Height          =   1125
      Left            =   0
      Picture         =   "frmmain.frx":1FF884
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   1755
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Develope By Ravi Vaghela"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   11760
      TabIndex        =   9
      Top             =   8760
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   0
      Picture         =   "frmmain.frx":203C59
      Stretch         =   -1  'True
      Top             =   10080
      Width           =   15615
   End
   Begin VB.Image Image1 
      Height          =   6825
      Left            =   3720
      Picture         =   "frmmain.frx":203FF8
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   10065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EasY InventorY ManageR"
      BeginProperty Font 
         Name            =   "BankGothic Md BT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   750
      Left            =   4560
      TabIndex        =   7
      Top             =   600
      Width           =   10515
   End
   Begin VB.Image Image4 
      Height          =   1650
      Left            =   -240
      Picture         =   "frmmain.frx":266E17
      Top             =   0
      Width           =   15855
   End
   Begin VB.Menu mnutran 
      Caption         =   "Transaction"
      Visible         =   0   'False
      Begin VB.Menu mnuadd 
         Caption         =   "Add Stock"
      End
      Begin VB.Menu mnuissue 
         Caption         =   "Issue"
      End
      Begin VB.Menu mnureturn 
         Caption         =   "Return"
      End
      Begin VB.Menu mnudeffective 
         Caption         =   "Deffective"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "Report"
      Visible         =   0   'False
      Begin VB.Menu mnuitems 
         Caption         =   "Items Wise"
      End
   End
   Begin VB.Menu mnuempmst 
      Caption         =   "Employe Mater"
      Visible         =   0   'False
      Begin VB.Menu mnuempmast 
         Caption         =   "Employe Master"
      End
      Begin VB.Menu mnuempdetail 
         Caption         =   "Employe Detail"
      End
   End
   Begin VB.Menu mnulogin 
      Caption         =   "login"
      Visible         =   0   'False
      Begin VB.Menu mnunewuser 
         Caption         =   "New User"
      End
   End
   Begin VB.Menu MNUEXIT 
      Caption         =   "EXIT"
      Visible         =   0   'False
      Begin VB.Menu MNUETE 
         Caption         =   "Exit From EIM"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Cmdexit_Click()
PopupMenu MNUEXIT, 0, 1500, 8500

End Sub

Private Sub cmdcstock_Click()

On Error Resume Next
    frmcstock.Show , Me
    frmcstock.Left = 3000
    frmcstock.Top = 2030
    frmcstock.Width = Me.Width - frmcstock.Left
    frmcstock.Height = Image2.Top - frmcstock.Top
frmadd.Visible = False
frmdeffective.Visible = False
frmempdetail.Visible = False
frmempmst.Visible = False
frmitems.Visible = False
frmitemwise.Visible = False
frmnewuser.Visible = False
frmreturn.Visible = False
frmissue.Visible = False
frmmain.Caption = "Easy Inventory Manager - Current Stock"

End Sub

Private Sub cmdempmst_Click()
PopupMenu mnuempmst, 0, 1500, 3500
End Sub

Private Sub Cmdnewitems_Click()
frmitems.Show , Me
frmitems.Left = 3000
frmitems.Top = 2030
frmitems.Width = Me.Width - frmitems.Left
frmitems.Height = Image2.Top - frmitems.Top
frmadd.Visible = False
frmcstock.Visible = False
frmdeffective.Visible = False
frmempdetail.Visible = False
frmempmst.Visible = False
frmitemwise.Visible = False
frmnewuser.Visible = False
frmreturn.Visible = False
frmissue.Visible = False
frmmain.Caption = "Easy Inventory Manager - New Items"
End Sub

Private Sub Cmdtran_Click()
PopupMenu mnutran, 0, 1500, 5300
End Sub

Private Sub Cmditemre_Click()
PopupMenu mnureport, 0, 1500, 6300
End Sub

Private Sub Cmdlog_Click()
PopupMenu mnulogin, 0, 1500, 7400
End Sub

Private Sub Command8_Click()
Shell "c:\windows\system32\calc.exe"
End Sub

Private Sub Form_Load()
BSE1.InitSubClassing
frmmain.Caption = "Easy Inventory Manager - Main Form"
End Sub






Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
    result = MsgBox("Do you want to exit to Windows ?", vbQuestion + vbYesNo + vbDefaultButton2, "Want to Exit ?")
        If result = vbYes Then
         Unload Me
         frmSplash.Show
         
         
       Else
        Cancel = True
        End If
    End If
End Sub



Private Sub mnuadd_Click()
frmadd.Show , Me
frmadd.Left = 3000
frmadd.Top = 2030
frmadd.Width = Me.Width - frmadd.Left
frmadd.Height = Image2.Top - frmadd.Top
frmcstock.Visible = False
frmdeffective.Visible = False
frmempdetail.Visible = False
frmempmst.Visible = False
frmitems.Visible = False
frmitemwise.Visible = False
frmnewuser.Visible = False
frmreturn.Visible = False
frmissue.Visible = False
frmmain.Caption = "Easy Inventory Manager - Add / Receive Stock"

End Sub

Private Sub mnudeffective_Click()
frmdeffective.Show , Me
frmdeffective.Left = 3000
    frmdeffective.Top = 2030
frmdeffective.Width = Me.Width - frmdeffective.Left
frmdeffective.Height = Image2.Top - frmdeffective.Top
frmadd.Visible = False
frmempdetail.Visible = False
frmcstock.Visible = False
frmempmst.Visible = False
frmitems.Visible = False
frmitemwise.Visible = False
frmnewuser.Visible = False
frmreturn.Visible = False
frmissue.Visible = False
frmmain.Caption = "Easy Inventory Manager - Deffective Items"
End Sub

Private Sub mnuempdetail_Click()
frmempdetail.Show , Me
frmempdetail.Left = 3000
    frmempdetail.Top = 2030
    frmempdetail.Width = Me.Width - frmempdetail.Left
    frmempdetail.Height = Image2.Top - frmempdetail.Top
frmadd.Visible = False
frmdeffective.Visible = False
frmcstock.Visible = False
frmempmst.Visible = False
frmitems.Visible = False
frmitemwise.Visible = False
frmnewuser.Visible = False
frmreturn.Visible = False
frmissue.Visible = False
frmmain.Caption = "Easy Inventory Manager - Employee Detail"

End Sub

Private Sub mnuempmast_Click()
frmempmst.Show , Me
frmempmst.Left = 3000
    frmempmst.Top = 2030
    frmempmst.Width = Me.Width - frmempmst.Left
    frmempmst.Height = Image2.Top - frmempmst.Top
    Image1.Enabled = False
frmadd.Visible = False
frmdeffective.Visible = False
frmempdetail.Visible = False
frmcstock.Visible = False
frmitems.Visible = False
frmitemwise.Visible = False
frmnewuser.Visible = False
frmreturn.Visible = False
frmissue.Visible = False
frmmain.Caption = "Easy Inventory Manager - Employee Master "

End Sub


Private Sub MNUETE_Click()
If MsgBox("Do you want to exit to Windows ?", vbQuestion + vbYesNo + vbDefaultButton2, "Want to Exit ?") = vbYes Then
        Unload Me
        frmSplash.Show
  End If

End Sub

Private Sub mnuissue_Click()
frmissue.Show , Me
frmissue.Left = 3000
    frmissue.Top = 2030
    frmissue.Width = Me.Width - frmissue.Left
    frmissue.Height = Image2.Top - frmissue.Top
frmadd.Visible = False
frmdeffective.Visible = False
frmempdetail.Visible = False
frmempmst.Visible = False
frmitems.Visible = False
frmcstock.Visible = False
frmnewuser.Visible = False
frmreturn.Visible = False
frmitemwise.Visible = False

frmmain.Caption = "Easy Inventory Manager - Issue Itmes "
End Sub

Private Sub mnuitems_Click()
frmitemwise.Show , Me
frmitemwise.Left = 3000
    frmitemwise.Top = 2030
    frmitemwise.Width = Me.Width - frmitemwise.Left
    frmitemwise.Height = Image2.Top - frmitemwise.Top
frmadd.Visible = False
frmdeffective.Visible = False
frmcstock.Visible = False
frmempmst.Visible = False
frmitems.Visible = False
frmnewuser.Visible = False
frmreturn.Visible = False
frmissue.Visible = False
frmmain.Caption = "Easy Inventory Manager - ItemWise Report"
End Sub

Private Sub mnunewuser_Click()
frmnewuser.Show , Me
  frmnewuser.Left = 3000
    frmnewuser.Top = 2030
    frmnewuser.Width = Me.Width - frmnewuser.Left
    frmnewuser.Height = Image2.Top - frmnewuser.Top
    Image1.Enabled = False
frmadd.Visible = False
frmdeffective.Visible = False
frmempdetail.Visible = False
frmempmst.Visible = False
frmitems.Visible = False
frmitemwise.Visible = False
frmcstock.Visible = False
frmreturn.Visible = False
frmitemwise.Visible = False
frmmain.Caption = "Easy Inventory Manager - New User"
End Sub

Private Sub mnureturn_Click()
frmreturn.Show , Me
frmreturn.Left = 3000
   frmreturn.Top = 2030
   frmreturn.Width = Me.Width - frmreturn.Left
   frmreturn.Height = Image2.Top - frmreturn.Top
frmcstock.Visible = False
frmadd.Visible = False
frmdeffective.Visible = False
frmempdetail.Visible = False
frmempmst.Visible = False
frmitems.Visible = False
frmitemwise.Visible = False
frmnewuser.Visible = False

frmissue.Visible = False
frmmain.Caption = "Easy Inventory Manager - Return Items"

End Sub

Private Sub mnusold_Click()
frmissue.Show , Me
    frmissue.Left = 3000
    frmissue.Top = 2030
    frmissue.Width = Me.Width - frmissue.Left
    frmissue.Height = Image2.Top - frmissue.Top
frmadd.Visible = False
frmdeffective.Visible = False
frmempdetail.Visible = False
frmempmst.Visible = False
frmitems.Visible = False
frmitemwise.Visible = False
frmnewuser.Visible = False
frmreturn.Visible = False
frmcstock.Visible = False
frmmain.Caption = "Easy Inventory Manager - Issue Items "


End Sub


Private Sub Timer1_Timer()
If cmdcstock.Left >= 1080 Then
cmdcstock.Move cmdcstock.Left - 1500
End If
If cmdempmst.Left >= 1080 Then
cmdempmst.Move cmdempmst.Left - 1500
End If
If cmdnewitems.Left >= 1080 Then
cmdnewitems.Move cmdnewitems.Left - 1500
End If
If cmdtran.Left >= 1080 Then
cmdtran.Move cmdtran.Left - 1500
End If
If cmditemre.Left >= 1080 Then
cmditemre.Move cmditemre.Left - 1500
End If
If cmdlog.Left >= 1080 Then
cmdlog.Move cmdlog.Left - 1500
End If
If cmdexit.Left >= 1080 Then
cmdexit.Move cmdexit.Left - 1500
End If

End Sub
