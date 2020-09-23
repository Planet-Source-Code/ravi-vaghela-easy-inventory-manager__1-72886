VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{41BC0B27-5B30-4FD1-AE28-32EA831E20D4}#1.0#0"; "bse_button.ocx"
Begin VB.Form frmdeffective 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   3690
   ClientTop       =   1155
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   Picture         =   "frmdeffective.frx":0000
   ScaleHeight     =   8625
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Begin BSE_Engine.BSE BSE1 
      Left            =   1800
      Top             =   6960
      _ExtentX        =   6588
      _ExtentY        =   1085
   End
   Begin VB.CommandButton Cmdprint 
      BackColor       =   &H80000013&
      Caption         =   "&Print"
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
   End
   Begin VB.ComboBox cmbDadIName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   0
      Top             =   3120
      Width           =   3855
   End
   Begin VB.ComboBox cmbISize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   1
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox txtDadIQty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   355
      Left            =   3840
      TabIndex        =   2
      Top             =   4080
      Width           =   3855
   End
   Begin VB.ComboBox cmbDadIBy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   3
      Top             =   4560
      Width           =   3855
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000013&
      Caption         =   "&Close"
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H80000013&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000013&
      Caption         =   "&Save"
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H80000013&
      Caption         =   "&New"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txtSrNo 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   19
      Text            =   "0"
      Top             =   2640
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DadDate 
      Height          =   330
      Left            =   3840
      TabIndex        =   9
      Top             =   5040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   39274
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   2160
      Picture         =   "frmdeffective.frx":51D9A
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEFECTIVE STOCK ENTRY"
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
      Left            =   3360
      TabIndex        =   18
      Top             =   1440
      Width           =   6975
   End
   Begin VB.Label lblCStock 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      TabIndex        =   17
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Return Stock"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   16
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dad Item Name :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   15
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dad Item Qty :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   14
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Size :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   13
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dad  By :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   12
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dad Date :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Trans. No :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image ImgClose 
      Height          =   540
      Left            =   10320
      Picture         =   "frmdeffective.frx":528B0
      Stretch         =   -1  'True
      Top             =   840
      Width           =   540
   End
End
Attribute VB_Name = "frmdeffective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmbDadIBy_LostFocus()
On Error Resume Next
If cmbDadIBy = "" Then
    Exit Sub
Else
    CheckData "EmpMaster", "EmpName", cmbDadIBy.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Name , " & cmbDadIBy.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbDadIBy.Text = ""
        cmbDadIBy.SetFocus
    'Else
    '    RcvDate.SetFocus
    End If
End If
End Sub

Private Sub cmbDadIName_LostFocus()
On Error Resume Next
If cmbDadIName = "" Then
    Exit Sub
Else
    CheckData "Items", "IName", cmbDadIName.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Item , " & cmbDadIName.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbDadIName.Text = ""
        cmbDadIName.SetFocus
   
   
    End If
End If

End Sub

Private Sub cmbISize_Change()
On Error Resume Next
    cmbISize = UCase(cmbISize)
    SendKeys "{End}"
End Sub

Private Sub cmbISize_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbISize, cmbISize.Text)
End Sub

Private Sub cmbDadIBy_Change()
On Error Resume Next
    cmbDadIBy = UCase(cmbDadIBy)
    SendKeys "{End}"
End Sub

Private Sub cmbDadIBy_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbDadIBy, cmbDadIBy.Text)
End Sub
Private Sub cmbDadIName_Change()
On Error Resume Next
    cmbDadIName = UCase(cmbDadIName)
    SendKeys "{End}"
End Sub

Private Sub cmbDadIName_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbDadIName, cmbDadIName.Text)
End Sub

Private Sub cmbISize_LostFocus()
On Error Resume Next
If cmbISize.Text = "" Then
    Exit Sub
Else
    CheckData "Items", "ISize", cmbISize.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Item Size , " & cmbISize.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbISize.Text = ""
        cmbISize.SetFocus
        Exit Sub
    'Else
    '    txtRcvIQty.SetFocus
    End If
End If
Call RtnStock
If lblCStock = 0 Then
    MsgBox "You can't Dad " & UCase(cmbIssIName) & ", 0 (Zero) Stock Return.", vbCritical, Me.Caption
    Call ClearAll
End If
End Sub

Private Sub cmdClose_Click()
frmmain.Caption = "Easy Inventory Manager - Main Form"
On Error Resume Next
    Unload Me
End Sub
Private Sub cmdDelete_Click()
On Error Resume Next
    con.Execute "Delete * from Dad Where Dad.SrNo = " & Val(txtSrNo)
    MsgBox " Information is Deleted ", vbInformation, Me.Caption
    Call ClearAll
End Sub

Private Sub cmdNew_Click()
On Error Resume Next
    Call ClearAll
End Sub
Private Sub cmdSave_Click()
On Error Resume Next
If cmbDadIName = "" Then
    MsgBox "Plese Select Dad Item Name.", vbCritical, Me.Caption
    cmbDadIName.SetFocus
    Exit Sub
End If
If cmbISize = "" Then
    MsgBox "Please Select Dad Item Size ", vbCritical, Me.Caption
    cmbISize.SetFocus
    Exit Sub
End If
If txtDadIQty = "" Then
    MsgBox "Please Enter Dad Quantity ", vbCritical, Me.Caption
    txtDadIQty.SetFocus
    Exit Sub
End If
If cmbDadIBy = "" Then
    MsgBox "Please Select Dad By Name ", vbCritical, Me.Caption
    cmbDadIBy.SetFocus
    Exit Sub
End If

With rs
       .Open "Select * from Dad where SrNo = '" & Val(txtSrNo) & "'", con, adOpenDynamic, adLockOptimistic
    If .EOF = True And .BOF = True Then
        .Close
        .Open "Select * from Dad", con, adOpenDynamic, adLockOptimistic
        .AddNew
        !SrNo = GetNewNo("Dad")
        !DadItems = UCase(cmbDadIName)
        !DadSize = UCase(cmbISize)
        !Dad = txtDadIQty
        !Dadby = UCase(cmbDadIBy)
        !DadDate = DadDate
        .Update
        .Close
        MsgBox "Information is Saved", vbInformation, Me.Caption
    Else
        !SrNo = txtSrNo
        !DadItems = UCase(cmbDadIName)
        !DadSize = UCase(cmbISize)
        !Dad = txtDadIQty
        !Dadby = UCase(cmbDadIBy)
        !DadDate = DadDate
        .Update
        .Close
        MsgBox "Information is Updated", vbInformation, Me.Caption
    End If
    
End With
Set rs = Nothing
Call ClearAll
End Sub

Private Sub Cmdprint_Click()
DataReport5.Show
End Sub

Private Sub Form_Load()
BSE1.InitSubClassing
On Error Resume Next
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Invdata.mdb;Persist Security Info=False"
    con.Open
Call ClearAll
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set rs = Nothing
    Set con = Nothing
End Sub

Public Function ClearAll()
On Error Resume Next

FeedData "Items", "IName", cmbDadIName
FeedData "Items", "ISize", cmbISize
FeedData "EmpMaster", "EmpName", cmbDadIBy

txtSrNo = GetNewNo("Dad")
cmbDadIName.Text = ""
cmbISize.Text = ""
txtDadIQty.Text = ""
cmbDadIBy.Text = ""
DadDate = Date

cmbDadIName.SetFocus
End Function
Public Function RtnStock()
On Error Resume Next
Dim rstmp As New ADODB.Recordset
    rstmp.Open "Select sum(Return) from Return where Return.RtnItems ='" & UCase(cmbDadIName) & "' and Return.RtnSize ='" & UCase(cmbISize) & "'", con, adOpenDynamic, adLockOptimistic
        If IsNull(rstmp(0)) Then
            lblCStock = 0
        Else
            lblCStock = rstmp(0)
        End If
        rstmp.Close
Set rstmp = Nothing
End Function

Private Sub ImgClose_Click()
frmmain.Caption = "Easy Inventory Manager - Main Form"
Me.Hide
Unload Me

End Sub

Private Sub txtDadIQty_LostFocus()
On Error Resume Next
If Val(lblCStock) < Val(txtDadIQty) Then
    MsgBox " Dad Quantity can not greater then Return Stock", vbCritical, Me.Caption
    txtDadIQty = ""
    txtDadIQty.SetFocus
End If
End Sub


