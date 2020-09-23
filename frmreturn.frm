VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{41BC0B27-5B30-4FD1-AE28-32EA831E20D4}#1.0#0"; "bse_button.ocx"
Begin VB.Form frmreturn 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7695
   ClientLeft      =   4470
   ClientTop       =   1725
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   Picture         =   "frmreturn.frx":0000
   ScaleHeight     =   7695
   ScaleMode       =   0  'User
   ScaleWidth      =   15000
   ShowInTaskbar   =   0   'False
   Begin BSE_Engine.BSE BSE1 
      Left            =   1080
      Top             =   6840
      _ExtentX        =   6588
      _ExtentY        =   1085
   End
   Begin VB.CommandButton Cmdprint 
      BackColor       =   &H80000013&
      Caption         =   "&Print"
      Height          =   495
      Left            =   5160
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H80000013&
      Caption         =   "&New"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000013&
      Caption         =   "&Save"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H80000013&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000013&
      Caption         =   "&Close"
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   1215
   End
   Begin VB.ComboBox cmbRRcvIBy 
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
      Left            =   3240
      TabIndex        =   4
      Top             =   4680
      Width           =   3855
   End
   Begin VB.TextBox txtRtnIQty 
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
      Left            =   3240
      TabIndex        =   2
      Top             =   3720
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
      Left            =   3240
      TabIndex        =   1
      Top             =   3240
      Width           =   3855
   End
   Begin VB.ComboBox cmbRtnIName 
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
      Left            =   3240
      TabIndex        =   0
      Top             =   2760
      Width           =   3855
   End
   Begin VB.ComboBox cmbRtnIBy 
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
      Left            =   3240
      TabIndex        =   3
      Top             =   4200
      Width           =   3855
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
      Left            =   3240
      TabIndex        =   22
      Text            =   "0"
      Top             =   2280
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker RtnDate 
      Height          =   330
      Left            =   3240
      TabIndex        =   9
      Top             =   5160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Format          =   55705601
      CurrentDate     =   39274
   End
   Begin VB.Label lblBStock 
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
      Left            =   8280
      TabIndex        =   25
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Stock"
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
      Left            =   8160
      TabIndex        =   21
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label lblRStock 
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
      Left            =   9120
      TabIndex        =   24
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label10 
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
      Left            =   9000
      TabIndex        =   20
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Image Image7 
      Height          =   1020
      Left            =   960
      Picture         =   "frmreturn.frx":51D9A
      Top             =   600
      Width           =   1020
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RETURN STOCK ENTRY"
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
      Left            =   2160
      TabIndex        =   19
      Top             =   840
      Width           =   6135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Stock"
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
      Left            =   7320
      TabIndex        =   18
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblIStock 
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
      Left            =   7440
      TabIndex        =   23
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date :-"
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
      Left            =   1680
      TabIndex        =   17
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Return Receive  By :-"
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
      Left            =   960
      TabIndex        =   16
      Top             =   4680
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
      Left            =   1800
      TabIndex        =   15
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Return Item Qty :-"
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
      Left            =   960
      TabIndex        =   14
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Return Item Name :-"
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
      Left            =   1200
      TabIndex        =   13
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Return  By :-"
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
      Left            =   1440
      TabIndex        =   12
      Top             =   4200
      Width           =   1575
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
      Left            =   1800
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image ImgClose 
      Height          =   540
      Left            =   9960
      Picture         =   "frmreturn.frx":529A2
      Stretch         =   -1  'True
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frmreturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmbISize_Change()
On Error Resume Next
    cmbISize = UCase(cmbISize)
    SendKeys "{End}"
End Sub

Private Sub cmbISize_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbISize, cmbISize.Text)
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
    '    txtRtnIQty.SetFocus
    End If

End If
Call IssStock

If lblIStock = 0 Then
    MsgBox "You can't Return " & UCase(cmbRtnIName) & ", 0 (Zero) Stock Issue.", vbCritical, Me.Caption
    Call ClearAll
End If
Call RtnStock

lblBStock = Val(lblIStock) - Val(lblRStock)
If lblBStock = 0 Then
    MsgBox "You can't Return  " & cmbRtnIName.Text & "  because Balance Stock is 0 ", vbCritical, Me.Caption
End If
End Sub

Private Sub cmbRRcvIBy_LostFocus()
On Error Resume Next
If cmbRRcvIBy = "" Then
    Exit Sub
Else
    CheckData "EmpMaster", "EmpName", cmbRRcvIBy.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Return Receiver Name , " & cmbRRcvIBy.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbRRcvIBy.Text = ""
        cmbRRcvIBy.SetFocus
    'Else
    '    RtnDate.SetFocus
    End If
End If

End Sub

Private Sub cmbRtnIBy_Change()
On Error Resume Next
    cmbRtnIBy = UCase(cmbRtnIBy)
    SendKeys "{End}"
End Sub

Private Sub cmbRtnIBy_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbRtnIBy, cmbRtnIBy.Text)
End Sub
Private Sub cmbRRcvIBy_Change()
On Error Resume Next
    cmbRRcvIBy = UCase(cmbRRcvIBy)
    SendKeys "{End}"
End Sub

Private Sub cmbRRcvIBy_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbRRcvIBy, cmbRRcvIBy.Text)
End Sub

Private Sub cmbRtnIBy_LostFocus()
On Error Resume Next
Dim rstmp As New ADODB.Recordset
Dim rstmp1 As New ADODB.Recordset

If cmbRtnIBy = "" Then
    Exit Sub
Else
    CheckData "EmpMaster", "EmpName", cmbRtnIBy.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Reterner Name , " & cmbRtnIBy.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbRtnIBy.Text = ""
        cmbRtnIBy.SetFocus
        Exit Sub
    'Else
    '    cmbRRcvIBy.SetFocus
    End If
End If



    rstmp.Open "Select sum(Issue) from Issue where  Issue.IssItems ='" & UCase(cmbRtnIName) & "' and Issue.IssSize ='" & UCase(cmbISize) & "' and Issue.IReceiveby ='" & UCase(cmbRtnIBy) & "'", con, adOpenDynamic, adLockOptimistic
        If IsNull(rstmp(0)) Then
            lblIStock = 0
        Else
            lblIStock = rstmp(0)
        End If
    rstmp1.Open "Select sum(Return) from Return where Return.RtnItems ='" & UCase(cmbRtnIName) & "' and Return.RtnSize ='" & UCase(cmbISize) & "' and Return.Returnby ='" & UCase(cmbRtnIBy) & "'", con, adOpenDynamic, adLockOptimistic
        If IsNull(rstmp1(0)) Then
            lblRStock = 0
        Else
            lblRStock = rstmp1(0)
        End If
    lblBStock = Val(lblIStock) - Val(lblRStock)
            
    If lblBStock = 0 Or lblBStock < 0 Then
        MsgBox UCase(cmbRtnIBy) & " can't Return " & UCase(cmbRtnIName) & ", 0 (Zero) Stock Issue.", vbCritical, Me.Caption
    End If
    rstmp1.Close
    rstmp.Close
Set rstmp1 = Nothing
Set rstmp = Nothing

If Val(lblBStock) < Val(txtRtnIQty) Then
    MsgBox cmbRtnIBy & " can't Return, because Return Quantity not grater than Balance Stock ", vbCritical, Me.Caption
    txtRtnIQty = ""
    txtRtnIQty.SetFocus
End If
End Sub

Private Sub cmbRtnIName_Change()
On Error Resume Next
    cmbRtnIName = UCase(cmbRtnIName)
    SendKeys "{End}"
End Sub

Private Sub cmbRtnIName_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbRtnIName, cmbRtnIName.Text)
End Sub

Private Sub cmbRtnIName_LostFocus()
If cmbRtnIName = "" Then
    Exit Sub
Else
    CheckData "Items", "IName", cmbRtnIName.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Item , " & cmbRtnIName.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbRtnIName.Text = ""
        cmbRtnIName.SetFocus
    'Else
    '    cmbISize.SetFocus
    End If
End If
End Sub

Private Sub cmdClose_Click()
frmmain.Caption = "Easy Inventory Manager - Main Form"
On Error Resume Next
    Unload Me
End Sub
Private Sub cmdDelete_Click()
On Error Resume Next
    con.Execute "Delete * from Return Where Return.SrNo = " & txtSrNo
    MsgBox "Information is Deleted", vbInformation, Me.Caption
    Call ClearAll
End Sub


Private Sub cmdNew_Click()
On Error Resume Next
    Call ClearAll
End Sub
Private Sub cmdSave_Click()
On Error Resume Next
If cmbRtnIName = "" Then
    MsgBox "Plese Select Return Item Name.", vbCritical, Me.Caption
    cmbRtnIName.SetFocus
    Exit Sub
End If
If cmbISize = "" Then
    MsgBox "Please Select Return Item Size ", vbCritical, Me.Caption
    cmbISize.SetFocus
    Exit Sub
End If
If txtRtnIQty = "" Then
    MsgBox "Please Enter Return Quantity ", vbCritical, Me.Caption
    txtRtnIQty.SetFocus
    Exit Sub
End If
If cmbRtnIBy = "" Then
    MsgBox "Please Select Return By Name ", vbCritical, Me.Caption
    cmbRtnIBy.SetFocus
    Exit Sub
End If
If cmbRRcvIBy = "" Then
    MsgBox "Please Select Return Receive By Name ", vbCritical, Me.Caption
    cmbRRcvIBy.SetFocus
    Exit Sub
End If
If txtRtnIQty > Val(lblBStock) Then
    MsgBox "Return Quantity not greater then Balance Quantity ", vbCritical, Me.Caption
    txtRtnIQty.SetFocus
    Exit Sub
End If
With rs
    '.Open "Select * from Return where RtnItems = '" & UCase(cmbRtnIName) & "'", con, adOpenDynamic, adLockOptimistic
    .Open "Select * from Return where SrNo = '" & Val(txtSrNo) & "'", con, adOpenDynamic, adLockOptimistic
    If .EOF = True And .BOF = True Then
        .Close
        .Open "Select * from Return", con, adOpenDynamic, adLockOptimistic
        .AddNew
        !SrNo = GetNewNo("Return")
        !RtnItems = UCase(cmbRtnIName)
        !RtnSize = UCase(cmbISize)
        !Return = txtRtnIQty
        !Returnby = UCase(cmbRtnIBy)
        !RReceiveby = UCase(cmbRRcvIBy)
        !RtnDate = RtnDate
        .Update
        .Close
        MsgBox "Information is Saved", vbInformation, Me.Caption
    Else
        !SrNo = txtSrNo
        !RtnItems = UCase(cmbRtnIName)
        !RtnSize = UCase(cmbISize)
        !Return = txtRtnIQty
        !Returnby = UCase(cmbRtnIBy)
        !RReceiveby = UCase(cmbRRcvIBy)
        !RtnDate = RtnDate
        .Update
        .Close
        MsgBox "Information is Updated", vbInformation, Me.Caption
    End If
    
End With
Set rs = Nothing
Call ClearAll
End Sub

Private Sub Cmdprint_Click()
DataReport6.Show
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

FeedData "Items", "IName", cmbRtnIName
FeedData "Items", "ISize", cmbISize
FeedData "EmpMaster", "EmpName", cmbRtnIBy
FeedData "EmpMaster", "EmpName", cmbRRcvIBy
txtSrNo = GetNewNo("Return")
cmbRtnIName.Text = ""
cmbISize.Text = ""
txtRtnIQty.Text = ""
cmbRtnIBy.Text = ""
cmbRRcvIBy.Text = ""
RtnDate = Date

cmbRtnIName.SetFocus
End Function

Public Function IssStock()
On Error Resume Next
Dim rstmp As New ADODB.Recordset
    
    rstmp.Open "Select sum(Issue) from Issue where Issue.IssItems ='" & UCase(cmbRtnIName) & "' and Issue.IssSize ='" & UCase(cmbISize) & "'", con, adOpenDynamic, adLockOptimistic
        If IsNull(rstmp(0)) Then
            lblIStock = 0
        Else
            lblIStock = rstmp(0)
        End If
    rstmp.Close
Set rstmp = Nothing
End Function

Public Function RtnStock()
On Error Resume Next
Dim rstmp As New ADODB.Recordset
    
    rstmp.Open "Select sum(Return) from Return where Return.RtnItems ='" & UCase(cmbRtnIName) & "' and Return.RtnSize ='" & UCase(cmbISize) & "'", con, adOpenDynamic, adLockOptimistic
        If IsNull(rstmp(0)) Then
            lblRStock = 0
        Else
            lblRStock = rstmp(0)
        End If
    rstmp.Close
Set rstmp = Nothing
End Function

Private Sub ImgClose_Click()
frmmain.Caption = "Easy Inventory Manager - Main Form"
Me.Hide
Unload Me

End Sub

Private Sub txtRtnIQty_LostFocus()
On Error Resume Next
If Val(lblBStock) < Val(txtRtnIQty) Then
    MsgBox "Return Quantity not grater than Balance Stock ", vbCritical, Me.Caption
    txtRtnIQty = ""
    txtRtnIQty.SetFocus
End If
    
End Sub
