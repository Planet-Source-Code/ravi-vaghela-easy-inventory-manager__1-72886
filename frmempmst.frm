VERSION 5.00
Object = "{41BC0B27-5B30-4FD1-AE28-32EA831E20D4}#1.0#0"; "bse_button.ocx"
Begin VB.Form frmempmst 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   4275
   ClientTop       =   1530
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmempmst.frx":0000
   ScaleHeight     =   8265
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin BSE_Engine.BSE BSE1 
      Left            =   1800
      Top             =   7080
      _ExtentX        =   6588
      _ExtentY        =   1085
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
      Left            =   7440
      TabIndex        =   20
      Text            =   "0"
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H80000013&
      Caption         =   "&New"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000013&
      Caption         =   "&Save"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H80000013&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000013&
      Caption         =   "&Close"
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   1215
   End
   Begin VB.ComboBox cmbDesig 
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
      Left            =   3960
      TabIndex        =   1
      Top             =   3840
      Width           =   5655
   End
   Begin VB.ComboBox cmbCity 
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
      Left            =   3960
      TabIndex        =   3
      Top             =   4800
      Width           =   3735
   End
   Begin VB.ComboBox cmbEmpName 
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
      Left            =   3960
      TabIndex        =   0
      Top             =   3360
      Width           =   5655
   End
   Begin VB.TextBox txtMobile 
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
      Left            =   7320
      TabIndex        =   5
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox txtPhone 
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
      Left            =   3960
      TabIndex        =   4
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox txtAdd 
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
      Left            =   3960
      TabIndex        =   2
      Top             =   4320
      Width           =   5655
   End
   Begin VB.TextBox txtEmpCode 
      Alignment       =   2  'Center
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
      Left            =   3960
      TabIndex        =   19
      Text            =   "c-1"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1020
      Left            =   2520
      Picture         =   "frmempmst.frx":51D9A
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE  MASTER"
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
      Left            =   3720
      TabIndex        =   18
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Emp.  No :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      TabIndex        =   17
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile :-"
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
      Left            =   6360
      TabIndex        =   16
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone :-"
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
      Left            =   2760
      TabIndex        =   15
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City :-"
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
      Left            =   2760
      TabIndex        =   14
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address :-"
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
      Left            =   2760
      TabIndex        =   13
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desig :-"
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
      Left            =   2760
      TabIndex        =   12
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name :-"
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
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code :-"
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
      TabIndex        =   10
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Image ImgClose 
      Height          =   540
      Left            =   9360
      Picture         =   "frmempmst.frx":5281E
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   540
   End
   Begin VB.Image Image2 
      Height          =   11010
      Left            =   0
      Picture         =   "frmempmst.frx":5332A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "frmempmst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmbCity_Change()
On Error Resume Next
    cmbCity = UCase(cmbCity)
    SendKeys "{End}"
End Sub

Private Sub cmbDesig_Change()
On Error Resume Next
    cmbDesig = UCase(cmbDesig)
    SendKeys "{End}"
End Sub

Private Sub cmbDesig_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If checkCharacter(KeyCode) Then Call findString(cmbDesig, cmbDesig.Text)
End Sub

Private Sub cmbEmpName_Change()
On Error Resume Next
    cmbEmpName = UCase(cmbEmpName)
    SendKeys "{End}"
End Sub
Private Sub cmbCity_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If checkCharacter(KeyCode) Then Call findString(cmbCity, cmbCity.Text)
End Sub
Private Sub cmbEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
   If checkCharacter(KeyCode) Then Call findString(cmbEmpName, cmbEmpName.Text)
End Sub

Private Sub cmbEmpName_LostFocus()
On Error Resume Next
Dim rsf As New ADODB.Recordset
    rsf.Open "Select * from EmpMaster where EmpMaster.EmpName = '" & UCase(cmbEmpName) & "'", con, adOpenDynamic, adLockOptimistic
    If rsf.BOF = True And rsf.EOF = True Then
        cmbDesig = ""
        txtAdd = ""
        cmbCity = ""
        txtPhone = ""
        txtMobile = ""
    Else
        'txtEmpCode = rsf!EmpCode
        'txtSrNo = rsf!SrNo
        cmbDesig = rsf!Desig
        txtAdd = rsf!Add
        cmbCity = rsf!City
        txtPhone = rsf!Phone
        txtMobile = rsf!Mobile
    End If
    rsf.Close
    Set rsf = Nothing
    
End Sub

Private Sub cmdClose_Click()
frmmain.Caption = "Easy Inventory Manager - Main Form"
On Error Resume Next
    Unload Me
End Sub

Public Function ClearAll()
On Error Resume Next

FeedData "EmpMaster", "EmpName", cmbEmpName
FeedData "EmpMaster", "Desig", cmbDesig
FeedData "EmpMaster", "City", cmbCity
txtEmpCode.Text = txtEmpCode.Text
txtSrNo = GetNewNo("EmpMaster")
cmbEmpName.Text = ""
cmbDesig.Text = ""
txtAdd.Text = ""
cmbCity.Text = ""
txtPhone.Text = ""
txtMobile.Text = ""

cmbEmpName.SetFocus

End Function

Private Sub cmdDelete_Click()
On Error Resume Next
    con.Execute "Delete * from EmpMaster Where EmpMaster.SrNo = " & Val(txtSrNo)
    MsgBox "Information is Deleted", vbInformation, Me.Caption
    Call ClearAll
End Sub



Private Sub cmdNew_Click()
On Error Resume Next
    Call ClearAll
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
If cmbEmpName = "" Then
    MsgBox "Please Enter Employee Name ", vbCritical, Me.Caption
    Exit Sub
End If
If cmbDesig = "" Then
    MsgBox "Please Enter Employee Designation ", vbCritical, Me.Caption
    Exit Sub
End If
If txtPhone = "" And txtMobile = "" Then
    MsgBox "Please Enter Employee's Phone or Mobile ", vbCritical, Me.Caption
    Exit Sub
End If
With rs
    .Open "Select * from EmpMaster where EmpName = '" & UCase(cmbEmpName) & "'", con, adOpenDynamic, adLockOptimistic
    If .EOF = True And .BOF = True Then
        .Close
        .Open "Select * from EmpMaster", con, adOpenDynamic, adLockOptimistic
        .AddNew
        !EmpCode = UCase(txtEmpCode)
        !SrNo = GetNewNo("EmpMaster")
        !EmpName = UCase(cmbEmpName)
        !Desig = UCase(cmbDesig)
        !Add = UCase(txtAdd)
        !City = UCase(cmbCity)
        !Phone = txtPhone
        !Mobile = txtMobile
        .Update
        .Close
        MsgBox "Information is Saved", vbInformation, Me.Caption
    Else
        !EmpCode = UCase(txtEmpCode)
        !SrNo = txtSrNo
        !EmpName = UCase(cmbEmpName)
        !Desig = UCase(cmbDesig)
        !Add = UCase(txtAdd)
        !City = UCase(cmbCity)
        !Phone = txtPhone
        !Mobile = txtMobile
        .Update
        .Close
        MsgBox "Information is Updated", vbInformation, Me.Caption
    End If

End With
Set rs = Nothing
Call ClearAll
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

Private Sub ImgClose_Click()
frmmain.Caption = "Easy Inventory Manager - Main Form"
Me.Hide
Unload Me
End Sub

