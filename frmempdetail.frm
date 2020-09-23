VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{41BC0B27-5B30-4FD1-AE28-32EA831E20D4}#1.0#0"; "bse_button.ocx"
Begin VB.Form frmempdetail 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   4470
   ClientTop       =   1350
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   Picture         =   "frmempdetail.frx":0000
   ScaleHeight     =   8220
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   Begin BSE_Engine.BSE BSE1 
      Left            =   2400
      Top             =   6840
      _ExtentX        =   6588
      _ExtentY        =   1085
   End
   Begin VB.CommandButton Cmdprint 
      BackColor       =   &H80000013&
      Caption         =   "&Print Details"
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin MSComctlLib.ListView LVEmp 
      Height          =   4890
      Left            =   570
      TabIndex        =   0
      Top             =   1320
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8625
      SortKey         =   1
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SrNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Emp Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Desig"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Address"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "City"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Phone"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Mobile"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Image ImgClose 
      Height          =   540
      Left            =   10320
      Picture         =   "frmempdetail.frx":51D9A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE  MASTER  DETAILS"
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
      Left            =   1830
      TabIndex        =   1
      Top             =   480
      Width           =   8085
   End
   Begin VB.Image Image1 
      Height          =   1020
      Left            =   480
      Picture         =   "frmempdetail.frx":528A6
      Top             =   240
      Width           =   1020
   End
End
Attribute VB_Name = "frmempdetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim i As Integer

Private Sub Cmdprint_Click()
DataReport4.Show
End Sub

Private Sub Form_Load()
BSE1.InitSubClassing
On Error Resume Next
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Invdata.mdb;Persist Security Info=False"
    con.Open
   
 i = 1
    rs.Open "Select * from EmpMaster ", con, adOpenDynamic, adLockOptimistic
        Do While Not rs.EOF
            LVEmp.ListItems.Add i, , i
            LVEmp.ListItems(i).SubItems(1) = rs!EmpCode & " - " & rs!SrNo
            LVEmp.ListItems(i).SubItems(2) = rs!EmpName
            LVEmp.ListItems(i).SubItems(3) = rs!Desig
            LVEmp.ListItems(i).SubItems(4) = rs!Add
            LVEmp.ListItems(i).SubItems(5) = rs!City
            LVEmp.ListItems(i).SubItems(6) = rs!Phone
            LVEmp.ListItems(i).SubItems(7) = rs!Mobile
            
            i = i + 1
            rs.MoveNext
        Loop
    rs.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set rs = Nothing
    Set con = Nothing
End Sub

Private Sub ImgClose_Click()
frmmain.Caption = "Easy Inventory Manager - Main Form"
On Error Resume Next
    Unload Me
End Sub


