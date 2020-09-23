VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{41BC0B27-5B30-4FD1-AE28-32EA831E20D4}#1.0#0"; "bse_button.ocx"
Begin VB.Form frmcstock 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10995
   ClientLeft      =   3885
   ClientTop       =   570
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   Picture         =   "frmcstock.frx":0000
   ScaleHeight     =   549.75
   ScaleMode       =   2  'Point
   ScaleWidth      =   570
   ShowInTaskbar   =   0   'False
   Begin BSE_Engine.BSE BSE1 
      Left            =   2280
      Top             =   9000
      _ExtentX        =   6588
      _ExtentY        =   1085
   End
   Begin VB.CommandButton Cmdprint 
      BackColor       =   &H80000013&
      Caption         =   "&Print"
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Caption         =   "&Close"
      Height          =   375
      Left            =   10080
      MaskColor       =   &H80000009&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin MSComctlLib.ListView LVCStock 
      Height          =   6375
      Left            =   1200
      TabIndex        =   4
      Top             =   2160
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   11245
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sr.No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Items Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Open Stock"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Receive"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Issue"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Return"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Deffective"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Cur. Stock"
         Object.Width           =   2291
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK  INVENTORY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   600
      Left            =   3600
      TabIndex        =   3
      Top             =   360
      Width           =   5295
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   960
      Picture         =   "frmcstock.frx":51D9A
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CURRENT STOCK  INVENTORY"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   7995
   End
   Begin VB.Image ImgClose 
      Height          =   540
      Left            =   10560
      Picture         =   "frmcstock.frx":528B0
      Stretch         =   -1  'True
      Top             =   480
      Width           =   540
   End
End
Attribute VB_Name = "frmcstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As String


Private Sub cmdClose_Click()
frmmain.Caption = "Easy Inventory Manager - Main Form"
On Error Resume Next
    Unload Me
    Image1.Enabled = True
End Sub

Private Sub Cmdprint_Click()
DataReport1.Show
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
Public Function CStock()
On Error Resume Next
Dim rstmp As New ADODB.Recordset
Dim rstmp1 As New ADODB.Recordset
Dim rstmp2 As New ADODB.Recordset
Dim rstmp3 As New ADODB.Recordset
Dim rstmp4 As New ADODB.Recordset
Dim rss As New ADODB.Recordset

Dim i As Integer
i = 1

    rstmp.Open "select * from Items ", con, adOpenDynamic, adLockOptimistic
        Do While Not rstmp.EOF
            LVCStock.ListItems.Add i, , rstmp!SrNo
            LVCStock.ListItems(i).SubItems(1) = rstmp!IName
            LVCStock.ListItems(i).SubItems(2) = rstmp!ISize
            LVCStock.ListItems(i).SubItems(3) = rstmp!OpnStock
            
            rstmp1.Open "Select sum(Receive) from Receive where Receive.RcvItems ='" & rstmp!IName & "' and Receive.RcvSize ='" & rstmp!ISize & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rstmp1.EOF
                    If IsNull(rstmp1(0)) Then
                        LVCStock.ListItems(i).SubItems(4) = 0
                    Else
                        LVCStock.ListItems(i).SubItems(4) = rstmp1(0)
                    End If
                    rstmp1.MoveNext
                Loop
                If rstmp1.EOF = True Then
                    rstmp1.Close
                End If
                
            rstmp2.Open "Select sum(Issue) from Issue where Issue.IssItems ='" & rstmp!IName & "' and Issue.IssSize ='" & rstmp!ISize & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rstmp2.EOF
                    If IsNull(rstmp2(0)) Then
                        LVCStock.ListItems(i).SubItems(5) = 0
                    Else
                        LVCStock.ListItems(i).SubItems(5) = rstmp2(0)
                    End If
                    rstmp2.MoveNext
                Loop
                If rstmp2.EOF = True Then
                    rstmp2.Close
                End If
            
            rstmp3.Open "Select sum(Return) from Return where Return.RtnItems ='" & rstmp!IName & "' and Return.RtnSize ='" & rstmp!ISize & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rstmp3.EOF
                    If IsNull(rstmp3(0)) Then
                        LVCStock.ListItems(i).SubItems(6) = 0
                    Else
                        LVCStock.ListItems(i).SubItems(6) = rstmp3(0)
                    End If
                    rstmp3.MoveNext
                Loop
                If rstmp3.EOF = True Then
                    rstmp3.Close
                End If
            
            rstmp4.Open "Select sum(Dad) from Dad where Dad.DadItems ='" & rstmp!IName & "' and Dad.DadSize ='" & rstmp!ISize & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rstmp4.EOF
                    If IsNull(rstmp4(0)) Then
                        LVCStock.ListItems(i).SubItems(7) = 0
                    Else
                        LVCStock.ListItems(i).SubItems(7) = rstmp4(0)
                    End If
                    rstmp4.MoveNext
                Loop
                If rstmp4.EOF = True Then
                    rstmp4.Close
                End If
            LVCStock.ListItems(i).SubItems(8) = Val(LVCStock.ListItems(i).SubItems(3)) + Val(LVCStock.ListItems(i).SubItems(4)) - Val(LVCStock.ListItems(i).SubItems(5)) + Val(LVCStock.ListItems(i).SubItems(6)) - Val(LVCStock.ListItems(i).SubItems(7))
                 rs1 = "insert into cstock values (" & i & "," & LVCStock.ListItems(i).SubItems(8) & ")"
                con.Execute rs1
            i = i + 1
            rstmp.MoveNext
        Loop
    
  
    rstmp.Close
    rstmp1.Close
    rstmp2.Close
    rstmp3.Close
    rstmp4.Close
    
Set rstmp = Nothing
Set rstmp1 = Nothing
Set rstmp2 = Nothing
Set rstmp3 = Nothing
Set rstmp4 = Nothing

End Function
Public Function ClearAll()
On Error Resume Next
    LV.ListItems.clear
    rs1 = "delete from cstock"
    con.Execute rs1
    Call CStock
End Function

Private Sub ImgClose_Click()
frmmain.Caption = "Easy Inventory Manager - Main Form"
On Error Resume Next
Me.Hide
Unload Me
Image1.Show
End Sub


