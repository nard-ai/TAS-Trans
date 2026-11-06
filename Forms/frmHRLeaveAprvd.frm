VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRLeaveAprvd 
   BackColor       =   &H80000009&
   Caption         =   "Approved"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Picture         =   "frmHRLeaveAprvd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Refresh"
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   4695
      Begin VB.OptionButton OptForAprv 
         BackColor       =   &H8000000E&
         Caption         =   "For Approval"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptAprv 
         BackColor       =   &H80000009&
         Caption         =   "Aprroved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptAllAprv 
         BackColor       =   &H80000009&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton btnAprvdExit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   10200
      Picture         =   "frmHRLeaveAprvd.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   360
      Width           =   855
   End
   Begin MSComctlLib.ListView lvwAprvLeave 
      Height          =   6135
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10821
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Leave No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Employee No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "FullName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type of Leave"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Leave Served"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Applied?"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Approved?"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Noted?"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Department"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Check if leave is approved"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   2415
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmHRLeaveAprvd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim list As ListItem
Dim rs As ADODB.Recordset
Dim str, lno As String
Dim allaprv, aprvd, foraprv As Boolean

Private Sub btnAprvdExit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub btnRefresh_Click()
    If HR = "MG" Then
        If allaprv = True Then 'All leaves
            str = "Select * from HREmpLeaves where applied=" & True
        ElseIf aprvd = True Then ' Approved leaves
            str = "Select * from HREmpLeaves where applied=" & True & " And applied =" & True & " and approved=" & True
        ElseIf foraprv = True Then 'For Approval
            str = "Select * from HREmpLeaves where applied=" & True & " And applied =" & True & " and approved=" & False
        End If
    Else
        If allaprv = True Then 'All leaves
            str = "Select * from HREmpLeaves where deptcode=" & Format(DEP, "00000") & " And applied =" & True
        ElseIf aprvd = True Then ' Approved leaves
            str = "Select * from HREmpLeaves where deptcode=" & Format(DEP, "00000") & " And applied =" & True & " and approved=" & True
        ElseIf foraprv = True Then 'For Approval
            str = "Select * from HREmpLeaves where deptcode=" & Format(DEP, "00000") & " And applied =" & True & " and approved=" & False
        End If
   End If
        ViewList
        allaprv = False
        aprvd = False
        foraprv = False
End Sub

Private Sub Form_Load()
    If HR = "MG" Then
        str = "Select * from HREmpLeaves where applied=" & True & " And applied =" & True & " and approved=" & False
    Else
        str = "Select * from HREmpLeaves where deptcode=" & Format(DEP, "00000") & " And applied =" & True & " and approved=" & False
    End If
    ViewList
    
End Sub

Private Sub lvwAprvLeave_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If Item.Checked = True Then
        'call query to tagged this record as approved
        cnDCPI.Execute "Update HREmpLeaves Set Approved = 1 where leavenum=" & Item.Text
    Else
        cnDCPI.Execute "Update HREmpLeaves Set Approved = 0 where leavenum=" & Item.Text
    End If
    
End Sub

Private Sub OptAllAprv_Click()
    'All leaves
    allaprv = True
End Sub

Private Sub OptAprv_Click()
    'All approved leaves
    aprvd = True
End Sub

Private Sub OptForAprv_Click()
    'For approval
    foraprv = True
End Sub

Private Sub ViewList()
    Set rs = cnDCPI.Execute(str)
    lvwAprvLeave.ListItems.Clear
        Do While Not rs.EOF
            Set list = lvwAprvLeave.ListItems.add(, , rs!leavenum)
            list.SubItems(1) = IIf(IsNull(rs!employeenum), "", rs!employeenum)
            list.SubItems(2) = IIf(IsNull(rs!FullName), "", rs!FullName)
            list.SubItems(3) = IIf(IsNull(rs!leavetypedesc), "", rs!leavetypedesc)
            list.SubItems(4) = IIf(IsNull(rs!leavedatefr), "", rs!leavedatefr)
            list.SubItems(5) = IIf(IsNull(rs!leavedateto), "", rs!leavedateto)
            list.SubItems(6) = IIf(IsNull(rs!leavebal), "", rs!leavebal)
            list.SubItems(7) = IIf(IsNull(rs!leaveserved), "", rs!leaveserved)
            list.SubItems(8) = IIf(IsNull(rs!applied), "", IIf(rs!applied, "Yes", "No"))
            list.SubItems(9) = IIf(IsNull(rs!approved), "", IIf(rs!approved, "Yes", "No"))
            list.SubItems(10) = IIf(IsNull(rs!noted), "", IIf(rs!noted, "Yes", "No"))
            list.SubItems(11) = IIf(IsNull(rs!Department), "", rs!Department)
            list.Checked = IIf(rs!approved, True, False)
            
        rs.MoveNext
        Loop
        rs.Close


End Sub
