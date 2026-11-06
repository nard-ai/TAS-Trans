VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchTimeKeep 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Employee"
   ClientHeight    =   4230
   ClientLeft      =   8130
   ClientTop       =   6195
   ClientWidth     =   4980
   ControlBox      =   0   'False
   Icon            =   "frmSearchTimeKeep.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   90
      TabIndex        =   0
      Top             =   540
      Width           =   4740
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   3840
         TabIndex        =   9
         Top             =   3240
         Width           =   825
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   3000
         TabIndex        =   8
         Top             =   3240
         Width           =   825
      End
      Begin MSComctlLib.ListView lvwdisplay 
         Height          =   2970
         Left            =   135
         TabIndex        =   1
         Top             =   270
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   5239
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Employee Name"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee No"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblRecount 
         BackStyle       =   0  'Transparent
         Caption         =   "count"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   2
         Top             =   3330
         Width           =   1560
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   900
      ButtonWidth     =   609
      ButtonHeight    =   741
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   960
         Left            =   90
         TabIndex        =   5
         Top             =   90
         Width           =   7980
         Begin VB.TextBox txtSearch 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1980
            TabIndex        =   6
            Top             =   0
            Width           =   2715
         End
         Begin VB.Label Label1 
            Caption         =   "Employee:"
            Height          =   195
            Left            =   1170
            TabIndex        =   7
            Top             =   90
            Width           =   915
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6210
         TabIndex        =   4
         Top             =   900
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmSearchTimeKeep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsHREmployees As ADODB.Recordset

Private Sub Command1_Click()
    frmTimeKeep.txtCancelSw.Text = "0"
    Call txtSearch_KeyPress(13)
End Sub

Private Sub Command2_Click()
frmTimeKeep.txtCancelSw.Text = "1"
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Dim StrQuery As String

    If TimeKeepFlag = True Then
        Me.Top = (frmTimeKeep.Height - Me.Height) / 8
        Me.Left = (frmTimeKeep.Width - Me.Width) / 2
    End If
    'Get all employee for a particular client not yet included in payroll.
    Set rsHREmployees = New ADODB.Recordset
    If TimeKeepFlag = True Then
        StrQuery = "SELECT a.surname, a.firstname, a.EmployeeID, a.EmpNo, a.ecolaamount, b.employmentstatusdesc FROM hremployees AS a LEFT JOIN hremploymentstatus as b ON a.EmploymentStatus = b.EmploymentStatus WHERE a.clientcode = " & _
            frmTimeKeep.txtClientCode & " and a.EmployeeID not in (select EmployeeID from employeebasicsalary) ORDER BY a.surname, a.firstname"
        Me.Caption = "Add Employee for " & frmTimeKeep.cmbCompany
    End If
    Set rsHREmployees = ReturnRst(StrQuery, cnnstr, mstatic)

    With lvwdisplay
        Call FillHREmployeeList(lvwdisplay, rsHREmployees)
        lvwdisplay.ColumnHeaders(1).Width = 2000
        lvwdisplay.ColumnHeaders(2).Width = 1500
        lblRecount = .ListItems.Count & " records "
        If .ListItems.Count = 0 Then
            MsgBox "No new employee selected from HR table!", vbOKOnly + vbInformation, PayrollAlert
        End If
    End With
End Sub

Private Sub FillHREmployeeList(ByRef mLsEmp As ListView, _
                        mRs As Recordset)
    Dim litem1 As ListItem

    mLsEmp.View = lvwReport
    mLsEmp.ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
        Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!Surname), "Blank Name", mRs!Surname & ", " & mRs!FirstName))
        litem1.SubItems(1) = IIf(IsNull(mRs!EmpNo), "", mRs!EmpNo) 'EmployeeID
        mRs.MoveNext
    Wend
End Sub

Private Sub lvwdisplay_DblClick()
    strrow = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).SubItems(1)
    txtSearch = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).Text
    If lvwdisplay.ListItems.Count = 0 Then
        MsgBox "List doesn't have any record from the table selected !", vbOKOnly + vbInformation, "Select..."
        Exit Sub
    End If
    Call Search(strrow)
    Unload Me
End Sub

Private Sub lvwdisplay_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtSearch = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).Text
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        strrow = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).SubItems(1)
        Call Search(strrow)
    End If
End Sub

Private Sub Search(ByVal strSearch As String)
Dim res, bkmrk

With rsHREmployees
    If .EOF Then
        .MoveLast
    ElseIf .BOF Then
        .MoveFirst
    End If
    
    bkmrk = .Bookmark
    .MoveFirst
    .Find "EmpNo= " & strSearch
    If Not .EOF Then
        Call Showdata(rsHREmployees)
    Else
        .Bookmark = bkmrk
    End If
    Unload Me
End With
End Sub

Private Sub Showdata(ByRef rec As ADODB.Recordset)
    
If TimeKeepFlag = True Then
    frmTimeKeep.txtEmpNum = IIf(IsNull(rec!EmployeeID), 0, rec!EmployeeID)
    frmTimeKeep.txtClient(0) = IIf(IsNull(rec!EmpNo), "", rec!EmpNo)
    frmTimeKeep.txtClient(1) = IIf(IsNull(rec!EmploymentStatusDesc), "", rec!EmploymentStatusDesc)
    frmTimeKeep.txtClient(2) = IIf(IsNull(rec!Surname), "", rec!Surname & ", " & rec!FirstName)
    frmTimeKeep.txtRates(4) = IIf(IsNull(rec!EcolaAmount), 0, rec!EcolaAmount)
End If

End Sub

