VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUserTransactions 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of User Transactions"
   ClientHeight    =   8100
   ClientLeft      =   1575
   ClientTop       =   3525
   ClientWidth     =   14505
   Icon            =   "frmUserTransactions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   14505
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   7350
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   14415
      Begin VB.ComboBox cmbTransaction 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmUserTransactions.frx":058A
         Left            =   10920
         List            =   "frmUserTransactions.frx":058C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox cmbEmployee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmUserTransactions.frx":058E
         Left            =   6600
         List            =   "frmUserTransactions.frx":0590
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox cmbStaff 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmUserTransactions.frx":0592
         Left            =   1320
         List            =   "frmUserTransactions.frx":0594
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   315
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   900
      End
      Begin MSComctlLib.ListView lvwdisplay 
         Height          =   6210
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   10954
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SEX"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ADDRESS"
            Object.Width           =   4586
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transactions:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   9825
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   5280
         TabIndex        =   15
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Staff:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblRecount 
         Alignment       =   1  'Right Justify
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
         Left            =   12960
         TabIndex        =   2
         Top             =   6960
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   847
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   0
         TabIndex        =   4
         Top             =   -120
         Width           =   14340
         Begin VB.CheckBox chkHistory 
            Caption         =   "Salary History"
            Height          =   255
            Left            =   9960
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   21
            Top             =   240
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.ComboBox cmbCompany 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmUserTransactions.frx":0596
            Left            =   6600
            List            =   "frmUserTransactions.frx":0598
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   200
            Width           =   3015
         End
         Begin VB.CommandButton cmdExport 
            Caption         =   "Export to Excel"
            Height          =   315
            Left            =   12600
            TabIndex        =   16
            Top             =   200
            Width           =   1500
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1920
            TabIndex        =   8
            Top             =   160
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   61865985
            CurrentDate     =   40124
         End
         Begin VB.ComboBox cmbmonth 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmUserTransactions.frx":059A
            Left            =   10110
            List            =   "frmUserTransactions.frx":059C
            Style           =   2  'Dropdown List
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   135
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.ComboBox cmbYear 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmUserTransactions.frx":059E
            Left            =   11415
            List            =   "frmUserTransactions.frx":05A0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   135
            Visible         =   0   'False
            Width           =   930
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   3720
            TabIndex        =   11
            Top             =   165
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   61865985
            CurrentDate     =   40124
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   5280
            TabIndex        =   17
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "To:"
            Height          =   195
            Index           =   1
            Left            =   3360
            TabIndex        =   12
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Transaction Date From:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1665
         End
      End
   End
End
Attribute VB_Name = "frmUserTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsClient As New ADODB.Recordset
Dim rsEmployees As ADODB.Recordset
Dim rsForExcel As ADODB.Recordset
Dim SelectString As String, whereCon As String
Dim rs As ADODB.Recordset
Dim listmove As ListItem
Dim NoTransFlag As Boolean

Private Function Filter_var() As String
Dim StaffFlag As Boolean, Companyflag As Boolean
Dim EmployeeFlag As Boolean, TransFlag As Boolean

Filter_var = ""
StaffFlag = False: Companyflag = False: EmployeeFlag = False: TransFlag = False
If cmbStaff.Text <> "--ALL STAFF--" Then
    Filter_var = " UserID = " & cmbStaff.ItemData(cmbStaff.ListIndex)
End If
If cmbCompany.Text <> "--ALL--" And cmbCompany <> "" Then
    If Filter_var <> "" Then
        Filter_var = Filter_var & " AND ClientCode = " & cmbCompany.ItemData(cmbCompany.ListIndex)
    Else
        Filter_var = " ClientCode = " & cmbCompany.ItemData(cmbCompany.ListIndex)
    End If
End If
If cmbEmployee.Text <> "--ALL--" And cmbEmployee <> "" Then
    If Filter_var <> "" Then
        Filter_var = Filter_var & " AND b.EmployeeID = " & cmbEmployee.ItemData(cmbEmployee.ListIndex)
    Else
        Filter_var = " b.EmployeeID = " & cmbEmployee.ItemData(cmbEmployee.ListIndex)
    End If
End If
If cmbTransaction.Text <> "--ALL--" And cmbTransaction <> "" Then
    If Filter_var <> "" Then
        Filter_var = Filter_var & " AND TransRemarks = '" & cmbTransaction.Text & "'"
    Else
        Filter_var = " TransRemarks = '" & cmbTransaction.Text & "'"
    End If
End If

'Filter_var = "MONTH(TransDate) = " & month(DTPicker1) & " and YEAR(TransDate) = " & Year(DTPicker1) & " and DAY(TransDate) = " & Day(DTPicker1)
If Filter_var <> "" Then
    Filter_var = Filter_var & " AND TransDate BETWEEN '" & DTPicker1 & "' and '" & DTPicker2 + 1 & "' "
Else
    Filter_var = " TransDate BETWEEN '" & DTPicker1 & "' and '" & DTPicker2 + 1 & "' "
End If
whereCon = Filter_var
End Function

Private Sub cmbCompany_Click()
Loadlist (Filter_var)
If cmbCompany.Text <> "--ALL--" Then
    Call FilterEmployee(cmbCompany.Text)
End If
End Sub

Private Sub cmbEmployee_Click()
Loadlist (Filter_var)
End Sub

Private Sub cmbStaff_Click()
Loadlist (Filter_var)
End Sub

Private Sub cmbTransaction_Click()
Loadlist (Filter_var)
End Sub

Private Sub FilterEmployee(ByVal mCompany As String)
Dim mCCode As Integer
    
mClientCode = cmbCompany.ItemData(cmbCompany.ListIndex)
rsClient.MoveFirst
rsClient.Find "clientname = '" & Trim(mCompany) & "'"
If rsClient.EOF = True Then
    mCCode = 0 ' updated by bernard 3/9/2009
    'Exit Sub
Else
    mCCode = rsClient!ClientCode
End If
With rsEmployees
    If mCCode <> 0 Then
        .Filter = " clientcode = " & mCCode
    Else
        .Filter = ""
    End If
    If .RecordCount <> 0 Then
        Call oCommonFunction.FillCombo(cmbEmployee, rsEmployees, 11)
    Else
        Call oCommonFunction.FillCombo(cmbEmployee, rsEmployees, 888)
    End If
End With

End Sub

Private Sub cmdExport_Click()
Call GenerateExcelReport
End Sub

Private Sub cmdRefresh_Click()
Loadlist (Filter_var)

If NoTransFlag = True Then
    Me.lblRecount.Caption = "0 record"
    MsgBox "No transaction records selected.", vbInformation, PayrollAlert
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim rsStaff As Recordset

DTPicker1.VALUE = Date
DTPicker2.VALUE = Date
Set rsStaff = cnn.Execute("Select Operator_Name, OperatorID FROM UsersList ORDER BY Operator_Name")
Do While Not rsStaff.EOF
    DoEvents
    cmbStaff.AddItem rsStaff!Operator_name
    cmbStaff.ItemData(cmbStaff.NewIndex) = rsStaff!operatorID
    rsStaff.MoveNext
Loop
rsStaff.Close
cmbStaff.AddItem "--ALL STAFF--", 0

y = 1
Do While y <= 10
    DoEvents
    cmbYear.AddItem 2003 + y
    cmbYear.ItemData(cmbYear.NewIndex) = y
    y = y + 1
Loop

NoTransFlag = True
cmbStaff.ListIndex = 0
Call GetCompanyList
Call GetEmployeeList
Call GetTransactions
End Sub

Private Sub GetCompanyList()
Dim StrQuery As String

'Get all active company.
StrQuery = GetClientRecord(0, True, " ORDER BY CLientName")
Set rsClient = ReturnRst(StrQuery, cnnstr, mstatic)

If rsClient.RecordCount <> 0 Then
    Call FillComboData(cmbCompany, rsClient)
End If
cmbCompany.AddItem "--ALL--", 0
cmbCompany.ListIndex = 0
End Sub

Private Sub GetEmployeeList()
Dim StrQuery As String

StrQuery = "SELECT DISTINCT EmployeeID, RTRIM(LTRIM(surname)) + ', ' + RTRIM(LTRIM(firstname)) + ' ' + (CASE WHEN Middlename IS NULL THEN '' ELSE LEFT(Middlename,1) + '.' END ) AS Fullname, ClientCode " & _
    "FROM HREmployees WHERE EmpStatus = 1"

'Get all Active employees.
Set rsEmployees = ReturnRst(StrQuery, cnnstr, mstatic)

If rsEmployees.RecordCount <> 0 Then
    Call FillComboData(cmbEmployee, rsEmployees)
End If
cmbEmployee.AddItem "--ALL--", 0
cmbEmployee.ListIndex = 0
End Sub

Private Sub GetTransactions()

cmbTransaction.AddItem "--ALL--", 0
cmbTransaction.AddItem "Civil Status change.", 1
cmbTransaction.AddItem "Employment Status change.", 2
cmbTransaction.AddItem "Date Hired change.", 3
cmbTransaction.AddItem "Start Contract change.", 4
cmbTransaction.AddItem "End Contract change.", 5
cmbTransaction.AddItem "Payment Basis change.", 6
cmbTransaction.AddItem "Daily Rate change.", 7
cmbTransaction.AddItem "Monthly Rate change.", 8
cmbTransaction.AddItem "Cash Bond Rate change.", 9
cmbTransaction.AddItem "Accnt# change.", 10
cmbTransaction.AddItem "Client change.", 11
cmbTransaction.ListIndex = 0

End Sub
  
Private Function FillComboData(ByRef mCmb As ComboBox, _
                        mRs As Recordset, Optional FilterFlag As Boolean, _
                        Optional ComboList As String = "")
mCmb.Clear
While Not (mRs.EOF Or mRs.BOF)
    If FilterFlag = False Then
        mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1)), "'", "''"))
        mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
    Else    ' need to filter the lists
        If InStr(1, ComboList, LTrim(RTrim(str(mRs.Fields(0)))), 1) = 0 Then
            mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1)), "'", "''"))
            mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
        End If
    End If
    mRs.MoveNext
Wend

End Function

Public Sub Loadlist(ByVal whereCon As String)
Dim rsEmp1 As ADODB.Recordset

On Error Resume Next

Me.lvwdisplay.ListItems.Clear
    
    Set rsEmp1 = New ADODB.Recordset
    With lvwdisplay
'    StrQuery = "SELECT UserName, LoginTime, LogoutTime, TransDate, TransType, TransDetails, TransStatus, TransRemarks, AgencyName" & _
'             " FROM UserTransactions WHERE " & whereCon & "  ORDER BY userID, TransDate"

    If chkHistory.VALUE = 1 Then
'        StrQuery = "SELECT b.Surname, b.Firstname, b.Middlename as MI, TransDate, TransDetails," & _
'                "TransRemarks, ClientName, UserName, LoginTime, LogoutTime, " & _
'                "MonthlyRate, DailyRate, StartContract, EndContract, b.EmployeeID " & _
'                "FROM UserTransactions a LEFT JOIN (SELECT EmployeeID, Surname, Firstname, Middlename FROM HREmployees) b ON a.EmployeeID = b.EmployeeID " & _
'                "WHERE " & whereCon & " AND TransDetails LIKE '%Rate%' ORDER BY Surname, Firstname, TransDetails, TransDate"
'        StrQuery = "SELECT a.Surname, a.Firstname, a.Middlename as MI, b.sdate AS TransDate, 'Salary History' AS TransDetails, " & _
'            "'Salary History' AS TransRemarks, c.ClientName, 'ADMIN' AS UserName, '10:00AM' AS LoginTime, '11:00AM' AS LogoutTime, " & _
'            "(b.DailyRate * 26) AS MonthlyRate, b.DailyRate, a.StartContract, a.EndContract, a.EmployeeID FROM hremployees a LEFT JOIN ClientTable c ON a.ClientCode = c.ClientCode RIGHT JOIN " & _
'            "(Select employeeid, ROUND(CONVERT(float, SUM(basicpay) / SUM((dayspresent) / 8)), 0) as dailyrate, " & _
'            "CONVERT(datetime, RIGHT('00' + RTRIM(LTRIM(CAST(MONTH(MAX(startpaydate)) as char(2)))),2) + '/' + " & _
'            "RIGHT('00' + RTRIM(LTRIM(CAST(DAY(MAX(startpaydate)) as char(2)))),2) + '/' + CAST(YEAR(MAX(startpaydate)) as char(4))) as sdate " & _
'            "FROM PayrollMaster WHERE (DaysPresent/8) > 0 and employeeid in (SELECT employeeid FROM hremployees WHERE paymentbasis = 'Daily') " & _
'            "GROUP BY EmployeeID, basicpay/(dayspresent/8) ) b on a.employeeid = b.employeeid WHERE b.dailyrate > 0 UNION " & _
'            "SELECT a.Surname, a.Firstname, a.Middlename as MI, b.sdate AS TransDate, 'Salary History' AS TransDetails, " & _
'            "'Salary History' AS TransRemarks, c.ClientName, 'ADMIN' AS UserName, '10:00AM' AS LoginTime, '11:00AM' AS LogoutTime, " & _
'            "(b.DailyRate * 26) AS MonthlyRate, b.DailyRate, a.StartContract, a.EndContract, a.EmployeeID FROM hremployees a LEFT JOIN ClientTable c ON a.ClientCode = c.ClientCode RIGHT JOIN " & _
'            "(Select employeeid, ROUND(CONVERT(float, SUM(basicpay)/(count(basicpay)*13)), 0) as dailyrate, " & _
'            "CONVERT(DateTime, RIGHT('00' + RTRIM(LTRIM(CAST(MONTH(MAX(startpaydate)) as char(2)))),2) + '/' + " & _
'            "RIGHT('00' + RTRIM(LTRIM(CAST(DAY(MAX(startpaydate)) as char(2)))),2) + '/' + CAST(YEAR(MAX(startpaydate)) as char(4))) as sdate " & _
'            "FROM PayrollMaster WHERE (basicpay /13) > 0 and employeeid in (SELECT employeeid FROM hremployees WHERE paymentbasis = 'Monthly') " & _
'            "GROUP BY EmployeeID, basicpay/(dayspresent/8) ) b on a.employeeid = b.employeeid WHERE b.dailyrate > 0 " & _
'            "ORDER BY a.surname, a.firstname, b.dailyrate "
        StrQuery = Replace("SELECT * FROM vwSalaryHistory WHERE " & whereCon & " ORDER BY Surname, Firstname, DailyRate", "b.EmployeeID", "EmployeeID")
    Else
        StrQuery = "SELECT b.Surname, b.Firstname, b.Middlename as MI, TransDate, TransDetails," & _
                "TransRemarks, ClientName, UserName, LoginTime, LogoutTime, " & _
                "MonthlyRate, DailyRate, StartContract, EndContract, b.EmployeeID " & _
                "FROM UserTransactions a LEFT JOIN (SELECT EmployeeID, Surname, Firstname, Middlename FROM HREmployees) b ON a.EmployeeID = b.EmployeeID " & _
                "WHERE " & whereCon & "  ORDER BY userID, TransDate"
    End If
    
    Set rsEmp1 = ReturnRst(StrQuery, CnnstrHR, mstatic)
    If rsEmp1.EOF = True Then
        NoTransFlag = True
        Exit Sub
    Else
        NoTransFlag = False
    End If

    If rsEmp1.EOF = False Then
        PopulateList Me.lvwdisplay, rsEmp1
   '    lvwdisplay.ColumnHeaders(1).Text = "  " & .ListItems.Count & " RECORD/S "
        lvwdisplay.ColumnHeaders(1).Width = 1300
        lvwdisplay.ColumnHeaders(2).Width = 1300
        lvwdisplay.ColumnHeaders(3).Width = 780
        lvwdisplay.ColumnHeaders(4).Width = 2000
        lvwdisplay.ColumnHeaders(5).Width = 4500
        lvwdisplay.ColumnHeaders(6).Width = 2200
        lvwdisplay.ColumnHeaders(7).Width = 3000
        lvwdisplay.ColumnHeaders(8).Width = 980
        lvwdisplay.ColumnHeaders(9).Width = 1980
        lblRecount = .ListItems.Count & " record/s "
        
'        X = 1
'        Do While .ListItems.Count >= X
'             If Val(Format(.ListItems.Item(X).SubItems(5), "d")) <= Val(Format(Date, "d")) Then
'            lvwdisplay.ListItems.Item(X).ForeColor = vbRed
'            lvwdisplay.ListItems(X).ListSubItems(2).ForeColor = vbRed
'            lvwdisplay.ListItems(X).ListSubItems(3).ForeColor = vbRed
'            lvwdisplay.ListItems(X).ListSubItems(4).ForeColor = vbRed
'            lvwdisplay.ListItems(X).ListSubItems(5).ForeColor = vbRed
'
'            End If
'
'            X = X + 1
'        Loop
        
'        strQuery = "select * from hremployees where employmentstatus <> 'Z' and endcontract between #" & Val(Format(Date, "m")) & "/1/" & Val(Format(Date, "yyyy")) & "# and #" & Format(Date, "mm/dd/yyyy") & "#"
'
'        Dim rsHReoc As New adodb.Recordset
'           Set rsHReoc = ReturnRst(strQuery, Cnnstr, mstatic)
'        If rsHReoc.EOF = False Then
'
'         strQuery = "UPDATE hremployees SET employmentstatus = 'C' WHERE  employmentstatus <> 'Z' and endcontract between #" & Val(Format(Date, "m")) & "/1/" & Val(Format(Date, "yyyy")) & "# and #" & Format(Date, "mm/dd/yyyy") & "#"
'
'         Call Action_Query(strQuery, Cnnstr)
'        End If
Else
    Me.lblRecount.Caption = "0 record"
    
    MsgBox "No transaction records selected.", vbInformation, PayrollAlert
    
End If
        End With
        
End Sub

'Private Sub lvwdisplay_DblClick()
'    strrow = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).SubItems(1)
'      If lvwdisplay.ListItems.Count = 0 Then
'        MsgBox "List doesn't have any record from the table selected !", vbOKOnly + vbInformation, "Select..."
'    Exit Sub
'    Else
'    End If
'    Call search(strrow)
'    Me.SetFocus
'End Sub

'Private Sub Search(ByVal strSearch As Integer)
'Dim res, bkmrk
'
'On Error Resume Next
'
'With frmHRFinalEmployees
'   Set .rsout = New ADODB.Recordset
'    StrQuery = "Select * from HREmployees where agencycode= " & CompanyCode & " order by  surname"
'    Set .rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'    '.rsout.Open "Select * from HREmployees where agencycode= " & CompanyCode & " order by  surname", cnnHR, adOpenKeyset, adLockOptimistic
'
'    bkmrk = .rsout.Bookmark
'    .rsout.MoveFirst
'    .rsout.Find "EmployeeID=" & Trim(strSearch) & ""
'    If Not .rsout.EOF Then
'        .Showdata .rsout
'    Else
'        .rsout.Bookmark = bkmrk
'    End If
'End With
'
'End Sub

Private Sub lvwdisplay_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtSearch = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).Text
End Sub

Private Sub lvwdisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Call lvwdisplay_DblClick
End Sub

Private Sub lvwdisplay_KeyUp(KeyCode As Integer, Shift As Integer)
'Call lvwdisplay_DblClick
End Sub

Private Sub GenerateExcelReport()
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mCompany As String, mHead As String
Dim StrQuery As String, i

If chkHistory.VALUE = 1 Then
    StrQuery = Replace("SELECT * FROM vwSalaryHistory WHERE " & whereCon & " ORDER BY Surname, Firstname, DailyRate", "b.EmployeeID", "EmployeeID")
Else
    StrQuery = "SELECT b.Surname, b.Firstname, b.Middlename AS MI, TransDate, TransType, TransDetails," & _
    "TransRemarks, ClientName, UserName, LoginTime, LogoutTime, " & _
    "MonthlyRate, DailyRate, StartContract, EndContract, b.EmployeeID " & _
    "FROM UserTransactions a LEFT JOIN (SELECT EmployeeID, Surname, Firstname, Middlename FROM HREmployees) b ON a.EmployeeID = b.EmployeeID " & _
    "WHERE " & whereCon & " AND TransRemarks NOT IN ('Exit.', 'Closed.') ORDER BY userID, TransDate"
End If
    
Set rsForExcel = ReturnRst(StrQuery, CnnstrHR, mstatic)
If rsForExcel.RecordCount = 0 Then
    Set rsForExcel = Nothing
    MsgBox "No records to process.", vbInformation, "Report Alert"
    Exit Sub
End If

i = 1
ctr = 8

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)
If cmbCompany.Text <> "--ALL--" Then
    mCompany = cmbCompany.Text
Else
    mCompany = "ALL COMPANY"
End If
mHead = "Transaction History Report for (" & mCompany & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 12, True)
If cmbEmployee.Text <> "--ALL--" Then
    mHead = "Record Changes for " & cmbEmployee.Text
Else
    mHead = "Record Changes for All Employees"
End If
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, False)
mHead = "Transaction Date Range from " & DTPicker1.VALUE & " to " & DTPicker2.VALUE
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, False)
If cmbTransaction.Text <> "--ALL--" Then
    mHead = "For " & cmbTransaction
Else
    mHead = "For All Transactions Done"
End If
Call PrintReportHeader(xlSheet, xlApp1, "A5", mHead, "", xlLeft, , False, False, False, , , 10, False)

rsForExcel.MoveFirst

If chkHistory.VALUE = 1 Then
    Call PrintReportHeader(xlSheet, xlApp1, "A7:A7", "Number", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 8)
    Call PrintReportHeader(xlSheet, xlApp1, "B7:B7", "Surname", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 15)
    Call PrintReportHeader(xlSheet, xlApp1, "C7:C7", "FirstName", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 15)
    Call PrintReportHeader(xlSheet, xlApp1, "D7:D7", "M.I.", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 6)
    Call PrintReportHeader(xlSheet, xlApp1, "E7:E7", "Company Name", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 28)
    Call PrintReportHeader(xlSheet, xlApp1, "F7:F7", "Effectivity", "Date", xlCenter, xlBottom, True, False, True, 15, , 9, False, 11)
    Call PrintReportHeader(xlSheet, xlApp1, "G7:G7", "Daily Rate", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 11)
    Call PrintReportHeader(xlSheet, xlApp1, "H7:H7", "Monthly Rate", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 9)
    Call PrintReportHeader(xlSheet, xlApp1, "I7:I7", "Start Contract", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 11)
    Call PrintReportHeader(xlSheet, xlApp1, "J7:J7", "End Contract", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 11)
Else
    Call PrintReportHeader(xlSheet, xlApp1, "A7:A7", "Number", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 8)
    Call PrintReportHeader(xlSheet, xlApp1, "B7:B7", "Surname", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 15)
    Call PrintReportHeader(xlSheet, xlApp1, "C7:C7", "FirstName", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 15)
    Call PrintReportHeader(xlSheet, xlApp1, "D7:D7", "M.I.", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 6)
    Call PrintReportHeader(xlSheet, xlApp1, "E7:E7", "Company Name", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 28)
    Call PrintReportHeader(xlSheet, xlApp1, "F7:F7", "Trans Date", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 11)
    Call PrintReportHeader(xlSheet, xlApp1, "G7:G7", "Trans Details", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 48)
    Call PrintReportHeader(xlSheet, xlApp1, "H7:H7", "Trans Remarks", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 20.5)
    Call PrintReportHeader(xlSheet, xlApp1, "I7:I7", "UserName", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 10)
    Call PrintReportHeader(xlSheet, xlApp1, "J7:J7", "Login Time", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 21)
    Call PrintReportHeader(xlSheet, xlApp1, "K7:K7", "Logout Time", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 21)
'    Call PrintReportHeader(xlSheet, xlApp1, "L7:L7", "Monthly Rate", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 11)
'    Call PrintReportHeader(xlSheet, xlApp1, "M7:M7", "Daily Rate", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 9)
'    Call PrintReportHeader(xlSheet, xlApp1, "N7:N7", "Start Contract", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 11)
'    Call PrintReportHeader(xlSheet, xlApp1, "O7:O7", "End Contract", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 11)
End If

With rsForExcel
Do While Not .EOF
    If chkHistory.VALUE = 1 Then
        xlSheet.Cells(ctr, 1).VALUE = i
        xlSheet.Cells(ctr, 2).VALUE = !Surname
        xlSheet.Cells(ctr, 3).VALUE = !FirstName
        xlSheet.Cells(ctr, 4).VALUE = !MI
        xlSheet.Cells(ctr, 5).VALUE = !ClientName
        xlSheet.Cells(ctr, 6).VALUE = IIf(IsNull(!TransDate), "01/01/1900", Format(!TransDate, "mm/dd/yyyy"))
        xlSheet.Cells(ctr, 7).VALUE = !DailyRate
        xlSheet.Cells(ctr, 7).NumberFormat = "#,##0.00"
        xlSheet.Cells(ctr, 8).VALUE = !MonthlyRate
        xlSheet.Cells(ctr, 8).NumberFormat = "#,##0.00"
        xlSheet.Cells(ctr, 9).VALUE = IIf(IsNull(!StartContract), "01/01/1900", Format(!StartContract, "mm/dd/yyyy"))
        xlSheet.Cells(ctr, 9).NumberFormat = "mm/dd/yyyy"
        xlSheet.Cells(ctr, 10).VALUE = IIf(IsNull(!endcontract), "01/01/1900", Format(!endcontract, "mm/dd/yyyy"))
        xlSheet.Cells(ctr, 10).NumberFormat = "mm/dd/yyyy"
    Else
        xlSheet.Cells(ctr, 1).VALUE = i
        xlSheet.Cells(ctr, 2).VALUE = !Surname
        xlSheet.Cells(ctr, 3).VALUE = !FirstName
        xlSheet.Cells(ctr, 4).VALUE = !MI
        xlSheet.Cells(ctr, 5).VALUE = !ClientName
        xlSheet.Cells(ctr, 6).VALUE = IIf(IsNull(!TransDate), "01/01/1900", Format(!TransDate, "mm/dd/yyyy"))
        xlSheet.Cells(ctr, 6).NumberFormat = "mm/dd/yyyy"
        xlSheet.Cells(ctr, 7).VALUE = !TransDetails
        xlSheet.Cells(ctr, 8).VALUE = !TransRemarks
        xlSheet.Cells(ctr, 9).VALUE = !username
        xlSheet.Cells(ctr, 10).VALUE = IIf(IsNull(!LoginTime), "01/01/1900", Format(!LoginTime, "mm/dd/yyyy hh:mm:sss"))
        xlSheet.Cells(ctr, 10).NumberFormat = "mm/dd/yyyy HH:MM:SS"
        xlSheet.Cells(ctr, 11).VALUE = IIf(IsNull(!LogoutTime), "01/01/1900", Format(!LogoutTime, "mm/dd/yyyy hh:mm:sss"))
        xlSheet.Cells(ctr, 11).NumberFormat = "mm/dd/yyyy hh:mm:ss"
'        xlSheet.Cells(ctr, 12).VALUE = !MonthlyRate
'        xlSheet.Cells(ctr, 12).NumberFormat = "#,##0.00"
'        xlSheet.Cells(ctr, 13).VALUE = !DailyRate
'        xlSheet.Cells(ctr, 13).NumberFormat = "#,##0.00"
'        xlSheet.Cells(ctr, 14).VALUE = IIf(IsNull(!StartContract), "01/01/1900", Format(!StartContract, "mm/dd/yyyy"))
'        xlSheet.Cells(ctr, 14).NumberFormat = "mm/dd/yyyy"
'        xlSheet.Cells(ctr, 15).VALUE = IIf(IsNull(!endcontract), "01/01/1900", Format(!endcontract, "mm/dd/yyyy"))
'        xlSheet.Cells(ctr, 15).NumberFormat = "mm/dd/yyyy"
    End If
    ctr = ctr + 1
    i = i + 1
    .MoveNext
Loop
End With
xlSheet.Range("A1").Select

End Sub

