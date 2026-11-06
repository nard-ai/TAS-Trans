VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmOtherReport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Contributions and Deductions Report"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "frmOtherReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   7515
   Begin VB.Frame Frame4 
      Caption         =   "Select Client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   180
      TabIndex        =   16
      Top             =   3600
      Width           =   3660
      Begin VB.ComboBox cmbClient 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   270
         Width           =   3345
      End
   End
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   3990
      TabIndex        =   15
      Top             =   3600
      Width           =   3315
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   330
         Left            =   1800
         TabIndex        =   27
         Top             =   265
         Width           =   1290
      End
      Begin VB.CommandButton cmdProcess 
         Caption         =   "Generate"
         Height          =   330
         Left            =   240
         TabIndex        =   25
         Top             =   265
         Width           =   1290
      End
   End
   Begin VB.Frame Employees 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2040
      TabIndex        =   14
      Top             =   2700
      Width           =   5280
      Begin VB.CheckBox chkLoans 
         Caption         =   "For Loans"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   1395
      End
      Begin VB.ComboBox cmbEmployee 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmOtherReports.frx":08CA
         Left            =   2760
         List            =   "frmOtherReports.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   280
         Width           =   2400
      End
      Begin VB.CheckBox chkAllEmployees 
         Caption         =   "All Employees"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Selected Only:"
         Height          =   300
         Index           =   3
         Left            =   1680
         TabIndex        =   26
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Report Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2040
      TabIndex        =   11
      Top             =   1875
      Width           =   5280
      Begin VB.ComboBox cmbQuarter 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmOtherReports.frx":08CE
         Left            =   2640
         List            =   "frmOtherReports.frx":08DE
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   315
         Width           =   1185
      End
      Begin VB.ComboBox cmbYearof 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmOtherReports.frx":0916
         Left            =   4425
         List            =   "frmOtherReports.frx":0932
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   315
         Width           =   795
      End
      Begin VB.ComboBox cmbMonthof 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmOtherReports.frx":0966
         Left            =   870
         List            =   "frmOtherReports.frx":098E
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Quarter"
         Height          =   180
         Index           =   0
         Left            =   2040
         TabIndex        =   29
         Top             =   405
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Year"
         Height          =   180
         Index           =   2
         Left            =   3930
         TabIndex        =   22
         Top             =   405
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Month of"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   21
         Top             =   405
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scope of Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   180
      TabIndex        =   7
      Top             =   1875
      Width           =   1815
      Begin VB.OptionButton OptScope 
         Caption         =   "Per Cut-Off"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   30
         Top             =   320
         Width           =   1530
      End
      Begin VB.OptionButton OptScope 
         Caption         =   "Annual Report"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   10
         Top             =   1125
         Width           =   1530
      End
      Begin VB.OptionButton OptScope 
         Caption         =   "Quarterly Report"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   9
         Top             =   855
         Width           =   1530
      End
      Begin VB.OptionButton OptScope 
         Caption         =   "Monthly Report"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   8
         Top             =   585
         Value           =   -1  'True
         Width           =   1410
      End
   End
   Begin VB.Frame FraReport 
      Caption         =   "Type of Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   180
      TabIndex        =   0
      Top             =   675
      Width           =   7125
      Begin VB.OptionButton OptReport 
         Caption         =   "Checklist"
         Height          =   195
         Index           =   9
         Left            =   3120
         TabIndex        =   36
         Top             =   720
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "Deductions"
         Height          =   195
         Index           =   8
         Left            =   4200
         TabIndex        =   35
         Top             =   720
         Width           =   1170
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "HR Head Count"
         Height          =   195
         Index           =   7
         Left            =   1560
         TabIndex        =   34
         Top             =   720
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "Payroll Summary"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   33
         Top             =   720
         Width           =   1530
      End
      Begin VB.ComboBox cmbDeductions 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmOtherReports.frx":09F4
         Left            =   5400
         List            =   "frmOtherReports.frx":09F6
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   600
         Width           =   1560
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "Cash Bond"
         Height          =   195
         Index           =   6
         Left            =   5880
         TabIndex        =   6
         Top             =   300
         Width           =   1170
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "13th Month"
         Height          =   195
         Index           =   5
         Left            =   4530
         TabIndex        =   5
         Top             =   300
         Width           =   1170
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "PhilHealth"
         Height          =   195
         Index           =   4
         Left            =   3270
         TabIndex        =   4
         Top             =   300
         Width           =   1050
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "HDMF"
         Height          =   195
         Index           =   3
         Left            =   2100
         TabIndex        =   3
         Top             =   300
         Width           =   1050
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "Tax"
         Height          =   195
         Index           =   2
         Left            =   1050
         TabIndex        =   2
         Top             =   300
         Width           =   1050
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "SSS "
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1050
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   2
         X1              =   0
         X2              =   800
         Y1              =   1500
         Y2              =   1200
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3945
      Left            =   45
      TabIndex        =   19
      Top             =   600
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   6959
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmOtherReports.frx":09F8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame5 
         Height          =   2295
         Left            =   120
         TabIndex        =   37
         Top             =   3840
         Width           =   7095
         Begin VB.CommandButton cmdPhilHealth 
            Caption         =   "PhilHealth Report"
            Height          =   330
            Left            =   4680
            TabIndex        =   41
            Top             =   1800
            Width           =   2010
         End
         Begin VB.CommandButton cmdAddList 
            Caption         =   "Add to List"
            Height          =   330
            Left            =   240
            TabIndex        =   39
            Top             =   1800
            Width           =   1290
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Records"
            Height          =   330
            Left            =   1920
            TabIndex        =   38
            Top             =   1800
            Width           =   1290
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1455
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   2566
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Client Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Surname"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Given Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Middle Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "PhilHealth No."
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Basic Salary"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Personal Share"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Employer Share"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Label Label2 
         Height          =   240
         Left            =   4230
         TabIndex        =   20
         Top             =   45
         Width           =   1500
      End
   End
End
Attribute VB_Name = "FrmOtherReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StrQuery As String
Dim rsEmpBasicData As ADODB.Recordset
Dim rsClientData As ADODB.Recordset
Dim rsGetDeductionTypes As ADODB.Recordset
Dim rsDeductions As ADODB.Recordset

Dim whereCon As String
Dim ReportScope As Integer
Dim ReportType As Integer
Dim ReportHeader As String
Dim mSapi As String
Dim mShortName As String

Dim ScopeRep As String
Dim mClientCode As Integer
Dim TotalLoan As Double
Dim LoanStart As Integer

'ADDED BY RAMIL 3/27/2007
Dim ESumString As String
Dim FSumString As String
Dim GSumString As String
Dim HSumString As String
Dim ISumString As String
Dim JSumString As String

'ADDED by Bernard 4/19/2008
Dim KSumString As String
Dim LSumString As String
Dim MSumString As String
Dim NSumString As String
Dim PhilHealthFlag As Boolean

'Dim LocalClient As String

Private Sub cmbQuarter_Click()
    ScopeRep = "For the " & cmbQuarter.Text & " of " & cmbYearof.Text
End Sub

Private Sub cmbYearof_Click()
    ScopeRep = "For the Year " & cmbYearof.Text
End Sub

'Private Sub cmdContract_Click()
'Dim mEmpNo As Integer
'Dim rsContract1 As ADODB.Recordset
'Dim rsContract2 As ADODB.Recordset
'
'Me.Height = 5025 ' original height
'
'mEmpNo = cmbEmployee.ItemData(cmbEmployee.ListIndex)
'' Get Current Contract.
'strQuery = "SELECT EmployeeID, EmpNo, AgencyCode, ClientCode, StartContract, EndContract, CBAmount, CBPayment, CBStartPayment, CBEndPayment, Remarks, LastUpdate, ClientName FROM HREmployees WHERE EmployeeID = " & mEmpNo
'Set rsContract1 = ReturnRst(strQuery, cnnStr, mstatic)
'' Get Contract History.
'strQuery = "SELECT EmployeeID, EmpNo, AgencyCode, ClientCode, StartContract, EndContract, CBAmount, CBPayment, CBStartPayment, CBEndPayment, Remarks, LastUpdate, ClientName FROM HREmployees_History WHERE EmployeeID = " & mEmpNo
'Set rsContract2 = ReturnRst(strQuery, CnnstrHR, mstatic)
'
'End Sub

Private Sub FillEmployeeList3(ByRef mLsEmp As ListView, _
                        ByRef mRs As Recordset, Optional AppendSw As Boolean = True)
Dim litem1 As ListItem
Dim Listctr As Integer

    Listctr = 0
    'mLsEmp.View = lvwReport
    ListView1(0).View = lvwReport
    'ListView1(1).ListItems.Clear
    mRs.MoveFirst
    While Not (mRs.EOF Or mRs.BOF)
        Set litem1 = ListView1(0).ListItems.Add(, , mRs!ClientCode)
        litem1.SubItems(1) = IIf(IsNull(mRs!ClientName), "Blank", mRs!ClientName)
        litem1.SubItems(2) = IIf(IsNull(mRs!Surname), "Blank", mRs!Surname)
        litem1.SubItems(3) = IIf(IsNull(mRs!FirstName), "Blank", mRs!FirstName)
        litem1.SubItems(4) = IIf(IsNull(mRs!MiddleName), "Blank", mRs!MiddleName)
        litem1.SubItems(5) = IIf(IsNull(mRs!PhilHealth), "", mRs!PhilHealth)
        litem1.SubItems(6) = Format(mRs!BasicPay, "0.##")
        litem1.SubItems(7) = Format(mRs!PHealthEmployee, "0.##")
        litem1.SubItems(8) = Format(mRs!PHealthEmployer, "0.##")
        Listctr = Listctr + 1
        mRs.MoveNext
    Wend
    MsgBox "There were " & Listctr & " records added.", vbInformation, PayrollAlert
End Sub
'
'Private Sub FillContractList(ByRef mLsEmp As ListView, _
'                        ByRef mRs As Recordset)
'Dim litem1 As ListItem
'Dim litem2 As ListItem
'
'    mLsEmp.View = lvwReport
'    'mLsEmp.ListItems.Clear
'    While Not (mRs.EOF Or mRs.BOF)
'        Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName))
'            litem1.SubItems(1) = mRs(0)
'        mRs.MoveNext
'    Wend
'End Sub
'
'Private Sub cmdAddList_Click()
'    If ReportType = 4 Then
'    PhilHealthFlag = True
'    Call ProcessReports
'    If rsDeductions.RecordCount <> 0 Then
'        Call FillEmployeeList3(ListView1(0), rsDeductions)
'    Else
'        MsgBox "No records to add.", vbInformation, PayrollAlert
'    End If
'    PhilHealthFlag = False
'    Else
'        MsgBox "Select PhilHealth Button first, before you can add records.", vbInformation, PayrollAlert
'    End If
'End Sub

'Private Sub cmdClear_Click()
'    If ReportType = 4 Then
'        ListView1(0).ListItems.Clear
'    Else
'        MsgBox "Select PhilHealth Button first, before you can clear the list.", vbInformation, PayrollAlert
'    End If
'End Sub

'Private Sub cmdPhilHealth_Click()
'    If ReportType = 4 Then
'        Call GeneratePhilHealthReports(rsDeductions)
'    Else
'        MsgBox "Select PhilHealth Button first, before you can generate report.", vbInformation, PayrollAlert
'    End If
'End Sub

Private Sub cmdProcess_Click()
    Call ProcessReports
End Sub

Private Sub ProcessReports()

If optReport(0).VALUE = False And optReport(7).VALUE = False And optReport(9).VALUE = False Then
    Call GetReportScope
    If chkLoans.VALUE = 1 Then
        If ReportType = 1 Then      ' for SSS Loan
            whereCon = whereCon & " and sssloan <> 0"
        ElseIf ReportType = 3 Then  ' for Pagibig Loan
            whereCon = whereCon & " and pagibigloan <> 0 "
        End If
    End If
    mUserID = Trim(UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID))))
        
    If ReportScope <> 0 Then
        StrQuery = "DELETE FROM tmpPayReport WHERE useriD like '%" & mUserID & "%'"
        'strQuery = "DELETE FROM tmpPayReport"
        Call Action_Query(StrQuery, cnnstr)
        StrQuery = "INSERT INTO tmpPayReport " & _
                    "(EmployeeID, EmpNo, FirstName, Middlename, Surname, SSS, Philhealth, " & _
                    "Pagibig, TIN, BasicPay, ClientCode, Month13Payment, TASAesla, LMCDed, ValewithClient, Medical, " & _
                    "SalaryLoan, Uniform, Variances, Vale_IOU, BirthDate, " & _
                    "CashAdvance, WTaxAmount, MiscFee, CalamityLoan, SSSEmployee, SSSEmployer, PHealthEmployee, PHealthEmployer, " & _
                    "PagibigEmployee, PagibigEmployer, SSSLoan, PagibigLoan, OtherLoan, StartPayDate, EndPayDate, MonthNo, CutOffNo, ClientName, MonthOf, UserID) " & _
                    "SELECT PayrollMaster.EmployeeID, PayrollMaster.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, " & _
                    "HREmployees.Pagibig, HREmployees.TIN, PayrollMaster.BasicPay, PayrollMaster.ClientCode, PayrollMaster.Month13Payment, PayrollMaster.TASAesla, PayrollMaster.LMCDed, PayrollMaster.ValewithClient, PayrollMaster.Medical, " & _
                    "PayrollMaster.SalaryLoan, PayrollMaster.Uniform, PayrollMaster.Variances, PayrollMaster.Vale_IOU, HREmployees.BirthDate, " & _
                    "PayrollMaster.CashAdvance, PayrollMaster.WTaxAmount, PayrollMaster.MiscFee, PayrollMaster.CalamityLoan, PayrollMaster.SSSEmployee, PayrollMaster.SSSEmployer, PayrollMaster.PHealthEmployee, PayrollMaster.PHealthEmployer, " & _
                    "PayrollMaster.PagibigEmployee, PayrollMaster.PagibigEmployer, PayrollMaster.SSSLoan, PayrollMaster.PagibigLoan, PayrollMaster.OtherLoan, PayrollMaster.StartPayDate, PayrollMaster.EndPayDate, PayrollMaster.MonthNo, PayrollMaster.CutOffNo, ClientTable.ClientName, DATENAME(month, PayrollMaster.StartPayDate), ' " & mUserID & "' " & _
                    "FROM (PayrollMaster INNER JOIN HREmployees ON (PayrollMaster.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollMaster.ClientCode = ClientTable.ClientCode " & _
                    "WHERE " & whereCon  '" ORDER BY PayrollMaster.ClientCode, PayrollMaster.StartPayDate, PayrollMaster.StartPayDate, HREmployees.Surname, HREmployees.FirstName"
        Call Action_Query(StrQuery, cnnstr)
        
        ' Get also the data from current payroll transactions.
        whereCon = Replace(whereCon, "payrollmaster", "PayrollTrans")
        whereCon = Replace(whereCon, "PayrollMaster", "PayrollTrans")
        StrQuery = "INSERT INTO tmpPayReport " & _
                    "(EmployeeID, EmpNo, FirstName, Middlename, Surname, SSS, Philhealth, " & _
                    "Pagibig, TIN, BasicPay, ClientCode, Month13Payment, TASAesla, LMCDed, ValewithClient, Medical, " & _
                    "SalaryLoan, Uniform, Variances, Vale_IOU, BirthDate, " & _
                    "CashAdvance, WTaxAmount, MiscFee, CalamityLoan, SSSEmployee, SSSEmployer, PHealthEmployee, PHealthEmployer, " & _
                    "PagibigEmployee, PagibigEmployer, SSSLoan, PagibigLoan, OtherLoan, StartPayDate, EndPayDate, MonthNo, CutOffNo, ClientName, MonthOf, UserID) " & _
                    "SELECT PayrollTrans.EmployeeID, PayrollTrans.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, " & _
                    "HREmployees.Pagibig, HREmployees.TIN, PayrollTrans.BasicPay, PayrollTrans.ClientCode, PayrollTrans.Month13Payment, PayrollTrans.TASAesla, PayrollTrans.LMCDed, PayrollTrans.ValewithClient, PayrollTrans.Medical, " & _
                    "PayrollTrans.SalaryLoan, PayrollTrans.Uniform, PayrollTrans.Variances, PayrollTrans.Vale_IOU, HREmployees.BirthDate, " & _
                    "PayrollTrans.CashAdvance, PayrollTrans.WTaxAmount, PayrollTrans.MiscFee, PayrollTrans.CalamityLoan, PayrollTrans.SSSEmployee, PayrollTrans.SSSEmployer, PayrollTrans.PHealthEmployee, PayrollTrans.PHealthEmployer, " & _
                    "PayrollTrans.PagibigEmployee, PayrollTrans.PagibigEmployer, PayrollTrans.SSSLoan, PayrollTrans.PagibigLoan, PayrollTrans.OtherLoan, PayrollTrans.StartPayDate, PayrollTrans.EndPayDate, PayrollTrans.MonthNo, PayrollTrans.CutOffNo, ClientTable.ClientName, DATENAME(month, PayrollTrans.StartPayDate), ' " & mUserID & "' " & _
                    "FROM (PayrollTrans INNER JOIN HREmployees ON (PayrollTrans.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollTrans.ClientCode = ClientTable.ClientCode " & _
                    "WHERE " & whereCon  '" ORDER BY PayrollTrans.ClientCode, PayrollTrans.StartPayDate, PayrollTrans.StartPayDate, HREmployees.Surname, HREmployees.FirstName"
        Call Action_Query(StrQuery, cnnstr)
        
'        If chkLoans.VALUE = 1 Or chkLoans.VALUE = 4 Or chkLoans.VALUE = 3 Then
'            ' Delete all endcutoff date >= 16 with keyword (Backlog).
'            strQuery = "DELETE FROM tmpPayReport WHERE Format(EndPayDate, ""dd"") >= 16 AND Format(EndPayDate, ""mmmm"") = '" & cmbMonthof.Text & "' and Format(EndPayDate, ""yyyy"") = " & cmbYearof.Text & " AND CutOffDate like '%Backlog%' AND useriD like '%" & mUserID & "%'"
'            Call Action_Query(strQuery, cnnStr)
'            ' Get all the corresponding payroll transactions from previous month.
'            strQuery = "INSERT INTO tmpPayReport " & _
'                        "(EmployeeID, EmpNo, FirstName, Middlename, Surname, SSS, Philhealth, " & _
'                        "Pagibig, TIN, BasicPay, ClientCode, Month13Payment, TASAesla, LMCDed, ValewithClient, Medical, " & _
'                        "SalaryLoan, Uniform, Variances, Vale_IOU, BirthDate, " & _
'                        "CashAdvance, WTaxAmount, MiscFee, CalamityLoan, SSSEmployee, SSSEmployer, PHealthEmployee, PHealthEmployer, " & _
'                        "PagibigEmployee, PagibigEmployer, SSSLoan, PagibigLoan, OtherLoan, StartPayDate, EndPayDate, MonthNo, CutOffNo, ClientName, MonthOf, UserID) " & _
'                        "SELECT PayrollMaster.EmployeeID, PayrollMaster.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, " & _
'                        "HREmployees.Pagibig, HREmployees.TIN, PayrollMaster.BasicPay, PayrollMaster.ClientCode, PayrollMaster.Month13Payment, PayrollMaster.TASAesla, PayrollMaster.LMCDed, PayrollMaster.ValewithClient, PayrollMaster.Medical, " & _
'                        "PayrollMaster.SalaryLoan, PayrollMaster.Uniform, PayrollMaster.Variances, PayrollMaster.Vale_IOU, HREmployees.BirthDate, " & _
'                        "PayrollMaster.CashAdvance, PayrollMaster.WTaxAmount, PayrollMaster.MiscFee, PayrollMaster.CalamityLoan, PayrollMaster.SSSEmployee, PayrollMaster.SSSEmployer, PayrollMaster.PHealthEmployee, PayrollMaster.PHealthEmployer, " & _
'                        "PayrollMaster.PagibigEmployee, PayrollMaster.PagibigEmployer, PayrollMaster.SSSLoan, PayrollMaster.PagibigLoan, PayrollMaster.OtherLoan, PayrollMaster.StartPayDate, PayrollMaster.EndPayDate, PayrollMaster.MonthNo, PayrollMaster.CutOffNo, ClientTable.ClientName, FORMAT(PayrollMaster.StartPayDate, ""mmmm""), ' " & mUserID & "' " & _
'                        "FROM (PayrollMaster INNER JOIN HREmployees ON (PayrollMaster.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollMaster.ClientCode = ClientTable.ClientCode " & _
'                        " WHERE Format(EndPayDate, ""dd"") >= 16 AND Format(StartPayDate, ""mm"") = " & getMonthNum(cmbMonthof.Text, cmbYearof.Text) & " AND Format(StartPayDate, ""yyyy"") = " & getMonthNum(cmbMonthof.Text, cmbYearof.Text, False) & _
'                        " AND ClientTable.CutOffPeriod like '%Backlog%'"
'            Call Action_Query(strQuery, cnnStr)
'        End If
        
        If ReportType = 1 Then  ' for SSS Loan, Order by Client Code + Name
            StrQuery = "SELECT tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, min(tmpPayReport.SSS) as SSS, " & _
                    "MIN(tmpPayReport.Philhealth) as PhilHealth, min(tmpPayReport.Pagibig) as pagibig, min(tmpPayReport.TIN) as TIN, tmpPayReport.ClientCode, tmppayreport.monthof, " & _
                    "SUM(tmpPayReport.Month13Payment) as month13payment, sum(tmpPayReport.TASAesla) as TASAesla, sum(tmpPayReport.LMCDed) as LMCDed, " & _
                    "SUM(tmpPayReport.ValewithClient) as valewithclient, sum(tmpPayReport.CashAdvance) as CashAdvance, sum(tmpPayReport.WTaxAmount) as WTaxAmount, " & _
                    "SUM(tmpPayReport.Medical) as Medical, sum(tmpPayReport.SalaryLoan) as SalaryLoan, sum(tmpPayReport.Uniform) as Uniform, sum(tmpPayReport.Variances) as Variances, sum(tmpPayReport.Vale_IOU) as Vale_IOU, " & _
                    "SUM(tmpPayReport.MiscFee) as miscfee, sum(tmpPayReport.CalamityLoan) as CalamityLoan, sum(tmpPayReport.SSSEmployee) as SSSEmployee, " & _
                    "SUM(tmpPayReport.SSSEmployer) as sssemployer, SUM(tmppayreport.sssloan) as sssloan, SUM(tmppayreport.pagibigloan) as pagibigloan, sum(tmpPayReport.PHealthEmployee) as phealthemployee, sum(tmpPayReport.PHealthEmployer) as PHealthEmployer, " & _
                    "SUM(tmpPayReport.PagibigEmployee) as pagibigemployee, sum(tmpPayReport.PagibigEmployer) as pagibigemployer, min(tmpPayReport.MonthNo) as MonthNo, " & _
                    "MIN(tmpPayReport.CutOffNo) as cutoffno, min(tmpPayReport.ClientName) as clientname FROM tmpPayReport " & _
                    "WHERE LTRIM(RTRIM(userID)) = '" & mUserID & "' " & _
                    "GROUP BY tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, tmpPayReport.ClientCode, tmppayreport.monthof, tmppayreport.monthno " & _
                    "ORDER BY tmppayreport.ClientCode, tmppayreport.Surname, tmppayreport.FirstName, tmppayreport.monthno"
        Else
            StrQuery = "SELECT tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, min(tmpPayReport.SSS) as SSS, " & _
                    "MIN( tmpPayReport.BirthDate) as BirthDate, MIN(tmpPayReport.Philhealth) as PhilHealth, min(tmpPayReport.Pagibig) as pagibig, min(tmpPayReport.TIN) as TIN, tmpPayReport.ClientCode, tmppayreport.monthof, " & _
                    "SUM(tmpPayReport.BasicPay) as BasicPay, SUM(tmpPayReport.Month13Payment) as month13payment, sum(tmpPayReport.TASAesla) as TASAesla, sum(tmpPayReport.LMCDed) as LMCDed, " & _
                    "SUM(tmpPayReport.ValewithClient) as valewithclient, sum(tmpPayReport.CashAdvance) as CashAdvance, sum(tmpPayReport.WTaxAmount) as WTaxAmount, " & _
                    "SUM(tmpPayReport.Medical) as Medical, sum(tmpPayReport.SalaryLoan) as SalaryLoan, sum(tmpPayReport.Uniform) as Uniform, sum(tmpPayReport.Variances) as Variances, sum(tmpPayReport.Vale_IOU) as Vale_IOU, " & _
                    "SUM(tmpPayReport.MiscFee) as miscfee, sum(tmpPayReport.CalamityLoan) as CalamityLoan, sum(tmpPayReport.SSSEmployee) as SSSEmployee, " & _
                    "SUM(tmpPayReport.SSSEmployer) as sssemployer, SUM(tmppayreport.sssloan) as sssloan, SUM(tmppayreport.pagibigloan) as pagibigloan, sum(tmpPayReport.PHealthEmployee) as phealthemployee, sum(tmpPayReport.PHealthEmployer) as PHealthEmployer, " & _
                    "SUM(tmpPayReport.PagibigEmployee) as pagibigemployee, sum(tmpPayReport.PagibigEmployer) as pagibigemployer, min(tmpPayReport.MonthNo) as MonthNo, " & _
                    "MIN(tmpPayReport.CutOffNo) as cutoffno, min(tmpPayReport.ClientName) as clientname FROM tmpPayReport " & _
                    "WHERE LTRIM(RTRIM(userID)) = '" & mUserID & "' " & _
                    "GROUP BY tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, tmpPayReport.ClientCode, tmppayreport.monthof, tmppayreport.monthno " & _
                    "ORDER BY tmppayreport.ClientCode, tmppayreport.monthno, tmppayreport.Surname, tmppayreport.FirstName"
        End If
    Else
        StrQuery = "SELECT PayrollMaster.EmployeeID, PayrollMaster.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, " & _
                "HREmployees.Pagibig, HREmployees.TIN, PayrollMaster.ClientCode, PayrollMaster.Month13Payment, PayrollMaster.TASAesla, PayrollMaster.LMCDed, PayrollMaster.ValewithClient, " & _
                "PayrollMaster.CashAdvance, PayrollMaster.WTaxAmount, PayrollMaster.MiscFee, PayrollMaster.CalamityLoan, PayrollMaster.SSSEmployee, PayrollMaster.SSSEmployer, PayrollMaster.PHealthEmployee, PayrollMaster.PHealthEmployer, " & _
                "PayrollMaster.Medical, PayrollMaster.SalaryLoan, PayrollMaster.Uniform, PayrollMaster.Variances, PayrollMaster.Vale_IOU, " & _
                "PayrollMaster.PagibigEmployee, PayrollMaster.PagibigEmployer, PayrollMaster.SSSLoan, PayrollMaster.PagibigLoan, PayrollMaster.StartPayDate, PayrollMaster.EndPayDate, PayrollMaster.MonthNo, PayrollMaster.CutOffNo, ClientTable.ClientName " & _
                "FROM (PayrollMaster INNER JOIN HREmployees ON (PayrollMaster.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollMaster.ClientCode = ClientTable.ClientCode " & _
                "WHERE " & whereCon & " ORDER BY PayrollMaster.ClientCode, PayrollMaster.StartPayDate, PayrollMaster.EndPayDate, HREmployees.Surname, HREmployees.FirstName"
    
'        If chkLoans.VALUE = 1 Or chkLoans.VALUE = 4 Or chkLoans.VALUE = 3 Then
'            ' Delete all endcutoff date >= 16 with keyword (Backlog).
'            strQuery = "DELETE FROM tmpPayReport WHERE Format(EndPayDate, ""dd"") >= 16 AND Format(EndPayDate, ""mmmm"") = '" & cmbMonthof.Text & "' and Format(EndPayDate, ""yyyy"") = " & cmbYearof.Text & " AND CutOffDate like '%Backlog%' AND useriD like '%" & mUserID & "%'"
'            Call Action_Query(strQuery, cnnStr)
'            ' Get all the corresponding payroll transactions from previous month.
'            strQuery = "INSERT INTO tmpPayReport " & _
'                        "(EmployeeID, EmpNo, FirstName, Middlename, Surname, SSS, Philhealth, " & _
'                        "Pagibig, TIN, BasicPay, ClientCode, Month13Payment, TASAesla, LMCDed, ValewithClient, Medical, " & _
'                        "SalaryLoan, Uniform, Variances, Vale_IOU, BirthDate, " & _
'                        "CashAdvance, WTaxAmount, MiscFee, CalamityLoan, SSSEmployee, SSSEmployer, PHealthEmployee, PHealthEmployer, " & _
'                        "PagibigEmployee, PagibigEmployer, SSSLoan, PagibigLoan, OtherLoan, StartPayDate, EndPayDate, MonthNo, CutOffNo, ClientName, MonthOf, UserID) " & _
'                        "SELECT PayrollMaster.EmployeeID, PayrollMaster.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, " & _
'                        "HREmployees.Pagibig, HREmployees.TIN, PayrollMaster.BasicPay, PayrollMaster.ClientCode, PayrollMaster.Month13Payment, PayrollMaster.TASAesla, PayrollMaster.LMCDed, PayrollMaster.ValewithClient, PayrollMaster.Medical, " & _
'                        "PayrollMaster.SalaryLoan, PayrollMaster.Uniform, PayrollMaster.Variances, PayrollMaster.Vale_IOU, HREmployees.BirthDate, " & _
'                        "PayrollMaster.CashAdvance, PayrollMaster.WTaxAmount, PayrollMaster.MiscFee, PayrollMaster.CalamityLoan, PayrollMaster.SSSEmployee, PayrollMaster.SSSEmployer, PayrollMaster.PHealthEmployee, PayrollMaster.PHealthEmployer, " & _
'                        "PayrollMaster.PagibigEmployee, PayrollMaster.PagibigEmployer, PayrollMaster.SSSLoan, PayrollMaster.PagibigLoan, PayrollMaster.OtherLoan, PayrollMaster.StartPayDate, PayrollMaster.EndPayDate, PayrollMaster.MonthNo, PayrollMaster.CutOffNo, ClientTable.ClientName, FORMAT(PayrollMaster.StartPayDate, ""mmmm""), ' " & mUserID & "' " & _
'                        "FROM (PayrollMaster INNER JOIN HREmployees ON (PayrollMaster.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollMaster.ClientCode = ClientTable.ClientCode " & _
'                        " WHERE Format(EndPayDate, ""dd"") >= 16 AND Format(StartPayDate, ""mm"") = " & getMonthNum(cmbMonthof.Text, cmbYearof.Text) & " AND Format(StartPayDate, ""yyyy"") = " & getMonthNum(cmbMonthof.Text, cmbYearof.Text, False) & _
'                        " AND ClientTable.CutOffPeriod like '%Backlog%'"
'            Call Action_Query(strQuery, cnnStr)
'        End If
    End If
    
    Set rsDeductions = ReturnRst(StrQuery, cnnstr, mstatic)
    If PhilHealthFlag = True Then
        GoTo ExitRtn1
    End If

Else  ' Payroll Summary Report
    If optReport(9).VALUE = True Then
        If cmbClient.Text = "--ALL CLIENTS--" Then ' if all clients, do not proceed
            MsgBox "Please select specific client only! Cannot process.", vbInformation
            GoTo ExitRtn1
        End If
    End If
    Set rsDeductions = GetPayrollTrans(False, mClientCode, cmbMonthof.Text, cmbYearof.Text)
End If
    If rsDeductions.RecordCount = 0 Then
        MsgBox "No records gathered from master database. Cannot proceed!", vbInformation, PayrollAlert
        Exit Sub
    End If
        
    ' Check if selected employees only, then filter the recordset.
    If chkAllEmployees.VALUE = 0 Then
        With rsDeductions
        If cmbClient.Text <> "--ALL CLIENTS--" Then ' check for all clients
            .Filter = "EmployeeID" & " = " & "'" & Trim(cmbEmployee.ItemData(cmbEmployee.ListIndex)) & "'"
        'Else
        '    .Filter = "EmployeeID" & " = " & "'" & Trim(cmbEmployee.ItemData(cmbEmployee.ListIndex)) & "'"
        End If
        If .EOF Then
            MsgBox "No records left to process! Cannot process reports!", vbInformation, PayrollAlert
            .Filter = ""
            Exit Sub
        End If
        End With
    End If
    If optReport(7).VALUE = True Then ' for Head count report
        Call PrintHeadCountReport(rsDeductions, cmbMonthof.Text, cmbYearof.Text)
    ElseIf optReport(9).VALUE = True Then ' for payroll checklist report
        Call PrintChecklistReport(rsDeductions, cmbMonthof.Text, cmbYearof.Text)
    ElseIf optReport(0).VALUE = False Then  ' for all deductions report
        Call GenerateReports(rsDeductions)
    Else    ' Payroll Summary
        Call PrintPayrollSummary(rsDeductions, 4)
    End If

ExitRtn1:

End Sub

Private Function getMonthNum(ByVal mMonth As String, ByVal mYear As Integer, Optional mSw As Boolean = True) As Integer
Dim mYear2 As Integer

Select Case mMonth
    Case "January"
        getMonthNum = 12
        mYear2 = mYear - 1
    Case "February"
        getMonthNum = 1
        mYear2 = mYear
    Case "March"
        getMonthNum = 2
        mYear2 = mYear
    Case "April"
        getMonthNum = 3
        mYear2 = mYear
    Case "May"
        getMonthNum = 4
        mYear2 = mYear
    Case "June"
        getMonthNum = 5
        mYear2 = mYear
    Case "July"
        getMonthNum = 6
        mYear2 = mYear
    Case "August"
        getMonthNum = 7
        mYear2 = mYear
    Case "September"
        getMonthNum = 8
        mYear2 = mYear
    Case "October"
        getMonthNum = 9
        mYear2 = mYear
    Case "November"
        getMonthNum = 10
        mYear2 = mYear
    Case "December"
        getMonthNum = 11
        mYear2 = mYear
End Select

If mSw = False Then ' return the year
    getMonthNum = mYear2
End If

End Function

Public Sub PrintReportHeader1(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal mRange As String, ByVal Head1 As String, Optional Head2 As String = "", Optional HorAlign As Variant = xlCenter, _
    Optional VerAlign As Variant = xlCenter, Optional WrapText As Boolean = True, _
    Optional ShrinkToFit As Boolean = False, Optional MergedCells As Boolean = False, _
    Optional ColorIndex As Variant = xlNone, Optional FontName As String = "Arial", _
    Optional FontSize As Integer = 8, Optional FontBold As Boolean = False, Optional ColumnWidth As Integer = 7.5)
    
    xlSheet.Range(mRange).Select
    If Head2 <> "" Then
        xlApp1.ActiveCell.FormulaR1C1 = Head1 & Chr(10) & Head2
    Else
        xlApp1.ActiveCell.FormulaR1C1 = Head1
    End If
    With xlApp1
        .Selection.Interior.ColorIndex = ColorIndex '15 or xlnone
        .Selection.HorizontalAlignment = HorAlign 'xlCenter
        .Selection.VerticalAlignment = VerAlign 'xlCenter
        .Selection.WrapText = WrapText 'True
        .Selection.Orientation = 0
        .Selection.AddIndent = False
        .Selection.ShrinkToFit = ShrinkToFit 'False
        .Selection.MergeCells = MergedCells 'False
    End With
    With xlSheet.Range(mRange)
        .Font.Name = FontName '"Arial"    'Font.Name
        .Font.Size = FontSize '14         'Font.size
        .Font.Bold = FontBold 'True       'Font.Bold
    End With
    With xlSheet.Columns.Range(mRange)
        .ColumnWidth = ColumnWidth '7.5
    End With
End Sub

Private Sub PutBox(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal mRange As String)

'xlSheet.Range(mRange).Select   ' original code
xlApp1.ActiveSheet.Range(mRange).Select
With xlApp1
    '.Selection.Columns.AutoFit
    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With xlApp1.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End With
With xlSheet.Range(mRange)
    .Font.Name = "Arial"   'Font.Name
    .Font.Size = 8         'Font.size
End With

End Sub

Private Function GetPayrollTrans(ByVal PayrollSw As Boolean, _
    Optional ClientCd As Integer, Optional mMonth As String = "January", Optional mYear As Integer = 2008) As ADODB.Recordset
Dim rsPayTrans As ADODB.Recordset
Dim StartDate As String
Dim EndDate As String
'Dim WhereCon As String

If PayrollSw = True Then    ' for current payroll processing
    StartDate = frmSelectCutoff.txtCutoffDate(0).Text
    EndDate = frmSelectCutoff.txtCutoffDate(1).Text
    whereCon = "WHERE a.clientcode = " & mClientCode & _
        " and a.startpaydate = '" & StartDate & "' and a.endpaydate = '" & EndDate & "'" & _
        " ORDER BY a.employeename"
Else    ' for payroll summary per cut-off
    If ClientCd <> 0 Then ' if specific Clients
        mClientCode = ClientCd
'        If CutOffPeriod = "START" Then
'            whereCon = "WHERE a.clientcode = " & mClientCode & _
'                " and format(a.startpaydate, ""mmmm"") = '" & mMonth & "'" & _
'                " and format(a.startpaydate, ""yyyy"") = " & mYear & _
'                " ORDER BY a.startpaydate, a.endpaydate, a.employeename"
'        ElseIf CutOffPeriod = "END" Then    ' Cut-off period is END
'            whereCon = "WHERE a.clientcode = " & mClientCode & _
'                " and format(a.endpaydate, ""mmmm"") = '" & mMonth & "'" & _
'                " and format(a.endpaydate, ""yyyy"") = " & mYear & _
'                " ORDER BY a.startpaydate, a.endpaydate, a.employeename"
'        Else    ' Based from Payroll process Date
            whereCon = "WHERE a.clientcode = " & mClientCode & _
                " and DATENAME(month, a.EntryDate) = '" & mMonth & "'" & _
                " and YEAR(a.EntryDate) = " & mYear & _
                " ORDER BY a.startpaydate, a.endpaydate, a.employeename"

 '               " and format(a.entrydate, ""mmmm"") = '" & mMonth & "'" &
 '               " and format(a.entrydate, ""yyyy"") = " & mYear &

'        End If
    Else        ' ALL Clients
        If CutOffPeriod = "START" Then
            whereCon = "WHERE DATENAME(month, a.startpaydate) = '" & mMonth & "'" & _
                    " and YEAR(a.startpaydate) = " & mYear & _
                    " ORDER BY d.clientname, a.startpaydate, a.endpaydate, a.employeename"
        Else    ' Cut-off period is END
            whereCon = "WHERE DATENAME(month, a.EntryDate) = '" & mMonth & "'" & _
                    " and YEAR(a.EntryDate) = " & mYear & _
                    " ORDER BY d.clientname, a.startpaydate, a.endpaydate, a.employeename"
        
'            whereCon = "WHERE format(a.endpaydate, ""mmmm"") = '" & mMonth & "'" & _
'                    " and format(a.endpaydate, ""yyyy"") = " & mYear & _
'                    " ORDER BY d.clientname, a.startpaydate, a.endpaydate, a.employeename"
        
        End If
    End If
End If

Set rsPayTrans = New ADODB.Recordset

If optReport(7).VALUE = True Then   ' Head Count Report
    mUserID = Trim(UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID))))
    StrQuery = "DELETE FROM tmpHRCount WHERE useriD like '%" & mUserID & "%'"
    Call Action_Query(StrQuery, cnnstr)
    StrQuery = "DELETE FROM tmpPayrollCount WHERE useriD like '%" & mUserID & "%'"
    Call Action_Query(StrQuery, cnnstr)
    StrQuery = "INSERT INTO tmpHRCount (AgencyCode, ClientCode, ClientName, EmpCount, StartContract, EndContract) " & _
        "SELECT AgencyCode, ClientCode, ClientName, EmpCount, StartContract, EndContract FROM qryHRCount "
    Call Action_Query(StrQuery, cnnstr)
    StrQuery = "UPDATE tmpHRCount SET tmpHRCOUNT.UserID = '" & mUserID & "' WHERE userID IS NULL"
    Call Action_Query(StrQuery, cnnstr)
    
    ' Insert the Head Count from Master Payroll table.
    StrQuery = "INSERT INTO tmpPayrollCount (AgencyCode, ClientCode, HeadCount, PayMonth, PayYear) " & _
        "SELECT AgencyCode, ClientCode, HeadCount, StartPayDate, StartPayYear FROM qryPayMastCount " & _
        " WHERE StartPayDate = '" & mMonth & "' and StartPayYear = " & mYear
    Call Action_Query(StrQuery, cnnstr)
    ' Insert the Head Count from Temporary Payroll table.
    StrQuery = "INSERT INTO tmpPayrollCount (AgencyCode, ClientCode, HeadCount, PayMonth, PayYear) " & _
        "SELECT AgencyCode, ClientCode, HeadCount, StartPayDate, StartPayYear FROM qryPayTransCount " & _
        " WHERE StartPayDate = '" & mMonth & "' and StartPayYear = " & mYear
    Call Action_Query(StrQuery, cnnstr)
    
    StrQuery = "UPDATE tmpPayrollCount SET UserID = '" & mUserID & "' WHERE userID IS NULL"
    Call Action_Query(StrQuery, cnnstr)
    
    StrQuery = "UPDATE tmpHRCount INNER JOIN tmpPayrollCount " & _
        " ON (tmpHRCount.ClientCode = tmpPayrollCount.ClientCode) " & _
        " AND (tmpHRCount.AgencyCode = tmpPayrollCount.AgencyCode) " & _
        " SET tmpHRCount.HeadCount = tmpHRCount.HeadCount + tmppayrollcount.HeadCount, tmpHRCount.PayMonth = tmpPayrollCount.PayMonth, tmpHRCount.PayYear = tmpPayrollCount.PayYear"
    Call Action_Query(StrQuery, cnnstr)
    StrQuery = "SELECT AgencyCode, ClientCode, ClientName, SUM(EmpCount) As HRCount, MAX(HeadCount) AS TotalCount " & _
        " FROM tmpHRCount WHERE userID = '" & mUserID & "' GROUP BY AgencyCode, ClientCode, ClientName ORDER BY AgencyCode, ClientName "
ElseIf optReport(9).VALUE = True Then
    ' For Payroll Checklist
    mUserID = Trim(UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID))))
    StrQuery = "DELETE FROM tmpPayrollChecklist WHERE useriD like '%" & mUserID & "%'"
    Call Action_Query(StrQuery, cnnstr)
    StrQuery = "INSERT INTO tmpPayrollChecklist (AgencyCode, ClientCode, EmployeeID, AgencyBDOCardFee, SSSLoanAmount, SSSLoanPayment, StartSSSPayment, LastSSSPayment, PbigLoanAmount, PbigLoanPayment, StartPbigPayment, LastPbigPayment, ValeLoanAmount, ValeLoanPayment, StartValePayment, LastValePayment, UpdatedBy, DateUpdated, EmpNo, EmpName, DailyRate, StartContract, EndContract, CBAmount, CBPayment, CBStartPayment, CBEndPayment, MiscFee, MFPayment, MFStartPayment, MFEndPayment, CashAdvance, PFPayment, PFStartPayment, PFENdPayment, LastUpdate, AccountNo, CutOffPeriod, SalaryRelease, Remarks) " & _
        "SELECT AgencyCode, ClientCode, EmployeeID, AgencyBDOCardFee, SSSLoanAmount, SSSLoanPayment, StartSSSPayment, LastSSSPayment, PbigLoanAmount, PbigLoanPayment, StartPbigPayment, LastPbigPayment, ValeLoanAmount, ValeLoanPayment, StartValePayment, LastValePayment, UpdatedBy, DateUpdated, EmpNo, EmpName, DailyRate, StartContract, EndContract, CBAmount, CBPayment, CBStartPayment, CBEndPayment, MiscFee, MFPayment, MFStartPayment, MFEndPayment, CashAdvance, PFPayment, PFStartPayment, PFENdPayment, LastUpdate, AccountNo, CutOffPeriod, SalaryRelease, Remarks " & _
        " From qryDeductionChecklist WHERE ClientCode = " & mClientCode & " AND EmploymentStatus IN ('C')"
    Call Action_Query(StrQuery, cnnstr)
    StrQuery = "UPDATE tmpPayrollChecklist SET UserID = '" & mUserID & "' WHERE userID IS NULL"
    Call Action_Query(StrQuery, cnnstr)
    StrQuery = "SELECT DISTINCT * FROM tmpPayrollChecklist WHERE ClientCode = " & mClientCode & " AND userID = '" & mUserID & "' ORDER BY EmpName"
Else    ' Payroll Summary Report
    ' Read the Master payroll table.
    StrQuery = "SELECT a.EmpNo as controlnum, a.*, b.acctno, c.*, d.ClientName, d.ClientSeriesNo " & _
        "FROM ((PayrollMaster AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID) " & _
        "LEFT JOIN qryEmployeeBasics AS c ON a.EmployeeID = c.EmployeeID) LEFT JOIN ClientTable AS d ON a.clientcode = d.clientcode " & whereCon
End If

Set rsPayTrans = ReturnRst(StrQuery, cnnstr, mstatic)

'Check if there were records.
If rsPayTrans.EOF = True Then
    Set rsPayTrans = Nothing
    'Set rsPayTrans = New ADODB.Recordset
    
    StrQuery = "SELECT a.EmpNo as controlnum, a.*, b.acctno, c.*, d.ClientName, d.ClientSeriesNo " & _
        "FROM ((PayrollTrans AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID) " & _
        "LEFT JOIN qryEmployeeBasics AS c ON a.EmployeeID = c.EmployeeID) LEFT JOIN ClientTable AS d ON a.clientcode = d.clientcode " & whereCon

    Set rsPayTrans = ReturnRst(StrQuery, cnnstr, mstatic)
End If
Set GetPayrollTrans = rsPayTrans
End Function

Private Sub chkAllEmployees_Click()
If chkAllEmployees.VALUE = 0 Then
    cmbEmployee.Enabled = True
Else
    cmbEmployee.Enabled = False
End If
End Sub

Private Sub cmbClient_Click()
Call FilterEmployee(cmbClient.Text)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    'Populate Clients.
    StrQuery = "SELECT * FROM clienttable WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
    Set rsClientData = ReturnRst(StrQuery, cnnstr, mstatic)
    
    ReportType = 1
    ReportScope = 1
    PhilHealthFlag = False
    cmbMonthof.Text = Format(Date, "mmmm")
    ReportHeader = "SSS Contributions Report"
    cmbQuarter.Text = "1st Quarter"
    cmbQuarter.Enabled = False
    cmbYearof.Text = Format(Date, "yyyy")
    ScopeRep = "For the Month of "
    
    mSapi = CompanyName & " (" & CompanyShort & ")"
    mShortName = "(" & CompanyShort & ")"
    
'    If CompanyCode = 1 Then
'        mSapi = "Search Asia Placement Inc. (SAPI)"
'    ElseIf CompanyCode = 2 Then
'        mSapi = "Jobs & Careers Inc. (JCI)"
'    Else
'        mSapi = "Corporate Staffers Placement Inc. (CSPI)"
'    End If
   
    StrQuery = "SELECT a.*, " & _
            "b.EmpNo AS EmpNo2, b.sss, b.philhealth, b.pagibig, b.acctno, b.TIN, c.employmentstatusdesc, d.clientname " & _
            "FROM ((qryEmployeebasics AS a LEFT JOIN HREmployees AS b ON a.EmployeeID = b.EmployeeID) " & _
            "LEFT JOIN HREmploymentStatus AS c ON b.EmploymentStatus = c.EmploymentStatus) " & _
            "LEFT JOIN ClientTable AS d ON a.clientcode = d.clientcode " & _
            " ORDER BY b.surname, b.firstname"

    Set rsEmpBasicData = ReturnRst(StrQuery, cnnstr, mstatic)
    Call oCommonFunction.FillCombo(cmbClient, rsClientData, 9) ' 9 means add "All Clients"
    If rsClientData.RecordCount <> 0 Then cmbClient.ListIndex = 0
    
    ' Get List of deductions.
    StrQuery = "SELECT * FROM DeductionTypes WHERE DeductionCode <> 0 ORDER BY DeductionDescription"
    Set rsGetDeductionTypes = ReturnRst(StrQuery, CnnstrHR, mstatic)
        
    ' Fill the combo with New Deductions.
    Call FillDeductions(cmbDeductions, rsGetDeductionTypes)
        
    'LocalClient = cmbClient.Text
    
End Sub

Private Function FillDeductions(ByRef mCmb As ComboBox, _
                        mRs As Recordset)
    mCmb.Clear
    mCmb.AddItem "All Deductions"
    While Not (mRs.EOF Or mRs.BOF)
        mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1)), "'", "''"))
        mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
        mRs.MoveNext
    Wend
End Function

Private Sub FilterEmployee(ByVal mCompany As String)
Dim mCCode As Integer
    
    mClientCode = cmbClient.ItemData(cmbClient.ListIndex)
    rsClientData.MoveFirst
    rsClientData.Find "clientname = '" & Trim(mCompany) & "'"
    If rsClientData.EOF = True Then
        mCCode = 0 ' updated by bernard 3/9/2009
        'Exit Sub
    Else
        mCCode = rsClientData!ClientCode
    End If
    With rsEmpBasicData
        If mCCode <> 0 Then
            .Filter = " clientcode = " & mCCode
        Else
            .Filter = ""
        End If
        If .RecordCount <> 0 Then
            'If Not .EOF And Not .BOF Then
            'Call oCommonFunction.FillCombo(cmbClient, rsClientData)
            Call oCommonFunction.FillCombo(cmbEmployee, rsEmpBasicData, 61)
            'Call FillEmployeeList(ListView1(0), rsEmpBasicData)
            'rsEmpBasicData.MoveFirst
            'lblclient(10).Caption = "Total:  " & ActiveCtr
            '.Filter = "clientcode = " & mCCode
        Else
            Call oCommonFunction.FillCombo(cmbEmployee, rsEmpBasicData, 999)
            'Call ClearAllData
            'Call clearList
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
rsEmpBasicData.Close
Set rsEmpBasicData = Nothing
End Sub

'Private Sub Option3_Click()
'If Option3.VALUE = True Then
'    cmbDeductions.Enabled = True
'Else
'    cmbDeductions.Enabled = False
'End If
'End Sub

Private Sub optReport_Click(Index As Integer)

chkLoans.Visible = False
chkLoans.VALUE = 0
cmbDeductions.Enabled = False
OptScope(0).Enabled = True
OptScope(2).Enabled = True
OptScope(3).Enabled = True
'Select report type and header.
Select Case Index
    Case 0  ' Payroll Summary
        ReportType = 0
        ReportHeader = "Payroll Summary Report"
        OptScope(0).Enabled = False
        OptScope(2).Enabled = False
        OptScope(3).Enabled = False
        OptScope(1).VALUE = True
    Case 1  ' SSS
        ReportType = 1
        ReportHeader = "SSS Contributions Report"
        chkLoans.Visible = True
    Case 2  ' Tax
        ReportType = 2
        ReportHeader = "Income Tax Return Report"
    Case 3  ' HDMF
        ReportType = 3
        ReportHeader = "HDMF Contributions Report"
        chkLoans.Visible = True
    Case 4  ' PhilHealth
        ReportType = 4
        ReportHeader = "PhilHealth Contributions Report"
    Case 5  ' 13th Month
        ReportType = 5
        ReportHeader = "13th Month Pay Report"
    Case 6  ' Cash Bond
        ReportType = 6
        ReportHeader = "Cash Bond Summary Report"
    Case 7  ' HR HeadCount Report
        ReportType = 7
        ReportHeader = "HR Head Count Report"
        OptScope(0).Enabled = False
        OptScope(2).Enabled = False
        OptScope(3).Enabled = False
        OptScope(1).VALUE = True
    Case 8  ' Other Deductions
        ReportType = 8
        ReportHeader = "Payroll Deductions Report"
        cmbDeductions.Enabled = True
End Select

End Sub

Private Sub OptScope_Click(Index As Integer)
 
Select Case Index
    
    Case 0  ' Per Cut-off
        cmbQuarter.Enabled = False
        cmbMonthof.Enabled = True
        ReportScope = 0
        ScopeRep = "Per Cut-off for the Month of "
    Case 1  ' Monthly
        cmbQuarter.Enabled = False
        cmbMonthof.Enabled = True
        ReportScope = 1
        ScopeRep = "For the Month of "
    Case 2  ' Quarterly
        cmbQuarter.Enabled = True
        cmbMonthof.Enabled = False
        ReportScope = 2
        ScopeRep = "For the " & cmbQuarter.Text & " of " & cmbYearof.Text
    Case 3  ' Annually
        cmbQuarter.Enabled = False
        cmbMonthof.Enabled = False
        ReportScope = 3
        ScopeRep = "For the Year " & cmbYearof.Text
End Select

End Sub

Private Sub GetReportScope()

Select Case ReportScope
    Case 0  ' Per Cut-off
        whereCon = " DATENAME(month, payrollmaster.startpaydate) = '" & cmbMonthof.Text & "' and year(payrollmaster.startpaydate) = " & cmbYearof.Text
    Case 1  ' Monthly
        whereCon = " DATENAME(month, payrollmaster.startpaydate) = '" & cmbMonthof.Text & "' and year(payrollmaster.startpaydate) = " & cmbYearof.Text
    Case 2  ' Quarterly
        If cmbQuarter.Text = "1st Quarter" Then     ' 1st Quarter
            whereCon = " MONTH(payrollmaster.startpaydate) between 1 and 3 and year(payrollmaster.startpaydate) = " & cmbYearof.Text
        ElseIf cmbQuarter.Text = "2nd Quarter" Then ' 2nd Quarter
            whereCon = " MONTH(payrollmaster.startpaydate) between 4 and 6 and year(payrollmaster.startpaydate) = " & cmbYearof.Text
        ElseIf cmbQuarter.Text = "3rd Quarter" Then ' 3rd Quarter
            whereCon = " MONTH(payrollmaster.startpaydate) between 7 and 9 and year(payrollmaster.startpaydate) = " & cmbYearof.Text
        ElseIf cmbQuarter.Text = "4th Quarter" Then ' 4th Quarter
            whereCon = " MONTH(payrollmaster.startpaydate) between 10 and 12 and year(payrollmaster.startpaydate) = " & cmbYearof.Text
        End If
    Case 3  ' Annually
        whereCon = " year(payrollmaster.startpaydate) = " & cmbYearof.Text
End Select
If cmbClient.Text <> "--ALL CLIENTS--" Then ' for selected clients only
    whereCon = whereCon & " and PayrollMaster.clientcode = " & mClientCode
End If

End Sub

'Private Sub GeneratePhilHealthReports(ByRef rs As ADODB.Recordset)
'Dim xlApp1 As Excel.Application
'Dim xlBook1 As Excel.Workbook
'Dim xlBook2 As Excel.Workbook
'Dim xlSheet As Excel.Worksheet
'Dim ctr As Long
'Dim i As Integer, j As Integer, k As Integer, l As Integer
'Dim cPos As Integer
'Dim prevClientCode As Integer
'Dim prevControlNo As String
'Dim prevCutOff As Variant
'Dim StartOfLine As Integer
'Dim EndOfLine As Integer
'Dim ExcelTemplate As String
'Dim mRange As String
'Dim mRange1 As String
'Dim mRange2 As String
'Dim mRange3 As String
'Dim HeaderFlag As Boolean
'Dim MemberCtr As Integer
'
'ExcelTemplate = App.Path & "\Reports\RF-1 Template.xls"
'
'Set xlApp1 = New Excel.Application
''Set xlBook1 = xlApp1.Workbooks.Open(ExcelTemplate)
'Set xlBook1 = xlApp1.Workbooks.Add(ExcelTemplate)
'
'xlApp1.WindowState = xlMaximized
'xlApp1.Visible = True
'xlApp1.DisplayAlerts = False
'
''Set xlSheet = book1.Worksheets.Add
'Set xlSheet = xlBook1.Worksheets(1)
''Set xlSheet.Name = "PayrollBankCopy"
'
'cPos = 5
'ctr = 7
'TotalLoan = 0: LoanStart = 0
'HeaderFlag = True
''On Error GoTo Errh
'' Print Header
''xlSheet.Range("E10") = mSapi  '" (SAPI)"
'
'
''xlSheet.Range("A2") = ReportHeader & " " & IIf(chkLoans.VALUE = 1, "(for Loans)", "") 'mSapi
'
''xlSheet.Range("A3") = ScopeRep & " " & cmbMonthof & " " & cmbYearof
'
'ctr = ctr + 1
'
'i = 1
'j = 5 '18
'MemberCtr = 1
'With xlSheet
'Do While ListView1(0).ListItems.Count >= i
'
'    If HeaderFlag = True Then
'        l = 1
'        For k = Asc("C") To Asc("P") ' Philhealth no.
'            If (Chr(k) <> "E" And Chr(k) <> "O") Then
'                mRange = Chr(k) & j
'                If Mid(frmPHealthReport.txtPhilHealthNo, l, 1) = "-" Then
'                    l = l + 1
'                End If
'                .Range(mRange) = Mid(frmPHealthReport.txtPhilHealthNo, l, 1)
'                l = l + 1
'            End If
'        Next
'        j = j + 2
'        l = 1
'        For k = Asc("C") To Asc("Q") ' TIN no.
'            If (Chr(k) <> "F" And Chr(k) <> "J" And Chr(k) <> "N") Then
'                mRange = Chr(k) & j
'                If Mid(frmPHealthReport.txtEmpTIN, l, 1) = "-" Then
'                    l = l + 1
'                End If
'                .Range(mRange) = Mid(frmPHealthReport.txtEmpTIN, l, 1)
'                l = l + 1
'            End If
'        Next
'        j = j + 3
'        mRange = "E" & j
'        xlSheet.Range(mRange) = frmPHealthReport.txtEmployerName
'        j = j + 2
'        mRange = "E" & j
'        xlSheet.Range(mRange) = frmPHealthReport.txtMailingAddress
'        j = j + 2
'        mRange = "E" & j
'        xlSheet.Range(mRange) = frmPHealthReport.txtTelephoneNo
'        j = j + 4
'        HeaderFlag = False
'        MemberCtr = 0
'    End If
'
'    MemberCtr = MemberCtr + 1
'
'    mRange1 = "B" & j
'    mRange2 = "E" & j
'    mRange3 = "M" & j
'
'    .Range(mRange1) = ListView1(0).ListItems(i).SubItems(2)
'    .Range(mRange2) = ListView1(0).ListItems(i).SubItems(3)
'    .Range(mRange3) = ListView1(0).ListItems(i).SubItems(4)
'
'    l = 1
'    For k = Asc("U") To Asc("Z")
'        If Chr(k) <> "W" Then
'            mRange = Chr(k) & j
'            If Mid(ListView1(0).ListItems(i).SubItems(5), l, 1) = "-" Then
'                l = l + 1
'            End If
'            .Range(mRange) = Mid(ListView1(0).ListItems(i).SubItems(5), l, 1)
'            l = l + 1
'        End If
'    Next
'    For k = Asc("A") To Asc("H")
'        If Chr(k) <> "G" Then
'            mRange = "A" & Chr(k) & j
'            If Mid(ListView1(0).ListItems(i).SubItems(5), l, 1) = "-" Then
'                l = l + 1
'            End If
'            .Range(mRange) = Mid(ListView1(0).ListItems(i).SubItems(5), l, 1)
'            l = l + 1
'        End If
'    Next
'
'    mRange1 = "AJ" & j
'    .Range(mRange1) = ListView1(0).ListItems(i).SubItems(6)
'    mRange1 = "AM" & j
'    .Range(mRange1) = ListView1(0).ListItems(i).SubItems(7)
'    mRange1 = "AN" & j
'    .Range(mRange1) = ListView1(0).ListItems(i).SubItems(8)
'    j = j + 1
'
'    If MemberCtr = 15 Then
'        HeaderFlag = True
'        j = j + 12
'    End If
''
''        litem1.SubItems(1) = IIf(IsNull(mRs!Clientname), "Blank", mRs!Clientname)
''        litem1.SubItems(2) = IIf(IsNull(mRs!Surname), "Blank", mRs!Surname)
''        litem1.SubItems(3) = IIf(IsNull(mRs!Firstname), "Blank", mRs!Firstname)
''        litem1.SubItems(4) = IIf(IsNull(mRs!Middlename), "Blank", mRs!Middlename)
''        litem1.SubItems(5) = IIf(IsNull(mRs!PhilHealth), "No Number", mRs!PhilHealth)
''        litem1.SubItems(6) = Format(mRs!Basicpay, "0.##")
''        litem1.SubItems(7) = Format(mRs!PHealthEmployee, "0.##")
''        litem1.SubItems(8) = Format(mRs!PHealthEmployer, "0.##")
''
'
'    i = i + 1
'Loop
'   End With
'
'
'Do While Not ListView1(0).ListItems.Count <> 0
'    If prevClientCode <> rs!ClientCode Then
'        ' Print Again the Client Name.
'        mRange = "A" & ctr + 1
'        Call ComputeTotals(xlSheet, xlApp1, ctr, i, StartOfLine)
'        ctr = ctr + 2
'        i = i + 2
'        mRange = "A" & ctr + 1
'        xlSheet.Range(mRange) = rs!ClientName
'        With xlSheet.Range(mRange)
'            .Font.Name = "Tahoma"   'Font.Name
'            .Font.Size = 10         'Font.size
'            .Font.Bold = True       'Font.Bold
'            .Font.Italic = False    'Font.Italic
'        End With
'        prevClientCode = rs!ClientCode
'        StartOfLine = ctr + 2
'        ctr = ctr + 2
'    End If
'    'Check if there is a change in Control Number.
'    If prevControlNo <> rs!EmpNo And ReportType = 1 And chkLoans.VALUE = 1 And ReportScope = 3 Then  ' Quarterly report Then
'        ' Print sub-total of loan.
'        Call ComputeTotals(xlSheet, xlApp1, ctr, i, StartOfLine, True)
'        prevControlNo = rs!EmpNo
'        ctr = ctr + 1
'    End If
'    xlSheet.Cells(ctr, 1).VALUE = IIf(rs!EmpNo = "" Or IsNull(rs!EmpNo), "00-0000", rs!EmpNo)
'    xlSheet.Cells(ctr, 2).VALUE = rs!Surname
'    xlSheet.Cells(ctr, 3).VALUE = rs!Firstname
'    xlSheet.Cells(ctr, 4).VALUE = rs!Middlename
'
'    Select Case ReportType
'        Case 1  ' SSS
'            If chkLoans.VALUE = 1 Then ' SSS Loan
'                xlSheet.Cells(ctr, 5).VALUE = rs!sss
'                xlSheet.Cells(ctr, 6).VALUE = Format(rs!SSSLoan, "##,#00.00")
'                TotalLoan = TotalLoan + rs!SSSLoan
'                mRange = "F" & ctr & ":F" & ctr
'                xlSheet.Range(mRange).Select
'                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'                If ReportScope = 0 Then
'                    xlSheet.Cells(ctr, 7).VALUE = rs!StartPayDate
'                    xlSheet.Cells(ctr, 8).VALUE = rs!EndPayDate
''                ElseIf ReportScope = 3 Then ' for annual
''                    xlSheet.Cells(ctr, 7).VALUE = "Annual"
''                    xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
'                Else
'                    xlSheet.Cells(ctr, 7).VALUE = rs!Monthof
'                    xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
'                End If
'                mRange = "A" & StartOfLine & ":H" & i + 7
'                mRange2 = "B7:H" & i + 7
'            Else
'                xlSheet.Cells(ctr, 5).VALUE = rs!sss
'                xlSheet.Cells(ctr, 6).VALUE = Format(rs!SSSEmployee, "##,#00.00")
'                xlSheet.Cells(ctr, 7).VALUE = Format(rs!sssemployer, "##,#00.00")
'                xlSheet.Cells(ctr, 8).VALUE = Format(rs!SSSEmployee + rs!sssemployer, "##,#00.00")
'                mRange = "F" & ctr & ":H" & ctr
'                xlSheet.Range(mRange).Select
'                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'                If ReportScope = 0 Then
'                    xlSheet.Cells(ctr, 9).VALUE = rs!StartPayDate
'                    xlSheet.Cells(ctr, 10).VALUE = rs!EndPayDate
''                ElseIf ReportScope = 3 Then ' for annual
''                    xlSheet.Cells(ctr, 9).VALUE = "Annual"
''                    xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
'                Else
'                    xlSheet.Cells(ctr, 9).VALUE = rs!Monthof
'                    xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
'                End If
'                mRange = "A" & StartOfLine & ":J" & i + 7
'                mRange2 = "B7:J" & i + 7
'            End If
'        Case 2  ' Tax
'            xlSheet.Cells(ctr, 5).VALUE = rs!TIN
'            xlSheet.Cells(ctr, 6).VALUE = Format(rs!WTaxAmount, "##,#00.00")
'            mRange = "F" & ctr
'            xlSheet.Range(mRange).Select
'            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'            If ReportScope = 0 Then
'                xlSheet.Cells(ctr, 7).VALUE = rs!StartPayDate
'                xlSheet.Cells(ctr, 8).VALUE = rs!EndPayDate
''            ElseIf ReportScope = 3 Then ' for annual
''                xlSheet.Cells(ctr, 7).VALUE = "Annual"
''                xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
'            Else
'                xlSheet.Cells(ctr, 7).VALUE = rs!Monthof
'                xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
'            End If
'            mRange = "A" & StartOfLine & ":H" & i + 7
'            mRange2 = "B7:H" & i + 7
'        Case 3  ' HDMF
'            If chkLoans.VALUE = 1 Then ' SSS Loan
'                xlSheet.Cells(ctr, 5).VALUE = IIf(IsNull(rs!pagibig), Format(rs!BIRTHDATE, "mm/dd/yyyy"), rs!pagibig)
'                xlSheet.Cells(ctr, 6).VALUE = Format(rs!PagibigLoan, "##,#00.00")
'                mRange = "F" & ctr
'                xlSheet.Range(mRange).Select
'                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'                If ReportScope = 0 Then
'                    xlSheet.Cells(ctr, 7).VALUE = rs!StartPayDate
'                    xlSheet.Cells(ctr, 8).VALUE = rs!EndPayDate
''                ElseIf ReportScope = 3 Then ' for annual
''                    xlSheet.Cells(ctr, 7).VALUE = "Annual"
''                    xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
'                Else
'                    xlSheet.Cells(ctr, 7).VALUE = rs!Monthof
'                    xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
'                End If
'                mRange = "A" & StartOfLine & ":H" & i + 7
'                mRange2 = "B7:H" & i + 7
'            Else
'                xlSheet.Cells(ctr, 5).VALUE = IIf(IsNull(rs!pagibig), Format(rs!BIRTHDATE, "mm/dd/yyyy"), rs!pagibig)
'                xlSheet.Cells(ctr, 6).VALUE = Format(rs!PagibigEmployee, "##,#00.00")
'                xlSheet.Cells(ctr, 7).VALUE = Format(rs!PagibigEmployer, "##,#00.00")
'                xlSheet.Cells(ctr, 8).VALUE = Format(rs!PagibigEmployee + rs!PagibigEmployer, "##,#00.00")
'                mRange = "F" & ctr & ":H" & ctr
'                xlSheet.Range(mRange).Select
'                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'                If ReportScope = 0 Then
'                    xlSheet.Cells(ctr, 9).VALUE = rs!StartPayDate
'                    xlSheet.Cells(ctr, 10).VALUE = rs!EndPayDate
''                ElseIf ReportScope = 3 Then ' for annual
''                    xlSheet.Cells(ctr, 9).VALUE = "Annual"
''                    xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
'                Else
'                    xlSheet.Cells(ctr, 9).VALUE = rs!Monthof
'                    xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
'                End If
'                mRange = "A" & StartOfLine & ":J" & i + 7
'                mRange2 = "B7:J" & i + 7
'            End If
'        Case 4  ' PhilHealth
'            xlSheet.Cells(ctr, 5).VALUE = rs!PhilHealth
'            xlSheet.Cells(ctr, 6).VALUE = Format(rs!PHealthEmployee, "##,#00.00")
'            xlSheet.Cells(ctr, 7).VALUE = Format(rs!PHealthEmployer, "##,#00.00")
'            xlSheet.Cells(ctr, 8).VALUE = Format(rs!PHealthEmployee + rs!PHealthEmployer, "##,#00.00")
'            mRange = "F" & ctr & ":H" & ctr
'            xlSheet.Range(mRange).Select
'            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'            If ReportScope = 0 Then
'                xlSheet.Cells(ctr, 9).VALUE = rs!StartPayDate
'                xlSheet.Cells(ctr, 10).VALUE = rs!EndPayDate
''            ElseIf ReportScope = 3 Then ' for annual
''                xlSheet.Cells(ctr, 9).VALUE = "Annual"
''                xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
'            Else
'                xlSheet.Cells(ctr, 9).VALUE = rs!Monthof
'                xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
'            End If
'            mRange = "A" & StartOfLine & ":J" & i + 7
'            mRange2 = "B7:J" & i + 7
'        Case 5  ' 13th Month
'            xlSheet.Cells(ctr, 5).VALUE = Format(rs!Month13Payment, "##,#00.00")
'            mRange = "E" & ctr
'            xlSheet.Range(mRange).Select
'            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'            If ReportScope = 0 Then
'                xlSheet.Cells(ctr, 6).VALUE = rs!StartPayDate
'                xlSheet.Cells(ctr, 7).VALUE = rs!EndPayDate
''            ElseIf ReportScope = 3 Then ' for annual
''                xlSheet.Cells(ctr, 6).VALUE = "Annual"
''                xlSheet.Cells(ctr, 7).VALUE = cmbYearof.Text
'            Else
'                xlSheet.Cells(ctr, 6).VALUE = rs!Monthof
'                xlSheet.Cells(ctr, 7).VALUE = cmbYearof.Text
'            End If
'            mRange = "A" & StartOfLine & ":G" & i + 7
'            mRange2 = "B7:G" & i + 7
'        Case 6  ' Cash Bond
'            xlSheet.Cells(ctr, 5).VALUE = Format(rs!TASAesla, "##,#00.00")
'            mRange = "E" & ctr
'            xlSheet.Range(mRange).Select
'            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'            If ReportScope = 0 Then
'                xlSheet.Cells(ctr, 6).VALUE = rs!StartPayDate
'                xlSheet.Cells(ctr, 7).VALUE = rs!EndPayDate
'            Else
'                xlSheet.Cells(ctr, 6).VALUE = rs!Monthof
'                xlSheet.Cells(ctr, 7).VALUE = cmbYearof.Text
'            End If
'            mRange = "A" & StartOfLine & ":G" & i + 7
'            mRange2 = "B7:G" & i + 7
'        Case 7  ' Head Count
'        Case 8  ' Other Deductions
'            xlSheet.Cells(ctr, 5).VALUE = Format(rs!Medical, "##,#00.00")
'            xlSheet.Cells(ctr, 6).VALUE = Format(rs!SalaryLoan, "##,#00.00")
'            xlSheet.Cells(ctr, 7).VALUE = Format(rs!Uniform, "##,#00.00")
'            xlSheet.Cells(ctr, 8).VALUE = Format(rs!Variances, "##,#00.00")
'            xlSheet.Cells(ctr, 9).VALUE = Format(rs!Vale_IOU, "##,#00.00")
'
'            xlSheet.Cells(ctr, 10).VALUE = Format(rs!LMCDed, "##,#00.00")
'            xlSheet.Cells(ctr, 11).VALUE = Format(rs!ValeWithClient, "##,#00.00")
'            xlSheet.Cells(ctr, 12).VALUE = Format(rs!CashAdvance, "##,#00.00")
'            xlSheet.Cells(ctr, 13).VALUE = Format(rs!MiscFee, "##,#00.00")
'
'            mRange = "E" & ctr & ":M" & ctr
'            xlSheet.Range(mRange).Select
'            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'            If ReportScope = 0 Then
'                xlSheet.Cells(ctr, 14).VALUE = rs!StartPayDate
'                xlSheet.Cells(ctr, 15).VALUE = rs!EndPayDate
'            Else
'                xlSheet.Cells(ctr, 14).VALUE = rs!Monthof
'                xlSheet.Cells(ctr, 15).VALUE = cmbYearof.Text
'            End If
'            mRange = "A" & StartOfLine & ":O" & i + 7   'L
'            mRange2 = "B7:O" & i + 7
'    End Select
'    'mGrandPay = mGrandPay + (mTotalPay - mTotalDeduction)
'    i = i + 1
'    ctr = ctr + 1
'    rs.MoveNext
'Loop
'
'End Sub

Private Sub GenerateReports(ByRef rs As ADODB.Recordset)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mRange As String, sumString As String, mRange2 As String, mRange3 As String
Dim ctr As Long
Dim i As Integer, j As Integer
Dim cPos As Integer
Dim prevClientCode As Integer
Dim prevControlNo As String
Dim prevCutOff As Variant
Dim StartOfLine As Integer
Dim EndOfLine As Integer

If ReportType = 0 Or ReportType = 0 Then
    MsgBox "Not yet finished.", vbInformation
    Exit Sub
End If


Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

'Set xlSheet = book1.Worksheets.Add
Set xlSheet = xlBook1.Worksheets(1)
'Set xlSheet.Name = "PayrollBankCopy"

cPos = 5
ctr = 7
i = 1
TotalLoan = 0: LoanStart = 0

ESumString = ""
FSumString = ""
GSumString = ""
HSumString = ""
ISumString = ""
JSumString = ""
KSumString = ""
LSumString = ""
MSumString = ""
NSumString = ""

On Error GoTo Errh
' Print Header
xlSheet.Range("A1") = mSapi  '" (SAPI)"
With xlSheet.Range("A1")
    .Font.Name = "Tahoma"  'Font.Name
    .Font.Size = 12 'font.size
    .Font.Bold = True ' Font.Bold
    .Font.Italic = False 'font.Italic
End With
xlSheet.Range("A2") = ReportHeader & " " & IIf(chkLoans.VALUE = 1, "(for Loans)", "") 'mSapi
With xlSheet.Range("A2")
    .Font.Name = "Tahoma"  'Font.Name
    .Font.Size = 10 'font.size
    .Font.Bold = False ' Font.Bold
    .Font.Italic = False 'font.Italic
End With
If ReportScope = 2 Or ReportScope = 3 Then ' Quarterly report
    xlSheet.Range("A3") = ScopeRep
Else
    xlSheet.Range("A3") = ScopeRep & " " & cmbMonthof & " " & cmbYearof
End If

'Print the Header.
xlSheet.Cells(ctr, 1).VALUE = "Control #"
xlSheet.Cells(ctr, 2).VALUE = "Surname   "
xlSheet.Cells(ctr, 3).VALUE = "Firstname "
xlSheet.Cells(ctr, 4).VALUE = "Middlename"

Select Case ReportType
    Case 1  ' SSS
        If chkLoans.VALUE = 1 Then ' for SSS Loans
            xlSheet.Cells(ctr - 1, 6).VALUE = "S S S"
            xlSheet.Cells(ctr, 5).VALUE = "SSS Number"
            xlSheet.Cells(ctr, 6).VALUE = "Loan Amount"
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 7).VALUE = "Start Pay"
                xlSheet.Cells(ctr, 8).VALUE = "End Pay"
'            ElseIf ReportScope = 3 Then ' for annual
'                xlSheet.Cells(ctr, 7).VALUE = "Annual"
'                xlSheet.Cells(ctr, 8).VALUE = "Year"
            Else
                xlSheet.Cells(ctr, 7).VALUE = "Month"
                xlSheet.Cells(ctr, 8).VALUE = "Year"
            End If
            mRange = "A1:H8"
            mRange2 = "A6:H7"
            mRange3 = "F6:F6"
        Else
            xlSheet.Cells(ctr, 5).VALUE = "SSS Number  "
            xlSheet.Cells(ctr - 1, 6).VALUE = "S     S     S"
            xlSheet.Cells(ctr, 6).VALUE = "Employee"
            'xlSheet.Cells(ctr - 1, 7).VALUE = "S S S"
            xlSheet.Cells(ctr, 7).VALUE = "Employer"
            'xlSheet.Cells(ctr - 1, 8).VALUE = "S S S"
            xlSheet.Cells(ctr, 8).VALUE = "Total    "
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 9).VALUE = "Start Pay"
                xlSheet.Cells(ctr, 10).VALUE = "End Pay"
'            ElseIf ReportScope = 3 Then ' for annual
'                xlSheet.Cells(ctr, 9).VALUE = "Annual"
'                xlSheet.Cells(ctr, 10).VALUE = "Year"
            Else
                xlSheet.Cells(ctr, 9).VALUE = "Month"
                xlSheet.Cells(ctr, 10).VALUE = "Year"
            End If
            mRange = "A1:J10"
            mRange2 = "A6:J7"
            mRange3 = "F6:H6"
        End If
    Case 2  ' Tax
        xlSheet.Cells(ctr, 5).VALUE = "Tax ID No."
        xlSheet.Cells(ctr, 6).VALUE = "WTax Amount"
        If ReportScope = 0 Then
            xlSheet.Cells(ctr, 7).VALUE = "Start Pay"
            xlSheet.Cells(ctr, 8).VALUE = "End Pay"
'        ElseIf ReportScope = 3 Then ' for annual
'            xlSheet.Cells(ctr, 7).VALUE = "Annual"
'            xlSheet.Cells(ctr, 8).VALUE = "Year"
        Else
            xlSheet.Cells(ctr, 7).VALUE = "Month"
            xlSheet.Cells(ctr, 8).VALUE = "Year"
        End If
        mRange = "A1:H8"
        mRange2 = "A6:H7"
        mRange3 = "F6:F6"
    Case 3  ' HDMF
        If chkLoans.VALUE = 1 Then ' for Pagibig Loans
            xlSheet.Cells(ctr, 5).VALUE = "Pagibig No."
            xlSheet.Cells(ctr - 1, 6).VALUE = "PAGIBIG"
            xlSheet.Cells(ctr, 6).VALUE = "Loan Amount"
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 7).VALUE = "Start Pay"
                xlSheet.Cells(ctr, 8).VALUE = "End Pay"
'            ElseIf ReportScope = 3 Then ' for annual
'                xlSheet.Cells(ctr, 7).VALUE = "Annual"
'                xlSheet.Cells(ctr, 8).VALUE = "Year"
            Else
                xlSheet.Cells(ctr, 7).VALUE = "Month"
                xlSheet.Cells(ctr, 8).VALUE = "Year"
            End If
            mRange = "A1:H8"
            mRange2 = "A6:H7"
            mRange3 = "F6:F6"
        Else
            xlSheet.Cells(ctr, 5).VALUE = "Pagibig No."
            xlSheet.Cells(ctr - 1, 6).VALUE = "P  A  G  I  B  I  G"
            xlSheet.Cells(ctr, 6).VALUE = "Employee"
            'xlSheet.Cells(ctr - 1, 7).VALUE = "PAGIBIG"
            xlSheet.Cells(ctr, 7).VALUE = "Employer"
            'xlSheet.Cells(ctr - 1, 8).VALUE = "PAGIBIG"
            xlSheet.Cells(ctr, 8).VALUE = "Total    "
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 9).VALUE = "Start Pay"
                xlSheet.Cells(ctr, 10).VALUE = "End Pay"
'            ElseIf ReportScope = 3 Then ' for annual
'                xlSheet.Cells(ctr, 9).VALUE = "Annual"
'                xlSheet.Cells(ctr, 10).VALUE = "Year"
            Else
                xlSheet.Cells(ctr, 9).VALUE = "Month"
                xlSheet.Cells(ctr, 10).VALUE = "Year"
            End If
            mRange = "A1:J10"
            mRange2 = "A6:J7"
            mRange3 = "F6:H6"
        End If
    Case 4  ' PhilHealth
        xlSheet.Cells(ctr, 5).VALUE = "PhilHealth #"
        xlSheet.Cells(ctr - 1, 6).VALUE = "PhilHealth"
        xlSheet.Cells(ctr, 6).VALUE = "Employee"
        'xlSheet.Cells(ctr - 1, 7).VALUE = "PhilHealth"
        xlSheet.Cells(ctr, 7).VALUE = "Employer"
        'xlSheet.Cells(ctr - 1, 8).VALUE = "PhilHealth"
        xlSheet.Cells(ctr, 8).VALUE = "Total    "
        If ReportScope = 0 Then
            xlSheet.Cells(ctr, 9).VALUE = "Start Pay"
            xlSheet.Cells(ctr, 10).VALUE = "End Pay"
'        ElseIf ReportScope = 3 Then ' for annual
'            xlSheet.Cells(ctr, 9).VALUE = "Annual"
'            xlSheet.Cells(ctr, 10).VALUE = "Year"
        Else
            xlSheet.Cells(ctr, 9).VALUE = "Month"
            xlSheet.Cells(ctr, 10).VALUE = "Year"
        End If
        mRange = "A1:J10"
        mRange2 = "A6:J7"
        mRange3 = "F6:H6"
    Case 5  ' 13th Month
        xlSheet.Cells(ctr - 1, 5).VALUE = "13th Month"
        xlSheet.Cells(ctr, 5).VALUE = "Payment     "
        If ReportScope = 0 Then
            xlSheet.Cells(ctr, 6).VALUE = "Start Pay"
            xlSheet.Cells(ctr, 7).VALUE = "End Pay"
'        ElseIf ReportScope = 3 Then 'for annual
'            xlSheet.Cells(ctr, 6).VALUE = "Annual"
'            xlSheet.Cells(ctr, 7).VALUE = "Year"
        Else
            xlSheet.Cells(ctr, 6).VALUE = "Month"
            xlSheet.Cells(ctr, 7).VALUE = "Year"
        End If
        mRange = "A1:G7"
        mRange2 = "A6:G7"
        mRange3 = "C6:C6"
    Case 6  ' Cash Bond
        xlSheet.Cells(ctr - 1, 5).VALUE = "Cash   "
        xlSheet.Cells(ctr, 5).VALUE = "Bond   "
        If ReportScope = 0 Then
            xlSheet.Cells(ctr, 6).VALUE = "Start Pay"
            xlSheet.Cells(ctr, 7).VALUE = "End Pay"
        Else
            xlSheet.Cells(ctr, 6).VALUE = "Month"
            xlSheet.Cells(ctr, 7).VALUE = "Year"
        End If
        mRange = "A1:G7"
        mRange2 = "A6:G7"
        mRange3 = "C6:C6"
    Case 7  ' Head Count
    Case 8  ' Other Deductions
        xlSheet.Cells(ctr - 1, 5).VALUE = "Calamity"
        xlSheet.Cells(ctr, 5).VALUE = " Loan "
        xlSheet.Cells(ctr - 1, 6).VALUE = "Salary "
        xlSheet.Cells(ctr, 6).VALUE = "Loan   "
        xlSheet.Cells(ctr - 1, 7).VALUE = " Uniform "
        xlSheet.Cells(ctr, 7).VALUE = " Charge "
        xlSheet.Cells(ctr - 1, 8).VALUE = "Other Ded"
        xlSheet.Cells(ctr, 8).VALUE = "Charge   "
        xlSheet.Cells(ctr - 1, 9).VALUE = "Planters"
        xlSheet.Cells(ctr, 9).VALUE = " Loan "
        
        xlSheet.Cells(ctr - 1, 10).VALUE = "  LMC   "
        xlSheet.Cells(ctr, 10).VALUE = "Deduction"
        xlSheet.Cells(ctr - 1, 11).VALUE = " VALE "
        xlSheet.Cells(ctr, 11).VALUE = "Deduction"
        xlSheet.Cells(ctr - 1, 12).VALUE = "  Cash  "
        xlSheet.Cells(ctr, 12).VALUE = " Advance "
        xlSheet.Cells(ctr - 1, 13).VALUE = "Misc.  "
        xlSheet.Cells(ctr, 13).VALUE = "Fee    "
        If ReportScope = 0 Then
            xlSheet.Cells(ctr, 14).VALUE = "Start Pay"
            xlSheet.Cells(ctr, 15).VALUE = "End Pay"
        Else
            xlSheet.Cells(ctr, 14).VALUE = "Month"
            xlSheet.Cells(ctr, 15).VALUE = "Year"
        End If
        mRange = "A1:O15"
        mRange2 = "A6:O7"
        mRange3 = "B6:B6"
End Select
With xlSheet.Range(mRange2)
    .Font.Name = "Tahoma"  'Font.Name
    .Font.Size = 10 'Font.size
    .Font.Bold = True
    .Font.Italic = False
End With
xlSheet.Range(mRange3).Select
With xlApp1.Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .MergeCells = True
End With
'xlSheet.Range(mRange3).Select
'With xlSheet.Range(mRange3)
'    .Font.Name = "Arial Narrow"   'Font.Name
'    .Font.Size = 10         'Font.size
'End With
ctr = ctr + 1

rs.MoveFirst
prevClientCode = rs!ClientCode
prevControlNo = rs!EmpNo
If ReportScope = 0 Then ' per cut-off
    prevCutOff = rs!StartPayDate
'ElseIf ReportScope = 3 Then
'    prevCutOff = "Annual"
Else
    prevCutOff = rs!Monthof
End If
StartOfLine = 8
xlSheet.Range("A5") = rs!ClientName
With xlSheet.Range("A5")
    .Font.Name = "Tahoma"  'Font.Name
    .Font.Size = 10 'Font.size
    .Font.Bold = True ' Font.Bold
    .Font.Italic = False 'font.Italic
End With
mRange = "A" & ctr + 2

Do While Not rs.EOF
    If prevClientCode <> rs!ClientCode Then
        ' Print Again the Client Name.
        mRange = "A" & ctr + 1
        Call ComputeTotals(xlSheet, xlApp1, ctr, i, StartOfLine)
        ctr = ctr + 2
        i = i + 2
        mRange = "A" & ctr + 1
        xlSheet.Range(mRange) = rs!ClientName
        With xlSheet.Range(mRange)
            .Font.Name = "Tahoma"   'Font.Name
            .Font.Size = 10         'Font.size
            .Font.Bold = True       'Font.Bold
            .Font.Italic = False    'Font.Italic
        End With
        prevClientCode = rs!ClientCode
        StartOfLine = ctr + 2
        ctr = ctr + 2
    End If
    'Check if there is a change in Control Number.
    If prevControlNo <> rs!EmpNo And ReportType = 1 And chkLoans.VALUE = 1 And ReportScope = 3 Then  ' Quarterly report Then
        ' Print sub-total of loan.
        Call ComputeTotals(xlSheet, xlApp1, ctr, i, StartOfLine, True)
        prevControlNo = rs!EmpNo
        ctr = ctr + 1
    End If
    xlSheet.Cells(ctr, 1).VALUE = IIf(rs!EmpNo = "" Or IsNull(rs!EmpNo), "00-0000", rs!EmpNo)
    xlSheet.Cells(ctr, 2).VALUE = rs!Surname
    xlSheet.Cells(ctr, 3).VALUE = rs!FirstName
    xlSheet.Cells(ctr, 4).VALUE = rs!MiddleName
    
    Select Case ReportType
        Case 1  ' SSS
            If chkLoans.VALUE = 1 Then ' SSS Loan
                xlSheet.Cells(ctr, 5).VALUE = rs!sss
                xlSheet.Cells(ctr, 6).VALUE = Format(rs!SSSLoan, "##,#00.00")
                TotalLoan = TotalLoan + rs!SSSLoan
                mRange = "F" & ctr & ":F" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                If ReportScope = 0 Then
                    xlSheet.Cells(ctr, 7).VALUE = rs!StartPayDate
                    xlSheet.Cells(ctr, 8).VALUE = rs!EndPayDate
'                ElseIf ReportScope = 3 Then ' for annual
'                    xlSheet.Cells(ctr, 7).VALUE = "Annual"
'                    xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
                Else
                    xlSheet.Cells(ctr, 7).VALUE = rs!Monthof
                    xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
                End If
                mRange = "A" & StartOfLine & ":H" & i + 7
                mRange2 = "B7:H" & i + 7
            Else
                xlSheet.Cells(ctr, 5).VALUE = rs!sss
                xlSheet.Cells(ctr, 6).VALUE = Format(rs!SSSEmployee, "##,#00.00")
                xlSheet.Cells(ctr, 7).VALUE = Format(rs!SSSEmployer, "##,#00.00")
                xlSheet.Cells(ctr, 8).VALUE = Format(rs!SSSEmployee + rs!SSSEmployer, "##,#00.00")
                mRange = "F" & ctr & ":H" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                If ReportScope = 0 Then
                    xlSheet.Cells(ctr, 9).VALUE = rs!StartPayDate
                    xlSheet.Cells(ctr, 10).VALUE = rs!EndPayDate
'                ElseIf ReportScope = 3 Then ' for annual
'                    xlSheet.Cells(ctr, 9).VALUE = "Annual"
'                    xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
                Else
                    xlSheet.Cells(ctr, 9).VALUE = rs!Monthof
                    xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
                End If
                mRange = "A" & StartOfLine & ":J" & i + 7
                mRange2 = "B7:J" & i + 7
            End If
        Case 2  ' Tax
            xlSheet.Cells(ctr, 5).VALUE = rs!TIN
            xlSheet.Cells(ctr, 6).VALUE = Format(rs!WTaxAmount, "##,#00.00")
            mRange = "F" & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 7).VALUE = rs!StartPayDate
                xlSheet.Cells(ctr, 8).VALUE = rs!EndPayDate
'            ElseIf ReportScope = 3 Then ' for annual
'                xlSheet.Cells(ctr, 7).VALUE = "Annual"
'                xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
            Else
                xlSheet.Cells(ctr, 7).VALUE = rs!Monthof
                xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
            End If
            mRange = "A" & StartOfLine & ":H" & i + 7
            mRange2 = "B7:H" & i + 7
        Case 3  ' HDMF
            If chkLoans.VALUE = 1 Then ' SSS Loan
                xlSheet.Cells(ctr, 5).VALUE = IIf(IsNull(rs!pagibig), Format(rs!BIRTHDATE, "mm/dd/yyyy"), rs!pagibig)
                xlSheet.Cells(ctr, 6).VALUE = Format(rs!PagibigLoan, "##,#00.00")
                mRange = "F" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                If ReportScope = 0 Then
                    xlSheet.Cells(ctr, 7).VALUE = rs!StartPayDate
                    xlSheet.Cells(ctr, 8).VALUE = rs!EndPayDate
'                ElseIf ReportScope = 3 Then ' for annual
'                    xlSheet.Cells(ctr, 7).VALUE = "Annual"
'                    xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
                Else
                    xlSheet.Cells(ctr, 7).VALUE = rs!Monthof
                    xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
                End If
                mRange = "A" & StartOfLine & ":H" & i + 7
                mRange2 = "B7:H" & i + 7
            Else
                xlSheet.Cells(ctr, 5).VALUE = IIf(IsNull(rs!pagibig), Format(rs!BIRTHDATE, "mm/dd/yyyy"), rs!pagibig)
                xlSheet.Cells(ctr, 6).VALUE = Format(rs!pagibigemployee, "##,#00.00")
                xlSheet.Cells(ctr, 7).VALUE = Format(rs!PagibigEmployer, "##,#00.00")
                xlSheet.Cells(ctr, 8).VALUE = Format(rs!pagibigemployee + rs!PagibigEmployer, "##,#00.00")
                mRange = "F" & ctr & ":H" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                If ReportScope = 0 Then
                    xlSheet.Cells(ctr, 9).VALUE = rs!StartPayDate
                    xlSheet.Cells(ctr, 10).VALUE = rs!EndPayDate
'                ElseIf ReportScope = 3 Then ' for annual
'                    xlSheet.Cells(ctr, 9).VALUE = "Annual"
'                    xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
                Else
                    xlSheet.Cells(ctr, 9).VALUE = rs!Monthof
                    xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
                End If
                mRange = "A" & StartOfLine & ":J" & i + 7
                mRange2 = "B7:J" & i + 7
            End If
        Case 4  ' PhilHealth
            xlSheet.Cells(ctr, 5).VALUE = rs!PhilHealth
            xlSheet.Cells(ctr, 6).VALUE = Format(rs!PHealthEmployee, "##,#00.00")
            xlSheet.Cells(ctr, 7).VALUE = Format(rs!PHealthEmployer, "##,#00.00")
            xlSheet.Cells(ctr, 8).VALUE = Format(rs!PHealthEmployee + rs!PHealthEmployer, "##,#00.00")
            mRange = "F" & ctr & ":H" & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 9).VALUE = rs!StartPayDate
                xlSheet.Cells(ctr, 10).VALUE = rs!EndPayDate
'            ElseIf ReportScope = 3 Then ' for annual
'                xlSheet.Cells(ctr, 9).VALUE = "Annual"
'                xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
            Else
                xlSheet.Cells(ctr, 9).VALUE = rs!Monthof
                xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
            End If
            mRange = "A" & StartOfLine & ":J" & i + 7
            mRange2 = "B7:J" & i + 7
        Case 5  ' 13th Month
            xlSheet.Cells(ctr, 5).VALUE = Format(rs!Month13Payment, "##,#00.00")
            mRange = "E" & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 6).VALUE = rs!StartPayDate
                xlSheet.Cells(ctr, 7).VALUE = rs!EndPayDate
'            ElseIf ReportScope = 3 Then ' for annual
'                xlSheet.Cells(ctr, 6).VALUE = "Annual"
'                xlSheet.Cells(ctr, 7).VALUE = cmbYearof.Text
            Else
                xlSheet.Cells(ctr, 6).VALUE = rs!Monthof
                xlSheet.Cells(ctr, 7).VALUE = cmbYearof.Text
            End If
            mRange = "A" & StartOfLine & ":G" & i + 7
            mRange2 = "B7:G" & i + 7
        Case 6  ' Cash Bond
            xlSheet.Cells(ctr, 5).VALUE = Format(rs!TASAesla, "##,#00.00")
            mRange = "E" & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 6).VALUE = rs!StartPayDate
                xlSheet.Cells(ctr, 7).VALUE = rs!EndPayDate
            Else
                xlSheet.Cells(ctr, 6).VALUE = rs!Monthof
                xlSheet.Cells(ctr, 7).VALUE = cmbYearof.Text
            End If
            mRange = "A" & StartOfLine & ":G" & i + 7
            mRange2 = "B7:G" & i + 7
        Case 7  ' Head Count
        Case 8  ' Other Deductions
            xlSheet.Cells(ctr, 5).VALUE = Format(rs!Medical, "##,#00.00")
            xlSheet.Cells(ctr, 6).VALUE = Format(rs!SalaryLoan, "##,#00.00")
            xlSheet.Cells(ctr, 7).VALUE = Format(rs!Uniform, "##,#00.00")
            xlSheet.Cells(ctr, 8).VALUE = Format(rs!Variances, "##,#00.00")
            xlSheet.Cells(ctr, 9).VALUE = Format(rs!Vale_IOU, "##,#00.00")
            
            xlSheet.Cells(ctr, 10).VALUE = Format(rs!LMCDed, "##,#00.00")
            xlSheet.Cells(ctr, 11).VALUE = Format(rs!ValeWithClient, "##,#00.00")
            xlSheet.Cells(ctr, 12).VALUE = Format(rs!CashAdvance, "##,#00.00")
            xlSheet.Cells(ctr, 13).VALUE = Format(rs!MiscFee, "##,#00.00")
            
            mRange = "E" & ctr & ":M" & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 14).VALUE = rs!StartPayDate
                xlSheet.Cells(ctr, 15).VALUE = rs!EndPayDate
            Else
                xlSheet.Cells(ctr, 14).VALUE = rs!Monthof
                xlSheet.Cells(ctr, 15).VALUE = cmbYearof.Text
            End If
            mRange = "A" & StartOfLine & ":O" & i + 7   'L
            mRange2 = "B7:O" & i + 7
    End Select
    'mGrandPay = mGrandPay + (mTotalPay - mTotalDeduction)
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext
Loop

' Print the last sub-total.
If ReportType = 1 And chkLoans.VALUE = 1 And ReportScope = 3 Then  ' Quarterly report Then
    Call ComputeTotals(xlSheet, xlApp1, ctr, i, StartOfLine, True)
    ctr = ctr + 1
End If

Call ComputeTotals(xlSheet, xlApp1, ctr, i, StartOfLine)

'ADDED BY RAMIL 03/10/2007
'start
If ReportType = 8 Then 'COMPUTE TOTAL OTHER DEDUCTIONS
    xlSheet.Range("E" & ctr + 5, "E" & ctr + 5).Select
    xlSheet.Application.ActiveCell.VALUE = ESumString & ")"
    xlApp1.Selection.NumberFormat = "#,##0.00"
     
    xlSheet.Range("F" & ctr + 5, "F" & ctr + 5).Select
    xlSheet.Application.ActiveCell.VALUE = FSumString & ")"
    xlApp1.Selection.NumberFormat = "#,##0.00"

    xlSheet.Range("G" & ctr + 5, "G" & ctr + 5).Select
    xlSheet.Application.ActiveCell.VALUE = GSumString & ")"
    xlApp1.Selection.NumberFormat = "#,##0.00"

    xlSheet.Range("H" & ctr + 5, "H" & ctr + 5).Select
    xlSheet.Application.ActiveCell.VALUE = HSumString & ")"
    xlApp1.Selection.NumberFormat = "#,##0.00"

    xlSheet.Range("I" & ctr + 5, "I" & ctr + 5).Select
    xlSheet.Application.ActiveCell.VALUE = ISumString & ")"
    xlApp1.Selection.NumberFormat = "#,##0.00"

    xlSheet.Range("J" & ctr + 5, "J" & ctr + 5).Select
    xlSheet.Application.ActiveCell.VALUE = JSumString & ")"
    xlApp1.Selection.NumberFormat = "#,##0.00"

    xlSheet.Range("K" & ctr + 5, "K" & ctr + 5).Select
    xlSheet.Application.ActiveCell.VALUE = KSumString & ")"
    xlApp1.Selection.NumberFormat = "#,##0.00"
    
    xlSheet.Range("L" & ctr + 5, "L" & ctr + 5).Select
    xlSheet.Application.ActiveCell.VALUE = LSumString & ")"
    xlApp1.Selection.NumberFormat = "#,##0.00"
    
    xlSheet.Range("M" & ctr + 5, "M" & ctr + 5).Select
    xlSheet.Application.ActiveCell.VALUE = MSumString & ")"
    xlApp1.Selection.NumberFormat = "#,##0.00"
    
End If
'end

' Display date and time printed.
xlSheet.Range("A" & ctr + 2).Select
xlSheet.Application.ActiveCell.VALUE = "Date Printed: " & Format(Now, "dddd    mm/dd/yyyy hh:mm:ss ampm")
xlSheet.Application.ActiveCell.Font.Name = "Tahoma"
xlSheet.Application.ActiveCell.Font.Size = 8

        
xlSheet.Range(mRange2).Select
xlApp1.Selection.Columns.AutoFit
xlSheet.Columns("A:A").ColumnWidth = 8.43
xlSheet.Range("A1").Select

errR:

    Set xlApp1 = Nothing
    Set xlBook1 = Nothing
    Set xlSheet = Nothing
    Exit Sub
    
Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
    Resume Next
    
End Sub

Private Sub PrintHeadCountReport(ByRef rs As ADODB.Recordset, ByVal mMonth As String, ByVal mYear As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String, mRange2 As String
Dim sumString As String
Dim ctr As Long
Dim i As Integer
Dim cPos As Integer
Dim LineCtr As Integer
Dim mHead As String

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)

ctr = 7
i = 1

'Updated by Bernard May 13, 2008.
mHead = CompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 12, True)
mHead = "MANPOWER HEADCOUNT UPDATE FOR THE MONTH OF " & mMonth & " " & mYear
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)

Call PrintReportHeader(xlSheet, xlApp1, "A5:A6", "Number", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "B5:B6", "Client Name", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 40)
Call PrintReportHeader(xlSheet, xlApp1, "C5:C6", "HR Head Count", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "D5:D6", "Payroll Head Count", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "E5:E6", "Difference", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 10)
rs.MoveFirst

Do While Not rs.EOF
    ' DoEvents
    xlSheet.Cells(ctr, 1).VALUE = i
    xlSheet.Cells(ctr, 2).VALUE = rs!ClientName
    xlSheet.Cells(ctr, 3).VALUE = rs!HRCount
    xlSheet.Cells(ctr, 4).VALUE = rs!TotalCount
    xlSheet.Cells(ctr, 5).FormulaR1C1 = "=RC[-2]-RC[-1]" '"=RC[-2]+RC[-1]"
    ctr = ctr + 1
    i = i + 1
    rs.MoveNext
Loop
ctr = ctr + 3
xlSheet.Cells(ctr, 2).VALUE = "TOTAL HEADCOUNT"
  
mRange = "C" & 7 & ":C" & ctr
xlSheet.Range(mRange).Select
mRange2 = "C" & ctr
xlSheet.Range(mRange2).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-3]C)"

mRange = "D" & 7 & ":D" & ctr
xlSheet.Range(mRange).Select
mRange2 = "D" & ctr
xlSheet.Range(mRange2).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-3]C)"

mRange = "E" & 7 & ":E" & ctr
xlSheet.Range(mRange).Select
mRange2 = "E" & ctr
xlSheet.Range(mRange2).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-3]C)"

'Box it.
mRange = "A5:E" & ctr + 3
Call PutBox(xlSheet, xlApp1, mRange)

ctr = ctr + 7
xlSheet.Cells(ctr, 2).VALUE = "Prepared by: ______________"
xlSheet.Cells(ctr + 3, 2).VALUE = "Note by: ______________"
xlSheet.Range("A1").Select

End Sub

Private Sub PrintChecklistReport(ByRef rs As ADODB.Recordset, ByVal mMonth As String, ByVal mYear As Integer)
Dim rsDedRecords As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String, mRange2 As String
Dim sumString As String
Dim ctr As Long
Dim i As Integer
Dim cPos As Integer
Dim LineCtr As Integer
Dim mHead As String
Dim mDeduc As Double

' Get all deduction records for specific client.
'strQuery = "SELECT * FROM qryDeductionRecords WHERE ClientCode = " & mClientCode & " order by clientCode"
'Set rsDedRecords = ReturnRst(strQuery, cnnStr, mstatic)
'If rsDedRecords.EOF = True Then
'    MsgBox "No payroll deduction records selected! Cannot proceed.", vbInformation
'    Set rsDedRecords = Nothing
'    GoTo ExitRtn2
'End If

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)

ctr = 8
i = 1
rs.MoveFirst

'Updated by Bernard May 13, 2008.
mHead = CompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "PAYROLL CHECKLIST FOR THE MONTH OF " & UCase(mMonth) & " " & mYear & " EVERY " & rs!CutOffPeriod
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "PAYROLL RELEASE DATE EVERY " & rs!SalaryRelease & " OF " & UCase(mMonth) & " " & mYear
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "COMPANY NAME: " & cmbClient.Text
Call PrintReportHeader(xlSheet, xlApp1, "A5", mHead, "", xlLeft, , False, False, False, , , 10, True)

Call PrintReportHeader(xlSheet, xlApp1, "A6:A7", "No.", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 3.14)
Call PrintReportHeader(xlSheet, xlApp1, "B6:B7", "Name", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 22)
Call PrintReportHeader(xlSheet, xlApp1, "C6:C7", "Daily", "Rate", xlCenter, xlBottom, True, True, True, 15, , 9, False, 4.3)
Call PrintReportHeader(xlSheet, xlApp1, "D6:D7", "Staff", "Status", xlCenter, xlBottom, True, True, True, 15, , 9, False, 6)
Call PrintReportHeader(xlSheet, xlApp1, "E6:H6", "Original DTR", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)
Call PrintReportHeader(xlSheet, xlApp1, "E7:E7", "Bandy Clock", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 4.7)
Call PrintReportHeader(xlSheet, xlApp1, "F7:F7", "Chrono Type", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 5)
Call PrintReportHeader(xlSheet, xlApp1, "G7:G7", "Hand Written", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 4.85)
Call PrintReportHeader(xlSheet, xlApp1, "H7:H7", "Xerox", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 4.56)
Call PrintReportHeader(xlSheet, xlApp1, "I6:Q6", "DEDUCTIONS", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)
Call PrintReportHeader(xlSheet, xlApp1, "I7:I7", "C.B. Balance", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.71)
Call PrintReportHeader(xlSheet, xlApp1, "J7:J7", "P/F Balance", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.71)
Call PrintReportHeader(xlSheet, xlApp1, "K7:K7", "M/F Balance", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.71)
Call PrintReportHeader(xlSheet, xlApp1, "L7:L7", "Trans. Fee", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 4.57)
Call PrintReportHeader(xlSheet, xlApp1, "M7:M7", "VALE C/A", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5)
Call PrintReportHeader(xlSheet, xlApp1, "N7:N7", "SLOAN Amount", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
Call PrintReportHeader(xlSheet, xlApp1, "O7:O7", "Variance w/ Client", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 6.57)
Call PrintReportHeader(xlSheet, xlApp1, "P7:P7", "Medical Ded.", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
Call PrintReportHeader(xlSheet, xlApp1, "Q7:Q7", "Uniform Ded.", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
Call PrintReportHeader(xlSheet, xlApp1, "R6:R7", "Pag-ibig", "Loan", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
Call PrintReportHeader(xlSheet, xlApp1, "S6:S7", "SSS", "Loan", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
'Call PrintReportHeader(xlSheet, xlApp1, "T6:T7", "Daily", "Rate", xlCenter, xlBottom, True, True, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "T6:T7", "Start Contract", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 7.72)
Call PrintReportHeader(xlSheet, xlApp1, "U6:U7", "End Contract", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 7.72)
Call PrintReportHeader(xlSheet, xlApp1, "V6:V7", "Account Number", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 14.14)
Call PrintReportHeader(xlSheet, xlApp1, "W6:W7", "Remarks", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 12)
rs.MoveFirst

Do While Not rs.EOF
    ' DoEvents
    xlSheet.Cells(ctr, 1).VALUE = i
    xlSheet.Cells(ctr, 2).VALUE = rs!EmpName
    mDeduc = 0#
'    If rs!CbAmount <> 0 Then
'        mDeduc = RunningBalance(rsDedRecords, 1, rs!EmployeeID, rs!cbstartpayment, rs!cbendpayment)
'        xlSheet.Cells(ctr, 8).VALUE = mDeduc
'    End If
    xlSheet.Cells(ctr, 3).VALUE = Format(rs!DailyRate, "0.##")
    xlSheet.Cells(ctr, 20).VALUE = Format(rs!StartContract, "mm/dd/yyyy")
    xlSheet.Cells(ctr, 21).VALUE = Format(rs!endcontract, "mm/dd/yyyy")
    xlSheet.Cells(ctr, 22).VALUE = rs!AccountNo
    xlSheet.Cells(ctr, 23).VALUE = rs!Remarks
    ctr = ctr + 1
    i = i + 1
    rs.MoveNext
Loop
mRange = "B" & ctr + 1 & ":B" & ctr + 1
Call PrintReportHeader(xlSheet, xlApp1, mRange, "sub-total", "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 22)
mRange = "B" & ctr + 2 & ":B" & ctr + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total of Staff", "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 22)
mRange = "D" & ctr + 2 & ":D" & ctr + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, i - 1, "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 5)
mRange = "B" & ctr + 3 & ":B" & ctr + 3
Call PrintReportHeader(xlSheet, xlApp1, mRange, "No. of lacking", "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 22)
ctr = ctr + 3

'Box it.
mRange = "A6:W" & ctr + 3
Call PlaceBox(xlSheet, xlApp1, mRange)

ctr = ctr + 4
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
mRange = "B" & ctr
Call CustomizedHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 8, False, 25)

ctr = ctr + 3
xlSheet.Cells(ctr, 2).VALUE = "Prepared by: ______________"
xlSheet.Cells(ctr, 7).VALUE = "Note by: ______________"
xlSheet.Cells(ctr, 11).VALUE = "Received by: ______________"
xlSheet.Range("A1").Select

ExitRtn2:

End Sub

Private Function RunningBalance(ByRef rsDed As ADODB.Recordset, ByVal mFieldNo As Integer, ByVal EmpNo As Integer, ByVal Startpay As Variant, ByVal EndPay As Variant) As Double
Dim mBal As Double
Dim StartDate As Variant
Dim EndDate As Variant

mBal = 0#
With rsDed
    StartDate = Val(Format(Startpay, "yyyymmdd"))
    EndDate = Val(Format(EndPay, "yyyymmdd"))
    .Filter = "EmployeeID" & " = " & EmpNo '& " AND (VAL(format(StartPaydate, ""yyyymmdd"")) between " & StartDate & " and " & EndDate & _
        '" or val(format(EndPaydate, ""yyyymmdd"")) between " & StartDate & " and " & EndDate & ")"
    
'            If StartDate = 19000101 Then StartDate = 99999999
'            If LastDate = 19000101 Then LastDate = 99999999
'            ' Validate whether entry dates fall on cut-off dates.
'            If (StartCutOffDate >= StartDate And StartCutOffDate <= LastDate) Or _
'            (EndCutOffDate >= StartDate And EndCutOffDate <= LastDate) Then
'                IncludeFlag = True
'            Else
'                IncludeFlag = False
'                mRem = "Elpsd."
'            End If
    
    If .EOF Then
        RunningBalance = 0#
    Else
        .MoveFirst
        Do While Not .EOF
            Select Case mFieldNo
                Case 1: ' CB Payments
                    mBal = mBal + !TASAesla
                Case 1: ' CB Payments
                    mBal = mBal + !TASAesla
                Case 1: ' CB Payments
                    mBal = mBal + !TASAesla
                Case 1: ' CB Payments
                    mBal = mBal + !TASAesla
                Case 1: ' CB Payments
                    mBal = mBal + !TASAesla
            End Select
            .MoveNext
        Loop
    End If
    .Filter = ""
End With
RunningBalance = 0#  'mBal
End Function

Private Sub PrintPayrollSummary(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim mRange2 As String
Dim sumString As String
Dim ctr As Long
Dim i As Integer, j As Integer, k As Integer
Dim cPos As Integer
Dim EndSw As Boolean
Dim mHead As String

Dim SubBDO(50) As Double
Dim SubATM(50) As Double
Dim SubCash(50) As Double
Dim SubHold(50) As Double
Dim bdoctr, atmctr, cashctr, holdctr
Dim LineCtr As Integer
Dim bdopos As Integer
Dim atmpos As Integer
Dim cashpos As Integer
Dim holdpos As Integer
Dim mSeriesNo As String
Dim prevClientName As String
Dim SaveTotal As String, ColString As String
Dim firstRun As Boolean
Dim lf, aa
Dim prevStartCutoff As Variant
Dim prevEndCutoff As Variant

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

'Set xlSheet = book1.Worksheets.Add
Set xlSheet = xlBook1.Worksheets(1)
xlSheet.Name = "Payroll Summary"
mSeriesNo = "45"
SaveTotal = ""
firstRun = True

cPos = 5
ctr = 7
LineCtr = 7
For i = 1 To 50
    SubBDO(i) = 0
    SubATM(i) = 0
    SubCash(i) = 0
    SubHold(i) = 0
    bdoctr = 0
    atmctr = 0
    cashctr = 0
    holdctr = 0
Next
i = 1
k = 7

'On Error GoTo Errh
'Print Header
'xlSheet.Range("A2") = mCompanyName & msapi  '" (SAPI)"

'Updated by Bernard May 13, 2008.
mHead = CompanyName & " (" & CompanyShort & ")"
'xlSheet.Range("A2") = mCompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 14, True)
If CutOffPeriod = "START" Then
    mHead = "Payroll Summary Report for the Month of " & Format(rs!StartPayDate, "mmmm") & " " & Format(rs!StartPayDate, "yyyy")
Else
    mHead = "Payroll Summary Report for the Month of " & Format(rs!EntryDate, "mmmm") & " " & Format(rs!EntryDate, "yyyy")
    'mHead = "Payroll Summary Report for the Month of " & Format(rs!PayrollDate, "mmmm") & " " & Format(rs!PayrollDate, "yyyy")
End If

''If Format(rs!StartPayDate, "mmm") = Format(rs!EndPayDate, "mmm") Then
'    'xlSheet.Range("A3") = "Payroll Summary Report for Cut-Off " & Format(rs!startPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "dd, yyyy")
'    mHead = "Payroll Summary Report for Cut-Off " & Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "dd, yyyy")
'Else
'    'xlSheet.Range("A3") = "Payroll Summary Report for Cut-Off " & Format(rs!startPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd, yyyy")
'    mHead = "Payroll Summary Report for Cut-Off " & Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd, yyyy")
'End If
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)

'xlSheet.Range("AO3") = mSeriesNo
'Call PrintReportHeader(xlSheet, xlApp1, "AS3", mSeriesNo, "", xlLeft, , False, False, False, , , 8, False)
'xlSheet.Range("AO4") = Format(Now, "mm/dd/yyyy")
'Call PrintReportHeader(xlSheet, xlApp1, "AS4", Format(Now, "mm/dd/yyyy"), "", xlLeft, xlBottom, False, False, False, , , 8, False)

'xlSheet.Cells(cPos, 1).VALUE = "Control"
Call PrintReportHeader(xlSheet, xlApp1, "A5:A6", "Client Name", "", xlCenter, xlBottom, True, False, True, 15, , 8, False, 30)
Call PrintReportHeader(xlSheet, xlApp1, "B5:B6", "Start Cut-Off", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "C5:C6", "End Cut-Off", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 10)

Call PrintReportHeader(xlSheet, xlApp1, "D5:D6", "Employee ID", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 13)
Call PrintReportHeader(xlSheet, xlApp1, "E5:E6", "Employee Name", "", xlCenter, xlBottom, True, False, True, 15, , 8, False, 27)
Call PrintReportHeader(xlSheet, xlApp1, "F5:F6", "Tax", "Status", xlCenter, xlCenter, True, True, True, 15, , 8, False, 5)
Call PrintReportHeader(xlSheet, xlApp1, "G5:G6", "Daily", "Rate", xlCenter, xlCenter, True, True, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "H5:H6", "Reg.", "Days", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "I5:I6", "This Pay", "Period", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9.5)
Call PrintReportHeader(xlSheet, xlApp1, "J5:J6", "Absent", "Deduction", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)
Call PrintReportHeader(xlSheet, xlApp1, "K5:K6", "Tardiness", "Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)
Call PrintReportHeader(xlSheet, xlApp1, "L5:L6", "Tardiness", "Deduction", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)
Call PrintReportHeader(xlSheet, xlApp1, "M5:M6", "Reg OT", "Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 6)
Call PrintReportHeader(xlSheet, xlApp1, "N5:N6", "Reg OT", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
Call PrintReportHeader(xlSheet, xlApp1, "O5:O6", "Night Diff", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
Call PrintReportHeader(xlSheet, xlApp1, "P5:P6", "Night Diff", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)

Call PrintReportHeader(xlSheet, xlApp1, "Q5:Q6", "Rest Day", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "R5:R6", "Rest Day", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "S5:S6", "Sunday OT", "Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "T5:T6", "Sunday OT", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)

Call PrintReportHeader(xlSheet, xlApp1, "U5:U6", "Spl Hol.", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "V5:V6", "Spl Hol.", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "W5:W6", "Lgl Hol.", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "X5:X6", "Lgl Hol.", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "Y5:Y6", "Salary", "Adjustment", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "Z5:Z6", "Incentive", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "AA5:AA6", "E-cola/Allow.", "Payment", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9.5)
Call PrintReportHeader(xlSheet, xlApp1, "AB5:AB6", "GROSS PAY", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 9)
Call PrintReportHeader(xlSheet, xlApp1, "AC5:AC6", "WTAX", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AD5:AF5", "Employee", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)
Call PrintReportHeader(xlSheet, xlApp1, "AD6:AD6", "SSS", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AE6:AE6", "PhilHealth", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AF6:AF6", "Pag-ibig", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AG5:AG6", "Cash", "Bond", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AH5:AH6", "Placement", "Fee", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AI5:AI6", "Misc.", "Fee", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
Call PrintReportHeader(xlSheet, xlApp1, "AJ5:AJ6", "Card", "Trans Fee", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AK5:AK6", "Salary", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AL5:AL6", "Vale", "(IOU)", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AM5:AM6", "Uniform", "Charges", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AN5:AN6", "Medical", "Charges", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AO5:AO6", "Other", "Ded.", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AP5:AP6", "Adjustment", "Audit", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AQ5:AQ6", "SSS", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AR5:AR6", "PBIG", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AS5:AS6", "Vale", "w/ Client", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
Call PrintReportHeader(xlSheet, xlApp1, "AT5:AT6", "TOTAL", "Deductions", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "AU5:AU6", "NET PAY", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "AV5:AV6", "ChinaBank/BDO", "Account Number", xlCenter, xlCenter, True, True, True, 15, , 8, False, 15.87)

Call PrintReportHeader(xlSheet, xlApp1, "AW5:AW6", "13th Month", "Payment", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "AX5:BA5", "Employer", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 18)
Call PrintReportHeader(xlSheet, xlApp1, "AX6:AX6", "SSS", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
Call PrintReportHeader(xlSheet, xlApp1, "AY6:AY6", "EC", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AZ6:AZ6", "PhilHealth", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
Call PrintReportHeader(xlSheet, xlApp1, "BA6:BA6", "Pag-ibig", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
Call PrintReportHeader(xlSheet, xlApp1, "BB5:BB6", "TOTAL", "Charges", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "BC5:BC6", "Agency", "Charge", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "BD5:BD6", "GROSS", "BILLING", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.7)

Do While Not rs.EOF
    ' DoEvents
    If prevClientName <> rs!ClientName Then
        If i > 1 Then
            Call PrintTotals(xlSheet, xlApp1, ctr, k)
            ' Print grand total in Sheet 2.
            mRange = "H" & ctr & ":BD" & ctr
            xlSheet.Range(mRange).Select
            xlApp1.Selection.Copy
            Call CopyToSheet2(xlSheet, xlApp1, LineCtr, prevClientName & " (" & prevStartCutoff & "-" & prevEndCutoff & ")", False)
            LineCtr = LineCtr + 1
            ctr = ctr + 1
            i = i + 1
            k = ctr
            SaveTotal = SaveTotal & ctr - 1 & ":"
        End If
        xlSheet.Cells(ctr, 1).VALUE = IIf(IsNull(rs!ClientName), "Deleted Client", rs!ClientName)
        'xlSheet.Cells(ctr, 1).Select
        'xlApp1.Selection.Copy
        prevClientName = rs!ClientName
        'Call CopyToSheet2(xlSheet, xlApp1, LineCtr)
    ElseIf prevStartCutoff <> rs!StartPayDate Then
        If i > 2 Then
            Call PrintTotals(xlSheet, xlApp1, ctr, k)
            ' Print grand total in Sheet 2.
            mRange = "H" & ctr & ":BD" & ctr
            xlSheet.Range(mRange).Select
            xlApp1.Selection.Copy
            Call CopyToSheet2(xlSheet, xlApp1, LineCtr, prevClientName & " (" & prevStartCutoff & "-" & prevEndCutoff & ")", False)
            LineCtr = LineCtr + 1
            ctr = ctr + 1
            i = i + 1
            k = ctr
            SaveTotal = SaveTotal & ctr - 1 & ":"
        End If
        xlSheet.Cells(ctr, 1).VALUE = IIf(IsNull(rs!ClientName), "Deleted Client", rs!ClientName)
        'xlSheet.Cells(ctr, 1).Select
        'xlApp1.Selection.Copy
        'prevClientName = rs!ClientName
        'Call CopyToSheet2(xlSheet, xlApp1, LineCtr)
        prevStartCutoff = rs!StartPayDate
        prevEndCutoff = rs!EndPayDate
    Else
        xlSheet.Cells(ctr, 1).VALUE = IIf(IsNull(rs!ClientName), "Deleted Client", rs!ClientName)
    End If
    
    xlSheet.Cells(ctr, 2).VALUE = rs!StartPayDate
    xlSheet.Cells(ctr, 3).VALUE = rs!EndPayDate
    
    xlSheet.Cells(ctr, 4).VALUE = rs!ControlNum
    xlSheet.Cells(ctr, 5).VALUE = rs!EmployeeName
    xlSheet.Cells(ctr, 6).VALUE = rs!TaxCode 'rs!EmpCivilStatus
    xlSheet.Cells(ctr, 7).VALUE = rs!DailyRate
    xlSheet.Cells(ctr, 8).VALUE = rs!DaysPresent
    
    xlSheet.Cells(ctr, 9).VALUE = Format(rs!BasicPay, "##,##0.00")
    xlSheet.Cells(ctr, 10).VALUE = Format(rs!AbsentDeduction, "##,#0.00")
    xlSheet.Cells(ctr, 11).VALUE = Format(rs!HoursTardiness, "##,#0.00")
    xlSheet.Cells(ctr, 12).VALUE = Format(rs!TardinessDeduction, "##,#0.00")
    
    xlSheet.Cells(ctr, 13).VALUE = Format(rs!OTHours, "##,##0.00")
    xlSheet.Cells(ctr, 14).VALUE = Format(rs!OTPay, "##,##0.00")
    xlSheet.Cells(ctr, 15).VALUE = Format(rs!NightDiffOT, "##,##0.00")
    xlSheet.Cells(ctr, 16).VALUE = Format(rs!NightDiffPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 17).VALUE = Format(rs!TaxRefundHours, "##,##0.00")
    xlSheet.Cells(ctr, 18).VALUE = Format(rs!TaxRefundPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 19).VALUE = Format(rs!SundayOT, "##,##0.00")
    xlSheet.Cells(ctr, 20).VALUE = Format(rs!SundayOTPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 21).VALUE = Format(rs!NoRegHoliday, "##,##0.00")
    xlSheet.Cells(ctr, 22).VALUE = Format(rs!RegHolidayPay, "##,##0.00")
    xlSheet.Cells(ctr, 23).VALUE = Format(rs!NoSplHoliday, "##,##0.00")
    xlSheet.Cells(ctr, 24).VALUE = Format(rs!SplHolidayPay, "##,##0.00")
    
'    xlSheet.Cells(ctr, 21).VALUE = Format(rs!NoSickLeave, "##,##0.00")
'    xlSheet.Cells(ctr, 22).VALUE = Format(rs!SickLeavePay, "##,##0.00")
'    xlSheet.Cells(ctr, 23).VALUE = Format(rs!NoVacationLeave, "##,##0.00")
'    xlSheet.Cells(ctr, 24).VALUE = Format(rs!VacationLeavePay, "##,##0.00")

    xlSheet.Cells(ctr, 25).VALUE = Format(rs!OtherIncome, "##,##0.00")
    xlSheet.Cells(ctr, 26).VALUE = Format(rs!Allowance, "##,##0.00")
    xlSheet.Cells(ctr, 27).VALUE = Format(rs!EcolaPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 28).VALUE = Format(rs!GrossPay, "##,##0.00")
    xlSheet.Cells(ctr, 29).VALUE = Format(rs!WTaxAmount, "##,##0.00")
    xlSheet.Cells(ctr, 30).VALUE = Format(rs!SSSEmployee, "##,##0.00")
    xlSheet.Cells(ctr, 31).VALUE = Format(rs!PHealthEmployee, "##,##0.00")
    xlSheet.Cells(ctr, 32).VALUE = Format(rs!pagibigemployee, "##,##0.00")
    
    xlSheet.Cells(ctr, 33).VALUE = Format(rs!TASAesla, "##,##0.00")
    xlSheet.Cells(ctr, 34).VALUE = Format(rs!CashAdvance, "##,##0.00")
    xlSheet.Cells(ctr, 35).VALUE = Format(rs!MiscFee, "##,##0.00")
    xlSheet.Cells(ctr, 36).VALUE = Format(rs!LMCDed, "##,##0.00")
    
    xlSheet.Cells(ctr, 37).VALUE = Format(rs!SalaryLoan, "##,##0.00")
    xlSheet.Cells(ctr, 38).VALUE = Format(rs!Vale_IOU, "##,##0.00")
    xlSheet.Cells(ctr, 39).VALUE = Format(rs!Uniform, "##,##0.00")
    xlSheet.Cells(ctr, 40).VALUE = Format(rs!Medical, "##,##0.00")
    xlSheet.Cells(ctr, 41).VALUE = Format(rs!Variances, "##,##0.00")
    xlSheet.Cells(ctr, 42).VALUE = Format(rs!OverUnderPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 43).VALUE = Format(rs!SSSLoan, "##,##0.00")
    xlSheet.Cells(ctr, 44).VALUE = Format(rs!PagibigLoan, "##,##0.00")
    xlSheet.Cells(ctr, 45).VALUE = Format(rs!ValeWithClient, "##,##0.00")
    xlSheet.Cells(ctr, 46).VALUE = Format(rs!Deductions, "##,##0.00")
    xlSheet.Cells(ctr, 47).VALUE = Format(rs!NetPay, "##,##0.00")
    xlSheet.Cells(ctr, 48).VALUE = IIf(rs!AcctNo = "", "n/a", rs!AcctNo)

    xlSheet.Cells(ctr, 49).VALUE = Format(rs!Month13Payment, "##,##0.00")
    If rs!SSSEmployer <> 0 Then
        xlSheet.Cells(ctr, 50).VALUE = Format(rs!SSSEmployer - IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer), "##,##0.00")
        xlSheet.Cells(ctr, 51).VALUE = Format(IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer), "##,##0.00")
    Else
        xlSheet.Cells(ctr, 50).VALUE = Format(0, "##,##0.00")
        xlSheet.Cells(ctr, 51).VALUE = Format(0, "##,##0.00")
    End If
    xlSheet.Cells(ctr, 52).VALUE = Format(rs!PHealthEmployer, "##,##0.00")
    xlSheet.Cells(ctr, 53).VALUE = Format(rs!PagibigEmployer, "##,##0.00")
    xlSheet.Cells(ctr, 54).VALUE = Format(rs!CalamityLoan, "##,##0.00")
    xlSheet.Cells(ctr, 55).FormulaR1C1 = "=RC[-21]+RC[-5]+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
    xlSheet.Cells(ctr, 56).FormulaR1C1 = "=RC[-2]+RC[-1]"
    
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext

Loop

Call PrintTotals(xlSheet, xlApp1, ctr, k)
mRange = "H" & ctr & ":BD" & ctr
xlSheet.Range(mRange).Select
xlApp1.Selection.Copy
Call CopyToSheet2(xlSheet, xlApp1, LineCtr, prevClientName & " (" & prevStartCutoff & "-" & prevEndCutoff & ")", False)

sumString = ""
If Len(SaveTotal) = 0 Then
    SaveTotal = ctr & ":"
Else
    SaveTotal = SaveTotal & ctr & ":"
End If
ColString = SaveTotal
ctr = ctr + 1
'Compute for Grand Total.
mRange = "E" & ctr & ":E" & ctr
xlSheet.Range(mRange).Select
xlApp1.ActiveCell.FormulaR1C1 = "  GRAND TOTAL"
firstRun = True

Dim mGTotal As Double

mGTotal = 0#
For j = Asc("H") To Asc("Z")
    If firstRun = True Then
        If Asc("H") = j Then
            j = Asc("D")
        Else
            j = Asc("H")
            firstRun = False
        End If
    End If
    SaveTotal = ColString
    While Len(SaveTotal) <> 0
        lf = InStr(SaveTotal, ":")
        aa = LTrim(RTrim(Left(SaveTotal, lf - 1)))
        'If sumString = "" Then ' orig code
            sumString = "=R[" & ((ctr - aa) * -1) & "]C"
            mRange = Chr(j) & aa
            mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
            'xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
        'Else
        '    sumString = sumString & "+R[" & ((ctr - aa) * -1) & "]C"
        'End If
        SaveTotal = Mid(SaveTotal, lf + 1, Len(SaveTotal) - lf)
    Wend
    mRange = Chr(j) & ctr
    xlSheet.Range(mRange).Select
    'xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)" orig code
    xlApp1.ActiveCell.VALUE = mGTotal
    'ActiveCell.FormulaR1C1 = "=R[-5]C+R[-3]C"
    xlSheet.Range(mRange).Select
    If xlApp1.ActiveCell.VALUE = 0 Then
        xlApp1.Selection.EntireColumn.Hidden = True
    End If
    sumString = ""
    mGTotal = 0#
Next
For j = Asc("A") To Asc("U")
    SaveTotal = ColString
    While Len(SaveTotal) <> 0
        lf = InStr(SaveTotal, ":")
        aa = LTrim(RTrim(Left(SaveTotal, lf - 1)))
        'If sumString = "" Then
            sumString = "=R[" & ((ctr - aa) * -1) & "]C"
            mRange = "A" & Chr(j) & aa
            mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
        'Else
        '    sumString = sumString & "+R[" & ((ctr - aa) * -1) & "]C"
        'End If
        SaveTotal = Mid(SaveTotal, lf + 1, Len(SaveTotal) - lf)
    Wend
    mRange = "A" & Chr(j) & ctr
    xlSheet.Range(mRange).Select
    'xlApp1.ActiveCell.FormulaR1C1 = sumString
    xlApp1.ActiveCell.VALUE = mGTotal
    xlSheet.Range(mRange).Select
    If xlApp1.ActiveCell.VALUE = 0 Then
        xlApp1.Selection.EntireColumn.Hidden = True
    End If
    sumString = ""
    mGTotal = 0#
Next
For j = Asc("W") To Asc("Z")
    SaveTotal = ColString
    While Len(SaveTotal) <> 0
        lf = InStr(SaveTotal, ":")
        aa = LTrim(RTrim(Left(SaveTotal, lf - 1)))
        'If sumString = "" Then
            sumString = "=R[" & ((ctr - aa) * -1) & "]C"
            mRange = "A" & Chr(j) & aa
            mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
        'Else
        '    sumString = sumString & "+R[" & ((ctr - aa) * -1) & "]C"
        'End If
        SaveTotal = Mid(SaveTotal, lf + 1, Len(SaveTotal) - lf)
    Wend
    mRange = "A" & Chr(j) & ctr
    xlSheet.Range(mRange).Select
    'xlApp1.ActiveCell.FormulaR1C1 = sumString
    xlApp1.ActiveCell.VALUE = mGTotal
    xlSheet.Range(mRange).Select
    If xlApp1.ActiveCell.VALUE = 0 Then
        xlApp1.Selection.EntireColumn.Hidden = True
    End If
    sumString = ""
    mGTotal = 0#
Next
For j = Asc("A") To Asc("D")
    SaveTotal = ColString
    While Len(SaveTotal) <> 0
        lf = InStr(SaveTotal, ":")
        aa = LTrim(RTrim(Left(SaveTotal, lf - 1)))
        'If sumString = "" Then
            sumString = "=R[" & ((ctr - aa) * -1) & "]C"
            mRange = "B" & Chr(j) & aa
            mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
        'Else
        '    sumString = sumString & "+R[" & ((ctr - aa) * -1) & "]C"
        'End If
        SaveTotal = Mid(SaveTotal, lf + 1, Len(SaveTotal) - lf)
    Wend
    mRange = "B" & Chr(j) & ctr
    xlSheet.Range(mRange).Select
    'xlApp1.ActiveCell.FormulaR1C1 = sumString
    xlApp1.ActiveCell.VALUE = mGTotal
    xlSheet.Range(mRange).Select
    If xlApp1.ActiveCell.VALUE = 0 Then
        xlApp1.Selection.EntireColumn.Hidden = True
    End If
    sumString = ""
    mGTotal = 0#
Next

''Compute for Grand Total.
'EndSw = False
'SumString = "=SUM(R[" & i * -1 & "]C:R[-2]C)"
'mRange = "E" & i + 7 & ":E" & i + 7
'xlsheet.Range(mRange).Select
'xlapp1.ActiveCell.FormulaR1C1 = "           TOTAL"
''Print Total Employee.
'mRange = "D" & i + 7 & ":D" & i + 7
'xlsheet.Range(mRange).Select
'xlapp1.ActiveCell.FormulaR1C1 = i - 1
'
'For j = Asc("H") To Asc("Z")
'    If EndSw = False And Chr(j) = "Z" Then
'        EndSw = True
'        'xlApp1.ActiveCell.FormulaR1C1 = atmctr
'    End If
'    If EndSw = True Then
'        If Chr(j) = "Z" Then
'            mRange = Chr(j) & i + 7 & ":" & Chr(j) & i + 7
'        Else
'            mRange = "A" & Chr(j) & i + 7 & ":A" & Chr(j) & i + 7
'        End If
'    Else
'        mRange = Chr(j) & i + 7 & ":" & Chr(j) & i + 7
'    End If
'    xlsheet.Range(mRange).Select
'    xlapp1.ActiveCell.FormulaR1C1 = SumString '"=SUM(R[i-6]C:R[-2]C)"
'    xlsheet.Range(mRange).Select
'Next
'For j = Asc("A") To Asc("U")    ' K
'    mRange = "A" & Chr(j) & i + 7 & ":A" & Chr(j) & i + 7
'    xlsheet.Range(mRange).Select
'    xlapp1.ActiveCell.FormulaR1C1 = SumString '"=SUM(R[i-6]C:R[-2]C)"
'    xlsheet.Range(mRange).Select
'Next
'For j = Asc("W") To Asc("Z")    ' K
'    mRange = "A" & Chr(j) & i + 7 & ":A" & Chr(j) & i + 7
'    xlsheet.Range(mRange).Select
'    xlapp1.ActiveCell.FormulaR1C1 = SumString '"=SUM(R[i-6]C:R[-2]C)"
'    xlsheet.Range(mRange).Select
'Next
'For j = Asc("A") To Asc("D")    ' K
'    mRange = "B" & Chr(j) & i + 7 & ":B" & Chr(j) & i + 7
'    xlsheet.Range(mRange).Select
'    xlapp1.ActiveCell.FormulaR1C1 = SumString '"=SUM(R[i-6]C:R[-2]C)"
'    xlsheet.Range(mRange).Select
'Next

mRange = "A5:BD" & i + 7
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
End With

'Format amount to 0.00.
mRange = "E7:BD" & i + 7    ' E7:AP
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
With xlApp1.Selection
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .MergeCells = False
    .Style = "Comma"
End With

mRange = "B7:B" & i + 7
xlSheet.Range(mRange).Select
xlApp1.Selection.HorizontalAlignment = xlCenter
mRange = "C7:C" & i + 7
xlSheet.Range(mRange).Select
xlApp1.Selection.HorizontalAlignment = xlCenter
mRange = "E7:E" & i + 7
xlSheet.Range(mRange).Select
xlApp1.Selection.HorizontalAlignment = xlLeft

mRange = "D7:D" & i + 8
xlSheet.Range(mRange).Select
xlApp1.Selection.HorizontalAlignment = xlCenter
mRange = "F7:F" & i + 7
xlSheet.Range(mRange).Select
xlApp1.Selection.HorizontalAlignment = xlCenter
mRange = "G7:G" & i + 7
xlSheet.Range(mRange).Select
xlApp1.Selection.HorizontalAlignment = xlCenter

' Center the BDO.
mRange = "AV7:AV" & i + 7
xlSheet.Range(mRange).Select
xlApp1.Selection.HorizontalAlignment = xlCenter

''Bold the total.
'mRange = "E" & i + 7 & ":BD" & i + 7
'xlSheet.Range(mRange).Select
'xlApp1.ActiveCell.Font.Bold = True

' Box IT.
mRange = "A5:BD" & i + 7
Call PutBox(xlSheet, xlApp1, mRange)

i = i + 5
''Print for Prepared by:, Checked by: and Approved by:
'mRange = "B" & i + 10
'xlSheet.Range(mRange) = "Prepared by:"
'mRange = "F" & i + 10
'xlSheet.Range(mRange) = "Checked by:"
'mRange = "K" & i + 10
'xlSheet.Range(mRange) = "Approved by:"
'
'mRange = "B" & i + 12
'xlSheet.Range(mRange) = "       LAI        "
'xlSheet.Range(mRange).Font.Underline = xlUnderlineStyleSingle
'mRange = "F" & i + 12
'xlSheet.Range(mRange) = "         /JQP  "
'xlSheet.Range(mRange).Font.Underline = xlUnderlineStyleSingle
'mRange = "K" & i + 12
'xlSheet.Range(mRange) = "   EOP/NLW        "
'xlSheet.Range(mRange).Font.Underline = xlUnderlineStyleSingle

'Change the font.
mRange = "E" & i + 2 & ":BD" & i + 14
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = True       'Font.Bold
    .Font.Italic = False    'Font.Italic
End With

With xlApp1.Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .MergeCells = False
    .Style = "Comma"
End With

' Copy Company Information to Sheet2.
xlApp1.ActiveSheet.Rows("2:3").Select
xlApp1.Selection.Copy
xlApp1.Sheets("Sheet2").Select
xlApp1.ActiveSheet.Range("A2").Select
xlApp1.ActiveSheet.Paste
xlApp1.Sheets("Payroll Summary").Select
'Copy Client Header to Sheet 2.
xlApp1.ActiveSheet.Range("A5:A6").Select
xlApp1.Selection.Copy
xlApp1.Sheets("Sheet2").Select
xlApp1.ActiveSheet.Range("A5:A6").Select
xlApp1.ActiveSheet.Paste
xlApp1.Sheets("Payroll Summary").Select
' Copy Column headers to Sheet 2.
xlApp1.ActiveSheet.Range("H5:BD6").Select
xlApp1.Selection.Copy
xlApp1.Sheets("Sheet2").Select
xlApp1.ActiveSheet.Range("B5:AX6").Select
xlApp1.ActiveSheet.Paste

xlApp1.Sheets("Sheet2").Select

' Compute for Grand Total in Sheet 2.
mRange = "A" & LineCtr + 1
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.ActiveSheet.Range(mRange).VALUE = "Grand Total"
For j = Asc("B") To Asc("Z")
    mRange = Chr(j) & 7 & ":" & Chr(j) & LineCtr + 1
    xlApp1.ActiveSheet.Range(mRange).Select
    mRange2 = Chr(j) & LineCtr + 1
    xlApp1.ActiveSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((LineCtr - 6) * -1) & "]C:R[-1]C)"
Next
For j = Asc("A") To Asc("O")
    mRange = "A" & Chr(j) & 7 & ":A" & Chr(j) & LineCtr + 1
    xlApp1.ActiveSheet.Range(mRange).Select
    mRange2 = "A" & Chr(j) & LineCtr + 1
    xlApp1.ActiveSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((LineCtr - 6) * -1) & "]C:R[-1]C)"
Next

For j = Asc("Q") To Asc("X")
    mRange = "A" & Chr(j) & 7 & ":A" & Chr(j) & LineCtr
    xlApp1.ActiveSheet.Range(mRange).Select
    mRange2 = "A" & Chr(j) & LineCtr + 1
    xlApp1.ActiveSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((LineCtr - 6) * -1) & "]C:R[-1]C)"
Next

'Format Sheet 2.
mRange = "A5:AX" & LineCtr + 3
xlApp1.ActiveSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
End With

' Make Grand Total in Bold Font.
mRange = "A" & LineCtr + 1 & ":AX" & LineCtr + 1
xlApp1.ActiveSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = True       'Font.Bold
End With

'Format amount to 0.00.
mRange = "B7:AX" & LineCtr + 3    ' E7:AP
xlApp1.ActiveSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
With xlApp1.Selection
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .MergeCells = False
    .Style = "Comma"
End With
' Box Sheet 2.
mRange = "A5:AX" & LineCtr + 3
Call PutBox(xlSheet, xlApp1, mRange)

mRange = "A5:A6"
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.ActiveSheet.Columns.Range(mRange).ColumnWidth = 30#
mRange = "AP5:AP6"
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.ActiveSheet.Columns.Range(mRange).ColumnWidth = 16
xlApp1.ActiveSheet.Range("A2").Select
xlApp1.ActiveSheet.Name = "Grand Total"
xlApp1.Sheets("Payroll Summary").Select

xlSheet.Range("A2").Select

errR:

    Set xlApp1 = Nothing
    Set xlBook1 = Nothing
    Set xlSheet = Nothing
    Exit Sub
    
Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
    GoTo errR
    
End Sub

Private Sub CopyToSheet2(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal ctr As Integer, ByVal mClientName As String, Optional mDataSw As Boolean = True)

Dim mRange As String

'    xlSheet.Rows("11:11").Select
'    Selection.Copy
'    Sheets("Sheet2").Select
'    Range("A4").Select
'    ActiveSheet.Paste
    
    xlApp1.Sheets("Sheet2").Select
'    If mDataSw = True Then ' Print Company name
'        mRange = "A" & ctr & ":A" & ctr
'        xlApp1.ActiveSheet.Range(mRange).Select
'        xlApp1.ActiveSheet.Paste
'
''        xlSheet.Range(mRange).Select
''        'xlSheet.Rows(mRange).Select
''        xlApp1.ActiveSheet.Paste
'    Else                    ' Print Totals
        mRange = "A" & ctr & ":A" & ctr
        xlApp1.ActiveSheet.Range(mRange).Select
        xlApp1.ActiveSheet.Range(mRange).VALUE = mClientName
        mRange = "B" & ctr & ":AX" & ctr
        xlApp1.ActiveSheet.Range(mRange).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
'    End If
    xlApp1.Sheets("Payroll Summary").Select
    
'    Range("D11:BD11").Select
'    Selection.Copy
'    Sheets("Sheet2").Select
'    Range("C12").Select
'    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'        False, Transpose:=False
    
'    xlApp1.ActiveSheet.Paste
'    Selection.Copy
'    xlSheet.Sheets("Sheet2").Select
'    Rows("10:10").Select
'    Range("G10").Select

End Sub

Private Sub PrintTotals(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal i As Integer, ByVal k As Integer)

Dim EndSw As Boolean
Dim sumString As String
Dim mRange As String
Dim j As Integer
Dim addno As Integer
Dim mRange2 As String

addno = 6   '7

'Compute for Grand Total.
EndSw = False
'sumString = "=SUM(R[" & i * -1 & "]C:R[-2]C)"
sumString = "=SUM(R[" & (i - k) * -1 & "]C:R[-1]C)"
'mRange = "E" & i + addno & ":E" & i + addno
mRange = "E" & i & ":E" & i
xlSheet.Range(mRange).Select
xlApp1.ActiveCell.FormulaR1C1 = "           TOTAL"
'xlApp1.Range(mRange).Font.Bold = True

'Print Total Employee.
mRange = "D" & i & ":D" & i
xlSheet.Range(mRange).Select
xlApp1.ActiveCell.FormulaR1C1 = IIf(i = k, i - 1, i - k)
'xlApp1.Range(mRange).Font.Bold = True

For j = Asc("H") To Asc("Z")
'    If EndSw = False And Chr(j) = "Z" Then
'        EndSw = True
'    End If
'    If EndSw = True Then
'        If Chr(j) = "Z" Then
'            mRange = Chr(j) & k + addno & ":" & Chr(j) & i + addno
'        Else
'            mRange = "A" & Chr(j) & k + addno & ":A" & Chr(j) & i + addno
'        End If
'    Else
'        mRange = Chr(j) & k & ":" & Chr(j) & i
'    End If
    mRange = Chr(j) & k & ":" & Chr(j) & i
    xlSheet.Range(mRange).Select
    mRange2 = Chr(j) & i
    xlSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
 '   xlApp1.Range(mRange2).Font.Bold = True
    xlSheet.Range(mRange).Select
Next
For j = Asc("A") To Asc("U")    ' K
    mRange = "A" & Chr(j) & k & ":A" & Chr(j) & i
    xlSheet.Range(mRange).Select
    mRange2 = "A" & Chr(j) & i
    xlSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
  '  xlApp1.Range(mRange2).Font.Bold = True
    xlSheet.Range(mRange).Select
Next
For j = Asc("W") To Asc("Z")    ' K
    mRange = "A" & Chr(j) & k & ":A" & Chr(j) & i
    xlSheet.Range(mRange).Select
    mRange2 = "A" & Chr(j) & i
    xlSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
   ' xlApp1.Range(mRange2).Font.Bold = True
    xlSheet.Range(mRange).Select
Next
For j = Asc("A") To Asc("D")    ' K
    mRange = "B" & Chr(j) & k & ":B" & Chr(j) & i
    xlSheet.Range(mRange).Select
    mRange2 = "B" & Chr(j) & i
    xlSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
    'xlApp1.Range(mRange2).Font.Bold = True
    xlSheet.Range(mRange).Select
Next
End Sub
    
Private Sub BoxDetails(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal mRange2 As String)

xlSheet.Range(mRange2).Select
With xlApp1
    .Selection.Columns.AutoFit
    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With xlApp1.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End With
        With xlSheet.Range(mRange2)
            .Font.Name = "Arial Narrow"   'Font.Name
            .Font.Size = 10         'Font.size
        End With

End Sub

Private Sub ComputeTotals(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal ctr As Integer, ByVal i As Integer, ByVal StartOfLine As Integer, Optional LoanFlag As Boolean = False)

Dim mRange As String
Dim mRange2 As String
Dim sumString As String

mRange = "A" & ctr + 1
If StartOfLine = 8 Then
    StartOfLine = 6
End If
If LoanFlag = False Then
    xlSheet.Range(mRange).VALUE = "Total:"
End If
Select Case ReportType
    Case 1  ' SSS
        If chkLoans.VALUE = 1 Then
            If LoanFlag = True Then
                mRange = "F" & ctr
            Else
                mRange = "F" & ctr + 1
            End If
            mRange2 = "A" & StartOfLine & ":H" & ctr + 1
        Else
            mRange = "F" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            xlApp1.Selection.NumberFormat = "#,##0.00"
            mRange = "G" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            xlApp1.Selection.NumberFormat = "#,##0.00"
            mRange = "H" & ctr + 1
            mRange2 = "A" & StartOfLine & ":J" & ctr + 1
        End If
    Case 2  ' Tax
        mRange = "F" & ctr + 1
        mRange2 = "A" & StartOfLine & ":H" & ctr + 1
    Case 3  ' HDMF
        If chkLoans.VALUE = 1 Then
            mRange = "F" & ctr + 1
            mRange2 = "A" & StartOfLine & ":H" & ctr + 1
        Else
            mRange = "F" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            xlApp1.Selection.NumberFormat = "#,##0.00"
            mRange = "G" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            xlApp1.Selection.NumberFormat = "#,##0.00"
            mRange = "H" & ctr + 1
            'mrange2 = "A7:J" & ctr + 1
            mRange2 = "A" & StartOfLine & ":J" & ctr + 1
        End If
    Case 4  ' PhilHealth
        mRange = "F" & ctr + 1
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        mRange = "G" & ctr + 1
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        mRange = "H" & ctr + 1
        'mrange2 = "A7:J" & ctr + 1
        mRange2 = "A" & StartOfLine & ":J" & ctr + 1
    Case 5  ' 13th Month
        mRange = "E" & ctr + 1
        'mrange2 = "A7:G" & ctr + 1
        mRange2 = "A" & StartOfLine & ":G" & ctr + 1
    Case 6  ' Cash Bond
        mRange = "E" & ctr + 1
        'mrange2 = "A7:G" & ctr + 1
        mRange2 = "A" & StartOfLine & ":G" & ctr + 1
    Case 7  ' Head Count
    Case 8  ' Other Deductions
        mRange = "E" & ctr + 1
        xlSheet.Range(mRange).Select
        'ADDED BY RAMIL 3/10/2007
        If ESumString = "" Then ESumString = "=SUM(" & mRange Else ESumString = ESumString & "," & mRange
        
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        mRange = "F" & ctr + 1
        'ADDED BY RAMIL 3/10/2007
        If FSumString = "" Then FSumString = "=SUM(" & mRange Else FSumString = FSumString & "," & mRange
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
       
        mRange = "G" & ctr + 1
        'ADDED BY RAMIL 3/10/2007
        If GSumString = "" Then GSumString = "=SUM(" & mRange Else GSumString = GSumString & "," & mRange
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        mRange = "H" & ctr + 1
        'ADDED BY RAMIL 3/10/2007
        If HSumString = "" Then HSumString = "=SUM(" & mRange Else HSumString = HSumString & "," & mRange
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        mRange = "I" & ctr + 1
        'ADDED BY RAMIL 3/10/2007
        If ISumString = "" Then ISumString = "=SUM(" & mRange Else ISumString = ISumString & "," & mRange
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        mRange = "J" & ctr + 1
        'ADDED BY Bernard April 19, 2008
        If JSumString = "" Then JSumString = "=SUM(" & mRange Else JSumString = JSumString & "," & mRange
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        mRange = "K" & ctr + 1
        'ADDED BY Bernard April 19, 2008
        If KSumString = "" Then KSumString = "=SUM(" & mRange Else KSumString = KSumString & "," & mRange
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        mRange = "L" & ctr + 1
        'ADDED BY Bernard April 19, 2008
        If LSumString = "" Then LSumString = "=SUM(" & mRange Else LSumString = LSumString & "," & mRange
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        mRange = "M" & ctr + 1
        'ADDED BY Bernard April 19, 2008
        If MSumString = "" Then MSumString = "=SUM(" & mRange Else MSumString = MSumString & "," & mRange
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        mRange = "N" & ctr + 1
        mRange2 = "A" & StartOfLine & ":O" & ctr + 1 'L
    End Select
        
    xlSheet.Range(mRange).Select
    If LoanFlag = True Then
        If LoanStart <> 0 Then StartOfLine = LoanStart + 1
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
    Else
        If ReportType = 1 And chkLoans.VALUE = 1 Then
            xlSheet.Range(mRange).VALUE = TotalLoan
        Else
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
        End If
    End If
    xlApp1.Selection.NumberFormat = "#,##0.00"
    If LoanFlag = False Then
        xlSheet.Range(mRange2).Select
        Call BoxDetails(xlSheet, xlApp1, mRange2)
    End If
    LoanStart = ctr + 1
End Sub

