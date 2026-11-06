VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmDedMonitoring 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Deductions Monitoring"
   ClientHeight    =   6795
   ClientLeft      =   1155
   ClientTop       =   1485
   ClientWidth     =   8580
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmDedMonitoring.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   8580
   Begin TabDlg.SSTab SSTab1 
      Height          =   6165
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   540
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   10874
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Employee Basic Salary"
      TabPicture(0)   =   "frmDedMonitoring.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   5820
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   8175
         Begin VB.CheckBox chkReportType 
            Caption         =   "Cash Bond"
            Height          =   285
            Index           =   0
            Left            =   480
            TabIndex        =   15
            Top             =   4200
            Value           =   1  'Checked
            Width           =   1230
         End
         Begin VB.CommandButton cmdGenerate 
            Caption         =   "Generate Report"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   14
            Top             =   5280
            Width           =   1455
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6480
            TabIndex        =   12
            Top             =   5280
            Width           =   1455
         End
         Begin VB.CommandButton cmdShowContracts 
            Caption         =   "Show Contracts"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6720
            TabIndex        =   10
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox cmbEmployees 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            ItemData        =   "frmDedMonitoring.frx":05A6
            Left            =   4080
            List            =   "frmDedMonitoring.frx":05B3
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   480
            Width           =   3840
         End
         Begin VB.ComboBox cmbCompany 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            ItemData        =   "frmDedMonitoring.frx":05D6
            Left            =   120
            List            =   "frmDedMonitoring.frx":05E3
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   480
            Width           =   3720
         End
         Begin VB.TextBox txtEmpNum 
            Height          =   285
            Left            =   1320
            TabIndex        =   6
            Text            =   "txtclientcode"
            Top             =   180
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtClientCode 
            Height          =   285
            Left            =   2160
            TabIndex        =   5
            Text            =   "txtclientcode"
            Top             =   180
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2535
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   7950
            _ExtentX        =   14023
            _ExtentY        =   4471
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
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Agency Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Client Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Control No."
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Start Contract"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "End Contract"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Total Cash Bond"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Cash Bond Payment"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Start CB Payment"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "End CB Payment"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   " Employee Contracts:"
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
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   1845
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Employee:"
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
            Height          =   255
            Index           =   0
            Left            =   4080
            TabIndex        =   9
            Top             =   240
            Width           =   2325
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Client:"
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
            Height          =   255
            Index           =   38
            Left            =   180
            TabIndex        =   4
            Top             =   240
            Width           =   1365
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9240
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDedMonitoring.frx":0606
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   800
         Y1              =   1500
         Y2              =   1200
      End
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "Agency Rate:"
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
      Height          =   255
      Index           =   23
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmDedMonitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public i As Integer
Dim rsEmployeeBasicData As ADODB.Recordset
Dim rsClientData As ADODB.Recordset
Dim rsCurrentContract As ADODB.Recordset
Dim rsContractHistory As ADODB.Recordset

Dim StrQuery As String
Dim mRecNo As Integer
Dim mClientCode As Integer
Dim mCompany As String
Dim mEmpName As String
Dim mEmpNo As Integer
Dim ActiveCtr As Integer
Dim InActiveCtr As Integer
Dim RefreshSw As Boolean
Dim paymentbase As String
Dim CurrEmployee As String
Dim CurrAmount As String    ' holds the current value of salaries
Dim AmountSw As Boolean     ' variable for basic salary entries
Dim ActiveSw As Boolean
Dim FormLoadSw As Boolean

Private Sub cmbCompany_Click()
    Call FilterEmployee(cmbCompany.Text)
End Sub

Private Sub cmbEmployees_Click()
    mEmpNo = cmbEmployees.ItemData(cmbEmployees.ListIndex)
End Sub

Private Sub cmdClose_Click()
    Set rsClientData = Nothing
    Set rsEmployeeBasicData = Nothing
    Set rsCurrentContract = Nothing
    Set rsContractHistory = Nothing
    Unload Me
End Sub

Private Sub cmdGenerate_Click()

    Call GenerateReport

End Sub

Private Sub cmdShowContracts_Click()
Dim mCurrentAgency As String

mCurrentAgency = """Current Agency (" & CompanyShort & ")"""
StrQuery = "SELECT a.ClientCode, " & mCurrentAgency & " AS AgencyName, a.EmployeeID, a.EmpNo, a.AgencyCode, a.StartContract, " & _
    " a.EndContract, a.CBAmount, a.CBPayment, a.CBStartPayment, a.CBEndPayment, a.Remarks, a.LastUpdate, " & _
    " b.ClientName FROM HREmployees AS a LEFT JOIN ClientTable AS b ON a.clientCode = b.clientCode " & _
    " WHERE a.EmployeeID = " & mEmpNo & " ORDER BY a.startContract DESC"
Set rsCurrentContract = ReturnRst(StrQuery, cnnstr, mstatic)
Call FillEmployeeList(ListView1(0), rsCurrentContract, False)

StrQuery = "SELECT ClientCode, AgencyName, EmployeeID, EmpNo, AgencyCode, StartContract, " & _
    " EndContract, CBAmount, CBPayment, CBStartPayment, CBEndPayment, Remarks, LastUpdate, " & _
    " ClientName FROM HREmployees_History WHERE EmployeeID = '" & mEmpNo & "' ORDER BY startContract DESC"

'strQuery = "SELECT b.ClientCode, a.AgencyName, a.EmployeeID, a.EmpNo, a.AgencyCode, a.StartContract, " & _
'    " a.EndContract, a.CBAmount, a.CBPayment, a.CBStartPayment, a.CBEndPayment, a.Remarks, a.LastUpdate, " & _
'    " b.ClientName FROM HREmployees_History a INNER JOIN ClientTable b ON a.ClientCode = b.ClientCode " & _
'    " WHERE a.EmployeeID = '" & mEmpNo & "' ORDER BY startContract DESC"

Set rsContractHistory = ReturnRst(StrQuery, CnnstrHR, mstatic)
Call FillEmployeeList(ListView1(0), rsContractHistory)
End Sub

Private Sub Form_Load()
Dim mComp As String
Dim mEmpN As String

    If ActiveSw = False Then
        Me.Top = (mdiMain.Height - Me.Height) / 8
        Me.Left = (mdiMain.Width - Me.Width) / 2
    End If
    FormLoadSw = False
    
    ' Get Account Name.
    cmbCompany.Enabled = True
    mUserID = UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID)))
    ActiveSw = True
    RefreshSw = False
    Call GetEmployeeBasic
    FormLoadSw = True
    
End Sub

Private Sub GetEmployeeBasic()
Dim rsTaxCode As ADODB.Recordset

' Get all active company.
StrQuery = "SELECT * FROM clienttable WHERE clientstatus = 1 and clientname is not null ORDER BY clientname"
Set rsClientData = ReturnRst(StrQuery, cnnstr, mstatic)

If rsClientData.RecordCount <> 0 Then
    Call FillComboData(cmbCompany, rsClientData)
    cmbCompany.ListIndex = 0
    rsClientData.MoveFirst
    mClientCode = rsClientData!ClientCode
Else
    MsgBox "No Client record selected. Cannot proceed.", vbInformation, PayrollAlert
    Exit Sub
End If

'strQuery = "SELECT a.EmployeeID AS Empno, a.EmpName As Employee, b.ClientName, a.* FROM qryHRDeductions AS a " & _
'            "LEFT JOIN ClientTable AS b ON a.clientcode = b.clientcode " & _
'            "WHERE a.EmploymentStatus not in ('T','W','X','Y','Z')" & _
'            " ORDER BY a.EmpName"

StrQuery = "SELECT a.EmployeeID AS Empno, a.EmpName As Employee, b.ClientName, a.* FROM qryHRDeductions AS a " & _
            "LEFT JOIN ClientTable AS b ON a.clientcode = b.clientcode " & _
            " ORDER BY a.EmpName"

Set rsEmployeeBasicData = ReturnRst(StrQuery, cnnstr, mstatic)
If rsEmployeeBasicData.RecordCount <> 0 Then
    Call FillComboData(cmbEmployees, rsEmployeeBasicData, True)
    cmbEmployees.ListIndex = 0
    Call FilterEmployee(rsClientData!ClientName)
    mEmpNo = cmbEmployees.ItemData(cmbEmployees.ListIndex)
Else
    MsgBox "No Employee information selected. Cannot proceed.", vbInformation, PayrollAlert
    Exit Sub
End If

End Sub

Private Sub FilterEmployee(ByVal mCompany As String, Optional NotYet As Boolean = False)
    rsClientData.MoveFirst
    rsClientData.Find "clientname = '" & Trim(mCompany) & "'"
    If rsClientData.EOF = True Then
        mClientCode = 0
        Exit Sub
    Else
        mClientCode = rsClientData!ClientCode
        txtClientCode = rsClientData!ClientCode
    End If
    If FormLoadSw = True Then
    With rsEmployeeBasicData
        If Trim(mCompany) <> "*** All Clients ***" Then
            .Filter = "clientcode = " & mClientCode
        Else
            .Filter = ""
        End If
        If .RecordCount <> 0 Then
            'If Not .EOF And Not .BOF Then
            Call FillComboData(cmbEmployees, rsEmployeeBasicData, True)
            cmbEmployees.ListIndex = 0
            mEmpNo = cmbEmployees.ItemData(cmbEmployees.ListIndex)
        Else
            Call clearList
        End If
    End With
    End If
End Sub

Private Sub FillEmployeeList(ByRef mLsEmp As ListView, _
                        ByRef mRs As Recordset, Optional AppendSw As Boolean = True)
Dim litem1 As ListItem

    InActiveCtr = 0
    ActiveCtr = 0
    'mLsEmp.View = lvwReport
    ListView1(1).View = lvwReport
    If AppendSw = False Then
        'mLsEmp.ListItems.Clear
        ListView1(1).ListItems.Clear
    End If
    'Get all employees contracts regardless of agency.
    'mRs.MoveFirst
    While Not (mRs.EOF Or mRs.BOF)
        Set litem1 = ListView1(1).ListItems.Add(, , mRs!ClientCode)
        litem1.SubItems(1) = IIf(IsNull(mRs!AgencyName), "Blank", mRs!AgencyName)
        litem1.SubItems(2) = IIf(IsNull(mRs!ClientName), "Blank", mRs!ClientName)
        litem1.SubItems(3) = IIf(IsNull(mRs!EmpNo), "Blank", mRs!EmpNo)
        litem1.SubItems(4) = IIf(IsNull(mRs!StartContract) Or mRs!StartContract = #1/1/1900#, "__/__/____", Format(mRs!StartContract, "mm/dd/yyyy"))
        litem1.SubItems(5) = IIf(IsNull(mRs!endcontract) Or mRs!endcontract = #1/1/1900#, "__/__/____", Format(mRs!endcontract, "mm/dd/yyyy"))
        litem1.SubItems(6) = Format(mRs!Cbamount, "0.##")
        litem1.SubItems(7) = Format(mRs!CBPayment, "0.##")
        litem1.SubItems(8) = IIf(IsNull(mRs!CBStartPayment) Or mRs!CBStartPayment = #1/1/1900#, "__/__/____", Format(mRs!CBStartPayment, "mm/dd/yyyy"))
        litem1.SubItems(9) = IIf(IsNull(mRs!CBEndPayment) Or mRs!CBEndPayment = #1/1/1900#, "__/__/____", Format(mRs!CBEndPayment, "mm/dd/yyyy"))
        mRs.MoveNext
    Wend
End Sub

Private Function FillComboData(ByRef mCmb As ComboBox, _
                        mRs As Recordset, Optional FilterFlag As Boolean = False, _
                        Optional ComboList As String = "")
    mCmb.Clear
    If FilterFlag = False Then
        mCmb.AddItem "*** All Clients ***"
    End If
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

Private Sub Form_Unload(Cancel As Integer)
ActiveSw = False
FormLoadSw = False
End Sub

Private Sub ListView1_Click(Index As Integer)
Dim res
    
'    If ActiveCtr = 0 Then
'        Call DisplayBasicData("", False)
'    Else
'        If SSTab1(1).Tab = 1 Then ' in-active tab
'            If ListView1(1).ListItems.Count = 0 Then
'                Call DisplayBasicData("", False)
'            Else
'                Call DisplayBasicData(ListView1(1).SelectedItem.Text, False)
'                CurrEmployee = ListView1(1).SelectedItem.Text
'            End If
'        Else
'            If ListView1(0).ListItems.Count = 0 Then
'                Call DisplayBasicData("", False)
'            Else
'                Call DisplayBasicData(ListView1(0).SelectedItem.Text, False)
'                CurrEmployee = ListView1(0).SelectedItem.Text
'                If AddDeductions = True Or EditDeductions = True Then
'                    'res = MsgBox("Cancel changes to deductions?", vbQuestion + vbYesNo, PayrollAlert)
'                    'If res = vbYes Then ' cancel the changes
'                        Call EnableDisableDeductions(False)
'                        Call EnableDisableDeductionButtons("Add", "Edit", True)
'                        EditDeductions = False
'                        AddDeductions = False
'                        cmbTypes.Enabled = True
'                        MsgBox "Changes to deductions was cancelled!", vbInformation, PayrollAlert
'                    'End If
'                End If
'            End If
'        End If
'    End If
End Sub

Private Sub ListView1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1(Index).SortKey = ColumnHeader.Index - 1
    ListView1(Index).Sorted = True
End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    CurrEmployee = Item.Text
End Sub

Private Sub clearList()
    ListView1(1).ListItems.Clear
End Sub

Private Sub GenerateReport()
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
Dim StartDeduction As Variant
Dim EndDeduction As Variant
Dim DeductionYear As Integer

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

'Updated by Bernard May 13, 2008.
mHead = CompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Cash Bond Summary Report"
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy")
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "COMPANY NAME: " & cmbCompany.Text
Call PrintReportHeader(xlSheet, xlApp1, "A5", mHead, "", xlLeft, , False, False, False, , , 10, True)

Call PrintReportHeader(xlSheet, xlApp1, "A6:A7", "Agency Name", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 20)
Call PrintReportHeader(xlSheet, xlApp1, "B6:B7", "Client Name", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 25)
Call PrintReportHeader(xlSheet, xlApp1, "C6:C7", "Control No.", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 11)
Call PrintReportHeader(xlSheet, xlApp1, "D6:D7", "Employee Name", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 23)
Call PrintReportHeader(xlSheet, xlApp1, "E6:E7", "Start Contract", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 6.8)
Call PrintReportHeader(xlSheet, xlApp1, "F6:F7", "End Contract", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 6.8)
Call PrintReportHeader(xlSheet, xlApp1, "G6:G7", "Cash Bond Total", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
Call PrintReportHeader(xlSheet, xlApp1, "H6:H7", "Cash Bond Payment", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
Call PrintReportHeader(xlSheet, xlApp1, "I6:I7", "Start", "Cut-Off", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)
Call PrintReportHeader(xlSheet, xlApp1, "J6:J7", "End", "Cut-Off", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)

For i = 1 To ListView1(1).Index
    If ListView1(1).ListItems.Item(i).Checked = True And Val(ListView1(1).ListItems.Item(i).SubItems(6)) Then
'        StartDeduction = Format(ListView1(1).ListItems.Item(i).SubItems(4), "mm/dd/yyyy") ' Start of Contract
'        EndDeduction = Format(ListView1(1).ListItems.Item(i).SubItems(5), "mm/dd/yyyy") ' End of Contract
        
        StartDeduction = Val(Format(ListView1(1).ListItems.Item(i).SubItems(4), "yyyymmdd")) ' Start of Contract
        EndDeduction = Val(Format(ListView1(1).ListItems.Item(i).SubItems(5), "yyyymmdd")) ' End of Contract
        
        
        'DeductionYear = Format(ListView1(1).ListItems.Item(i).SubItems(8), "yyyy")
'        strQuery = "SELECT EmployeeID, TASAesla, StartPaydate, EndPayDate FROM qryContractHistory WHERE EmployeeID = " & _
'            mEmpNo & " AND StartPaydate >= '" & StartDeduction & _
'            "' AND EndPayDate <= '" & EndDeduction & "' ORDER BY StartPayDate"
        
        StrQuery = "SELECT EmployeeID, TASAesla, StartPaydate, EndPayDate FROM qryContractHistory WHERE EmployeeID = " & _
            mEmpNo & " AND (StartPayNum between " & StartDeduction & _
            " AND " & EndDeduction & " OR EndPayDate Between " & StartDeduction & _
            " AND " & EndDeduction & ") ORDER BY StartPayDate"
        
        
        Set rsDedRecords = ReturnRst(StrQuery, cnnstr, mstatic)
        If rsDedRecords.EOF = False Then
            xlSheet.Cells(ctr, 1).VALUE = ListView1(1).ListItems.Item(i).SubItems(1) ' Agency Name
            xlSheet.Cells(ctr, 2).VALUE = ListView1(1).ListItems.Item(i).SubItems(2) ' Client Name
            xlSheet.Cells(ctr, 3).VALUE = ListView1(1).ListItems.Item(i).SubItems(3) ' Control Number
            xlSheet.Cells(ctr, 4).VALUE = cmbEmployees.Text  ' Employee Name
            xlSheet.Cells(ctr, 5).VALUE = ListView1(1).ListItems.Item(i).SubItems(4) ' Start Contract
            xlSheet.Cells(ctr, 6).VALUE = ListView1(1).ListItems.Item(i).SubItems(5) ' End Contract
            xlSheet.Cells(ctr, 7).VALUE = Format(ListView1(1).ListItems.Item(i).SubItems(6), "#,##0.00") ' Total Cash Bond
            Do While Not rsDedRecords.EOF
                xlSheet.Cells(ctr, 8).VALUE = Format(rsDedRecords!TASAesla, "#,##0.00") ' cash bond Payment
                xlSheet.Cells(ctr, 9).VALUE = Format(rsDedRecords!StartPayDate, "mm/dd/yyyy") ' start cutoff
                xlSheet.Cells(ctr, 10).VALUE = Format(rsDedRecords!EndPayDate, "mm/dd/yyyy") ' end cutoff
                ctr = ctr + 1
                rsDedRecords.MoveNext
            Loop
            ' Print Sub-Totals.
            xlSheet.Cells(ctr, 4).VALUE = "     Total Payments"
            mRange = "H" & 8 & ":H" & ctr
            xlSheet.Range(mRange).Select
            mRange = "H" & ctr
            xlSheet.Range(mRange).Activate
            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-1]C)"
            xlApp1.Range(mRange).Font.Bold = True
            mRange = "G8:H" & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
        Else
            xlSheet.Cells(ctr, 1).VALUE = ListView1(1).ListItems.Item(i).SubItems(1) ' Agency Name
            xlSheet.Cells(ctr, 2).VALUE = ListView1(1).ListItems.Item(i).SubItems(2) ' Client Name
            xlSheet.Cells(ctr, 3).VALUE = ListView1(1).ListItems.Item(i).SubItems(3) ' Control Number
            xlSheet.Cells(ctr, 4).VALUE = cmbEmployees.Text  ' Employee Name
            xlSheet.Cells(ctr, 5).VALUE = ListView1(1).ListItems.Item(i).SubItems(4) ' Start Contract
            xlSheet.Cells(ctr, 6).VALUE = ListView1(1).ListItems.Item(i).SubItems(5) ' End Contract
            xlSheet.Cells(ctr, 7).VALUE = Format(ListView1(1).ListItems.Item(i).SubItems(6), "#,##0.00") ' Total Cash Bond
            ctr = ctr + 1
            xlSheet.Cells(ctr, 1).VALUE = "No payments found."
        End If
    Else
        xlSheet.Cells(ctr, 1).VALUE = ListView1(1).ListItems.Item(i).SubItems(1) ' Agency Name
        xlSheet.Cells(ctr, 2).VALUE = ListView1(1).ListItems.Item(i).SubItems(2) ' Client Name
        xlSheet.Cells(ctr, 3).VALUE = ListView1(1).ListItems.Item(i).SubItems(3) ' Control Number
        xlSheet.Cells(ctr, 4).VALUE = cmbEmployees.Text  ' Employee Name
        xlSheet.Cells(ctr, 5).VALUE = ListView1(1).ListItems.Item(i).SubItems(4) ' Start Contract
        xlSheet.Cells(ctr, 6).VALUE = ListView1(1).ListItems.Item(i).SubItems(5) ' End Contract
        xlSheet.Cells(ctr, 7).VALUE = Format(ListView1(1).ListItems.Item(i).SubItems(6), "#,##0.00") ' Total Cash Bond
        ctr = ctr + 1
        xlSheet.Cells(ctr, 1).VALUE = "Zero (0) deductions."
    End If
Next i
ctr = ctr + 2

'Box it.
mRange = "A6:J" & ctr
Call PlaceBox(xlSheet, xlApp1, mRange)

ctr = ctr + 1
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
mRange = "A" & ctr
Call CustomizedHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 8, False, 25)

ctr = ctr + 3
xlSheet.Cells(ctr, 1).VALUE = "Prepared by: ______________"
xlSheet.Cells(ctr, 3).VALUE = "Note by: ______________"
xlSheet.Cells(ctr, 7).VALUE = "Received by: ______________"
xlSheet.Range("A1").Select


'
'
'Do While Not rs.EOF
'    ' DoEvents
'    xlSheet.Cells(ctr, 1).VALUE = i
'    xlSheet.Cells(ctr, 2).VALUE = rs!EmpName
'    mDeduc = 0#
''    If rs!CbAmount <> 0 Then
''        mDeduc = RunningBalance(rsDedRecords, 1, rs!EmployeeID, rs!cbstartpayment, rs!cbendpayment)
''        xlSheet.Cells(ctr, 8).VALUE = mDeduc
''    End If
'    xlSheet.Cells(ctr, 3).VALUE = Format(rs!DailyRate, "0.##")
'    xlSheet.Cells(ctr, 20).VALUE = Format(rs!StartContract, "mm/dd/yyyy")
'    xlSheet.Cells(ctr, 21).VALUE = Format(rs!endcontract, "mm/dd/yyyy")
'    xlSheet.Cells(ctr, 22).VALUE = rs!AccountNo
'    xlSheet.Cells(ctr, 23).VALUE = rs!Remarks
'    ctr = ctr + 1
'    i = i + 1
'    rs.MoveNext
'Loop
'mRange = "B" & ctr + 1 & ":B" & ctr + 1
'Call PrintReportHeader(xlSheet, xlApp1, mRange, "sub-total", "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 22)
'mRange = "B" & ctr + 2 & ":B" & ctr + 2
'Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total of Staff", "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 22)
'mRange = "D" & ctr + 2 & ":D" & ctr + 2
'Call PrintReportHeader(xlSheet, xlApp1, mRange, i - 1, "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 5)
'mRange = "B" & ctr + 3 & ":B" & ctr + 3
'Call PrintReportHeader(xlSheet, xlApp1, mRange, "No. of lacking", "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 22)

ExitRtn2:

End Sub

