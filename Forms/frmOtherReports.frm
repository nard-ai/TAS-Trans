VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmOtherReport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Contributions and Deductions Report"
   ClientHeight    =   4650
   ClientLeft      =   4140
   ClientTop       =   3360
   ClientWidth     =   7740
   Icon            =   "frmOtherReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   7740
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
      Left            =   300
      TabIndex        =   13
      Top             =   3720
      Width           =   3660
      Begin VB.ComboBox cmbClient 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   270
         Width           =   3345
      End
   End
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   4110
      TabIndex        =   12
      Top             =   3720
      Width           =   3315
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   330
         Left            =   1800
         TabIndex        =   21
         Top             =   265
         Width           =   1290
      End
      Begin VB.CommandButton cmdProcess 
         Caption         =   "Generate"
         Height          =   330
         Left            =   240
         TabIndex        =   20
         Top             =   265
         Width           =   1290
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Payroll Cutoff"
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
      Left            =   2160
      TabIndex        =   9
      Top             =   1995
      Width           =   5280
      Begin VB.ComboBox cmbMonthof 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmOtherReports.frx":08CA
         Left            =   480
         List            =   "frmOtherReports.frx":08F2
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   720
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.ComboBox cmbEndMonth 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmOtherReports.frx":0958
         Left            =   3240
         List            =   "frmOtherReports.frx":0980
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   795
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.ComboBox cmbQuarter 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmOtherReports.frx":09E6
         Left            =   2880
         List            =   "frmOtherReports.frx":09F6
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.ComboBox cmbYearof 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmOtherReports.frx":0A2E
         Left            =   4425
         List            =   "frmOtherReports.frx":0A4A
         TabIndex        =   10
         Text            =   "cmbYearof"
         Top             =   555
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtStartDate 
         Height          =   315
         Left            =   1080
         TabIndex        =   34
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   144048129
         CurrentDate     =   40262
      End
      Begin MSComCtl2.DTPicker dtEndDate 
         Height          =   315
         Left            =   3360
         TabIndex        =   35
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   144048129
         CurrentDate     =   40262
      End
      Begin VB.Label Label1 
         Caption         =   "End Date:"
         Height          =   180
         Index           =   0
         Left            =   2520
         TabIndex        =   23
         Top             =   405
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Year"
         Height          =   180
         Index           =   2
         Left            =   3930
         TabIndex        =   19
         Top             =   405
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Start Date:"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   18
         Top             =   405
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Type"
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
      Left            =   300
      TabIndex        =   5
      Top             =   1995
      Width           =   1815
      Begin VB.OptionButton OptScope 
         Caption         =   "Gross Summary"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   7
         Top             =   900
         Width           =   1530
      End
      Begin VB.OptionButton OptScope 
         Caption         =   "Per Cut-off"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   105
         TabIndex        =   24
         Top             =   320
         Width           =   1530
      End
      Begin VB.OptionButton OptScope 
         Caption         =   "Annual Report"
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   105
         TabIndex        =   8
         Top             =   1125
         Width           =   1530
      End
      Begin VB.OptionButton OptScope 
         Caption         =   "Detailed Report"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   6
         Top             =   600
         Value           =   -1  'True
         Width           =   1530
      End
   End
   Begin VB.Frame FraReport 
      Caption         =   "Report Selection"
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
      Left            =   360
      TabIndex        =   0
      Top             =   795
      Width           =   7125
      Begin VB.CheckBox chkAllTax 
         Caption         =   "All"
         Height          =   195
         Left            =   6480
         TabIndex        =   47
         Top             =   720
         Width           =   555
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "Tax"
         Height          =   195
         Index           =   2
         Left            =   1170
         TabIndex        =   37
         Top             =   300
         Width           =   810
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "13th Month"
         Height          =   195
         Index           =   5
         Left            =   1170
         TabIndex        =   36
         Top             =   720
         Width           =   1170
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "Alphalist/Tax"
         Height          =   195
         Index           =   9
         Left            =   5160
         TabIndex        =   29
         Top             =   720
         Width           =   1410
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "Deduction Payments"
         Height          =   195
         Index           =   8
         Left            =   5160
         TabIndex        =   28
         Top             =   300
         Width           =   1770
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "Masterlist"
         Height          =   195
         Index           =   7
         Left            =   3750
         TabIndex        =   27
         Top             =   720
         Width           =   1290
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "Payroll"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   26
         Top             =   720
         Width           =   930
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
         ItemData        =   "frmOtherReports.frx":0A7E
         Left            =   7170
         List            =   "frmOtherReports.frx":0A80
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "Absent/Lates"
         Height          =   195
         Index           =   6
         Left            =   2340
         TabIndex        =   4
         Top             =   720
         Width           =   1290
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "PhilHealth"
         Height          =   195
         Index           =   4
         Left            =   3750
         TabIndex        =   3
         Top             =   300
         Width           =   1050
      End
      Begin VB.OptionButton OptReport 
         Caption         =   "HDMF"
         Height          =   195
         Index           =   3
         Left            =   2340
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
      TabIndex        =   15
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
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
      Left            =   2160
      TabIndex        =   11
      Top             =   2820
      Width           =   5280
      Begin VB.CheckBox chkCal 
         Caption         =   "For Calamity"
         Height          =   195
         Left            =   1200
         TabIndex        =   59
         Top             =   480
         Width           =   1395
      End
      Begin VB.Frame Frame6 
         Height          =   825
         Left            =   2040
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   5295
         Begin VB.CommandButton cmdTotals 
            Caption         =   "Compute Totals"
            Height          =   330
            Left            =   3480
            TabIndex        =   58
            Top             =   360
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.TextBox txtTax 
            Height          =   305
            Index           =   1
            Left            =   1800
            TabIndex        =   52
            Text            =   "0.00"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtTax 
            Height          =   305
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Text            =   "0.00"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Total Tax Withheld"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   53
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Total Compensation"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   120
            Width           =   1575
         End
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
         ItemData        =   "frmOtherReports.frx":0A82
         Left            =   2760
         List            =   "frmOtherReports.frx":0A84
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   165
         Width           =   2400
      End
      Begin VB.CheckBox chkLoans 
         Caption         =   "For Loans"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   1395
      End
      Begin VB.CheckBox chkAllEmployees 
         Caption         =   "All Employees"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   120
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Selected Only:"
         Height          =   300
         Index           =   3
         Left            =   1320
         TabIndex        =   56
         Top             =   240
         Width           =   1275
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   4065
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   7170
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "_"
      TabPicture(0)   =   "frmOtherReports.frx":0A86
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame5 
         Height          =   2415
         Left            =   120
         TabIndex        =   30
         Top             =   3840
         Visible         =   0   'False
         Width           =   7215
         Begin VB.TextBox txtChangeID 
            Height          =   285
            Left            =   480
            TabIndex        =   45
            Text            =   "0.00"
            Top             =   1560
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtChange 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   5040
            TabIndex        =   43
            Text            =   "0.00"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CommandButton cmd13thMonth 
            Caption         =   "Return"
            Height          =   330
            Index           =   5
            Left            =   6000
            TabIndex        =   42
            Top             =   1920
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.CommandButton cmd13thMonth 
            Caption         =   "Save"
            Height          =   330
            Index           =   4
            Left            =   6240
            TabIndex        =   41
            Top             =   1560
            Width           =   810
         End
         Begin VB.CommandButton cmd13thMonth 
            Caption         =   "Excel File"
            Height          =   330
            Index           =   3
            Left            =   3840
            TabIndex        =   40
            Top             =   1920
            Width           =   1050
         End
         Begin VB.CommandButton cmd13thMonth 
            Caption         =   "Print Payslip"
            Height          =   330
            Index           =   2
            Left            =   2760
            TabIndex        =   39
            Top             =   1920
            Width           =   1050
         End
         Begin VB.CommandButton cmd13thMonth 
            Caption         =   "Retrieve"
            Height          =   330
            Index           =   1
            Left            =   1680
            TabIndex        =   38
            Top             =   1920
            Width           =   1050
         End
         Begin VB.CommandButton cmd13thMonth 
            Caption         =   "Generate"
            Height          =   330
            Index           =   0
            Left            =   240
            TabIndex        =   31
            Top             =   1920
            Visible         =   0   'False
            Width           =   1050
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1215
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   2143
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   9
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Surname"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Firstname"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "MIddlename"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "13th Month"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Year"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Width           =   9
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Width           =   9
            EndProperty
         End
         Begin VB.Label Label4 
            Caption         =   "Edit 13th Month Pay:"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Edit 13th Month Pay:"
            Height          =   255
            Left            =   3480
            TabIndex        =   44
            Top             =   1560
            Width           =   1575
         End
      End
      Begin VB.Label Label2 
         Height          =   240
         Left            =   4230
         TabIndex        =   17
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
Dim rsOtherDed As ADODB.Recordset
Dim rsOtherDed2 As ADODB.Recordset
Dim rsTaxTable As ADODB.Recordset
Dim rs13thMonth As ADODB.Recordset

Dim whereCon As String
Dim ReportScope As Integer
Dim ReportType As Integer
Dim ReportHeader As String
Dim mSapi As String
Dim mShortName As String
Dim mdedname(1, 50) As String

Dim ScopeRep As String
Dim mClientCode As Integer
Dim TotalLoan As Double
Dim LoanStart As Integer

'ADDED BY Bernard 3/27/2007
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

Dim Flag1601 As Boolean

Private Sub chkCal_Click()
If chkCal.VALUE = 1 Then
    chkLoans.VALUE = 0
    If ReportType = 1 Then ' SSS
        ReportHeader = "Calamity Loan Report"
    End If
Else
    If ReportType = 1 Then ' SSS
        ReportHeader = "SSS Contributions Report"
    End If
End If
End Sub

Private Sub chkLoans_Click()
If chkLoans.VALUE = 1 Then
    chkCal.VALUE = 0
    If ReportType = 3 Then
        ReportHeader = "HDMF Contributions Report"
    ElseIf ReportType = 1 Then
        ReportHeader = "SSS Loan Report"
    End If
Else
    If ReportType = 1 Then ' SSS
        ReportHeader = "SSS Contributions Report"
    End If
End If
End Sub

'Private Sub PreviewBIRForm(Optional ReportName As String = "DTR Records.rpt")
'
'Dim rsReport As New ADODB.Recordset
'Dim cApp As CRAXDRT.Application
'Dim cAppl As CRAXDRT.Application
'Dim cReport As CRAXDRT.Report
'Dim StartDate As String
'Dim EndDate As String
'Dim strPay As String
'
'On Error GoTo errHandler
'
'Me.MousePointer = vbHourglass
'
'If FMSFlag = False Then
'    'Get DTR records.
'    If chkAllDTR.VALUE = 1 Then 'all
'        strPay = "EXEC GetDTRRecords 1, " & txtClientCode & ", '" & mDTRKEY & "', '" & Trim$(txtClient(9).Text) & "'"
'    Else
'        strPay = "EXEC GetDTRRecords 0, " & txtClientCode & ", '" & mDTRKEY & "', '" & Trim$(txtClient(9).Text) & "'"
'    End If
'    'ConnectToSQL
'Else
'    'Get DTR records for FMS.
'    If chkAllDTR.VALUE = 1 Then 'all
'        strPay = "EXEC GetDTRRecordsCon 1, " & txtClientCode & ", '" & mDTRKEY & "', '" & Trim$(txtClient(9).Text) & "'"
'    Else
'        strPay = "EXEC GetDTRRecordsCon 0, " & txtClientCode & ", '" & mDTRKEY & "', '" & Trim$(txtClient(9).Text) & "'"
'    End If
'End If
'Call Action_Query(strPay, cnnstr)
'
'If FMSFlag = False Then
'    strPay = "SELECT * FROM tmpDTR order by Fullname, ActualDate "
'Else
'    strPay = "SELECT * FROM tmpDTRCon order by Fullname, ActualDate "
'    ReportName = "DTR Records Conductors.rpt"
'End If
''Set rsReport = ReturnRst(strPay, cnnstr, mstatic)
'Set rsReport = ReturnRst(strPay, cnnstr, mstatic)
'
'If rsReport.RecordCount = 0 Then
'    MsgBox "No record/s selected! Cannot proceed.", vbExclamation + vbOKOnly, PayrollAlert
'    rsReport.Close
'    Set rsReport = Nothing
'    Me.MousePointer = vbDefault
'    Exit Sub
'End If
'
'    Set cAppl = New CRAXDRT.Application
'
'    'Set cReport = cAppl.OpenReport(App.Path & "\Reports\Payroll Register.rpt")
'    Set cReport = cAppl.OpenReport(App.Path & "\Reports\" & ReportName)
'
'    cReport.DiscardSavedData
'    cReport.Database.SetDataSource rsReport
'
'    frmReportViewer.CRViewer91.ReportSource = cReport
'    frmReportViewer.CRViewer91.ViewReport
'
''    frmReportViewer.CRViewer91.Height = Me.Height - 150
''    frmReportViewer.CRViewer91.Width = Me.Width - 150
'
'    frmReportViewer.CRViewer91.Zoom 100
'
'    Do While frmReportViewer.CRViewer91.IsBusy
'        DoEvents
'    Loop
'
'    frmReportViewer.Show
'
'Set cAppl = Nothing
'Set cReport = Nothing
'rsReport.Close
'Set rsReport = Nothing
''CON.Close
'
'Me.MousePointer = vbDefault
'
''Call frmTimeKeep.SetFocus
'
'Exit Sub
'errHandler:
'
'    Me.MousePointer = vbDefault
'
'    MsgBox "Error : " & Err.Number & ";" & Err.Description, vbExclamation + vbOKOnly, "Error"
'    Err.Clear
'
'Set cAppl = Nothing
'Set cReport = Nothing
'
'End Sub

'Dim LocalClient As String

Private Sub cmbQuarter_Click()
    ScopeRep = "For the " & cmbQuarter.Text & " of " & cmbYearof.Text
End Sub

Private Sub cmbYearof_Click()
    ScopeRep = "For the Year " & cmbYearof.Text
End Sub

Private Sub Display13thMonth(ByRef mLsEmp As ListView, _
                        ByRef mRs As Recordset, Optional AppendSw As Boolean = True)
Dim litem1 As ListItem
Dim Listctr As Integer

Listctr = 0
ListView1(0).View = lvwReport
mLsEmp.ListItems.Clear
mRs.MoveFirst
While Not (mRs.EOF Or mRs.BOF)
    Set litem1 = ListView1(0).ListItems.Add(, , mRs!EmployeeID)
    litem1.SubItems(1) = IIf(IsNull(mRs!Surname), "Blank", mRs!Surname)
    litem1.SubItems(2) = IIf(IsNull(mRs!FirstName), "Blank", mRs!FirstName)
    litem1.SubItems(3) = IIf(IsNull(mRs!MiddleName), "Blank", mRs!MiddleName)
    litem1.SubItems(4) = Format(mRs!Month13Payment, "0.##")
    litem1.SubItems(5) = mRs!YearOf
    litem1.SubItems(6) = IIf(IsNull(mRs!Surname), "Blank", mRs!EmployeeID)
    Listctr = Listctr + 1
    mRs.MoveNext
Wend
Label4.Caption = "Total Emp. " & Listctr
End Sub

Private Sub cmd13thMonth_Click(Index As Integer)
Dim mYear As Integer
Dim res

mYear = Format(dtEndDate.VALUE, "yyyy")
Select Case Index
    Case 0 ' Generate
        Call ProcessReports
    Case 1 ' Retrieve 13th month records
        If cmbClient.Text = "--ALL CLIENTS--" Then ' if all clients, do not proceed
            StrQuery = "SELECT A.*, B.DateHired FROM [HR13thMonthPay] A LEFT JOIN [HREmployees] B ON A.EmployeeID = B.EmployeeID Where yearof = " & mYear & " ORDER BY A.Surname, A.Firstname"
        Else
            StrQuery = "SELECT A.*, B.DateHired FROM [HR13thMonthPay] A LEFT JOIN [HREmployees] B ON A.EmployeeID = B.EmployeeID Where yearof = " & mYear & " And A.clientCode = " & mClientCode & " ORDER BY A.Surname, A.Firstname"
        End If
        Set rsDeductions = ReturnRst(StrQuery, cnnstr, mstatic)
        If rsDeductions.RecordCount = 0 Then
            ListView1(0).ListItems.Clear
            MsgBox "No records selected!", vbInformation, PayrollAlert
        Else
            Call Display13thMonth(ListView1(0), rsDeductions)
        End If
    Case 2 ' Print Payslip
        Call Print13thMonthPayslip
    Case 3 ' Save to Excel File
        rsDeductions.MoveFirst
        Call GenerateReports(rsDeductions)
    Case 4 ' Save
        StrQuery = "UPDATE HR13thMonthPay SET month13payment = " & txtChange.Text & _
            " WHERE yearof = " & mYear & " And clientCode = " & mClientCode & " and employeeid = " & Val(txtChangeID.Text)
        
        res = MsgBox("Change 13th month pay to " & txtChange.Text & "?", vbQuestion + vbYesNo, PayrollAlert)
        If res = vbYes Then ' want to delete and reprocess
            Call Action_Query(StrQuery, cnnstr)
            If cmbClient.Text = "--ALL CLIENTS--" Then ' if all clients, do not proceed
                StrQuery = "SELECT * FROM [HR13thMonthPay] Where yearof = " & mYear & " ORDER BY Surname, Firstname"
            Else
                StrQuery = "SELECT * FROM [HR13thMonthPay] Where yearof = " & mYear & " And clientCode = " & mClientCode & " ORDER BY Surname, Firstname"
            End If
            Set rsDeductions = ReturnRst(StrQuery, cnnstr, mstatic)
            If rsDeductions.RecordCount = 0 Then
                MsgBox "No records selected!", vbInformation, PayrollAlert
            Else
                Call Display13thMonth(ListView1(0), rsDeductions)
            End If
        End If
    Case 5 ' Return
        Frame5.Visible = False
        FraReport.Visible = True
        Frame1.Visible = True
        Frame2.Visible = True
        Employees.Visible = True
        Frame3.Visible = True
        Frame4.Visible = True
End Select

End Sub

Private Sub cmdProcess_Click()
If InStr(1, ReportAccess, LoginName) > 0 Then
    ReportType = 7
    Call ProcessReports
ElseIf ForPayroll(AccountType) = False Then
'    If OptReport(0).VALUE = True And OptScope(2).VALUE = True Then
'        Call ProcessReports
'    Else
        MsgBox "You have no rights to generate this report!", vbInformation, PayrollAlert
'    End If
ElseIf (OptReport(0).VALUE = True Or OptReport(7).VALUE = True) And AdminRights = False Then
    MsgBox "You do have rights to process this report!", vbInformation, PayrollAlert
Else
    Call ProcessReports
End If

End Sub

Private Sub Generate13thMonthData()
Dim mYear As Integer
Dim rsYear As ADODB.Recordset
Dim strQuery2 As String
Dim res

'FraReport.Visible = False
'Frame1.Visible = False
'Frame2.Visible = False
'Employees.Visible = False
'Frame3.Visible = False
'Frame4.Visible = False
'Frame5.Visible = True
'
'Frame5.Caption = "13th Month for " & Format(dtStartDate.VALUE, "mmmm") & " " & Format(dtStartDate.VALUE, "yyyy") & " to " & Format(dtEndDate.VALUE, "mmmm") & " " & Format(dtEndDate.VALUE, "yyyy")

ScopeRep = "For Month Range " & Format(dtStartDate.VALUE, "mmmm") & " " & Format(dtStartDate.VALUE, "yyyy") & " to " & Format(dtEndDate.VALUE, "mmmm") & " " & Format(dtEndDate.VALUE, "yyyy")
' Update with specified date range.
StrQuery = "UPDATE tmpPayReport SET StartPayDate = '" & Format(dtStartDate.VALUE, "mm/dd/yyyy") & "', EndPayDate = '" & Format(dtEndDate.VALUE, "mm/dd/yyyy") & "' " & _
    "FROM tmpPayReport  " & _
    "WHERE LTRIM(RTRIM(tmpPayReport.UserID)) = '" & mUserID & "'"
Call Action_Query(StrQuery, cnnstr)

'StrQuery = "SELECT tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, MAX(HREmployees.MonthlyRate) as MonthlyRate, MAX(tmppayreport.ClientName) AS ClientName, MAX(StartPayDate) AS StartPayDate, MAX(EndPayDate) AS EndPayDate, " & _
'        "(CONVERT(Decimal(9,2), MAX(HREmployees.MonthlyRate)) / 12 ) * (CONVERT(Decimal(9,2), COUNT(tmpPayReport.MonthNo))/2) AS Month13Payment, (CONVERT(Decimal(9,2), COUNT(tmpPayReport.MonthNo))/2) AS NoofMonths, tmpPayReport.ClientCode, MAX(tmpPayReport.MonthOf) AS MonthoF, " & _
'        "MAX(HREmployees.AcctNo) AS AccountNo, YEAR(MAX(EndPayDate)) AS YearOf, MAX(HREmployees.DateHired) AS DateHired " & _
'        "FROM tmpPayReport INNER JOIN HREmployees ON (tmpPayReport.EmployeeID = HREmployees.EmployeeID) " & _
'        "WHERE LTRIM(RTRIM(tmpPayReport.UserID)) = '" & Trim$(mUserID) & "' " & _
'        "GROUP BY tmppayreport.ClientName, tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, tmpPayReport.ClientCode " & _
'        "ORDER BY tmppayreport.ClientName, tmppayreport.Surname, tmppayreport.FirstName"

'Updated Nov. 1, 2016, Compute 13th month pay based from Basic Pay.
StrQuery = "SELECT tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, MAX(HREmployees.MonthlyRate) as MonthlyRate, MAX(tmppayreport.ClientName) AS ClientName, MAX(StartPayDate) AS StartPayDate, MAX(EndPayDate) AS EndPayDate, " & _
        "CONVERT(Decimal(9,2), ((SUM(TmpPayReport.BasicPay) + SUM(TmpPayReport.OTPay) + SUM(TmpPayReport.Month13Payment)) - (SUM(TmpPayReport.AbsentDeduction) + SUM(TmpPayReport.TardinessDeduction))) / (CONVERT(Decimal(9,2), COUNT(tmpPayReport.MonthNo))/2) ) AS Month13Payment, " & _
        "(CONVERT(Decimal(9,2), COUNT(tmpPayReport.MonthNo))/2) AS NoofMonths, tmpPayReport.ClientCode, MAX(tmpPayReport.MonthOf) AS MonthoF, " & _
        "MAX(HREmployees.AcctNo) AS AccountNo, YEAR(MAX(EndPayDate)) AS YearOf, MAX(HREmployees.DateHired) AS DateHired " & _
        "FROM tmpPayReport INNER JOIN HREmployees ON (tmpPayReport.EmployeeID = HREmployees.EmployeeID) " & _
        "WHERE LTRIM(RTRIM(tmpPayReport.UserID)) = '" & Trim$(mUserID) & "' " & _
        "GROUP BY tmppayreport.ClientName, tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, tmpPayReport.ClientCode " & _
        "ORDER BY tmppayreport.ClientName, tmppayreport.Surname, tmppayreport.FirstName"

Set rsDeductions = ReturnRst(StrQuery, cnnstr, mstatic)
If rsDeductions.RecordCount = 0 Then
    MsgBox "No records gathered from master database. Cannot proceed!", vbInformation, PayrollAlert
'    Frame5.Visible = False
'    FraReport.Visible = True
'    Frame1.Visible = True
'    Frame2.Visible = True
'    Employees.Visible = True
'    Frame3.Visible = True
'    Frame4.Visible = True
    GoTo ExitRtn1
End If

mYear = Format(dtEndDate.VALUE, "yyyy")
If cmbClient.Text = "--ALL CLIENTS--" Then ' if all clients, do not proceed
    StrQuery = "SELECT A.*, B.DateHired FROM [HR13thMonthPay] A LEFT JOIN HREmployees B ON A.EmployeeID = B.EmployeeID Where yearof = " & mYear & " ORDER BY A.Surname, A.Firstname"
    strQuery2 = "DELETE FROM [HR13thMonthPay] Where yearof = " & mYear
Else
    StrQuery = "SELECT A.*, B.DateHired FROM [HR13thMonthPay] A LEFT JOIN HREmployees B ON A.EmployeeID = B.EmployeeID Where yearof = " & mYear & " And A.clientCode = " & mClientCode & " ORDER BY A.Surname, A.Firstname"
    strQuery2 = "DELETE FROM [HR13thMonthPay] Where yearof = " & mYear & " And clientCode = " & mClientCode
End If

Set rsYear = ReturnRst(StrQuery, cnnstr, mstatic)
If rsYear.RecordCount <> 0 Then
    res = MsgBox("13th month pay for " & cmbClient.Text & " already processed." & Chr(13) & "Do you want to delete and reprocess again?", vbQuestion + vbYesNo, PayrollAlert)
    If res = vbYes Then ' want to delete and reprocess
        Action_Query strQuery2, cnnstr
    Else
    Set rsDeductions = ReturnRst(StrQuery, cnnstr, mstatic)
    If rsDeductions.RecordCount <> 0 Then
        Call Display13thMonth(ListView1(0), rsDeductions)
    Else
        MsgBox "No records generated!", vbInformation, PayrollAlert
    End If
    GoTo ExitRtn1
    End If
End If
Set rsYear = Nothing

'    [EmployeeID] [int] NULL ,
'    [EmpNo] [nvarchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [FirstName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [Middlename] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [Surname] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [MonthlyRate] [decimal](9, 2) NULL ,
'    [ClientName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [StartPayDate] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [EndPayDate] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [Month13Payment] [decimal](27, 12) NULL ,
'    [NoofMonths] [decimal](13, 6) NULL ,
'    [ClientCode] [int] NULL ,
'    [MonthoF] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [YearOf] [int] NULL ,
'    [AccountNo] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL
        
'strQuery2 = "INSERT INTO HR13thMonthPay (EmployeeID, EmpNo, Firstname, Middlename, Surname, MonthlyRate, CLientName, StartPayDate, " & _
'        "EndPayDate, Month13Payment, NoofMonths, ClientCode, MonthOF, AccountNo, Yearof) " & _
'        "SELECT tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, MAX(HREmployees.MonthlyRate) as MonthlyRate, MAX(tmppayreport.ClientName) AS ClientName, MAX(StartPayDate) AS StartPayDate, MAX(EndPayDate) AS EndPayDate, " & _
'        "CONVERT(Decimal(9,2), (MAX(GrossPay) + MAX(OtherIncome) + MAX(Otpay) + MAX(Month13payment) ) -( MAX(Absentdeduction) + MAX(TardinessDeduction)) / 12 ) AS Month13Payment, (CONVERT(Decimal(9,2), COUNT(tmpPayReport.MonthNo))/2) AS NoofMonths, tmpPayReport.ClientCode, MAX(tmpPayReport.MonthOf) AS MonthoF, " & _
'        "MAX(HREmployees.AcctNo) AS AccountNo, YEAR(MAX(EndPayDate)) AS YearOf " & _
'        "FROM tmpPayReport INNER JOIN HREmployees ON (tmpPayReport.EmployeeID = HREmployees.EmployeeID) " & _
'        "WHERE LTRIM(RTRIM(tmpPayReport.UserID)) = '" & mUserID & "' " & _
'        "GROUP BY tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, tmpPayReport.ClientCode "

'Updated Nov. 1, 2016, get the 13th month pay based from basic pay.
strQuery2 = "INSERT INTO HR13thMonthPay (EmployeeID, EmpNo, Firstname, Middlename, Surname, MonthlyRate, CLientName, StartPayDate, " & _
        "EndPayDate, Month13Payment, NoofMonths, ClientCode, MonthOF, AccountNo, Yearof) " & _
        "SELECT tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, MAX(HREmployees.MonthlyRate) as MonthlyRate, MAX(tmppayreport.ClientName) AS ClientName, MAX(StartPayDate) AS StartPayDate, MAX(EndPayDate) AS EndPayDate, " & _
        "CONVERT(Decimal(9,2), ((SUM(TmpPayReport.BasicPay) + SUM(TmpPayReport.OTPay) + SUM(TmpPayReport.Month13Payment)) - (SUM(TmpPayReport.AbsentDeduction) + SUM(TmpPayReport.TardinessDeduction))) / 12 ) AS Month13Payment, " & _
        "(CONVERT(Decimal(9,2), COUNT(tmpPayReport.MonthNo))/2) AS NoofMonths, tmpPayReport.ClientCode, MAX(tmpPayReport.MonthOf) AS MonthoF, " & _
        "MAX(HREmployees.AcctNo) AS AccountNo, YEAR(MAX(EndPayDate)) AS YearOf " & _
        "FROM tmpPayReport INNER JOIN HREmployees ON (tmpPayReport.EmployeeID = HREmployees.EmployeeID) " & _
        "WHERE LTRIM(RTRIM(tmpPayReport.UserID)) = '" & mUserID & "' " & _
        "GROUP BY tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, tmpPayReport.ClientCode"
Action_Query strQuery2, cnnstr

' Orig code        "(CONVERT(Decimal(9,2), MAX(HREmployees.MonthlyRate)) / 12 ) * (CONVERT(Decimal(9,2), COUNT(tmpPayReport.MonthNo))/2) AS Month13Payment, (CONVERT(Decimal(9,2), COUNT(tmpPayReport.MonthNo))/2) AS NoofMonths, tmpPayReport.ClientCode, MAX(tmpPayReport.MonthOf) AS MonthoF, "
'        "CONVERT(Decimal(9,2), ((SUM(TmpPayReport.BasicPay) + SUM(TmpPayReport.OTPay) + SUM(TmpPayReport.Month13Payment)) - (SUM(TmpPayReport.AbsentDeduction) + SUM(TmpPayReport.TardinessDeduction))) / (CONVERT(Decimal(9,2), COUNT(tmpPayReport.MonthNo))/2) ) AS Month13Payment, "

Set rsDeductions = ReturnRst(StrQuery, cnnstr, mstatic)
If rsDeductions.RecordCount <> 0 Then
    Call Display13thMonth(ListView1(0), rsDeductions)
Else
    MsgBox "No records generated!", vbInformation, PayrollAlert
End If

ExitRtn1:

End Sub

Private Sub Print13thMonthPayslip()
Dim rsMonth As Recordset
Dim mYear As Integer

mYear = Format(dtEndDate.VALUE, "yyyy")
StrQuery = "EXEC Generate13thMonthPayslip3 '" & Trim$(mUserID) & "', " & mClientCode & ", '" & Format(dtStartDate.VALUE, "mm/dd/yyyy") & "', '" & Format(dtEndDate.VALUE, "mm/dd/yyyy") & "', " & mYear
Call Action_Query(StrQuery, cnnstr)
StrQuery = "SELECT * FROM tmpPayrollTemplate where EmployeeID in (select EmployeeID from hremployees where EmploymentStatus in ('R', 'C', 'P')) ORDER BY EmployeeName"
Set rsMonth = ReturnRst(StrQuery, cnnstr, mstatic)
        
If rsMonth.EOF = True Then
    GoTo ExitRtn1
End If

With rptPayslipTAS13th  ' with 13th month payslip
    Set .DataSource = rsMonth
        .DataMember = rsMonth.DataMember
        DoEvents
        .Show vbModal
End With

ExitRtn1:

Set rsMonth = Nothing

End Sub

Private Function CrystalReportConnection() As Boolean

On Error GoTo errcon

CrystalReportConnection = False

Set CON = New ADODB.Connection
CON.CursorLocation = adUseClient
CON.Provider = "SQLOLEDB"
CON.Open "File name=" & App.Path & "\ReportConnection.udl"
'CON.Open ReportConnection

ReportConnection = CON.ConnectionString
CrystalReportConnection = True

Exit Function

errcon:

MsgBox "Unable to to establish the connection! Please check connection settings.", vbCritical, PayrollAlert
CON.Close
Err.Clear

'ShellExecute hWnd, "open", App.Path & "\connection.udl", vbNullString, vbNullString, conSwNormal

End Function

Private Sub PreviewBIRForm(ByVal ReportName As String, Optional F1601 As Boolean = True)
Dim rsReport As New ADODB.Recordset
Dim cApp As CRAXDRT.Application
Dim cAppl As CRAXDRT.Application
Dim cReport As CRAXDRT.Report

On Error GoTo errHandler
    
If CrystalReportConnection = False Then
    Exit Sub
End If

If rsTaxTable.RecordCount = 0 Then
    MsgBox "No record/s selected! Cannot proceed.", vbExclamation + vbOKOnly, PayrollAlert
    Exit Sub
End If

'Unload Me
If F1601 = True Then
    frmReportViewer.Caption = "BIR Form 1601-C Preview"
Else
    frmReportViewer.Caption = "BIR Form 2316 Preview"
End If
Me.MousePointer = vbHourglass

Set cAppl = New CRAXDRT.Application

'Set cReport = cAppl.OpenReport(App.Path & "\Reports\Payroll Register.rpt")
Set cReport = cAppl.OpenReport(App.Path & "\Reports\" & ReportName)

cReport.DiscardSavedData
cReport.Database.SetDataSource rsTaxTable

frmReportViewer.CRViewer91.ReportSource = cReport
frmReportViewer.CRViewer91.ViewReport

'    frmReportViewer.CRViewer91.Height = Me.Height - 150
'    frmReportViewer.CRViewer91.Width = Me.Width - 150

frmReportViewer.CRViewer91.Zoom 100

Do While frmReportViewer.CRViewer91.IsBusy
    DoEvents
Loop

frmReportViewer.Show

Set cAppl = Nothing
Set cReport = Nothing

CON.Close
Me.MousePointer = vbDefault
'frmReportViewer.Caption = "Payroll Summary Report"

Exit Sub

errHandler:
    
    Me.MousePointer = vbDefault
    
    MsgBox "Error : " & Err.Number & ";" & Err.Description, vbExclamation + vbOKOnly, "Error"
    Err.Clear
    
    Set cAppl = Nothing
    Set cReport = Nothing

End Sub

Private Sub GetLatesFromDTR() ' For DTR based Lates Report
                
whereCon = Replace(whereCon, "payrollmaster.startpaydate", "b.ActualDate")
whereCon = Replace(whereCon, "PayrollMaster.clientcode", "a.clientcode")

'StrQuery = "INSERT INTO tmpPayReport " & _
'        "(EmployeeID, EmpNo, FirstName, Middlename, Surname, ClientCode, " & _
'        "StartPayDate, HoursTardiness, MonthOf) " & _
'        "SELECT EmployeeID, EmpNo, FirstName, Middlename, Surname, clientcode, " & _
'        "ActualDate, DTRLates, NameOfDay " & _
'        "FROM vwDTRLates " & _
'        "WHERE " & whereCon

StrQuery = "INSERT INTO tmpPayReport " & _
        "(EmployeeID, EmpNo, FirstName, Middlename, Surname, ClientCode, ClientName, " & _
        "StartPayDate, HoursTardiness, MonthOf, UserID) " & _
        "SELECT a.EmployeeID, a.EmpNo, a.FirstName, a.Middlename, a.Surname, a.ClientCode, '" & cmbClient.Text & "', " & _
        "CONVERT(varchar(10), b.ActualDate, 101) , b.dtrlates, b.NameOfDay, '" & LTrim(mUserID) & "' " & _
        "FROM dtrpayrollrecords b left join hremployees a " & _
        "ON a.dtrid = b.dtrid WHERE dtrlates <> 0 " & _
        "AND " & whereCon

Call Action_Query(StrQuery, cnnstr)
StrQuery = "SELECT * FROM tmppayreport " & _
    "WHERE LTRIM(RTRIM(userID)) = '" & mUserID & "' ORDER BY Surname, Firstname, StartPaydate"

End Sub

Private Sub ProcessReports(Optional FormSw As Boolean = False)
Dim mQuery As String
Dim mDayNo As Integer
Dim res

'Check if for Form 1601. Updated May 14, 2013.
If Flag1601 = True And OptScope(1).VALUE = True And OptReport(9).VALUE = True Then
    PreviewBIRForm ("Form 1601-C.rpt")
    GoTo ExitRtn1
ElseIf FormSw = False And Flag1601 = False And OptScope(1).VALUE = True And OptReport(9).VALUE = True Then
    MsgBox "Compute for Totals first before you generate the report.", vbInformation, PayrollAlert
    GoTo ExitRtn1
End If

Me.MousePointer = vbHourglass
If OptReport(0).VALUE = False Then
    Call GetReportScope
    mQuery = "PayrollMaster.PagibigLoan"
    If chkLoans.VALUE = 1 Then
        If ReportType = 1 Then      ' for SSS Loan
            whereCon = whereCon & " and sssloan <> 0"
        ElseIf ReportType = 3 Then  ' for Pagibig Loan
            whereCon = whereCon & " and pagibigloan <> 0 "
        End If
    ElseIf chkCal.VALUE = 1 Then    ' Check for calamity Loan, updated Oct. 12, 2013
        If ReportType = 1 Then      ' for SSS Calamity Loan
            whereCon = whereCon & " and medical <> 0"
        ElseIf ReportType = 3 Then  ' for Pagibig Loan
            whereCon = whereCon & " and ExcessLGLholotpay <> 0 "
        End If
        mQuery = "PayrollMaster.ExcessLGLholotpay"
    End If
    mUserID = Trim(UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID))))
        
    If ReportScope <> 0 Then
        StrQuery = "DELETE FROM tmpPayReport WHERE LTRIM(RTRIM(useriD)) like '%" & mUserID & "%'"
        Call Action_Query(StrQuery, cnnstr)
        
        ' Need to get the sickleave (OTPay field) and vacation leave (month13Payment) payments
        If ReportType = 5 Then      ' for 13th month pay
            StrQuery = "INSERT INTO tmpPayReport " & _
                        "(EmployeeID, EmpNo, FirstName, Middlename, Surname, SSS, Philhealth, TaxCode, Exempt, " & _
                        "Pagibig, TIN, BasicPay, GrossPay, OTHours, OTPay, ClientCode, Month13Payment, TASAesla, LMCDed, ValewithClient, Medical, " & _
                        "SalaryLoan, Uniform, Variances, Vale_IOU, BirthDate, DateHired, " & _
                        "DaysAbsent, AbsentDeduction, HoursTardiness, TardinessDeduction, YearOf, " & _
                        "UnionDues, HazardPay, CashAdvance, WTaxAmount, MiscFee, CalamityLoan, SSSEmployee, SSSEmployer, ECCEmployer, PHealthEmployee, PHealthEmployer, " & _
                        "PagibigEmployee, PagibigEmployer, SSSLoan, PagibigLoan, OtherLoan, StartPayDate, EndPayDate, MonthNo, CutOffNo, ClientName, MonthOf, UserID) " & _
                        "SELECT PayrollMaster.EmployeeID, PayrollMaster.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, HREmployees.TaxCode, TaxExemptionCode.ExemptionAmount, " & _
                        "HREmployees.Pagibig, HREmployees.TIN, PayrollMaster.BasicPay, PayrollMaster.GrossPay, PayrollMaster.OTHours, PayrollMaster.SickLeavePay, PayrollMaster.ClientCode, PayrollMaster.VacationLeavePay, PayrollMaster.TASAesla, PayrollMaster.LMCDed, PayrollMaster.ValewithClient, PayrollMaster.Medical, " & _
                        "PayrollMaster.SalaryLoan, PayrollMaster.Uniform, PayrollMaster.Variances, PayrollMaster.Vale_IOU, HREmployees.BirthDate, HREmployees.DateHired, " & _
                        "PayrollMaster.DaysAbsent, PayrollMaster.AbsentDeduction, PayrollMaster.HoursTardiness, PayrollMaster.TardinessDeduction, YEAR(PayrollMaster.StartPayDate), " & _
                        "PayrollMaster.UnionDues, PayrollMaster.HazardPay, PayrollMaster.CashAdvance, PayrollMaster.WTaxAmount, PayrollMaster.MiscFee, PayrollMaster.CalamityLoan, PayrollMaster.SSSEmployee, PayrollMaster.SSSEmployer, PayrollMaster.ECCEmployer, PayrollMaster.PHealthEmployee, PayrollMaster.PHealthEmployer, " & _
                        "PayrollMaster.PagibigEmployee, PayrollMaster.PagibigEmployer, PayrollMaster.SSSLoan, " & mQuery & ", PayrollMaster.OtherLoan, PayrollMaster.StartPayDate, PayrollMaster.EndPayDate, PayrollMaster.MonthNo, PayrollMaster.CutOffNo, ClientTable.ClientName, DATENAME(month, PayrollMaster.StartPayDate), '" & LTrim(mUserID) & "' " & _
                        "FROM (PayrollMaster INNER JOIN HREmployees ON (PayrollMaster.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollMaster.ClientCode = ClientTable.ClientCode LEFT JOIN TaxExemptionCode ON HREmployees.TaxCode = TaxExemptionCode.TaxExemptionCode " & _
                        "WHERE " & whereCon  '" ORDER BY PayrollMaster.ClientCode, PayrollMaster.StartPayDate, PayrollMaster.StartPayDate, HREmployees.Surname, HREmployees.FirstName"
            Call Action_Query(StrQuery, cnnstr)
            
            ' Get also the data from current payroll transactions.
            whereCon = Replace(whereCon, "payrollmaster", "PayrollTrans")
            whereCon = Replace(whereCon, "PayrollMaster", "PayrollTrans")
            mQuery = Replace(mQuery, "PayrollMaster", "PayrollTrans")
            StrQuery = "INSERT INTO tmpPayReport " & _
                        "(EmployeeID, EmpNo, FirstName, Middlename, Surname, SSS, Philhealth, TaxCode, Exempt, " & _
                        "Pagibig, TIN, BasicPay, GrossPay, OTHours, OTPay, ClientCode, Month13Payment, TASAesla, LMCDed, ValewithClient, Medical, " & _
                        "SalaryLoan, Uniform, Variances, Vale_IOU, BirthDate, DateHired, " & _
                        "DaysAbsent, AbsentDeduction, HoursTardiness, TardinessDeduction, YearOF, " & _
                        "UnionDues, HazardPay, CashAdvance, WTaxAmount, MiscFee, CalamityLoan, SSSEmployee, SSSEmployer, ECCEmployer, PHealthEmployee, PHealthEmployer, " & _
                        "PagibigEmployee, PagibigEmployer, SSSLoan, PagibigLoan, OtherLoan, StartPayDate, EndPayDate, MonthNo, CutOffNo, ClientName, MonthOf, UserID) " & _
                        "SELECT PayrollTrans.EmployeeID, PayrollTrans.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, HREmployees.TaxCode, TaxExemptionCode.ExemptionAmount, " & _
                        "HREmployees.Pagibig, HREmployees.TIN, PayrollTrans.BasicPay, PayrollTrans.GrossPay, PayrollTrans.OTHours, PayrollTrans.SickLeavePay, PayrollTrans.ClientCode, PayrollTrans.VacationLeavePay, PayrollTrans.TASAesla, PayrollTrans.LMCDed, PayrollTrans.ValewithClient, PayrollTrans.Medical, " & _
                        "PayrollTrans.SalaryLoan, PayrollTrans.Uniform, PayrollTrans.Variances, PayrollTrans.Vale_IOU, HREmployees.BirthDate, HREmployees.DateHired, " & _
                        "PayrollTrans.DaysAbsent, PayrollTrans.AbsentDeduction, PayrollTrans.HoursTardiness, PayrollTrans.TardinessDeduction, YEAR(PayrollTrans.StartPayDate), " & _
                        "PayrollTrans.UnionDues, PayrollTrans.HazardPay, PayrollTrans.CashAdvance, PayrollTrans.WTaxAmount, PayrollTrans.MiscFee, PayrollTrans.CalamityLoan, PayrollTrans.SSSEmployee, PayrollTrans.SSSEmployer, PayrollTrans.ECCEmployer, PayrollTrans.PHealthEmployee, PayrollTrans.PHealthEmployer, " & _
                        "PayrollTrans.PagibigEmployee, PayrollTrans.PagibigEmployer, PayrollTrans.SSSLoan, " & mQuery & ", PayrollTrans.OtherLoan, PayrollTrans.StartPayDate, PayrollTrans.EndPayDate, PayrollTrans.MonthNo, PayrollTrans.CutOffNo, ClientTable.ClientName, DATENAME(month, PayrollTrans.StartPayDate), '" & LTrim(mUserID) & "' " & _
                        "FROM (PayrollTrans INNER JOIN HREmployees ON ((PayrollTrans.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollTrans.ClientCode = ClientTable.ClientCode) LEFT JOIN TaxExemptionCode ON HREmployees.TaxCode = TaxExemptionCode.TaxExemptionCode " & _
                        "WHERE " & whereCon  '" ORDER BY PayrollTrans.ClientCode, PayrollTrans.StartPayDate, PayrollTrans.StartPayDate, HREmployees.Surname, HREmployees.FirstName"
            Call Action_Query(StrQuery, cnnstr)
        
        Else
            If ReportType = 6 Then  ' for DTR Based Lates 3/18/18
                res = MsgBox("Get Lates from DTR?", vbQuestion + vbYesNo, PayrollAlert)
            End If
            If ReportType = 6 And res = vbYes Then ' Dtr based
                Call GetLatesFromDTR
                ReportType = 66 ' change the type
                GoTo GetData1
            Else
                StrQuery = "INSERT INTO tmpPayReport " & _
                        "(EmployeeID, EmpNo, FirstName, Middlename, Surname, SSS, Philhealth, TaxCode, Exempt, " & _
                        "Pagibig, TIN, BasicPay, GrossPay, OTHours, OTPay, ClientCode, Month13Payment, TASAesla, LMCDed, ValewithClient, Medical, " & _
                        "SalaryLoan, Uniform, Variances, Vale_IOU, BirthDate, DateHired, " & _
                        "DaysAbsent, AbsentDeduction, HoursTardiness, TardinessDeduction, YearOf, " & _
                        "UnionDues, HazardPay, CashAdvance, WTaxAmount, MiscFee, CalamityLoan, SSSEmployee, SSSEmployer, ECCEmployer, PHealthEmployee, PHealthEmployer, " & _
                        "PagibigEmployee, PagibigEmployer, SSSLoan, PagibigLoan, OtherLoan, StartPayDate, EndPayDate, MonthNo, CutOffNo, ClientName, MonthOf, UserID) " & _
                        "SELECT PayrollMaster.EmployeeID, PayrollMaster.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, HREmployees.TaxCode, TaxExemptionCode.ExemptionAmount, " & _
                        "HREmployees.Pagibig, HREmployees.TIN, PayrollMaster.BasicPay, PayrollMaster.GrossPay, PayrollMaster.OTHours, PayrollMaster.OTPay, PayrollMaster.ClientCode, PayrollMaster.Month13Payment, PayrollMaster.TASAesla, PayrollMaster.LMCDed, PayrollMaster.ValewithClient, PayrollMaster.Medical, " & _
                        "PayrollMaster.SalaryLoan, PayrollMaster.Uniform, PayrollMaster.Variances, PayrollMaster.Vale_IOU, HREmployees.BirthDate, HREmployees.DateHired, " & _
                        "PayrollMaster.DaysAbsent, PayrollMaster.AbsentDeduction, PayrollMaster.HoursTardiness, PayrollMaster.TardinessDeduction, YEAR(PayrollMaster.StartPayDate), " & _
                        "PayrollMaster.UnionDues, PayrollMaster.HazardPay, PayrollMaster.CashAdvance, PayrollMaster.WTaxAmount, PayrollMaster.MiscFee, PayrollMaster.CalamityLoan, PayrollMaster.SSSEmployee, PayrollMaster.SSSEmployer, PayrollMaster.ECCEmployer, PayrollMaster.PHealthEmployee, PayrollMaster.PHealthEmployer, " & _
                        "PayrollMaster.PagibigEmployee, PayrollMaster.PagibigEmployer, PayrollMaster.SSSLoan, " & mQuery & ", PayrollMaster.OtherLoan, PayrollMaster.StartPayDate, PayrollMaster.EndPayDate, PayrollMaster.MonthNo, PayrollMaster.CutOffNo, ClientTable.ClientName, DATENAME(month, PayrollMaster.StartPayDate), '" & LTrim(mUserID) & "' " & _
                        "FROM (PayrollMaster INNER JOIN HREmployees ON (PayrollMaster.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollMaster.ClientCode = ClientTable.ClientCode LEFT JOIN TaxExemptionCode ON HREmployees.TaxCode = TaxExemptionCode.TaxExemptionCode " & _
                        "WHERE " & whereCon  '" ORDER BY PayrollMaster.ClientCode, PayrollMaster.StartPayDate, PayrollMaster.StartPayDate, HREmployees.Surname, HREmployees.FirstName"
                Call Action_Query(StrQuery, cnnstr)
            
                ' Get also the data from current payroll transactions.
                whereCon = Replace(whereCon, "payrollmaster", "PayrollTrans")
                whereCon = Replace(whereCon, "PayrollMaster", "PayrollTrans")
                mQuery = Replace(mQuery, "PayrollMaster", "PayrollTrans")
                StrQuery = "INSERT INTO tmpPayReport " & _
                        "(EmployeeID, EmpNo, FirstName, Middlename, Surname, SSS, Philhealth, TaxCode, Exempt, " & _
                        "Pagibig, TIN, BasicPay, GrossPay, OTHours, OTPay, ClientCode, Month13Payment, TASAesla, LMCDed, ValewithClient, Medical, " & _
                        "SalaryLoan, Uniform, Variances, Vale_IOU, BirthDate, DateHired, " & _
                        "DaysAbsent, AbsentDeduction, HoursTardiness, TardinessDeduction, YearOF, " & _
                        "UnionDues, HazardPay, CashAdvance, WTaxAmount, MiscFee, CalamityLoan, SSSEmployee, SSSEmployer, ECCEmployer, PHealthEmployee, PHealthEmployer, " & _
                        "PagibigEmployee, PagibigEmployer, SSSLoan, PagibigLoan, OtherLoan, StartPayDate, EndPayDate, MonthNo, CutOffNo, ClientName, MonthOf, UserID) " & _
                        "SELECT PayrollTrans.EmployeeID, PayrollTrans.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, HREmployees.TaxCode, TaxExemptionCode.ExemptionAmount, " & _
                        "HREmployees.Pagibig, HREmployees.TIN, PayrollTrans.BasicPay, PayrollTrans.GrossPay, PayrollTrans.OTHours, PayrollTrans.OTPay, PayrollTrans.ClientCode, PayrollTrans.Month13Payment, PayrollTrans.TASAesla, PayrollTrans.LMCDed, PayrollTrans.ValewithClient, PayrollTrans.Medical, " & _
                        "PayrollTrans.SalaryLoan, PayrollTrans.Uniform, PayrollTrans.Variances, PayrollTrans.Vale_IOU, HREmployees.BirthDate, HREmployees.DateHired, " & _
                        "PayrollTrans.DaysAbsent, PayrollTrans.AbsentDeduction, PayrollTrans.HoursTardiness, PayrollTrans.TardinessDeduction, YEAR(PayrollTrans.StartPayDate), " & _
                        "PayrollTrans.UnionDues, PayrollTrans.HazardPay, PayrollTrans.CashAdvance, PayrollTrans.WTaxAmount, PayrollTrans.MiscFee, PayrollTrans.CalamityLoan, PayrollTrans.SSSEmployee, PayrollTrans.SSSEmployer, PayrollTrans.ECCEmployer, PayrollTrans.PHealthEmployee, PayrollTrans.PHealthEmployer, " & _
                        "PayrollTrans.PagibigEmployee, PayrollTrans.PagibigEmployer, PayrollTrans.SSSLoan, " & mQuery & ", PayrollTrans.OtherLoan, PayrollTrans.StartPayDate, PayrollTrans.EndPayDate, PayrollTrans.MonthNo, PayrollTrans.CutOffNo, ClientTable.ClientName, DATENAME(month, PayrollTrans.StartPayDate), '" & LTrim(mUserID) & "' " & _
                        "FROM (PayrollTrans INNER JOIN HREmployees ON ((PayrollTrans.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollTrans.ClientCode = ClientTable.ClientCode) LEFT JOIN TaxExemptionCode ON HREmployees.TaxCode = TaxExemptionCode.TaxExemptionCode " & _
                        "WHERE " & whereCon  '" ORDER BY PayrollTrans.ClientCode, PayrollTrans.StartPayDate, PayrollTrans.StartPayDate, HREmployees.Surname, HREmployees.FirstName"
                Call Action_Query(StrQuery, cnnstr)
            End If
        End If
        
        If ReportType = 1 Then  ' for SSS Loan, Order by Client Code + Name
            StrQuery = "SELECT tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, min(tmpPayReport.SSS) as SSS, " & _
                    "MIN(tmpPayReport.BirthDate) as BirthDate, MIN( tmpPayReport.DateHired) as DateHired, MIN(tmpPayReport.TaxCode) as TaxCode, " & _
                    "MIN(tmpPayReport.Philhealth) as PhilHealth, min(tmpPayReport.Pagibig) as Pagibig, min(tmpPayReport.TIN) as TIN, tmpPayReport.ClientCode, tmppayreport.monthof, " & _
                    "SUM(tmpPayReport.BasicPay) as BasicPay, sum(tmpPayReport.GrossPay) as GrossPay, " & _
                    "SUM(tmpPayReport.Month13Payment) as month13payment, sum(tmpPayReport.TASAesla) as TASAesla, sum(tmpPayReport.LMCDed) as LMCDed, " & _
                    "SUM(tmpPayReport.ValeWithClient) as valewithclient, sum(tmpPayReport.CashAdvance) as CashAdvance, sum(tmpPayReport.WTaxAmount) as WTaxAmount, " & _
                    "SUM(tmpPayReport.Medical) as Medical, sum(tmpPayReport.SalaryLoan) as SalaryLoan, sum(tmpPayReport.Uniform) as Uniform, sum(tmpPayReport.Variances) as Variances, sum(tmpPayReport.Vale_IOU) as Vale_IOU, " & _
                    "SUM(tmpPayReport.UnionDues) as UnionDues, SUM(tmpPayReport.HazardPay) as HazardPay, SUM(tmpPayReport.MiscFee) as miscfee, sum(tmpPayReport.CalamityLoan) as CalamityLoan, sum(tmpPayReport.SSSEmployee) as SSSEmployee, " & _
                    "SUM(tmpPayReport.SSSEmployer) as SSSEmployer, SUM(tmpPayReport.ECCEmployer) as ECCEmployer, SUM(tmppayreport.sssloan) as sssloan, SUM(tmppayreport.pagibigloan) as pagibigloan, sum(tmpPayReport.PHealthEmployee) as phealthemployee, sum(tmpPayReport.PHealthEmployer) as PHealthEmployer, " & _
                    "SUM(tmpPayReport.PagibigEmployee) as pagibigemployee, sum(tmpPayReport.PagibigEmployer) as pagibigemployer, min(tmpPayReport.MonthNo) as MonthNo, " & _
                    "MIN(tmpPayReport.CutOffNo) as cutoffno, min(tmpPayReport.ClientName) as clientname, YearOF FROM tmpPayReport " & _
                    "WHERE LTRIM(RTRIM(userID)) = '" & mUserID & "' " & _
                    "GROUP BY tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, tmpPayReport.ClientCode, tmppayreport.YearOf, tmppayreport.monthof, tmppayreport.monthno " & _
                    "ORDER BY tmppayreport.ClientCode, tmppayreport.Surname, tmppayreport.FirstName, tmppayreport.YearOf, tmppayreport.monthno"
        Else
            If ReportType = 5 Then      ' for 13th month pay
                Call Generate13thMonthData
                GoTo ExitRtn1
            ElseIf ReportType = 7 Then  ' MasterList
                Call CreateMasterList
                GoTo ExitRtn1
            ElseIf ReportType = 9 Then  ' Alphalist Report
                ' Get Annual Tax Table.
                StrQuery = "SELECT salrate1, salrate2, basetax, perover FROM YearlyTaxTable order by id"
                Set rsTaxTable = ReturnRst(StrQuery, cnnstr, mstatic)
                ' Get 13th month pay.
                StrQuery = "SELECT EmployeeID, Month13Payment, YearOF FROM HR13thMonthPay WHERE yearof = " & Year(dtEndDate.VALUE)
                Set rs13thMonth = ReturnRst(StrQuery, cnnstr, mstatic)
                
                StrQuery = "SELECT tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, min(tmpPayReport.SSS) as SSS, min(tmpPayReport.TaxCode) as TaxCode, max(tmpPayReport.Exempt) as Exempt," & _
                    "MIN( tmpPayReport.BirthDate) as BirthDate, MIN( tmpPayReport.DateHired) as DateHired, MIN(tmpPayReport.Philhealth) as PhilHealth, min(tmpPayReport.Pagibig) as pagibig, min(tmpPayReport.TIN) as TIN, tmpPayReport.ClientCode, MIN(tmppayreport.monthof) AS Monthof, MIN(tmppayreport.Yearof) AS YearOf, " & _
                    "SUM(tmpPayReport.BasicPay) as BasicPay, SUM(tmpPayReport.GrossPay) as GrossPay, SUM(tmpPayReport.Month13Payment) as month13payment, sum(tmpPayReport.TASAesla) as TASAesla, sum(tmpPayReport.LMCDed) as LMCDed, " & _
                    "SUM(tmpPayReport.ValewithClient) as valewithclient, sum(tmpPayReport.CashAdvance) as CashAdvance, sum(tmpPayReport.WTaxAmount) as WTaxAmount, " & _
                    "SUM(tmpPayReport.DaysAbsent) as DaysAbsent, sum(tmpPayReport.AbsentDeduction) as AbsentDeduction, sum(tmpPayReport.HoursTardiness) as HoursTardiness, sum(tmpPayReport.TardinessDeduction) as TardinessDeduction, " & _
                    "SUM(tmpPayReport.Medical) as Medical, sum(tmpPayReport.SalaryLoan) as SalaryLoan, sum(tmpPayReport.Uniform) as Uniform, sum(tmpPayReport.Variances) as Variances, sum(tmpPayReport.Vale_IOU) as Vale_IOU, " & _
                    "SUM(tmpPayReport.UnionDues) as UnionDues, SUM(tmpPayReport.HazardPay) as HazardPay, SUM(tmpPayReport.MiscFee) as miscfee, sum(tmpPayReport.CalamityLoan) as CalamityLoan, sum(tmpPayReport.SSSEmployee) as SSSEmployee, " & _
                    "SUM(tmpPayReport.SSSEmployer) as sssemployer, SUM(tmpPayReport.ECCEmployer) as ECCEmployer, SUM(tmppayreport.sssloan) as sssloan, SUM(tmppayreport.pagibigloan) as pagibigloan, sum(tmpPayReport.PHealthEmployee) as phealthemployee, sum(tmpPayReport.PHealthEmployer) as PHealthEmployer, " & _
                    "SUM(tmpPayReport.OTHours) as OTHours, sum(tmpPayReport.OTPay) as OTPay, " & _
                    "SUM(tmpPayReport.PagibigEmployee) as pagibigemployee, sum(tmpPayReport.PagibigEmployer) as pagibigemployer, min(tmpPayReport.MonthNo) as MonthNo, " & _
                    "MIN(tmpPayReport.CutOffNo) as cutoffno, min(tmpPayReport.ClientName) as clientname FROM tmpPayReport " & _
                    "WHERE LTRIM(RTRIM(userID)) = '" & mUserID & "' " & _
                    "GROUP BY tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, tmpPayReport.ClientCode " & _
                    "ORDER BY tmppayreport.ClientCode, tmppayreport.Surname, tmppayreport.FirstName"
            Else  ' Updated Feb. 09, 2013 to include YearOF
'                If ReportType = 6 Then ' get number of days late, updated 12/2/2017 Bernard
'                    StrQuery = "TRUNCATE TABLE tmpDaysLate"
'                    Call Action_Query(StrQuery, cnnstr)
'                    whereCon = " DTRPayrollRecords.DTRLates <> 0.00 AND CONVERT(varchar(10), CONVERT(datetime, DTRPayrollRecords.ActualDate), 112) BETWEEN " & Val(Format(dtStartDate.VALUE, "yyyymmdd")) & " and " & Val(Format(dtEndDate.VALUE, "yyyymmdd"))
'                    StrQuery = "INSERT INTO tmpDaysLate SELECT MIN(a.DtrID) AS DTRID, COUNT(ActualDate) AS DaysLate, MIN(b.EmployeeID) AS EmpID " & _
'                        "FROM DTRPayrollRecords a LEFT JOIN HREmployees b ON a.DtrID = b.DtrID WHERE " & whereCon
'                    Call Action_Query(StrQuery, cnnstr)
'                    StrQuery = "UPDATE a SET cutoffno = b.DaysLate FROM tmpPayReport a WHERE a.EmployeeID = "
'
'select min(a.dtrid), count(actualdate) as dy, min(b.employeeid)
'from dtrpayrollrecords a left join hremployees b
'on a.dtrid = b.dtrid
'
'
'select top 100 * from dtrpayrollrecords
'where actualdate > '2017-10-01'
'and dtrid = '00002929'
'and dtrlates <> 0.00
'
'                End If
                
                StrQuery = "SELECT tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, min(tmpPayReport.SSS) as SSS, min(tmpPayReport.TaxCode) as TaxCode, max(tmpPayReport.Exempt) as Exempt," & _
                    "MIN( tmpPayReport.BirthDate) as BirthDate, MIN( tmpPayReport.DateHired) as DateHired, MIN(tmpPayReport.Philhealth) as PhilHealth, min(tmpPayReport.Pagibig) as pagibig, min(tmpPayReport.TIN) as TIN, tmpPayReport.ClientCode, tmppayreport.monthof, " & _
                    "SUM(tmpPayReport.BasicPay) as BasicPay, SUM(tmpPayReport.GrossPay) as GrossPay, SUM(tmpPayReport.Month13Payment) as month13payment, sum(tmpPayReport.TASAesla) as TASAesla, sum(tmpPayReport.LMCDed) as LMCDed, " & _
                    "SUM(tmpPayReport.ValewithClient) as valewithclient, sum(tmpPayReport.CashAdvance) as CashAdvance, sum(tmpPayReport.WTaxAmount) as WTaxAmount, " & _
                    "SUM(tmpPayReport.DaysAbsent) as DaysAbsent, sum(tmpPayReport.AbsentDeduction) as AbsentDeduction, sum(tmpPayReport.HoursTardiness) as HoursTardiness, sum(tmpPayReport.TardinessDeduction) as TardinessDeduction, " & _
                    "SUM(tmpPayReport.Medical) as Medical, sum(tmpPayReport.SalaryLoan) as SalaryLoan, sum(tmpPayReport.Uniform) as Uniform, sum(tmpPayReport.Variances) as Variances, sum(tmpPayReport.Vale_IOU) as Vale_IOU, " & _
                    "SUM(tmpPayReport.UnionDues) as UnionDues, SUM(tmpPayReport.HazardPay) as HazardPay, SUM(tmpPayReport.MiscFee) as miscfee, sum(tmpPayReport.CalamityLoan) as CalamityLoan, sum(tmpPayReport.SSSEmployee) as SSSEmployee, " & _
                    "SUM(tmpPayReport.SSSEmployer) as sssemployer, SUM(tmpPayReport.ECCEmployer) as ECCEmployer, SUM(tmppayreport.sssloan) as sssloan, SUM(tmppayreport.pagibigloan) as pagibigloan, sum(tmpPayReport.PHealthEmployee) as phealthemployee, sum(tmpPayReport.PHealthEmployer) as PHealthEmployer, " & _
                    "SUM(tmpPayReport.OTHours) as OTHours, sum(tmpPayReport.OTPay) as OTPay, " & _
                    "SUM(tmpPayReport.PagibigEmployee) as pagibigemployee, sum(tmpPayReport.PagibigEmployer) as pagibigemployer, min(tmpPayReport.MonthNo) as MonthNo, " & _
                    "MIN(tmpPayReport.CutOffNo) as cutoffno, min(tmpPayReport.ClientName) as clientname, tmpPayReport.YearOF FROM tmpPayReport " & _
                    "WHERE LTRIM(RTRIM(userID)) = '" & mUserID & "' " & _
                    "GROUP BY tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, tmpPayReport.ClientCode, tmppayreport.YearOF, tmppayreport.monthof, tmppayreport.monthno " & _
                    "ORDER BY tmppayreport.ClientCode, tmppayreport.YearOF, tmppayreport.monthno, tmppayreport.Surname, tmppayreport.FirstName"
                
                If ReportType = 8 Then ' for Deduction payments
                    mQuery = "SELECT * FROM vwOtherDed WHERE " & Replace(whereCon, "PayrollTrans.", "")
                    Set rsOtherDed = ReturnRst(mQuery, cnnstr, mstatic)
                    mQuery = "SELECT DISTINCT VoucherNo FROM vwOtherDed WHERE VoucherNo IS NOT NULL AND VoucherNo <> '' AND " & Replace(whereCon, "PayrollTrans.", "") & " ORDER BY VoucherNo"
                    Set rsOtherDed2 = ReturnRst(mQuery, cnnstr, mstatic)
                End If
            
            End If
        End If
    Else
        StrQuery = "SELECT PayrollMaster.EmployeeID, PayrollMaster.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, " & _
                "HREmployees.Pagibig, HREmployees.TIN, HREmployees.BirthDate, HREmployees.DateHired, PayrollMaster.ClientCode, PayrollMaster.Month13Payment, PayrollMaster.TASAesla, PayrollMaster.LMCDed, PayrollMaster.ValewithClient, " & _
                "PayrollMaster.CashAdvance, PayrollMaster.WTaxAmount, PayrollMaster.MiscFee, PayrollMaster.CalamityLoan, PayrollMaster.SSSEmployee, PayrollMaster.SSSEmployer, PayrollMaster.PHealthEmployee, PayrollMaster.PHealthEmployer, " & _
                "PayrollMaster.UnionDues, PayrollMaster.HazardPay, PayrollMaster.Medical, PayrollMaster.SalaryLoan, PayrollMaster.Uniform, PayrollMaster.Variances, PayrollMaster.Vale_IOU, PayrollMaster.OTHours, PayrollMaster.OTPay, " & _
                "PayrollMaster.DaysAbsent, PayrollMaster.HoursTardiness, PayrollMaster.AbsentDeduction, PayrollMaster.TardinessDeduction, datename(month, payrollmaster.StartPaydate) AS MonthOF, " & _
                "PayrollMaster.PagibigEmployee, PayrollMaster.PagibigEmployer, PayrollMaster.SSSLoan, PayrollMaster.PagibigLoan, PayrollMaster.StartPayDate, Year(PayrollMaster.StartPayDate) AS YearOf, PayrollMaster.EndPayDate, PayrollMaster.MonthNo, PayrollMaster.CutOffNo, ClientTable.ClientName " & _
                "FROM (PayrollMaster INNER JOIN HREmployees ON (PayrollMaster.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollMaster.ClientCode = ClientTable.ClientCode " & _
                "WHERE " & whereCon & " ORDER BY PayrollMaster.ClientCode, PayrollMaster.StartPayDate, PayrollMaster.EndPayDate, HREmployees.Surname, HREmployees.FirstName"
    End If
    
GetData1:

    Set rsDeductions = ReturnRst(StrQuery, cnnstr, mstatic)
    If PhilHealthFlag = True Then
        GoTo ExitRtn1
    End If
ElseIf OptReport(9).VALUE = True Then
        If cmbClient.Text = "--ALL CLIENTS--" Then ' if all clients, do not proceed
            MsgBox "Please select specific client only! Cannot process.", vbInformation
            GoTo ExitRtn1
        End If
        Set rsDeductions = GetPayrollTrans(False, mClientCode, dtStartDate.VALUE, dtEndDate.VALUE)
Else
    'If OptScope(2).VALUE = True and instr("Accounting, Administrator", accounttype) = Then
    Set rsDeductions = GetPayrollTrans(False, mClientCode, dtStartDate.VALUE, dtEndDate.VALUE)
End If
    
If rsDeductions.RecordCount = 0 Then
    MsgBox "No records gathered from master database. Cannot proceed!", vbInformation, PayrollAlert
    GoTo ExitRtn1
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
        GoTo ExitRtn1
    End If
    End With
End If

'If OptReport(7).VALUE = True Then           ' for Head count report
'    Call CreateMasterList
'    'Call SalariesAndWagesReport(rsDeductions, cmbMonthof.Text, cmbYearof.Text)
''ElseIf OptReport(9).VALUE = True Then      ' for payroll checklist report
''    Call PrintChecklistReport(rsDeductions, cmbMonthof.Text, cmbYearof.Text)
If OptReport(0).VALUE = True Then       ' for Payroll Summary Report
    Call PrintPayrollSummary(rsDeductions, 4)
ElseIf OptReport(6).VALUE = True Then       ' for lates and absent Report
    Call GenerateReports(rsDeductions)
ElseIf OptReport(0).VALUE = True Then       ' Payroll Summary
    If OptScope(2).VALUE = True Then
        Call PrintGrossSummaryReport(rsDeductions, Format(dtStartDate.VALUE, "mmmm"), Year(dtEndDate.VALUE))
    Else
        Call GenerateReports(rsDeductions)
        'Call PrintPayrollSummary(rsDeductions, 4)
    End If
Else
    ' Check first if for Form 1601. Updated May 14, 2013.
    If OptScope(1).VALUE = True And OptReport(9).VALUE = True Then
        If FormSw = True Then
            ComputeTaxSummary
        End If
        GoTo ExitRtn1
    ElseIf OptScope(2).VALUE = True And OptReport(9).VALUE = True Then
        ComputeTaxSummary False
        PreviewBIRForm "Form 2316.rpt", False
        GoTo ExitRtn1
    Else
        Call GenerateReports(rsDeductions)
    End If
End If

If ReportType = 66 Then
    ReportType = 6
End If

ExitRtn1:

Me.MousePointer = vbDefault

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

Private Function GetSalariesAndWages(ByVal PayrollSw As Boolean, _
    Optional ClientCd As Integer, Optional mSMonth As Date, Optional mEMonth As Date) As ADODB.Recordset
Dim rsPayTrans As ADODB.Recordset

'mSMonth = Format(mSMonth, "mm/dd/yyyy")
'mEMonth = Format(mEMonth, "mm/dd/yyyy")

StrQuery = "SELECT * FROM dbo.vwSalaryTotal " & _
    "WHERE ClientCode = " & mClientCode & " and StartPaydate = '" & Format(mSMonth, "mm/dd/yyyy") & "' and EndPaydate = '" & Format(mEMonth, "mm/dd/yyyy") & "'"
Set rsPayTrans = ReturnRst(StrQuery, cnnstr, mstatic)
If rsPayTrans.RecordCount = 0 Then
    rsPayTrans.Close
End If
StrQuery = "SELECT * FROM dbo.vwSalaryMasterTotal " & _
    "WHERE ClientCode = " & mClientCode & " and StartPaydate = '" & Format(mSMonth, "mm/dd/yyyy") & "' and EndPaydate = '" & Format(mEMonth, "mm/dd/yyyy") & "'"
Set rsPayTrans = ReturnRst(StrQuery, cnnstr, mstatic)

Set GetSalariesAndWages = rsPayTrans

End Function

Private Function GetPayrollTrans(ByVal PayrollSw As Boolean, _
    Optional ClientCd As Integer, Optional mSMonth As Date, Optional mEMonth As Date) As ADODB.Recordset
Dim rsPayTrans As ADODB.Recordset
Dim StartDate As String
Dim EndDate As String
Dim CodeList As String, CodeFlag As Boolean
'Dim WhereCon As String

CodeFlag = False
If PayrollSw = True Then    ' for current payroll processing
    StartDate = frmSelectCutoff.txtCutOffDate(0).Text
    EndDate = frmSelectCutoff.txtCutOffDate(1).Text
    whereCon = "WHERE a.clientcode = " & mClientCode & _
        " and a.startpaydate = '" & StartDate & "' and a.endpaydate = '" & EndDate & "'" & _
        " ORDER BY a.employeename"
Else    ' for payroll summary per cut-off
    If ClientCd <> 0 Then ' if specific Clients
        mClientCode = ClientCd
        If OptScope(1) = True Then
            whereCon = "WHERE a.clientcode = " & mClientCode & _
                " and CONVERT(datetime, a.StartPayDate) >= '" & Format(mSMonth, "mm/dd/yyyy") & "'" & _
                " and CONVERT(datetime, a.EndPayDate) <= '" & Format(mEMonth, "mm/dd/yyyy") & "'" & _
                " ORDER BY a.startpaydate, a.endpaydate, a.employeename"
        ElseIf OptScope(2) = True Then
            CodeFlag = True
            CodeList = "(SELECT ClientCode FROM ClientTable WHERE ClientCode2 <> 'FMS' AND LEFT(ClientName,10) IN (SELECT LEFT(CLientName,10) FROM ClientTable WHERE ClientCode = " & mClientCode & "))"
            whereCon = "WHERE a.clientcode IN " & CodeList & _
                " and CONVERT(datetime, a.StartPayDate) >= '" & Format(mSMonth, "mm/dd/yyyy") & "'" & _
                " and CONVERT(datetime, a.EndPayDate) <= '" & Format(mEMonth, "mm/dd/yyyy") & "'" & _
                " ORDER BY d.ClientName, a.startpaydate, a.endpaydate, a.employeename"
        ElseIf OptScope(3) = True Then  ' for payroll summary combined Admin, Inspector and Mechanic
            CodeFlag = True
            CodeList = LTrim(RTrim(Replace(cmbClient.Text, "ADMIN", "")))
            CodeList = "%" & CodeList & "%"
            CodeList = "(SELECT ClientCode FROM ClientTable WHERE ClientName IN (SELECT CLientName FROM ClientTable WHERE ClientName like '" & CodeList & "')"
            whereCon = "WHERE a.clientcode IN " & CodeList & _
                " and CONVERT(datetime, a.StartPayDate) >= '" & Format(mSMonth, "mm/dd/yyyy") & "'" & _
                " and CONVERT(datetime, a.EndPayDate) <= '" & Format(mEMonth, "mm/dd/yyyy") & "')" & _
                " ORDER BY d.ClientName, a.startpaydate, a.endpaydate, a.employeename"
        Else
            whereCon = "WHERE a.clientcode = " & mClientCode & _
                " and CONVERT(datetime, a.StartPayDate) >= '" & Format(mSMonth, "mm/dd/yyyy") & "'" & _
                " and CONVERT(datetime, a.EndPayDate) <= '" & Format(mEMonth, "mm/dd/yyyy") & "'"
        End If
    Else        ' ALL Clients
        If OptScope(1) = True Then
            whereCon = "WHERE CONVERT(datetime, a.StartPayDate) >= '" & Format(mSMonth, "mm/dd/yyyy") & "'" & _
                    " and CONVERT(datetime, a.EndPayDate) <= '" & Format(mEMonth, "mm/dd/yyyy") & "'" & _
                    " ORDER BY d.clientname, a.startpaydate, a.endpaydate, a.employeename"
        ElseIf OptScope(2) = True Then
            MsgBox "ALL Client selection for Gross Summary Report is not allow!", vbInformation, PayrollAlert
            whereCon = "WHERE a.clientcode = 999"
        Else
            whereCon = "WHERE CONVERT(datetime, a.StartPayDate) >= '" & Format(mSMonth, "mm/dd/yyyy") & "'" & _
                    " and CONVERT(datetime, a.EndPayDate) <= '" & Format(mEMonth, "mm/dd/yyyy") & "'"
        End If
    End If
End If
'Set rsPayTrans = New ADODB.Recordset

If OptReport(7).VALUE = True Then   ' Head Count Report
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
    
'    ' Insert the Head Count from Master Payroll table.
'    StrQuery = "INSERT INTO tmpPayrollCount (AgencyCode, ClientCode, HeadCount, PayMonth, PayYear) " & _
'        "SELECT AgencyCode, ClientCode, HeadCount, StartPayDate, StartPayYear FROM qryPayMastCount " & _
'        " WHERE StartPayDate = '" & mMonth & "' and StartPayYear = " & mYear
'    Call Action_Query(StrQuery, cnnstr)
'    ' Insert the Head Count from Temporary Payroll table.
'    StrQuery = "INSERT INTO tmpPayrollCount (AgencyCode, ClientCode, HeadCount, PayMonth, PayYear) " & _
'        "SELECT AgencyCode, ClientCode, HeadCount, StartPayDate, StartPayYear FROM qryPayTransCount " & _
'        " WHERE StartPayDate = '" & mMonth & "' and StartPayYear = " & mYear
'    Call Action_Query(StrQuery, cnnstr)
'
'    StrQuery = "UPDATE tmpPayrollCount SET UserID = '" & mUserID & "' WHERE userID IS NULL"
'    Call Action_Query(StrQuery, cnnstr)
    
    StrQuery = "UPDATE tmpHRCount INNER JOIN tmpPayrollCount " & _
        " ON (tmpHRCount.ClientCode = tmpPayrollCount.ClientCode) " & _
        " AND (tmpHRCount.AgencyCode = tmpPayrollCount.AgencyCode) " & _
        " SET tmpHRCount.HeadCount = tmpHRCount.HeadCount + tmppayrollcount.HeadCount, tmpHRCount.PayMonth = tmpPayrollCount.PayMonth, tmpHRCount.PayYear = tmpPayrollCount.PayYear"
    Call Action_Query(StrQuery, cnnstr)
    StrQuery = "SELECT AgencyCode, ClientCode, ClientName, SUM(EmpCount) As HRCount, MAX(HeadCount) AS TotalCount " & _
        " FROM tmpHRCount WHERE userID = '" & mUserID & "' GROUP BY AgencyCode, ClientCode, ClientName ORDER BY AgencyCode, ClientName "
ElseIf OptReport(9).VALUE = True Then

'    ' For Payroll Checklist
'    mUserID = Trim(UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID))))
'    StrQuery = "DELETE FROM tmpPayrollChecklist WHERE useriD like '%" & mUserID & "%'"
'    Call Action_Query(StrQuery, cnnstr)
'    StrQuery = "INSERT INTO tmpPayrollChecklist (AgencyCode, ClientCode, EmployeeID, AgencyBDOCardFee, SSSLoanAmount, SSSLoanPayment, StartSSSPayment, LastSSSPayment, PbigLoanAmount, PbigLoanPayment, StartPbigPayment, LastPbigPayment, ValeLoanAmount, ValeLoanPayment, StartValePayment, LastValePayment, UpdatedBy, DateUpdated, EmpNo, EmpName, DailyRate, StartContract, EndContract, CBAmount, CBPayment, CBStartPayment, CBEndPayment, MiscFee, MFPayment, MFStartPayment, MFEndPayment, CashAdvance, PFPayment, PFStartPayment, PFENdPayment, LastUpdate, AccountNo, CutOffPeriod, SalaryRelease, Remarks) " & _
'        "SELECT AgencyCode, ClientCode, EmployeeID, AgencyBDOCardFee, SSSLoanAmount, SSSLoanPayment, StartSSSPayment, LastSSSPayment, PbigLoanAmount, PbigLoanPayment, StartPbigPayment, LastPbigPayment, ValeLoanAmount, ValeLoanPayment, StartValePayment, LastValePayment, UpdatedBy, DateUpdated, EmpNo, EmpName, DailyRate, StartContract, EndContract, CBAmount, CBPayment, CBStartPayment, CBEndPayment, MiscFee, MFPayment, MFStartPayment, MFEndPayment, CashAdvance, PFPayment, PFStartPayment, PFENdPayment, LastUpdate, AccountNo, CutOffPeriod, SalaryRelease, Remarks " & _
'        " FROM qryDeductionChecklist WHERE ClientCode = " & mClientCode & " AND EmploymentStatus IN ('C')"
'    Call Action_Query(StrQuery, cnnstr)
'    StrQuery = "UPDATE tmpPayrollChecklist SET UserID = '" & mUserID & "' WHERE userID IS NULL"
'    Call Action_Query(StrQuery, cnnstr)
'    StrQuery = "SELECT DISTINCT * FROM tmpPayrollChecklist WHERE ClientCode = " & mClientCode & " AND userID = '" & mUserID & "' ORDER BY EmpName"
Else    ' Payroll Summary Report
    If OptScope(1) = True Or OptScope(2) = True Or OptScope(3) = True Then
        ' Read the Master payroll table.
        'Enhanced by Bernard a.EmpNo changed to b.DTRID
        StrQuery = "SELECT b.DTRID as controlnum, a.CashAdvance AS CashAdvance2, a.*, b.acctno, c.*, d.ClientName, d.ClientSeriesNo " & _
            " FROM ((PayrollMaster AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID) " & _
            "LEFT JOIN qryEmployeeBasics AS c ON a.EmployeeID = c.EmployeeID) LEFT JOIN ClientTable AS d ON a.clientcode = d.clientcode " & whereCon
    Else
        ' Read the Master payroll table.
        StrQuery = "select empno, min(employeename) as EmpName, sum(grosspay) as GPay, sum(month13rate) as m13Rate From payrollmaster a " & whereCon & " group by a.empno"
    End If
End If

Set rsPayTrans = ReturnRst(StrQuery, cnnstr, mstatic)
'Check if there were records.
If rsPayTrans.EOF = True Then
    Set rsPayTrans = Nothing
    'Set rsPayTrans = New ADODB.Recordset
    
    If OptScope(1) = True Or OptScope(2) = True Or OptScope(3) = True Then
        ' Enhanced by Bernard a.EmpNo changed to b.DTRID
        StrQuery = "SELECT b.DTRID as controlnum, a.CashAdvance AS CashAdvance2, a.*, b.acctno, c.*, d.ClientName, d.ClientSeriesNo " & _
            "FROM ((PayrollTrans AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID) " & _
            "LEFT JOIN qryEmployeeBasics AS c ON a.EmployeeID = c.EmployeeID) LEFT JOIN ClientTable AS d ON a.clientcode = d.clientcode " & whereCon
    Else
        ' Read the Master payroll table.
        StrQuery = "select empno, min(employeename) as EmpName, sum(grosspay) as GPay, sum(month13rate) as m13Rate From payrollmaster AS a " & whereCon & " Group by a.empno "
    End If
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
If cmdClose.Caption = "Cancel" Then
    cmdClose.Caption = "Close"
    cmdProcess.Caption = "Generate"
    cmdProcess.ToolTipText = "Generate report..."
Else
    Unload Me
End If
End Sub

Private Sub cmdTotals_Click()
If cmdTotals.Caption = "Compute Totals" Then
    ' Initially get those records subject for computation.
    Flag1601 = False
    Call ProcessReports(True)
Else 'Save Totals
    If Flag1601 = True Then
        Call UpdateBIRFormRecords
    End If
    cmdTotals.Caption = "Compute Totals"
End If
End Sub

Private Sub UpdateBIRFormRecords()
Dim rs As Recordset

' Update the totals of BIR data.
StrQuery = "UPDATE Form1601C SET GrossPay = " & txtTax(0).Text & ", Tax = " & txtTax(1).Text
Call Action_Query(StrQuery, cnnstr)

StrQuery = "SELECT TOP 1 * FROM Form1601C"
Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
Set rsTaxTable = rs
cmdTotals.Caption = "Compute Totals"
Set rs = Nothing

End Sub

Private Sub dtStartDate_LostFocus()
    
If Day(dtStartDate.VALUE) <= 15 Then
    CutOffPeriod = "START"
Else
    CutOffPeriod = "END"
End If

End Sub

Private Sub Form_Load()
Dim j As Integer

'    If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'        StrQuery = "SELECT * FROM clienttable WHERE CLIENTCODE2 = 'FMS' AND clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
'    Else
'        StrQuery = "SELECT * FROM clienttable WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
'    End If

'Populate Clients.
StrQuery = GetClientRecord(0, False, " ORDER BY ClientName")
'StrQuery = "SELECT * FROM clienttable WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
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

dtStartDate.VALUE = Format(Now - 31, "mm/dd/yyyy")
dtEndDate.VALUE = Format(Now, "mm/dd/yyyy")

If Day(dtStartDate.VALUE) <= 15 Then
    CutOffPeriod = "START"
Else
    CutOffPeriod = "END"
End If

'mSapi = CompanyName & " (" & CompanyShort & ")"
'mSapi = "BENSAN TRANS CORPORATION"
mSapi = "TAS TRANS (Group of Companies)"
mShortName = "(" & CompanyShort & ")"

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
StrQuery = "SELECT * FROM DeductionTypes2 WHERE DeductionCode <> 0 ORDER BY DeductionDescription"
Set rsGetDeductionTypes = ReturnRst(StrQuery, CnnstrHR, mstatic)
    
' Fill the combo with New Deductions.
Call FillDeductions(cmbDeductions, rsGetDeductionTypes)

dtStartDate.VALUE = Now - 30
dtEndDate.VALUE = Now
'LocalClient = cmbClient.Text
    
If InStr(1, ReportAccess, LoginName) > 0 Then
    For j = 0 To 9
        OptReport(j).Enabled = False
    Next j
    OptReport(7).Enabled = True
    OptReport(7).VALUE = True
End If
    
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

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
txtChange.Text = ListView1(Index).SelectedItem.ListSubItems(4).Text
txtChangeID.Text = ListView1(Index).SelectedItem.ListSubItems(6).Text
End Sub

Private Sub optReport_Click(Index As Integer)

chkLoans.Visible = False
chkLoans.VALUE = 0
chkCal.Visible = False
chkCal.VALUE = 0

If Index = 1 Then
    chkCal.Visible = True
End If
cmbDeductions.Enabled = False
OptScope(0).Enabled = False
OptScope(2).Enabled = True
OptScope(3).Enabled = False
OptScope(1).Enabled = True
OptScope(1).VALUE = True

OptScope(0).Caption = "Per Cut-off"
OptScope(1).Caption = "Detailed Report"
OptScope(2).Caption = "Gross Summary"

cmbEmployee.Enabled = True
chkAllEmployees.Enabled = True

Frame5.Visible = False
SSTab2.Height = 4065
FrmOtherReport.Height = 5100

cmdTotals.Visible = False
Frame6.Visible = False
Flag1601 = False
Label1(3).Caption = "Selected Only"

'Select report type and header.
Select Case Index
    Case 0  ' Payroll Summary Report
        ReportType = 0
        Label1(3).Caption = "Select Dept: "
        OptScope(3).Caption = "Payroll Summary"
        ReportHeader = "Payroll Summary Report"
        cmbEmployee.Enabled = True
        chkAllEmployees.Enabled = True 'False updated Feb. 1, 2015
        
        OptScope(0).Enabled = False
        OptScope(2).Enabled = True ' for Gross Summary Report
        OptScope(3).Enabled = True ' for payroll summary
        OptScope(1).VALUE = True
        
        'If InStr("Accounting, Administrator", AccountType) = 0 Then
        If ForPayroll(AccountType) = False Then
            OptScope(1).Enabled = False
        End If
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
        If chkCal.VALUE = 1 Then
            ReportHeader = "Calamity Loan Report"
        End If
        chkLoans.Visible = True
        chkCal.Visible = True
    Case 4  ' PhilHealth
        ReportType = 4
        ReportHeader = "PhilHealth Contributions Report"
    Case 5  ' 13th Month
        Frame5.Visible = True
        SSTab2.Height = 6465
        FrmOtherReport.Height = 7500
        ReportType = 5
        ReportHeader = "13th Month Pay Report"
    Case 6  ' Late and Absent
        ReportType = 6
        ReportHeader = "Lates and Absent Report"
    Case 7  ' Masterlist Report
        ReportType = 7
        ReportHeader = "Employee Masterlist"
        OptScope(0).Enabled = False
        OptScope(1).VALUE = True
        OptScope(1).Enabled = False
        OptScope(2).Enabled = False
        OptScope(3).Enabled = False
    Case 777  ' Tax Forms (Previously Emp Salary)
        ReportType = 7
        OptScope(0).Caption = "Form 1601-C"
        OptScope(1).Caption = "Form 2316"
        
        ReportHeader = "BIR Forms"
        OptScope(0).Enabled = True
        OptScope(0).VALUE = True
        OptScope(1).Enabled = True
        OptScope(2).Enabled = False
        OptScope(3).Enabled = False
    Case 8  ' Other Deductions
        ReportType = 8
        ReportHeader = "Payroll Deductions Report"
        cmbDeductions.Enabled = True
    Case 9  ' Tax Alphalist
        ReportType = 9
        ReportHeader = "Alphalist Report"

        OptScope(0).Caption = "Alphalist"
        OptScope(1).Caption = "Form 1601-C"
        OptScope(2).Caption = "Form 2316"
        
        'ReportHeader = "BIR Forms"
        OptScope(0).Enabled = True
        OptScope(0).VALUE = True
        OptScope(1).Enabled = True
        OptScope(2).Enabled = True
        OptScope(3).Enabled = False

End Select

End Sub

Private Sub ComputeTaxSummary(Optional F1601Sw As Boolean = True)
Dim rs As Recordset
Dim PerEmployee As String

On Error GoTo Exit1

PerEmployee = ""
'Truncate the contents of temporary table.
StrQuery = "EXEC TruncateTableForm1601C " & IIf(F1601Sw = True, 1, 0)
Call Action_Query(StrQuery, cnnstr)

' Update temporary summary table.
If F1601Sw = True Then
    StrQuery = "INSERT INTO Form1601C (Company, Address, TelNo, ContactPerson, Position, TIN, ZipCode, GrossPay, Tax, MonthOf) " & _
        "SELECT MAX(b.ClientReportName) As Company, MAX(b.ClientAddress) AS Address, MAX(b.ClientTelNo) AS TelNo, " & _
        "MAX(b.ClientContactPerson) AS ContactPerson, MAX(b.ClientContactPosition) AS Position, MAX(b.ClientTaxNumber) AS TIN, " & _
        "MAX(b.ClientPagibig) AS ZipCode, SUM(a.Grosspay) as GrossPay, SUM(a.WtaxAmount) as Tax, RIGHT('00' + RTRIM(CONVERT(Char(2), MONTH(MAX(a.StartPaydate)))), 2) + '   ' + RIGHT(RTRIM(CONVERT(Char(4), YEAR(MAX(a.StartPayDate)))), 4) " & _
        "FROM tmpPayReport a LEFT JOIN Clienttable b on a.clientname = b.clientname " & _
        "WHERE LTRIM(RTRIM(userID)) = '" & Trim$(mUserID) & "' AND a.WTaxAmount <> 0 GROUP BY Monthof, YearOf"
Else    ' for Form 2316
    ' Check if selected employees only, then filter the recordset.
    If chkAllEmployees.VALUE = 0 Then
        If cmbClient.Text <> "--ALL CLIENTS--" Then ' check for all clients
            PerEmployee = " AND a.EmployeeID = '" & Trim(cmbEmployee.ItemData(cmbEmployee.ListIndex)) & "'"
        End If
    End If
    StrQuery = "INSERT INTO Form1601C (EmployeeID, Surname, FirstName, Middlename, Fullname, " & _
        "TIN, TaxCode, MinWagePerDay, MinWagePerMonth, BasicPay, GrossPay, Allowance, SSSEmployee, pHealthEmployee, PagIbigEmployee, WtaxAmount, BirthDate, " & _
        "Company, Address, Owner, EmpTIN, ZipCode, " & _
        "Month13, Exemptions, OTPay, HazardPay, MonthOF, YearOf) " & _
        "SELECT a.EmployeeID, a.Surname, a.FirstName, a.Middlename, " & _
        "LTRIM(RTRIM(a.Surname)) + ', ' + LTRIM(RTRIM(a.Firstname)) + ' ' + CASE WHEN a.Middlename IS NULL THEN '' ELSE LEFT(a.Middlename,1) + '.' END AS Fullname, " & _
        "MAX(a.TIN) AS TIN, MAX(a.TaxCode) AS TaxCode, MAX(b.DailyRate), MAX(b.MonthlyRate), SUM(a.BasicPay) AS BasicPay, SUM(a.GrossPay) AS GrossPay, SUM(a.Allowance), SUM(a.SSSEmployee), SUM(a.pHealthEmployee), SUM(a.PagIbigEmployee), SUM(a.WtaxAmount), MAX(a.BirthDate), " & _
        "MAX(c.ClientReportName), MAX(c.ClientAddress), MAX(c.ClientOwner), MAX(c.ClientTaxNumber), MAX(c.ClientPagibig), " & _
        "SUM(a.Month13Payment) AS Month13, SUM(a.Exempt) AS Exemptions, Sum(a.OTPay), Sum(a.HazardPay), Max(a.MonthOf), Max(a.YearOf) " & _
        "FROM tmpPayReport a LEFT JOIN HREmployees b ON a.EmployeeID = b.EmployeeID AND a.Surname = b.Surname LEFT JOIN ClientTable c ON a.ClientName = c.CLientName WHERE LTRIM(RTRIM(a.UserID)) = '" & Trim$(mUserID) & "' " & PerEmployee & " GROUP BY a.Surname, a.Firstname, a.Middlename, a.EmployeeID "
End If
Call Action_Query(StrQuery, cnnstr)

If F1601Sw = True Then
    StrQuery = "SELECT TOP 1 * FROM Form1601C"
    Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
    If rs.EOF = False Then ' Save the value for possible editing
        txtTax(0).Text = rs!GrossPay
        txtTax(1).Text = rs!TAX
        Set rsTaxTable = rs
        Flag1601 = True
        cmdTotals.Caption = "Save Totals"
    Else
        txtTax(0).Text = 0#
        txtTax(1).Text = 0#
        Flag1601 = False
    End If
Else
    StrQuery = "SELECT * FROM Form1601C ORDER BY Surname, Firstname, Middlename"
    Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
    If rs.EOF = False Then ' check if there is record then save to recordset tax table.
        Set rsTaxTable = rs
    End If
End If
Set rs = Nothing
Exit Sub

Exit1:

    Err.Clear
    
End Sub

Private Sub CreateMasterList()
Dim str As String, SELECTED As String, WhereConn As String
Dim rsTemp As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet1 As Excel.Worksheet
Dim ctr As Long
Dim i As Long

SELECTED = " Empname, ClientName, EmpNo, DTRID, Deptname, Empposition, PaymentBasis, DateHired, StartContract, EndContract, SSS, Philhealth, TIN, Pagibig, BirthDate, TaxCode, EcolaAmount, AcctNo, Supervisor, EmploymentStatusDesc, CommissionRate "
WhereConn = ""
If cmbClient.Text = "--ALL CLIENTS--" Then
    MsgBox "Too many records selected. Please select specific Company.", vbInformation, PayrollAlert
    Exit Sub
Else
    WhereConn = " WHERE clientname = '" & cmbClient.Text & "' And EmploymentStatus in ('C', 'P', 'R') ORDER BY ClientName, EmpName"
End If

sqlString = "SELECT " & _
             "" & Replace(SELECTED, "controlno", "empno") & "" & _
              " From vwEmployeeCustomInfo " & WhereConn
             
Set rsTemp = ReturnRst(sqlString, cnnstr, mstatic)
If rsTemp.EOF = False Then
    Set xlApp1 = New Excel.Application
   Set xlBook1 = xlApp1.Workbooks.Add()
   xlApp1.Visible = True
   CreateExcel rsTemp, xlBook1, xlApp1
   xlApp1.DisplayAlerts = False
Else
   MsgBox "No Records!!!", vbCritical, PayrollAlert
End If

Set rsTemp = Nothing

End Sub

Function CreateExcel(rs As ADODB.Recordset, book1 As Excel.Workbook, ByVal xlApp1 As Excel.Application)
    Dim xlSheet As Excel.Worksheet
    Dim ctr As Long
    Dim i As Long
    Set xlSheet = book1.Worksheets.Add
    ctr = 7
    i = 1
    xlSheet.Cells(1, 1).VALUE = "EMPLOYEE MASTERLIST " & IIf(cmbClient.ListIndex >= 0, " ASSIGNED TO " & cmbClient.Text, "")
    xlSheet.Cells(2, 1).VALUE = "As of:"
    xlSheet.Cells(2, 3).VALUE = Format(Date, "mmmm dd, yyyy")
    xlSheet.Cells(3, 1).VALUE = "Company:"
    xlSheet.Cells(3, 3).VALUE = cmbClient.Text
    xlSheet.Cells(4, 1).VALUE = "Prepared by:"
    xlSheet.Cells(4, 3).VALUE = UCase(GETuSERNAME(operatorID))
    
    xlSheet.Cells(6, 1).VALUE = "No."
    xlSheet.Cells(6, 2).VALUE = "Name"
    xlSheet.Cells(6, 3).VALUE = "Company Name"
    xlSheet.Cells(6, 4).VALUE = "ID No."
    xlSheet.Cells(6, 5).VALUE = "DTR ID"
    xlSheet.Cells(6, 6).VALUE = "Department"
    xlSheet.Cells(6, 7).VALUE = "Position"
    xlSheet.Cells(6, 8).VALUE = "Payment Base"
    xlSheet.Cells(6, 9).VALUE = "Date Hired"
    xlSheet.Cells(6, 10).VALUE = "Years of Service" ' updated 8/30/2012
    xlSheet.Cells(6, 11).VALUE = "Date From"
    xlSheet.Cells(6, 12).VALUE = "Date To"
    
    xlSheet.Cells(6, 13).VALUE = "SSS Number"
    xlSheet.Cells(6, 14).VALUE = "PhilHealth"
    xlSheet.Cells(6, 15).VALUE = "TIN Number"
    xlSheet.Cells(6, 16).VALUE = "Pagibig"
    xlSheet.Cells(6, 17).VALUE = "Birthday"
    xlSheet.Cells(6, 18).VALUE = "Age"  ' updated 8/30/2012
    xlSheet.Cells(6, 19).VALUE = "CS"
    xlSheet.Cells(6, 20).VALUE = "Ecola Rate"
    xlSheet.Cells(6, 21).VALUE = "Account No"
    xlSheet.Cells(6, 22).VALUE = "Supervisor"
    xlSheet.Cells(6, 23).VALUE = "Employment Status"
    
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Cells.EntireRow.AutoFit
    xlSheet.Cells.Range("A1:Z6").Font.Bold = True
    xlSheet.Cells.Range("A1:A5").HorizontalAlignment = xlLeft
    xlSheet.Cells.Range("C1:D5").HorizontalAlignment = xlLeft
    xlSheet.Cells.Range("A6:Z6").Interior.ColorIndex = 15
    xlSheet.Columns("A:A").ColumnWidth = 5.71
    
    Do While Not rs.EOF
        DoEvents
        xlSheet.Cells(ctr, 1).VALUE = i
        xlSheet.Cells(ctr, 2).VALUE = rs!EmpName
        xlSheet.Cells(ctr, 3).VALUE = rs!ClientName
        xlSheet.Cells(ctr, 4).VALUE = rs!EmpNo
        xlSheet.Cells(ctr, 5).VALUE = rs!DTRID
        xlSheet.Cells(ctr, 6).VALUE = rs!deptname
        xlSheet.Cells(ctr, 7).VALUE = rs!empPosition
        xlSheet.Cells(ctr, 8).VALUE = rs!PaymentBasis
        xlSheet.Cells(ctr, 9).VALUE = rs!datehired
        xlSheet.Cells(ctr, 9).NumberFormat = "mm/dd/yyyy"
        If IsNull(rs!datehired) = True Then
            xlSheet.Cells(ctr, 10).VALUE = 0
        Else
            xlSheet.Cells(ctr, 10).VALUE = Year(Now) - Year(rs!datehired)
        End If
        xlSheet.Cells(ctr, 11).VALUE = rs!StartContract
        xlSheet.Cells(ctr, 11).NumberFormat = "mm/dd/yyyy"
        xlSheet.Cells(ctr, 12).VALUE = rs!endcontract
        xlSheet.Cells(ctr, 12).NumberFormat = "mm/dd/yyyy"
    
        xlSheet.Cells(ctr, 13).VALUE = IIf(rs!sss = "__-_______-_", "", rs!sss)
        xlSheet.Cells(ctr, 14).VALUE = IIf(rs!PhilHealth = "__-_________-_", "", rs!PhilHealth)
        xlSheet.Cells(ctr, 15).VALUE = IIf(rs!TIN = "___-___-___", "", rs!TIN)
        xlSheet.Cells(ctr, 16).VALUE = rs!Pagibig
        xlSheet.Cells(ctr, 17).VALUE = rs!BIRTHDATE
        xlSheet.Cells(ctr, 17).NumberFormat = "mm/dd/yyyy"
        If IsNull(rs!BIRTHDATE) = True Then
            xlSheet.Cells(ctr, 18).VALUE = 0
        Else
            xlSheet.Cells(ctr, 18).VALUE = Year(Now) - Year(rs!BIRTHDATE)
        End If
        xlSheet.Cells(ctr, 19).VALUE = rs!TaxCode
        
        If rs!EcolaAmount <> "" Or IsNull(rs!EcolaAmount) = False Then
            xlSheet.Cells(ctr, 20).VALUE = Format(rs!EcolaAmount, "#,##0.00")
        Else
            xlSheet.Cells(ctr, 20).VALUE = Format(0#, "#,##0.00")
        End If
        
        xlSheet.Cells(ctr, 21).VALUE = rs!AcctNo
        xlSheet.Cells(ctr, 22).VALUE = rs!Supervisor
        xlSheet.Cells(ctr, 23).VALUE = rs!EmploymentStatusDesc
        
        i = i + 1
        ctr = ctr + 1
        rs.MoveNext
    Loop
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Cells.EntireRow.AutoFit
    xlSheet.Columns("A:A").ColumnWidth = 5.71
    xlSheet.Columns("D:D").ColumnWidth = 10
    'Center
    xlSheet.Range("A6:W6").Select
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    
    'center id
    
      xlSheet.Range("A7:A" & ctr).Select
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
     
       'line
    'Range("A6:Q14").Select
    xlSheet.Range("A6:W" & ctr).Select
    xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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

    xlSheet.Name = "Masterlist"
    
End Function

Private Sub OptScope_Click(Index As Integer)

'Populate Clients.
StrQuery = GetClientRecord(0, False, " ORDER BY ClientName")
Set rsClientData = ReturnRst(StrQuery, cnnstr, mstatic)
Call oCommonFunction.FillCombo(cmbClient, rsClientData, 9999) ' 99999 do not add any
If rsClientData.RecordCount <> 0 Then cmbClient.ListIndex = 0

cmdTotals.Visible = False
Frame6.Visible = False
Select Case Index
    Case 0  ' Per Cut-off
        cmbQuarter.Enabled = False
        cmbMonthof.Enabled = True
        If ReportType = 9 Then
            ReportScope = 1
        Else
            ReportScope = 0
        End If
        ScopeRep = "Per Cut-off Date " & Format(dtStartDate.VALUE, "mm/dd/yyyy") & " to " & Format(dtEndDate.VALUE, "mm/dd/yyyy")
    Case 1  ' Detailed Report
        cmbQuarter.Enabled = False
        cmbMonthof.Enabled = True
        ReportScope = 1
        ScopeRep = "For Detailed Month Range " & Format(dtStartDate.VALUE, "mmmm") & " " & Format(dtStartDate.VALUE, "yyyy") & " to " & Format(dtEndDate.VALUE, "mmmm") & " " & Format(dtEndDate.VALUE, "yyyy")
        cmdTotals.Caption = "Compute Totals"
        If ReportType = 9 Then
            cmdTotals.Visible = True
            Frame6.Visible = True
        End If
    Case 2  ' Gross Summary
        cmbQuarter.Enabled = True
        cmbMonthof.Enabled = False
        ReportScope = 1
        ScopeRep = "For Gross Summary Month " & Format(dtStartDate.VALUE, "mmmm") & " " & Format(dtStartDate.VALUE, "yyyy") & " to " & Format(dtEndDate.VALUE, "mmmm") & " " & Format(dtEndDate.VALUE, "yyyy")
    Case 3  ' Annually
        cmbQuarter.Enabled = False
        cmbMonthof.Enabled = False
        ReportScope = 3
        ScopeRep = "For the Year " & Format(dtStartDate.VALUE, "yyyy")
        If OptReport(0).Index = 0 Then  ' for Payroll
            ' Re-filter the selection of Company.
            StrQuery = GetClientRecord(88, False, " ORDER BY ClientName")
            Set rsClientData = ReturnRst(StrQuery, cnnstr, mstatic)
            Call oCommonFunction.FillCombo(cmbClient, rsClientData, 9999) ' 99999 do not add any
            If rsClientData.RecordCount <> 0 Then cmbClient.ListIndex = 0
        
        End If
End Select

End Sub

Private Sub GetReportScope()
Dim StartMonth As Integer
Dim EndMonth As Integer
Dim mQuery As String

'StartMonth = Val(getMonthNo(cmbMonthof.Text))
'EndMonth = Val(getMonthNo(cmbEndMonth.Text))

StartMonth = Val(getMonthNo(Format(dtStartDate.VALUE, "mmmm")))
EndMonth = Val(getMonthNo(Format(dtEndDate.VALUE, "mmmm")))
cmbYearof.Text = Year(dtEndDate.VALUE)
'VAL(Format(a.StartContract, ""mm"")) >= " & getMonthNum(txtAppMonth.Text, txtAppYear.Text) & " AND VAL(Format(a.StartContract, ""yyyy"")) = " & txtAppYear.Text

Select Case ReportScope
    Case 0  ' Per Cut-off
        'whereCon = " DATENAME(month, payrollmaster.startpaydate) = '" & cmbMonthof.Text & "' and year(payrollmaster.startpaydate) = " & cmbYearof.Text
        'whereCon = " MONTH(payrollmaster.startpaydate) between " & StartMonth & " and " & EndMonth & " and year(payrollmaster.startpaydate) = " & cmbYearof.Text
        whereCon = " CONVERT(varchar(10), CONVERT(datetime, payrollmaster.startpaydate), 112) BETWEEN " & Val(Format(dtStartDate.VALUE, "yyyymmdd")) & " and " & Val(Format(dtEndDate.VALUE, "yyyymmdd"))
    Case 1  ' Detailed Report
        'whereCon = " DATENAME(month, payrollmaster.startpaydate) = '" & cmbMonthof.Text & "' and year(payrollmaster.startpaydate) = " & cmbYearof.Text
        'whereCon = " MONTH(payrollmaster.startpaydate) between " & StartMonth & " and " & EndMonth & " and year(payrollmaster.startpaydate) = " & cmbYearof.Text
        whereCon = " CONVERT(varchar(10), CONVERT(datetime, payrollmaster.startpaydate), 112) BETWEEN " & Val(Format(dtStartDate.VALUE, "yyyymmdd")) & " and " & Val(Format(dtEndDate.VALUE, "yyyymmdd"))
    Case 2  ' Summary Report
        'whereCon = " MONTH(payrollmaster.startpaydate) between " & StartMonth & " and " & EndMonth & " and year(payrollmaster.startpaydate) = " & cmbYearof.Text
        whereCon = " CONVERT(varchar(10), CONVERT(datetime, payrollmaster.startpaydate), 112) BETWEEN " & Val(Format(dtStartDate.VALUE, "yyyymmdd")) & " and " & Val(Format(dtEndDate.VALUE, "yyyymmdd"))
    Case 3  ' Annually
        'whereCon = " MONTH(payrollmaster.startpaydate) between " & StartMonth & " and " & EndMonth & " and year(payrollmaster.startpaydate) = " & cmbYearof.Text
        whereCon = " YEAR(payrollmaster.startpaydate) = " & Year(dtStartDate.VALUE)
End Select

If cmbClient.Text <> "--ALL CLIENTS--" Then ' for selected clients only
    If ReportType = 9 And chkAllTax.VALUE = 1 Then
        mQuery = " and PayrollMaster.ClientCode IN (SELECT ClientCode FROM ClientTable WHERE ClientCode2 <> 'FMS' AND LEFT(ClientName,10) IN (SELECT LEFT(CLientName,10) FROM ClientTable WHERE ClientCode = " & mClientCode & "))"
        whereCon = whereCon & mQuery
'    ElseIf ReportType = 7 Then ' for BIR Forms
'        mQuery = " and PayrollMaster.ClientCode IN (SELECT ClientCode FROM ClientTable WHERE ClientCode2 <> 'FMS' AND LEFT(ClientName,10) IN (SELECT LEFT(CLientName,10) FROM ClientTable WHERE ClientCode = " & mClientCode & "))"
'        whereCon = whereCon & mQuery
    Else
        whereCon = whereCon & " and PayrollMaster.clientcode = " & mClientCode
    End If
End If

End Sub

Private Sub GenerateReports(ByRef rs As ADODB.Recordset)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mRange As String, sumString As String, mRange2 As String, mRange3 As String
Dim ctr As Long, mRange11 As String, mRange12 As String
Dim i As Integer, j As Integer, jj As Integer
Dim cPos As Integer
Dim prevClientCode As Integer
Dim prevControlNo As String
Dim prevCutOff As Variant
Dim StartOfLine As Integer
Dim EndOfLine As Integer
Dim SaveTotal As String, ColString As String, EndRange As String
Dim mRemarks As String, RecNum As Integer
Dim DedTotals As Double, prevVoucher As String

If ReportType = 0 Then
    MsgBox "Not yet finished.", vbInformation
    Exit Sub
ElseIf rs.EOF = True Or rs.BOF = True Then
    MsgBox "No records to process.", vbInformation
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

'Check if Excel has only 1 default sheet. Updated 3/8/2020
If xlApp1.Sheets.Count = 1 Then
    With xlBook1
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Sheet2"
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Sheet3"
    End With
End If
xlApp1.Sheets("Sheet1").Select

jj = 0
cPos = 5
ctr = 7
i = 1
TotalLoan = 0: LoanStart = 0

EndRange = "E"
SaveTotal = ""
ColString = ""

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
    .Font.Size = 12         'font.size
    .Font.Bold = True       ' Font.Bold
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
    ScopeRep = "For Date Range " & Format(dtStartDate.VALUE, "mm/dd/yyyy") & " to " & Format(dtEndDate.VALUE, "mm/dd/yyyy")
    xlSheet.Range("A3") = ScopeRep '& " " & cmbMonthof & " " & cmbYearof
End If

If ReportType <> 0 Then
    'Print the Header.
    If ReportType = 1 Then ' for sss
        xlSheet.Cells(ctr, 1).VALUE = "Surname   "
        xlSheet.Cells(ctr, 2).VALUE = "Firstname "
        xlSheet.Cells(ctr, 3).VALUE = "Middlename"
        xlSheet.Cells(ctr, 4).VALUE = "Date of Birth"
        xlSheet.Cells(ctr, 5).VALUE = "SSS Number"
    ElseIf ReportType = 4 Then  ' for Phealth Updates 9/16/2018
        xlSheet.Cells(ctr, 1).VALUE = "Surname   "
        xlSheet.Cells(ctr, 2).VALUE = "Firstname "
        xlSheet.Cells(ctr, 3).VALUE = "Middlename"
        xlSheet.Cells(ctr, 4).VALUE = "Period Covered"
    Else
        xlSheet.Cells(ctr, 1).VALUE = "Employee #"
        xlSheet.Cells(ctr, 2).VALUE = "Surname   "
        xlSheet.Cells(ctr, 3).VALUE = "Firstname "
        xlSheet.Cells(ctr, 4).VALUE = "Middlename"
    End If
End If

If ReportType = 3 And chkLoans.VALUE = 0 Then ' for Pagibig Contributions enhanced 9/19/2012
    xlSheet.Cells(ctr, 1).VALUE = "PAG-IBIG NO"
    xlSheet.Cells(ctr, 2).VALUE = "FIRSTNAME  "
    xlSheet.Cells(ctr, 3).VALUE = "LASTNAME   "
    xlSheet.Cells(ctr, 4).VALUE = "MIDDLENAME "
End If

If ReportType = 5 Then  ' for 13th month Pay Reporting, updated May 10, 2013
    xlSheet.Cells(ctr, 1).VALUE = "No."
    xlSheet.Cells(ctr, 2).VALUE = "Fullname   "
    xlSheet.Cells(ctr, 3).VALUE = "Date Hired "
    xlSheet.Cells(ctr, 4).VALUE = "Basic Pay  "
End If

Select Case ReportType
    Case 1  ' SSS
        If chkLoans.VALUE = 1 Or chkCal.VALUE = 1 Then ' for SSS Loans
            xlSheet.Cells(ctr - 1, 6).VALUE = "S S S"
            xlSheet.Cells(ctr, 5).VALUE = "SSS Number"
            If chkCal.VALUE = 1 Then
                xlSheet.Cells(ctr, 6).VALUE = "Calamity Amount"
            Else
                xlSheet.Cells(ctr, 6).VALUE = "Loan Amount"
            End If
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
'            xlSheet.Cells(ctr, 5).VALUE = "SSS Number  "
'            xlSheet.Cells(ctr - 1, 6).VALUE = "S     S     S"
'            xlSheet.Cells(ctr, 6).VALUE = "Employee"
'            xlSheet.Cells(ctr, 7).VALUE = "Employer"
'            xlSheet.Cells(ctr, 8).VALUE = "Total    "
            
            xlSheet.Cells(ctr - 1, 6).VALUE = "CONTRIBUTION"
            xlSheet.Cells(ctr, 6).VALUE = "Employee"
            xlSheet.Cells(ctr, 7).VALUE = "Employer"
            xlSheet.Cells(ctr, 8).VALUE = "SSS Prem"
            
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 9).VALUE = "Start Pay"
                xlSheet.Cells(ctr, 10).VALUE = "End Pay"
            Else
                If ReportScope = 1 Then ' new sss format
                    xlSheet.Cells(ctr, 9).VALUE = "EC"
                    xlSheet.Cells(ctr, 10).VALUE = "Total"
                Else
                    xlSheet.Cells(ctr, 9).VALUE = "Month"
                    xlSheet.Cells(ctr, 10).VALUE = "Year"
                End If
            End If
            xlApp1.Sheets("Sheet2").Select
            xlApp1.ActiveSheet.Cells(ctr, 1).VALUE = "SSS Num   "
            xlApp1.ActiveSheet.Cells(ctr, 2).VALUE = "Last Name "
            xlApp1.ActiveSheet.Cells(ctr, 3).VALUE = "First Name"
            xlApp1.ActiveSheet.Cells(ctr, 4).VALUE = "   MI     "
            xlApp1.ActiveSheet.Cells(ctr, 5).VALUE = "Date Hired"
            xlApp1.ActiveSheet.Cells(ctr, 6).VALUE = "  Status  "
            xlApp1.ActiveSheet.Cells(ctr, 7).VALUE = " SSS Prem "
            xlApp1.ActiveSheet.Cells(ctr, 8).VALUE = " EC Premium"
            xlApp1.Sheets("Sheet3").Select
            xlApp1.ActiveSheet.Cells(ctr, 1).VALUE = "SSS Num   "
            xlApp1.ActiveSheet.Cells(ctr, 2).VALUE = "Last Name "
            xlApp1.ActiveSheet.Cells(ctr, 3).VALUE = "First Name"
            xlApp1.ActiveSheet.Cells(ctr, 4).VALUE = "   MI     "
            xlApp1.ActiveSheet.Cells(ctr, 5).VALUE = "Date of Birth"
            xlApp1.Sheets("Sheet1").Select
            mRange = "A1:L10"
            mRange2 = "A6:L7"
            'mRange3 = "E6:G6"
            mRange3 = "F6:H6"
        End If
    Case 2, 9 ' Tax or Alphalist
        If OptReport(9).VALUE = True Then
            xlSheet.Cells(ctr, 5).VALUE = "Tax ID No."
            xlSheet.Cells(ctr, 6).VALUE = "Tax Code"
            xlSheet.Cells(ctr, 7).VALUE = "Gross Pay"
            xlSheet.Cells(ctr, 8).VALUE = "Tax Exemption"
            xlSheet.Cells(ctr, 9).VALUE = "  SSS  "
            xlSheet.Cells(ctr, 10).VALUE = "Pagibig"
            xlSheet.Cells(ctr, 11).VALUE = "PhilHealth"
            xlSheet.Cells(ctr, 12).VALUE = "WTax Amount"
            xlSheet.Cells(ctr, 13).VALUE = "Gross Taxable"
            xlSheet.Cells(ctr, 14).VALUE = "TAX DUE"
            xlSheet.Cells(ctr, 15).VALUE = "13th Month"
            mRange = "A1:O8"
            mRange2 = "A6:O7"
            mRange3 = "H6:H6"
        Else
            xlSheet.Cells(ctr, 5).VALUE = "Tax ID No."
            xlSheet.Cells(ctr, 6).VALUE = "WTax Amount"
            xlSheet.Cells(ctr, 7).VALUE = "Gross Pay"
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 8).VALUE = "Start Pay"
                xlSheet.Cells(ctr, 9).VALUE = "End Pay"
            Else
                xlSheet.Cells(ctr, 8).VALUE = "Month"
                xlSheet.Cells(ctr, 9).VALUE = "Year"
            End If
            mRange = "A1:I8"
            mRange2 = "A6:I7"
            mRange3 = "F6:G6"
        End If
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
        ElseIf chkCal.VALUE = 1 Then ' for Calamity Loans
            xlSheet.Cells(ctr, 5).VALUE = "Pagibig No."
            xlSheet.Cells(ctr - 1, 6).VALUE = "PAGIBIG"
            xlSheet.Cells(ctr, 6).VALUE = "Calamity Loan"
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
            'xlSheet.Cells(ctr, 5).VALUE = "Pagibig No."
            xlSheet.Cells(ctr, 5).VALUE = "COVERED PERIOD"
            'xlSheet.Cells(ctr - 1, 6).VALUE = "P  A  G  I  B  I  G"
            'xlSheet.Cells(ctr, 6).VALUE = "Employee"
            xlSheet.Cells(ctr, 6).VALUE = "EMPLOYEE SHARE"
            'xlSheet.Cells(ctr - 1, 7).VALUE = "PAGIBIG"
            'xlSheet.Cells(ctr, 7).VALUE = "Employer"
            xlSheet.Cells(ctr, 7).VALUE = "EMPLOYER SHARE"
            'xlSheet.Cells(ctr - 1, 8).VALUE = "PAGIBIG"
            'xlSheet.Cells(ctr, 8).VALUE = "Total    "
            xlSheet.Cells(ctr, 8).VALUE = "TIN NO.   "
'            If ReportScope = 0 Then
'                xlSheet.Cells(ctr, 9).VALUE = "Start Pay"
'                xlSheet.Cells(ctr, 10).VALUE = "End Pay"
''            ElseIf ReportScope = 3 Then ' for annual
''                xlSheet.Cells(ctr, 9).VALUE = "Annual"
''                xlSheet.Cells(ctr, 10).VALUE = "Year"
'            Else
'                'xlSheet.Cells(ctr, 9).VALUE = "Month"
'                'xlSheet.Cells(ctr, 10).VALUE = "Year"
                xlSheet.Cells(ctr, 9).VALUE = "BIRTH DATE"
                xlSheet.Cells(ctr, 10).VALUE = "REMARKS   "
'            End If
            mRange = "A1:J10"
            mRange2 = "A6:J7"
            mRange3 = "F6:H6"
        End If
    Case 4  ' PhilHealth Updated 9/16/2018
        xlSheet.Cells(ctr, 5).VALUE = "Date of Birth"
        xlSheet.Cells(ctr, 6).VALUE = "PhilHealth No."
        xlSheet.Cells(ctr - 1, 7).VALUE = "CONTRIBUTIONS"
        xlSheet.Cells(ctr, 7).VALUE = "Employee"
        xlSheet.Cells(ctr, 8).VALUE = "Employer"
        xlSheet.Cells(ctr, 9).VALUE = "Total   "
        If ReportScope = 0 Then
            xlSheet.Cells(ctr, 10).VALUE = "Start Pay"
            xlSheet.Cells(ctr, 11).VALUE = "End Pay"
        Else
            'xlSheet.Cells(ctr, 10).VALUE = "Month"
            'xlSheet.Cells(ctr, 11).VALUE = "Year"
        End If
        mRange = "A1:K10"
        mRange2 = "A6:K7"
        mRange3 = "F6:H6"
    Case 5  ' 13th Month
        'xlSheet.Cells(ctr - 1, 5).VALUE = "13th Month"
        'xlSheet.Cells(ctr, 5).VALUE = "Payment     "
        'xlSheet.Cells(ctr, 6).VALUE = "Basic Pay"
        xlSheet.Cells(ctr, 5).VALUE = "No of Months"
        xlSheet.Cells(ctr, 6).VALUE = "13th Month Payment"
        'xlSheet.Cells(ctr, 7).VALUE = "Date Hired"
        mRange = "A1:F7"
        mRange2 = "A6:F7"
        mRange3 = "C6:C6"
    
        xlApp1.Sheets("Sheet2").Select
        xlApp1.ActiveSheet.Range("A1") = mSapi
        xlApp1.ActiveSheet.Range("A2") = ReportHeader
        'xlSheet.Range("A2") = ReportHeader
        xlApp1.ActiveSheet.Range("A3") = "For Bank Copy Report"
        'xlSheet.Range("A3") = "13th Month Pay Bank Report"
        
        xlApp1.ActiveSheet.Cells(ctr, 1).VALUE = "TARGET ACCOUNT"
        xlApp1.ActiveSheet.Cells(ctr, 2).VALUE = "AMOUNT"
        xlApp1.ActiveSheet.Cells(ctr, 3).VALUE = "E-MAIL (Optional)"
        xlApp1.ActiveSheet.Cells(ctr, 4).VALUE = "PARTICULARS (Optional)"
        
        xlApp1.Sheets("Sheet1").Select
    
    Case 6  ' Late and Absent
        xlSheet.Cells(ctr, 5).VALUE = "Late Hrs"
        xlSheet.Cells(ctr, 6).VALUE = "Late Pay"
        xlSheet.Cells(ctr, 7).VALUE = "Abs. Hrs"
        xlSheet.Cells(ctr, 8).VALUE = "Abs. Pay"
        xlSheet.Cells(ctr, 9).VALUE = "Month of"
        mRange = "A1:I7"
        mRange2 = "A6:I7"
        mRange3 = "C6:C6"
    Case 66  ' DTR Lates per Day
        xlSheet.Cells(ctr, 5).VALUE = "Late Hrs"
        xlSheet.Cells(ctr, 6).VALUE = "Date Late"
        xlSheet.Cells(ctr, 7).VALUE = "Day"
        mRange = "A1:G7"
        mRange2 = "A6:G7"
        mRange3 = "C6:C6"
    Case 0  ' Gross Summary
        xlSheet.Cells(ctr, 1).VALUE = "Employee #"
        xlSheet.Cells(ctr, 2).VALUE = "Employee Name   "
        
        xlSheet.Cells(ctr, 3).VALUE = "Gross Pay  "
        xlSheet.Cells(ctr, 6).VALUE = "13 Month Rate "
        If ReportScope = 0 Then
            xlSheet.Cells(ctr, 7).VALUE = "Start Pay"
            xlSheet.Cells(ctr, 8).VALUE = "End Pay"
        Else
            xlSheet.Cells(ctr - 1, 7).VALUE = "Date  "
            xlSheet.Cells(ctr, 7).VALUE = "From  "
            xlSheet.Cells(ctr - 1, 8).VALUE = "Date  "
            xlSheet.Cells(ctr, 8).VALUE = "To    "
        End If
        mRange = "A1:H7"
        mRange2 = "A6:H7"
        mRange3 = "C6:C6"
    
    Case 7  ' Head Count
        mRange2 = "A6:O7"
    Case 8  ' Other Deductions
        xlSheet.Cells(ctr - 1, 5).VALUE = "Calamity"
        xlSheet.Cells(ctr, 5).VALUE = " Loan "
        xlSheet.Cells(ctr - 1, 6).VALUE = "Salary "
        xlSheet.Cells(ctr, 6).VALUE = "Loan   "
        xlSheet.Cells(ctr - 1, 7).VALUE = " Uniform "
        xlSheet.Cells(ctr, 7).VALUE = " Charge "
'        xlSheet.Cells(ctr - 1, 8).VALUE = "Other Ded"
'        xlSheet.Cells(ctr, 8).VALUE = "Charge   "
        xlSheet.Cells(ctr - 1, 8).VALUE = "Planters"
        xlSheet.Cells(ctr, 8).VALUE = " Loan "
        
        xlSheet.Cells(ctr - 1, 9).VALUE = "  LMC   "
        xlSheet.Cells(ctr, 9).VALUE = "Deduction"
        xlSheet.Cells(ctr - 1, 10).VALUE = " VALE "
        xlSheet.Cells(ctr, 10).VALUE = "Deduction"
        xlSheet.Cells(ctr - 1, 11).VALUE = "  Cash  "
        xlSheet.Cells(ctr, 11).VALUE = " Advance "
        xlSheet.Cells(ctr - 1, 12).VALUE = " Union "
        xlSheet.Cells(ctr, 12).VALUE = " Dues  "
        
        ' Get Other Deductions as field names. Updated by Bernard June 29, 2013.
        If rsOtherDed2.State = 1 And rsOtherDed2.RecordCount > 0 Then
            With rsOtherDed2
                .MoveFirst
                Do While Not .EOF
                    jj = jj + 1
                    xlSheet.Cells(ctr, 12 + jj).VALUE = !VoucherNo
                    mdedname(0, jj) = !VoucherNo
                    .MoveNext
                Loop
            End With
        End If
        
        If ReportScope = 0 Then
            xlSheet.Cells(ctr, 13 + jj).VALUE = "Start Pay"
            xlSheet.Cells(ctr, 14 + jj).VALUE = "End Pay"
        Else
            xlSheet.Cells(ctr, 13 + jj).VALUE = "Month"
            xlSheet.Cells(ctr, 14 + jj).VALUE = "Year"
        End If
        xlSheet.Cells(ctr, 15 + jj).VALUE = "Remarks"
        mRange = "A1:" & Chr(79 + jj) & 15 + jj
        mRange2 = "A6:" & Chr(79 + jj) & "7"
        mRange3 = "B6:B6"
End Select
With xlSheet.Range(mRange2)
    .Font.Name = "Tahoma"   'Font.Name
    .Font.Size = 10         'Font.size
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
ctr = ctr + 1

rs.MoveFirst

prevClientCode = rs!ClientCode
prevControlNo = rs!EmpNo
If ReportScope = 0 Then ' per cut-off
    prevCutOff = rs!StartPaydate
'ElseIf ReportScope = 3 Then
'    prevCutOff = "Annual"
Else
    prevCutOff = rs!MonthOf
End If
StartOfLine = 8

xlSheet.Range("A5") = rs!ClientName
If ReportType = 1 And chkLoans.VALUE <> 1 Then ' SSS premium
    xlApp1.Sheets("Sheet2").Select
    xlApp1.ActiveSheet.Range("A5") = rs!ClientName
    xlApp1.Sheets("Sheet3").Select
    xlApp1.ActiveSheet.Range("A5") = rs!ClientName
End If
xlApp1.Sheets("Sheet1").Select

With xlSheet.Range("A5")
    .Font.Name = "Tahoma"  'Font.Name
    .Font.Size = 10 'Font.size
    .Font.Bold = True ' Font.Bold
    .Font.Italic = False 'font.Italic
End With
mRange = "A" & ctr + 2
RecNum = 0

Do While Not rs.EOF
NextRec:
    ' For Lates and absent.
    If ReportType = 6 Then
        If rs!DaysAbsent = 0 And rs!HoursTardiness = 0 Then
            rs.MoveNext
            If rs.EOF = True Then
                GoTo NextProcess
            End If
        GoTo NextRec
        End If
    End If

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
        SaveTotal = SaveTotal & ctr - 1 & ":"
        prevClientCode = rs!ClientCode
        StartOfLine = ctr + 2
        ctr = ctr + 2
        RecNum = 0
    End If
    RecNum = RecNum + 1
    'Check if there is a change in Control Number.
    If prevControlNo <> rs!EmpNo And ReportType = 1 And chkLoans.VALUE = 1 And ReportScope = 3 Then  ' Quarterly report Then
        ' Print sub-total of loan.
        Call ComputeTotals(xlSheet, xlApp1, ctr, i, StartOfLine, True)
        prevControlNo = rs!EmpNo
        ctr = ctr + 1
    End If
    
    DoEvents ' Updated 9/23/2018 caused by Error in subscript
    ' Updated May 10, 2013, for 13th month pay report adjustment
    If ReportType = 5 Then ' for 13th month pay
        xlSheet.Cells(ctr, 1).VALUE = RecNum
        xlSheet.Cells(ctr, 2).VALUE = Trim$(rs!Surname) & ", " & Trim$(rs!FirstName) & " " & IIf(rs!MiddleName = "", "", Left(rs!MiddleName, 1) & ".")
        xlSheet.Cells(ctr, 3).VALUE = IIf(IsNull(rs!datehired), "", Format(rs!datehired, "mm/dd/yyyy"))
        xlSheet.Cells(ctr, 4).VALUE = Format(rs!MonthlyRate, "###,#00.00")
    Else
        If ReportType = 1 Or ReportType = 4 Then ' New sss format
            xlSheet.Cells(ctr, 1).VALUE = rs!Surname
            xlSheet.Cells(ctr, 2).VALUE = rs!FirstName
            xlSheet.Cells(ctr, 3).VALUE = IIf(IsNull(rs!MiddleName) Or rs!MiddleName = "", "", rs!MiddleName)
            If ReportType = 4 Then ' if Philhealth
                xlSheet.Cells(ctr, 4).VALUE = LTrim(RTrim(cmbYearof.Text)) & Right("00" & rs!MonthNo, 2) 'getMonthNo(rs!MonthoF)
            Else
                xlSheet.Cells(ctr, 4).VALUE = IIf(rs!sss = "" Or IsNull(rs!sss), "00-0000000-0", rs!sss)
            End If
        Else
            xlSheet.Cells(ctr, 1).VALUE = IIf(rs!EmpNo = "" Or IsNull(rs!EmpNo), "00-0000", rs!EmpNo)
            xlSheet.Cells(ctr, 2).VALUE = rs!Surname
            xlSheet.Cells(ctr, 3).VALUE = rs!FirstName
            xlSheet.Cells(ctr, 4).VALUE = IIf(IsNull(rs!MiddleName) Or rs!MiddleName = "", "", rs!MiddleName)
        End If
    End If
    
    Select Case ReportType
        Case 1  ' SSS
            If chkLoans.VALUE = 1 Or chkCal.VALUE = 1 Then ' SSS Loan
                xlSheet.Cells(ctr, 5).VALUE = rs!sss
                If chkCal.VALUE = 1 Then ' if SSS calamity
                    xlSheet.Cells(ctr, 6).VALUE = Format(rs!Medical, "##,#00.00")
                    TotalLoan = TotalLoan + rs!Medical
                Else
                    xlSheet.Cells(ctr, 6).VALUE = Format(rs!SSSLoan, "##,#00.00")
                    TotalLoan = TotalLoan + rs!SSSLoan
                End If
                mRange = "F" & ctr & ":F" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                If ReportScope = 0 Then
                    xlSheet.Cells(ctr, 7).VALUE = rs!StartPaydate
                    xlSheet.Cells(ctr, 8).VALUE = rs!EndPaydate
                Else
                    xlSheet.Cells(ctr, 7).VALUE = rs!MonthOf
                    xlSheet.Cells(ctr, 8).VALUE = rs!YearOf 'cmbYearof.Text
                End If
                mRange = "A" & StartOfLine & ":H" & i + 7
                mRange2 = "B7:H" & i + 7
            Else
                If ReportType = 1 Then 'New sss format, 9/15/2018
                    xlSheet.Cells(ctr, 4).VALUE = Format(IIf(IsNull(rs!BIRTHDATE), "01/01/1900", rs!BIRTHDATE), "mm/dd/yyyy")
                    xlSheet.Cells(ctr, 5).VALUE = IIf(IsNull(rs!sss), "No SSS #", rs!sss)
                    xlSheet.Cells(ctr, 6).VALUE = Format(rs!SSSEmployee, "##,#00.00")
                    xlSheet.Cells(ctr, 7).VALUE = Format(rs!SSSEmployer - rs!ECCEmployer, "##,#00.00")
                    xlSheet.Cells(ctr, 8).VALUE = Format(rs!SSSEmployee + rs!SSSEmployer - rs!ECCEmployer, "##,#00.00")
                    xlSheet.Cells(ctr, 9).VALUE = Format(rs!ECCEmployer, "##,#00.00")
                    xlSheet.Cells(ctr, 10).VALUE = Format(rs!SSSEmployee + rs!SSSEmployer, "##,#00.00")
                    mRange = "F" & ctr & ":J" & ctr
                    xlSheet.Range(mRange).Select
                    xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                Else
                    xlSheet.Cells(ctr, 5).VALUE = IIf(IsNull(rs!sss), "No SSS #", rs!sss)
                    xlSheet.Cells(ctr, 6).VALUE = Format(rs!SSSEmployee, "##,#00.00")
                    xlSheet.Cells(ctr, 7).VALUE = Format(rs!SSSEmployer, "##,#00.00")
                    xlSheet.Cells(ctr, 8).VALUE = Format(rs!SSSEmployee + rs!SSSEmployer, "##,#00.00")
                    mRange = "F" & ctr & ":H" & ctr
                    xlSheet.Range(mRange).Select
                    xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                End If
                If ReportScope = 0 Then
                    xlSheet.Cells(ctr, 9).VALUE = rs!StartPaydate
                    xlSheet.Cells(ctr, 10).VALUE = rs!EndPaydate
                Else
                    If ReportScope = 1 Then 'New SSS Format
                        'xlSheet.Cells(ctr, 9).VALUE = Format(rs!ECCEmployer, "##,#00.00")
                        'xlSheet.Cells(ctr, 10).VALUE = Format(rs!SSSEmployee + rs!SSSEmployer, "##,#00.00")
                    Else
                        xlSheet.Cells(ctr, 9).VALUE = rs!MonthOf
                        xlSheet.Cells(ctr, 10).VALUE = rs!YearOf 'cmbYearof.Text
                    End If
                End If
                mRange = "A" & ctr & ":E" & ctr
                xlSheet.Range(mRange).Select
                xlApp1.Selection.Copy
                Call CopyToOtherSheet(xlSheet, xlApp1, ctr, IIf(IsNull(rs!datehired), "__/__/____", rs!datehired), rs!SSSEmployee + rs!SSSEmployer - rs!ECCEmployer, rs!ECCEmployer, "MCL", "Sheet2")
                xlSheet.Range(mRange).Select
                xlApp1.Selection.Copy
                'Call CopyToOtherSheet(xlSheet, xlApp1, ctr, IIf(IsNull(rs!BirthDate), "__/__/____", rs!BirthDate), rs!SSSEmployee + rs!SSSEmployer - rs!ECCEmployer, rs!ECCEmployer, "EPF", "Sheet3")
                Call CopyToOtherSheet(xlSheet, xlApp1, ctr, IIf(IsNull(rs!BIRTHDATE), "01/01/1900", rs!BIRTHDATE), rs!SSSEmployee + rs!SSSEmployer - rs!ECCEmployer, rs!ECCEmployer, "EPF", "Sheet3")
                mRange = "A" & StartOfLine & ":J" & i + 7
                mRange2 = "B7:J" & i + 7
            End If
        Case 2, 9  ' Alphalist
            If OptReport(9).VALUE = True Then
                xlSheet.Cells(ctr, 5).VALUE = rs!TIN
                xlSheet.Cells(ctr, 6).VALUE = rs!TaxCode
                xlSheet.Cells(ctr, 7).VALUE = Format(rs!GrossPay, "##,#00.00")
                xlSheet.Cells(ctr, 8).VALUE = Format(rs!Exempt, "##,#00.00")
                xlSheet.Cells(ctr, 9).VALUE = Format(rs!SSSEmployee, "##,#00.00")
                xlSheet.Cells(ctr, 10).VALUE = Format(rs!pagibigemployee, "##,#00.00")
                xlSheet.Cells(ctr, 11).VALUE = Format(rs!PHealthEmployee, "##,#00.00")
                xlSheet.Cells(ctr, 12).VALUE = Format(rs!WTaxAmount, "##,#00.00")
                xlSheet.Cells(ctr, 13).VALUE = Format(rs!GrossPay - (rs!Exempt + rs!SSSEmployee + rs!pagibigemployee + rs!PHealthEmployee), "###,#00.00")
                xlSheet.Cells(ctr, 14).VALUE = Format(ComputeAnnualTax(rs!GrossPay, rsTaxTable, (rs!Exempt + rs!SSSEmployee + rs!pagibigemployee + rs!PHealthEmployee)), "###,#00.00")
                xlSheet.Cells(ctr, 15).VALUE = Format(Get13thMonthPay(rs13thMonth, rs!EmployeeID), "###,#00.00")
                mRange = "I" & ctr & ":N" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                mRange = "A" & StartOfLine & ":O" & i + 7
                mRange2 = "B7:I" & i + 7
            Else    ' Tax
                xlSheet.Cells(ctr, 5).VALUE = rs!TIN
                xlSheet.Cells(ctr, 6).VALUE = Format(rs!WTaxAmount, "##,#00.00")
                xlSheet.Cells(ctr, 7).VALUE = Format(rs!GrossPay, "##,#00.00")
                mRange = "F" & ctr & ":G" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                xlSheet.Cells(ctr, 8).VALUE = rs!MonthOf
                xlSheet.Cells(ctr, 9).VALUE = rs!YearOf 'cmbYearof.Text
                mRange = "A" & StartOfLine & ":I" & i + 7
                mRange2 = "B7:I" & i + 7
            End If
        Case 3  ' HDMF
            If chkLoans.VALUE = 1 Then ' SSS Loan
                xlSheet.Cells(ctr, 5).VALUE = IIf(IsNull(rs!Pagibig), Format(rs!BIRTHDATE, "mm/dd/yyyy"), rs!Pagibig)
                xlSheet.Cells(ctr, 6).VALUE = Format(rs!PagibigLoan, "##,#00.00")
                mRange = "F" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                If ReportScope = 0 Then
                    xlSheet.Cells(ctr, 7).VALUE = rs!StartPaydate
                    xlSheet.Cells(ctr, 8).VALUE = rs!EndPaydate
'                ElseIf ReportScope = 3 Then ' for annual
'                    xlSheet.Cells(ctr, 7).VALUE = "Annual"
'                    xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
                Else
                    xlSheet.Cells(ctr, 7).VALUE = rs!MonthOf
                    xlSheet.Cells(ctr, 8).VALUE = rs!YearOf 'cmbYearof.Text
                End If
                mRange = "A" & StartOfLine & ":H" & i + 7
                mRange2 = "B7:H" & i + 7
            ElseIf chkCal.VALUE = 1 Then ' calamity Loan
                xlSheet.Cells(ctr, 5).VALUE = IIf(IsNull(rs!Pagibig), Format(rs!BIRTHDATE, "mm/dd/yyyy"), rs!Pagibig)
                xlSheet.Cells(ctr, 6).VALUE = Format(rs!PagibigLoan, "##,#00.00")
                mRange = "F" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                If ReportScope = 0 Then
                    xlSheet.Cells(ctr, 7).VALUE = rs!StartPaydate
                    xlSheet.Cells(ctr, 8).VALUE = rs!EndPaydate
'                ElseIf ReportScope = 3 Then ' for annual
'                    xlSheet.Cells(ctr, 7).VALUE = "Annual"
'                    xlSheet.Cells(ctr, 8).VALUE = cmbYearof.Text
                Else
                    xlSheet.Cells(ctr, 7).VALUE = rs!MonthOf
                    xlSheet.Cells(ctr, 8).VALUE = rs!YearOf 'cmbYearof.Text
                End If
                mRange = "A" & StartOfLine & ":H" & i + 7
                mRange2 = "B7:H" & i + 7
            Else
                'xlSheet.Cells(ctr, 5).VALUE = IIf(IsNull(rs!pagibig), Format(rs!BirthDate, "mm/dd/yyyy"), rs!pagibig)
                xlSheet.Cells(ctr, 1).VALUE = IIf(IsNull(rs!Pagibig), Format(rs!BIRTHDATE, "mm/dd/yyyy"), rs!Pagibig)
                xlSheet.Cells(ctr, 5).VALUE = GetMonthName(rs!MonthNo) & " " & rs!YearOf 'cmbYearof.Text
                mRange = "E" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "yyyymm"  ' or mmm-yyyy
                'xlSheet.Cells(ctr, 6).VALUE = Format(rs!pagibigemployee, "##,#00.00")
                xlSheet.Cells(ctr, 6).VALUE = Format(rs!pagibigemployee, "##,#00.00")
                'xlSheet.Cells(ctr, 7).VALUE = Format(rs!PagibigEmployer, "##,#00.00")
                xlSheet.Cells(ctr, 7).VALUE = Format(rs!PagibigEmployer, "##,#00.00")
                'xlSheet.Cells(ctr, 8).VALUE = Format(rs!pagibigemployee + rs!PagibigEmployer, "##,#00.00")
                xlSheet.Cells(ctr, 8).VALUE = IIf(IsNull(rs!TIN), "", rs!TIN)
                xlSheet.Cells(ctr, 9).VALUE = IIf(IsNull(rs!BIRTHDATE), "__/__/____", Format(rs!BIRTHDATE, "mm/dd/yyyy"))
                mRange = "F" & ctr & ":G" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'                If ReportScope = 0 Then
'                    xlSheet.Cells(ctr, 9).VALUE = rs!StartPayDate
'                    xlSheet.Cells(ctr, 10).VALUE = rs!EndPayDate
''                ElseIf ReportScope = 3 Then ' for annual
''                    xlSheet.Cells(ctr, 9).VALUE = "Annual"
''                    xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
'                Else
'                    xlSheet.Cells(ctr, 9).VALUE = rs!MonthoF
'                    xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
'                End If
                mRange = "A" & StartOfLine & ":J" & i + 7
                mRange2 = "B7:J" & i + 7
            End If
        Case 4  ' PhilHealth
            'xlSheet.Cells(ctr, 5).VALUE = LTrim(RTrim(cmbYearof.Text)) & Right("00" & rs!MonthNo, 2) 'getMonthNo(rs!MonthoF)
            xlSheet.Cells(ctr, 5).VALUE = IIf(IsNull(rs!BIRTHDATE), "__/__/____", Format(rs!BIRTHDATE, "mm/dd/yyyy"))
            xlSheet.Cells(ctr, 6).VALUE = rs!PhilHealth
            xlSheet.Cells(ctr, 7).VALUE = Format(rs!PHealthEmployee, "##,#00.00")
            xlSheet.Cells(ctr, 8).VALUE = Format(rs!PHealthEmployer, "##,#00.00")
            xlSheet.Cells(ctr, 9).VALUE = Format(rs!PHealthEmployee + rs!PHealthEmployer, "##,#00.00")
            mRange = "G" & ctr & ":I" & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 10).VALUE = rs!StartPaydate
                xlSheet.Cells(ctr, 11).VALUE = rs!EndPaydate
'            ElseIf ReportScope = 3 Then ' for annual
'                xlSheet.Cells(ctr, 9).VALUE = "Annual"
'                xlSheet.Cells(ctr, 10).VALUE = cmbYearof.Text
'            Else
'                xlSheet.Cells(ctr, 10).VALUE = rs!MonthOf
'                xlSheet.Cells(ctr, 11).VALUE = rs!YearOf 'cmbYearof.Text
            End If
            mRange = "A" & StartOfLine & ":K" & i + 7
            mRange2 = "B7:K" & i + 7
        
        Case 5  ' 13th Month, Updated May 3, 2013
            'xlSheet.Cells(ctr, 6).VALUE = rs!MonthlyRate
            xlSheet.Cells(ctr, 5).VALUE = rs!NoofMonths
            xlSheet.Cells(ctr, 6).VALUE = Format(rs!Month13Payment, "##,#00.00")
            mRange = "D" & ctr & ":F" & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            'xlSheet.Cells(ctr, 8).VALUE = Trim$(rs!Surname) & ", " & Trim$(rs!FirstName) & " " & IIf(rs!Middlename = "", "", Left(rs!Middlename, 1) & ".")
            'xlSheet.Cells(ctr, 9).VALUE = IIf(IsNull(rs!datehired), "", Format(rs!datehired, "mm/dd/yyyy"))
            mRange = "A" & StartOfLine & ":F" & i + 7
            mRange2 = "B7:F" & i + 7
        
            xlApp1.Sheets("Sheet2").Select
            xlApp1.ActiveSheet.Cells(ctr, 1).VALUE = rs!AccountNo
            xlApp1.ActiveSheet.Cells(ctr, 2).VALUE = Format(rs!Month13Payment, "##,#00.00")
            xlApp1.ActiveSheet.Cells(ctr, 3).VALUE = "tastranspayroll@ yahoo.com"
            xlApp1.ActiveSheet.Cells(ctr, 4).VALUE = Trim$(rs!Surname) & ", " & Trim$(rs!FirstName) & " " & IIf(rs!MiddleName = "", "", Left(rs!MiddleName, 1) & ".")
            mRange = "B" & ctr & ":B" & ctr
            xlApp1.ActiveSheet.Range(mRange).Select
            xlApp1.ActiveSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            mRange11 = "A" & StartOfLine & ":D" & i + 8
            mRange12 = "B8:B" & i + 7
            xlApp1.Sheets("Sheet1").Select
        
        Case 6  ' Late and Absent
            'DaysAbsent, AbsentDeduction, HoursTardiness, TardinessDeduction,
            xlSheet.Cells(ctr, 5).VALUE = Format(rs!HoursTardiness, "##,#00.00")
            xlSheet.Cells(ctr, 6).VALUE = Format(rs!TardinessDeduction, "##,#00.00")
            xlSheet.Cells(ctr, 7).VALUE = Format(rs!DaysAbsent, "0.00")
            xlSheet.Cells(ctr, 8).VALUE = Format(rs!AbsentDeduction, "##,#00.00")
            xlSheet.Cells(ctr, 9).VALUE = rs!MonthOf
            mRange = "E" & ctr & ":H" & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            mRange = "A" & StartOfLine & ":I" & i + 7
            mRange2 = "B7:I" & i + 7
            EndRange = "H"
        Case 66  ' DTR Lates per Day
            'DaysAbsent, AbsentDeduction, HoursTardiness, TardinessDeduction,
            xlSheet.Cells(ctr, 5).VALUE = Format(rs!HoursTardiness, "##,#00.00")
            xlSheet.Cells(ctr, 6).VALUE = Format(rs!StartPaydate, "mm/dd/yyyy")
            xlSheet.Cells(ctr, 7).VALUE = rs!MonthOf
            mRange = "E" & ctr & ":E" & ctr
            xlSheet.Range(mRange).Select
            'xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            mRange = "A" & StartOfLine & ":G" & i + 7
            mRange2 = "B7:G" & i + 7
            EndRange = "H"
        
        
        Case 7  ' Head Count
        Case 8  ' Other Deductions.
            xlSheet.Cells(ctr, 5).VALUE = Format(rs!Medical, "##,#00.00")
            xlSheet.Cells(ctr, 6).VALUE = Format(rs!SalaryLoan, "##,#00.00")
            xlSheet.Cells(ctr, 7).VALUE = Format(rs!Uniform, "##,#00.00")
            'xlSheet.Cells(ctr, 8).VALUE = Format(rs!TASAesla, "##,#00.00") ' OTher Ded Charge
            mRemarks = ""
'            ' Updated by Bernard july 30, 2012. For Other Deductions.
'            If rsOtherDed.State = 1 And rsOtherDed.RecordCount > 0 Then
'                rsOtherDed.Filter = ""
'                rsOtherDed.MoveFirst
'                rsOtherDed.Filter = "EmpNo = '" & rs!EmpNo & "'"
'                If rsOtherDed.RecordCount > 0 Then
'                    mRemarks = "Other Deductions:" & Chr(10) ' vbCrLf
'                End If
'                Do While Not rsOtherDed.EOF
'                    mRemarks = mRemarks & Format(rsOtherDed!DeductionPayment, "##,#00.00") & "-" & IIf(IsNull(rsOtherDed!Remarks), "Other Deductions", rsOtherDed!Remarks) & Chr(10) 'vbCrLf   ' remarks
'                    rsOtherDed.MoveNext
'                Loop
'                xlSheet.Cells(ctr, 16).VALUE = mRemarks
'                rsOtherDed.Filter = ""
'                rsOtherDed.MoveFirst
'            End If
            
            xlSheet.Cells(ctr, 8).VALUE = Format(rs!Vale_IOU, "##,#00.00")   ' Planters loan
            xlSheet.Cells(ctr, 9).VALUE = Format(rs!LMCDed, "##,#00.00")        ' LMC Ded
            xlSheet.Cells(ctr, 10).VALUE = Format(rs!ValeWithClient, "##,#00.00")
            xlSheet.Cells(ctr, 11).VALUE = Format(rs!CashAdvance, "##,#00.00")
            xlSheet.Cells(ctr, 12).VALUE = Format(rs!UnionDues, "##,#00.00")
            
            ' Updated by Bernard july 30, 2012. For Other Deductions.
            If rsOtherDed.State = 1 And rsOtherDed.RecordCount > 0 Then
                With rsOtherDed
                    .Filter = ""
                    .MoveFirst
                    .Filter = "EmpNo = '" & rs!EmpNo & "'"
                    If .RecordCount > 0 Then
                        mRemarks = "Other Deductions:" & Chr(10) ' vbCrLf
                        prevVoucher = !VoucherNo
                    Else
                        xlSheet.Cells(ctr, 12 + 1).VALUE = 0#
                        prevVoucher = ""
                    End If
                    DedTotals = 0
                    Do While Not .EOF
                        If !PaymentsMade <> 0 Then
                            mRemarks = mRemarks & Format(rsOtherDed!DeductionPayment, "##,#00.00") & "-" & IIf(IsNull(rsOtherDed!Remarks), "Other Deductions", rsOtherDed!Remarks) & Chr(10) 'vbCrLf   ' remarks
                            xlSheet.Cells(ctr, GetDedColumn(12, !VoucherNo, jj)).VALUE = Val(xlSheet.Cells(ctr, GetDedColumn(12, !VoucherNo, jj)).VALUE) + Format(!DeductionPayment, "##,#00.00")
                        End If
                        .MoveNext
                    Loop
                    xlSheet.Cells(ctr, 15 + jj).VALUE = mRemarks
                    .Filter = ""
                    .MoveFirst
                End With
            End If
            mRange = "E" & ctr & ":" & Chr(76 + jj) & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            If ReportScope = 0 Then
                xlSheet.Cells(ctr, 13 + jj).VALUE = rs!StartPaydate
                xlSheet.Cells(ctr, 14 + jj).VALUE = rs!EndPaydate
            Else
                xlSheet.Cells(ctr, 13 + jj).VALUE = rs!MonthOf
                xlSheet.Cells(ctr, 14 + jj).VALUE = rs!YearOf 'cmbYearof.Text
            End If
            mRange = "A" & StartOfLine & ":" & Chr(79 + jj) & i + 7   'L
            mRange2 = "B7:" & Chr(79 + jj) & i + 7
    End Select
    'mGrandPay = mGrandPay + (mTotalPay - mTotalDeduction)
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext
Loop

NextProcess:

' Print the last sub-total.
If ReportType = 1 And chkLoans.VALUE = 1 And ReportScope = 3 Then  ' Quarterly report Then
    Call ComputeTotals(xlSheet, xlApp1, ctr, i, StartOfLine, True)
    ctr = ctr + 1
End If

Call ComputeTotals(xlSheet, xlApp1, ctr, i, StartOfLine, , jj)

' Updated August 20, 2011. Bernard.
Dim mGTotal As Double, aa As Double, lf As Double

sumString = ""
If Len(SaveTotal) = 0 Then
    SaveTotal = ctr + 1 & ":"
Else
    SaveTotal = SaveTotal & ctr & ":"
End If
ColString = SaveTotal
ctr = ctr + 3

' Display date and time printed.
xlSheet.Range("A" & ctr).Select
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
    'Resume Next ' bernard 2019
    
End Sub

Private Function GetDedColumn(StartPos As Integer, mDed As String, ii As Integer) As Integer
Dim kk

GetDedColumn = StartPos
For kk = 1 To ii
    If Trim$(mdedname(0, kk)) = Trim$(mDed) Then
        GetDedColumn = StartPos + kk
        GoTo ExitDed
    End If
Next kk

ExitDed:

End Function

Private Function GetMonthName(MonthNo As Integer) As String

Select Case MonthNo
    Case 1:
        GetMonthName = "January"
    Case 2:
        GetMonthName = "February"
    Case 3:
        GetMonthName = "March"
    Case 4:
        GetMonthName = "April"
    Case 5:
        GetMonthName = "May"
    Case 6:
        GetMonthName = "June"
    Case 7:
        GetMonthName = "July"
    Case 8:
        GetMonthName = "August"
    Case 9:
        GetMonthName = "September"
    Case 10:
        GetMonthName = "October"
    Case 11:
        GetMonthName = "November"
    Case Else
        GetMonthName = "December"
End Select
End Function


Private Sub PrintGrossSummaryReport(ByRef rs As ADODB.Recordset, ByVal mMonth As String, ByVal mYear As Integer)
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
If cmbClient.Text = "--ALL CLIENTS--" Then ' if all clients, do not proceed
    mHead = CompanyName & " (" & CompanyShort & ")"
Else
    mHead = cmbClient.Text
End If
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 12, True)
mHead = "Gross Summary Report from " & Format(dtStartDate.VALUE, "mm/dd/yyyy") & " to " & Format(dtEndDate.VALUE, "mm/dd/yyyy")
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)

Call PrintReportHeader(xlSheet, xlApp1, "A5:A6", "ID No.", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "B5:B6", "Employee Name", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 40)
Call PrintReportHeader(xlSheet, xlApp1, "C5:C6", "Total Gross Pay", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "D5:D6", "13th Month", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "E5:E6", "Remarks", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 10)
rs.MoveFirst

Do While Not rs.EOF
    ' DoEvents
    xlSheet.Cells(ctr, 1).VALUE = rs!EmpNo
    xlSheet.Cells(ctr, 2).VALUE = rs!EmpName
    xlSheet.Cells(ctr, 3).VALUE = rs!GPay
    xlSheet.Cells(ctr, 4).VALUE = rs!m13Rate
    xlSheet.Cells(ctr, 5).VALUE = "___________"
    'xlSheet.Cells(ctr, 5).FormulaR1C1 = "=RC[-2]-RC[-1]" '"=RC[-2]+RC[-1]"
    ctr = ctr + 1
    i = i + 1
    rs.MoveNext
Loop
ctr = ctr + 3
xlSheet.Cells(ctr, 2).VALUE = "Grand Total"
  
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

'ctr = ctr + 7
'xlSheet.Cells(ctr, 2).VALUE = "Prepared by: ______________"
'xlSheet.Cells(ctr + 3, 2).VALUE = "Note by: ______________"
xlSheet.Range("A1").Select

End Sub

'Private Sub PrintChecklistReport(ByRef rs As ADODB.Recordset, ByVal mMonth As String, ByVal mYear As Integer)
'Dim rsDedRecords As ADODB.Recordset
'Dim xlApp1 As Excel.Application
'Dim xlBook1 As Excel.Workbook
'Dim xlSheet As Excel.Worksheet
'Dim mTotalPay As Double
'Dim mTotalDeduction As Double
'Dim mRange As String, mRange2 As String
'Dim sumString As String
'Dim ctr As Long
'Dim i As Integer
'Dim cPos As Integer
'Dim LineCtr As Integer
'Dim mHead As String
'Dim mDeduc As Double
'
'' Get all deduction records for specific client.
''strQuery = "SELECT * FROM qryDeductionRecords WHERE ClientCode = " & mClientCode & " order by clientCode"
''Set rsDedRecords = ReturnRst(strQuery, cnnStr, mstatic)
''If rsDedRecords.EOF = True Then
''    MsgBox "No payroll deduction records selected! Cannot proceed.", vbInformation
''    Set rsDedRecords = Nothing
''    GoTo ExitRtn2
''End If
'
'Set xlApp1 = New Excel.Application
'Set xlBook1 = xlApp1.Workbooks.Add()
'
'xlApp1.WindowState = xlMaximized
'xlApp1.Visible = True
'xlApp1.DisplayAlerts = False
'
'Set xlSheet = xlBook1.Worksheets(1)
'
'ctr = 8
'i = 1
'rs.MoveFirst
'
''Updated by Bernard May 13, 2008.
'mHead = CompanyName & " (" & CompanyShort & ")"
'Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 10, True)
'mHead = "PAYROLL CHECKLIST FOR THE MONTH OF " & UCase(mMonth) & " " & mYear & " EVERY " & rs!CutOffPeriod
'Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
'mHead = "PAYROLL RELEASE DATE EVERY " & rs!SalaryRelease & " OF " & UCase(mMonth) & " " & mYear
'Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)
'mHead = "COMPANY NAME: " & cmbClient.Text
'Call PrintReportHeader(xlSheet, xlApp1, "A5", mHead, "", xlLeft, , False, False, False, , , 10, True)
'
'Call PrintReportHeader(xlSheet, xlApp1, "A6:A7", "No.", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 3.14)
'Call PrintReportHeader(xlSheet, xlApp1, "B6:B7", "Name", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 22)
'Call PrintReportHeader(xlSheet, xlApp1, "C6:C7", "Daily", "Rate", xlCenter, xlBottom, True, True, True, 15, , 9, False, 4.3)
'Call PrintReportHeader(xlSheet, xlApp1, "D6:D7", "Staff", "Status", xlCenter, xlBottom, True, True, True, 15, , 9, False, 6)
'Call PrintReportHeader(xlSheet, xlApp1, "E6:H6", "Original DTR", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)
'Call PrintReportHeader(xlSheet, xlApp1, "E7:E7", "Bandy Clock", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 4.7)
'Call PrintReportHeader(xlSheet, xlApp1, "F7:F7", "Chrono Type", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 5)
'Call PrintReportHeader(xlSheet, xlApp1, "G7:G7", "Hand Written", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 4.85)
'Call PrintReportHeader(xlSheet, xlApp1, "H7:H7", "Xerox", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 4.56)
'Call PrintReportHeader(xlSheet, xlApp1, "I6:Q6", "DEDUCTIONS", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)
'Call PrintReportHeader(xlSheet, xlApp1, "I7:I7", "C.B. Balance", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.71)
'Call PrintReportHeader(xlSheet, xlApp1, "J7:J7", "P/F Balance", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.71)
'Call PrintReportHeader(xlSheet, xlApp1, "K7:K7", "M/F Balance", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.71)
'Call PrintReportHeader(xlSheet, xlApp1, "L7:L7", "Trans. Fee", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 4.57)
'Call PrintReportHeader(xlSheet, xlApp1, "M7:M7", "VALE C/A", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5)
'Call PrintReportHeader(xlSheet, xlApp1, "N7:N7", "SLOAN Amount", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
'Call PrintReportHeader(xlSheet, xlApp1, "O7:O7", "Variance w/ Client", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 6.57)
'Call PrintReportHeader(xlSheet, xlApp1, "P7:P7", "Medical Ded.", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
'Call PrintReportHeader(xlSheet, xlApp1, "Q7:Q7", "Uniform Ded.", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
'Call PrintReportHeader(xlSheet, xlApp1, "R6:R7", "Pag-ibig", "Loan", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
'Call PrintReportHeader(xlSheet, xlApp1, "S6:S7", "SSS", "Loan", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
''Call PrintReportHeader(xlSheet, xlApp1, "T6:T7", "Daily", "Rate", xlCenter, xlBottom, True, True, True, 15, , 9, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "T6:T7", "Start Contract", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 7.72)
'Call PrintReportHeader(xlSheet, xlApp1, "U6:U7", "End Contract", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 7.72)
'Call PrintReportHeader(xlSheet, xlApp1, "V6:V7", "Account Number", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 14.14)
'Call PrintReportHeader(xlSheet, xlApp1, "W6:W7", "Remarks", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 12)
'rs.MoveFirst
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
'ctr = ctr + 3
'
''Box it.
'mRange = "A6:W" & ctr + 3
'Call PlaceBox(xlSheet, xlApp1, mRange)
'
'ctr = ctr + 4
'mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
'mRange = "B" & ctr
'Call CustomizedHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 8, False, 25)
'
'ctr = ctr + 3
'xlSheet.Cells(ctr, 2).VALUE = "Prepared by: ______________"
'xlSheet.Cells(ctr, 7).VALUE = "Note by: ______________"
'xlSheet.Cells(ctr, 11).VALUE = "Received by: ______________"
'xlSheet.Range("A1").Select
'
'ExitRtn2:
'
'End Sub

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
Dim sumString As String, ExcelTemplate As String
Dim ctr As Long, ctr2 As Long
Dim i As Integer, j As Integer, k As Integer
Dim cPos As Integer
Dim EndSw As Boolean
Dim mHead As String

Dim LineCtr As Integer
Dim mSeriesNo As String
Dim prevClientName As String, pClient As String
Dim SaveTotal As String, ColString As String
Dim firstRun As Boolean
Dim lf, aa, mTemplate As String
Dim prevStartCutoff As Variant
Dim prevEndCutoff As Variant
Dim WConsolidated As Boolean

'Set xlApp1 = New Excel.Application
'Set xlBook1 = xlApp1.Workbooks.Add()

'xlApp1.WindowState = xlMaximized
'xlApp1.Visible = True
'xlApp1.DisplayAlerts = False
'
''Set xlSheet = book1.Worksheets.Add
'Set xlSheet = xlBook1.Worksheets(1)
'xlSheet.Name = "Payroll Summary"

'-----------
'If CutOffPeriod = "START" Then
mTemplate = ".xls"
WConsolidated = False

If OptScope(3).VALUE = True Then
    mTemplate = " 4.xls"
    'mTemplate = " 3.xls"
    WConsolidated = True
Else
    mTemplate = " 2.xls"
End If

If OptScope(2).VALUE = True Then
    If InStr(1, cmbClient.Text, "SAINT ANTHONY", vbTextCompare) <> 0 Then
        ExcelTemplate = App.Path & "\Templates\Payroll Summary Saint Anthony Template" & mTemplate
    ElseIf InStr(1, cmbClient.Text, "TAS TRANSPORT", vbTextCompare) <> 0 Then
        ExcelTemplate = App.Path & "\Templates\Payroll Summary Tastrans Template" & mTemplate
    ElseIf InStr(1, cmbClient.Text, "SAN AGUSTIN", vbTextCompare) <> 0 Then
        ExcelTemplate = App.Path & "\Templates\Payroll Summary San Agustin Template" & mTemplate
    Else
        ExcelTemplate = App.Path & "\Templates\Payroll Summary Bensan Template" & mTemplate
    End If
ElseIf OptScope(3).VALUE = True Then    ' with consolidation, updated may 31, 2015
    ExcelTemplate = App.Path & "\Templates\Payroll Summary Template" & mTemplate
Else
    ExcelTemplate = App.Path & "\Templates\Payroll Summary Template.xls"
End If

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add(ExcelTemplate)

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)

'--------

mSeriesNo = "45"
SaveTotal = ""
firstRun = True

cPos = 5
ctr = 7: ctr2 = 7
LineCtr = 7
i = 1
k = 7

If OptScope(3).VALUE = True Then    ' for Updated Summary June 13, 2015
    xlApp1.Sheets("Payroll Summary").Select
    mHead = Replace(cmbClient.Text, "ADMIN", "") & " (" & CompanyShort & ")"
    Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 14, True)
    xlApp1.Sheets("Grand Total").Select
    Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 14, True)
    xlApp1.Sheets("Consolidation").Select
    Call PrintReportHeader(xlSheet, xlApp1, "B1", mHead, "", xlLeft, , False, False, False, , , 14, True)
    xlApp1.Sheets("Payroll Summary").Select
Else
    'Updated by Bernard May 13, 2008.
    mHead = CompanyName & " (" & CompanyShort & ")"
    Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 14, True)
End If
If CutOffPeriod = "START" Then
    mHead = "Payroll Summary Report for the Month of " & Format(rs!StartPaydate, "mmmm") & " " & Format(rs!StartPaydate, "yyyy")
Else
    mHead = "Payroll Summary Report for the Month of " & Format(rs!EntryDate, "mmmm") & " " & Format(rs!EntryDate, "yyyy")
End If

Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
Call PrintReportHeader(xlSheet, xlApp1, "A5:A6", "Client Name", "", xlCenter, xlBottom, True, False, True, 15, , 8, False, 30)
Call PrintReportHeader(xlSheet, xlApp1, "B5:B6", "Start Cut-Off", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "C5:C6", "End Cut-Off", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 10)

Call PrintReportHeader(xlSheet, xlApp1, "D5:D6", "DTR ID", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 13)
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

'Call PrintReportHeader(xlSheet, xlApp1, "Q5:Q6", "Tax", "Refund", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3) 'Updated April 13, 2015
Call PrintReportHeader(xlSheet, xlApp1, "Q5:Q6", "Vacation", "Leave", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3) 'Updated April 13, 2015
Call PrintReportHeader(xlSheet, xlApp1, "R5:R6", "Sick", "Leave", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "S5:S6", "Sunday OT", "Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "T5:T6", "Sunday OT", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)

Call PrintReportHeader(xlSheet, xlApp1, "U5:U6", "Lgl Hol.", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "V5:V6", "Lgl Hol.", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "W5:W6", "Spl Hol.", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "X5:X6", "Spl Hol.", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "Y5:Y6", "Other", "Income", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "Z5:Z6", "Allowance", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)

Call PrintReportHeader(xlSheet, xlApp1, "AA5:AA6", "ECOLA", "Payment", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9.5)
Call PrintReportHeader(xlSheet, xlApp1, "AB5:AB6", "GROSS PAY", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 9)
Call PrintReportHeader(xlSheet, xlApp1, "AC5:AC6", "WTAX", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AD5:AF5", "Employee", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)
Call PrintReportHeader(xlSheet, xlApp1, "AD6:AD6", "SSS", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AE6:AE6", "PhilHealth", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AF6:AF6", "Pag-ibig", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AG5:AG6", "HDMF", "Calamity", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.86)
Call PrintReportHeader(xlSheet, xlApp1, "AH5:AH6", "Cash", "Advance", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.86)
Call PrintReportHeader(xlSheet, xlApp1, "AI5:AI6", "Misc.", "Fee", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
Call PrintReportHeader(xlSheet, xlApp1, "AJ5:AJ6", "Association", "/SANAWU", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AK5:AK6", "Salary", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AL5:AL6", "Planters", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.87)
Call PrintReportHeader(xlSheet, xlApp1, "AM5:AM6", "Uniform", "Charges", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AN5:AN6", "Other", "Deduction", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)

Call PrintReportHeader(xlSheet, xlApp1, "AO5:AO6", "Variances", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)

Call PrintReportHeader(xlSheet, xlApp1, "AP5:AP6", "SSS", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AQ5:AQ6", "PBIG", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.8)
Call PrintReportHeader(xlSheet, xlApp1, "AR5:AR6", "Vale", "w/ Client", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.8)
Call PrintReportHeader(xlSheet, xlApp1, "AS5:AS6", "Union", "Dues", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
Call PrintReportHeader(xlSheet, xlApp1, "AT5:AT6", "Hazard", "Pay", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
Call PrintReportHeader(xlSheet, xlApp1, "AU5:AU6", "Salary", "Adjustment", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)

Call PrintReportHeader(xlSheet, xlApp1, "AV5:AV6", "TOTAL", "Deductions", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "AW5:AW6", "NET PAY", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 8.7)

'Updated 2/10/2018
Call PrintReportHeader(xlSheet, xlApp1, "AX5:AX5", "Employer", "SSS", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)
'Call PrintReportHeader(xlSheet, xlApp1, "AX6:AX6", "SSS", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AY6:AY6", "Pagibig", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 8)
' Updated to include SSS Calamity Loan
Call PrintReportHeader(xlSheet, xlApp1, "AZ6:AZ6", "SSS Cal", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 8)

xlApp1.Sheets("Grand Total").Select
Call PrintReportHeader(xlSheet, xlApp1, "K5:K6", "Vacation", "Leave", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3) 'Updated April 13, 2015
Call PrintReportHeader(xlSheet, xlApp1, "L5:L6", "Sick", "Leave", xlCenter, xlCenter, True, False, True, 15, , 8, False, 1)
xlApp1.Sheets("Payroll Summary").Select

Do While Not rs.EOF
    If prevClientName <> rs!ClientName Then
        If i > 1 Then
            Call PrintTotals(xlSheet, xlApp1, ctr, k)
            ' Print grand total in Sheet 2.
            mRange = "H" & ctr & ":AZ" & ctr ' AY updated 2/10/18
            'xlSheet.Range(mRange).Select
            xlApp1.ActiveSheet.Range(mRange).Select
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
        prevStartCutoff = rs!StartPaydate   ' Updated 10/31/2016
        prevEndCutoff = rs!EndPaydate       ' Updated 10/31/2016
        'Call CopyToSheet2(xlSheet, xlApp1, LineCtr)
    ElseIf prevStartCutoff <> rs!StartPaydate Then
        If i > 2 Then
            Call PrintTotals(xlSheet, xlApp1, ctr, k)
            ' Print grand total in Sheet 2.
            mRange = "H" & ctr & ":AZ" & ctr   ' AY
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
        prevStartCutoff = rs!StartPaydate
        prevEndCutoff = rs!EndPaydate
    Else
        xlSheet.Cells(ctr, 1).VALUE = IIf(IsNull(rs!ClientName), "Deleted Client", rs!ClientName)
    End If
    
    xlSheet.Cells(ctr, 2).VALUE = rs!StartPaydate
    xlSheet.Cells(ctr, 3).VALUE = rs!EndPaydate
    
    xlSheet.Cells(ctr, 4).VALUE = rs!controlnum
    xlSheet.Cells(ctr, 5).VALUE = rs!EmployeeName
    
'    ' Print the cash advance in Excel Report.
'    If rs!CashAdvance2 > 0 And rs!EmployeeName <> "" Then
'        If pClient <> rs!ClientName Then
'            Call CopyToConsolidation(xlSheet, xlApp1, ctr, rs!EmployeeName, rs!CashAdvance2, "Consolidation", "New")
'        Else
'            Call CopyToConsolidation(xlSheet, xlApp1, ctr, rs!EmployeeName, rs!CashAdvance2, "Consolidation")
'        End If
'        pClient = rs!ClientName
'    End If
    
    xlSheet.Cells(ctr, 6).VALUE = rs!EmpTaxCode 'rs!EmpCivilStatus
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
    
'    xlSheet.Cells(ctr, 17).VALUE = Format(rs!TaxRefundHours, "##,##0.00")
'    xlSheet.Cells(ctr, 18).VALUE = Format(0, "##,##0.00")
'    'xlSheet.Cells(ctr, 18).VALUE = Format(rs!TaxRefundPay, "##,##0.00")
    ' Replace with Vacation and Sick Leave. Nov. 1, 2016
    xlSheet.Cells(ctr, 17).VALUE = Format(rs!VacationLeavePay, "##,##0.00")
    xlSheet.Cells(ctr, 18).VALUE = Format(rs!SickLeavePay, "##,##0.00")
    
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
    
'    xlSheet.Cells(ctr, 33).VALUE = Format(rs!TASAesla, "##,##0.00")
'    xlSheet.Cells(ctr, 33).VALUE = Format(rs!SSSEmployer, "##,##0.00")
    xlSheet.Cells(ctr, 33).VALUE = Format(rs!ExcessLglHolOTPay, "##,##0.00") ' previously SSS Employer 7/19/2014
    xlSheet.Cells(ctr, 34).VALUE = Format(rs!CashAdvance2, "##,##0.00") ' for Cash Advance
    xlSheet.Cells(ctr, 35).VALUE = Format(rs!MiscFee, "##,##0.00")
    xlSheet.Cells(ctr, 36).VALUE = Format(rs!LMCDed, "##,##0.00")
    
    xlSheet.Cells(ctr, 37).VALUE = Format(rs!SalaryLoan, "##,##0.00")
    xlSheet.Cells(ctr, 38).VALUE = Format(rs!Vale_IOU, "##,##0.00")
    xlSheet.Cells(ctr, 39).VALUE = Format(rs!Uniform, "##,##0.00")
'    xlSheet.Cells(ctr, 40).VALUE = Format(rs!Medical, "##,##0.00")
    xlSheet.Cells(ctr, 40).VALUE = Format(rs!TASAesla, "##,##0.00")
    
    'xlSheet.Cells(ctr, 41).VALUE = Format(rs!Variances, "##,##0.00")    ' for Sal. Adjmnt. 2/7/15
    xlSheet.Cells(ctr, 41).VALUE = Format(rs!OverUnderPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 42).VALUE = Format(rs!SSSLoan, "##,##0.00")
    xlSheet.Cells(ctr, 43).VALUE = Format(rs!PagibigLoan, "##,##0.00")
    xlSheet.Cells(ctr, 44).VALUE = Format(rs!ValeWithClient, "##,##0.00")
    xlSheet.Cells(ctr, 45).VALUE = Format(rs!UnionDues, "##,##0.00")
    xlSheet.Cells(ctr, 46).VALUE = Format(rs!HazardPay, "##,##0.00")
    xlSheet.Cells(ctr, 47).VALUE = Format(rs!Variances, "##,##0.00")    ' for Sal. Adjmnt. 2/7/15
    
    xlSheet.Cells(ctr, 48).VALUE = Format(rs!Deductions, "##,##0.00")
    xlSheet.Cells(ctr, 49).VALUE = Format(rs!NetPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 50).VALUE = Format(rs!SSSEmployer, "##,##0.00")
    xlSheet.Cells(ctr, 51).VALUE = Format(rs!PagibigEmployer, "##,##0.00")
    xlSheet.Cells(ctr, 52).VALUE = Format(rs!Medical, "##,##0.00") ' for SSS Calamity Loan 2/26/19

    
'    xlSheet.Cells(ctr, 48).VALUE = IIf(rs!AcctNo = "", "n/a", rs!AcctNo)
'
'    xlSheet.Cells(ctr, 49).VALUE = Format(rs!Month13Payment, "##,##0.00")
'    If rs!SSSEmployer <> 0 Then
'        xlSheet.Cells(ctr, 50).VALUE = Format(rs!SSSEmployer - IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer), "##,##0.00")
'        xlSheet.Cells(ctr, 51).VALUE = Format(IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer), "##,##0.00")
'    Else
'        xlSheet.Cells(ctr, 50).VALUE = Format(0, "##,##0.00")
'        xlSheet.Cells(ctr, 51).VALUE = Format(0, "##,##0.00")
'    End If
'    xlSheet.Cells(ctr, 52).VALUE = Format(rs!PHealthEmployer, "##,##0.00")
'    xlSheet.Cells(ctr, 53).VALUE = Format(rs!PagibigEmployer, "##,##0.00")
'    xlSheet.Cells(ctr, 54).VALUE = Format(rs!CalamityLoan, "##,##0.00")
'    xlSheet.Cells(ctr, 55).FormulaR1C1 = "=RC[-21]+RC[-5]+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
'    xlSheet.Cells(ctr, 56).FormulaR1C1 = "=RC[-2]+RC[-1]"
    
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext

Loop

Call PrintTotals(xlSheet, xlApp1, ctr, k)
mRange = "H" & ctr & ":AY" & ctr  'AW updated 2/10/18
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
        sumString = "=R[" & ((ctr - aa) * -1) & "]C"
        mRange = Chr(j) & aa
        mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
        SaveTotal = mID(SaveTotal, lf + 1, Len(SaveTotal) - lf)
    Wend
    mRange = Chr(j) & ctr
    xlSheet.Range(mRange).Select
    xlApp1.ActiveCell.VALUE = mGTotal
    xlSheet.Range(mRange).Select
    If xlApp1.ActiveCell.VALUE = 0 Then
        xlApp1.Selection.EntireColumn.Hidden = True
    End If
    sumString = ""
    mGTotal = 0#
Next

For j = Asc("A") To Asc("Z")  'Y Updated 2/26/19
    SaveTotal = ColString
    While Len(SaveTotal) <> 0
        lf = InStr(SaveTotal, ":")
        aa = LTrim(RTrim(Left(SaveTotal, lf - 1)))
        sumString = "=R[" & ((ctr - aa) * -1) & "]C"
        mRange = "A" & Chr(j) & aa
        mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
        SaveTotal = mID(SaveTotal, lf + 1, Len(SaveTotal) - lf)
    Wend
    mRange = "A" & Chr(j) & ctr
    xlSheet.Range(mRange).Select
    xlApp1.ActiveCell.VALUE = mGTotal
    xlSheet.Range(mRange).Select
    If xlApp1.ActiveCell.VALUE = 0 Then
        xlApp1.Selection.EntireColumn.Hidden = True
    End If
    sumString = ""
    mGTotal = 0#
Next

For j = Asc("X") To Asc("Z")
    SaveTotal = ColString
    While Len(SaveTotal) <> 0
        lf = InStr(SaveTotal, ":")
        aa = LTrim(RTrim(Left(SaveTotal, lf - 1)))
        sumString = "=R[" & ((ctr - aa) * -1) & "]C"
        mRange = "A" & Chr(j) & aa
        mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
        SaveTotal = mID(SaveTotal, lf + 1, Len(SaveTotal) - lf)
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
        SaveTotal = mID(SaveTotal, lf + 1, Len(SaveTotal) - lf)
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

mRange = "A5:AZ" & i + 7 'AY
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
End With

'Format amount to 0.00.
mRange = "E7:AZ" & i + 7    ' E7:AY
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

' Box IT.
mRange = "A5:AZ" & i + 7  ' AY
Call PutBox(xlSheet, xlApp1, mRange)

i = i + 5
'Change the font.
mRange = "E" & i + 2 & ":AZ" & i + 14 'AY
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
xlApp1.Sheets("Grand Total").Select
xlApp1.ActiveSheet.Range("A2").Select
xlApp1.ActiveSheet.Paste
xlApp1.Sheets("Payroll Summary").Select
'Copy Client Header to Sheet 2.
xlApp1.ActiveSheet.Range("A5:A6").Select
xlApp1.Selection.Copy
xlApp1.Sheets("Grand Total").Select
xlApp1.ActiveSheet.Range("A5:A6").Select
xlApp1.ActiveSheet.Paste
xlApp1.Sheets("Payroll Summary").Select
' Copy Column headers to Sheet 2.
xlApp1.ActiveSheet.Range("H5:AZ6").Select 'AY6
xlApp1.Selection.Copy
xlApp1.Sheets("Grand Total").Select
xlApp1.ActiveSheet.Range("B5:AT6").Select  'B5:AS6
xlApp1.ActiveSheet.Paste

xlApp1.Sheets("Grand Total").Select

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
For j = Asc("A") To Asc("T")  ' S
    mRange = "A" & Chr(j) & 7 & ":A" & Chr(j) & LineCtr + 1
    xlApp1.ActiveSheet.Range(mRange).Select
    mRange2 = "A" & Chr(j) & LineCtr + 1
    xlApp1.ActiveSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((LineCtr - 6) * -1) & "]C:R[-1]C)"
Next

'Format Sheet 2.
mRange = "A5:AT" & LineCtr + 3 ' AS
xlApp1.ActiveSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
End With

' Make Grand Total in Bold Font.
mRange = "A" & LineCtr + 1 & ":AT" & LineCtr + 1  ' AS
xlApp1.ActiveSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = True       'Font.Bold
End With

'Format amount to 0.00.
mRange = "B7:AT" & LineCtr + 3    ' AX E7:AS
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
mRange = "A5:AT" & LineCtr + 3 ' AS
Call PutBox(xlSheet, xlApp1, mRange)

mRange = "A5:A6"
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.ActiveSheet.Columns.Range(mRange).ColumnWidth = 30#
mRange = "AP5:AP6"
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.ActiveSheet.Columns.Range(mRange).ColumnWidth = 16
xlApp1.ActiveSheet.Range("A2").Select

'' Move the last 3 fields. 2/26/19
'xlApp1.Columns("AR:AT").Select
'xlApp1.Selection.Cut
'xlApp1.Columns("AL:AL").Select
'xlApp1.Selection.Insert Shift:=xlToRight

'Hide the last 2 columns. 2/10/18
xlApp1.Columns("AR:AT").Select
xlApp1.Selection.EntireColumn.Hidden = True

'xlApp1.ActiveSheet.Name = "Grand Total"
xlApp1.Sheets("Payroll Summary").Select

'If LoginName = "ARCRUZ" Then
If ForPayroll(AccountType) = False Then
    xlApp1.ActiveWindow.SelectedSheets.Select
    xlApp1.Columns("G:G").Select
    xlApp1.Selection.Delete Shift:=xlToLeft

    'xlApp1.ActiveWindow.SelectedSheets.Delete
    xlApp1.ActiveWindow.SelectedSheets.Select
    'xlApp1.Sheets("Grand Total").Select
Else
    xlApp1.ActiveWindow.SelectedSheets.Select
    'xlSheet.Range("A2").Select
End If
            
If WConsolidated = True Then
    Call CopyToConsolidation(rs, xlSheet, xlApp1, ctr2, rs!EmployeeName, rs!CashAdvance2, "Consolidation", "New")
End If

'Hide the last 2 columns. 2/10/18 Updated 2/26/19
xlApp1.Columns("AX:AZ").Select
xlApp1.Selection.EntireColumn.Hidden = True

'' Print the cash advance in Excel Report.
'ctr2 = 7
'rs.MoveFirst
'rs.Filter = "Cashadvance2 > 0"
'pClient = ""  'rs!ClientName ' for changing the clientname
'
'Do While Not rs.EOF
'    If rs!CashAdvance2 > 0 And rs!EmployeeName <> "" Then
'        If pClient <> rs!ClientName Then
'            Call CopyToConsolidation(xlSheet, xlApp1, ctr2, rs!EmployeeName, rs!CashAdvance2, "Consolidation", "New")
'        Else
'            Call CopyToConsolidation(xlSheet, xlApp1, ctr2, rs!EmployeeName, rs!CashAdvance2, "Consolidation")
'        End If
'        pClient = rs!ClientName
'    End If
'    ctr2 = ctr2 + 1
'    rs.MoveNext
'Loop

xlApp1.Range("A2").Select
xlApp1.Visible = True

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

    'xlApp1.Sheets("Sheet2").Select
    xlApp1.Sheets("Grand Total").Select
    mRange = "A" & ctr & ":A" & ctr
    xlApp1.ActiveSheet.Range(mRange).Select
    xlApp1.ActiveSheet.Range(mRange).VALUE = mClientName
    mRange = "B" & ctr & ":AT" & ctr  ' AS updated 2/10/18 updated 2/26/19
    xlApp1.ActiveSheet.Range(mRange).Select
    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    xlApp1.Sheets("Payroll Summary").Select

End Sub

Private Sub CopyToOtherSheet(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal ctr As Integer, ByVal mDate As Date, ByVal mSSS As Double, ByVal mECC As Double, _
    Optional SheetType As String = "MCL", _
    Optional SheetName As String = "Sheet2")

Dim mRange As String

xlApp1.Sheets(SheetName).Select
'mRange = "A" & ctr & ":A" & ctr
mRange = "B" & ctr & ":D" & ctr
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False

' Get SSS Number.
'mRange = "E" & ctr & ":E" & ctr
mRange = "F" & ctr & ":F" & ctr
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.ActiveSheet.Range(mRange).VALUE = Replace(xlApp1.ActiveSheet.Range(mRange).VALUE, "-", "")
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.Selection.Copy

' Paste SSS Number.
mRange = "A" & ctr & ":A" & ctr
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False
xlApp1.ActiveSheet.Range(mRange).NumberFormat = "0000000000"

If SheetType = "MCL" Then
    ' Date Hired
    mRange = "E" & ctr & ":E" & ctr
    xlApp1.ActiveSheet.Range(mRange).VALUE = mDate
    ' Status
    mRange = "F" & ctr & ":F" & ctr
    xlApp1.ActiveSheet.Range(mRange).VALUE = "3"
    ' Get Employee Share
    mRange = "G" & ctr & ":G" & ctr
    xlApp1.ActiveSheet.Range(mRange).VALUE = mSSS
    xlApp1.ActiveSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
    ' Get ECC Share
    mRange = "H" & ctr & ":H" & ctr
    xlApp1.ActiveSheet.Range(mRange).VALUE = mECC
    xlApp1.ActiveSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
Else ' for EPF
    ' Date of Birth
    mRange = "E" & ctr & ":E" & ctr
    xlApp1.ActiveSheet.Range(mRange).VALUE = mDate
    ' Clear this column.
    mRange = "F" & ctr & ":F" & ctr
    xlApp1.ActiveSheet.Range(mRange).VALUE = ""
End If
xlApp1.Sheets("Sheet1").Select

End Sub

Private Sub CopyToConsolidation(ByRef rsRec As Recordset, ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal ctr As Integer, ByVal mName, ByVal mValue, Optional SheetName As String = "Consolidation", _
    Optional mClient As String = "")

Dim mRange As String, mRangeAmt As String, mRangeName As String
Dim LastPos As Integer
Dim mInspector As Integer
Dim mSumCtr As Integer
Dim pClient As String

On Error GoTo FixIT1

' Print the cash advance in Excel Report.
rsRec.MoveFirst
rsRec.Filter = "Cashadvance2 > 0"
pClient = ""  'rs!ClientName ' for changing the clientname

xlApp1.Sheets(SheetName).Select

Do While Not rsRec.EOF
    mName = rsRec!EmployeeName
    mValue = rsRec!CashAdvance2
    If pClient <> rsRec!ClientName Then
        mClient = "New"
    Else
        mClient = ""
    End If
'            Call CopyToConsolidation(xlSheet, xlApp1, ctr2, rs!EmployeeName, rs!CashAdvance2, "Consolidation", "New")
'        Else
'            Call CopyToConsolidation(xlSheet, xlApp1, ctr2, rs!EmployeeName, rs!CashAdvance2, "Consolidation")
'        End If

If ctr = 7 Then
    'Paste the Employee Name.
    mRangeName = "C" & ctr & ":C" & ctr
    xlApp1.ActiveSheet.Range(mRangeName).Select
    xlApp1.ActiveCell.FormulaR1C1 = mName
    'Paste the Cash Advance Amount.
    mRangeAmt = "E" & ctr & ":E" & ctr
    xlApp1.ActiveSheet.Range(mRangeAmt).Select
    xlApp1.ActiveCell.FormulaR1C1 = mValue
    mRange = "A" & (ctr) & ":N" & (ctr)
    LastPos = ctr
    mInspector = 1
    mSumCtr = 1
Else
    If mClient = "New" Then
        If mInspector = 1 Then
            xlApp1.ActiveSheet.Range("F6").Select
        ElseIf mInspector = 2 Then
            xlApp1.ActiveSheet.Range("I6").Select
            'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[4]C[-1]:R[9]C[-1])"
        'Else
        '    xlApp1.ActiveSheet.Range("L6").Select
        End If
        xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-1]:R[" & (LastPos - 7) & "]C[-1])"
        'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-1]:R[3]C[-1])"
        If mInspector = 1 Then
            mInspector = 2
            mRangeName = "C" & LastPos & ":C" & LastPos
            mRangeAmt = "H" & LastPos & ":H" & LastPos
            xlApp1.ActiveSheet.Range(mRangeName).Select
            xlApp1.ActiveCell.FormulaR1C1 = mName
            xlApp1.ActiveSheet.Range(mRangeAmt).Select
            xlApp1.ActiveCell.FormulaR1C1 = mValue
            GoTo Jump1
        Else
            mInspector = 3
            mRangeName = "C" & LastPos & ":C" & LastPos
            mRangeAmt = "K" & LastPos & ":K" & LastPos
            mRange = "A" & (LastPos) & ":N" & (LastPos)
            xlApp1.ActiveSheet.Range(mRangeName).Select
            xlApp1.ActiveCell.FormulaR1C1 = mName
            xlApp1.ActiveSheet.Range(mRangeAmt).Select
            xlApp1.ActiveCell.FormulaR1C1 = mValue
            GoTo Jump1
        End If
    Else
        mRange = "A" & (LastPos - 1) & ":N" & (LastPos - 1)
        mRangeName = "C" & LastPos & ":C" & LastPos
        If mInspector = 1 Then
            mRangeAmt = "E" & LastPos & ":E" & LastPos
        Else
            If mInspector = 2 Then
                mRangeAmt = "H" & LastPos & ":H" & LastPos
            Else
                mRangeAmt = "K" & LastPos & ":K" & LastPos
            End If
        End If
    End If
   
    'Paste the Employee Name.
    xlApp1.ActiveSheet.Range(mRange).Select
    xlApp1.Selection.Copy
    mRange = "A" & LastPos & ":N" & LastPos
    xlApp1.ActiveSheet.Range(mRange).Select
    xlApp1.Selection.Insert Shift:=xlDown
    
    xlApp1.ActiveSheet.Range(mRangeName).Select
    xlApp1.ActiveCell.FormulaR1C1 = mName
    'xlApp1.ActiveCell.VALUE = mName
    'xlApp1.ActiveSheet.Range(mRangeName).VALUE = mName
    
    'Paste the Cash Advance Amount.
    xlApp1.ActiveSheet.Range(mRangeAmt).Select
    xlApp1.ActiveCell.FormulaR1C1 = mValue
    'xlApp1.ActiveCell.VALUE = mValue

    If mInspector = 3 Then
        xlApp1.ActiveSheet.Range("L6").Select
        xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-1]:R[" & (LastPos - 6) & "]C[-1])"
    End If

Jump1:

End If

LastPos = LastPos + 1

    pClient = rsRec!ClientName
    ctr = ctr + 1
    rsRec.MoveNext
Loop

xlApp1.ActiveSheet.Range("B1").Select
xlApp1.Sheets("Payroll Summary").Select

Exit Sub

FixIT1:

    Err.Clear
    Resume Next
    
End Sub

'Private Sub CopyToConsolidation(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
'    ByVal ctr As Integer, ByVal mName, ByVal mValue, Optional SheetName As String = "Consolidation", _
'    Optional mClient As String = "")
'
'Static mRange As String, mRangeAmt As String, mRangeName As String
'Static LastPos As Integer
'Static mInspector As Integer
'Static mSumCtr As Integer
'
'On Error GoTo FixIT1
'
'xlApp1.Sheets(SheetName).Select
'If ctr = 7 Then
'    'Paste the Employee Name.
'    mRangeName = "C" & ctr & ":C" & ctr
'    xlApp1.ActiveSheet.Range(mRangeName).Select
'    xlApp1.ActiveCell.FormulaR1C1 = mName
'    'Paste the Cash Advance Amount.
'    mRangeAmt = "E" & ctr & ":E" & ctr
'    xlApp1.ActiveSheet.Range(mRangeAmt).Select
'    xlApp1.ActiveCell.FormulaR1C1 = mValue
'    mRange = "A" & (ctr) & ":N" & (ctr)
'    LastPos = ctr
'    mInspector = 1
'    mSumCtr = 1
'Else
'    If mClient = "New" Then
'        If mInspector = 1 Then
'            xlApp1.ActiveSheet.Range("F6").Select
'        ElseIf mInspector = 2 Then
'            xlApp1.ActiveSheet.Range("I6").Select
'            'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[4]C[-1]:R[9]C[-1])"
'        'Else
'        '    xlApp1.ActiveSheet.Range("L6").Select
'        End If
'        xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-1]:R[" & (LastPos - 7) & "]C[-1])"
'        'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-1]:R[3]C[-1])"
'        If mInspector = 1 Then
'            mInspector = 2
'            mRangeName = "C" & LastPos & ":C" & LastPos
'            mRangeAmt = "H" & LastPos & ":H" & LastPos
'            xlApp1.ActiveSheet.Range(mRangeName).Select
'            xlApp1.ActiveCell.FormulaR1C1 = mName
'            xlApp1.ActiveSheet.Range(mRangeAmt).Select
'            xlApp1.ActiveCell.FormulaR1C1 = mValue
'            GoTo Jump1
'        Else
'            mInspector = 3
'            mRangeName = "C" & LastPos & ":C" & LastPos
'            mRangeAmt = "K" & LastPos & ":K" & LastPos
'            mRange = "A" & (LastPos) & ":N" & (LastPos)
'            xlApp1.ActiveSheet.Range(mRangeName).Select
'            xlApp1.ActiveCell.FormulaR1C1 = mName
'            xlApp1.ActiveSheet.Range(mRangeAmt).Select
'            xlApp1.ActiveCell.FormulaR1C1 = mValue
'            GoTo Jump1
'        End If
'    Else
'        mRange = "A" & (LastPos - 1) & ":N" & (LastPos - 1)
'        mRangeName = "C" & LastPos & ":C" & LastPos
'        If mInspector = 1 Then
'            mRangeAmt = "E" & LastPos & ":E" & LastPos
'        Else
'            If mInspector = 2 Then
'                mRangeAmt = "H" & LastPos & ":H" & LastPos
'            Else
'                mRangeAmt = "K" & LastPos & ":K" & LastPos
'            End If
'        End If
'    End If
'
'    'Paste the Employee Name.
'    xlApp1.ActiveSheet.Range(mRange).Select
'    xlApp1.Selection.Copy
'    mRange = "A" & LastPos & ":N" & LastPos
'    xlApp1.ActiveSheet.Range(mRange).Select
'    xlApp1.Selection.Insert Shift:=xlDown
'
'    xlApp1.ActiveSheet.Range(mRangeName).Select
'    xlApp1.ActiveCell.FormulaR1C1 = mName
'    'xlApp1.ActiveCell.VALUE = mName
'    'xlApp1.ActiveSheet.Range(mRangeName).VALUE = mName
'
'    'Paste the Cash Advance Amount.
'    xlApp1.ActiveSheet.Range(mRangeAmt).Select
'    xlApp1.ActiveCell.FormulaR1C1 = mValue
'    'xlApp1.ActiveCell.VALUE = mValue
'
'    If mInspector = 3 Then
'        xlApp1.ActiveSheet.Range("L6").Select
'        xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-1]:R[" & (LastPos - 6) & "]C[-1])"
'    End If
'
'Jump1:
'
'End If
'
'LastPos = LastPos + 1
'xlApp1.Sheets("Payroll Summary").Select
'
'Exit Sub
'
'FixIT1:
'
'    Err.Clear
'    xlApp1.ActiveSheet.Range(mRangeName).VALUE = mName
'    Resume Next
'
'End Sub

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
sumString = "=SUM(R[" & (i - k) * -1 & "]C:R[-1]C)"
mRange = "E" & i & ":E" & i
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.ActiveCell.FormulaR1C1 = "           TOTAL"

'Print Total Employee.
mRange = "D" & i & ":D" & i
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.ActiveCell.FormulaR1C1 = IIf(i = k, i - 1, i - k)

For j = Asc("H") To Asc("Z")
    mRange = Chr(j) & k & ":" & Chr(j) & i
    xlApp1.ActiveSheet.Range(mRange).Select
    mRange2 = Chr(j) & i
    xlApp1.ActiveSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
    xlApp1.ActiveSheet.Range(mRange).Select
Next
For j = Asc("A") To Asc("W")    ' K
    mRange = "A" & Chr(j) & k & ":A" & Chr(j) & i
    xlApp1.ActiveSheet.Range(mRange).Select
    mRange2 = "A" & Chr(j) & i
    xlApp1.ActiveSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
    xlApp1.ActiveSheet.Range(mRange).Select
Next
For j = Asc("X") To Asc("Z")    ' K
    mRange = "A" & Chr(j) & k & ":A" & Chr(j) & i
    xlApp1.ActiveSheet.Range(mRange).Select
    mRange2 = "A" & Chr(j) & i
    xlApp1.ActiveSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
    xlApp1.ActiveSheet.Range(mRange).Select
Next
For j = Asc("A") To Asc("D")    ' K
    mRange = "B" & Chr(j) & k & ":B" & Chr(j) & i
    xlApp1.ActiveSheet.Range(mRange).Select
    mRange2 = "B" & Chr(j) & i
    xlApp1.ActiveSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
    xlApp1.ActiveSheet.Range(mRange).Select
Next
End Sub
   
Private Sub BoxDetails(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal mRange2 As String)

xlApp1.ActiveSheet.Range(mRange2).Select
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
    .Font.Name = "Arial Narrow"     'Font.Name
    .Font.Size = 10                 'Font.size
End With

End Sub

Private Sub ComputeTotals(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal ctr As Integer, ByVal i As Integer, ByVal StartOfLine As Integer, Optional LoanFlag As Boolean = False, _
    Optional arrMax As Integer = 0)

Dim mRange As String
Dim mRange2 As String
Dim sumString As String
Dim kk As Integer

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
'            mRange = "E" & ctr + 1
'            xlSheet.Range(mRange).Select
'            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'            xlApp1.ActiveCell.FormulaR1C1 = sumString
'            xlApp1.Selection.NumberFormat = "#,##0.00"
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
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            xlApp1.Selection.NumberFormat = "#,##0.00"
            mRange = "I" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            xlApp1.Selection.NumberFormat = "#,##0.00"
            mRange = "J" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            xlApp1.Selection.NumberFormat = "#,##0.00"
            
            mRange = "H" & ctr + 1
            mRange2 = "A" & StartOfLine & ":J" & ctr + 1
            
'            mRange = "F" & ctr + 1
'            xlSheet.Range(mRange).Select
'            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'            xlApp1.ActiveCell.FormulaR1C1 = sumString
'            xlApp1.Selection.NumberFormat = "#,##0.00"
'            mRange = "G" & ctr + 1
'            xlSheet.Range(mRange).Select
'            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'            xlApp1.ActiveCell.FormulaR1C1 = sumString
'            xlApp1.Selection.NumberFormat = "#,##0.00"
'            mRange = "H" & ctr + 1
'            mRange2 = "A" & StartOfLine & ":J" & ctr + 1
        
        End If
    Case 2, 9 ' Alphalist
        If OptReport(9).VALUE = True Then
            mRange = "I" & ctr + 1
            mRange2 = "A" & StartOfLine & ":O" & ctr + 1
        Else ' Tax
            mRange = "F" & ctr + 1
            mRange2 = "A" & StartOfLine & ":I" & ctr + 1
        End If
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
            'mRange = "H" & ctr + 1
            'mrange2 = "A7:J" & ctr + 1
            mRange2 = "A" & StartOfLine & ":J" & ctr + 1
        End If
    Case 4  ' PhilHealth
        mRange = "G" & ctr + 1
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        mRange = "H" & ctr + 1
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        mRange = "I" & ctr + 1
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
        'mrange2 = "A7:J" & ctr + 1
        mRange2 = "A" & StartOfLine & ":I" & ctr + 1
    Case 5  ' 13th Month
    
'        xlApp1.Sheets("Sheet2").Select
'        mRange = "B" & ctr + 2
'        mRange2 = "A" & StartOfLine & ":D" & ctr + 1
'        sumString = "=SUM(R[" & (ctr - StartOfLine - 2) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'
'        xlApp1.Sheets("Sheet1").Select
        mRange = "F" & ctr + 1
        'mrange2 = "A7:G" & ctr + 1
        mRange2 = "A" & StartOfLine & ":F" & ctr + 1
    
    Case 6, 66 ' Lates and Absences
        mRange = "E" & ctr + 1
        xlSheet.Range(mRange).Select
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
        xlApp1.Selection.NumberFormat = "#,##0.00"
'        mRange = "F" & ctr + 1
'        xlSheet.Range(mRange).Select
'        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'        xlApp1.Selection.NumberFormat = "#,##0.00"
'        mRange = "G" & ctr + 1
'        xlSheet.Range(mRange).Select
'        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'        xlApp1.Selection.NumberFormat = "#,##0.00"
'        mRange = "H" & ctr + 1
'        xlSheet.Range(mRange).Select
'        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'        xlApp1.Selection.NumberFormat = "#,##0.00"
'        mRange = "F" & ctr + 1
        mRange2 = "A" & StartOfLine & ":I" & ctr + 1
    Case 7  ' Head Count
    Case 8  ' Other Deductions
        mRange = "E" & ctr + 1
        ' Updated July 3, 2013, Bernard.
        For kk = Asc("E") To Asc("E") + 6 + arrMax
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            xlApp1.Selection.NumberFormat = "#,##0.00"
            mRange = Chr(kk + 1) & ctr + 1
        Next kk
        mRange2 = "A" & StartOfLine & ":" & Chr(79 + arrMax) & ctr + 1 'L
        
'        'ADDED BY Bernard 3/10/2007
'        If ESumString = "" Then ESumString = "=SUM(" & mRange Else ESumString = ESumString & "," & mRange
'        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'        xlApp1.Selection.NumberFormat = "#,##0.00"
'        mRange = "F" & ctr + 1
'        'ADDED BY Bernard 3/10/2007
'        If FSumString = "" Then FSumString = "=SUM(" & mRange Else FSumString = FSumString & "," & mRange
'        xlSheet.Range(mRange).Select
'        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'        xlApp1.Selection.NumberFormat = "#,##0.00"
'
'        mRange = "G" & ctr + 1
'        'ADDED BY Bernard 3/10/2007
'        If GSumString = "" Then GSumString = "=SUM(" & mRange Else GSumString = GSumString & "," & mRange
'        xlSheet.Range(mRange).Select
'        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'        xlApp1.Selection.NumberFormat = "#,##0.00"
'        mRange = "H" & ctr + 1
'        'ADDED BY Bernard 3/10/2007
'        If HSumString = "" Then HSumString = "=SUM(" & mRange Else HSumString = HSumString & "," & mRange
'        xlSheet.Range(mRange).Select
'        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'        xlApp1.Selection.NumberFormat = "#,##0.00"
'        mRange = "I" & ctr + 1
'        'ADDED BY Bernard 3/10/2007
'        If ISumString = "" Then ISumString = "=SUM(" & mRange Else ISumString = ISumString & "," & mRange
'        xlSheet.Range(mRange).Select
'        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'        xlApp1.Selection.NumberFormat = "#,##0.00"
'        mRange = "J" & ctr + 1
'        'ADDED BY Bernard April 19, 2008
'        If JSumString = "" Then JSumString = "=SUM(" & mRange Else JSumString = JSumString & "," & mRange
'        xlSheet.Range(mRange).Select
'        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'        xlApp1.Selection.NumberFormat = "#,##0.00"
'        mRange = "K" & ctr + 1
'        'ADDED BY Bernard April 19, 2008
'        If KSumString = "" Then KSumString = "=SUM(" & mRange Else KSumString = KSumString & "," & mRange
'        xlSheet.Range(mRange).Select
'        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'        xlApp1.Selection.NumberFormat = "#,##0.00"
'        mRange = "L" & ctr + 1
'        'ADDED BY Bernard April 19, 2008
'        If LSumString = "" Then LSumString = "=SUM(" & mRange Else LSumString = LSumString & "," & mRange
'        xlSheet.Range(mRange).Select
'        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'        xlApp1.Selection.NumberFormat = "#,##0.00"
'        mRange = "M" & ctr + 1
'        'ADDED BY Bernard April 19, 2008
'        If MSumString = "" Then MSumString = "=SUM(" & mRange Else MSumString = MSumString & "," & mRange
'        xlSheet.Range(mRange).Select
'        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
'        xlApp1.ActiveCell.FormulaR1C1 = sumString
'        xlApp1.Selection.NumberFormat = "#,##0.00"
'        mRange = "N" & ctr + 1
'        mRange2 = "A" & StartOfLine & ":P" & ctr + 1 'L
    End Select
        
    xlSheet.Range(mRange).Select
    If LoanFlag = True Then
        If LoanStart <> 0 Then StartOfLine = LoanStart + 1
        sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
        xlApp1.ActiveCell.FormulaR1C1 = sumString
    Else
        If ReportType = 1 And chkLoans.VALUE = 1 Then
            xlSheet.Range(mRange).VALUE = TotalLoan
        ElseIf ReportType = 9 Then
            mRange = "I" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            mRange = "J" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            mRange = "K" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            mRange = "L" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            mRange = "M" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            mRange = "N" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
            mRange = "O" & ctr + 1
            xlSheet.Range(mRange).Select
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
        Else
            sumString = "=SUM(R[" & (ctr - StartOfLine + 1) * -1 & "]C:R[-1]C)"
            xlApp1.ActiveCell.FormulaR1C1 = sumString
        End If
    End If
    xlApp1.Selection.NumberFormat = "#,##0.00"
    If LoanFlag = False Then
        xlApp1.Sheets("Sheet1").Select
        xlSheet.Range(mRange2).Select
        Call BoxDetails(xlSheet, xlApp1, mRange2)
        If ReportType = 1 And chkLoans.VALUE <> 1 Then   ' SSS Premium
            xlApp1.Sheets("Sheet2").Select
            mRange2 = "A" & StartOfLine & ":H" & ctr + 1
            Call BoxDetails(xlSheet, xlApp1, mRange2)
            xlApp1.ActiveSheet.Name = "MCL Data Entry"
            xlApp1.ActiveSheet.Range("A1").Select
            xlApp1.Sheets("Sheet3").Select
            mRange2 = "A" & StartOfLine & ":E" & ctr + 1
            Call BoxDetails(xlSheet, xlApp1, mRange2)
            xlApp1.ActiveSheet.Name = "EPF Data Entry"
            xlApp1.ActiveSheet.Range("A1").Select
            xlApp1.Sheets("Sheet1").Select
        End If
    End If
    LoanStart = ctr + 1
    
End Sub

Private Sub SalariesAndWagesReport(ByRef rs As ADODB.Recordset, ByVal mMonth As String, ByVal mYear As Integer)
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
Dim rsSalCodes As Recordset

MsgBox "Under development!", vbInformation, PayrollAlert
Exit Sub


Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)

ctr = 9
i = 1

'Updated by Bernard May 13, 2008.
mHead = CompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 12, True)
mHead = "Salaries And Wages Report"
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, False)
mHead = "Cut-off Date: " & dtStartDate.VALUE & " to " & dtEndDate.VALUE
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, False)

Call PrintReportHeader(xlSheet, xlApp1, "A7:A8", "Number", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "B7:B8", "Account No", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 40)
Call PrintReportHeader(xlSheet, xlApp1, "C7:C8", "Particulars", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "D7:D8", "Amount", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 10)

rs.MoveFirst
mHead = "Department Name: " & rs!ClientName
Call PrintReportHeader(xlSheet, xlApp1, "A5", mHead, "", xlLeft, , False, False, False, , , 10, False)

On Error GoTo NextRec
StrQuery = "SELECT SalaryCode, Particular, AccountNo from HRSalaryCode WHERE Includedflag = 1 order by FieldSeq"
Set rsSalCodes = ReturnRst(StrQuery, cnnstr, mstatic)
With rsSalCodes

Do While Not rs.EOF
    xlSheet.Cells(ctr, 1).VALUE = ctr
    xlSheet.Cells(ctr, 2).VALUE = !Particular
    xlSheet.Cells(ctr, 3).VALUE = !AccountNo
    mRange2 = Trim$(!SalaryCode)
    If CheckField(rsSalCodes, mRange2) Then
        xlSheet.Cells(ctr, 4).VALUE = rs.Fields(mRange2).VALUE
    Else
        xlSheet.Cells(ctr, 4).VALUE = "0.00"
    End If
    'xlSheet.Cells(ctr, 4).FormulaR1C1 = "=RC[-2]-RC[-1]" '"=RC[-2]+RC[-1]"
    ctr = ctr + 1
    i = i + 1
    .MoveNext
Loop

NextRec:
    MsgBox Err.Description, vbInformation, PayrollAlert
    Err.Clear
    'Resume Next

End With

ctr = ctr + 3
xlSheet.Cells(ctr, 2).VALUE = "TOTAL HEADCOUNT"
  
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
ctr = ctr + 3
xlSheet.Cells(ctr, 2).VALUE = "TOTAL HEADCOUNT"
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)
  
  
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

Private Function CheckField(ByRef rsCodes As Recordset, mFld As String) As Boolean
Dim bkmrk

With rsCodes
    If .EOF = False Then
        bkmrk = .Bookmark
    Else
        bkmrk = 1
    End If
    
    
    
    .Bookmark = bkmrk
    
    

End With

End Function

Private Sub PrintCircuitPayroll(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim mRange2 As String
Dim sumString As String, ExcelTemplate As String
Dim ctr As Long
Dim i As Integer, j As Integer, k As Integer
Dim cPos As Integer
Dim EndSw As Boolean
Dim mHead As String, mCutoff As String

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
Dim lf, aa, mTemplate As String
Dim prevStartCutoff As Variant
Dim prevEndCutoff As Variant

mTemplate = ".xls"
ExcelTemplate = App.Path & "\Templates\Circuit Payroll Summary Template.xls"

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add(ExcelTemplate)

xlApp1.WindowState = xlMaximized
xlApp1.Visible = False
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)

'--------

mSeriesNo = "45"
SaveTotal = ""
firstRun = True

cPos = 5
ctr = 7
LineCtr = 7
i = 1
k = 7

'Updated by Bernard April 5, 2014.
mHead = CompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 14, True)
mCutoff = Format(rs!EndPaydate, "mm/dd/yyyy")
If CutOffPeriod = "START" Then
    mHead = "Payroll Summary Report for the Month of " & Format(rs!StartPaydate, "mmmm") & " " & Format(rs!StartPaydate, "yyyy")
Else
    mHead = "Payroll Summary Report for the Month of " & Format(rs!EntryDate, "mmmm") & " " & Format(rs!EntryDate, "yyyy")
End If
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
' Start printing of Headings.
Call PrintReportHeader(xlSheet, xlApp1, "A5:A6", mCutoff, "", xlCenter, xlBottom, True, False, True, 15, , 8, False, 30)

'Call PrintReportHeader(xlSheet, xlApp1, "B5:B6", "Start Cut-Off", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "C5:C6", "End Cut-Off", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 10)
'
'Call PrintReportHeader(xlSheet, xlApp1, "D5:D6", "DTR ID", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 13)
'Call PrintReportHeader(xlSheet, xlApp1, "E5:E6", "Employee Name", "", xlCenter, xlBottom, True, False, True, 15, , 8, False, 27)
'Call PrintReportHeader(xlSheet, xlApp1, "F5:F6", "Tax", "Status", xlCenter, xlCenter, True, True, True, 15, , 8, False, 5)
'Call PrintReportHeader(xlSheet, xlApp1, "G5:G6", "Daily", "Rate", xlCenter, xlCenter, True, True, True, 15, , 8, False, 7)
'Call PrintReportHeader(xlSheet, xlApp1, "H5:H6", "Reg.", "Days", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
'Call PrintReportHeader(xlSheet, xlApp1, "I5:I6", "This Pay", "Period", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9.5)
'Call PrintReportHeader(xlSheet, xlApp1, "J5:J6", "Absent", "Deduction", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)
'Call PrintReportHeader(xlSheet, xlApp1, "K5:K6", "Tardiness", "Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)
'Call PrintReportHeader(xlSheet, xlApp1, "L5:L6", "Tardiness", "Deduction", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)
'Call PrintReportHeader(xlSheet, xlApp1, "M5:M6", "Reg OT", "Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 6)
'Call PrintReportHeader(xlSheet, xlApp1, "N5:N6", "Reg OT", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'Call PrintReportHeader(xlSheet, xlApp1, "O5:O6", "Night Diff", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
'Call PrintReportHeader(xlSheet, xlApp1, "P5:P6", "Night Diff", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
'
'Call PrintReportHeader(xlSheet, xlApp1, "Q5:Q6", "Rest Day", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
'Call PrintReportHeader(xlSheet, xlApp1, "R5:R6", "Rest Day", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
'Call PrintReportHeader(xlSheet, xlApp1, "S5:S6", "Sunday OT", "Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
'Call PrintReportHeader(xlSheet, xlApp1, "T5:T6", "Sunday OT", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
'
'Call PrintReportHeader(xlSheet, xlApp1, "U5:U6", "Spl Hol.", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
'Call PrintReportHeader(xlSheet, xlApp1, "V5:V6", "Spl Hol.", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
'Call PrintReportHeader(xlSheet, xlApp1, "W5:W6", "Lgl Hol.", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
'Call PrintReportHeader(xlSheet, xlApp1, "X5:X6", "Lgl Hol.", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
'Call PrintReportHeader(xlSheet, xlApp1, "Y5:Y6", "Salary", "Adjustment", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
'Call PrintReportHeader(xlSheet, xlApp1, "Z5:Z6", "Incentive", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
'Call PrintReportHeader(xlSheet, xlApp1, "AA5:AA6", "E-cola/Allow.", "Payment", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9.5)
'Call PrintReportHeader(xlSheet, xlApp1, "AB5:AB6", "GROSS PAY", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 9)
'Call PrintReportHeader(xlSheet, xlApp1, "AC5:AC6", "WTAX", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
'Call PrintReportHeader(xlSheet, xlApp1, "AD5:AF5", "Employee", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)
'Call PrintReportHeader(xlSheet, xlApp1, "AD6:AD6", "SSS", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
'Call PrintReportHeader(xlSheet, xlApp1, "AE6:AE6", "PhilHealth", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
'Call PrintReportHeader(xlSheet, xlApp1, "AF6:AF6", "Pag-ibig", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
''Call PrintReportHeader(xlSheet, xlApp1, "AG5:AG6", "Other", "Deduction", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
'Call PrintReportHeader(xlSheet, xlApp1, "AG5:AG6", "SSS", "Employer", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
'Call PrintReportHeader(xlSheet, xlApp1, "AH5:AH6", "Cash", "Advance", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
'Call PrintReportHeader(xlSheet, xlApp1, "AI5:AI6", "Misc.", "Fee", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
''Call PrintReportHeader(xlSheet, xlApp1, "AJ5:AJ6", "Card", "Trans Fee", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
'Call PrintReportHeader(xlSheet, xlApp1, "AJ5:AJ6", "Association", "/SANAWU", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
'Call PrintReportHeader(xlSheet, xlApp1, "AK5:AK6", "Salary", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
'Call PrintReportHeader(xlSheet, xlApp1, "AL5:AL6", "Planters", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.87)
'Call PrintReportHeader(xlSheet, xlApp1, "AM5:AM6", "Uniform", "Charges", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
''Call PrintReportHeader(xlSheet, xlApp1, "AN5:AN6", "Medical", "Charges", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
'Call PrintReportHeader(xlSheet, xlApp1, "AN5:AN6", "Other", "Deduction", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
'Call PrintReportHeader(xlSheet, xlApp1, "AO5:AO6", "Variances", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
'Call PrintReportHeader(xlSheet, xlApp1, "AP5:AP6", "Adjustment", "Audit", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
'Call PrintReportHeader(xlSheet, xlApp1, "AQ5:AQ6", "SSS", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.8)
'Call PrintReportHeader(xlSheet, xlApp1, "AR5:AR6", "PBIG", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.8)
'Call PrintReportHeader(xlSheet, xlApp1, "AS5:AS6", "Vale", "w/ Client", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
'Call PrintReportHeader(xlSheet, xlApp1, "AT5:AT6", "Union", "Dues", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
'Call PrintReportHeader(xlSheet, xlApp1, "AU5:AU6", "Hazard", "Pay", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
'
'Call PrintReportHeader(xlSheet, xlApp1, "AV5:AV6", "TOTAL", "Deductions", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.7)
'Call PrintReportHeader(xlSheet, xlApp1, "AW5:AW6", "NET PAY", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 8.7)

Do While Not rs.EOF
    If prevClientName <> rs!ClientName Then
        If i > 1 Then
            Call PrintTotals(xlSheet, xlApp1, ctr, k)
            ' Print grand total in Sheet 2.
            mRange = "H" & ctr & ":AW" & ctr ' BD
            xlApp1.ActiveSheet.Range(mRange).Select
            xlApp1.Selection.Copy
            Call CopyToSheet2(xlSheet, xlApp1, LineCtr, prevClientName & " (" & prevStartCutoff & "-" & prevEndCutoff & ")", False)
            LineCtr = LineCtr + 1
            ctr = ctr + 1
            i = i + 1
            k = ctr
            SaveTotal = SaveTotal & ctr - 1 & ":"
        End If
        xlSheet.Cells(ctr, 1).VALUE = IIf(IsNull(rs!ClientName), "Deleted Client", rs!ClientName)
        prevClientName = rs!ClientName
        'Call CopyToSheet2(xlSheet, xlApp1, LineCtr)
    ElseIf prevStartCutoff <> rs!StartPaydate Then
        If i > 2 Then
            Call PrintTotals(xlSheet, xlApp1, ctr, k)
            ' Print grand total in Sheet 2.
            mRange = "H" & ctr & ":AW" & ctr
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
        prevStartCutoff = rs!StartPaydate
        prevEndCutoff = rs!EndPaydate
    Else
        xlSheet.Cells(ctr, 1).VALUE = IIf(IsNull(rs!ClientName), "Deleted Client", rs!ClientName)
    End If
    
    ' Start of Detail Printing.
    xlSheet.Cells(ctr, 2).VALUE = "000-000-000"
    xlSheet.Cells(ctr, 3).VALUE = Format(rs!BasicPay, "##,##0.00")  ' Reg pay
    xlSheet.Cells(ctr, 4).VALUE = Format(0, "##,##0.00")  ' LWOP
    xlSheet.Cells(ctr, 5).VALUE = Format(0, "##,##0.00")  ' UT
    xlSheet.Cells(ctr, 6).VALUE = Format(rs!TardinessDeduction, "##,#0.00") ' Tardy
    xlSheet.Cells(ctr, 7).VALUE = Format(rs!BasicPay, "##,##0.00")  ' Basic Pay
    xlSheet.Cells(ctr, 8).VALUE = Format(rs!SundayOTPay, "##,##0.00")
    
    
    xlSheet.Cells(ctr, 2).VALUE = rs!StartPaydate
    xlSheet.Cells(ctr, 3).VALUE = rs!EndPaydate
    
    xlSheet.Cells(ctr, 4).VALUE = rs!controlnum
    xlSheet.Cells(ctr, 5).VALUE = rs!EmployeeName
    xlSheet.Cells(ctr, 6).VALUE = rs!EmpTaxCode 'rs!EmpCivilStatus
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
    
'    xlSheet.Cells(ctr, 33).VALUE = Format(rs!TASAesla, "##,##0.00")
    xlSheet.Cells(ctr, 33).VALUE = Format(rs!SSSEmployer, "##,##0.00")
    xlSheet.Cells(ctr, 34).VALUE = Format(rs!CashAdvance2, "##,##0.00") ' for Cash Advance
    xlSheet.Cells(ctr, 35).VALUE = Format(rs!MiscFee, "##,##0.00")
    xlSheet.Cells(ctr, 36).VALUE = Format(rs!LMCDed, "##,##0.00")
    
    xlSheet.Cells(ctr, 37).VALUE = Format(rs!SalaryLoan, "##,##0.00")
    xlSheet.Cells(ctr, 38).VALUE = Format(rs!Vale_IOU, "##,##0.00")
    xlSheet.Cells(ctr, 39).VALUE = Format(rs!Uniform, "##,##0.00")
'    xlSheet.Cells(ctr, 40).VALUE = Format(rs!Medical, "##,##0.00")
    xlSheet.Cells(ctr, 40).VALUE = Format(rs!TASAesla, "##,##0.00")
    xlSheet.Cells(ctr, 41).VALUE = Format(rs!Variances, "##,##0.00")
    xlSheet.Cells(ctr, 42).VALUE = Format(rs!OverUnderPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 43).VALUE = Format(rs!SSSLoan, "##,##0.00")
    xlSheet.Cells(ctr, 44).VALUE = Format(rs!PagibigLoan, "##,##0.00")
    xlSheet.Cells(ctr, 45).VALUE = Format(rs!ValeWithClient, "##,##0.00")
    xlSheet.Cells(ctr, 46).VALUE = Format(rs!UnionDues, "##,##0.00")
    xlSheet.Cells(ctr, 47).VALUE = Format(rs!HazardPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 48).VALUE = Format(rs!Deductions, "##,##0.00")
    xlSheet.Cells(ctr, 49).VALUE = Format(rs!NetPay, "##,##0.00")
    
    
'    xlSheet.Cells(ctr, 48).VALUE = IIf(rs!AcctNo = "", "n/a", rs!AcctNo)
'
'    xlSheet.Cells(ctr, 49).VALUE = Format(rs!Month13Payment, "##,##0.00")
'    If rs!SSSEmployer <> 0 Then
'        xlSheet.Cells(ctr, 50).VALUE = Format(rs!SSSEmployer - IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer), "##,##0.00")
'        xlSheet.Cells(ctr, 51).VALUE = Format(IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer), "##,##0.00")
'    Else
'        xlSheet.Cells(ctr, 50).VALUE = Format(0, "##,##0.00")
'        xlSheet.Cells(ctr, 51).VALUE = Format(0, "##,##0.00")
'    End If
'    xlSheet.Cells(ctr, 52).VALUE = Format(rs!PHealthEmployer, "##,##0.00")
'    xlSheet.Cells(ctr, 53).VALUE = Format(rs!PagibigEmployer, "##,##0.00")
'    xlSheet.Cells(ctr, 54).VALUE = Format(rs!CalamityLoan, "##,##0.00")
'    xlSheet.Cells(ctr, 55).FormulaR1C1 = "=RC[-21]+RC[-5]+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
'    xlSheet.Cells(ctr, 56).FormulaR1C1 = "=RC[-2]+RC[-1]"
    
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext

Loop

Call PrintTotals(xlSheet, xlApp1, ctr, k)
mRange = "H" & ctr & ":AW" & ctr  'BD
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
        sumString = "=R[" & ((ctr - aa) * -1) & "]C"
        mRange = Chr(j) & aa
        mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
        SaveTotal = mID(SaveTotal, lf + 1, Len(SaveTotal) - lf)
    Wend
    mRange = Chr(j) & ctr
    xlSheet.Range(mRange).Select
    xlApp1.ActiveCell.VALUE = mGTotal
    xlSheet.Range(mRange).Select
    If xlApp1.ActiveCell.VALUE = 0 Then
        xlApp1.Selection.EntireColumn.Hidden = True
    End If
    sumString = ""
    mGTotal = 0#
Next
For j = Asc("A") To Asc("W")
    SaveTotal = ColString
    While Len(SaveTotal) <> 0
        lf = InStr(SaveTotal, ":")
        aa = LTrim(RTrim(Left(SaveTotal, lf - 1)))
        sumString = "=R[" & ((ctr - aa) * -1) & "]C"
        mRange = "A" & Chr(j) & aa
        mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
        SaveTotal = mID(SaveTotal, lf + 1, Len(SaveTotal) - lf)
    Wend
    mRange = "A" & Chr(j) & ctr
    xlSheet.Range(mRange).Select
    xlApp1.ActiveCell.VALUE = mGTotal
    xlSheet.Range(mRange).Select
    If xlApp1.ActiveCell.VALUE = 0 Then
        xlApp1.Selection.EntireColumn.Hidden = True
    End If
    sumString = ""
    mGTotal = 0#
Next
For j = Asc("X") To Asc("Z")
    SaveTotal = ColString
    While Len(SaveTotal) <> 0
        lf = InStr(SaveTotal, ":")
        aa = LTrim(RTrim(Left(SaveTotal, lf - 1)))
        sumString = "=R[" & ((ctr - aa) * -1) & "]C"
        mRange = "A" & Chr(j) & aa
        mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
        SaveTotal = mID(SaveTotal, lf + 1, Len(SaveTotal) - lf)
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
        SaveTotal = mID(SaveTotal, lf + 1, Len(SaveTotal) - lf)
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

mRange = "A5:AW" & i + 7 'BD
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
End With

'Format amount to 0.00.
mRange = "E7:AW" & i + 7    ' BD E7:AP
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
mRange = "A5:AW" & i + 7  ' BD
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
mRange = "E" & i + 2 & ":AW" & i + 14 'BD
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
xlApp1.Sheets("Grand Total").Select
xlApp1.ActiveSheet.Range("A2").Select
xlApp1.ActiveSheet.Paste
xlApp1.Sheets("Payroll Summary").Select
'Copy Client Header to Sheet 2.
xlApp1.ActiveSheet.Range("A5:A6").Select
xlApp1.Selection.Copy
xlApp1.Sheets("Grand Total").Select
xlApp1.ActiveSheet.Range("A5:A6").Select
xlApp1.ActiveSheet.Paste
xlApp1.Sheets("Payroll Summary").Select
' Copy Column headers to Sheet 2.
xlApp1.ActiveSheet.Range("H5:AW6").Select
xlApp1.Selection.Copy
xlApp1.Sheets("Grand Total").Select
xlApp1.ActiveSheet.Range("B5:AQ6").Select
xlApp1.ActiveSheet.Paste

xlApp1.Sheets("Grand Total").Select

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
For j = Asc("A") To Asc("Q")
    mRange = "A" & Chr(j) & 7 & ":A" & Chr(j) & LineCtr + 1
    xlApp1.ActiveSheet.Range(mRange).Select
    mRange2 = "A" & Chr(j) & LineCtr + 1
    xlApp1.ActiveSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((LineCtr - 6) * -1) & "]C:R[-1]C)"
Next

'For j = Asc("Q") To Asc("X")
'    mRange = "A" & Chr(j) & 7 & ":A" & Chr(j) & LineCtr
'    xlApp1.ActiveSheet.Range(mRange).Select
'    mRange2 = "A" & Chr(j) & LineCtr + 1
'    xlApp1.ActiveSheet.Range(mRange2).Activate
'    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((LineCtr - 6) * -1) & "]C:R[-1]C)"
'Next

'Format Sheet 2.
mRange = "A5:AQ" & LineCtr + 3 ' AX
xlApp1.ActiveSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
End With

' Make Grand Total in Bold Font.
mRange = "A" & LineCtr + 1 & ":AQ" & LineCtr + 1  ' AX
xlApp1.ActiveSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = True       'Font.Bold
End With

'Format amount to 0.00.
mRange = "B7:AQ" & LineCtr + 3    ' AX E7:AP
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
mRange = "A5:AQ" & LineCtr + 3 ' AX
Call PutBox(xlSheet, xlApp1, mRange)

mRange = "A5:A6"
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.ActiveSheet.Columns.Range(mRange).ColumnWidth = 30#
mRange = "AP5:AP6"
xlApp1.ActiveSheet.Range(mRange).Select
xlApp1.ActiveSheet.Columns.Range(mRange).ColumnWidth = 16
xlApp1.ActiveSheet.Range("A2").Select
'xlApp1.ActiveSheet.Name = "Grand Total"
xlApp1.Sheets("Payroll Summary").Select

'If LoginName = "ARCRUZ" Then
If ForPayroll(AccountType) = False Then
    xlApp1.ActiveWindow.SelectedSheets.Select
    xlApp1.Columns("G:G").Select
    xlApp1.Selection.Delete Shift:=xlToLeft

    'xlApp1.ActiveWindow.SelectedSheets.Delete
    xlApp1.ActiveWindow.SelectedSheets.Select
    'xlApp1.Sheets("Grand Total").Select
Else
    xlApp1.ActiveWindow.SelectedSheets.Select
    'xlSheet.Range("A2").Select
End If
xlApp1.Range("A2").Select
xlApp1.Visible = True

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


'End of Program

