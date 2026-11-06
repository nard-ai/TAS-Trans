VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{90AF9238-EA14-11D8-B352-8B0F54D58444}#1.0#0"; "VisualMenuEx.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "TAS TRANS GROUP OF COMPANIES HR and Payroll System"
   ClientHeight    =   9630
   ClientLeft      =   1935
   ClientTop       =   2070
   ClientWidth     =   15135
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   11640
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":15F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1B90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":29E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":32BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":370E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":45E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5314
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":701E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":78F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A0AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A4FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A656
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":AAA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":AEFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":B214
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":B666
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":BAB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":BF0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":C35C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":D1AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tlbrmnu 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   900
      ButtonWidth     =   767
      ButtonHeight    =   741
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Employee Information"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Payroll Transactions"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Employee Details"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Client Master"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "HR Search"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "HR Reports"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Time Keeping"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin PSVisualMenuEx.VisualMenuEx VisualMenuEx1 
      Left            =   11760
      Top             =   2040
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   6
      Bmp:1           =   "mdiMain.frx":DA88
      Key:1           =   "#ClientMaster"
      Bmp:2           =   "mdiMain.frx":DEB0
      Key:2           =   "#PayrollSummaryReport"
      Bmp:3           =   "mdiMain.frx":E2D8
      Key:3           =   "#ChangePassword"
      Bmp:4           =   "mdiMain.frx":E700
      Key:4           =   "#PayrollTransactions"
      Bmp:5           =   "mdiMain.frx":EB28
      Key:5           =   "#mnuCustomizedReport"
      Mask:6          =   16711935
      Key:6           =   "#mnuEmpDetails"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   11880
      Top             =   2760
   End
   Begin MSComctlLib.StatusBar statbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9255
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12876
            MinWidth        =   12876
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu Login 
         Caption         =   "Log in"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLogOut 
         Caption         =   "&Log out"
      End
      Begin VB.Menu ChangePassword 
         Caption         =   "Change password"
      End
      Begin VB.Menu mnuCopyEmployee 
         Caption         =   "Copy Employee"
         Visible         =   0   'False
      End
      Begin VB.Menu sep90 
         Caption         =   "-"
      End
      Begin VB.Menu ShowToolbar 
         Caption         =   "Show Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu ShowStatusbar 
         Caption         =   "Show Status bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHumanResources 
      Caption         =   "Human Resources"
      Begin VB.Menu mnuFilterEmp 
         Caption         =   "Filter Employee"
         Visible         =   0   'False
      End
      Begin VB.Menu Sep7 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu HREmployees 
         Caption         =   "Employee Information"
      End
      Begin VB.Menu mnuEmpDetails 
         Caption         =   "Employee Details"
      End
      Begin VB.Menu SS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBenefits 
         Caption         =   "Benefits"
      End
      Begin VB.Menu mnuAssignment 
         Caption         =   "Assignments"
      End
      Begin VB.Menu mnuTraining 
         Caption         =   "Trainings"
      End
      Begin VB.Menu mnuOffenses 
         Caption         =   "Offenses"
      End
      Begin VB.Menu mnuLeaves 
         Caption         =   "Leaves"
      End
      Begin VB.Menu mnuAppraisal 
         Caption         =   "Appraisals"
      End
   End
   Begin VB.Menu mnuPayroll 
      Caption         =   "Payroll"
      Begin VB.Menu PayrollTransactions 
         Caption         =   "Payroll Processing"
      End
      Begin VB.Menu PayrollSummaryReport 
         Caption         =   "Payroll Summary Report"
         Enabled         =   0   'False
      End
      Begin VB.Menu dashpay 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayMaster 
         Caption         =   "Employee Payroll Master"
      End
      Begin VB.Menu payted 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayrollReports 
         Caption         =   "Payroll Reports"
      End
   End
   Begin VB.Menu mnuAdministrative 
      Caption         =   "Administrative"
      Begin VB.Menu UserAccounts 
         Caption         =   "User Accounts"
      End
      Begin VB.Menu dash33 
         Caption         =   "-"
         Index           =   33
      End
      Begin VB.Menu UserTrans 
         Caption         =   "User Transactions/History"
      End
      Begin VB.Menu dash35 
         Caption         =   "-"
         Index           =   35
      End
      Begin VB.Menu CompanyInformation 
         Caption         =   "Company Information"
      End
      Begin VB.Menu ClientMaster 
         Caption         =   "Client Master"
      End
      Begin VB.Menu ddss 
         Caption         =   "-"
      End
      Begin VB.Menu EmployeeMaster 
         Caption         =   "Employee Payroll Master"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShiftSched 
         Caption         =   "Shift Schedule"
      End
   End
   Begin VB.Menu mnuTimeKeep 
      Caption         =   "&Time Keep"
      Begin VB.Menu ImportDTR 
         Caption         =   "Import DTR Records"
      End
      Begin VB.Menu mnuEmpTimeKeep 
         Caption         =   "Employee Time Keeping"
      End
      Begin VB.Menu dash34 
         Caption         =   "-"
         Index           =   34
      End
      Begin VB.Menu HolTable 
         Caption         =   "Holiday Table"
      End
      Begin VB.Menu LeaveDayoff 
         Caption         =   "Leave and Dayoff"
      End
      Begin VB.Menu holdash 
         Caption         =   "-"
         Index           =   22
      End
      Begin VB.Menu EmpDed 
         Caption         =   "Employee Deductions"
      End
      Begin VB.Menu test 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu fmsGross 
         Caption         =   "Driver/Conductor Contribution"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuReporting 
      Caption         =   "&Reports"
      Begin VB.Menu mnuCustomizedReport 
         Caption         =   "HR Reports"
      End
      Begin VB.Menu sep9 
         Caption         =   "-"
      End
      Begin VB.Menu EOC 
         Caption         =   "List of EOC"
         Visible         =   0   'False
      End
      Begin VB.Menu dd 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu PayrollReports 
         Caption         =   "Payroll Reports"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu sep10 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu DedMonitoring 
         Caption         =   "Payroll Reports"
      End
      Begin VB.Menu sep11 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu PhilhealthReport 
         Caption         =   "PhilHealth Report"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu SSSBankReport 
         Caption         =   "SSS Report"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu HDMFReport 
         Caption         =   "HDMF Report"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu bill1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu bill2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBillingReports 
         Caption         =   "Billing Reports"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMaintenance 
      Caption         =   "Maintenance"
      Begin VB.Menu mnubank 
         Caption         =   "Bank"
      End
      Begin VB.Menu mnuBranch 
         Caption         =   "Branch"
      End
      Begin VB.Menu mnuDepartment 
         Caption         =   "Department"
      End
      Begin VB.Menu sep34 
         Caption         =   "-"
      End
      Begin VB.Menu sssTable 
         Caption         =   "SSS table"
      End
      Begin VB.Menu PagibigTable 
         Caption         =   "Pag-ibig table"
      End
      Begin VB.Menu PhilHealthtable 
         Caption         =   "PhilHealth table"
      End
      Begin VB.Menu Taxtable 
         Caption         =   "Tax table"
      End
      Begin VB.Menu TaxExemptionTable 
         Caption         =   "Tax Exemption Table"
      End
   End
   Begin VB.Menu winlist 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnHelp 
      Caption         =   "&About"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------
'PROGRAM : TASTRANS
'DEVELOPER : Bernard O. Sahagun
'DATE DEVELOPED : 022505
'STATUS: TO ENHANCE
'------------------------------------

Option Explicit

Private Sub ChangePassword_Click()
frmChange.Show
frmChange.SetFocus
End Sub

Private Sub ClientMaster_Click()
frmClientTable.Show
frmClientTable.SetFocus
End Sub

Private Sub CompanyInformation_Click()
frmCompany.Show
End Sub

Private Sub DedMonitoring_Click()
FrmOtherReport.Show
FrmOtherReport.SetFocus
'frmDedMonitoring.Show
'frmDedMonitoring.SetFocus
End Sub

Private Sub DriverCon_Click()
frmFMSDeduction.Show
frmFMSDeduction.SetFocus
End Sub

Private Sub EmpDed_Click()
frmEmpDeductions.Show
frmEmpDeductions.SetFocus
End Sub

Private Sub EmployeeMaster_Click()
frmBasicSalary.Show
frmBasicSalary.SetFocus
End Sub

Private Sub EOC_Click()
MsgBox "Not Yet available!", vbInformation, PayrollAlert
'frmHrEOC.Show
'frmHrEOC.SetFocus
End Sub

Private Sub fmsGross_Click()
frmFMSDeduction.Show
frmFMSDeduction.SetFocus
'frmFMSBonus.Show
'frmFMSBonus.SetFocus
End Sub

Private Sub HDMFReport_Click()
    frmHDMFReport.Show
    frmHDMFReport.SetFocus
End Sub

Private Sub HolidayTable_Click()
    frmHolidays.Show
    frmHolidays.SetFocus
End Sub

Private Sub Holtable_Click()
    frmHolidays.Show
    frmHolidays.SetFocus
End Sub

Private Sub HREmployees_Click()
Dim rsout As New ADODB.Recordset

FeMP = True

If rsout.State = 1 Then rsout.Close
    
StrQuery = "SELECT * FROM HREmployees WHERE agencycode = " & CompanyCode & " order by surname, firstname"
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HREmployees where agencycode=" & CompanyCode & " order by surname, firstname", cnnHR, adOpenKeyset, adLockOptimistic
If rsout.EOF = False Then
    Call frmHRFinalEmployees.Showdata(rsout)
Else
    MsgBox "Record not found..pls add new employee.", vbInformation
    AddFlag = True
    frmHRFinalEmployees.Show
    Exit Sub
End If

End Sub

Private Sub ImportDTR_Click()
frmImportDTR.Show
frmImportDTR.SetFocus
End Sub

Private Sub LeaveDayoff_Click()
    Call frmLeaveDayoff.Show
    frmLeaveDayoff.SetFocus
End Sub

Private Sub Login_Click()
frmLogin.Show
End Sub

Public Sub InitializeMdi()
   
   On Error GoTo Err
   'mdiMain.Picture = LoadPicture(App.Path & "\logos\taslogo.gif")
   'mdiMain.Picture = LoadPicture(App.Path & "\logos\" & CompanyCode & " background.gif")

Err:

StrDbpath = GetSetting("DataSearch", CompanyShort, "Path2Version2")

      CompanyName = getAgencyNamebyID(CompanyCode)
      mdiMain.Caption = CompanyName
      mUserID = Trim(UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID))))
      
        If LoginName = "EVA" Then
            Call frmUserTransactions.Show
            frmUserTransactions.SetFocus
        Else
            frmHRSearch.Show
            frmHRSearch.SetFocus
            'frmHrEOC.Show
            'frmHrEOC.SetFocus
        End If

   ' statbar.Panels(1).Text = FormatDateTime(Now, vbLongDate) & "    " & FormatDateTime(Now, vbLongTime)
     statbar.Panels(1).Text = "DATABASE : " & StrDbpath
   ' Timer1.Interval = 1000

End Sub

Private Sub mnTa_Click()
    frmHRTraining.Show
    frmHRTraining.SetFocus
End Sub

Private Sub MDIForm_Resize()

If mHRSEARCH = True Then
If frmHRSearch.WindowState = 1 Then Exit Sub
    If mdiMain.Height > 3000 Then
'Call frmHRSearch.cmdArrange_Click
If Me.Height <> frmHRSearch.ScaleHeight Then

'    With frmHRSearch
'
'    .Frame1.Height = Me.Height - 500
'    .lvwdisplay.Height = Me.Height - 1600
'    .cmbClient.Top = Me.Height - 700
'    .cmbClient.Width = 2200
'    .lvwdisplay.Width = Me.Width - 100
'    .lvwdisplay.ColumnHeaders(1).Width = 2000
'    .lvwdisplay.ColumnHeaders(2).Width = 0
'    .lvwdisplay.ColumnHeaders(3).Width = 0
'    .lvwdisplay.ColumnHeaders(4).Width = 0
'    .lvwdisplay.ColumnHeaders(5).Width = 0
'    .fraSearch.Left = 0
'    .txtSearch.Width = 2200
'    .Label2.Top = 120
'    .Label2.Left = 40
'    .txtSearch.Left = 60
'    .txtSearch.Top = 350
'    .cmdArrange.Top = 120
'    .cmdArrange.Left = 1890
'    .cmdSettings.Top = 30
'    .cmdSettings.Left = 1485
'    'Call cmdSettings_Click
'.FraSettings.Width = 2355
'
'End With

    frmHRSearch.Height = Me.ScaleHeight
    frmHRSearch.Left = Me.ScaleWidth - frmHRSearch.Width
    If frmHRSearch.Height > 1600 Then
        frmHRSearch.Frame1.Height = frmHRSearch.Height - 500
    frmHRSearch.lvwdisplay.Height = frmHRSearch.Height - 1600
    frmHRSearch.cmbClient.Top = frmHRSearch.Height - 700
    frmHRSearch.cmbClient.Width = 2200

    End If

    End If

End If
 
End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Call SaveLoginStatus(0, LoginName, LoginPass)
    LogoutTime = Format(Now, "mm/dd/yyyy hh:mm:ss")
    Call SaveUserTransactions("System Closed.", "Closed.")
    End
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
    frmAbout.SetFocus
End Sub

Private Sub mnuAppraisal_Click()
    frmHRAppraisal.Show
    frmHRAppraisal.SetFocus
End Sub

Private Sub mnuAssignment_Click()
    frmHRProjAssign.Show
    frmHRProjAssign.SetFocus
End Sub

Private Sub mnubank_Click()
frmBank.Show
frmBank.SetFocus
End Sub

Private Sub mnuBenefits_Click()
    frmHRBenefits.Show
    frmHRBenefits.SetFocus
End Sub

'Private Sub mnuBillingReports_Click()
'frmBilling.Show
'End Sub

Private Sub mnuBranch_Click()
'frmBranch.Show
'frmBranch.SetFocus
End Sub

Private Sub mnuCopyEmployee_Click()
'frmHRCopy.Show
End Sub

Private Sub mnuCustomizedReport_Click()
FrmReportCustomized.Show
FrmReportCustomized.SetFocus
End Sub

Private Sub mnuDbSettings_Click()
FrmSetDatabase.Show
FrmSetDatabase.SetFocus
End Sub

Private Sub mnuDepartment_Click()
frmDept.Show
frmDept.SetFocus
End Sub

Private Sub mnuEmpBenefits_Click()
frmHRBenefitSub.Show
frmHRBenefitSub.SetFocus
End Sub

Private Sub mnuEmpDetails_Click()
MMEmp = True
frmHRMainEmployee.Show
frmHRMainEmployee.SetFocus
End Sub

Private Sub mnuEmpTimeKeep_Click()
Call frmTimeKeep.Show
frmTimeKeep.SetFocus
End Sub

Private Sub mnuExit_Click()

    If MsgBox("Exit HR and Payroll System?", vbCritical + vbOKCancel, CompanyName) = vbOK Then
        Call SaveLoginStatus(0, LoginName, LoginPass)
        
        LogoutTime = Format(Now, "mm/dd/yyyy hh:mm:ss")
        Call SaveUserTransactions("System exit.", "Exit.")
        
        End
    End If
    
End Sub

Private Sub mnuLogIn_Click()
    frmLogin.Show 1
End Sub

Private Sub mnuQuit_Click()
    If MsgBox("Are you sure you want to quit?", vbCritical + vbOKCancel, CompanyName) = vbOK Then Unload Me
End Sub

Private Sub mnuFilterEmp_Click()
frmHRFilter.Show
frmHRFilter.SetFocus
End Sub

Private Sub mnuLeaves_Click()
frmHRLeave.Show
frmHRLeave.SetFocus
End Sub

Private Sub mnuLogOut_Click()
    If MsgBox("Are you sure you want to logout?", vbYesNo + vbQuestion) = vbYes Then
        Call UnloadAll
        'Unload Me
        Call SaveLoginStatus(0, LoginName, LoginPass)   ' updated 11/7/2009
        
        LogoutTime = Format(Now, "mm/dd/yyyy hh:mm:ss")
        Call SaveUserTransactions("Logout.", "Logout Time.")
        
        AddFlag = False
        addflagship = False
        On Error Resume Next
        If FrmReportCustomized.Visible = True Then Unload FrmReportCustomized
        frmLogin.Show
        mdiMain.Hide
    End If
End Sub

Public Sub CloseAllWindows()
    On Error Resume Next
    Unload frmHRFinalEmployees
    Dim i  As Long
    For i = 0 To Forms.Count - 1
          If Forms(i) Is Me Then
                'Already unloading
          Else
                 Unload Forms(i)
          End If
    Next i

End Sub

Private Sub mnuOffenses_Click()
frmHROffense.Show
frmHROffense.SetFocus
End Sub

Private Sub mnuPayMaster_Click()
frmBasicSalary.Show
frmBasicSalary.SetFocus
End Sub

Private Sub mnuPayrollReports_Click()
FrmOtherReport.Show
FrmOtherReport.SetFocus
End Sub

Private Sub mnuShiftSched_Click()
    Call frmShiftSchedule.Show
frmShiftSchedule.SetFocus
End Sub

Private Sub mnuTraining_Click()
    frmHRTraining.Show
    frmHRTraining.SetFocus
End Sub

Private Sub PagibigTable_Click()
    frmPagibig.Show
    frmPagibig.SetFocus
End Sub

'Private Sub PayrollDed_Click()
'    Call frmSalaryDeductions.Show
'    Call frmSalaryDeductions.SetFocus
'End Sub

Private Sub PayrollReports_Click()
Call FrmOtherReport.Show
FrmOtherReport.SetFocus
End Sub

Private Sub PayrollTransactions_Click()
    frmSelectCutoff.Show
    frmSelectCutoff.SetFocus
End Sub

Private Sub PhilhealthReport_Click()
    PHealthFlag = True
    frmPHealthReport.Show
    frmPHealthReport.SetFocus
End Sub

Private Sub PhilHealthtable_Click()
    frmPhilHealth.Show
    frmPhilHealth.SetFocus
End Sub

Private Sub ShowStatusbar_Click()
If ShowStatusbar.Checked = True Then
ShowStatusbar.Checked = False
Me.statbar.Visible = False
Else
ShowStatusbar.Checked = True
Me.statbar.Visible = True
End If

End Sub

Private Sub ShowToolbar_Click()
If ShowToolbar.Checked = True Then
    ShowToolbar.Checked = False
    Tlbrmnu.Visible = False
Else
    ShowToolbar.Checked = True
    Tlbrmnu.Visible = True
End If
End Sub

Private Sub SSSReports_Click()
    FrmOtherReport.Show
    FrmOtherReport.SetFocus
End Sub

Private Sub SSSBankReport_Click()
    PHealthFlag = False
    frmSSSReport.Show
    frmSSSReport.SetFocus
End Sub

Private Sub sssTable_Click()
    frmSSS.Show
    frmSSS.SetFocus
End Sub

Private Sub StatementOfAccounts_Click()
    frmPrintPayslip.Show
    frmPrintPayslip.SetFocus
End Sub

Private Sub TaxExemptionTable_Click()
    frmTaxExemption.Show
    frmTaxExemption.SetFocus
End Sub

Private Sub Taxtable_Click()
    FrmTaxTable.Show
    FrmTaxTable.SetFocus
End Sub

Private Sub Timer1_Timer()
'    statbar.Panels(1).Text = FormatDateTime(Now, vbLongDate) & "    " & FormatDateTime(Now, vbLongTime)
End Sub

Private Sub tlbrMnu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            FeMP = True
            Call HREmployees_Click
        Case 3
            frmSelectCutoff.Show
            frmSelectCutoff.SetFocus
        Case 5
            frmHRMainEmployee.Show
            frmHRMainEmployee.SetFocus
        Case 7
            frmClientTable.Show
            frmClientTable.SetFocus
        Case 9
            frmHRSearch.Show
            frmHRSearch.SetFocus
'        Case 11
'            frmSelectCutoff.Show
'            frmPrintPayslip.Show
        Case 11
            FrmReportCustomized.Show
            FrmReportCustomized.SetFocus
        Case 13
            frmTimeKeep.Show
            frmTimeKeep.SetFocus
'        Case 15
'            FrmSetDatabase.Show
'            FrmSetDatabase.SetFocus
'        Case 17
'            frmUsers.Show
'            frmUsers.SetFocus
'        Case 19
'            FrmSetDatabase.Show
'            FrmSetDatabase.SetFocus
        Case 21
             MsgBox "Help not yet available!"
            '     Case 23
           ' MsgBox "help is under construction"
     End Select
End Sub

Private Sub UserAccounts_Click()
frmUsers.Show
frmUsers.SetFocus
End Sub

Private Sub UserTrans_Click()
frmUserTransactions.Show
frmUserTransactions.SetFocus
End Sub

Private Sub UnloadAll()
Dim frm As Form

On Error Resume Next

For Each frm In Forms
    If frm.Name <> Me.Name Then
        If frm.MDIChild = True Then ActiveForm.CommandPass "Close"
    End If
    'If frm.Name <> Me.Name Then Unload Me
Next
Set frm = Nothing
End Sub

Public Sub CommandPass(ByVal strCommand As String)

Select Case strCommand
    Case "Close"
        Unload Me
End Select

End Sub


