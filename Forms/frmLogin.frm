VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3315
   ClientLeft      =   2250
   ClientTop       =   3030
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1954.667
   ScaleMode       =   0  'User
   ScaleWidth      =   5408.329
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDeductions 
      Caption         =   "Ded&uctions"
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdLeaveForm 
      Caption         =   "&Leave Form"
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   2137
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdDTRImport 
      Caption         =   "&DTR Import"
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   2490
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdDTREntries 
      Caption         =   "DTR &Entries"
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   2850
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.ComboBox cboAgency 
      Height          =   330
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   "Select agency here"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   315
      Left            =   4350
      TabIndex        =   5
      Top             =   2850
      Width           =   1005
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3180
      TabIndex        =   4
      Top             =   2850
      Width           =   1005
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save password"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3810
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3315
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3300
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1305
      Width           =   2175
   End
   Begin VB.CommandButton cmdDeactivate 
      Caption         =   "&Deactivate User"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   2850
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Agency:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   2160
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "HR AND PAYROLL SYSTEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   660
      TabIndex        =   8
      Top             =   540
      Width           =   2130
   End
   Begin VB.Image imgLogo 
      Height          =   360
      Left            =   210
      Picture         =   "frmLogin.frx":27A2
      Stretch         =   -1  'True
      Top             =   405
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FADAD1&
      Height          =   420
      Left            =   165
      Top             =   360
      Width           =   420
   End
   Begin VB.Label LblCompany 
      BackStyle       =   0  'Transparent
      Caption         =   "TAS TRANS GROUP OF COMPANIES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   615
      TabIndex        =   7
      Top             =   270
      Width           =   5145
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   0
      Picture         =   "frmLogin.frx":306C
      Top             =   765
      Width           =   315
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   570
      Picture         =   "frmLogin.frx":372E
      Top             =   720
      Width           =   4425
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   900
      Picture         =   "frmLogin.frx":56A8
      Top             =   1455
      Width           =   720
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   2175
      TabIndex        =   6
      Top             =   1305
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   2175
      TabIndex        =   2
      Top             =   1665
      Width           =   1080
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FADAD1&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   8
      FillColor       =   &H00FADAD1&
      Height          =   960
      Left            =   540
      Top             =   495
      Width           =   1005
   End
   Begin VB.Image Image5 
      Height          =   630
      Left            =   1395
      Picture         =   "frmLogin.frx":6572
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------
'PROGRAM : HR AND PAYROLL SYSTEM
'DEPARTMENT : HR
'DATE DEVELOPED : 022505
'UPDATED BY: Bernard O. Sahagun
'DATE OF UPDATES: January 9, 2010

Option Explicit

Dim ctr As Integer
Dim firstRun As String

Private Sub cboAgency_Click()
    LblCompany.Caption = getAgencyName(cboAgency.Text)
End Sub

Private Sub chkConnection_Click()
    FrmSetDatabase.Show
    FrmSetDatabase.SetFocus
End Sub

Private Sub chkDTREntries_Click()

End Sub

Private Sub cmdDeactivate_Click()

Call SaveLoginStatus(0, txtUserName.Text, txtPassword.Text)
MsgBox "Account " & txtUserName.Text & " was successfully deactivated!", vbInformation, PayrollAlert
txtUserName.Text = ""
txtPassword.Text = ""
cmdDeactivate.Visible = False

End Sub

Private Sub cmdDeductions_Click()
Call frmEmpDeductions.Show
End Sub

Private Sub cmdDTREntries_Click()
If cmdDTREntries.Caption = "FMS Records" Then
    Call frmFMSDeduction.Show
Else
    Call frmTimeKeep.Show
End If
End Sub

Private Sub cmdDTRImport_Click()
Call frmImportDTR.Show
End Sub

Private Sub cmdLeaveForm_Click()
Call frmLeaveDayoff.Show
End Sub

Private Sub Form_Load()

'CnnstrHR = cnnstr()
Call LoadAgency
'LblCompany.Caption = getAgencyName(cboAgency.Text)
'CompanyShort = cboAgency.Text
'CompanyName = LblCompany.Caption
    Me.Picture = LoadPicture(App.Path & "\pics\bg.gif")
    
    On Error Resume Next
    firstRun = GetSetting("DataSearch", "Data", "firstRun")
    If firstRun = "" Then 'first run of the system
        SaveSetting "DataSearch", "Data", "firstRun", "YES"
        SaveSetting "DataSearch", "Data", "check", "False"
        SaveSetting "DataSearch", "Data", "username", ""
        SaveSetting "DataSearch", "Data", "password", ""
    Else
        Check1.VALUE = GetSetting("DataSearch", "Data", "check")   'save check
        txtUserName = GetSetting("DataSearch", "Data", "username") 'save username
        txtPassword = GetSetting("DataSearch", "Data", "password") 'password
    End If
End Sub

Private Sub cmdCancel_Click()
  End
End Sub

Private Sub LoadAgency() 'Get information from Agency
Dim rstAgency As New ADODB.Recordset

StrQuery = "select agencyname, agencyshort, agencycode from companytable where agencystatus = 1 order by agencycode"
Set rstAgency = ReturnRst(StrQuery, cnnstr, mstatic)
'Set rstAgency = returnrst("select agencyname, agencyshort, agencycode from companytable where agencystatus = 1 order by agencycode", CnnstrHR, mstatic)
CompanyName = rstAgency!AgencyName
CompanyCode = rstAgency!agencycode
CompanyShort = rstAgency!AgencyShort
LblCompany.Caption = rstAgency!AgencyName ' updated July 28, 2011

'Do While Not rstAgency.EOF
'    Me.cboAgency.AddItem rstAgency!AgencyShort
'    cboAgency.ItemData(cboAgency.NewIndex) = rstAgency!AgencyCode
'    rstAgency.MoveNext
'Loop
'cboAgency.ListIndex = 0

End Sub

Private Function checkCompanyCode() As Boolean
    Dim strcondb    As String
    Dim intStrDbpath   As Integer
    Dim intStrDbHRpath As Integer
    
'    CompanyCode = cboAgency.ItemData(cboAgency.ListIndex) ' cboAgency.ListIndex + 1 'GETINI("Company", "CompanyName", "", App.Path & "\Application.ini")
'    CompanyShort = cboAgency.Text
'    StrDbpath = GetSetting("DataSearch", CompanyShort, "Path2Version2")
    
'    cnnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrDbpath
'    CnnstrHR = cnnstr() '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrDbHRpath
     
    logoPath = PublicFunctions.getLogoPath(CompanyCode)
    
    'Cnnstr = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & StrDbpath & ";Uid=;Pwd=alpha;"
    intStrDbpath = OpenDatabaseAdo(cnn, cnnstr)
    intStrDbHRpath = OpenDatabaseAdo(cnn, CnnstrHR)

    If intStrDbpath <> 0 Or intStrDbHRpath <> 0 Then
        MsgBox "Unable To Establish the Connection. ", vbCritical
        checkCompanyCode = False
        Exit Function
    Else
        checkCompanyCode = True
    End If
End Function

Private Sub cmdOK_Click()

    If cmdOK.Caption = "Quit" Then
        cmdOK.Caption = "Ok"
        End
    End If
    If checkCompanyCode = False Then
        FrmSetDatabase.cmdOK.Enabled = False
        FrmSetDatabase.Show
    Else
        MMEmp = False
        FeMP = False
        
        'ConnectDB
        Dim rs_log As New ADODB.Recordset
        Set rs_log = ReturnRst("Select Username, Password, operatorid, superuser, LoginStatus, AccountType, administrativerights, AddRight, EditRight, DeleteRight from Userslist " & _
                    " Where Username= '" & txtUserName.Text & _
                    "' And Password = '" & txtPassword.Text & "'", cnnstr, mstatic)
        If Not rs_log.EOF Then
            If rs_log!LoginStatus = True Then 'Active
                MsgBox "YOUR ACCOUNT IS ACTIVE!" & Chr(13) & Chr(13) & _
                    "Somebody else is using your account, if you properly logout last time!" & Chr(13) & _
                    "Otherwise, click Deactivate ID, to gain access.", vbInformation, PayrollAlert
                    cmdDeactivate.Visible = True
                GoTo ExitLogin1
            End If
            mdiMain.statbar.Panels(2).Text = "USER : " & Trim(txtUserName.Text) & ""
            Call pWord_Save
            Call CheckUser(rs_log.Fields("operatorid"))
            Call SaveLoginStatus(1, txtUserName.Text, txtPassword.Text)
            
            LoginTime = Format(Now, "mm/dd/yyyy hh:mm:ss")
            LogoutTime = ""
            AccountType = IIf(IsNull(rs_log.Fields("AccountType")), "HR", rs_log.Fields("AccountType"))
            
            operatorID = rs_log.Fields("operatorid")
            SuperUser = rs_log.Fields("superuser")
            AdminRights = rs_log.Fields("administrativerights")
            LoginName = rs_log.Fields("UserName")
            LoginPass = rs_log.Fields("Password")
            AddRights = IIf(rs_log.Fields("Addright") = True, True, False)
            EditRights = IIf(rs_log.Fields("Editright") = True, True, False)
            DeleteRights = IIf(rs_log.Fields("DeleteRight") = True, True, False)
            NewCutoff = 15  ' set the cutoff to old cut-off of 1-15 and 16-30
            
            'Call SaveUserTransactions("test message", "Test", "Bernard")
            
'            If AccountType = "TK" Then  ' for Time Keeping
'                txtusername.Enabled = False
'                Me.Caption = "DTR Records"
'                lblLabels(1).Visible = False
'                txtpassword.Visible = False
'                Image2.Visible = False
'                cmdcancel.Visible = False
'                cmdOK.Caption = "Quit"
'                cmdDTREntries.Visible = True
'                cmdDTRImport.Visible = True
'                cmdLeaveForm.Visible = True
'                cmdDeductions.Visible = True
'            ElseIf AccountType = "FMS" Then ' for FMS
'                txtusername.Enabled = False
'                Me.Caption = "FMS Records"
'                lblLabels(1).Visible = False
'                txtpassword.Visible = False
'                Image2.Visible = False
'                cmdcancel.Visible = False
'                cmdOK.Caption = "Quit"
'                cmdDTREntries.Caption = "FMS Records"
'                cmdDTREntries.Visible = True
'                cmdDTRImport.Visible = False
'                cmdLeaveForm.Visible = False
'                cmdDeductions.Visible = False
'            Else
                txtUserName.Enabled = True
                Me.Caption = "Login"
                lblLabels(1).Visible = True
                txtPassword.Visible = True
                Image2.Visible = True
                cmdcancel.Visible = True
                cmdOK.Caption = "Ok"
                
                cmdDTREntries.Visible = False
                cmdDTRImport.Visible = False
                cmdLeaveForm.Visible = False
                cmdDeductions.Visible = False

                Unload Me
                mdiMain.CloseAllWindows
                mdiMain.InitializeMdi
                mdiMain.Show
'            End If
            
         Else
            MsgBox "Login failed.", vbCritical + vbOKOnly, CompanyName
            txtUserName.SetFocus
            ctr = ctr + 1
            If ctr > 3 Then MsgBox "Too many Invalid Logon Attempts !", vbCritical + vbOKOnly, "PROGRAM TERMINATED": Set rs_log = Nothing: End
        End If
        If rs_log.State = 1 Then rs_log.Close
        Set rs_log = Nothing
    End If

    Exit Sub
    
ExitLogin1:
        
    If rs_log.State = 1 Then rs_log.Close
    Set rs_log = Nothing

End Sub

Public Function CheckUser(ByVal operatorID As Integer)
 Dim RsUsers As ADODB.Recordset
 
 Set RsUsers = New ADODB.Recordset
 
 StrQuery = "Select * from userslist where operatorid = " & operatorID
 Set RsUsers = ReturnRst(StrQuery, cnnstr, mstatic)
 'RsUsers.Open "select * from userslist where operatorid=" & operatorID & "", cnnstr, adOpenDynamic, adLockOptimistic
 
 Call AbleMenu(False, False, False, False, False, False, False, False, False, False)
 With RsUsers
     If RsUsers.Fields("superuser") = True Then
        Call AbleMenu(True, True, True, True, True, True, True, True, True, True)
        Exit Function
     Else
        Call AbleMenu(.Fields("payrollrights"), .Fields("employeeinforights"), _
            .Fields("administrativerights"), .Fields("reportrights"), _
            .Fields("maintenancerights"), .Fields("clientmaster"), _
            .Fields("employeemaster"), .Fields("useraccounts"), _
            .Fields("TimeKeeping"), .Fields("connectionsettings"))
     End If
 End With
End Function

Public Sub AbleMenu(Optional ByVal mPay As Boolean, Optional ByVal mEmp As Boolean, _
    Optional ByVal c As Boolean, Optional ByVal mRep As Boolean, _
    Optional ByVal mMaintenance As Boolean, Optional ByVal mClient As Boolean, _
    Optional ByVal mMaster As Boolean, Optional ByVal mUser As Boolean, _
    Optional ByVal mTime As Boolean, Optional ByVal j As Boolean)

    mdiMain.mnuPayroll.Enabled = mPay
    mdiMain.mnuPayMaster.Enabled = mPay
    mdiMain.Tlbrmnu.Buttons(3).Enabled = mPay
    mdiMain.Tlbrmnu.Buttons(7).Enabled = mClient
    
    'Administrative
    mdiMain.mnuHumanResources.Enabled = mEmp
    mdiMain.Tlbrmnu.Buttons(1).Enabled = mEmp
    mdiMain.Tlbrmnu.Buttons(5).Enabled = mEmp
    mdiMain.Tlbrmnu.Buttons(9).Enabled = mEmp
    
    'mdiMain.mnuDbSettings.Enabled = i
    mdiMain.UserAccounts.Enabled = mUser
    mdiMain.CompanyInformation.Enabled = mUser
    mdiMain.PhilhealthReport.Enabled = mUser
    mdiMain.SSSBankReport.Enabled = mUser
    mdiMain.HDMFReport.Enabled = mUser
    mdiMain.Tlbrmnu.Buttons(11).Enabled = mUser
    
    mdiMain.mnuShiftSched.Enabled = mUser
    mdiMain.EmployeeMaster.Enabled = mMaster
    
    mdiMain.UserTrans.Enabled = mUser   ' for administrative rights only
    
    mdiMain.ClientMaster.Enabled = mClient
  '  mdiMain.mnuAdministrative.Enabled = c
    mdiMain.mnuReporting.Enabled = mRep
    mdiMain.mnuMaintenance.Enabled = mMaintenance
    mdiMain.DedMonitoring.Enabled = mRep
    mdiMain.PhilhealthReport.Enabled = mRep
    
    mdiMain.fmsGross.Enabled = True 'mUser
    mdiMain.mnuMaintenance.Enabled = mUser
    
    mdiMain.mnuCustomizedReport.Enabled = mPay
    'mdiMain.Tlbrmnu.Buttons(11).Enabled = mPay     ' hr reports
    mdiMain.Tlbrmnu.Buttons(13).Enabled = mTime     ' time keeping
    'mdiMain.Tlbrmnu.Buttons(15).Enabled = i
    'mdiMain.Tlbrmnu.Buttons(17).Enabled = h
    mdiMain.mnuTimeKeep.Enabled = mTime
    mdiMain.Tlbrmnu.Buttons(13).Enabled = mTime     ' time keeping
    
End Sub

Private Sub pWord_Save()

    If Check1.VALUE = 1 Then
        SaveSetting "DataSearch", "Data", "check", "0"  ' Check1.VALUE, do not save password, Bernard
        SaveSetting "DataSearch", "Data", "username", "" 'txtusername.Text
        SaveSetting "DataSearch", "Data", "password", "" 'txtpassword.Text
    Else
        SaveSetting "DataSearch", "Data", "check", "0"
        SaveSetting "DataSearch", "Data", "username", ""
        SaveSetting "DataSearch", "Data", "password", ""
    End If
    
End Sub

Private Sub txtpassword_GotFocus()
    txtPassword.BackColor = &HFFFFC0
    txtPassword.ForeColor = vbBlue
    Call txtHLigths(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub txtpassword_LostFocus()
    txtPassword.BackColor = vbWhite
End Sub

Private Sub txtusername_GotFocus()
    txtUserName.BackColor = &HFFFFC0
    txtUserName.ForeColor = vbBlue
    Call txtHLigths(txtUserName)
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
 If KeyAscii = 13 Then SendKeys "{tab}"   ' BOSSSS
End Sub

Private Sub txtusername_LostFocus()
txtUserName.BackColor = vbWhite
End Sub

