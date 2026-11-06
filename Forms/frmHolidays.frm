VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHolidays 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Holiday Entry"
   ClientHeight    =   6540
   ClientLeft      =   5475
   ClientTop       =   4155
   ClientWidth     =   5430
   Icon            =   "frmHolidays.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   5430
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   5925
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5205
      Begin VB.CheckBox chkAllEmployees 
         BackColor       =   &H80000009&
         Caption         =   "All Employees"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   5040
         Value           =   1  'Checked
         Width           =   1275
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
         ItemData        =   "frmHolidays.frx":08CA
         Left            =   1560
         List            =   "frmHolidays.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   4965
         Width           =   3480
      End
      Begin VB.ComboBox cmbClient 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   4560
         Width           =   4905
      End
      Begin VB.TextBox txtFlag 
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   4560
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   3960
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply for DTR Entry"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Last record"
         Top             =   5400
         Width           =   4935
      End
      Begin VB.TextBox txtHolidayName 
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   645
         Left            =   1320
         TabIndex        =   14
         Top             =   1320
         Width           =   3720
      End
      Begin VB.ComboBox cmbHolidayType 
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
         ItemData        =   "frmHolidays.frx":08CE
         Left            =   1320
         List            =   "frmHolidays.frx":08D0
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   2445
      End
      Begin MSComCtl2.DTPicker dtHoliday 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   19988481
         CurrentDate     =   40359
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   3413
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   9
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Holiday Type"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Holiday Date"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Holiday Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   9
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   9
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   9
         EndProperty
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Client"
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   4320
         Width           =   1290
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Holiday Name"
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
         Height          =   255
         Left            =   135
         TabIndex        =   10
         Top             =   1260
         Width           =   1020
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Holiday List"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   2025
         Width           =   1290
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Holiday Date"
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
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Holiday Type"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   375
         Width           =   1065
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton cmdDelete 
         Height          =   405
         Left            =   1410
         Picture         =   "frmHolidays.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   495
         Picture         =   "frmHolidays.frx":09D4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   960
         Picture         =   "frmHolidays.frx":0B1E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   45
         Picture         =   "frmHolidays.frx":1050
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "New"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1845
         Picture         =   "frmHolidays.frx":1582
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmHolidays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsHoliday As ADODB.Recordset
Dim rsEmpBasicData As ADODB.Recordset
Dim rsClientData As ADODB.Recordset
Dim mEmpNum As String
Dim mDTRID As String
Dim StartFlag As Boolean

Private Sub cmbEmployee_LostFocus()
Dim mEmpId As Integer

On Error GoTo ExitRtn

mEmpId = Trim(cmbEmployee.ItemData(cmbEmployee.ListIndex))

If rsEmpBasicData.RecordCount <> 0 Then
    rsEmpBasicData.MoveFirst
    rsEmpBasicData.Find "EmployeeID = " & mEmpId
    If rsEmpBasicData.EOF = True Then ' does not exists
        Exit Sub
    Else
        mEmpNum = rsEmpBasicData!EmpNo
        mDTRID = rsEmpBasicData!DTRID
    End If
End If

ExitRtn:

End Sub

Private Sub cmdApply_Click()
Dim msg As String

On Error GoTo ExitRtn

If chkAllEmployees.VALUE = 0 Then
    msg = "Reflect " & UCase(txtHolidayName) & " to " & cmbEmployee.Text & " for DTR Entry?"
Else
    msg = "Reflect " & UCase(txtHolidayName) & " to all employees of " & cmbClient.Text & " for DTR Entry?"
End If

If MsgBox(msg, vbQuestion + vbYesNo) = vbYes Then
    If chkAllEmployees.VALUE = 0 Then ' single employee
        Call UpdateLeaveCalendarWithHolidayEntries(mEmpNum, True)
    Else    ' reflect to all employees
        If rsEmpBasicData.RecordCount <> 0 Then
            With rsEmpBasicData
                .MoveFirst
                Do While Not .EOF
                    Call UpdateLeaveCalendarWithHolidayEntries(!EmpNo)
                    .MoveNext
                Loop
                MsgBox "Successfully reflected to " & .RecordCount & " employees of " & cmbClient.Text, vbInformation, PayrollAlert
            End With
        Else
            MsgBox "No employees reflected for " & cmbClient.Text, vbInformation, PayrollAlert
        End If
    End If
Else
    MsgBox UCase(txtHolidayName) & " not reflected to DTR Entry!", vbInformation, PayrollAlert
End If

ExitRtn:

End Sub

Private Sub UpdateLeaveCalendarWithHolidayEntries(ByVal mEmpNo As String, _
    Optional mAlert As Boolean = False)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim rsDayoff As ADODB.Recordset
Dim mString As String
Dim mDayOff As String
Dim mOffDay As Date
Dim mCalYearMonth As String

On Error GoTo ExitRtn1

mDayOff = ""
mString = ""
mCalYearMonth = LTrim(RTrim(str(Year(dtHoliday.VALUE)))) & Right("00" & LTrim(RTrim(str(month(dtHoliday.VALUE)))), 2)
mString = " Days" & LTrim(RTrim(str(Day(dtHoliday.VALUE)))) & " = '" & Left(cmbHolidayType.Text, 1) & "', "

StrQuery = "SELECT EMPNO from HRDayoffCalendar WHERE ltrim(rtrim(empno)) = '" & mEmpNo & "' and calyearmonth = " & mCalYearMonth
Set rsDayoff = ReturnRst(StrQuery, cnnstr, mstatic)

If rsDayoff.RecordCount = 0 Then
    StrQuery = "INSERT INTO HRDayoffCalendar (EmpNo, DTRID, CalYear, CalMonth, CalYearMonth, UpdatedBy) VALUES " & _
        "(' " & mEmpNo & "', '" & mDTRID & "', " & Year(dtHoliday.VALUE) & ", " & month(dtHoliday.VALUE) & _
        ", " & mCalYearMonth & ", '" & LoginName & "')"
    Action_Query StrQuery, cnnstr
End If
rsDayoff.Close

'Update the schedules.
StrQuery = "UPDATE HRDayoffCalendar SET " & mString & " UpdatedBy = '" & LoginName & "', Updated = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'" & _
    " WHERE LTRIM(RTRIM(empno)) = '" & mEmpNo & "' and calyearmonth = " & mCalYearMonth

If Action_Query(StrQuery, cnnstr) = True Then
    If mAlert = True Then
        MsgBox UCase(txtHolidayName.Text) & " successfully reflected for DTR Entry!", vbInformation, PayrollAlert
    End If
Else
    If mAlert = True Then
        MsgBox "Failed to reflect " & UCase(txtHolidayName) & " for DTR Entry!", vbInformation, PayrollAlert
    End If
End If

ExitRtn1:

End Sub

Private Sub Form_Load()
Dim rst As ADODB.Recordset

dtHoliday.VALUE = Now
StrQuery = "SELECT LTRIM(RTRIM(HolidayType)) + '-' + LTRIM(RTRIM(HolidayDesc)) AS HolidayType, HolidayDate, HolidayName, HolidayID, LTRIM(RTRIM(HolidayType)) + '-' + LTRIM(RTRIM(HolidayDesc)), HolidayFlag FROM tblHolidays ORDER BY HolidayDate"
Set rsHoliday = ReturnRst(StrQuery, cnnstr, mstatic)
If rsHoliday.EOF = False Then
    ListView1(1).ListItems.Clear
    PopulateList ListView1(1), rsHoliday
'    ' Initialized text boxes.
'    rsHoliday.MoveFirst
'    cmbHolidayType.Text = rsHoliday!HolidayType
'    dtHoliday.VALUE = rsHoliday!HolidayDate
'    txtHolidayName.Text = rsHoliday!HolidayName
'    txtID.Text = rsHoliday!HolidayID
'    txtFlag.Text = rsHoliday!Flag
End If

StrQuery = "SELECT * FROM HRLeaveType WHERE HolidayFlag = 1"
Set rst = ReturnRst(StrQuery, cnnstr, mstatic)

cmbHolidayType.Clear
Do While Not rst.EOF
    cmbHolidayType.AddItem LTrim(RTrim(rst!leavetype)) & "-" & LTrim(RTrim(rst!leavetypedesc))
    rst.MoveNext
Loop
Set rst = Nothing
cmbHolidayType.ListIndex = 0
TextBox (False)
    
StrQuery = "SELECT * FROM clienttable WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
Set rsClientData = ReturnRst(StrQuery, cnnstr, mstatic)

StrQuery = "SELECT EmployeeID, FullName, * FROM vwEmployeeBasicSalary ORDER BY Fullname"

'StrQuery = "SELECT a.EmployeeID AS EmpID, RTRIM(LTRIM(a.surname)) + ', ' + RTRIM(LTRIM(a.firstname)) + ' ' + (CASE WHEN a.Middlename IS NULL THEN '' ELSE LEFT(a.Middlename,1) + '.' END ) AS Fullname, " & _
'        "a.*, b.Clientname " & _
'        "FROM HREmployees AS a " & _
'        "LEFT JOIN ClientTable AS b ON a.clientcode = b.clientcode " & _
'        " WHERE     (a.EmploymentStatus NOT IN ('T', 'W', 'X', 'Y', 'Z'))" & _
'        " ORDER BY a.surname, a.firstname"
Set rsEmpBasicData = ReturnRst(StrQuery, cnnstr, mstatic)

Call oCommonFunction.FillCombo(cmbClient, rsClientData) ' 9 means add "All Clients"
If rsClientData.RecordCount <> 0 Then cmbClient.ListIndex = 0

End Sub

'Private Sub FillEmployeeList(ByRef mLsEmp As ListView, _
'                        ByRef mRs As Recordset)
'Dim litem1 As ListItem
'Dim litem2 As ListItem
'
'ActiveCtr = 0
'mLsEmp.View = lvwReport
'mLsEmp.ListItems.Clear
'ListView1(0).ListItems.Clear
'While Not (mRs.EOF Or mRs.BOF)
'    'Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!EmpName), "Blank Name", mRs!EmpName))
'    Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName))
'        litem1.SubItems(1) = mRs(0)
'        ActiveCtr = ActiveCtr + 1
'    mRs.MoveNext
'Wend
'
'End Sub


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
        Call oCommonFunction.FillCombo(cmbEmployee, rsEmpBasicData, 1)
    Else
        Call oCommonFunction.FillCombo(cmbEmployee, rsEmpBasicData, 999)
    End If
End With
    
End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    If Index = 1 Then
        If ListView1(1).ListItems.Count <> 0 Then
            cmbHolidayType.Text = ListView1(1).ListItems(ListView1(1).SelectedItem.Index).ListSubItems(4).Text
            dtHoliday.VALUE = ListView1(1).ListItems(ListView1(1).SelectedItem.Index).ListSubItems(1).Text
            txtHolidayName.Text = Trim$(ListView1(1).ListItems(ListView1(1).SelectedItem.Index).ListSubItems(2).Text)
            txtID.Text = ListView1(1).ListItems(ListView1(1).SelectedItem.Index).ListSubItems(3).Text
            txtFlag.Text = ListView1(1).ListItems(ListView1(1).SelectedItem.Index).ListSubItems(5).Text
            If Val(txtFlag.Text) = 0 Then   ' enable apply to Payroll button
                cmdApply.Enabled = True
            Else
                cmdApply.Enabled = False
            End If
        Else
            txtFlag.Text = ""
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Call EnableButton(True, True, False, True, False)
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Delete this Holiday Date?", vbQuestion + vbYesNo) = vbYes Then
       Call Action_Query("DELETE FROM tblHolidays WHERE HolidayID = " & txtID & "", cnnstr)
       Call EnableButton(True, True, False, True, False)
       Call Form_Load
    End If
End Sub

Private Sub cmdEdit_Click()
    Call EnableButton(False, False, True, False, True)
    TextBox (True)
    cmbHolidayType.SetFocus
End Sub

Private Sub TextBox(ByRef TFalse As Boolean)
cmbHolidayType.Enabled = TFalse
dtHoliday.Enabled = TFalse
txtHolidayName.Enabled = TFalse
End Sub

Private Sub cmdAdd_Click()
    TextBox (True)
    
    txtID = ""
    txtHolidayName.Text = ""
    dtHoliday.VALUE = Now
    
    cmbHolidayType.SetFocus
    Call EnableButton(False, False, True, False, True)
End Sub

Private Sub cmdSave_Click()
If Trim(txtHolidayName.Text) = "" Then
    MsgBox "Please Enter Holiday name.", vbInformation
    txtHolidayName.SetFocus
Else
    If IsNumeric(txtID) = True Then ' Edit
        StrQuery = "UPDATE tblHolidays SET HolidayType = '" & Left(cmbHolidayType.Text, 1) & "', HolidayDesc = '" & Mid(cmbHolidayType.Text, 3, Len(cmbHolidayType.Text) - 1) & "', HolidayDate = '" & dtHoliday.VALUE & "', HolidayName = '" & txtHolidayName.Text & "', UpdatedBy = '" & LoginName & "' WHERE HolidayID = " & txtID
        If Action_Query(StrQuery, cnnstr) = True Then
            MsgBox "Holiday entry was Updated!", vbInformation, PayrollAlert
            Call Form_Load
        Else
            MsgBox "Failed to update Holiday table!", vbInformation, PayrollAlert
        End If
    Else    ' Insert
        StrQuery = "INSERT INTO tblHolidays (HolidayType, HolidayDesc, HolidayDate, HolidayName, UpdatedBy) VALUES ('" & Left(cmbHolidayType.Text, 1) & "','" & Mid(cmbHolidayType.Text, 3, Len(cmbHolidayType.Text) - 1) & "','" & dtHoliday.VALUE & "','" & txtHolidayName & "','" & LoginName & "')"
        If Action_Query(StrQuery, cnnstr) = True Then
            MsgBox "Holiday entry was saved!", vbInformation, PayrollAlert
        Else
            MsgBox "Failed to insert new Holiday date!", vbInformation, PayrollAlert
        End If
        Call Form_Load
    End If
   TextBox (False)
    
    Call EnableButton(True, True, False, True, False)
End If
End Sub

Private Sub EnableButton(Optional ByVal Add As Boolean, Optional ByVal Edit As Boolean, Optional ByVal save As Boolean, Optional ByVal Delete As Boolean, Optional ByVal Cancel As Boolean)
    cmdadd.Enabled = Add
    cmdedit.Enabled = Edit
    cmdsave.Enabled = save
    cmdDelete.Enabled = Delete
    cmdcancel.Enabled = Cancel
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsHoliday = Nothing
End Sub

