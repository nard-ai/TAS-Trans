VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRSearch 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HR Search"
   ClientHeight    =   4305
   ClientLeft      =   4155
   ClientTop       =   3510
   ClientWidth     =   7020
   Icon            =   "frmHRSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   7020
   Begin VB.Frame FraSettings 
      Height          =   1500
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   2355
      Begin VB.CheckBox chkStatus 
         Caption         =   "Resigned"
         Height          =   255
         Index           =   5
         Left            =   1305
         TabIndex        =   12
         Top             =   135
         Width           =   1005
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "Early EOC"
         Height          =   255
         Index           =   6
         Left            =   1305
         TabIndex        =   20
         Top             =   405
         Width           =   1020
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "EOC"
         Height          =   255
         Index           =   7
         Left            =   1305
         TabIndex        =   19
         Top             =   675
         Width           =   960
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "AWOL"
         Height          =   255
         Index           =   8
         Left            =   1305
         TabIndex        =   18
         Top             =   945
         Width           =   960
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "Contractual"
         Height          =   255
         Index           =   1
         Left            =   45
         TabIndex        =   17
         Top             =   405
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "Project Base"
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   16
         Top             =   135
         Width           =   1590
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "Probationary"
         Height          =   255
         Index           =   2
         Left            =   45
         TabIndex        =   15
         Top             =   675
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "Regular"
         Height          =   255
         Index           =   3
         Left            =   45
         TabIndex        =   14
         Top             =   945
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "Terminated"
         Height          =   255
         Index           =   4
         Left            =   45
         TabIndex        =   13
         Top             =   1215
         Width           =   1590
      End
   End
   Begin VB.ComboBox cmbClient 
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
      ItemData        =   "frmHRSearch.frx":058A
      Left            =   60
      List            =   "frmHRSearch.frx":058C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3870
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   6990
      Begin MSComctlLib.ListView lvwdisplay 
         Height          =   2970
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   6960
         _ExtentX        =   12277
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
         Left            =   5400
         TabIndex        =   2
         Top             =   3285
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   847
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdSettings 
         BackColor       =   &H00FFFFC0&
         Caption         =   "[ ]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   330
      End
      Begin VB.Frame fraSearch 
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   2880
         TabIndex        =   6
         Top             =   -90
         Width           =   3930
         Begin VB.CommandButton cmdArrange 
            BackColor       =   &H00FFFFC0&
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3555
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   225
            Width           =   330
         End
         Begin VB.TextBox txtSearch 
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
            Height          =   285
            Left            =   720
            TabIndex        =   7
            Top             =   180
            Width           =   2760
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "NAME:"
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
            Left            =   180
            TabIndex        =   8
            Top             =   225
            Width           =   1560
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   510
         Left            =   3105
         TabIndex        =   4
         Top             =   0
         Width           =   3750
      End
   End
End
Attribute VB_Name = "frmHRSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim whereCon As String
Dim Filter_var As String
Dim LoadFlag As Boolean

Public Sub chkStatus_Click(Index As Integer)
Dim str As String
str = ""

If chkStatus(0).VALUE = 1 Then
str = "'B'"
End If

If chkStatus(1).VALUE = 1 Then
If str <> "" Then str = str & ",'C'" Else str = "'C'"
End If

If chkStatus(2).VALUE = 1 Then
If str <> "" Then str = str & ",'P'" Else str = "'P'"
End If

If chkStatus(3).VALUE = 1 Then
If str <> "" Then str = str & ",'R'" Else str = "'R'"
End If

If chkStatus(4).VALUE = 1 Then
If str <> "" Then str = str & ",'T'" Else str = "'T'"
End If

If chkStatus(5).VALUE = 1 Then
If str <> "" Then str = str & ",'X'" Else str = "'X'"
End If

If chkStatus(6).VALUE = 1 Then
If str <> "" Then str = str & ",'Y'" Else str = "'Y'"
End If

If chkStatus(7).VALUE = 1 Then
    If str <> "" Then str = str & ",'Z'" Else str = "'Z'"
End If

If chkStatus(8).VALUE = 1 Then
    If str <> "" Then str = str & ",'W'" Else str = "'W'"
End If

If str = "" Then MsgBox "Please check at least 1 check box.", vbInformation:  chkStatus(1).VALUE = 1: Exit Sub

whereCon = ""
whereCon = " EMPLOYMENTSTATUS IN " & "(" & str & ")"

If Me.cmbClient.Text <> "--ALL COMPANIES--" Then
    whereCon = whereCon & " and Clientcode= " & Me.cmbClient.ItemData(Me.cmbClient.ListIndex)
End If
      
whereCon = " WHERE " & whereCon & " and Agencycode = " & CompanyCode
 
'Set rsEmp1 = New ADODB.Recordset
'StrQuery = "SELECT Surname + ', ' + FirstName + ' ' + CASE WHEN Middlename IS NULL THEN '' ELSE LEFT(Middlename, 1) + '.' END AS [Employee Name], " & _
'                " EmployeeID AS [Employee ID], EmpNo AS [Emp No], StartContract AS [Start Contract], EndContract AS [End Contract], * " & _
'                " FROM HREmployees " & whereCon & " ORDER BY Surname, Firstname "
'Set rsEmp1 = ReturnRst(StrQuery, cnnstr, mstatic)

' Comment by Bernard 11/08/2012
'If LoadFlag = True Then
    Loadlist (whereCon)
'End If
'LoadFlag = True

End Sub

Public Sub cmbClient_Click()
AutoSizeDropDownWidth cmbClient
If Me.cmbClient.Text = "--ALL COMPANIES--" Then
    Filter_var = ""
    CurrentClientCode = 0
Else
    Filter_var = " WHERE clientcode=" & Me.cmbClient.ItemData(Me.cmbClient.ListIndex)
    CurrentClientCode = Me.cmbClient.ItemData(Me.cmbClient.ListIndex)
End If
Loadlist (Filter_var)  ' comment by bernard 11/08/2012
Call Me.chkStatus_Click(1)
End Sub

Public Sub cmdArrange_Click()

If Me.Height <> mdiMain.ScaleHeight Then
    cmdArrange.Caption = "><"
    Me.Height = mdiMain.ScaleHeight
    Me.Width = 2400
    Me.Left = mdiMain.ScaleWidth - Me.Width
    Me.Top = 0
    Frame1.Height = Me.Height - 500
    lvwdisplay.Height = Me.Height - 1600
    cmbClient.Top = Me.Height - 830
    cmbClient.Width = 2200
    lvwdisplay.Width = Me.Width - 100
    lvwdisplay.ColumnHeaders(1).Width = 2500  ' 2000
    lvwdisplay.ColumnHeaders(2).Width = 0
    lvwdisplay.ColumnHeaders(3).Width = 0
    lvwdisplay.ColumnHeaders(4).Width = 0
    lvwdisplay.ColumnHeaders(5).Width = 0
    fraSearch.Left = 0
    txtSearch.Width = 2200
    Label2.Top = 120
    Label2.Left = 40
    txtSearch.Left = 60
    txtSearch.Top = 350
    cmdArrange.Top = 120
    cmdArrange.Left = 1890
    cmdSettings.Top = 30
    cmdSettings.Left = 1485
    Call cmdSettings_Click
    FraSettings.Width = 2355
Else
    FraSettings.Width = 6600
    cmdArrange.Caption = ">>"
    cmdArrange.Top = 225
    cmdArrange.Left = 3555
    Me.Height = 4680 + 2800  ' + 2800
    Me.Width = 7110
    Me.Top = (mdiMain.Height - Me.Height) / 4
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.Top = 0
    Frame1.Height = 4000 + 2800
    lvwdisplay.Height = 2970
    
    cmbClient.Top = 3870
    cmbClient.Width = 2775
    cmbClient.Top = Me.Height - 830
    lvwdisplay.Width = 6950
    lvwdisplay.ColumnHeaders(1).Width = 2500  '2000
    lvwdisplay.ColumnHeaders(2).Width = 0
    lvwdisplay.ColumnHeaders(3).Width = 1500
    lvwdisplay.ColumnHeaders(4).Width = 1500
    lvwdisplay.ColumnHeaders(5).Width = 1500
    fraSearch.Left = 2880
    txtSearch.Width = 2760
    Label2.Top = 225
    Label2.Left = 180
    txtSearch.Left = 720
    txtSearch.Top = 180
    cmdSettings.Top = 90
    cmdSettings.Left = 90
    Call cmdSettings_Click
End If

End Sub

Private Sub cmdSettings_Click()
On Error Resume Next

If FraSettings.Visible = True Then
    FraSettings.Visible = False
    lvwdisplay.Height = Me.Height - 1750
    lvwdisplay.Top = 225
    FraSettings.Width = 2355
Else
    FraSettings.Visible = True
    lvwdisplay.Height = Me.Height - 3100
    lvwdisplay.Top = 1600
    FraSettings.Width = 6950
End If
End Sub

Private Sub Form_Activate()
Call Me.chkStatus_Click(1)
With frmHRFinalEmployees
    Set .rsHREmp = New ADODB.Recordset
    StrQuery = "SELECT * FROM HREmployees " & whereCon & "  order by  surname, firstname"
    Set .rsHREmp = ReturnRst(StrQuery, cnnstr, mstatic)
End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Dim rstclient As Recordset

mHRSEARCH = True
LoadFlag = True
Me.Top = (mdiMain.Height - Me.Height) / 4
Me.Left = (mdiMain.Width - Me.Width) / 2

StrQuery = GetClientRecord(0, True, " ORDER BY ClientName")
Set rstclient = ReturnRst(StrQuery, cnnstr, mstatic)
    
Do While Not rstclient.EOF
    cmbClient.AddItem rstclient!ClientName
    cmbClient.ItemData(cmbClient.NewIndex) = rstclient!ClientCode
    rstclient.MoveNext
Loop
rstclient.Close

If AccountType <> "FMS" Then ' Updated by Bernard 6/30/2011.
    cmbClient.AddItem "--ALL COMPANIES--", 0 ' updated by bernard 06/30/2011
End If

cmbClient.ListIndex = 0
CurrentClientCode = 0

'LoadList (" WHERE AgencyCode = " & CompanyCode) ' orig

Call Me.chkStatus_Click(1)
With frmHRFinalEmployees
    Set .rsHREmp = New ADODB.Recordset
    StrQuery = "SELECT * FROM HREmployees " & whereCon & "  order by  surname, firstname"
    Set .rsHREmp = ReturnRst(StrQuery, cnnstr, mstatic)
End With
Call cmdArrange_Click
    
End Sub

Public Sub Loadlist(whereCon As String)
Dim rsEmp1 As ADODB.Recordset

StrQuery = "SELECT Surname + ', ' + FirstName + ' ' + CASE WHEN Middlename IS NULL THEN '' ELSE LEFT(Middlename, 1) + '.' END AS [Employee Name], " & _
                " EmployeeID AS [Employee ID], EmpNo AS [Emp No], StartContract AS [Start Contract], EndContract AS [End Contract] " & _
                " FROM HREmployees " & whereCon & " ORDER BY Surname, Firstname "
Set rsEmp1 = ReturnRst(StrQuery, cnnstr, mstatic)

'Refresh HR Form recordset.
'If lvwdisplay.ListItems.Count <> 0 And CurrentClientCode <> 0 Then
'If rsEmp1.RecordCount > 0 And CurrentClientCode <> 0 Then
    With frmHRFinalEmployees
        Set .rsHREmp = New ADODB.Recordset
        StrQuery = "SELECT * FROM HREmployees " & whereCon & "  order by  surname, firstname"
        Set .rsHREmp = ReturnRst(StrQuery, cnnstr, mstatic)
    End With
'End If

Me.lvwdisplay.ListItems.Clear
PopulateList Me.lvwdisplay, rsEmp1

lvwdisplay.ColumnHeaders(1).Text = "  " & lvwdisplay.ListItems.Count & " RECORD/S "
If cmdArrange.Caption <> "><" Then
   lvwdisplay.ColumnHeaders(1).Width = 2500
   lvwdisplay.ColumnHeaders(2).Width = 0
   lvwdisplay.ColumnHeaders(3).Width = 1500
   lvwdisplay.ColumnHeaders(4).Width = 1500
   lvwdisplay.ColumnHeaders(5).Width = 1500
Else
   lvwdisplay.ColumnHeaders(1).Width = 2500
   lvwdisplay.ColumnHeaders(2).Width = 0
   lvwdisplay.ColumnHeaders(3).Width = 0
   lvwdisplay.ColumnHeaders(4).Width = 0
   lvwdisplay.ColumnHeaders(5).Width = 0
End If
lblRecount = lvwdisplay.ListItems.Count & " record/s "
   
Set rsEmp1 = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
mHRSEARCH = False
End Sub

Private Sub lvwdisplay_DblClick()
Dim strRow As String

On Error GoTo Exit32

strRow = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).SubItems(1)
txtSearch = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).Text
If lvwdisplay.ListItems.Count = 0 Then
    MsgBox "No HR records selected!", vbOKOnly + vbInformation, "Select..."
    Exit Sub
End If
Call Search(strRow)
Exit32:

End Sub

Private Sub lvwdisplay_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtSearch = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).Text
End Sub

Private Sub lvwdisplay_KeyDown(KeyCode As Integer, Shift As Integer)
If FeMP = True Then
    Call lvwdisplay_DblClick
End If
End Sub

Private Sub lvwdisplay_KeyUp(KeyCode As Integer, Shift As Integer)
If FeMP = True Then
    Call lvwdisplay_DblClick
End If
End Sub

Private Sub txtSearch_Change()
Autocomplete lvwdisplay, (Trim$(txtSearch)), txtSearch
End Sub

Private Sub txtSearch_GotFocus()
Call txtHLigths(txtSearch)
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Execute = True
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
    Execute = False
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
Dim strRow As String

strRow = ""
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
If KeyAscii = 13 Then
    If lvwdisplay.ListItems.Count > 0 Then
        strRow = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).SubItems(1)
        Call Search(strRow)
    Else
        MsgBox "No records Found.", vbInformation, PayrollAlert
    End If
End If
    
End Sub

Private Sub Search(ByVal strSearch As Integer)
Dim res, bkmrk, mmEmpTemp

On Error GoTo ErrRtn1

Repeat1:

mmEmpTemp = MMEmp
If MMEmp = False And FeMP = False Then FeMP = True

If FeMP = True Then
    With frmHRFinalEmployees
        If .rsHREmp.EOF = False Then
            bkmrk = .rsHREmp.Bookmark
        Else
            bkmrk = 1
        End If
        
        .rsHREmp.MoveFirst
        .rsHREmp.Find "EmployeeID = '" & Val(strSearch) & "'"
        
        ' Refresh the display in HR Form.
        If .rsHREmp.EOF Then
            .rsHREmp.Bookmark = bkmrk
        End If
        .Showdata .rsHREmp
        If MMEmp = True Then
            Showdata .rsHREmp
        End If
    End With
        
        
        
'    ' Original code ' Bernard 11/08/2012
'    Set .rsHREmp = New ADODB.Recordset
'    LoadFlag = False
'    Me.chkStatus_Click (1)
'    StrQuery = "SELECT * FROM HREmployees " & whereCon & " order by  surname, firstname"
'    Set .rsHREmp = ReturnRst(StrQuery, cnnstr, mstatic)
'
    
'    If rsEmp1.RecordCount > 0 Then
'        rsEmp1.MoveFirst
'        rsEmp1.Find "EmployeeID = '" & Val(strSearch) & "'"
'        If rsEmp1.EOF = True Then
'            .rsHREmp.Bookmark = bkmrk
'        End If
'        Showdata rsEmp1
'        .Showdata rsEmp1
'    End If
    
    
'    If rsRec.RecordCount > 0 Then
'        rsRec.MoveFirst
'        rsRec.Find "EmployeeID = '" & Val(strSearch) & "'"
'        If rsRec.EOF = True Then
'            frmHRFinalEmployees.rsHREmp.Bookmark = bkmrk
'        End If
'        Showdata rsRec
'        frmHRFinalEmployees.Showdata rsRec
'    End If
    
    If MMEmp = False Then MMEmp = mmEmpTemp
End If

'' Comment by Bernard 11/08/2012

'If MMEmp = True Then
'
'Set frmHRMainEmployee.rs = New ADODB.Recordset
'
'If frmHRMainEmployee.rs.State = 1 Then rs.Close
'
'    With frmHRMainEmployee
''    StrQuery = "SELECT * FROM HREmployees WHERE agencycode = " & CompanyCode & " order by surname, firstname"
''    Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
''    '.rs.Open "Select * from HREmployees where agencycode=" & CompanyCode & " order by surname, firstname", cnnHR, adOpenKeyset, adLockOptimistic
''
''    .rs.Find "EmployeeID=" & Trim(strSearch) & ""
''    If Not .rs.EOF Then
''        .Showdata .rs
''    Else
''        .rs.Bookmark = bkmrk
''    End If
'
'    If .rs.EOF = False Then
'        bkmrk = .rs.Bookmark
'    Else
'        bkmrk = 1
'    End If
'    Set .rs = New ADODB.Recordset
'    StrQuery = "SELECT * FROM HREmployees " & whereCon & " order by  surname, firstname"
'    Set .rs = ReturnRst(StrQuery, cnnstr, mstatic)
'
'    .rs.MoveFirst
'    .rs.Find "EmployeeID = '" & Val(strSearch) & "'"
'    If Not .rs.EOF Then
'        .Showdata .rs
'    Else
'        .rs.Bookmark = bkmrk
'        .Showdata .rs
'    End If
'    End With
'
'End If

Exit Sub

ErrRtn1:

    'MsgBox Err.Description & " " & Err.Number, vbInformation, PayrollAlert

If Err.Number = 3704 Then
    CurrentClientCode = 1
    Err.Clear
    On Error GoTo ErrRtn1
    Me.chkStatus_Click (1)
    GoTo Repeat1
End If

End Sub

Private Sub Showdata(ByRef rec As ADODB.Recordset)
On Error Resume Next

frmHRMainEmployee.TxtEmployeeNum.Text = IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
frmHRMainEmployee.txtEmpName.Text = rec!Surname & ", " & rec!FirstName
'frmHRMainEmployee.txtClient.Text = IIf(IsNull(rec!ClientCode), "", rec!ClientCode)
frmHRMainEmployee.txtClient.Text = frmHRFinalEmployees.lblCompany

frmHRMainEmployee.mskDtHired.Text = IIf(IsNull(rec!datehired), "__/__/____", rec!datehired)

End Sub

'Private Sub Info(EmpNo As String)
'Dim str, str1, str2, str3, str4, str5, str6, str7
'
'
' 'Assignment
'    str1 = "Select * From HREmpAssignment where EmployeeID='" & EmpNo & "'"
'        Set rsA = cnnHR.Execute(str1)
'        Do While Not rsA.EOF
'            Set ListAssgn = frmHRMainEmployee.lvwAssignment.ListItems.Add(, , rsA!EmployeeID)
'            ListAssgn.SubItems(1) = IIf(IsNull(rsA!assignmentnum), "", rsA!assignmentnum)
'            ListAssgn.SubItems(2) = IIf(IsNull(rsA!assigndate), "", rsA!assigndate)
'            ListAssgn.SubItems(3) = IIf(IsNull(rsA!AssignType), "", rsA!AssignType)
'            ListAssgn.SubItems(4) = IIf(IsNull(rsA!endassigndate), "", rsA!endassigndate)
'            'ListAssgn.SubItems(5) = IIf(IsNull(rsA!CountyName), "", rsA!CountyName)
'            'ListAssgn.SubItems(6) = IIf(IsNull(rsA!projectname), "", rsA!projectname)
''           ListAssgn.SubItems(7) = IIf(IsNull(rsA!shiftsched), "", rsA!shiftsched)
'        rsA.MoveNext
'        Loop
'        rsA.Close
'
'    'Training
'
'    str2 = "Select * From HREmployeesTraining where EmployeeID='" & EmpNo & "'"
'        Set rst = cnnHR.Execute(str2)
'        Do While Not rst.EOF
'            Set ListTrng = frmHRMainEmployee.lvwTraining.ListItems.Add(, , rst!EmployeeID)
'            ListTrng.SubItems(1) = IIf(IsNull(rst!trainingnum), "", rst!trainingnum)
'            ListTrng.SubItems(2) = IIf(IsNull(rst!trainingtype), "", rst!trainingtype)
'            ListTrng.SubItems(3) = IIf(IsNull(rst!training), "", rst!training)
'            ListTrng.SubItems(4) = IIf(IsNull(rst!trainingdate), "", rst!trainingdate)
'            ListTrng.SubItems(5) = IIf(IsNull(rst!trainingduration), "", rst!trainingduration)
'            ListTrng.SubItems(6) = IIf(IsNull(rst!institute), "", rst!institute)
'            ListTrng.SubItems(7) = IIf(IsNull(rst!facilitator), "", rst!facilitator)
'            ListTrng.SubItems(8) = IIf(IsNull(rst!trainingvenue), "", rst!trainingvenue)
'            ListTrng.SubItems(9) = IIf(IsNull(rst!trainingcost), "", rst!trainingcost)
'            ListTrng.SubItems(10) = IIf(IsNull(rst!starttrainingcontract), "", rst!starttrainingcontract)
''           ListTrng.SubItems(11) = IIf(IsNull(rsT!endtrainingcontract), "", rsT!endtrainingcontract)
'        rst.MoveNext
'        Loop
'        rst.Close
'
'
'
'    'Offenses
'    str3 = "Select * From HREmployeesOffenses where EmployeeID='" & EmpNo & "'"
'        Set rsO = cnnHR.Execute(str3)
'        Do While Not rsO.EOF
'            Set ListOff = frmHRMainEmployee.lvwOffenses.ListItems.Add(, , rsO!EmployeeID)
'            ListOff.SubItems(1) = IIf(IsNull(rsO!offensenum), "", rsO!offensenum)
'            ListOff.SubItems(2) = IIf(IsNull(rsO!offensemade), "", rsO!offensemade)
'            ListOff.SubItems(3) = IIf(IsNull(rsO!offensetype), "", rsO!offensetype)
'            ListOff.SubItems(4) = IIf(IsNull(rsO!offensedate), "", rsO!offensedate)
'            ListOff.SubItems(5) = IIf(IsNull(rsO!disciplinaryaction), "", rsO!disciplinaryaction)
'            ListOff.SubItems(6) = IIf(IsNull(rsO!actiontype), "", rsO!actiontype)
'            ListOff.SubItems(7) = IIf(IsNull(rsO!dateserved), "", rsO!dateserved)
''           ListOff.SubItems(8) = IIf(IsNull(rsO!counsellingdate), "", rsO!counsellingdate)
''           ListOff.SubItems(9) = IIf(IsNull(rsO!counsellingfacilitator), "", rsO!counsellingfacilitator)
'        rsO.MoveNext
'        Loop
'        rsO.Close
'
'    'Leaves
'    str4 = "Select * From HREmployeesLeaves where EmployeeID='" & EmpNo & "'"
'        Set rsL = cnnHR.Execute(str4)
'        Do While Not rsL.EOF
'            Set ListLeave = frmHRMainEmployee.lvwLeaves.ListItems.Add(, , rsL!leavenum)
'            ListLeave.SubItems(1) = IIf(IsNull(rsL!leavetype), "", rsL!leavetype)
'            ListLeave.SubItems(2) = IIf(IsNull(rsL!leavedatefr), "", rsL!leavedatefr)
'            ListLeave.SubItems(3) = IIf(IsNull(rsL!leavedateto), "", rsL!leavedateto)
'            ListLeave.SubItems(4) = IIf(IsNull(rsL!leavebal), "", rsL!leavebal)
'            ListLeave.SubItems(5) = IIf(IsNull(rsL!LeaveApprovedBy), "", rsL!LeaveApprovedBy)
'            ListLeave.SubItems(6) = IIf(IsNull(rsL!leaveserved), "", rsL!leaveserved)
'        rsL.MoveNext
'        Loop
'        rsL.Close
'
'
'    'Benefits
'       str6 = "Select * From HREmployeeBenefits where EmployeeID = " & EmpNo & "'"
'       'str6 = "Select * From HREmployeeBenefits where EmployeeID = " & EmpNo & "'"
'        Set rsB = cnnHR.Execute(str6)
'        Do While Not rsB.EOF
'            Set ListBen = frmHRMainEmployee.lvwBenefits.ListItems.Add(, , rsB!EmployeeID)
'            ListBen.SubItems(1) = IIf(IsNull(rsB!id), "", rsB!id)
'            ListBen.SubItems(2) = IIf(IsNull(rsB!benefitnum), "", rsB!benefitnum)
'      '      ListBen.SubItems(3) = IIf(IsNull(rsB!benefitacquiredate), "", rsB!benefitacquiredate)
'        rsB.MoveNext
'        Loop
'        rsB.Close
'
'    'Appraisal
'        str7 = "Select * From HREmployeesAppraisal where EmployeeID='" & EmpNo & "'"
'        Set rsP = cnnHR.Execute(str7)
'        Do While Not rsP.EOF
'            Set ListAppr = frmHRMainEmployee.lvwAppraisal.ListItems.Add(, , rsP!EmployeeID)
'            ListAppr.SubItems(1) = IIf(IsNull(rsP!appraisalnum), "", rsP!appraisalnum)
''            ListAppr.SubItems(2) = IIf(IsNull(rsP!appraisaldate), "", rsP!appraisaldate)
'        rsP.MoveNext
'        Loop
'        rsP.Close
'
'End Sub

Public Sub CommandPass(ByVal strCommand As String)

Select Case strCommand
    Case "Close"
        Unload Me
End Select

End Sub

