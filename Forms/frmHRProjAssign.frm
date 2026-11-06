VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHRProjAssign 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HR Work Assignment"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   Icon            =   "frmHRProjAssign.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   7800
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   3015
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   7575
      Begin VB.TextBox txtRemarks 
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
         Left            =   1380
         TabIndex        =   6
         Top             =   2505
         Width           =   5550
      End
      Begin VB.TextBox txtAssignType 
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
         Left            =   1380
         TabIndex        =   5
         Top             =   2115
         Width           =   5550
      End
      Begin VB.TextBox txtID 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Text            =   "txtID"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdfirst 
         Height          =   255
         Left            =   6390
         Picture         =   "frmHRProjAssign.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Top record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPrev 
         Height          =   255
         Left            =   6630
         Picture         =   "frmHRProjAssign.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Previous record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdNext 
         Height          =   255
         Left            =   6870
         Picture         =   "frmHRProjAssign.frx":0B5E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Next record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdlast 
         Height          =   255
         Left            =   7110
         Picture         =   "frmHRProjAssign.frx":0CA8
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Last record"
         Top             =   0
         Width           =   255
      End
      Begin VB.ComboBox cmbempname 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "frmHRProjAssign.frx":0DF2
         Left            =   1380
         List            =   "frmHRProjAssign.frx":0DF4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1215
         Width           =   3990
      End
      Begin VB.TextBox txtempno 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1380
         TabIndex        =   1
         Top             =   825
         Width           =   2070
      End
      Begin VB.TextBox txtassignno 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1380
         TabIndex        =   0
         Top             =   420
         Width           =   2070
      End
      Begin VB.ComboBox cmbshiftsched 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "frmHRProjAssign.frx":0DF6
         Left            =   5490
         List            =   "frmHRProjAssign.frx":0DF8
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   495
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSMask.MaskEdBox mskassigndate 
         Height          =   285
         Left            =   1380
         TabIndex        =   3
         Top             =   1680
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   12582912
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskendassigndate 
         Height          =   285
         Left            =   4185
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   12582912
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Assignment"
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
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   2205
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   1335
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Assignment ID"
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
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee No"
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   930
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Assign Date"
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
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1770
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   2595
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "End Assign Date"
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
         Height          =   195
         Left            =   2895
         TabIndex        =   20
         Top             =   1755
         Width           =   1170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Schedule"
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
         Height          =   195
         Left            =   4260
         TabIndex        =   19
         Top             =   540
         Visible         =   0   'False
         Width           =   1020
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   600
         Picture         =   "frmHRProjAssign.frx":0DFA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   2340
         Picture         =   "frmHRProjAssign.frx":0F44
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1170
         Picture         =   "frmHRProjAssign.frx":1476
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmHRProjAssign.frx":19A8
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "New"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   2940
         Picture         =   "frmHRProjAssign.frx":1EDA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   1755
         Picture         =   "frmHRProjAssign.frx":2024
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   2
         X1              =   0
         X2              =   800
         Y1              =   1500
         Y2              =   1200
      End
   End
End
Attribute VB_Name = "frmHRProjAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsout As ADODB.Recordset
Dim rsHREmp As ADODB.Recordset

Dim id As String
Dim trannum As String
Dim outputfilename As String
Dim rsassigntype As ADODB.Recordset
Dim rscountycode As ADODB.Recordset
Dim rsProjectcode As ADODB.Recordset
Dim rstclient As ADODB.Recordset
Dim rsshift As ADODB.Recordset

Dim EmpID As Integer

Private Sub Form_Unload(Cancel As Integer)
Set rsHREmp = Nothing
End Sub

Private Sub txtAssignType_GotFocus()
txtAssignType.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtAssignType
End Sub

'Private Sub txtAssignType_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub

Private Sub txtAssignType_LostFocus()
txtAssignType.BackColor = vbWhite
End Sub

Private Sub cmbempname_Click()
EmpID = cmbempname.ItemData(cmbempname.ListIndex)
txtID.Text = EmpID

With rsHREmp
    .MoveFirst
    .Find "EmployeeID = " & EmpID
    If .EOF = False Then
        txtempno.Text = rsHREmp!EmpNo
    End If
    
'Call EmpField
End With

End Sub

Private Sub EmpField()
'************
'Search additional field
'************
Dim res
Dim Search, search1, Proj, str As String
Dim SelectString As String
Dim rs As ADODB.Recordset
Dim PRcn As ADODB.Connection
Set PRcn = New ADODB.Connection
Set rs = New ADODB.Recordset

If cmbempname.Text = "" Then
    If res = vbOK Then
    End If
Else
    Search = cmbempname.Text
End If

If IsEmpty(cmbempname.Text) = False Then
    SelectString = "Select * from NAME where name='" & Search & "' "
End If
Set rs = cnn.Execute(SelectString)
If rs.EOF Then
End If
ShowdataProd rs
End Sub

Private Sub ShowdataProd(ByRef rec As ADODB.Recordset)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    On Error Resume Next
'    txtavekshr.Text = IIf(IsNull(rec!avekshr), "", rec!avekshr)
'    txtavekor.Text = IIf(IsNull(rec!avekor), "", rec!avekor)
    txtempno.Text = IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
End Sub

Private Sub cmbempname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
'Call FindMatch(cmbempname, KeyAscii)
End Sub

Private Sub cmbshiftsched_GotFocus()
cmbshiftsched.BackColor = &HFFFFC0
End Sub

Private Sub cmbshiftsched_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbshiftsched_LostFocus()
cmbshiftsched.BackColor = vbWhite

End Sub

Private Sub cmdAdd_Click()
'On Error GoTo Errh

addflagship = True
editflagship = False

Call Add
Call ClearScreen

txtempno.Enabled = False
cmbempname.Enabled = True
'cmbempname.Text = ""
cmbempname.ListIndex = 0 '-1
mskassigndate.Enabled = True
txtAssignType.Enabled = True
txtRemarks.Enabled = True
mskendassigndate.Enabled = True
'cmbcountycode.Enabled = True
'cmbprojectcode.Enabled = True
'cmbclientcode.Enabled = True
cmbshiftsched.Enabled = True
cmbempname.SetFocus

Exit Sub
Errh:
    If Err.Number = 5 Then
        Resume Next
    End If
End Sub

Private Sub Add()
cmdadd.Enabled = False
cmdedit.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
cmddelete.Enabled = False
cmdrefresh.Enabled = False
cmdfirst.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdlast.Enabled = False
'cmdExit.Enabled = False
End Sub

Private Sub Edit()
cmdadd.Enabled = False
cmdedit.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
cmddelete.Enabled = False
cmdrefresh.Enabled = False
cmdfirst.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdlast.Enabled = False
'    cmdExit.Enabled = False
End Sub

Private Sub cmdAssigntype_Click()
frmAssignType.Show 1
End Sub

Private Sub cmdCancel_Click()
'If AddFlag = True Then
'        Unload Me
'        Exit Sub
'    End If

    addflagship = False
    editflagship = False

Set rsout = New ADODB.Recordset
StrQuery = "Select * from HREmployeesAssignment "
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HREmployeesAssignment ", cnnHR, adOpenKeyset, adLockOptimistic

  If rsout.EOF = False Then
        Showdata rsout
        cmddelete.Enabled = True
        cmdedit.Enabled = True
    End If
Call Cancel
End Sub

Private Sub cmdDelete_Click()
Dim res
Dim str As String
Dim assignmentnum As String

Set rsout = New ADODB.Recordset
StrQuery = "Select * FROM HREmployeesAssignment"
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HREmployeesAssignment ", cnnHR, adOpenKeyset, adLockOptimistic

res = MsgBox("Delete record", vbQuestion + vbYesNo, CompanyName)
If res = vbYes Then
    str = "Delete from HREmployeesAssignment where assignmentnum = " & txtassignno.Text
    Set rsout = cnn.Execute(str)
    'res = MsgBox("Record Deleted", vbCritical, CompanyName)
    StrQuery = "Select * from HREmployeesAssignment"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "Select * from HREmployeesAssignment", cnnHR, adOpenKeyset, adLockOptimistic
    If rsout.EOF = False Then
        Showdata rsout
          cmdNext.Enabled = True
          cmdlast.Enabled = True
          Else
          cmdNext.Enabled = False
          cmdlast.Enabled = False
          cmdfirst.Enabled = False
          cmddelete.Enabled = False
          
          Call ClearScreen
          Call Cancel
          cmdedit.Enabled = False
          cmddelete.Enabled = False
         End If
End If
End Sub

Private Sub cmdEdit_Click()
editflagship = True
addflagship = False
Call Edit

txtempno.Enabled = False
mskassigndate.Enabled = True
txtAssignType.Enabled = True
txtRemarks.Enabled = True
mskendassigndate.Enabled = True
cmbempname.Enabled = True
'cmbcountycode.Enabled = True
'cmbprojectcode.Enabled = True
'cmbclientcode.Enabled = True
cmbshiftsched.Enabled = True
cmbempname.SetFocus
End Sub

Private Sub cmdExit_Click()
'Dim res
'res = MsgBox("Exit QC TransmissionZip Form", vbQuestion + vbYesNo, CompanyName)
  '  If res = vbYes Then
        Unload Me
        'End
   ' End If
End Sub

Private Sub cmdfirst_Click()
rsout.MoveFirst
cmdfirst.Enabled = False
cmdPrev.Enabled = False
cmdlast.Enabled = True
cmdNext.Enabled = True
Showdata rsout
End Sub

Private Sub cmdlast_Click()
rsout.MoveLast
cmdlast.Enabled = False
cmdNext.Enabled = False
cmdPrev.Enabled = True
cmdfirst.Enabled = True
Showdata rsout
End Sub

Private Sub cmdNext_Click()
 On Error GoTo X
 rsout.MoveNext
 cmdPrev.Enabled = True
 cmdfirst.Enabled = True
   If rsout.EOF Then
        rsout.MoveLast
        cmdNext.Enabled = False
        cmdlast.Enabled = False
        MsgBox "End of file encountered", vbInformation, CompanyName
End If

    
    Showdata rsout
X:
    Exit Sub
 
End Sub

Private Sub cmdPrev_Click()

 On Error GoTo X
 rsout.MovePrevious
  cmdNext.Enabled = True
      cmdlast.Enabled = True
   If rsout.BOF Then
       rsout.MoveFirst
      cmdPrev.Enabled = False
      cmdfirst.Enabled = False
      MsgBox "Beginning of file encountered", vbInformation, CompanyName
    End If
    
   assignmentnum = rsout!assignmentnum
    Showdata rsout
X:
    Exit Sub
End Sub

Private Sub cmdSave_Click()

If cmbempname.Text = "" Then MsgBox "Select employee name. ", vbInformation: cmbempname.SetFocus: Exit Sub

txtempno.Enabled = False
mskassigndate.Enabled = False
txtAssignType.Enabled = False
txtRemarks.Enabled = False
mskendassigndate.Enabled = False
'cmbcountycode.Enabled = False
cmbempname.Enabled = False
'cmbprojectcode.Enabled = False
'cmbclientcode.Enabled = False
cmbshiftsched.Enabled = False

'Add new record
If AddFlag = True Or addflagship = True Then
    'Call Save
    Call SaveNewHREmployeeRecord
    MsgBox "New Assignment was successfully saved!", vbInformation, PayrollAlert
    AddFlag = False
    addflagship = False
End If
        
If EditFlag = True Or editflagship = True Then
    Call Edit
    Call EditSave
End If

Call Save

addflagship = False
editflagship = False

End Sub

Private Sub Save()
cmdadd.Enabled = True
cmdedit.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdrefresh.Enabled = False
cmddelete.Enabled = True
cmdfirst.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdlast.Enabled = True
'cmdExit.Enabled = True
End Sub

Private Sub Cancel()
cmdadd.Enabled = True
'cmdedit.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdrefresh.Enabled = False
'cmddelete.Enabled = True
cmdfirst.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdlast.Enabled = True
'cmdExit.Enabled = True
txtempno.Enabled = False
mskassigndate.Enabled = False
txtAssignType.Enabled = False
txtRemarks.Enabled = False
mskendassigndate.Enabled = False
'cmbcountycode.Enabled = False
cmbempname.Enabled = False
'cmbprojectcode.Enabled = False
'cmbclientcode.Enabled = False
cmbshiftsched.Enabled = False
End Sub

Private Sub DTmovementdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Initialize()
Set oCommonFunction = New clsCommon
End Sub

Private Sub Form_Load()

Me.Top = 0 '(mdiMain.Height - Me.Height) / 8
Me.Left = (mdiMain.Width - Me.Width) / 2

'Set rstclient = New ADODB.Recordset
'Set rsshift = New ADODB.Recordset
'Set rsemp = New ADODB.Recordset

addflagship = False
editflagship = False
txtempno.Enabled = False
mskassigndate.Enabled = False
txtAssignType.Enabled = False
txtRemarks.Enabled = False
mskendassigndate.Enabled = False
cmbempname.Enabled = False
cmbshiftsched.Enabled = False
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdrefresh.Enabled = False

'Populate ComboBox
'StrQuery = "Select * FROM HREmployees WHERE EmpStatus = 1 order by Surname, Firstname"
StrQuery = "Select EmployeeID, [Name] AS EmpName, EmpNo FROM NAME WHERE EmpStatus = 1 ORDER BY [name]"
Set rsHREmp = ReturnRst(StrQuery, cnnstr, mstatic)

Do While Not rsHREmp.EOF
    cmbempname.AddItem IIf(IsNull(rsHREmp![EmpName]), "No Name", rsHREmp![EmpName])
    cmbempname.ItemData(cmbempname.NewIndex) = rsHREmp!EmployeeID
    rsHREmp.MoveNext
Loop

'StrQuery = "Select * From HRShiftSchedule WHERE ShiftStatus = 1 Order by Shift ID"
'Set rsshift = ReturnRst(StrQuery, cnnstr, mstatic)
''Set rsshift = cnn.Execute("Select * From hrShiftSchedule")
'Do While Not rsshift.EOF
'    cmbshiftsched.AddItem rsshift!shiftcode
'    'cmbshiftsched.AddItem rsshift!shiftsched & "-" & rsshift!shiftscheddesc
'    rsshift.MoveNext
'Loop
'rsshift.Close

Call ShowAssignment

End Sub

Public Sub ShowAssignment(Optional ByVal assignmentnum As Integer)

Set rsout = New ADODB.Recordset

If rsout.State = 1 Then rsout.Close

If assignmentnum = 0 Then
    StrQuery = "Select * FROM HREmployeesAssignment "
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
Else
    StrQuery = "Select * from HREmployeesAssignment WHERE assignmentnum = " & assignmentnum & ""
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
End If

'**********
'Check if there is an existing record.
'*********

If rsout.RecordCount > 0 Then
    Showdata rsout
Else
   Call ClearScreen
   cmdfirst.Enabled = False
   cmdPrev.Enabled = False
   cmdNext.Enabled = False
   cmdlast.Enabled = False
   cmdedit.Enabled = False
   cmddelete.Enabled = False
End If

End Sub

Private Sub SaveNewHREmployeeRecord()
Dim res

' For Update 2010.
'Set rsout = New ADODB.Recordset
'rsout.Open "Select * from HREmployeesAssignment", cnnHR, adOpenKeyset, adLockOptimistic

StrQuery = "INSERT INTO HREmployeesAssignment (EmployeeID, EmployeeName, AssignDate, AssignType, AssignRemarks, EndAssignDate, UserID) " & _
    " VALUES (" & txtID.Text & ", '" & cmbempname.Text & "', '" & mskassigndate.Text & "', '" & txtAssignType.Text & "', '" & txtRemarks.Text & _
    "', '" & mskendassigndate.Text & "', '" & LoginName & "')"
Action_Query StrQuery, cnnstr

'With rsout
'    If Not (.BOF And .EOF) Then
'    .MoveLast
'    End If
'    .AddNew
'    !EmployeeID = txtempno.Text
'    !assigndate = IIf(mskassigndate.Text = "__/__/____", Null, Format(mskassigndate.Text, "mm/dd/yyyy"))
'    !AssignType = txtAssignType.Text
'    !EmployeeName = cmbempname.Text
'    !assignremarks = txtRemarks.Text
'    !endassigndate = IIf(mskendassigndate.Text = "__/__/____", Null, Format(mskendassigndate.Text, "mm/dd/yyyy"))
'    '!countycode = cmbcountycode.Text
'    '!projectcode = cmbprojectcode.Text
'    '!clientcode = Left(cmbclientcode.Text, 1)
'    !shiftsched = Left(cmbshiftsched.Text, 1)
'    !USERID = operatorID
'
'
'            .Update
'End With
''res = MsgBox("Record Added", vbInformation, CompanyName)
'Me.txtassignno.Text = Getmax

End Sub

Function Getmax() As Integer
    Dim rs As ADODB.Recordset
    Dim str As String
    str = "Select max(assignmentnum) From HREmployeesAssignment"
        Set rs = cnn.Execute(str)
    If Not rs.EOF Then
         Getmax = rs.Fields(0)
    End If
End Function

Private Sub ClearScreen()

txtassignno.Text = "<auto number>"
txtempno.Text = ""
cmbempname.ListIndex = 0 '-1
mskassigndate.Text = "__/__/____"
txtAssignType.Text = ""
txtRemarks.Text = ""
mskendassigndate.Text = "__/__/____"
'cmbcountycode.ListIndex = -1
'cmbprojectcode.Text = ""
'cmbclientcode.ListIndex = -1
cmbshiftsched.ListIndex = -1
 
' If Not Trim(projectlong) = "" Then
 '       cmbprojectcode.Text = projectlong
  '  Else
   '     cmbprojectcode.ListIndex = 0
   ' End If
    
 'If Not Trim(countylong) = "" Then
  '      cmbcountycode.Text = countylong
  '  Else
   '     cmbcountycode.ListIndex = 0
   ' End If

'If Not Trim(projectlong) = "" Then
'        cmbclientcode.Text = projectlong
'    Else
'        cmbclientcode.ListIndex = -1
'        End If
End Sub

Private Sub Showdata(ByRef rec As ADODB.Recordset)
Dim rs As ADODB.Recordset

If rec.EOF = True Then
    Call ClearScreen
    cmdedit.Enabled = False
    cmddelete.Enabled = False

Else
    txtassignno.Text = IIf(IsNull(rec!assignmentnum), "000000", rec!assignmentnum)
    txtempno.Text = IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
    mskassigndate.Text = IIf(IsNull(rec!assigndate), "__/__/____", Format(rec!assigndate, "mm/dd/yyyy"))
    txtAssignType.Text = IIf(IsNull(rec!AssignType), "", rec!AssignType)
    txtRemarks.Text = IIf(IsNull(rec!assignremarks), "", rec!assignremarks)
    mskendassigndate.Text = IIf(IsNull(rec!endassigndate), "__/__/____", Format(rec!endassigndate, "mm/dd/yyyy"))
    
With rsHREmp
    .MoveFirst
    .Find "EmployeeID = " & EmpID
    If .EOF = False Then
        txtempno.Text = rsHREmp!EmpNo
        cmbempname.Text = rsHREmp!EmpName
    End If
End With
    
    
'    strSQL = "Select * From hrShiftSchedule " & IIf(IsNull(rec!shiftsched) = True Or rec!shiftsched = "", "", "Where ShiftSched =" & rec!shiftsched) & ""
'    Set rs = ReturnRst(strSQL, cnnstr, mstatic)
'    'Set rs = cnn.Execute(strSQL)
'    If Not rs.EOF Then
'        cmbshiftsched = rs.Fields(1)
'    Else
'        cmbshiftsched.ListIndex = -1
'    End If
'    rs.Close
    
'    StrQuery = "Select * FROM HREmployees Where EmployeeID = " & IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
'    Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
'    'Set rs = cnn.Execute("Select * From HREmployees Where EmployeeID =" & IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID) & "")
'    If Not rs.EOF Then
'        cmbempname.Text = rs.Fields("EmpName") '& "," & rs.Fields(2)
'    Else
'        cmbempname.ListIndex = 0 '-1
'    End If
'    rs.Close
'
'    StrQuery = "Select EmployeeID, [Name] AS EmpName, EmpNo FROM NAME WHERE EmployeeID = " & IIf(IsNull(rec!EmployeeID), "0", rec!EmployeeID)
'    Set rsHREmp = ReturnRst(StrQuery, cnnstr, mstatic)
'
'    Do While Not rsHREmp.EOF
'        cmbempname.AddItem IIf(IsNull(rsHREmp![EmpName]), "No Name", rsHREmp![EmpName])
'        cmbempname.ItemData(cmbempname.NewIndex) = rsHREmp!EmployeeID
'        rsHREmp.MoveNext
'    Loop
    
    
End If
End Sub

Private Sub EditSave()
Dim res

StrQuery = "UPDATE HREmployeesAssignment SET EmployeeID = " & txtID.Text & ", EmployeeName = '" & cmbempname.Text & "', AssignDate = '" & mskassigndate.Text & "', AssignType = '" & _
    txtAssignType.Text & "', AssignRemarks = '" & txtRemarks.Text & "', EndAssignDate = '" & mskendassigndate.Text & "', UserID = '" & LoginName & "'"
Action_Query StrQuery, cnnstr

'With rsout
'    !EmployeeID = txtempno.Text
'    !assigndate = IIf(mskassigndate.Text = "__/__/____", Null, Format(mskassigndate.Text, "mm/dd/yyyy"))
'    !AssignType = txtAssignType.Text
'    !assignremarks = txtRemarks.Text
'    !endassigndate = IIf(mskendassigndate.Text = "__/__/____", Null, Format(mskendassigndate.Text, "mm/dd/yyyy"))
'    !EmployeeName = cmbempname.Text
''    !countycode = cmbcountycode.Text
''    !projectcode = cmbprojectcode.Text
' '   !clientcode = Left(cmbclientcode.Text, 1)
'    !shiftsched = Left(cmbshiftsched.Text, 1)
'        !USERID = operatorID
'    .Update
'End With

'res = MsgBox("Record Updated", vbInformation, CompanyName)
End Sub

Private Sub txttransactionnum_GotFocus()
txttransactionnum.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txttransactionnum
End Sub

Private Sub txttransactionnum_LostFocus()
txttransactionnum.BackColor = vbWhite
End Sub

Private Sub txttransactionnum_Validate(Cancel As Boolean)
Dim res
If Not IsNumeric(txttransactionnum.Text) Then
    res = MsgBox("Invalid Input: should be numeric", vbExclamation, CompanyName)
    Cancel = True
Else
    Cancel = False
End If
End Sub

Private Sub txtzipfilename_GotFocus()
txtzipfilename.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtzipfilename
End Sub


Private Sub txtzipfilename_LostFocus()
txtzipfilename.BackColor = vbWhite
End Sub

Private Sub txtclientcode_GotFocus()
txtClientCode.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtClientCode
End Sub

Private Sub txtclientcode_LostFocus()
txtClientCode.BackColor = vbWhite
End Sub

'Private Sub txtAssignType_GotFocus()
'txtAssignType.BackColor = &HFFFFC0
'Set oCommonFunction.ValidAlpha = txtAssignType
'End Sub
'
'Private Sub txtAssignType_LostFocus()
'txtAssignType.BackColor = vbWhite
'End Sub

Private Sub txtemployeeno_GotFocus()
txtEmployeeNo.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtEmployeeNo
End Sub

Private Sub txtemployeeno_LostFocus()
txtEmployeeNo.BackColor = vbWhite
End Sub

Private Sub txtjobclass_GotFocus()
txtjobclass.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtjobclass
End Sub

Private Sub txtjobclass_LostFocus()
txtjobclass.BackColor = vbWhite
End Sub

Private Sub txtjobcode_GotFocus()
txtjobcode.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtjobcode
End Sub

Private Sub txtjobcode_LostFocus()
txtjobcode.BackColor = vbWhite
End Sub

Private Sub txtmovementremarks_GotFocus()
txtmovementremarks.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtmovementremarks
End Sub

Private Sub txtmovementremarks_LostFocus()
txtmovementremarks.BackColor = vbWhite
End Sub

Private Sub txtmovementseq_GotFocus()
txtmovementseq.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtmovement
End Sub

Private Sub txtmovementseq_LostFocus()
txtmovementseq.BackColor = vbWhite
End Sub

Private Sub txtsalary_GotFocus()
txtsalary.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtsalary
End Sub

Private Sub txtsalary_LostFocus()
txtsalary.BackColor = vbWhite
End Sub

Private Sub txtstatus_GotFocus()
txtstatus.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtstatus
End Sub


Private Sub txtstatus_LostFocus()
txtstatus.BackColor = vbWhite
End Sub

Private Sub mskassigndate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskassigndate_LostFocus()
mskassigndate.BackColor = vbWhite
End Sub

Private Sub mskassigndate_Validate(Cancel As Boolean)
If mskassigndate <> "__/__/____" And IsDate(mskassigndate) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskassigndate.SetFocus
End Sub

Private Sub mskendassigndate_GotFocus()
mskendassigndate.BackColor = &HFFFFC0
End Sub

Private Sub mskendassigndate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskendassigndate_LostFocus()
mskendassigndate.BackColor = vbWhite

End Sub

Private Sub mskendassigndate_Validate(Cancel As Boolean)
If mskendassigndate <> "__/__/____" And IsDate(mskendassigndate) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskendassigndate.SetFocus
End Sub

Private Sub txtempno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtremarks_GotFocus()
txtRemarks.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtRemarks
End Sub

Private Sub txtremarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtremarks_LostFocus()
txtRemarks.BackColor = vbWhite
End Sub

Private Sub cmbempname_GotFocus()
cmbempname.BackColor = &HFFFFC0
'Set oCommonFunction.ValidAlpha = cmbempname
End Sub

Private Sub cmbempname_LostFocus()
cmbempname.BackColor = vbWhite
End Sub

