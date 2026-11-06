VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHRMovement 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HR Movement Data Entry"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   Icon            =   "frmHRMovement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   7755
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   90
      TabIndex        =   6
      Top             =   675
      Width           =   7575
      Begin VB.CommandButton cmdfirst 
         Height          =   255
         Left            =   6480
         Picture         =   "frmHRMovement.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Top record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPrev 
         Height          =   255
         Left            =   6720
         Picture         =   "frmHRMovement.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Previous record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdNext 
         Height          =   255
         Left            =   6960
         Picture         =   "frmHRMovement.frx":081E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Next record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdlast 
         Height          =   255
         Left            =   7200
         Picture         =   "frmHRMovement.frx":0968
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   1620
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1215
         Width           =   2055
      End
      Begin VB.TextBox txtempno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1620
         TabIndex        =   1
         Top             =   810
         Width           =   2010
      End
      Begin VB.TextBox txtmovementno 
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
         Height          =   285
         Left            =   1590
         TabIndex        =   0
         Top             =   405
         Width           =   2010
      End
      Begin VB.TextBox txtmovementremarks 
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
         Height          =   735
         Left            =   3825
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1170
         Width           =   3495
      End
      Begin VB.ComboBox cmbmovementtype 
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
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   405
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskmovementdate 
         Height          =   285
         Left            =   1590
         TabIndex        =   3
         Top             =   1635
         Width           =   2010
         _ExtentX        =   3545
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
         Format          =   "mm/dd/yy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
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
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement No."
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
         TabIndex        =   11
         Top             =   495
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee No."
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
         TabIndex        =   10
         Top             =   885
         Width           =   990
      End
      Begin VB.Label label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Date"
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
         TabIndex        =   9
         Top             =   1755
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Type"
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
         Left            =   3840
         TabIndex        =   8
         Top             =   495
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movement Remarks"
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
         Left            =   3840
         TabIndex        =   7
         Top             =   885
         Width           =   1410
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   3000
         Picture         =   "frmHRMovement.frx":0AB2
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   2400
         Picture         =   "frmHRMovement.frx":0BB4
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmHRMovement.frx":0CFE
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "New"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1200
         Picture         =   "frmHRMovement.frx":1230
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1800
         Picture         =   "frmHRMovement.frx":1762
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   600
         Picture         =   "frmHRMovement.frx":1C94
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   450
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
Attribute VB_Name = "frmHRMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Set oCommonFunction.ValidAlpha = txt

Dim rsout As ADODB.Recordset
Dim id As String
Dim trannum As String
Dim outputfilename As String
Dim rsMoveType As ADODB.Recordset
Dim rsEmpType As ADODB.Recordset
Dim rsjobclass As ADODB.Recordset
Dim rsjobcode As ADODB.Recordset
Dim rstClient As ADODB.Recordset
Dim rs As ADODB.Recordset

Private Sub cmbemploymenttype_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub cmbclientcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbempname_Click()
Call EmpField
End Sub
Private Sub EmpField()
'************
'search additional field
'************
Dim res
Dim search, search1, Proj, str As String
Dim SelectString As String
Dim rs As ADODB.Recordset
Dim PRcn As ADODB.Connection
Set PRcn = New ADODB.Connection
Set rs = New ADODB.Recordset
'prcn.Open "dcpi"
'On Error Resume Next
If cmbempname.Text = "" Then
     If res = vbOK Then
        'cmbProjectName.SetFocus
       End If
     Else
        'search = txtcustomercode.Text
        'search1 = txtprojectcode.Text
        'str = search & search1
        search = cmbempname.Text
End If

'If IsEmpty(txtcustomercode.Text) = False And IsEmpty(txtprojectcode.Text) = False Then
If IsEmpty(cmbempname.Text) = False Then
    'SelectString = "Select * from dcprojectcounty where projectshort='" & str & "' "
    SelectString = "select * from name where name='" & search & "' "
End If
Set rs = cnDCPI.Execute(SelectString)
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
    txtEmpNo.Text = IIf(IsNull(rec!employeenum), "", rec!employeenum)
End Sub


Private Sub cmbmovementtype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmdAdd_Click()
'On Error GoTo Errh
addflagship = True
editflagship = False

Call Add
Call clearscreen

'txtmovementno.Enabled = True
txtEmpNo.Enabled = False
mskmovementdate.Enabled = True
cmbmovementtype.Enabled = True
txtmovementremarks.Enabled = True
cmbempname.Enabled = True
'cmbclientcode.Enabled = True
'txtsalary.Enabled = True
'txtmovementseq.Enabled = True
'cmbemploymenttype.Enabled = True
'cmbjobcode.Enabled = True
'cmbjobclass.Enabled = True
'txtstatus.Enabled = True
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
    'cmdExit.Enabled = False
End Sub
Private Sub cmdCancel_Click()
If AddFlag = True Then
        Unload Me
        Exit Sub
    End If

    addflagship = False
    editflagship = False
    Showdata rsout
Call Cancel
End Sub
Private Sub cmdDelete_Click()
Dim res
Dim str As String
Dim movementnum As String

Set rsout = New ADODB.Recordset
rsout.Open "Select * from HREmployeesMovement ", cnDCPI, adOpenKeyset, adLockOptimistic

res = MsgBox("Delete record", vbQuestion + vbYesNo, CompanyName)
If res = vbYes Then
    str = "Delete from HREmployeesMovement where movementnum=" & txtmovementno.Text
    Set rsout = cnDCPI.Execute(str)
    'res = MsgBox("Record Deleted", vbCritical, CompanyName)
    rsout.Open "Select * from HREmployeesMovement", cnDCPI, adOpenKeyset, adLockOptimistic
         Showdata rsout
          cmdNext.Enabled = True
          cmdlast.Enabled = True
End If
End Sub
Private Sub cmdEdit_Click()
    editflagship = True
    addflagship = False
    Call Edit
'txtmovementno.Enabled = True
txtEmpNo.Enabled = False
cmbempname.Enabled = True
mskmovementdate.Enabled = True
cmbmovementtype.Enabled = True
txtmovementremarks.Enabled = True
'cmbclientcode.Enabled = True
'txtsalary.Enabled = True
'txtmovementseq.Enabled = True
'cmbemploymenttype.Enabled = True
'cmbjobcode.Enabled = True
'cmbjobclass.Enabled = True
'txtstatus.Enabled = True
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
Dim notnum As String
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
      MsgBox "Beginning of file encountered", vbInformation, "DCPI Infromation System"
    End If
    
   movementnum = rsout!movementnum
    Showdata rsout
X:
    Exit Sub
End Sub

Private Sub cmdSave_Click()
If cmbempname.Text = "" Then MsgBox "Select employee name. ", vbInformation: cmbempname.SetFocus: Exit Sub

txtmovementno.Enabled = False
txtEmpNo.Enabled = False
mskmovementdate.Enabled = False
cmbmovementtype.Enabled = False
txtmovementremarks.Enabled = False
cmbempname.Enabled = False
'cmbclientcode.Enabled = False
'txtsalary.Enabled = False
'txtmovementseq.Enabled = False
'cmbemploymenttype.Enabled = False
'cmbjobcode.Enabled = False
'cmbjobclass.Enabled = False
'txtstatus.Enabled = False
'Add new record

        If AddFlag = True Or addflagship = True Then

           Call save
           Call saveclick
           AddFlag = False
           addflagship = False
         End If
        
        If EditFlag = True Or editflagship = True Then
            Call Edit
            Call editsave
        End If
        
        Call save

        
        addflagship = False
        editflagship = False

End Sub
Private Sub save()
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
txtmovementno.Enabled = False
txtEmpNo.Enabled = False
mskmovementdate.Enabled = False
cmbmovementtype.Enabled = False
txtmovementremarks.Enabled = False
cmbempname.Enabled = False
'cmbclientcode.Enabled = False
'txtsalary.Enabled = False
'txtmovementseq.Enabled = False
'cmbemploymenttype.Enabled = False
'cmbjobcode.Enabled = False
'cmbjobclass.Enabled = False
'txtstatus.Enabled = False
End Sub

Private Sub Form_Initialize()
Set oCommonFunction = New clsCommon
End Sub

Private Sub mskmovementdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
   Me.Top = 0 '(mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
Set rsMoveType = New ADODB.Recordset
Set rsEmpType = New ADODB.Recordset
Set rsjobclass = New ADODB.Recordset
Set rsjobcode = New ADODB.Recordset
Set rstClient = New ADODB.Recordset
Set rsEmp = New ADODB.Recordset

addflagship = False
editflagship = False
txtmovementno.Enabled = False
txtEmpNo.Enabled = False
mskmovementdate.Enabled = False
cmbmovementtype.Enabled = False
txtmovementremarks.Enabled = False
cmbempname.Enabled = False
'cmbclientcode.Enabled = False
'txtsalary.Enabled = False
'txtmovementseq.Enabled = False
'cmbemploymenttype.Enabled = False
'cmbjobcode.Enabled = False
'cmbjobclass.Enabled = False
'txtstatus.Enabled = False
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdrefresh.Enabled = False
Set rsout = New ADODB.Recordset

'populate of combo boxes
Set rsEmp = cnDCPI.Execute("Select * From HREmployees")
Do While Not rsEmp.EOF
        cmbempname.AddItem rsEmp!Surname & "," & rsEmp!FirstName
        rsEmp.MoveNext
    Loop
    rsEmp.Close

Set rsMoveType = cnDCPI.Execute("Select * From HRMovementType")
Do While Not rsMoveType.EOF
        cmbmovementtype.AddItem rsMoveType!movementtype & "-" & rsMoveType!movementtypedesc
        rsMoveType.MoveNext
    Loop
    rsMoveType.Close

'Set rstClient = cnDCPI.Execute("select * from clienttable")
' Do While Not rstClient.EOF
'        cmbclientcode.AddItem rstClient!ClientName
'        cmbclientcode.ItemData(cmbclientcode.NewIndex) = rstClient!ClientCode
'        rstClient.MoveNext
'    Loop
'    rstClient.Close
'
'Set rsEmpType = cnDCPI.Execute("Select * from HREmploymentType")
'Do While Not rsEmpType.EOF
'    cmbemploymenttype.AddItem rsEmpType!employmenttype & "-" & rsEmpType!employmenttyhpesdesc
'    rsEmpType.MoveNext
'    Loop
'    rsEmpType.Close
'
'Set rsjobclass = cnDCPI.Execute("Select * from HRJobClass")
'Do While Not rsjobclass.EOF
'    cmbjobclass.AddItem rsjobclass!jobclass & "-" & rsjobclass!jobclassdesc
'    rsjobclass.MoveNext
'    Loop
'    rsjobclass.Close
'
'Set rsjobcode = cnDCPI.Execute("Select * from HRJobCode")
'Do While Not rsjobcode.EOF
'    cmbjobcode.AddItem rsjobcode!jobcode & "-" & rsjobcode!jobcodedesc
'    rsjobcode.MoveNext
'    Loop
'    rsjobcode.Close
showMovement




End Sub

Public Sub showMovement(Optional ByVal movementnum As Integer)
'open database
If rsout.State = 1 Then rsout.Close
If movementnum = 0 Then
rsout.Open "select * from HREmployeesMovement", cnDCPI, adOpenKeyset, adLockOptimistic
Else
rsout.Open "select * from HREmployeesMovement where movementnum=" & movementnum & "", cnDCPI, adOpenKeyset, adLockOptimistic
End If

'**************
'check if there is an existing record
'**************
If rsout.RecordCount > 0 Then
    Showdata rsout
Else
   Call clearscreen
   cmdfirst.Enabled = False
   cmdPrev.Enabled = False
   cmdNext.Enabled = False
   cmdlast.Enabled = False
   cmdedit.Enabled = False
   cmddelete.Enabled = False
End If

End Sub

Private Sub saveclick()
Dim res
Set rsout = New ADODB.Recordset
rsout.Open "Select * from HREmployeesMovement", cnDCPI, adOpenKeyset, adLockOptimistic
With rsout
    If Not (.BOF And .EOF) Then
    .MoveLast
    End If
    .AddNew
   ' !movementnum = txtmovementno.Text
    !employeenum = txtEmpNo.Text
    !movementdate = IIf(mskmovementdate.Text = "__/__/____", Null, Format(mskmovementdate.Text, "mm/dd/yyyy"))
    !movementtype = Left(cmbmovementtype.Text, 1)
    !movementremarks = txtmovementremarks.Text
    !EmployeeName = cmbempname.Text
    '!clientcode = Left(cmbclientcode.Text, 5)
'    !salary = txtsalary.Text
'    !movementsequence = txtmovementseq.Text
'    !employmenttype = Left(cmbemploymenttype.Text, 2)
'    !jobcode = Left(cmbjobcode.Text, 2)
'    !jobclass = Left(cmbjobclass.Text, 2)
 '   !Status = txtstatus.Text
        !userid = operatorid
'     If Trim(cmbclientcode.Text) = "" Then
'          !ClientCode = 0
'      Else
'        If Me.cmbclientcode.ListIndex = -1 Then
'            !ClientCode = ClientCode
'        Else
'            !ClientCode = Me.cmbclientcode.ItemData(Me.cmbclientcode.ListIndex)
'        End If
'         End If
         
    .Update
End With
txtmovementno.Text = Getmax
'res = MsgBox("Record Added", vbInformation, CompanyName)
End Sub
Function output() As String
    Dim rs As ADODB.Recordset
    Dim str As String
     str = "Select max(movementnum) From HREmployeesMovement"
    Set rs = cnDCPI.Execute(str)
    If Not rs.EOF Then
         output = Format(rs.Fields(0) + 1, "")
    Else
       output = "00000001"
    End If
End Function
Function Getmax() As Integer
    Dim rs As ADODB.Recordset
    Dim str As String
    str = "Select max(movementnum) From HREmployeesMovement"
        Set rs = cnDCPI.Execute(str)
    If Not rs.EOF Then
         Getmax = rs.Fields(0)
    End If
End Function
Private Sub clearscreen()
txtmovementno.Text = "<auto number>"
txtEmpNo.Text = ""
mskmovementdate.Text = "__/__/____"
cmbmovementtype.ListIndex = -1
txtmovementremarks.Text = ""
cmbempname.ListIndex = -1

End Sub
Private Sub Showdata(ByRef rec As ADODB.Recordset)
Dim rs As ADODB.Recordset
Dim rsout As ADODB.Recordset
If rec.EOF = True Then
Call clearscreen
cmdedit.Enabled = False
cmddelete.Enabled = False

Else
    txtmovementno.Text = IIf(IsNull(rec!movementnum), "000000", rec!movementnum)
    txtEmpNo.Text = IIf(IsNull(rec!employeenum), "", rec!employeenum)
    mskmovementdate.Text = IIf(IsNull(rec!movementdate), "__/__/____", Format(rec!movementdate, "mm/dd/yyyy"))
    'cmbmovementtype.Text = IIf(IsNull(Rec!movementtype), "", Rec!movementtype)
    txtmovementremarks.Text = IIf(IsNull(rec!movementremarks), "", rec!movementremarks)
End If

Set rs = cnDCPI.Execute("Select * From HRMovementType Where MovementType ='" & IIf(IsNull(rec!movementtype), "", rec!movementtype) & "'")
        If Not rs.EOF Then
            cmbmovementtype.Text = rs.Fields(0) & "-" & rs.Fields(1)
        Else
            cmbmovementtype.ListIndex = -1
        End If
        rs.Close


Set rs = cnDCPI.Execute("Select * From HREmployees Where employeenum =" & IIf(IsNull(rec!employeenum), "", rec!employeenum) & "")
        If Not rs.EOF Then
                 cmbempname.Text = rs.Fields(4) & "," & rs.Fields(2)
        Else
            cmbempname.ListIndex = -1
        End If
        rs.Close

End Sub

Private Sub editsave()
Dim res
With rsout
   ' !movementnum = txtmovementno.Text
    !employeenum = txtEmpNo.Text
    !movementdate = IIf(mskmovementdate.Text = "__/__/____", Null, Format(mskmovementdate.Text, "MM/DD/YYYY"))
    !movementtype = Left(cmbmovementtype.Text, 1)
    !movementremarks = txtmovementremarks.Text
    !EmployeeName = cmbempname.Text
        !userid = operatorid
    '!clientcode = Left(cmbclientcode.Text, 5)
'    !salary = txtsalary.Text
'    !movementsequence = txtmovementseq.Text
'    !employmenttype = Left(cmbemploymenttype.Text, 2)
'    !jobcode = Left(cmbjobcode.Text, 2)
'    !jobclass = Left(cmbjobclass.Text, 2)
 ' !Status = me.txtstatus.Text
    
'     If cmbclientcode.Text = "" Then
'            !ClientCode = 0
'        Else
'            !ClientCode = Me.cmbclientcode.ItemData(Me.cmbclientcode.ListIndex)
'        End If
    .Update
End With
'res = MsgBox("Record Updated", vbInformation, CompanyName)
End Sub


Private Sub Timer1_Timer()
'Label4.Left = Label4.Left + 60
'If Label4.Left > Transmissionzip.Width Then Label4.Left = 0
End Sub

Private Sub txttransactionnum_GotFocus()
txttransactionnum.BackColor = &HFFFFC0
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
End Sub


Private Sub txtzipfilename_LostFocus()
txtzipfilename.BackColor = vbWhite
End Sub

Private Sub txtclientcode_GotFocus()
txtClientCode.BackColor = &HFFFFC0
End Sub



Private Sub txtclientcode_LostFocus()
txtClientCode.BackColor = vbWhite
End Sub

Private Sub txtemployeeno_GotFocus()
txtEmployeeNo.BackColor = &HFFFFC0
End Sub



Private Sub txtemployeeno_LostFocus()
txtEmployeeNo.BackColor = vbWhite
End Sub

Private Sub txtjobclass_GotFocus()
txtjobclass.BackColor = &HFFFFC0
End Sub



Private Sub txtjobclass_LostFocus()
txtjobclass.BackColor = vbWhite
End Sub

Private Sub txtjobcode_GotFocus()
txtjobcode.BackColor = &HFFFFC0
End Sub


Private Sub txtjobcode_LostFocus()
txtjobcode.BackColor = vbWhite
End Sub

Private Sub mskmovementdate_Validate(Cancel As Boolean)
If mskmovementdate <> "__/__/____" And IsDate(mskmovementdate) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskmovementdate.SetFocus

End Sub

Private Sub txtempno_GotFocus()
Set oCommonFunction.ValidAlpha = txtEmpNo
End Sub

Private Sub txtmovementno_GotFocus()
txtmovementno.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtmovementno
End Sub



Private Sub txtmovementno_LostFocus()
txtmovementno.BackColor = vbWhite
If IsNumeric(txtmovementno.Text) = True Then
Else
    MsgBox "Should be numeric", vbInformation, CompanyName
    txtmovementno.SetFocus
End If
End Sub
Private Sub cmbempname_GotFocus()
cmbempname.BackColor = &HFFFFC0
End Sub

Private Sub cmbempname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
Call FindMatch(cmbempname, KeyAscii)
End Sub

Private Sub cmbempname_LostFocus()
cmbempname.BackColor = vbWhite
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
