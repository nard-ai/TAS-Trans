VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHRAppraisal 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HR Appraisal Data Entry"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "frmHRAppraisal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   7875
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   90
      TabIndex        =   10
      Top             =   585
      Width           =   7695
      Begin VB.TextBox txtremarks 
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
         Height          =   600
         Left            =   5625
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1575
         Width           =   1935
      End
      Begin VB.CommandButton cmdfirst 
         Height          =   255
         Left            =   6525
         Picture         =   "frmHRAppraisal.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Top record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPrev 
         Height          =   255
         Left            =   6765
         Picture         =   "frmHRAppraisal.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Previous record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdNext 
         Height          =   255
         Left            =   7005
         Picture         =   "frmHRAppraisal.frx":081E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Next record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdlast 
         Height          =   255
         Left            =   7245
         Picture         =   "frmHRAppraisal.frx":0968
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   1575
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1155
         Width           =   2205
      End
      Begin VB.TextBox txtempno 
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
         Left            =   1575
         TabIndex        =   1
         Top             =   765
         Width           =   2205
      End
      Begin VB.ComboBox cmbschedule 
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
         Left            =   1575
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "cmbschedule"
         Top             =   1530
         Width           =   2205
      End
      Begin VB.TextBox txttotal 
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
         Left            =   5640
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1215
         Width           =   1935
      End
      Begin VB.TextBox txtpersonal 
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
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   7
         Top             =   765
         Width           =   1935
      End
      Begin VB.TextBox txtperformance 
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
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   6
         Top             =   375
         Width           =   1935
      End
      Begin VB.TextBox txtappraisalno 
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
         Left            =   1575
         TabIndex        =   0
         Top             =   360
         Width           =   2205
      End
      Begin VB.TextBox txtevalby 
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
         Left            =   1575
         TabIndex        =   5
         Top             =   2430
         Width           =   2205
      End
      Begin MSMask.MaskEdBox mskappraisaldate 
         Height          =   330
         Left            =   1575
         TabIndex        =   4
         Top             =   1995
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   582
         _Version        =   393216
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks/Action taken"
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
         Index           =   2
         Left            =   4005
         TabIndex        =   31
         Top             =   1710
         Width           =   1575
      End
      Begin VB.Label Label2 
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
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   1290
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Performance/Work Factor"
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
         Height          =   465
         Index           =   3
         Left            =   3960
         TabIndex        =   18
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   4
         Left            =   3960
         TabIndex        =   17
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Factor"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   16
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appraisal Sched"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appraisal No."
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
         Left            =   240
         TabIndex        =   14
         Top             =   495
         Width           =   960
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
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appraisal Date"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2115
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evaluated By"
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
         Left            =   240
         TabIndex        =   11
         Top             =   2550
         Width           =   945
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdrefresh 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         Picture         =   "frmHRAppraisal.frx":0AB2
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Refresh"
         Top             =   -15
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         Picture         =   "frmHRAppraisal.frx":0BFC
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "New"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton cmdsave 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1215
         Picture         =   "frmHRAppraisal.frx":112E
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2385
         Picture         =   "frmHRAppraisal.frx":1660
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdedit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   600
         Picture         =   "frmHRAppraisal.frx":1B92
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         Picture         =   "frmHRAppraisal.frx":1CDC
         Style           =   1  'Graphical
         TabIndex        =   25
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
Attribute VB_Name = "frmHRAppraisal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsout As ADODB.Recordset
Dim id As String
Dim trannum As String
Dim outputfilename As String
Dim rsactionType As ADODB.Recordset
Dim rsemp As ADODB.Recordset
Dim rs As ADODB.Recordset
'Set oCommonFunction.ValidAlpha = txtaddress1
'If KeyAscii = 13 Then SendKeys "{tab}"

Private Sub cmbempname_Click()
Call EmpField
End Sub
Private Sub EmpField()
'************
'search additional field
'************
Dim res
Dim Search, search1, Proj, str As String
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
        Search = cmbempname.Text
End If

'If IsEmpty(txtcustomercode.Text) = False And IsEmpty(txtprojectcode.Text) = False Then
If IsEmpty(cmbempname.Text) = False Then
    'SelectString = "Select * from dcprojectcounty where projectshort='" & str & "' "
    SelectString = "select * from name where name='" & Search & "' "
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
Call FindMatch(cmbempname, KeyAscii)
End Sub

Private Sub cmbschedule_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub



Private Sub mskappraisaldate_Validate(Cancel As Boolean)
If mskappraisaldate <> "__/__/____" And IsDate(mskappraisaldate) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskappraisaldate.SetFocus

End Sub

Private Sub txtempno_GotFocus()
TxtControlNo.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = TxtControlNo
End Sub


Private Sub cmbemp_LostFocus()
cmbemp.BackColor = vbWhite
End Sub

Private Sub cmbschedule_GotFocus()
cmbschedule.BackColor = &HFFFFC0
End Sub

Private Sub cmbschedule_LostFocus()
cmbschedule.BackColor = vbWhite
End Sub

Private Sub cmdAdd_Click()
'On Error GoTo Errh
addflagship = True
editflagship = False
AddFlag = True
Call Add
Call clearscreen
'txtControlNo.Text = Format(Date, "yy") & "-" & Format(Date, "mm") & "-" & Format(NEWID, "00000")

txtappraisalno.Enabled = False
txtempno.Enabled = False
cmbempname.Enabled = True
cmbempname.Text = ""
cmbschedule.Enabled = True
txtperformance.Enabled = True
txtpersonal.Enabled = True
txttotal.Enabled = True
'cmbactiontaken.Enabled = True
mskappraisaldate.Enabled = True
txtevalby.Enabled = True
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
  '  cmdExit.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    addflagship = False
    editflagship = False
    Set rsout = New ADODB.Recordset
    StrQuery = "Select * from HREmployeesappraisal"
    Set rsout = ReturnRst(stquery, cnnstr, mstatic)
    'rsout.Open "Select * from HREmployeesappraisal ", cnnHR, adOpenKeyset, adLockOptimistic
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
Dim movementnum As String

Set rsout = New ADODB.Recordset
StrQuery = "Select * from HREmployeesappraisal "
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HREmployeesappraisal ", cnnHR, adOpenKeyset, adLockOptimistic

res = MsgBox("Delete record", vbQuestion + vbYesNo, CompanyName)
If res = vbYes Then
    str = "Delete from HREmployeesappraisal where appraisalnum=" & txtappraisalno.Text
    Set rsout = cnn.Execute(str)
    StrQuery = "Select * from HREmployeesappraisal"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "Select * from HREmployeesappraisal", cnnHR, adOpenKeyset, adLockOptimistic
         If rsout.EOF = False Then
         Showdata rsout
          cmdNext.Enabled = True
          cmdlast.Enabled = True
          Else
          cmdNext.Enabled = False
          cmdlast.Enabled = False
          cmdfirst.Enabled = False
          cmddelete.Enabled = False
          
          Call clearscreen
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
txtappraisalno.Enabled = True
txtempno.Enabled = False
cmbempname.Enabled = True
cmbschedule.Enabled = True
txtperformance.Enabled = True
txtpersonal.Enabled = True
txttotal.Enabled = True
'cmbactiontaken.Enabled = True
mskappraisaldate.Enabled = True
txtevalby.Enabled = True
txtRemarks.Enabled = True
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
      MsgBox "Beginning of file encountered", vbInformation, CompanyName
    End If
    
   appraisalnum = rsout!appraisalnum
    Showdata rsout
X:
    Exit Sub
End Sub

Private Sub cmdSave_Click()
If cmbempname.Text = "" Then MsgBox "Select employee name. ", vbInformation: cmbempname.SetFocus: Exit Sub

txtappraisalno.Enabled = False
txtempno.Enabled = False
cmbempname.Enabled = False
cmbschedule.Enabled = False
txtperformance.Enabled = False
txtpersonal.Enabled = False
txttotal.Enabled = False
txtRemarks.Enabled = False
'cmbactiontaken.Enabled = False
mskappraisaldate.Enabled = False
txtevalby.Enabled = False
'Add new record
        If AddFlag = True Or addflagship = True Then

           Call save
           Call SaveNewHREmployeeRecord
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
txtappraisalno.Enabled = False
txtempno.Enabled = False
cmbempname.Enabled = False
cmbschedule.Enabled = False
txtperformance.Enabled = False
txtpersonal.Enabled = False
txttotal.Enabled = False
txtRemarks.Enabled = False
'cmbactiontaken.Enabled = False
mskappraisaldate.Enabled = False
txtevalby.Enabled = False
End Sub

Private Sub Form_Load()
   Me.Top = 0 '(mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
Set rsMoveType = New ADODB.Recordset
Set rsEmpType = New ADODB.Recordset
Set rsjobclass = New ADODB.Recordset
Set rsjobcode = New ADODB.Recordset
Set rstclient = New ADODB.Recordset
Set rsemp = New ADODB.Recordset

addflagship = False
editflagship = False
txtappraisalno.Enabled = False
txtempno.Enabled = False
cmbempname.Enabled = False
cmbschedule.Enabled = False
txtperformance.Enabled = False
txtpersonal.Enabled = False
txttotal.Enabled = False
'cmbactiontaken.Enabled = False
mskappraisaldate.Enabled = False
txtevalby.Enabled = False
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdrefresh.Enabled = False
Call month
Set rsout = New ADODB.Recordset

'populate of combo boxes
StrQuery = "Select * From HREmployees order by surname, firstname"
Set rsemp = ReturnRst(StrQuery, cnnstr, mstatic)
'Set rsemp = cnn.Execute("Select * From HREmployees order by surname, firstname")
Do While Not rsemp.EOF
    cmbempname.AddItem rsemp!Surname & "," & rsemp!FirstName
    rsemp.MoveNext
Loop
rsemp.Close
    
'Set rsactionType = cnn.Execute("Select * From HRappraisalactiontaken")
'Do While Not rsactionType.EOF
'        cmbactiontaken.AddItem rsactionType!actiontaken & "-" & rsactionType!actiontakendesc
'        rsactionType.MoveNext
'    Loop
'    rsactionType.Close
ShowAppraisal

End Sub

Public Sub ShowAppraisal(Optional ByVal appraisalnum As Integer)

If rsout.State = 1 Then rsout.Close
If appraisalnum = 0 Then
    StrQuery = "select * from HREmployeesAppraisal"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "select * from HREmployeesAppraisal", cnnHR, adOpenKeyset, adLockOptimistic
Else
    StrQuery = "Select * FROM HREmployeesAppraisal where appraisalnum = " & appraisalnum & ""
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "select * from HREmployeesAppraisal where appraisalnum=" & appraisalnum & "", cnnHR, adOpenKeyset, adLockOptimistic
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

Private Sub SaveNewHREmployeeRecord()
Dim res

' For Update 2010, bernard.
Set rsout = New ADODB.Recordset

'strQuery = "Select * from HREmployeesAppraisal"
'Set rsout = ReturnRst(strQuery, cnnstr, mstatic)
''rsout.Open "Select * from HREmployeesAppraisal", cnnHR, adOpenKeyset, adLockOptimistic
'
'With rsout
'    If Not (.BOF And .EOF) Then
'    .MoveLast
'    End If
'    .AddNew
''    !appraisalnum = txtappraisalno.Text
'    !EmployeeID = txtempno.Text
'    !EmployeeName = cmbempname.Text
'    !appraisaldate = IIf(mskappraisaldate.Text = "__/__/____", Null, Format(mskappraisaldate.Text, "mm/dd/yyyy"))
'    !evalby = txtevalby.Text
'    !appraisalsched = cmbschedule.Text
'    !performancefactor = txtperformance.Text
'    !personalfactor = txtpersonal.Text
'    !Total = txttotal.Text
'    !actiontaken = txtRemarks.Text
'    !USERID = operatorID
'    .Update
'    txtappraisalno.Text = Getmax
'End With
''res = MsgBox("Record Added", vbInformation, CompanyName)
End Sub

Function Getmax() As Integer
    Dim rs As ADODB.Recordset
    Dim str As String
     str = "Select max(appraisalnum) From HREmployeesAppraisal"
    Set rs = cnn.Execute(str)
    If Not rs.EOF Then
         Getmax = rs.Fields(0)
    End If
End Function


Private Sub clearscreen()
txtappraisalno.Text = "<auto number>"
txtempno.Text = ""
cmbempname.ListIndex = -1
mskappraisaldate.Text = "__/__/____"
cmbschedule.Text = ""
txtpersonal.Text = ""
txtperformance.Text = ""
txttotal.Text = ""
'cmbactiontaken.ListIndex = -1
txtevalby.Text = ""
End Sub
Private Sub Showdata(ByRef rec As ADODB.Recordset)
Dim rs As ADODB.Recordset
Dim rsout As ADODB.Recordset
If rec.EOF = True Then
Call clearscreen
cmdedit.Enabled = False
cmddelete.Enabled = False

Else
    txtappraisalno.Text = IIf(IsNull(rec!appraisalnum), "000000", rec!appraisalnum)
    mskappraisaldate.Text = IIf(IsNull(rec!appraisaldate), "__/__/____", Format(rec!appraisaldate, "mm/dd/yyyy"))
    txtperformance.Text = IIf(IsNull(rec!performancefactor), "", rec!performancefactor)
    txtpersonal.Text = IIf(IsNull(rec!personalfactor), "", rec!personalfactor)
    txttotal.Text = IIf(IsNull(rec!Total), "", rec!Total)
    txtevalby.Text = IIf(IsNull(rec!evalby), "", rec!evalby)
    cmbschedule.Text = IIf(IsNull(rec!appraisalsched), "", rec!appraisalsched)
    txtempno.Text = IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
    txtRemarks.Text = IIf(IsNull(rec!actiontaken), "", rec!actiontaken)
'show populate combo
End If

    StrQuery = "SELECT * FROM HREmployees WHERE EmployeeID =" & IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
    Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
    'Set rs = cnn.Execute("Select * From HREmployees Where EmployeeID =" & IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID) & "")
        If Not rs.EOF Then
            cmbempname.Text = rs.Fields(4) & "," & rs.Fields(2)
        Else
            cmbempname.ListIndex = -1
        End If
        rs.Close
        
'

End Sub

Private Sub editsave()
Dim res
With rsout
'    !appraisalnum = txtappraisalno.Text
    !EmployeeID = txtempno.Text
    !EmployeeName = cmbempname.Text
    !appraisaldate = IIf(mskappraisaldate.Text = "__/__/____", Null, mskappraisaldate.Text)
    !evalby = txtevalby.Text
    !appraisalsched = cmbschedule.Text
    !performancefactor = txtperformance.Text
    !personalfactor = txtpersonal.Text
    !Total = txttotal.Text
    !USERID = operatorID
   !actiontaken = txtRemarks.Text
    .Update
End With
'res = MsgBox("Record Updated", vbInformation, CompanyName)
End Sub

Private Sub month()
cmbschedule.AddItem "01 - January"
cmbschedule.AddItem "02 - February"
cmbschedule.AddItem "03 - March"
cmbschedule.AddItem "04 - April"
cmbschedule.AddItem "05 - May"
cmbschedule.AddItem "06 - June"
cmbschedule.AddItem "07 - July"
cmbschedule.AddItem "08 - August"
cmbschedule.AddItem "09 - September"
cmbschedule.AddItem "10 - October"
cmbschedule.AddItem "11 - November"
cmbschedule.AddItem "12 - December"
End Sub

Private Sub mskappraisaldate_GotFocus()
mskappraisaldate.BackColor = &HFFFFC0
End Sub

Private Sub mskappraisaldate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskappraisaldate_LostFocus()
mskappraisaldate.BackColor = vbWhite
End Sub

Private Sub txtappraisalno_GotFocus()
txtappraisalno.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtappraisalno
End Sub


Private Sub txtappraisalno_LostFocus()
txtappraisalno.BackColor = vbWhite
End Sub

Private Sub txtevalby_GotFocus()
txtevalby.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtevalby
End Sub


Private Sub txtevalby_LostFocus()
txtevalby.BackColor = vbWhite
End Sub

Private Sub txtperformance_GotFocus()
txtperformance.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txtperformance
End Sub


Private Sub txtperformance_LostFocus()
txtperformance.BackColor = vbWhite
End Sub

Private Sub txtpersonal_GotFocus()
txtpersonal.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txtpersonal
End Sub


Private Sub txtpersonal_LostFocus()
txtpersonal.BackColor = vbWhite
End Sub

Private Sub txtremarks_GotFocus()
txtRemarks.BackColor = &HFFFFC0
End Sub

Private Sub txtremarks_LostFocus()
txtRemarks.BackColor = vbWhite
End Sub

Private Sub txttotal_GotFocus()
txttotal.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txttotal
'txttotal.Text = Format(Val(txtperformance.Text) + Val(txtpersonal.Text), "###.##")
End Sub


Private Sub txttotal_LostFocus()
txttotal.BackColor = vbWhite
End Sub
