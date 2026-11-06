VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHROffense 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HR Employee Offenses"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frmHROffense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   8040
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   120
      TabIndex        =   12
      Top             =   630
      Width           =   7815
      Begin VB.CommandButton cmdfirst 
         Height          =   255
         Left            =   6660
         Picture         =   "frmHROffense.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Top record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPrev 
         Height          =   255
         Left            =   6900
         Picture         =   "frmHROffense.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Previous record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdNext 
         Height          =   255
         Left            =   7140
         Picture         =   "frmHROffense.frx":081E
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Next record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdlast 
         Height          =   255
         Left            =   7380
         Picture         =   "frmHROffense.frx":0968
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Last record"
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtcoundate 
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
         Left            =   5640
         TabIndex        =   8
         Top             =   1200
         Width           =   1980
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
         Left            =   1530
         TabIndex        =   1
         Top             =   765
         Width           =   2160
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
         Left            =   1530
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   2160
      End
      Begin VB.TextBox txtoffenseno 
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
         Left            =   1530
         TabIndex        =   0
         Top             =   360
         Width           =   2160
      End
      Begin VB.TextBox txtoffensemade 
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
         Left            =   1530
         TabIndex        =   3
         Top             =   1620
         Width           =   2160
      End
      Begin VB.ComboBox cmboffensetype 
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
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2070
         Width           =   2160
      End
      Begin VB.ComboBox cmbactiontype 
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
         Left            =   5625
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   375
         Width           =   1980
      End
      Begin VB.TextBox txtcounfacilitator 
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
         Left            =   5640
         TabIndex        =   9
         Top             =   1635
         Width           =   1980
      End
      Begin VB.TextBox txtcouninfo 
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
         Left            =   5640
         TabIndex        =   10
         Top             =   2070
         Width           =   1980
      End
      Begin VB.TextBox txtremarks 
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
         Left            =   5640
         TabIndex        =   11
         Top             =   2505
         Width           =   1980
      End
      Begin MSMask.MaskEdBox mskoffensedate 
         Height          =   285
         Left            =   1530
         TabIndex        =   5
         Top             =   2505
         Width           =   2160
         _ExtentX        =   3810
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
      Begin MSMask.MaskEdBox mskdateserved 
         Height          =   285
         Left            =   5640
         TabIndex        =   7
         Top             =   765
         Width           =   1980
         _ExtentX        =   3493
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
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Offense No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   495
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
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
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   885
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Offense Made"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1755
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Nature of Offense"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   2190
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Offense Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2625
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Action Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3960
         TabIndex        =   18
         Top             =   495
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Served"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3960
         TabIndex        =   17
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Counseling/ Suspension Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   645
         Left            =   3960
         TabIndex        =   16
         Top             =   1230
         Width           =   1320
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Coun. Facilitator"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3960
         TabIndex        =   15
         Top             =   1755
         Width           =   1185
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Counseling Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3960
         TabIndex        =   14
         Top             =   2145
         Width           =   1125
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3960
         TabIndex        =   13
         Top             =   2580
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   8040
      _ExtentX        =   14182
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
         Picture         =   "frmHROffense.frx":0AB2
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   2340
         Picture         =   "frmHROffense.frx":0BFC
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1200
         Picture         =   "frmHROffense.frx":112E
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmHROffense.frx":1660
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "New"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   1755
         Picture         =   "frmHROffense.frx":1B92
         Style           =   1  'Graphical
         TabIndex        =   30
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
Attribute VB_Name = "frmHROffense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Set oCommonFunction.ValidAlpha = txtaddress1
'If KeyAscii = 13 Then SendKeys "{tab}"

Dim rsout As ADODB.Recordset
Dim id As String
Dim trannum As String
Dim outputfilename As String

Private Sub cmbemploymenttype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbemp_Click()
'Call EmpName
End Sub

Private Sub cmbactiontype_GotFocus()
cmbactiontype.BackColor = &HFFFFC0
End Sub

Private Sub cmbactiontype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbactiontype_LostFocus()
cmbactiontype.BackColor = vbWhite
End Sub

Private Sub cmbempname_Click()
Call EmpField
End Sub

Private Sub cmbempname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
Call FindMatch(cmbempname, KeyAscii)
End Sub

Private Sub cmbmovementtype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmboffensetype_GotFocus()
cmboffensetype.BackColor = &HFFFFC0
End Sub

Private Sub cmboffensetype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmboffensetype_LostFocus()
cmboffensetype.BackColor = vbWhite
End Sub

Private Sub cmdAdd_Click()
'On Error GoTo Errh
addflagship = True
editflagship = False

Call Add
Call clearscreen

TxtControlNo.Enabled = False
cmbempname.Enabled = True
txtoffensemade.Enabled = True
cmboffensetype.Enabled = True
mskoffensedate.Enabled = True
'txtdiscaction.Enabled = True
cmbactiontype.Enabled = True
mskdateserved.Enabled = True
txtcoundate.Enabled = True
txtcounfacilitator.Enabled = True
txtcouninfo.Enabled = True
txtRemarks.Enabled = True
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
    cmdfirst.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdlast.Enabled = False
    'cmdExit.Enabled = False
End Sub

Private Sub cmdCancel_Click()

addflagship = False
editflagship = False
       
Set rsout = New ADODB.Recordset
StrQuery = "Select * from HREmployeesOffenses "
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HREmployeesOffenses ", cnnHR, adOpenKeyset, adLockOptimistic

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
Dim offensenum As String

Set rsout = New ADODB.Recordset
StrQuery = "Select * from HREmployeesOffenses "
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HREmployeesOffenses ", cnnHR, adOpenKeyset, adLockOptimistic

res = MsgBox("Delete record", vbQuestion + vbYesNo, CompanyName)
If res = vbYes Then
    str = "Delete from HREmployeesOffenses where offensenum=" & txtoffenseno.Text
    Set rsout = cnnHR.Execute(str)
  '  res = MsgBox("Record Deleted", vbCritical, CompanyName)
    StrQuery = "Select * from HREmployeesOffenses"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "Select * from HREmployeesOffenses", cnnHR, adOpenKeyset, adLockOptimistic
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
TxtControlNo.Enabled = False
cmbempname.Enabled = True
txtoffensemade.Enabled = True
cmboffensetype.Enabled = True
mskoffensedate.Enabled = True
'txtdiscaction.Enabled = True
cmbactiontype.Enabled = True
mskdateserved.Enabled = True
txtcoundate.Enabled = True
txtcounfacilitator.Enabled = True
txtcouninfo.Enabled = True
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
    
   offensenum = rsout!offensenum
    Showdata rsout
X:
    Exit Sub
End Sub


Private Sub cmdSave_Click()
If cmbempname.Text = "" Then MsgBox "Select employee name. ", vbInformation: cmbempname.SetFocus: Exit Sub

TxtControlNo.Enabled = False
cmbempname.Enabled = False
txtoffensemade.Enabled = False
cmboffensetype.Enabled = False
mskoffensedate.Enabled = False
'txtdiscaction.Enabled = False
cmbactiontype.Enabled = False
mskdateserved.Enabled = False
txtcoundate.Enabled = False
txtcounfacilitator.Enabled = False
txtcouninfo.Enabled = False
txtRemarks.Enabled = False
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
'cmddelete.Enabled = True
cmdfirst.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdlast.Enabled = True
'cmdExit.Enabled = True
TxtControlNo.Enabled = False
cmbempname.Enabled = False
txtoffensemade.Enabled = False
cmboffensetype.Enabled = False
mskoffensedate.Enabled = False
'txtdiscaction.Enabled = False
cmbactiontype.Enabled = False
mskdateserved.Enabled = False
txtcoundate.Enabled = False
txtcounfacilitator.Enabled = False
txtcouninfo.Enabled = False
txtRemarks.Enabled = False
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

Dim rsTemp As ADODB.Recordset

Set rsTemp = New ADODB.Recordset
Set rsemp = New ADODB.Recordset
Call Load_ref_table
addflagship = False
editflagship = False
txtempno.Enabled = False
cmbempname.Enabled = False
txtoffensemade.Enabled = False
cmboffensetype.Enabled = False
mskoffensedate.Enabled = False
'txtdiscaction.Enabled = False
cmbactiontype.Enabled = False
mskdateserved.Enabled = False
txtcoundate.Enabled = False
txtcounfacilitator.Enabled = False
txtcouninfo.Enabled = False
txtRemarks.Enabled = False
cmdsave.Enabled = False
cmdcancel.Enabled = False
Set rsout = New ADODB.Recordset

StrQuery = "Select * From HREmployees order by surname, firstname"
Set rsTemp = ReturnRst(StrQuery, cnnstr, mstatic)
'Set rsTemp = cnn.Execute("Select * From HREmployees order by surname, firstname")
Do While Not rsTemp.EOF
        cmbempname.AddItem rsTemp!Surname & "," & rsTemp!FirstName
        rsTemp.MoveNext
Loop
rsTemp.Close
ShowOffenses
'open database
End Sub

Public Sub ShowOffenses(Optional ByVal offensenum As Integer)
If rsout.State = 1 Then rsout.Close
If offensenum = 0 Then
    StrQuery = "select * from HREmployeesOffenses"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "select * from HREmployeesOffenses", cnnHR, adOpenKeyset, adLockOptimistic
Else
    StrQuery = "select * from HREmployeesOffenses where offensenum=" & offensenum & ""
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "select * from HREmployeesOffenses where offensenum=" & offensenum & "", cnnHR, adOpenKeyset, adLockOptimistic
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



Private Sub Load_ref_table()
Set rsoffenseType = New ADODB.Recordset
Set rsactionType = New ADODB.Recordset

Set rsoffenseType = cnn.Execute("Select * From HRoffenseType")
Do While Not rsoffenseType.EOF
        cmboffensetype.AddItem rsoffenseType!offensetype & "-" & rsoffenseType!offensetypedesc
        rsoffenseType.MoveNext
    Loop
    rsoffenseType.Close

Set rsactionType = cnn.Execute("Select * From HRactionType")
Do While Not rsactionType.EOF
        cmbactiontype.AddItem rsactionType!actiontype & "-" & rsactionType!actiontypedesc
        rsactionType.MoveNext
    Loop
    rsactionType.Close

End Sub

Private Sub SaveNewHREmployeeRecord()
Dim res

' For Update 2010, bernard.
Set rsout = New ADODB.Recordset
'strQuery = "Select * from HREmployeesOffenses"
'Set rsout = ReturnRst(strQuery, cnnstr, mstatic)
''rsout.Open "Select * from HREmployeesOffenses", cnnHR, adOpenKeyset, adLockOptimistic
''**************
''save new record
''**************
'
'With rsout
'    If Not (.BOF And .EOF) Then
'    .MoveLast
'    End If
'    .AddNew
'    !EmployeeID = TxtControlNo.Text
'    !EmployeeName = cmbempname.Text
'    !offensemade = txtoffensemade.Text
'    !offensetype = Left(cmboffensetype.Text, 1)
'    !offensedate = IIf(mskoffensedate.Text = "__/__/____", Null, Format(mskoffensedate.Text, "mm/dd/yyyy"))
'    '!disciplinaryaction = txtdiscaction.Text
'    !actiontype = Left(cmbactiontype.Text, 1)
'    !dateserved = IIf(mskdateserved.Text = "__/__/____", Null, Format(mskdateserved.Text, "mm/dd/yyyy"))
'    !counselingdate = txtcoundate.Text
'    !counselingfacilitator = txtcounfacilitator.Text
'    !counselinginfo = txtcouninfo.Text
'    !Remarks = txtRemarks.Text
'        !USERID = operatorID
'    .Update
'End With
'Me.txtoffenseno.Text = Getmax
''res = MsgBox("Record Added", vbInformation, CompanyName)
End Sub

Function Getmax() As Integer
    Dim rs As ADODB.Recordset
    Dim str As String
    str = "Select max(offensenum) From HREmployeesOffenses"
        Set rs = cnn.Execute(str)
    If Not rs.EOF Then
         Getmax = rs.Fields(0)
    End If
End Function

Private Sub clearscreen()
txtoffenseno.Text = "<auto number>"
TxtControlNo.Text = ""
cmbempname.ListIndex = -1
txtoffensemade.Text = ""
cmboffensetype.ListIndex = -1
mskoffensedate.Text = "__/__/____"
'txtdiscaction.Text = ""
cmbactiontype.ListIndex = -1
mskdateserved.Text = "__/__/____"
txtcoundate.Text = ""
txtcounfacilitator.Text = ""
txtcouninfo.Text = ""
txtRemarks.Text = ""
End Sub

Private Sub Showdata(ByRef rec As ADODB.Recordset)
Dim rs As ADODB.Recordset
Dim rsout As ADODB.Recordset
'**************
'show record
'**************
If rec.EOF = True Then
Call clearscreen
cmdedit.Enabled = False
cmddelete.Enabled = False

Else
    
    txtoffenseno.Text = IIf(IsNull(rec!offensenum), "000000", rec!offensenum)
    txtempno.Text = IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
    txtoffensemade.Text = IIf(IsNull(rec!offensemade), "", rec!offensemade)
   ' cmboffensetype.Text = IIf(IsNull(Rec!offensetype), "", Rec!offensetype)
    mskoffensedate.Text = IIf(IsNull(rec!offensedate), "__/__/____", Format(rec!offensedate, "mm/dd/yyyy"))
  '  cmbactiontype.Text = IIf(IsNull(Rec!actiontype), "", Rec!actiontype)
    mskdateserved.Text = IIf(IsNull(rec!dateserved), "__/__/____", Format(rec!dateserved, "mm/dd/yyyy"))
    txtcoundate.Text = IIf(IsNull(rec!counselingdate), "", rec!counselingdate)
    txtcounfacilitator.Text = IIf(IsNull(rec!counselingfacilitator), "", rec!counselingfacilitator)
   ' txtdiscaction.Text = IIf(IsNull(Rec!disciplinaryaction), "", Rec!disciplinaryaction)
    txtcouninfo.Text = IIf(IsNull(rec!counselinginfo), "", rec!counselinginfo)
    txtRemarks.Text = IIf(IsNull(rec!Remarks), "", rec!Remarks)
 End If
 
Set rs = cnn.Execute("Select * From HRoffenseType Where offenseType ='" & IIf(IsNull(rec!offensetype), "", rec!offensetype) & "'")
        If Not rs.EOF Then
            cmboffensetype.Text = rs.Fields(0) & "-" & rs.Fields(1)
        Else
            cmboffensetype.ListIndex = -1
        End If
        rs.Close

Set rs = cnn.Execute("Select * From HRactionType Where actionType ='" & IIf(IsNull(rec!actiontype), "", rec!actiontype) & "'")
        If Not rs.EOF Then
            cmbactiontype.Text = rs.Fields(0) & "-" & rs.Fields(1)
        Else
            cmbactiontype.ListIndex = -1
        End If
        rs.Close

StrQuery = "SELECT * FROM HREmployees WHERE EmployeeID =" & IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID) & ""
Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
'Set rs = cnn.Execute("Select * From HREmployees Where EmployeeID =" & IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID) & "")
If Not rs.EOF Then
    cmbempname.Text = rs.Fields(4) & "," & rs.Fields(2)
Else
    cmbempname.ListIndex = -1
End If
rs.Close

End Sub

Private Sub editsave()
Dim res
'**************
'save edited record
'**************

' For Revision by Bernard.
With rsout
    !EmployeeID = TxtControlNo.Text
    !EmployeeName = cmbempname.Text
    !offensemade = txtoffensemade.Text
    !offensetype = Left(cmboffensetype.Text, 1)
    !offensedate = IIf(mskoffensedate.Text = "__/__/____", Null, mskoffensedate.Text)
    '!disciplinaryaction = txtdiscaction.Text
    !actiontype = Left(cmbactiontype.Text, 1)
    !dateserved = IIf(mskdateserved.Text = "__/__/____", Null, mskdateserved.Text)
    !counselingdate = txtcoundate.Text
    !counselingfacilitator = txtcounfacilitator.Text
    !counselinginfo = txtcouninfo.Text
    !Remarks = txtRemarks.Text
        !USERID = operatorID
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

Private Sub mskdateserved_Validate(Cancel As Boolean)
If mskdateserved <> "__/__/____" And IsDate(mskdateserved) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskdateserved.SetFocus
End Sub

Private Sub mskoffensedate_GotFocus()
mskoffensedate.BackColor = &HFFFFC0
End Sub

Private Sub mskoffensedate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskoffensedate_LostFocus()
mskoffensedate.BackColor = vbWhite
End Sub

Private Sub mskoffensedate_Validate(Cancel As Boolean)
If mskoffensedate <> "__/__/____" And IsDate(mskoffensedate) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskoffensedate.SetFocus
End Sub

Private Sub txtcoundate_GotFocus()
txtcoundate.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtcoundate
End Sub

Private Sub txtcoundate_LostFocus()
txtcoundate.BackColor = vbWhite
End Sub

Private Sub txtcounfacilitator_GotFocus()
txtcounfacilitator.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtcounfacilitator
End Sub


Private Sub txtcounfacilitator_LostFocus()
txtcounfacilitator.BackColor = vbWhite
End Sub

Private Sub txtcouninfo_GotFocus()
txtcouninfo.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtcouninfo
End Sub



Private Sub txtcouninfo_LostFocus()
txtcouninfo.BackColor = vbWhite
End Sub

Private Sub txtdiscaction_GotFocus()
txtdiscaction.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdiscaction
End Sub

Private Sub txtdiscaction_LostFocus()
txtdiscaction.BackColor = vbWhite
End Sub

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
Set oCommonFunction.ValidAlpha = txtmovementseq
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

Private Sub txtControlNo_GotFocus()
TxtControlNo.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = TxtControlNo
End Sub


Private Sub txtControlNo_LostFocus()
TxtControlNo.BackColor = vbWhite
End Sub
Private Sub cmbempname_GotFocus()
cmbempname.BackColor = &HFFFFC0
End Sub



Private Sub cmbempname_LostFocus()
cmbempname.BackColor = vbWhite
End Sub


Private Sub txtoffensemade_GotFocus()
txtoffensemade.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtoffensemade
End Sub


Private Sub txtoffensemade_LostFocus()
txtoffensemade.BackColor = vbWhite
End Sub

Private Sub txtoffenseno_GotFocus()
txtcounfacilitator.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtoffenseno
End Sub

Private Sub txtremarks_GotFocus()
txtRemarks.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtRemarks
End Sub


Private Sub txtremarks_LostFocus()
txtRemarks.BackColor = vbWhite
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
    TxtControlNo.Text = IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
End Sub

