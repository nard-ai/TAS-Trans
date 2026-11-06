VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmHRBenefits 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HR Benefits Data Entry"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmHRBenefits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2865
   ScaleWidth      =   7695
   Begin VB.TextBox cmdshowbenefitsubform 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Height          =   285
      Left            =   6300
      TabIndex        =   26
      Text            =   "Benefits List"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   45
      TabIndex        =   27
      Top             =   2925
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "   "
      TabPicture(0)   =   "frmHRBenefits.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "listbenefit"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.ListView listbenefit 
         Height          =   1965
         Left            =   180
         TabIndex        =   28
         Top             =   180
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3466
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "BenefitNum"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Benefit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "BenefitType"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "BenefitEligibility"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "BenefitRemarks"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Benefit Date Added"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   90
      TabIndex        =   7
      Top             =   630
      Width           =   7455
      Begin VB.CommandButton cmdSub 
         Caption         =   "..."
         Height          =   285
         Left            =   3330
         TabIndex        =   29
         Top             =   1575
         Width           =   330
      End
      Begin VB.CommandButton cmdfirst 
         Height          =   255
         Left            =   6300
         Picture         =   "frmHRBenefits.frx":05A6
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Top record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPrev 
         Height          =   255
         Left            =   6540
         Picture         =   "frmHRBenefits.frx":06F0
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Previous record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdNext 
         Height          =   255
         Left            =   6780
         Picture         =   "frmHRBenefits.frx":083A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Next record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdlast 
         Height          =   255
         Left            =   7020
         Picture         =   "frmHRBenefits.frx":0984
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Last record"
         Top             =   0
         Width           =   255
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
         Left            =   1290
         TabIndex        =   1
         Top             =   765
         Width           =   2010
      End
      Begin VB.TextBox txtid 
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
         Left            =   1305
         TabIndex        =   0
         Top             =   375
         Width           =   2010
      End
      Begin VB.ComboBox cmbbenefitno 
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
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1575
         Width           =   2010
      End
      Begin VB.ComboBox cmbempname 
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
         Left            =   1290
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1125
         Width           =   2010
      End
      Begin VB.TextBox txtbenremarks 
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
         Left            =   5280
         TabIndex        =   5
         Top             =   765
         Width           =   2055
      End
      Begin VB.TextBox txtbenamount 
         Alignment       =   1  'Right Justify
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
         Left            =   5280
         TabIndex        =   6
         Top             =   1155
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskbenefitacquire 
         Height          =   285
         Left            =   5265
         TabIndex        =   4
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
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
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1245
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   450
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit Acquire Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3480
         TabIndex        =   11
         Top             =   450
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit Remarks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   3480
         TabIndex        =   10
         Top             =   810
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   3480
         TabIndex        =   9
         Top             =   1230
         Width           =   1110
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   2280
      TabIndex        =   8
      Top             =   5025
      Width           =   1815
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   1755
         Picture         =   "frmHRBenefits.frx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   600
         Picture         =   "frmHRBenefits.frx":0BD0
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   2340
         Picture         =   "frmHRBenefits.frx":0D1A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1170
         Picture         =   "frmHRBenefits.frx":124C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmHRBenefits.frx":177E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "New"
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
Attribute VB_Name = "frmHRBenefits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsout As ADODB.Recordset
Dim id As String
Dim trannum As String
Dim outputfilename As String
Dim list As ListItem
Dim rsbenefit As ADODB.Recordset
'Set oCommonFunction.ValidAlpha = txtaddress1
'If KeyAscii = 13 Then SendKeys "{tab}"

Private Sub cmbemploymenttype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbmovementtype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbbenefitno_Click()
Dim intSelectedOption As Integer
Dim strFindMe As String
Dim itmFound As ListItem
      strFindMe = Left(cmbbenefitno.Text, 1)
      intSelectedOption = lvwText
   Set itmFound = listbenefit.FindItem(strFindMe, intSelectedOption, , lvwPartial)
   If itmFound Is Nothing Then
       Exit Sub
   Else
       itmFound.EnsureVisible
       itmFound.SELECTED = True
       listbenefit.SetFocus
End If
End Sub

Private Sub cmbbenefitno_GotFocus()
On Error Resume Next
cmbbenefitno.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = cmbbenefitno
End Sub


Private Sub cmbbenefitno_LostFocus()
cmbbenefitno.BackColor = vbWhite
End Sub

Private Sub cmbemp_GotFocus()
cmbemp.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = cmbemp
End Sub

Private Sub cmbemp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbemp_LostFocus()
cmbemp.BackColor = vbWhite
End Sub

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
    StrQuery = "select * from name where name='" & Search & "' "
End If
Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
'Set rs = cnn.Execute(SelectString)
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
    txtEmpNo.Text = IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
End Sub


Private Sub cmbempname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
Call FindMatch(cmbempname, KeyAscii)
End Sub

Private Sub cmdAdd_Click()
'On Error GoTo Errh
addflagship = True
editflagship = False
AddFlag = True
'AddFlag = True Or addflagship = True

Call Add
Call ClearScreen

txtEmpNo.Enabled = False
cmbempname.Enabled = True
cmbempname.Text = ""
mskbenefitacquire.Enabled = True
cmbbenefitno.Enabled = True
txtbenremarks.Enabled = True
txtbenamount.Enabled = True
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
'If AddFlag = True Then
'        Unload Me
'        Exit Sub
'    End If

    addflagship = False
    editflagship = False
    
Set rsout = New ADODB.Recordset
StrQuery = "Select * from HREmployeeBenefits "
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HREmployeeBenefits ", cnnHR, adOpenKeyset, adLockOptimistic
  
  If rsout.EOF = False Then
        Showdata rsout
        cmddelete.Enabled = True
        cmdedit.Enabled = True
        Else
        cmdNext.Enabled = False
        cmdfirst.Enabled = False
    cmdlast.Enabled = False
    cmdPrev.Enabled = False
       End If
    
    
Call Cancel
End Sub
Private Sub cmdDelete_Click()
Dim res
Dim str As String
Dim id As String

Set rsout = New ADODB.Recordset
StrQuery = "Select * from HREmployeeBenefits "
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HREmployeeBenefits ", cnnHR, adOpenKeyset, adLockOptimistic

res = MsgBox("Are you sure you want to remove the selected benefit?", vbQuestion + vbYesNo, CompanyName)
If res = vbYes Then
    str = "Delete from HREmployeeBenefits where EmployeeID = " & txtID.Text
    Set rsout = cnn.Execute(str)
 '   res = MsgBox("Record Deleted", vbCritical, CompanyName)
    StrQuery = "Select * from HREmployeeBenefits"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "Select * from HREmployeeBenefits", cnnHR, adOpenKeyset, adLockOptimistic
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
txtEmpNo.Enabled = False
cmbempname.Enabled = True
mskbenefitacquire.Enabled = True
cmbbenefitno.Enabled = True
txtbenremarks.Enabled = True
txtbenamount.Enabled = True
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
Call scroll
End Sub
Private Sub cmdlast_Click()
rsout.MoveLast
cmdlast.Enabled = False
cmdNext.Enabled = False
cmdPrev.Enabled = True
cmdfirst.Enabled = True
Showdata rsout
Call scroll
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
Call scroll
X:
    Exit Sub
 


End Sub
Private Sub scroll()
Dim intSelectedOption As Integer
Dim strFindMe As String
Dim itmFound As ListItem
      strFindMe = Left(cmbbenefitno.Text, 1)
      intSelectedOption = lvwText
   Set itmFound = listbenefit.FindItem(strFindMe, intSelectedOption, , lvwPartial)
   If itmFound Is Nothing Then
      MsgBox "No match found"
      Exit Sub
   Else
       itmFound.EnsureVisible
       itmFound.SELECTED = True
       listbenefit.SetFocus
End If
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
    
   id = rsout!id
    Showdata rsout
    Call scroll
X:
    Exit Sub
End Sub

Private Sub cmdSave_Click()

If cmbempname.Text = "" Then MsgBox "Select employee name. ", vbInformation: cmbempname.SetFocus: Exit Sub
If CheckTExtbox = False Then Exit Sub
txtEmpNo.Enabled = False
cmbempname.Enabled = False
mskbenefitacquire.Enabled = False
cmbbenefitno.Enabled = False
txtbenremarks.Enabled = False
txtbenamount.Enabled = False
'Add new record
        If AddFlag = True Or addflagship = True Then

            Call save
           Call SaveNewHREmployeeRecord
           
           AddFlag = False
           addflagship = False
         End If
        
        If EditFlag = True Or editflagship = True Then
            Call Edit
            Call EditSave
        End If
        
     Call save
'rsout.Requery
Set rsout = New ADODB.Recordset
StrQuery = "Select * from HREmployeeBenefits "
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
       
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
txtEmpNo.Enabled = False
cmbempname.Enabled = False
mskbenefitacquire.Enabled = False
cmbbenefitno.Enabled = False
txtbenremarks.Enabled = False
txtbenamount.Enabled = False
End Sub

Private Sub DTmovementdate_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub cmdshowbenefitsubform_Click()
Unload Me
frmHRBenefitSub.Show
frmHRBenefitSub.SetFocus
End Sub

Private Sub cmdSub_Click()
frmHRBenefitSub.Show
frmHRBenefitSub.SetFocus
End Sub

Private Sub Form_Load()

    Me.Top = 0 '(mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    
    Set rsbenefit = New ADODB.Recordset
    Set rsemp = New ADODB.Recordset
    Option1.VALUE = True
    addflagship = False
    editflagship = False
    txtEmpNo.Enabled = False
    cmbempname.Enabled = False
    mskbenefitacquire.Enabled = False
    txtbenremarks.Enabled = False
    txtbenamount.Enabled = False
    cmbbenefitno.Enabled = False
    cmdsave.Enabled = False
    cmdcancel.Enabled = False
    Set rsout = New ADODB.Recordset

    StrQuery = "SELECT * FROM HREmployees WHERE agencycode= " & CompanyCode & " order by surname, firstname"
    Set rsemp = ReturnRst(StrQuery, cnnstr, mstatic)
    'Set rsemp = cnnHR.Execute("Select * From HREmployees where agencycode= " & CompanyCode & " order by surname, firstname")
    Do While Not rsemp.EOF
        cmbempname.AddItem rsemp!Surname & "," & rsemp!FirstName
        rsemp.MoveNext
    Loop
    rsemp.Close
    
    'Set rsbenefit = cnnHR.Execute("Select * From HRbenefits")
    StrQuery = "SELECT * FROM HRbenefits"
    Set rsbenefit = ReturnRst(StrQuery, cnnstr, mstatic)
    Do While Not rsbenefit.EOF
            cmbbenefitno.AddItem rsbenefit!benefitnum & " -" & rsbenefit!benefit
            rsbenefit.MoveNext
    Loop
    rsbenefit.Close
    
    showBenefits
    Call benefitsub
End Sub

Public Sub showBenefits(Optional ByVal id As Integer)

If rsout.State = 1 Then rsout.Close
If id = 0 Then
    StrQuery = "select * from HREmployeeBenefits"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "select * from HREmployeeBenefits", cnnHR, adOpenKeyset, adLockOptimistic
Else
    StrQuery = "select * from HREmployeeBenefits where EmployeeID = " & id & ""
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "select * from HREmployeeBenefits where ID=" & id & "", cnnHR, adOpenKeyset, adLockOptimistic
End If

If rsout.RecordCount > 0 Then
    Showdata rsout
Else
   Call ClearScreen
   cmddelete.Enabled = False
   cmdfirst.Enabled = False
   cmdPrev.Enabled = False
   cmdNext.Enabled = False
   cmdlast.Enabled = False
   cmdedit.Enabled = False
   cmddelete.Enabled = False
End If

End Sub

Private Sub jeff()
Dim rs As ADODB.Recordset
Dim intSelectedOption As Integer
Dim strFindMe As String
Dim itmFound As ListItem
      strFindMe = Left(cmbbenefitno.Text, 1)
      intSelectedOption = lvwText
    '  rs.MoveFirst
   Set itmFound = listbenefit.FindItem(strFindMe, intSelectedOption, , lvwPartial)
   If itmFound Is Nothing Then
     ' MsgBox "No match found"
      Call cmbbenefitno_Click
      Exit Sub
   Else
       itmFound.EnsureVisible
       itmFound.SELECTED = True
       listbenefit.SetFocus
End If
End Sub

Private Sub SaveNewHREmployeeRecord()
Dim res

' For Update 2010
Set rsout = New ADODB.Recordset
'strQuery = "Select * from HREmployeeBenefits"
'Set rsout = ReturnRst(strQuery, cnnstr, mstatic)
''rsout.Open "Select * from HREmployeeBenefits", cnnHR, adOpenKeyset, adLockOptimistic
'
'With rsout
'    If Not (.BOF And .EOF) Then
'    .MoveLast
'    End If
'    .AddNew
'    !EmployeeID = txtempno.Text
'    !EmployeeName = cmbempname.Text
'    !benefitacquiredate = IIf(IsNull(mskbenefitacquire.Text) = True, "__/__/____", Format(mskbenefitacquire.Text, "mm/dd/yyyy"))
'    !benefitnum = RTrim(Left(cmbbenefitno.Text, 2))
'    !benefitremarks = IIf(Me.txtbenremarks.Text = "", "-", txtbenremarks.Text)
'    !benefitamount = IIf(txtbenamount.Text = "", 0, txtbenamount.Text)
'    !USERID = operatorID
'    .Update
'End With
'
'Me.txtID.Text = Getmax

'res = MsgBox("Record Added", vbInformation, CompanyName)
End Sub

Function Getmax() As Integer
    Dim rs As ADODB.Recordset
    Dim str As String
  str = "Select max(id) From HREmployeeBenefits"
      Set rs = cnnHR.Execute(str)
    If Not rs.EOF Then
         Getmax = rs.Fields(0)
    End If
End Function

Public Function CheckTExtbox() As Boolean
If mskbenefitacquire.Text = "__/__/____" Then MsgBox "Enter Acquire date", vbInformation: mskbenefitacquire.SetFocus: CheckTExtbox = False: Exit Function
If cmbbenefitno.Text = "" Or cmbbenefitno.ListIndex = -1 Then MsgBox "Enter benefit", vbInformation: cmbbenefitno.SetFocus: CheckTExtbox = False: Exit Function

CheckTExtbox = True

End Function
Private Sub ClearScreen()
txtID.Text = "<auto number>"
txtEmpNo.Text = ""
cmbempname.ListIndex = -1
mskbenefitacquire.Text = "__/__/____"
cmbbenefitno.ListIndex = -1
txtbenremarks.Text = ""
txtbenamount.Text = ""
End Sub
Private Sub Showdata(ByRef rec As ADODB.Recordset)
Dim rs As ADODB.Recordset
Dim rsout As ADODB.Recordset
If rec.EOF = True Then
Call ClearScreen
cmdedit.Enabled = False
cmddelete.Enabled = False

Else
    txtID.Text = IIf(IsNull(rec!id), "", rec!id)
    txtEmpNo.Text = IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
    mskbenefitacquire.Text = IIf(IsNull(rec!benefitacquiredate), "__/__/____", Format(rec!benefitacquiredate, "mm/dd/yyyy"))
   txtbenremarks.Text = IIf(IsNull(rec!benefitremarks), "", rec!benefitremarks)
   txtbenamount.Text = IIf(IsNull(rec!benefitamount), "", rec!benefitamount)
   ' cmbbenefitno.Text = IIf(IsNull(Rec!benefitnum), "", Rec!benefitnum)
End If

    StrQuery = "SELECT * FROM HREmployees WHERE EmployeeID =" & IIf(IsNull(rec!EmployeeID), 0, rec!EmployeeID) & " and agencycode=" & CompanyCode
    Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
    'Set rs = cnnHR.Execute("Select * From HREmployees Where EmployeeID =" & IIf(IsNull(rec!EmployeeID), 0, rec!EmployeeID) & " and agencycode=" & CompanyCode)
    If Not rs.EOF Then
        cmbempname.Text = rs.Fields(4) & "," & rs.Fields(2)
    Else
        cmbempname.ListIndex = -1
    End If
    rs.Close


Set rs = cnnHR.Execute("Select * From HRbenefits Where benefitnum =" & IIf(IsNull(rec!benefitnum), 0, rec!benefitnum))
        If Not rs.EOF Then
            cmbbenefitno.Text = rs.Fields(0) & " -" & rs.Fields(1)
        Else
            cmbbenefitno.ListIndex = -1
        End If
        rs.Close

'Call jeff
End Sub

Private Sub EditSave()
Dim res
With rsout
    !EmployeeID = txtEmpNo.Text
    !EmployeeName = cmbempname.Text
    !benefitacquiredate = mskbenefitacquire.Text
    !benefitnum = RTrim(Left(cmbbenefitno.Text, 2))
    !benefitremarks = txtbenremarks.Text
    !benefitamount = txtbenamount.Text
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
End Sub

Private Sub txttransactionnum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
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

Private Sub txtzipfilename_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
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

Private Sub txtbenefitno_GotFocus()
txtbenefitno.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtbenefitno
End Sub

Private Sub txtbenefitno_LostFocus()
txtbenefitno.BackColor = vbWhite
End Sub

Private Sub listbenefit_LostFocus()

   Dim i As Integer
   For i = 1 To listbenefit.ListItems.Count
      listbenefit.ListItems.Item(i).SELECTED = False
   Next i

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
On Error Resume Next
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
Private Sub benefitsub()
Dim NN As String
Dim str As String
Dim rs2 As ADODB.Recordset

'listbenefit.ListItems.Clear
StrQuery = "Select * From HRbenefits" ' where benefitnum=" & Me.cmbbenefitno.Text
Set rs2 = ReturnRst(StrQuery, cnnstr, mstatic)
'Set rs2 = cnnHR.Execute(str)
    Do While Not rs2.EOF
    Set list = listbenefit.ListItems.Add(, , rs2!benefitnum)
    list.SubItems(1) = rs2!benefit
   ' list.SubItems(2) = IIf(IsNull(rs2!benefittype), "", rs2!benefittype)
    list.SubItems(3) = IIf(IsNull(rs2!benefiteligibility), "", rs2!benefiteligibility)
    list.SubItems(4) = IIf(IsNull(rs2!benefitremarks), "", rs2!benefitremarks)
    list.SubItems(5) = IIf(IsNull(rs2!benefitdateadded), "", rs2!benefitdateadded)
    rs2.MoveNext
    Loop
    rs2.Close

End Sub

Private Sub mskbenefitacquire_GotFocus()
mskbenefitacquire.BackColor = &HFFFFC0
End Sub

Private Sub mskbenefitacquire_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskbenefitacquire_LostFocus()
mskbenefitacquire.BackColor = vbWhite
End Sub

Private Sub mskbenefitacquire_Validate(Cancel As Boolean)
If mskbenefitacquire <> "__/__/____" And IsDate(mskbenefitacquire) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskbenefitacquire.SetFocus

End Sub

Private Sub txtbenamount_GotFocus()
'On Error Resume Next
txtbenamount.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txtbenamount
End Sub


Private Sub txtbenamount_LostFocus()
txtbenamount.BackColor = vbWhite
End Sub

Private Sub txtbenremarks_GotFocus()
On Error Resume Next
txtbenremarks.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtbenremarks
End Sub

Private Sub txtbenremarks_LostFocus()
txtbenremarks.BackColor = vbWhite
End Sub

