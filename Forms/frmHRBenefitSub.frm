VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHRBenefitSub 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HR Benefits"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "frmHRBenefitSub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2280
   ScaleWidth      =   6645
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   135
      TabIndex        =   14
      Top             =   585
      Width           =   6435
      Begin VB.CommandButton cmdlast 
         Height          =   255
         Left            =   6105
         Picture         =   "frmHRBenefitSub.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Last record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdNext 
         Height          =   255
         Left            =   5865
         Picture         =   "frmHRBenefitSub.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Next record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPrev 
         Height          =   255
         Left            =   5625
         Picture         =   "frmHRBenefitSub.frx":081E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Previous record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdfirst 
         Height          =   255
         Left            =   5385
         Picture         =   "frmHRBenefitSub.frx":0968
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Top record"
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtbenefit 
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
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtbenefitno 
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
         Left            =   1200
         TabIndex        =   0
         Top             =   330
         Width           =   1935
      End
      Begin VB.TextBox txtbenefitel 
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
         Left            =   1170
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtbenefitrem 
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
         Left            =   4365
         TabIndex        =   3
         Top             =   315
         Width           =   1935
      End
      Begin MSMask.MaskEdBox mskbenefit 
         Height          =   285
         Left            =   4365
         TabIndex        =   4
         Top             =   675
         Width           =   1170
         _ExtentX        =   2064
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   " No."
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
         Left            =   90
         TabIndex        =   19
         Top             =   405
         Width           =   285
      End
      Begin VB.Label Label2 
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
         TabIndex        =   18
         Top             =   795
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   " Eligibility"
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
         Height          =   420
         Left            =   135
         TabIndex        =   17
         Top             =   945
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   3465
         TabIndex        =   16
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   " Date Added"
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
         Left            =   3420
         TabIndex        =   15
         Top             =   765
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
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
         Height          =   420
         Left            =   1200
         Picture         =   "frmHRBenefitSub.frx":0AB2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   480
      End
      Begin VB.CommandButton cmdedit 
         Height          =   420
         Left            =   600
         Picture         =   "frmHRBenefitSub.frx":0FE4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   420
         Left            =   1800
         Picture         =   "frmHRBenefitSub.frx":112E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Undo"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   420
         Left            =   0
         Picture         =   "frmHRBenefitSub.frx":1660
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Add"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         Height          =   420
         Left            =   2400
         Picture         =   "frmHRBenefitSub.frx":1B92
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
Attribute VB_Name = "frmHRBenefitSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsout As ADODB.Recordset
Dim id As String
Dim trannum As String
Dim outputfilename As String
Dim rsMoveType As ADODB.Recordset
Dim rsEmpType As ADODB.Recordset
Dim rsjobclass As ADODB.Recordset
Dim rsjobcode As ADODB.Recordset
Dim rstclient As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim bentyp, ben_add, benadd As String

Private Sub cmbemploymenttype_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub cmbmovementtype_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub cmbbenefittype_GotFocus()
'txttransactionnum.BackColor = &HFFFFC0
End Sub

Private Sub cmbbenefittype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmdAdd_Click()
'On Error GoTo Errh
addflagship = True
editflagship = False

Call Add
Call clearscreen

txtbenefitno.Enabled = False
txtbenefit.Enabled = True
'cmbbenefittype.Enabled = True
txtbenefitel.Enabled = True
txtbenefitrem.Enabled = True
mskbenefit.Enabled = True
txtbenefit.SetFocus
'Exit Sub
'Errh:
'    If err.Number = 5 Then
'        Resume Next
'    End If

End Sub
Private Sub Add()
cmdadd.Enabled = False
cmdedit.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
cmdDelete.Enabled = False
'cmdrefresh.Enabled = False
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
    cmdDelete.Enabled = False
    'cmdrefresh.Enabled = False
    cmdfirst.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdlast.Enabled = False
'    cmdExit.Enabled = False
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
Dim benefitnum As String

Set rsout = New ADODB.Recordset
StrQuery = "Select * from HRbenefits "
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HRbenefits ", cnnHR, adOpenKeyset, adLockOptimistic

res = MsgBox("Delete record", vbQuestion + vbYesNo, CompanyName)
If res = vbYes Then
    str = "Delete from HRbenefits where benefitnum=" & txtbenefitno.Text
    Set rsout = cnn.Execute(str)
    'res = MsgBox("Record Deleted", vbCritical, CompanyName)
    StrQuery = "Select * from HRbenefits"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "Select * from HRbenefits", cnnHR, adOpenKeyset, adLockOptimistic
         Showdata rsout
          cmdNext.Enabled = True
          cmdlast.Enabled = True
End If
End Sub
Private Sub cmdEdit_Click()
    editflagship = True
    addflagship = False
    Call Edit
'txtbenefitno.Enabled = True
txtbenefit.Enabled = True
'cmbbenefittype.Enabled = True
txtbenefitel.Enabled = True
txtbenefitrem.Enabled = True
mskbenefit.Enabled = True
txtbenefit.SetFocus
End Sub
Private Sub cmdExit_Click()
'Dim res
'res = MsgBox("Exit QC TransmissionZip Form", vbQuestion + vbYesNo, CompanyName)
  '  If res = vbYes Then
        Unload Me
        frmHRBenefits.Show vbModal
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
     '   MsgBox "End of file encountered", vbInformation, CompanyName
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
      'MsgBox "Beginning of file encountered", vbInformation, "DCPI Infromation System"
    End If
    
   benefitnum = rsout!benefitnum
    Showdata rsout
X:
    Exit Sub
End Sub

Private Sub cmdSave_Click()
txtbenefitno.Enabled = False
txtbenefit.Enabled = False
'cmbbenefittype.Enabled = False
txtbenefitel.Enabled = False
txtbenefitrem.Enabled = False
mskbenefit.Enabled = False
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
'cmdrefresh.Enabled = False
cmdDelete.Enabled = True
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
'cmdrefresh.Enabled = False
cmdDelete.Enabled = True
cmdfirst.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdlast.Enabled = True
'cmdExit.Enabled = True
txtbenefitno.Enabled = False
txtbenefit.Enabled = False
'cmbbenefittype.Enabled = False
txtbenefitel.Enabled = False
txtbenefitrem.Enabled = False
mskbenefit.Enabled = False
End Sub

Private Sub DTmovementdate_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Load()
Set rsbenefitType = New ADODB.Recordset
Set rsEmpType = New ADODB.Recordset
Set rsjobclass = New ADODB.Recordset
Set rsjobcode = New ADODB.Recordset
Set rstclient = New ADODB.Recordset

addflagship = False
editflagship = False
txtbenefitno.Enabled = False
txtbenefit.Enabled = False
'cmbbenefittype.Enabled = False
txtbenefitrem.Enabled = False
txtbenefitel.Enabled = False
mskbenefit.Enabled = False
cmdsave.Enabled = False
cmdcancel.Enabled = False
'cmdrefresh.Enabled = False
Set rsout = New ADODB.Recordset

'populate of combo boxes
'Set rsbenefitType = cnn.Execute("Select * From HRbenefits")
'Do While Not rsbenefitType.EOF
'        cmbbenefittype.AddItem rsbenefitType!benefitnum & "-" & rsbenefitType!benefit
'        rsbenefitType.MoveNext
'Loop
'    rsbenefitType.Close

StrQuery = "Select * from HRbenefits"
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "select * from HRbenefits", cnnHR, adOpenKeyset, adLockOptimistic

If rsout.RecordCount > 0 Then
    Showdata rsout
Else
   Call clearscreen
   cmdfirst.Enabled = False
   cmdPrev.Enabled = False
   cmdNext.Enabled = False
   cmdlast.Enabled = False
   cmdedit.Enabled = False
   cmdDelete.Enabled = False
End If

End Sub
Private Sub SaveNewHREmployeeRecord()

'bentyp = Left(cmbbenefittype.Text, 1)
ben_add = IIf(IsNull(mskbenefit.Text), "__/__/____", Format(mskbenefit.Text, "mm/dd/yyyy"))
benadd = IIf(ben_add = "__/__/____", "", ben_add)
'UN = "A"
    sqlString = "INSERT INTO hrbenefits(benefit,benefiteligibility,benefitremarks,benefitdateadded,userid) vaLues ('" & txtbenefit & "','" & txtbenefitel & "','" & txtbenefitrem & "','" & benadd & "','" & operatorID & "')"
     Call Action_Query(sqlString, cnnstr)
  
rsout.Requery
rsout.MoveLast

End Sub

Private Sub clearscreen()
txtbenefitno.Text = "<auto number>"
txtbenefitno.BackColor = vbWhite
txtbenefit.Text = ""
'cmbbenefittype.ListIndex = -1
txtbenefitel.Text = ""
txtbenefitrem.Text = ""
mskbenefit.Text = "__/__/____"
End Sub
Private Sub Showdata(ByRef rec As ADODB.Recordset)
Dim rs As ADODB.Recordset
Dim rsout As ADODB.Recordset
If rec.EOF = True Then
Call clearscreen
cmdedit.Enabled = False
cmdDelete.Enabled = False

Else
    txtbenefitno.Text = IIf(IsNull(rec!benefitnum), "", rec!benefitnum)
    txtbenefit.Text = IIf(IsNull(rec!benefit), "", rec!benefit)
    txtbenefitel.Text = IIf(IsNull(rec!benefiteligibility), "", rec!benefiteligibility)
    txtbenefitrem.Text = IIf(IsNull(rec!benefitremarks), "", rec!benefitremarks)
    mskbenefit.Text = IIf(IsNull(rec!benefitdateadded), "__/__/____", Format(rec!benefitdateadded, "mm/dd/yyyy"))
End If

'show populate combo
'Set rs = cnn.Execute("Select * From HRbenefitType Where benefitType ='" & IIf(IsNull(rec!benefittype), "", rec!benefittype) & "'")
'        If Not rs.EOF Then
'            cmbbenefittype.Text = rs.Fields(0) & "-" & rs.Fields(1)
'        Else
'            cmbbenefittype.ListIndex = -1
'        End If
'        rs.Close
'

End Sub

Private Sub editsave()

'bentyp = Left(cmbbenefittype.Text, 1)
ben_add = IIf(IsNull(mskbenefit.Text), "__/__/____", Format(mskbenefit.Text, "mm/dd/yyyy"))
benadd = IIf(ben_add = "__/__/____", "", ben_add)
         UN = "A"
         
     sqlString = "Update hrbenefits  set benefit='" & txtbenefit & "',benefiteligibility='" & txtbenefitel & "',benefitremarks='" & txtbenefitrem.Text & "',benefitdateadded='" & benadd & "',userid='" & operatorID & "' where benefitnum =" & txtbenefitno.Text & ""
         
         
         Call Action_Query(sqlString, cnnstr)
  
'cnn.Execute "EXEC UpdHRBenefits " & txtbenefitno & ", '" & txtbenefit & "', '" & _
bentyp & "', '" & txtbenefitel & "', '" & txtbenefitrem & _
"', '" & benadd & "', '" & UN & "'"

'rsout.Requery
'rsout.MoveLast

End Sub


Private Sub Timer1_Timer()
'Label4.Left = Label4.Left + 60
'If Label4.Left > Transmissionzip.Width Then Label4.Left = 0
End Sub

Private Sub txttransactionnum_GotFocus()
txttransactionnum.BackColor = &HFFFFC0
End Sub

Private Sub txttransactionnum_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
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

Private Sub txtzipfilename_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtzipfilename_LostFocus()
txtzipfilename.BackColor = vbWhite
End Sub

Private Sub txtclientcode_GotFocus()
txtClientCode.BackColor = &HFFFFC0
End Sub

Private Sub txtclientcode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtclientcode_LostFocus()
txtClientCode.BackColor = vbWhite
End Sub

Private Sub txtemployeeno_GotFocus()
txtEmployeeNo.BackColor = &HFFFFC0
End Sub

Private Sub txtemployeeno_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtemployeeno_LostFocus()
txtEmployeeNo.BackColor = vbWhite
End Sub

Private Sub txtjobclass_GotFocus()
txtjobclass.BackColor = &HFFFFC0
End Sub

Private Sub txtjobclass_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtjobclass_LostFocus()
txtjobclass.BackColor = vbWhite
End Sub

Private Sub txtjobcode_GotFocus()
txtjobcode.BackColor = &HFFFFC0
End Sub

Private Sub txtjobcode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtjobcode_LostFocus()
txtjobcode.BackColor = vbWhite
End Sub

Private Sub txtmovementremarks_GotFocus()
txtmovementremarks.BackColor = &HFFFFC0
End Sub

Private Sub txtmovementremarks_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtmovementremarks_LostFocus()
txtmovementremarks.BackColor = vbWhite
End Sub

Private Sub txtmovementseq_GotFocus()
txtmovementseq.BackColor = &HFFFFC0
End Sub

Private Sub txtmovementseq_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtmovementseq_LostFocus()
txtmovementseq.BackColor = vbWhite
End Sub

Private Sub txtsalary_GotFocus()
txtsalary.BackColor = &HFFFFC0
End Sub

Private Sub txtsalary_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtsalary_LostFocus()
txtsalary.BackColor = vbWhite

End Sub

Private Sub txtstatus_GotFocus()
txtstatus.BackColor = &HFFFFC0
End Sub



Private Sub txtstatus_LostFocus()
txtstatus.BackColor = vbWhite
End Sub


Private Sub mskbenefit_GotFocus()
 mskbenefit.BackColor = &HFFFFC0
End Sub

Private Sub mskbenefit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskbenefit_LostFocus()
mskbenefit.BackColor = vbWhite
End Sub

Private Sub mskbenefit_Validate(Cancel As Boolean)
If mskbenefit <> "__/__/____" And IsDate(mskbenefit) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskbenefit.SetFocus

End Sub

Private Sub txtbenefit_GotFocus()
txtbenefit.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtbenefit
End Sub

Private Sub txtbenefit_LostFocus()
txtbenefit.BackColor = vbWhite
End Sub

Private Sub txtbenefitel_GotFocus()
txtbenefitel.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtbenefitel
End Sub

Private Sub txtbenefitel_LostFocus()
 txtbenefitel.BackColor = vbWhite
End Sub

Private Sub txtbenefitno_GotFocus()
txtbenefitno.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtbenefitel
End Sub



Private Sub txtbenefitno_LostFocus()
txtbenefitno.BackColor = vbWhite
End Sub

Private Sub txtbenefitrem_GotFocus()
txtbenefitrem.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtbenefitrem
End Sub

Private Sub txtbenefitrem_LostFocus()
 txtbenefitrem.BackColor = vbWhite
End Sub
