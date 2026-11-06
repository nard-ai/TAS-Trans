VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHRAttendance 
   BackColor       =   &H80000009&
   Caption         =   "Attendance"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3675
   ScaleWidth      =   7275
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   7095
      Begin VB.CommandButton btnHRAtndcExit 
         Height          =   585
         Left            =   6360
         Picture         =   "frmHRAttendance.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton btnHRAtndcLast 
         Height          =   585
         Left            =   5400
         Picture         =   "frmHRAttendance.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Last record"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton btnHRAtndcNext 
         Height          =   585
         Left            =   4800
         Picture         =   "frmHRAttendance.frx":067C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Next record"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton btnHRAtndcPrev 
         Height          =   585
         Left            =   4200
         Picture         =   "frmHRAttendance.frx":0776
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Previous record"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton btnHRAtndcFirst 
         Height          =   585
         Left            =   3600
         Picture         =   "frmHRAttendance.frx":0878
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "First record"
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   2655
      Left            =   135
      TabIndex        =   0
      Top             =   585
      Width           =   7095
      Begin VB.TextBox txtEmployeeNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtDay 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtLogMode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtAtndcRem 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtShiftCd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtChronoID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtUnitID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   1680
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPAtndcDt 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         CheckBox        =   -1  'True
         Format          =   52232193
         CurrentDate     =   37528
         MinDate         =   29221
      End
      Begin MSComCtl2.DTPicker DTPAtndcTime 
         Height          =   375
         Left            =   4920
         TabIndex        =   9
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         CheckBox        =   -1  'True
         Format          =   52232194
         UpDown          =   -1  'True
         CurrentDate     =   9.33564814814815E-02
      End
      Begin VB.Label lblEmpNum 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Employee No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label lblShftCd 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Shift Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblChronoID 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Chronolog ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblUnitID 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Unit ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   510
      End
      Begin VB.Label lblReason 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   360
      End
      Begin VB.Label lblDay 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   3735
         TabIndex        =   13
         Top             =   360
         Width           =   285
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   3720
         TabIndex        =   12
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblLogMd 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Log Mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   3720
         TabIndex        =   11
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label lblLeaveRemarks 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   3720
         TabIndex        =   10
         Top             =   1680
         Width           =   750
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   794
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton btnHRAtndcDelete 
         Height          =   495
         Left            =   1800
         Picture         =   "frmHRAttendance.frx":0DAA
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton btnHRAtndcEdit 
         Height          =   495
         Left            =   600
         Picture         =   "frmHRAttendance.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton btnHRAtndcNew 
         Height          =   495
         Left            =   0
         Picture         =   "frmHRAttendance.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "New"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton btnHRAtndcSave 
         Height          =   495
         Left            =   1200
         Picture         =   "frmHRAttendance.frx":1910
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton btnHRAtndcCancel 
         Height          =   495
         Left            =   2400
         Picture         =   "frmHRAttendance.frx":1E42
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   720
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
Attribute VB_Name = "frmHRAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsHRAttendance As ADODB.Recordset
Dim rsHREmployees As ADODB.Recordset
Dim addatndc, editatndc As Boolean
Dim dept As String

Private Sub btnHRAtndcCancel_Click()
    If addatndc = True Then  'this is to refresh the fields
        clearscreen
    End If
    addatndc = False
    SaveCancelClick
    btnHRAtndcFirst.SetFocus
End Sub

Private Sub btnHRAtndcDelete_Click()
 Dim ans As Integer
    
    ans = MsgBox("Are you sure you to delete this record?", vbQuestion + vbYesNo)
    
    If ans = vbYes Then
        rsHRAttendance.Delete
    End If
    rsHRAttendance.MoveFirst
    Showdata rsHRAttendance
    
End Sub

Private Sub btnHRAtndcEdit_Click()
    editatndc = True
    addatndc = False
    AddEditClick
    txtemployeeno.SetFocus
End Sub

Private Sub btnHRAtndcExit_Click()
    Unload Me
End Sub

Private Sub btnHRAtndcFirst_Click()
    rsHRAttendance.MoveFirst
    Showdata rsHRAttendance
End Sub

Private Sub btnHRAtndcLast_Click()
    rsHRAttendance.MoveLast
    Showdata rsHRAttendance
End Sub

Private Sub btnHRAtndcNew_Click()
    addatndc = True
    editatndc = False
    clearscreen
    AddEditClick
    txtemployeeno.SetFocus
   
End Sub

Private Sub btnHRAtndcNext_Click()
    rsHRAttendance.MoveNext
    
    If rsHRAttendance.EOF Then
        rsHRAttendance.MoveLast
        MsgBox "End of file encountered", vbInformation, "Search Asia"
    End If

    Showdata rsHRAttendance
End Sub

Private Sub btnHRAtndcPrev_Click()
   rsHRAttendance.MovePrevious
    
    If rsHRAttendance.BOF Then
        rsHRAttendance.MoveFirst
        MsgBox "Beginning of file encountered", vbInformation, "Search Asia"
    End If

    Showdata rsHRAttendance
End Sub

Private Sub btnHRAtndcSave_Click()


 'If add new record
    If addatndc = True Then
        LoadTable
        AddSave
        addatndc = False
        
    
    'If edit record
    ElseIf editatndc = True Then
        editsave
        editatndc = False
        
    End If
    
    SaveCancelClick
    
End Sub

Private Sub Form_Activate()
    txtemployeeno.SetFocus
End Sub

Private Sub Form_Load()
Dim str As String
    'When clicked particular employee no in listview and choose edit button
    editatndc = True
    addatndc = False
    Set rsHRAttendance = New ADODB.Recordset
    'str = "Select * from HREmployeesAttendance where employeenum= & m_empno and & attendancetime=#" & atndctym & "#"""
    With rsHRAttendance
        .Open "Select * from HREmployeesAttendance ", cnDCPI, adOpenKeyset, adLockOptimistic
        .MoveFirst
       '.Find "employeenum = & m_empno and & attendancetime=#" & atndctym & "#"
         If Not .EOF Then
            AddEditClick
            'txtEmployeeNo.Text = m_empno
            txtemployeeno.Text = IIf(IsNull(!employeenum), 0, !employeenum)
            txtShiftCd.Text = IIf(IsNull(!shiftcode), 0, !shiftcode)
            txtChronoID.Text = IIf(IsNull(!chronoid), 0, !chronoid)
            txtUnitID.Text = IIf(IsNull(!unitid), "", !unitid)
            DTPAtndcDt.Value = IIf(IsNull(!attendancedate), "", !attendancedate)
            txtDay.Text = IIf(IsNull(!Day), 0, !Day)
            DTPAtndcTime.Value = IIf(IsNull(!attendancetime), "", !attendancetime)
            txtLogMode.Text = IIf(IsNull(!LogMode), "", !LogMode)
            txtAtndcRem.Text = IIf(IsNull(!remarks), "", !remarks)
            dept = IIf(IsNull(!deptCode), 0, !deptCode)
        End If
    End With
  
    
'            If rsHRAttendance.BOF And rsHRAttendance.EOF Then
'                addatndc = True
'                btnHRAtndcFirst.Enabled = False
'                btnHRAtndcNext.Enabled = False
'                btnHRAtndcPrev.Enabled = False
'                btnHRAtndcLast.Enabled = False
'                btnHRAtndcNew.Enabled = False
'                btnHRAtndcEdit.Enabled = False
'                btnHRAtndcCancel.Enabled = False
'                btnHRAtndcDelete.Enabled = False
'            Else
'                btnHRAtndcFirst.Enabled = True
'                btnHRAtndcNext.Enabled = True
'                btnHRAtndcPrev.Enabled = True
'                btnHRAtndcLast.Enabled = True
'                btnHRAtndcNew.Enabled = True
'                btnHRAtndcEdit.Enabled = True
'                btnHRAtndcCancel.Enabled = True
'                btnHRAtndcDelete.Enabled = True
'            End If
'    clearscreen
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsHRAttendance.State = 1 Then
        rsHRAttendance.Close
    End If
    
    addatndc = False
    editatndc = False
End Sub

Private Sub Timer1_Timer()
Me.Caption = "Attendance    " & Now
End Sub

Private Sub clearscreen()
  
    txtemployeeno.Text = ""
    txtShiftCd.Text = 0
    txtChronoID.Text = 0
    txtUnitID.Text = ""
    DTPAtndcDt.Value = ""
    txtDay.Text = 0
    DTPAtndcTime.Value = Format$(Now, "hh:mm:ss")
    txtLogMode.Text = ""
    txtAtndcRem.Text = ""
    dept = 0
     
End Sub

Private Sub Showdata(ByRef Rec As ADODB.Recordset)
Dim rs As ADODB.Recordset

    txtemployeeno.Text = IIf(IsNull(Rec!employeenum), "", Rec!employeenum)
    txtShiftCd.Text = IIf(IsNull(Rec!shiftcode), 0, Rec!shiftcode)
    txtChronoID.Text = IIf(IsNull(Rec!chronoid), 0, Rec!chronoid)
    txtUnitID.Text = IIf(IsNull(Rec!unitid), "", Rec!unitid)
    DTPAtndcDt.Value = IIf(IsNull(Rec!attendancedate), "", Rec!attendancedate)
    txtDay.Text = IIf(IsNull(Rec!Day), 0, Rec!Day)
    DTPAtndcTime.Value = IIf(IsNull(Rec!attendancetime), "", Rec!attendancetime)
    txtLogMode.Text = IIf(IsNull(Rec!LogMode), "", Rec!LogMode)
    txtAtndcRem.Text = IIf(IsNull(Rec!remarks), "", Rec!remarks)
    dept = IIf(IsNull(Rec!deptCode), 0, Rec!deptCode)
    
    Set rs = Nothing
End Sub

Sub UnlockCtl()
    'To unlock all textbox
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Locked = False
        End If
    Next

End Sub
Sub LockCtl()
'To lock all textbox
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Locked = True
        End If
    Next
End Sub

Private Sub AddEditClick()
 'Enabling and disabling command buttons when adding
        
    btnHRAtndcFirst.Enabled = False
    btnHRAtndcNext.Enabled = False
    btnHRAtndcPrev.Enabled = False
    btnHRAtndcLast.Enabled = False
    btnHRAtndcNew.Enabled = False
    btnHRAtndcEdit.Enabled = False
    btnHRAtndcSave.Enabled = True
    btnHRAtndcCancel.Enabled = True
    btnHRAtndcDelete.Enabled = False
    btnHRAtndcExit.Enabled = True

    UnlockCtl
    
End Sub

Private Sub AddSave()
'Save new record
    
    
    With rsHRAttendance
        If Not (.BOF And .EOF) Then
        .MoveLast
        End If
        .AddNew
        !employeenum = txtemployeeno.Text
        !shiftcode = txtShiftCd.Text
        !chronoid = txtChronoID.Text
        !unitid = txtUnitID.Text
        !attendancedate = DTPAtndcDt.Value
        !attendancetime = DTPAtndcTime.Value
        !LogMode = txtLogMode.Text
        !remarks = txtAtndcRem.Text
        !UserID = UN
        !deptCode = dept
        .Update
    End With
    
    

End Sub

Public Sub editsave()
'Save existing record
    
    With rsHRAttendance
        
        !employeenum = txtemployeeno.Text
        !shiftcode = txtShiftCd.Text
        !chronoid = txtChronoID.Text
        !unitid = txtUnitID.Text
        !attendancedate = DTPAtndcDt.Value
        !Day = txtDay.Text
        !attendancetime = DTPAtndcTime.Value
        !LogMode = txtLogMode.Text
        !remarks = txtAtndcRem.Text
        !UserID = UN
        !deptCode = dept
        .Update
    End With
    
End Sub
    
    
Private Sub SaveCancelClick()
   
   'Enabling and disabling command buttons when save
   
    btnHRAtndcFirst.Enabled = True
    btnHRAtndcNext.Enabled = True
    btnHRAtndcPrev.Enabled = True
    btnHRAtndcLast.Enabled = True
    btnHRAtndcNew.Enabled = True
    btnHRAtndcEdit.Enabled = True
    btnHRAtndcSave.Enabled = False
    btnHRAtndcCancel.Enabled = False
    btnHRAtndcDelete.Enabled = False
    btnHRAtndcExit.Enabled = True
   
   LockCtl
   
End Sub




Private Sub txtAtndcRem_GotFocus()
    txtAtndcRem.BackColor = &H8000000B
End Sub

Private Sub txtAtndcRem_LostFocus()
    txtAtndcRem.BackColor = &H80000009
End Sub

Private Sub txtChronoID_GotFocus()
txtChronoID.BackColor = &H8000000B
    
    If Len(txtChronoID.Text) > 0 Then
        txtChronoID.SelStart = 0
        txtChronoID.SelLength = Len(txtShiftCd)
    End If

End Sub

Private Sub txtDay_GotFocus()
    txtDay.BackColor = &H8000000B
     If Len(txtDay.Text) > 0 Then
        txtDay.SelStart = 0
        txtDay.SelLength = Len(txtDay)
    End If
End Sub

Private Sub txtDay_LostFocus()
    txtDay.BackColor = &H80000009
End Sub



Private Sub txtDay_Validate(Cancel As Boolean)
    If txtDay.Text < 1 Or txtDay.Text > 7 Then
        MsgBox "Day ranged from 1 to 7 only", vbInformation
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

Private Sub txtemployeeno_GotFocus()
    txtemployeeno.BackColor = &H8000000B
End Sub

Private Sub txtemployeeno_LostFocus()
  txtemployeeno.BackColor = &H80000009
   
End Sub

Private Sub txtEmployeeNo_Validate(Cancel As Boolean)
'If employee no is empty
    If txtemployeeno.Text = "" Then
        MsgBox "Pls input employee no", vbInformation
        Cancel = True
        txtemployeeno.SetFocus
    Else
        Cancel = False
    End If
End Sub

Private Sub txtLogMode_GotFocus()
    txtLogMode.BackColor = &H8000000B
End Sub

Private Sub txtLogMode_KeyPress(KeyAscii As Integer)
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtLogMode_LostFocus()
    txtLogMode.BackColor = &H80000009
End Sub

Private Sub txtLogMode_Validate(Cancel As Boolean)
    If Not (UCase(txtLogMode) = "I" Or UCase(txtLogMode) = "O") Then
        MsgBox "Log mode is either I or O only", vbInformation
        Cancel = True
    Else
        Cancel = False
    End If

End Sub

Private Sub txtShiftCd_GotFocus()
txtShiftCd.BackColor = &H8000000B
    
    If Len(txtShiftCd.Text) > 0 Then
        txtShiftCd.SelStart = 0
        txtShiftCd.SelLength = Len(txtShiftCd)
    End If
End Sub

Private Sub txtShiftCd_LostFocus()
   txtShiftCd.BackColor = &H80000009
End Sub

Private Sub txtShiftCd_Validate(Cancel As Boolean)
 If Not IsNumeric(txtShiftCd.Text) Then
        MsgBox "Invalid number", vbExclamation, "Search Asia"
        Cancel = True
 Else
        Cancel = False
        If txtShiftCd.Text < 1 Or txtShiftCd.Text > 3 Then
            MsgBox "Shift code ranged from 1 to 3 only", vbInformation
        Cancel = True
            Else
            Cancel = False
        End If
    End If
End Sub


Private Sub txtUnitID_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub LoadTable()
'this is getting dept field in HREmployees to save in HREmployeesAttendance
 
    Set rsHREmployees = cnDCPI.Execute("Select employeenum,deptcode from HREmployees")
    
    Do While Not rsHREmployees.EOF
        txtemployeeno.Text = rsHREmployees!employeenum
        dept = rsHREmployees!deptCode
        rsHREmployees.MoveNext
        Loop
        rsHREmployees.Close
    
End Sub
