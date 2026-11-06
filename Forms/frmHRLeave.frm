VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmHRLeave 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Leave Application"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "frmHRLeave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   7905
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3285
      Left            =   135
      TabIndex        =   0
      Top             =   630
      Width           =   7695
      Begin VB.TextBox cmbAprvdBy 
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
         Left            =   5535
         TabIndex        =   38
         Top             =   450
         Width           =   2055
      End
      Begin VB.TextBox cmbNotedBy 
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
         Left            =   5535
         TabIndex        =   37
         Top             =   855
         Width           =   2055
      End
      Begin VB.CommandButton cmdlast 
         Height          =   255
         Left            =   7335
         Picture         =   "frmHRLeave.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Last record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdNext 
         Height          =   255
         Left            =   7095
         Picture         =   "frmHRLeave.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Next record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPrev 
         Height          =   255
         Left            =   6855
         Picture         =   "frmHRLeave.frx":081E
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Previous record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdfirst 
         Height          =   255
         Left            =   6615
         Picture         =   "frmHRLeave.frx":0968
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Top record"
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtdatehired 
         Alignment       =   2  'Center
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
         TabIndex        =   11
         Top             =   3315
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtpagibig 
         Alignment       =   2  'Center
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
         Left            =   1470
         TabIndex        =   10
         Top             =   3330
         Visible         =   0   'False
         Width           =   2055
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
         Left            =   1485
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1125
         Width           =   2280
      End
      Begin VB.TextBox txtLeaveNum 
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
         Left            =   1485
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   2280
      End
      Begin VB.TextBox txtEmployeeNo 
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
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   765
         Width           =   2280
      End
      Begin VB.ComboBox cmbLeaveType 
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
         Left            =   1500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1545
         Width           =   2280
      End
      Begin VB.TextBox txtReason 
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
         Left            =   1500
         TabIndex        =   7
         Top             =   2805
         Width           =   2280
      End
      Begin VB.TextBox txtLeaveRem 
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
         Height          =   780
         Left            =   5520
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1260
         Width           =   2055
      End
      Begin MSMask.MaskEdBox msktax 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   3300
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   12582912
         Enabled         =   0   'False
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###-###-###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msksss 
         Height          =   285
         Left            =   1590
         TabIndex        =   8
         Top             =   3285
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   12582912
         Enabled         =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-#######-#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskleavedatefrom 
         Height          =   285
         Left            =   1500
         TabIndex        =   5
         Top             =   1980
         Width           =   2280
         _ExtentX        =   4022
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
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskleavedateto 
         Height          =   285
         Left            =   1500
         TabIndex        =   6
         Top             =   2415
         Width           =   2280
         _ExtentX        =   4022
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
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Date Hired"
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
         Index           =   16
         Left            =   285
         TabIndex        =   26
         Top             =   3390
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "PAGIBIG#"
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
         Index           =   6
         Left            =   150
         TabIndex        =   25
         Top             =   3405
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "SSS#"
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
         Left            =   210
         TabIndex        =   24
         Top             =   3360
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "TAX ID#"
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
         Left            =   240
         TabIndex        =   23
         Top             =   3375
         Visible         =   0   'False
         Width           =   615
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
         TabIndex        =   22
         Top             =   1275
         Width           =   1140
      End
      Begin VB.Label lblLeaveNo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Leave No."
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
         Top             =   450
         Width           =   735
      End
      Begin VB.Label lblEmpNum 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
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
         TabIndex        =   20
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblLeaveType 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Leave Type"
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
         TabIndex        =   19
         Top             =   1665
         Width           =   840
      End
      Begin VB.Label lblLeaveDtFr 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Leave Date From"
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
         TabIndex        =   18
         Top             =   2100
         Width           =   1230
      End
      Begin VB.Label lblLeaveDtTo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Leave Date To"
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
         TabIndex        =   17
         Top             =   2535
         Width           =   1050
      End
      Begin VB.Label lblReason 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Reason"
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
         TabIndex        =   16
         Top             =   2925
         Width           =   540
      End
      Begin VB.Label lblAprvdBy 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Approved By"
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
         Left            =   4200
         TabIndex        =   15
         Top             =   525
         Width           =   930
      End
      Begin VB.Label lblLeaveRemarks 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
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
         Left            =   4200
         TabIndex        =   14
         Top             =   1335
         Width           =   615
      End
      Begin VB.Label lblNotedBy 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Noted By"
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
         Left            =   4200
         TabIndex        =   13
         Top             =   915
         Width           =   660
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6240
      Top             =   7320
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   1800
         Picture         =   "frmHRLeave.frx":0AB2
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   600
         Picture         =   "frmHRLeave.frx":0BB4
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmHRLeave.frx":0CFE
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "New"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1215
         Picture         =   "frmHRLeave.frx":1230
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   2385
         Picture         =   "frmHRLeave.frx":1762
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Cancel"
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
Attribute VB_Name = "frmHRLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cn As ADODB.Connection
Dim rsHREmpLeave As ADODB.Recordset
Dim rstclient As ADODB.Recordset
Dim rsLeaveType As ADODB.Recordset
Dim rsUser As ADODB.Recordset
Dim addleave, editleave As Boolean
Dim rsout As ADODB.Recordset
Dim id As String
Dim trannum As String
Dim outputfilename As String
Dim str As String

Private Sub cmbemploymenttype_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub cmbmovementtype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
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
    SelectString = "select * from NAME where name='" & Search & "' "
End If
Set rs = ReturnRst(SelectString, cnnstr, mstatic)
'Set rs = cnnHR.Execute(SelectString)
        If rs.EOF Then
        End If
ShowdataProd rs
End Sub
Private Sub ShowdataProd(ByRef rec As ADODB.Recordset)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    On Error Resume Next
     txtEmployeeNo.Text = IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
     msksss.Text = IIf(IsNull(rec!sss), "__-_______-_", rec!sss)
     msktax.Text = IIf(IsNull(rec!TIN), "___-___-___", rec!TIN)
     txtpagibig.Text = IIf(IsNull(rec!Pagibig), "", rec!Pagibig)
     txtdatehired.Text = IIf(IsNull(rec!datehired), "", rec!datehired)
End Sub


Private Sub btnHRLeaveSave_Click()

End Sub

Private Sub cmbAprvdBy_GotFocus()
cmbAprvdBy.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = cmbAprvdBy
End Sub

Private Sub cmbAprvdBy_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbAprvdBy_LostFocus()
cmbAprvdBy.BackColor = vbWhite
End Sub

Private Sub cmbempname_Click()
Call EmpField
End Sub

Private Sub cmbempname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
Call FindMatch(cmbempname, KeyAscii)
End Sub

Private Sub cmbLeaveType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbNotedBy_GotFocus()
cmbNotedBy.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = cmbNotedBy
End Sub

Private Sub cmbNotedBy_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbNotedBy_LostFocus()
cmbNotedBy.BackColor = vbWhite
End Sub

Private Sub cmdAdd_Click()
'On Error GoTo Errh
addflagship = True
editflagship = False

Call Add
Call ClearScreen
cmbempname.Enabled = True
cmbempname.Text = ""
cmbLeaveType.Enabled = True
mskleavedatefrom.Enabled = True
mskleavedateto.Enabled = True
txtReason.Enabled = True
'txtLeaveBal.Enabled = True
cmbAprvdBy.Enabled = True
cmbNotedBy.Enabled = True
'txtLeaveSrvd.Enabled = True
txtLeaveRem.Enabled = True
cmbempname.SetFocus
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
'    cmdExit.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    
addflagship = False
editflagship = False
    
Set rsout = New ADODB.Recordset
StrQuery = "Select * from HREmployeesleaves "
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HREmployeesleaves ", cnnHR, adOpenKeyset, adLockOptimistic

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
Dim trainingnum As String

Set rsout = New ADODB.Recordset
StrQuery = "Select * from HREmployeesleaves "
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HREmployeesleaves ", cnnHR, adOpenKeyset, adLockOptimistic

res = MsgBox("Delete record", vbQuestion + vbYesNo, CompanyName)
If res = vbYes Then
    str = "Delete from HREmployeesLeaves where leavenum=" & txtLeaveNum.Text
    Set rsout = cnnHR.Execute(str)
   ' res = MsgBox("Record Deleted", vbCritical, CompanyName)
    StrQuery = "Select * from HREmployeesLeaves"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "Select * from HREmployeesLeaves", cnnHR, adOpenKeyset, adLockOptimistic
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
    txtLeaveNum.Enabled = False
cmbempname.Enabled = True
cmbLeaveType.Enabled = True
mskleavedatefrom.Enabled = True
mskleavedateto.Enabled = True
txtReason.Enabled = True
'txtLeaveBal.Enabled = True
cmbAprvdBy.Enabled = True
cmbNotedBy.Enabled = True
'txtLeaveSrvd.Enabled = True
txtLeaveRem.Enabled = True
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
    
   leavenum = rsout!leavenum
    Showdata rsout
X:
    Exit Sub
End Sub

Private Sub cmdSave_Click()
If cmbempname.Text = "" Then MsgBox "Select employee name. ", vbInformation: cmbempname.SetFocus: Exit Sub

cmbempname.Enabled = False
cmbLeaveType.Enabled = False
mskleavedatefrom.Enabled = False
mskleavedateto.Enabled = False
txtReason.Enabled = False
'txtLeaveBal.Enabled = False
cmbAprvdBy.Enabled = False
cmbNotedBy.Enabled = False
'txtLeaveSrvd.Enabled = False
txtLeaveRem.Enabled = False
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
cmdedit.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmddelete.Enabled = True
cmdfirst.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdlast.Enabled = True
'cmdExit.Enabled = True
cmbempname.Enabled = False
cmbLeaveType.Enabled = False
mskleavedatefrom.Enabled = False
mskleavedateto.Enabled = False
txtReason.Enabled = False
'txtLeaveBal.Enabled = False
cmbAprvdBy.Enabled = False
cmbNotedBy.Enabled = False
'txtLeaveSrvd.Enabled = False
txtLeaveRem.Enabled = False
End Sub

Private Sub DTmovementdate_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Initialize()
Set oCommonFunction = New clsCommon
End Sub

Private Sub Form_Load()
 
Me.Top = 0 '(mdiMain.Height - Me.Height) / 8
Me.Left = (mdiMain.Width - Me.Width) / 2
Set rsemp = New ADODB.Recordset
'Call Load_ref_table

addflagship = False
editflagship = False
cmbempname.Enabled = False
cmbLeaveType.Enabled = False
mskleavedatefrom.Enabled = False
mskleavedateto.Enabled = False
txtReason.Enabled = False
'txtLeaveBal.Enabled = False
cmbAprvdBy.Enabled = False
cmbNotedBy.Enabled = False
'txtLeaveSrvd.Enabled = False
txtLeaveRem.Enabled = False
cmdsave.Enabled = False
cmdcancel.Enabled = False
Set rsout = New ADODB.Recordset
Set rsUser = New ADODB.Recordset

'    Set cn = New ADODB.Connection
    'cn.Open "DSN=Sample"

    StrQuery = "SELECT * FROM HREmployees WHERE agencycode= " & CompanyCode & " order by surname, firstname"
    Set rsemp = ReturnRst(StrQuery, cnnstr, mstatic)
    'Set rsemp = cnnHR.Execute("Select * From HREmployees where agencycode= " & CompanyCode & " order by surname, firstname")
    Do While Not rsemp.EOF
        cmbempname.AddItem rsemp!Surname & "," & rsemp!FirstName
        rsemp.MoveNext
    Loop
    rsemp.Close

    StrQuery = "Select * from HRLeaveType"
    Set rsLeaveType = ReturnRst(StrQuery, cnnstr, mstatic)
    
'Set rsLeaveType = cnnHR.Execute("Select * from HRLeaveType")
     Do While Not rsLeaveType.EOF
        cmbLeaveType.AddItem rsLeaveType!leavetype & "-" & rsLeaveType!leavetypedesc
        rsLeaveType.MoveNext
     Loop
     rsLeaveType.Close
    
'    'edit by Bernard
'   ' str = "Select * from User where hr='" & "SU" & "' And clientcode= " & DEP
'
'    'Set rsUser = cnnHR.Execute(str)
'    'Do While Not rsUser.EOF
'     '   cmbAprvdBy.AddItem rsUser!surname & ", " & rsUser!FirstName
'      '  cmbAprvdBy.ItemData(cmbAprvdBy.NewIndex) = rsUser!EmployeeID
'       ' rsUser.MoveNext
'        'Loop
'        'rsUser.Close

'    'filter by manager
'    mgr = "Select * from Userstable where hr='" & "MG" & "'"
'
'    Set rsUser = cnnHR.Execute(mgr)
'    Do While Not rsUser.EOF
'        cmbNotedBy.AddItem rsUser!surname & ", " & rsUser!FirstName
'        cmbNotedBy.ItemData(cmbNotedBy.NewIndex) = rsUser!EmployeeID
'        rsUser.MoveNext
'        Loop
'        rsUser.Close
''end here
'open database
ShowLeaves

End Sub

Public Sub ShowLeaves(Optional ByVal leavenum As Integer)
'open database
If rsout.State = 1 Then rsout.Close
If leavenum = 0 Then
    StrQuery = "select * from HREmployeesleaves"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "select * from HREmployeesleaves", cnnHR, adOpenKeyset, adLockOptimistic
Else
    StrQuery = "select * from HREmployeesleaves where leavenum=" & leavenum & ""
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "select * from HREmployeesleaves where leavenum=" & leavenum & "", cnnHR, adOpenKeyset, adLockOptimistic
End If

'************
'check if there is an existing record
'************
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


Private Sub Load_ref_table()
Set rstrainingType = New ADODB.Recordset

'populate combo box
Set rstrainingType = cnnHR.Execute("Select * From HRTrainingType")
Do While Not rstrainingType.EOF
        cmbtrainingtype.AddItem rstrainingType!trainingtype & "-" & rstrainingType!trainingtypedesc
        rstrainingType.MoveNext
    Loop
    rstrainingType.Close

Set rstrainingType = cnnHR.Execute("Select * From HRTrainingDuration")
Do While Not rstrainingType.EOF
        cmbtrainingduration.AddItem rstrainingType!trainingtype & "-" & rstrainingType!trainingtypedesc
        rstrainingType.MoveNext
    Loop
    rstrainingType.Close

End Sub

Private Sub SaveNewHREmployeeRecord()
Dim res

Set rsout = New ADODB.Recordset
'
'strQuery = "Select * from HREmployeesleaves"
'Set rsout = ReturnRst(strQuery, cnnstr, mstatic)
''rsout.Open "Select * from HREmployeesleaves", cnnHR, adOpenKeyset, adLockOptimistic
''***********
''save new record
''***********
'With rsout
'    If Not (.BOF And .EOF) Then
'    .MoveLast
'    End If
'    .AddNew
'        !EmployeeID = txtEmployeeNo.Text
'        !EmployeeName = cmbempname.Text
''        !LeaveType = Left(cmbLeaveType.Text, 1)
'        !leavedatefr = IIf(mskleavedatefrom.Text = "__/__/____", Null, Format(mskleavedatefrom, "mm/dd/yyyy"))
'        !leavedateto = IIf(mskleavedateto.Text = "__/__/____", Null, Format(mskleavedateto, "mm/dd/yyyy"))
'        !leaversn = txtReason.Text
'       ' !sss = msksss.Text
'       ' !tax = msktax.Text
'       ' !pagibig = txtpagibig.Text
'       ' !datehired = txtdatehired.Text
'        '!leavebal = IIf(IsNull(txtLeaveBal.Text) Or txtLeaveBal.Text = "", 0, txtLeaveBal.Text)
'        '!leaveserved = IIf(IsNull(txtLeaveSrvd.Text) Or txtLeaveSrvd = "", 0, txtLeaveSrvd.Text)
'        !Remarks = txtLeaveRem.Text
'        !LeaveApprovedBy = cmbAprvdBy.Text
'     !NotedBy = cmbNotedBy.Text
'
'
''
''
'    If cmbLeaveType.Text = "" Or Me.cmbLeaveType.ListIndex = -1 Then
'            !leavetype = ""
'        Else
'            !leavetype = Left(cmbLeaveType.Text, 1)
'        End If
'
'    .Update
'End With
'Me.txtLeaveNum.Text = Getmax
''res = MsgBox("Record Added", vbInformation, CompanyName)
End Sub

Function Getmax() As Integer
    Dim rs As ADODB.Recordset
    Dim str As String
  str = "Select max(leavenum) From HREmployeesleaves"
      Set rs = cnnHR.Execute(str)
    If Not rs.EOF Then
         Getmax = rs.Fields(0)
    End If
    
  End Function
  
Private Sub ClearScreen()
'***********
'clear textbox
'***********

txtLeaveNum.Text = "<auto number>"
txtEmployeeNo.Text = ""
cmbempname.ListIndex = -1
cmbLeaveType.ListIndex = -1
mskleavedatefrom.Text = "__/__/____"
mskleavedateto.Text = "__/__/____"
txtReason.Text = ""
'msksss.Text = "__-_______-_"
'msktax.Text = "___-___-___"
'txtpagibig.Text = ""
'txtdatehired.Text = ""
'txtLeaveBal.Text = ""
'cmbAprvdBy.ListIndex = -1
'cmbNotedBy.ListIndex = -1
'txtLeaveSrvd.Text = ""
txtLeaveRem.Text = ""
cmbAprvdBy.Text = ""
cmbNotedBy.Text = ""
End Sub

Private Sub Showdata(ByRef rec As ADODB.Recordset)
Set rs = New ADODB.Recordset
Dim rsout As ADODB.Recordset
If rec.EOF = True Then
Call ClearScreen
cmdedit.Enabled = False
cmddelete.Enabled = False

Else

    txtLeaveNum.Text = IIf(IsNull(rec!leavenum), "00000000", rec!leavenum)
    txtEmployeeNo.Text = IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
    mskleavedatefrom.Text = IIf(IsNull(rec!leavedatefr), "__/__/____", Format(rec!leavedatefr, "mm/dd/yyyy"))
    mskleavedateto.Text = IIf(IsNull(rec!leavedateto), "__/__/____", Format(rec!leavedateto, "mm/dd/yyyy"))
    txtReason.Text = IIf(IsNull(rec!leaversn), "", rec!leaversn)
  Me.cmbAprvdBy.Text = IIf(IsNull(rec!LeaveApprovedBy), "", rec!LeaveApprovedBy)
  Me.cmbNotedBy.Text = IIf(IsNull(rec!NotedBy), "", rec!NotedBy)
  Me.txtLeaveRem.Text = IIf(IsNull(rec!Remarks), "", rec!Remarks)
 End If

    Set rs = cnnHR.Execute("Select * From HRLeaveType Where LeaveType ='" & IIf(IsNull(rec!leavetype), "", rec!leavetype) & "'")
        If Not rs.EOF Then
            cmbLeaveType.Text = rs.Fields(0) & "-" & rs.Fields(1)
        Else
            cmbLeaveType.ListIndex = -1
        End If
        rs.Close

    StrQuery = "SELECT * FROM HREmployees WHERE EmployeeID =" & IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID) & " and agencycode=" & CompanyCode
    Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
    'Set rs = cnn.Execute("Select * From HREmployees Where EmployeeID =" & IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID) & " and agencycode=" & CompanyCode)
        If Not rs.EOF Then
            cmbempname.Text = rs.Fields(4) & "," & rs.Fields(2)
        Else
            cmbempname.ListIndex = -1
        End If
        rs.Close


End Sub

Private Sub EditSave()
Dim res
'***********
'save edited record
'***********

With rsout
        !EmployeeID = txtEmployeeNo.Text
        !EmployeeName = cmbempname.Text
'        !LeaveType = Left(cmbLeaveType.Text, 1)
        !leavedatefr = IIf(mskleavedatefrom.Text = "__/__/____", Null, Format(mskleavedatefrom, "mm/dd/yyyy"))
        !leavedateto = IIf(mskleavedateto.Text = "__/__/____", Null, Format(mskleavedateto, "mm/dd/yyyy"))
        !leaversn = txtReason.Text
        '!sss = msksss.Text
        '!tax = msktax.Text
        '!pagibig = txtpagibig.Text
        '!datehired = txtdatehired.Text
     '   !leavebal = IIf(IsNull(txtLeaveBal.Text) Or txtLeaveBal.Text = "", 0, txtLeaveBal.Text)
     '   !leaveserved = IIf(IsNull(txtLeaveSrvd.Text) Or txtLeaveSrvd.Text = "", 0, txtLeaveSrvd.Text)
        !Remarks = txtLeaveRem.Text
     !LeaveApprovedBy = cmbAprvdBy.Text
    !NotedBy = cmbNotedBy.Text
    
 '   If cmbAprvdBy.Text = "" Or Me.cmbAprvdBy.ListIndex = -1 Then
  '          !leaveapprovedby = ""
   '     Else
    '        !leaveapprovedby = Me.cmbAprvdBy.ItemData(Me.cmbAprvdBy.ListIndex)
     '   End If

    
    'If cmbNotedBy.Text = "" Or Me.cmbNotedBy.ListIndex = -1 Then
     '       !notedby = ""
      '  Else
       '     !notedby = Me.cmbNotedBy.ItemData(Me.cmbNotedBy.ListIndex)
        'End If

    
    If cmbLeaveType.Text = "" Or Me.cmbLeaveType.ListIndex = -1 Then
            !leavetype = ""
        Else
            !leavetype = Left(cmbLeaveType.Text, 1)
        End If
    
    .Update
End With
'res = MsgBox("Record Updated", vbInformation, CompanyName)
End Sub


Private Sub mskleavedatefrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub mskleavedatefrom_Validate(Cancel As Boolean)
If mskleavedatefrom <> "__/__/____" And IsDate(mskleavedatefrom) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskleavedatefrom.SetFocus

End Sub

Private Sub mskleavedateto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub mskleavedateto_Validate(Cancel As Boolean)
If mskleavedateto <> "__/__/____" And IsDate(mskleavedateto) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskleavedateto.SetFocus
End Sub

Private Sub msksss_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub msktax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

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

Private Sub txtAssignType_GotFocus()
txtAssignType.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtAssignType
End Sub

Private Sub txtAssignType_LostFocus()
txtAssignType.BackColor = vbWhite
End Sub

Private Sub txtdatehired_KeyPress(KeyAscii As Integer)
Set oCommonFunction.ValidAlpha = txtdatehired
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
Set oCommonFunction.ValidAlpha = txtmovemetremarks
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

Private Sub txtfacilitator_GotFocus()
txtfacilitator.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtfacilitator
End Sub



Private Sub txtfacilitator_LostFocus()
txtfacilitator.BackColor = vbWhite
End Sub

Private Sub txtinstitute_GotFocus()
txtinstitute.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtinstitute
End Sub



Private Sub txtinstitute_LostFocus()
txtinstitute.BackColor = vbWhite
End Sub

Private Sub txtremarks_GotFocus()
txtRemarks.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtRemarks
End Sub

Private Sub txtremarks_LostFocus()
txtRemarks.BackColor = vbWhite
End Sub

Private Sub txttraining_GotFocus()
txttraining.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txttraining
End Sub



Private Sub txttraining_LostFocus()
txttraining.BackColor = vbWhite
End Sub



Private Sub txthrcom_GotFocus()
txthrcom.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txthrcom
End Sub


Private Sub txthrcom_LostFocus()
txthrcom.BackColor = vbWhite
End Sub
Private Sub mskscontract_GotFocus()
mskscontract.BackColor = &HFFFFC0

End Sub



Private Sub mskscontract_LostFocus()
mskscontract.BackColor = vbWhite
End Sub
Private Sub mskecontract_GotFocus()
mskecontract.BackColor = &HFFFFC0

End Sub



Private Sub mskecontract_LostFocus()
mskecontract.BackColor = vbWhite
End Sub

Private Sub txtsponsor_GotFocus()
txtsponsor.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtsponsor
End Sub



Private Sub txtsponsor_LostFocus()
txtsponsor.BackColor = vbWhite
End Sub
Private Sub mskEndContract_GotFocus()
mskEndContract.BackColor = &HFFFFC0
End Sub

Private Sub mskEndContract_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskEndContract_LostFocus()
mskEndContract.BackColor = vbWhite
End Sub

Private Sub mskStartContract_GotFocus()
mskStartContract.BackColor = &HFFFFC0
End Sub

Private Sub mskStartContract_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskStartContract_LostFocus()
mskStartContract.BackColor = vbWhite
End Sub
Private Sub msktrainingdate_GotFocus()
msktrainingdate.BackColor = &HFFFFC0

End Sub

Private Sub msktrainingdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub msktrainingdate_LostFocus()
msktrainingdate.BackColor = vbWhite
End Sub

Private Sub txttrainingcost_GotFocus()
txttrainingcost.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txttrainingcost
End Sub


Private Sub txttrainingcost_LostFocus()
txttrainingcost.BackColor = vbWhite
End Sub

Private Sub txttrainingduration_GotFocus()
txttrainingduration.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txttrainingduration
End Sub
Private Sub txttrainingduration_LostFocus()
txttrainingduration.BackColor = vbWhite
End Sub

Private Sub txttrainingtype_GotFocus()
txttrainingtype.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txttrainingtype
End Sub

Private Sub txttrainingtype_LostFocus()
txttraining.BackColor = vbWhite
End Sub

Private Sub txttrainingno_GotFocus()
txttrainingno.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txttrainingno
End Sub


Private Sub txttrainingno_LostFocus()
txttrainingno.BackColor = vbWhite
End Sub

Private Sub txttrainingvenue_GotFocus()
txttrainingvenue.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txttrainingvenue
End Sub



Private Sub txttrainingvenue_LostFocus()
txttrainingvenue.BackColor = vbWhite
End Sub

Private Sub txtLeaveBal_GotFocus()
txtLeaveBal.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtLeaveBal
End Sub

Private Sub txtLeaveBal_LostFocus()
txtLeaveBal.BackColor = vbWhite
End Sub

Private Sub txtLeaveBal_Validate(Cancel As Boolean)
If txtLeaveBal <> "" And IsNumeric(txtLeaveBal) = False Then Cancel = True: MsgBox "Leave balance must be numeric.", vbInformation: txtLeaveBal.SetFocus

End Sub

Private Sub txtLeaveNum_GotFocus()
txtLeaveNum.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtLeaveNum
End Sub
'
Private Sub txtLeaveRem_GotFocus()
txtLeaveRem.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtLeaveRem
End Sub

Private Sub txtLeaveRem_LostFocus()
txtLeaveRem.BackColor = vbWhite
End Sub


Private Sub txtpagibig_Change()
txtpagibig.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtpagibig
End Sub

Private Sub txtReason_GotFocus()
txtReason.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtReason
End Sub

Private Sub txtReason_LostFocus()
txtReason.BackColor = vbWhite
End Sub

Public Sub CommandPass(ByVal strCommand As String)

Select Case strCommand
    Case "Close"
        Unload Me
End Select

End Sub

