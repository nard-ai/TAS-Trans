VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHRTraining 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HR Employee Trainings"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "frmHRTraining.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   7770
   Begin VB.TextBox cmbtrainingduration 
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
      Left            =   1485
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   3375
      Width           =   2205
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3915
      Left            =   45
      TabIndex        =   15
      Top             =   630
      Width           =   7575
      Begin VB.CommandButton cmdfirst 
         Height          =   255
         Left            =   6525
         Picture         =   "frmHRTraining.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Top record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdPrev 
         Height          =   255
         Left            =   6765
         Picture         =   "frmHRTraining.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Previous record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdNext 
         Height          =   255
         Left            =   7005
         Picture         =   "frmHRTraining.frx":081E
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Next record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdlast 
         Height          =   255
         Left            =   7245
         Picture         =   "frmHRTraining.frx":0968
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Left            =   1440
         TabIndex        =   2
         Top             =   1125
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
         Left            =   1440
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   765
         Width           =   2205
      End
      Begin VB.TextBox txttrainingno 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   375
         Width           =   2205
      End
      Begin VB.TextBox txttraining 
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
         Left            =   1440
         TabIndex        =   4
         Text            =   "Text4"
         Top             =   1950
         Width           =   2205
      End
      Begin VB.TextBox txtinstitute 
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
         Left            =   1440
         TabIndex        =   7
         Text            =   "Text6"
         Top             =   3120
         Width           =   2205
      End
      Begin VB.TextBox txtfacilitator 
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
         Left            =   1440
         TabIndex        =   8
         Text            =   "Text7"
         Top             =   3510
         Width           =   2205
      End
      Begin VB.TextBox txttrainingvenue 
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
         Left            =   5520
         TabIndex        =   9
         Text            =   "Text8"
         Top             =   375
         Width           =   1935
      End
      Begin VB.TextBox txttrainingcost 
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
         Left            =   5520
         TabIndex        =   10
         Text            =   "Text9"
         Top             =   765
         Width           =   1935
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
         Height          =   915
         Left            =   5490
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "frmHRTraining.frx":0AB2
         Top             =   2325
         Width           =   1935
      End
      Begin VB.ComboBox cmbtrainingtype 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1560
         Width           =   2205
      End
      Begin VB.TextBox txtsponsor 
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
         Left            =   5520
         TabIndex        =   13
         Text            =   "Text10"
         Top             =   1950
         Width           =   1935
      End
      Begin MSMask.MaskEdBox msktrainingdate 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   2340
         Width           =   2205
         _ExtentX        =   3889
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
      Begin MSMask.MaskEdBox mskstartcontract 
         Height          =   285
         Left            =   5520
         TabIndex        =   11
         Top             =   1150
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSMask.MaskEdBox mskendcontract 
         Height          =   285
         Left            =   5535
         TabIndex        =   12
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
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
         Left            =   120
         TabIndex        =   30
         Top             =   1290
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Training No."
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
         TabIndex        =   29
         Top             =   450
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
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
         TabIndex        =   28
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Training Type"
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
         TabIndex        =   27
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Training"
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
         TabIndex        =   26
         Top             =   2070
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Training Date"
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
         TabIndex        =   25
         Top             =   2460
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Training Duration"
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
         TabIndex        =   24
         Top             =   2895
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Institute"
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
         TabIndex        =   23
         Top             =   3285
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Facilitator"
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
         TabIndex        =   22
         Top             =   3630
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Training Venue"
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
         Left            =   4080
         TabIndex        =   21
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Training Cost"
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
         Left            =   4095
         TabIndex        =   20
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Training"
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
         Left            =   4080
         TabIndex        =   19
         Top             =   1230
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "End Training"
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
         Left            =   4080
         TabIndex        =   18
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label label13 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
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
         Index           =   0
         Left            =   4080
         TabIndex        =   17
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label label13 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Sponsor"
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
         Left            =   4080
         TabIndex        =   16
         Top             =   2025
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   2340
         Picture         =   "frmHRTraining.frx":0AB9
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   2925
         Picture         =   "frmHRTraining.frx":0BBB
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmHRTraining.frx":0D05
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "New"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1170
         Picture         =   "frmHRTraining.frx":1237
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1755
         Picture         =   "frmHRTraining.frx":1769
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   600
         Picture         =   "frmHRTraining.frx":1C9B
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Edit"
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
Attribute VB_Name = "frmHRTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Set oCommonFunction.ValidAlpha = txtaddress1
'If KeyAscii = 13 Then SendKeys "{tab}"

Dim rsout As ADODB.Recordset
Dim id As String
Dim trannum As String
Dim outputfilename As String

Private Sub cmbemploymenttype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbmovementtype_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
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


Private Sub cmbempname_Click()
Call EmpField
End Sub

Private Sub cmbempname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
Call FindMatch(cmbempname, KeyAscii)
End Sub

Private Sub cmbtrainingtype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
Private Sub cmbtrainingduration_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmdAdd_Click()
'On Error GoTo Errh
addflagship = True
editflagship = False

Call Add
Call clearscreen
txttrainingno.Enabled = False
txtempno.Enabled = False
cmbempname.Enabled = True
cmbempname.Text = ""
cmbtrainingtype.Enabled = True
cmbtrainingduration.Enabled = True
txttraining.Enabled = True
msktrainingdate.Enabled = True
txtinstitute.Enabled = True
txtfacilitator.Enabled = True
txttrainingvenue.Enabled = True
txttrainingcost.Enabled = True
mskStartContract.Enabled = True
mskEndContract.Enabled = True
txtremarks.Enabled = True
txtsponsor.Enabled = True
'txthrcom.Enabled = True
'mskscontract.Enabled = True
'mskecontract.Enabled = True
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

Private Sub cmdCancel_Click()

addflagship = False
editflagship = False

Set rsout = New ADODB.Recordset
StrQuery = "Select * from HREmployeesTraining "
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HREmployeesTraining ", cnnHR, adOpenKeyset, adLockOptimistic

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
StrQuery = "Select * from HREmployeesTraining "
Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
'rsout.Open "Select * from HREmployeesTraining ", cnnHR, adOpenKeyset, adLockOptimistic

res = MsgBox("Delete record", vbQuestion + vbYesNo, CompanyName)
If res = vbYes Then
    str = "Delete from HREmployeesTraining where trainingnum=" & txttrainingno.Text
    Set rsout = cnn.Execute(str)
   ' res = MsgBox("Record Deleted", vbCritical, CompanyName)
    StrQuery = "Select * from HREmployeesTraining"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "Select * from HREmployeesTraining", cnnHR, adOpenKeyset, adLockOptimistic
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
txttrainingno.Enabled = False
txtempno.Enabled = False
cmbempname.Enabled = True
cmbtrainingtype.Enabled = True
cmbtrainingduration.Enabled = True

msktrainingdate.Enabled = True
txtinstitute.Enabled = True
txtfacilitator.Enabled = True
txttrainingvenue.Enabled = True
txttrainingcost.Enabled = True
mskStartContract.Enabled = True
mskEndContract.Enabled = True
txtremarks.Enabled = True
txtsponsor.Enabled = True
'txthrcom.Enabled = True
'mskscontract.Enabled = True
'mskecontract.Enabled = True
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
    
   trainingnum = rsout!trainingnum
    Showdata rsout
X:
    Exit Sub
End Sub

Private Sub cmdSave_Click()
If cmbempname.Text = "" Then MsgBox "Select employee name. ", vbInformation: cmbempname.SetFocus: Exit Sub

txttrainingno.Enabled = False
txtempno.Enabled = False
cmbempname.Enabled = False
cmbtrainingtype.Enabled = False
cmbtrainingduration.Enabled = False
txttraining.Enabled = False
msktrainingdate.Enabled = False
txtinstitute.Enabled = False
txtfacilitator.Enabled = False
txttrainingvenue.Enabled = False
txttrainingcost.Enabled = False
mskStartContract.Enabled = False
mskEndContract.Enabled = False
txtremarks.Enabled = False
txtsponsor.Enabled = False
'txthrcom.Enabled = False
'mskscontract.Enabled = False
'mskecontract.Enabled = False
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
txttrainingno.Enabled = False
txtempno.Enabled = False
cmbempname.Enabled = False
cmbtrainingtype.Enabled = False
cmbtrainingduration.Enabled = False
txttraining.Enabled = False
msktrainingdate.Enabled = False
txtinstitute.Enabled = False
txtfacilitator.Enabled = False
txttrainingvenue.Enabled = False
txttrainingcost.Enabled = False
mskStartContract.Enabled = False
mskEndContract.Enabled = False
txtremarks.Enabled = False
txtsponsor.Enabled = False
'txthrcom.Enabled = False
'mskscontract.Enabled = False
'mskecontract.Enabled = False
End Sub

Private Sub DTmovementdate_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Load()
   Me.Top = 0 '(mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    
Set rsemp = New ADODB.Recordset
Call Load_ref_table

addflagship = False
editflagship = False
txttrainingno.Enabled = False
txtempno.Enabled = False
cmbempname.Enabled = False
cmbtrainingtype.Enabled = False
cmbtrainingduration.Enabled = False
txttraining.Enabled = False
msktrainingdate.Enabled = False
txtinstitute.Enabled = False
txtfacilitator.Enabled = False
txttrainingvenue.Enabled = False
txttrainingcost.Enabled = False
mskStartContract.Enabled = False
mskEndContract.Enabled = False
txtremarks.Enabled = False
txtsponsor.Enabled = False
'txthrcom.Enabled = False
'mskscontract.Enabled = False
'mskecontract.Enabled = False
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdrefresh.Enabled = False
Set rsout = New ADODB.Recordset

StrQuery = "Select * From HREmployees order by surname, firstname"
Set rsemp = ReturnRst(StrQuery, cnnstr, mstatic)
'Set rsemp = cnn.Execute("Select * From HREmployees order by surname, firstname")
Do While Not rsemp.EOF
    cmbempname.AddItem rsemp!Surname & "," & rsemp!FirstName
    rsemp.MoveNext
Loop
rsemp.Close
    
'open database
ShowTraining

End Sub

Public Sub ShowTraining(Optional ByVal trainingnum As Integer)
If rsout.State = 1 Then rsout.Close
If trainingnum = 0 Then
    StrQuery = "Select * from HREmployeesTraining"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "select * from HREmployeesTraining", cnnHR, adOpenKeyset, adLockOptimistic
Else
    StrQuery = "Select * from HREmployeesTraining where trainingnum=" & trainingnum & ""
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "select * from HREmployeesTraining where trainingnum=" & trainingnum & "", cnnHR, adOpenKeyset, adLockOptimistic
End If
'************
'check if there is an existing record
'************
If rsout.RecordCount > 0 Then
    Showdata rsout
Else
   Call clearscreen
   cmddelete.Enabled = False
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
Set rstrainingType = cnn.Execute("Select * From HRTrainingType")
Do While Not rstrainingType.EOF
        cmbtrainingtype.AddItem rstrainingType!trainingtype & "-" & rstrainingType!trainingtypedesc
        rstrainingType.MoveNext
    Loop
    rstrainingType.Close

End Sub

Private Sub SaveNewHREmployeeRecord()
Dim res

' For Update 2010
Set rsout = New ADODB.Recordset
'
'strQuery = "Select * from HREmployeesTraining"
'Set rsout = ReturnRst(strQuery, cnnstr, mstatic)
''rsout.Open "Select * from HREmployeesTraining", cnnHR, adOpenKeyset, adLockOptimistic
''***********
''save new record
''***********
'With rsout
'    If Not (.BOF And .EOF) Then
'    .MoveLast
'    End If
'    .AddNew
'    '!trainingnum = IIf(Trim(txttrainingno.Text) = "", Null, txttraining.Text)
'    !EmployeeID = txtempno.Text
'    !EmployeeName = cmbempname.Text
'    !trainingtype = Left(cmbtrainingtype.Text, 1)
'    !training = txttraining.Text
'    !trainingdate = IIf(msktrainingdate.Text = "__/__/____", Null, Format(msktrainingdate.Text, "mm/dd/yyyy"))
'    !trainingduration = cmbtrainingduration.Text
'    !institute = txtinstitute.Text
'    !facilitator = txtfacilitator.Text
'    !trainingvenue = txttrainingvenue.Text
'    !trainingcost = IIf(IsNumeric(txttrainingcost.Text) = True, txttrainingcost.Text, 0)
'    !starttrainingcontract = IIf(mskStartContract.Text = "__/__/____", Null, Format(mskStartContract.Text, "mm/dd/yyyy"))
'    !endtrainingcontract = IIf(mskEndContract.Text = "__/__/____", Null, Format(mskEndContract.Text, "mm/dd/yyyy"))
'    !Remarks = txtRemarks.Text
'    !sponsor = IIf(txtsponsor.Text = "", Null, txtsponsor.Text)
'        !USERID = operatorID
'    .Update
'End With
'Me.txttrainingno.Text = Getmax
End Sub

Function Getmax() As Integer
    Dim rs As ADODB.Recordset
    Dim str As String
    str = "Select max(trainingnum) From HREmployeesTraining"
        Set rs = cnn.Execute(str)
    If Not rs.EOF Then
         Getmax = rs.Fields(0)
    End If
End Function

Private Sub clearscreen()
'***********
'clear textbox
'***********

txttrainingno.Text = "<auto number>"
txtempno.Text = ""
cmbempname.ListIndex = -1
cmbtrainingtype.ListIndex = -1
cmbtrainingduration.Text = ""
txttraining.Text = ""
msktrainingdate.Text = "__/__/____"
txtinstitute.Text = ""
txtfacilitator.Text = ""
txttrainingvenue.Text = ""
txttrainingcost.Text = ""
mskStartContract.Text = "__/__/____"
mskEndContract.Text = "__/__/____"
txtremarks.Text = ""
txtsponsor.Text = ""

End Sub
Private Sub Showdata(ByRef rec As ADODB.Recordset)
Set rs = New ADODB.Recordset
Dim rsout As ADODB.Recordset
'***********
'show record
'***********
If rec.EOF = True Then
Call clearscreen
cmdedit.Enabled = False
cmddelete.Enabled = False

Else
    txttrainingno.Text = IIf(IsNull(rec!trainingnum), "000000", rec!trainingnum)
    txtempno.Text = IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
    'cmbtrainingtype.Text = IIf(IsNull(Rec!trainingtype), "", Rec!trainingtype)
    txttraining.Text = IIf(IsNull(rec!training), "", rec!training)
    msktrainingdate.Text = IIf(IsNull(rec!trainingdate), "__/__/____", Format(rec!trainingdate, "mm/dd/yyyy"))
    txtinstitute.Text = IIf(IsNull(rec!institute), "", rec!institute)
    txtfacilitator.Text = IIf(IsNull(rec!facilitator), "", rec!facilitator)
    txttrainingvenue.Text = IIf(IsNull(rec!trainingvenue), "", rec!trainingvenue)
    txttrainingcost.Text = IIf(IsNull(rec!trainingcost), "", rec!trainingcost)
    mskStartContract.Text = IIf(IsNull(rec!starttrainingcontract), "__/__/____", Format(rec!starttrainingcontract, "mm/dd/yyyy"))
    mskEndContract.Text = IIf(IsNull(rec!endtrainingcontract), "__/__/____", Format(rec!endtrainingcontract, "mm/dd/yyyy"))
    txtremarks.Text = IIf(IsNull(rec!Remarks), "", rec!Remarks)
    txtsponsor.Text = IIf(IsNull(rec!sponsor), "", rec!sponsor)
    cmbtrainingduration.Text = IIf(IsNull(rec!trainingduration), "", rec!trainingduration)
    
    StrQuery = "Select * From HRtrainingType Where trainingType ='" & IIf(IsNull(rec!trainingtype), "", rec!trainingtype)
    Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
    'Set rs = cnn.Execute("Select * From HRtrainingType Where trainingType ='" & IIf(IsNull(rec!trainingtype), "", rec!trainingtype) & "'")
        If Not rs.EOF Then
            cmbtrainingtype.Text = rs.Fields(0) & "-" & rs.Fields(1)
        Else
            cmbtrainingtype.ListIndex = -1
        End If
        rs.Close

    StrQuery = "Select * From HREmployees Where EmployeeID =" & IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID)
    Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
    'Set rs = cnn.Execute("Select * From HREmployees Where EmployeeID =" & IIf(IsNull(rec!EmployeeID), "", rec!EmployeeID) & "")
        If Not rs.EOF Then
            cmbempname.Text = rs.Fields(4) & "," & rs.Fields(2)
        Else
            cmbempname.ListIndex = -1
        End If
        rs.Close

End If
End Sub

Private Sub editsave()
Dim res
'***********
'save edited record
'***********

With rsout
   ' !trainingnum = txttrainingno.Text
    !EmployeeID = txtempno.Text
    !EmployeeName = cmbempname.Text
    !trainingtype = Left(cmbtrainingtype.Text, 1)
    !training = txttraining.Text
    !trainingdate = IIf(msktrainingdate.Text = "__/__/____", Null, Format(msktrainingdate.Text, "mm/dd/yyyy"))
    !trainingduration = cmbtrainingduration.Text
    !institute = txtinstitute.Text
    !facilitator = txtfacilitator.Text
    !trainingvenue = txttrainingvenue.Text
    !trainingcost = IIf(txttrainingcost.Text = "", Null, txttrainingcost.Text)
    !starttrainingcontract = IIf(mskStartContract.Text = "__/__/____", Null, Format(mskStartContract.Text, "mm/dd/yyyy"))
    !endtrainingcontract = IIf(mskEndContract.Text = "__/__/____", Null, Format(mskEndContract.Text, "mm/dd/yyyy"))
    !Remarks = txtremarks.Text
    !sponsor = IIf(txtsponsor.Text = "", Null, txtsponsor.Text)
    '!hrcom = IIf(txthrcom.Text = "", Null, txthrcom.Text)
    '!StartContract = IIf(IsNull(mskscontract.Text), "__/__/____", Format(mskscontract.Text, "mm/dd/yyyy"))
   ' !EndContract = IIf(IsNull(mskecontract.Text), "__/__/____", Format(mskecontract.Text, "mm/dd/yyyy"))
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
Set oCommonFunction.ValidAlpha = txttransacionnum
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

Private Sub txtassigntype_GotFocus()
txtAssignType.BackColor = &HFFFFC0
End Sub

Private Sub txtassigntype_LostFocus()
txtAssignType.BackColor = vbWhite
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

Private Sub mskecontract_Validate(Cancel As Boolean)
If mskecontract <> "__/__/____" And IsDate(mskecontract) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskecontract.SetFocus

End Sub

Private Sub mskEndContract_Validate(Cancel As Boolean)
If mskEndContract <> "__/__/____" And IsDate(mskEndContract) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskEndContract.SetFocus
End Sub

Private Sub mskscontract_Validate(Cancel As Boolean)
If mskscontract <> "__/__/____" And IsDate(mskscontract) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskscontract.SetFocus
End Sub

Private Sub mskStartContract_Validate(Cancel As Boolean)
If mskStartContract <> "__/__/____" And IsDate(mskStartContract) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskStartContract.SetFocus
End Sub

Private Sub msktrainingdate_Validate(Cancel As Boolean)
If msktrainingdate <> "__/__/____" And IsDate(msktrainingdate) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: msktrainingdate.SetFocus
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
txtremarks.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtremarks
End Sub

Private Sub txtremarks_LostFocus()
txtremarks.BackColor = vbWhite
End Sub

Private Sub txttraining_GotFocus()
txttraining.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txttraining
End Sub


Private Sub txttraining_LostFocus()
txttraining.BackColor = vbWhite
End Sub



Private Sub mskscontract_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskscontract_LostFocus()
mskscontract.BackColor = vbWhite
End Sub
Private Sub mskecontract_GotFocus()
mskecontract.BackColor = &HFFFFC0

End Sub

Private Sub mskecontract_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
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
Set oCommonFunction.ValidDouble = txttrainingcost
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
