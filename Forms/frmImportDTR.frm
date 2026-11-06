VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportDTR 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import DTR Records"
   ClientHeight    =   6075
   ClientLeft      =   2745
   ClientTop       =   3105
   ClientWidth     =   11205
   Icon            =   "frmImportDTR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11205
   Begin VB.CommandButton cmdImportDTR 
      Appearance      =   0  'Flat
      Caption         =   "Import New &DTR Records from File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      ToolTipText     =   "Import DTR Records from File"
      Top             =   5160
      Width           =   2865
   End
   Begin VB.CheckBox chkNoFilter 
      BackColor       =   &H8000000E&
      Caption         =   "      All Employees"
      Height          =   375
      Left            =   7680
      TabIndex        =   24
      Top             =   1200
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkInvalidID 
      Caption         =   "&Invalid DTR ID"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5640
      Width           =   1425
   End
   Begin VB.CommandButton cmdPrintInvalidID 
      Appearance      =   0  'Flat
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   22
      ToolTipText     =   "Print Preview"
      Top             =   5640
      Width           =   1425
   End
   Begin VB.TextBox txtClientCode 
      Height          =   285
      Left            =   1920
      TabIndex        =   20
      Text            =   "txtclientcode"
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtEmpNum 
      Height          =   285
      Left            =   1200
      TabIndex        =   19
      Text            =   "txtclientcode"
      Top             =   600
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5430
      Left            =   3120
      TabIndex        =   8
      Top             =   525
      Width           =   7980
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   360
         TabIndex        =   37
         Top             =   1680
         Visible         =   0   'False
         Width           =   4815
         Begin VB.Label lblDisplay 
            Alignment       =   2  'Center
            Caption         =   "Label4"
            Height          =   375
            Left            =   480
            TabIndex        =   38
            Top             =   480
            Width           =   3855
         End
      End
      Begin VB.TextBox txtClient 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Text            =   "Client Name"
         Top             =   345
         Width           =   3015
      End
      Begin VB.TextBox txtClient 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Index           =   0
         Left            =   3240
         TabIndex        =   31
         Text            =   "Client Name"
         Top             =   345
         Width           =   1095
      End
      Begin VB.TextBox txtClient 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Index           =   8
         Left            =   5700
         MultiLine       =   -1  'True
         TabIndex        =   30
         Text            =   "frmImportDTR.frx":058A
         Top             =   345
         Width           =   2115
      End
      Begin VB.TextBox txtClient 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Index           =   9
         Left            =   4380
         MultiLine       =   -1  'True
         TabIndex        =   29
         Text            =   "frmImportDTR.frx":0596
         Top             =   345
         Width           =   1275
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   405
         Left            =   5760
         Picture         =   "frmImportDTR.frx":05A2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Refresh"
         Top             =   680
         Width           =   465
      End
      Begin MSComctlLib.ListView lstDTRDisplay 
         Height          =   4290
         Left            =   15
         TabIndex        =   9
         Top             =   1125
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   7567
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Machine"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DTR ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DTR Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Time In"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Time Out"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Work Hrs"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Import Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Remarks"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   0
         TabIndex        =   17
         Top             =   3720
         Visible         =   0   'False
         Width           =   7500
         Begin VB.Label Label3 
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
            Left            =   6480
            TabIndex        =   18
            Top             =   360
            Width           =   840
         End
      End
      Begin MSComCtl2.DTPicker DTRStart 
         Height          =   375
         Left            =   960
         TabIndex        =   26
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60096513
         CurrentDate     =   40262
      End
      Begin MSComCtl2.DTPicker DTREnd 
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60096513
         CurrentDate     =   40262
      End
      Begin VB.ComboBox cmbYear 
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
         ItemData        =   "frmImportDTR.frx":06EC
         Left            =   6345
         List            =   "frmImportDTR.frx":06EE
         Sorted          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "cmbYear"
         Top             =   1485
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.ComboBox cmbMonth 
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
         ItemData        =   "frmImportDTR.frx":06F0
         Left            =   4800
         List            =   "frmImportDTR.frx":06F2
         Style           =   2  'Dropdown List
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1485
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name:"
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
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee No:"
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
         Height          =   255
         Index           =   13
         Left            =   3240
         TabIndex        =   34
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "DTR ID:"
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
         Height          =   255
         Index           =   6
         Left            =   4380
         TabIndex        =   33
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Schedule:"
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
         Height          =   255
         Index           =   37
         Left            =   5700
         TabIndex        =   32
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "End DTR:"
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
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   28
         Top             =   795
         Width           =   795
      End
      Begin VB.Label lblRecord 
         BackStyle       =   0  'Transparent
         Caption         =   "New DTR:"
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
         Left            =   6360
         TabIndex        =   25
         Top             =   720
         Width           =   960
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Start DTR:"
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
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   15
         Top             =   795
         Width           =   795
      End
      Begin VB.Label lblRecount 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   7440
         TabIndex        =   10
         Top             =   720
         Width           =   480
      End
   End
   Begin VB.ComboBox cmbCompany 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      ItemData        =   "frmImportDTR.frx":06F4
      Left            =   120
      List            =   "frmImportDTR.frx":0701
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   855
      Width           =   2925
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   900
      ButtonWidth     =   609
      ButtonHeight    =   741
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   960
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   8940
         Begin VB.Label Label1 
            Caption         =   "Employee Daily Time Records"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3480
            TabIndex        =   3
            Top             =   45
            Width           =   3975
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   6210
         TabIndex        =   1
         Top             =   900
         Width           =   570
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3630
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   1485
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   6403
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Client Name"
         Object.Width           =   4974
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Client ID"
         Object.Width           =   2
      EndProperty
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Time Records:"
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
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   12
      Top             =   600
      Width           =   1725
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name:"
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
      Height          =   255
      Index           =   38
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name:"
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
      Height          =   255
      Index           =   39
      Left            =   120
      TabIndex        =   6
      Top             =   1245
      Width           =   1365
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Height          =   255
      Index           =   10
      Left            =   2190
      TabIndex        =   5
      Top             =   1245
      Width           =   705
   End
End
Attribute VB_Name = "frmImportDTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsClientData As ADODB.Recordset
Dim rsEmployeeBasicData As ADODB.Recordset
Dim rsDTRRecords As ADODB.Recordset
Dim rsHREmployees As ADODB.Recordset
Dim InvalidIDCtr As Long
Dim ActiveCtr As Long
Dim TotalCount As Long
Dim InvalidIDSw As Boolean
Dim DuplicateDTR As Long
Dim mDTRRec As Long

Dim StartFirstCutoff As Date
Dim EndFirstCutoff As Date
Dim StartSecondCutoff As Date
Dim EndSecondCutoff As Date
Dim FirstHalfFlag As Boolean    ' to determine whether first half or second half

Dim mDTRID As String
Dim firstRun As Boolean

Private Sub Command1_Click()
    'frmBasicSalary.txtCancelSw.Text = "0"
    Call txtSearch_KeyPress(13)
End Sub

Private Sub Command2_Click()
'If TimeKeepFlag = True Then
'    frmTimeKeep.txtCancelSw.Text = "1"
'Else
    'frmBasicSalary.txtCancelSw.Text = "1"
'End If
Unload Me
End Sub

Private Sub chkInvalidID_Click()

If chkInvalidID.VALUE = 1 Then ' view all invalidID for the month
    InvalidIDSw = True
    lblRecord.Caption = "Inv. DTR:"
Else
    InvalidIDSw = False
    lblRecord.Caption = "New DTR:"
End If
Call cmdRefresh_Click
End Sub

Private Sub cmbCompany_Click()
    Call FilterEmployee(cmbCompany.Text)
    If ActiveCtr = 0 Then
        Call DisplayBasicData("", False)
    Else
        If ListView1(0).ListItems.Count > 0 Then
            Call DisplayBasicData(ListView1(0).SelectedItem.Text, False)
        End If
    End If
End Sub

Private Sub cmdImportDTR_Click()
Dim DTRFilename As String
Dim mLine As StandardStreamTypes
Dim TextLine As String
Dim FSO As New FileSystemObject, txtfile, file1 As FILE, ts As TextStream
Dim InvRec As Long
Dim res

On Error GoTo errClear

' Set the selected path.
CommonDialog1.Filter = "DTR Files|*.txt; *.dat|"
CommonDialog1.ShowOpen
DTRFilename = CommonDialog1.Filename

mDTRRec = 0
InvRec = 0

With rsEmployeeBasicData
    .Filter = ""
End With

    ' Read the contents of the file.
    Set file1 = FSO.GetFile(DTRFilename)
    Set ts = file1.OpenAsTextStream(ForReading)
    Call DisplayMsg("Opening DTR file...")
    Do While Not ts.AtEndOfLine = True  ' Loop until end of file.
        DoEvents
        TextLine = ts.ReadLine          ' Read line into variable.
        If Len(TextLine) <> 49 And Len(TextLine) <> 33 Then ' Updated 8/16/2012. old machine 49, new machine 33
        'If Len(TextLine) <> 49 Then
            ' Invalid records
            InvRec = InvRec + 1
        Else
            mDTRRec = mDTRRec + 1
        End If
       Call DisplayMsg("Validating record " & mDTRRec & ". Total Invalid records: " & InvRec)
    Loop
    ts.Close
    MsgBox "Validation Done. Total records validated: " & mDTRRec, vbInformation, payrollert
    Call DisplayMsg("Validation Done. Total records validated: " & mDTRRec, True)
    
    mDTRRec = 0
    InvRec = 0
    If MsgBox("Proceed with DTR Importing?", vbYesNo + vbQuestion) = vbYes Then
        Set ts = file1.OpenAsTextStream(ForReading)
        Call DisplayMsg("Reading DTR file...")
        Do While Not ts.AtEndOfLine = True  ' Loop until end of file.
            DoEvents
            TextLine = ts.ReadLine          ' Read line into variable.
'            If ImportDTRRecords(TextLine) = True Then
'                mDTRRec = mDTRRec + 1
'            Else
'                InvRec = InvRec + 1
'            End If
            If Len(TextLine) = 49 Then      ' Updated 8/16/2012 using old machine 49
                If ImportDTRRecords(TextLine) = True Then
                    mDTRRec = mDTRRec + 1
                Else
                    InvRec = InvRec + 1
                End If
            Else
                If ImportNEWDTRRecords(TextLine) = True Then    ' using new machine 33
                    mDTRRec = mDTRRec + 1
                Else
                    InvRec = InvRec + 1
                End If
            End If
            
            Call DisplayMsg("Records Imported: " & mDTRRec & "   Duplicate Records: " & InvRec)
        Loop
        ts.Close
        If mDTRRec <> 0 Then
            MsgBox "There were " & mDTRRec & " new DTR records inserted!", vbInformation, PayrollAlert
        Else
            MsgBox "There were " & DuplicateDTR & " duplicate records found. No DTR records inserted!", vbInformation, PayrollAlert
        End If
        Call DisplayMsg("Importing done.", True)
    End If

'Restore the filtering.
Call cmbCompany_Click

lblRecount.Caption = RecordCount

Exit Sub

errClear:

    'MsgBox Err.Description, vbInformation, PayrollAlert
    Err.Clear
    
End Sub

Private Sub DisplayMsg(ByVal mText As String, Optional mFlag As Boolean = False)
    lblDisplay.Caption = mText
    If mFlag = False Then
        Frame4.Visible = True
    Else
        Frame4.Visible = False
    End If
End Sub

'Sub Read_Files()
'    Dim fso As New FileSystemObject, txtfile, _
'      fil1 As FILE, ts As TextStream
'    Set txtfile = fso.CreateTextFile("c:\testfile.txt", True)
'    MsgBox "Writing file"
'    ' Write a line.
'    Set fil1 = fso.GetFile("c:\testfile.txt")
'    Set ts = fil1.OpenAsTextStream(ForWriting)
'    ts.Write "Hello World"
'    ts.Close
'    ' Read the contents of the file.
'    Set ts = fil1.OpenAsTextStream(ForReading)
'    s = ts.ReadLine
'    MsgBox s
'    ts.Close
'End Sub

Private Function ImportDTRRecords(ByVal DTRrec As String) As Boolean
Dim mSeq As String
Dim mMachine As String
Dim mDTRID As String
Dim DtrName As String
Dim mMode As String
Dim IOMode As String
Dim DtrDate As String
Dim DtrTime As String
Dim InvalidID As Long

On Error GoTo errImport

ImportDTRRecords = False
If Len(DTRrec) <> 49 Then
    Exit Function
End If

'DTRrec = Replace(Replace(Replace(Replace(DTRrec, Chr(32), "|"), Chr(13), "|"), Chr(17), "|"), Chr(9), "|")
DTRrec = Replace(DTRrec, Chr(9), "|")
mSeq = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)
DTRrec = Mid(DTRrec, Len(mSeq) + 2, Len(DTRrec) - Len(mSeq))
mMachine = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)
DTRrec = Mid(DTRrec, Len(mMachine) + 2, Len(DTRrec) - Len(mMachine))
mDTRID = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)
DTRrec = Mid(DTRrec, Len(mDTRID) + 2, Len(DTRrec) - Len(mDTRID))
DtrName = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)
DTRrec = Mid(DTRrec, Len(DtrName) + 2, Len(DTRrec) - Len(DtrName))
mMode = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)
DTRrec = Mid(DTRrec, Len(mMode) + 2, Len(DTRrec) - Len(mMode))
IOMode = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)
DTRrec = Mid(DTRrec, Len(IOMode) + 2, Len(DTRrec) - Len(IOMode))
DtrDate = Mid(DTRrec, 1, Len(DTRrec))
DtrTime = Right(DtrDate, 5)

If mSeq <> "No" Then
    ' Search if DTR ID in machine does not match in HR list.
    With rsEmployeeBasicData
        .MoveFirst
        .Find "DTRID = '" & mDTRID & "'"
        If .EOF = True Then
            InvalidID = 1
            InvalidIDCtr = InvalidIDCtr + 1
        Else
            InvalidID = 0
        End If
    End With

    StrQuery = "INSERT INTO DTRRecords (DTR_Seq, Machine_ID, DTRID, DTR_Mode, DTR_IOMode, DTR_Date, DTR_Time, ImportID, InvalidIDFlag) " & _
        " VALUES ('" & mSeq & "', '" & mMachine & "', '" & mDTRID & "', '" & mMode & "', '" & IOMode & "', '" & DtrDate & "', '" & DtrTime & "', " & operatorID & ", " & InvalidID & ")"
    If Action_Query(StrQuery, cnnstr) = True Then
        ImportDTRRecords = True
    Else
        DuplicateDTR = DuplicateDTR + 1
    End If
  
End If

Exit Function

errImport:

    MsgBox Err.Description, vbInformation, PayrollAlert
    Err.Clear
    Resume Next
'No  Mchn    EnNo        Name        Mode    IOMd    DateTime
'00001   1   00000001                1   0   2010/01/16  16:58

End Function

Private Function ImportNEWDTRRecords(ByVal DTRrec As String) As Boolean  ' For New Biometric Machine
Dim mSeq As String
Dim mMachine As String
Dim mDTRID As String
Dim DtrName As String
Dim mMode As String
Dim IOMode As String, DTR_IN As String, DTR_OUT As String
Dim DtrDate As String
Dim DtrTime As String
Dim InvalidID As Long

On Error GoTo errImport

ImportNEWDTRRecords = False
If Len(DTRrec) <> 33 Then ' 49
    Exit Function
End If

mMachine = "2"
'DTRrec = Replace(Replace(Replace(Replace(DTRrec, Chr(32), "|"), Chr(13), "|"), Chr(17), "|"), Chr(9), "|")
DTRrec = Replace(DTRrec, Chr(9), "|")
StrQuery = DTRrec
mDTRID = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)   ' DTR ID
mSeq = mDTRRec

'mSeq = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)
DTRrec = Mid(DTRrec, Len(mDTRID) + 2, Len(DTRrec) - Len(mDTRID))
mDTRID = Right("00000000" & LTrim(mDTRID), 8)   ' Format DTR ID

'mMachine = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)
DtrDate = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)  ' DTR Date and Time
DTRrec = Mid(DTRrec, Len(DtrDate) + 2, Len(DTRrec) - Len(DtrDate))
DtrTime = Right(LTrim(RTrim(DtrDate)), 8)
DtrTime = Mid(DtrTime, 1, 5)

DTR_IN = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)   ' Get Time IN
DTRrec = Mid(DTRrec, Len(DTR_IN) + 2, Len(DTRrec) - Len(DTR_IN))

'DtrName = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)
DTR_OUT = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)  ' Get Time OUT
DTRrec = Mid(DTRrec, Len(DTR_OUT) + 2, Len(DTRrec) - Len(DTR_OUT))

' Machine Mode.
mMode = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)

If LTrim(RTrim(DTR_OUT)) = "1" Then
    IOMode = 1  ' 1 means OUT
Else
    IOMode = 0  ' 0 means IN
End If
'DTRrec = Mid(DTRrec, Len(mMode) + 2, Len(DTRrec) - Len(mMode))
'IOMode = Mid(DTRrec, 1, InStr(1, DTRrec, "|", vbTextCompare) - 1)
'DTRrec = Mid(DTRrec, Len(IOMode) + 2, Len(DTRrec) - Len(IOMode))
'DtrDate = Mid(DTRrec, 1, Len(DTRrec))

If mSeq <> "No" Then
    ' Search if DTR ID in machine does not match in HR list.
    With rsEmployeeBasicData
        .MoveFirst
        .Find "DTRID = '" & mDTRID & "'"
        If .EOF = True Then
            InvalidID = 1
            InvalidIDCtr = InvalidIDCtr + 1
        Else
            InvalidID = 0
        End If
    End With
    StrQuery = "INSERT INTO DTRRecords (DTR_Seq, Machine_ID, DTRID, DTR_Mode, DTR_IOMode, DTR_Date, DTR_Time, ImportID, InvalidIDFlag) " & _
        " VALUES ('" & mSeq & "', '" & mMachine & "', '" & mDTRID & "', '" & mMode & "', '" & IOMode & "', '" & DtrDate & "', '" & DtrTime & "', " & operatorID & ", " & InvalidID & ")"
    If Action_Query(StrQuery, cnnstr) = True Then
        ImportNEWDTRRecords = True
    Else
        DuplicateDTR = DuplicateDTR + 1
    End If
  
End If

Exit Function

errImport:

    MsgBox Err.Description, vbInformation, PayrollAlert
    Err.Clear
    'Resume Next

'Old Machine layout.
'No  Mchn    EnNo        Name        Mode    IOMd    DateTime
'00001   1   00000001                1   0   2010/01/16  16:58

'New Machine Layout.
'No  Mchn    EnNo        Name        Mode    IOMd    DateTime
'    1   2012-05-29 10:37:02 1   0   0   0
'    1   2012-05-29 10:37:16 1   0   0   0
'    1   2012-06-02 13:43:26 1   0   0   0
'    1   2012-06-02 13:43:33 1   0   0   0
'    1   2012-06-02 13:43:44 1   0   0   0
'    1   2012-06-02 13:51:25 1   0   0   0
' 1075   2012-06-02 13:53:23 1   0   0   0
' 1075   2012-06-02 13:53:28 1   0   0   0
'    1   2012-06-02 13:53:33 1   0   0   0
'    1   2012-06-02 13:53:39 1   0   0   0
' 1858   2012-06-02 13:55:25 1   0   0   0

End Function

Private Sub cmdPrintInvalidID_Click()
If lstDTRDisplay.ListItems.Count <> 0 Then
    Call PrintPreviewDTR
Else
    MsgBox "No records selected!", vbInformation, PayrollAlert
End If
End Sub

Private Sub PrintPreviewDTR()
Dim rsDedRecords As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String, mRange2 As String
Dim sumString As String
Dim ctr As Long
Dim i As Long
Dim cPos As Long
Dim LineCtr As Long
Dim mHead As String
Dim mDeduc As Double
Dim StartDeduction As Variant
Dim EndDeduction As Variant
Dim DeductionYear As Long

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)

ctr = 8
i = 1

'Updated by Bernard May 13, 2008.
mHead = CompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 10, True)
If chkInvalidID.VALUE = 1 Then
    mHead = "List of Invalid DTR ID Report"
ElseIf chkNoFilter = 1 Then 'for all employees
    mHead = "DTR Records Report for All Employees"
Else
    mHead = "DTR Records Report for " & ListView1(0).SelectedItem.Text
End If
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy")
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "COMPANY NAME: " & cmbCompany.Text
Call PrintReportHeader(xlSheet, xlApp1, "A5", mHead, "", xlLeft, , False, False, False, , , 10, True)

Call PrintReportHeader(xlSheet, xlApp1, "A6:A7", "Record #", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "B6:B7", "DTR ID", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "C6:C7", "DTR Date", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "D6:D7", "Time In", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "E6:E7", "Time Out", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "F6:F7", "Work Hours", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "G6:G7", "Import Date", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "H6:H7", "Remarks", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 20)

For i = 1 To lstDTRDisplay.ListItems.Count
        xlSheet.Cells(ctr, 1).VALUE = i
        xlSheet.Cells(ctr, 2).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(1) ' DTR ID
        xlSheet.Cells(ctr, 2).NumberFormat = "00000000"
        xlSheet.Cells(ctr, 3).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(2) ' DTR Date
        xlSheet.Cells(ctr, 4).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(3) ' Time In
        xlSheet.Cells(ctr, 5).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(4) ' Time Out
        xlSheet.Cells(ctr, 6).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(5) ' Work Hours
        xlSheet.Cells(ctr, 7).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(6) ' Import Date
        xlSheet.Cells(ctr, 8).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(7) ' Import Date
        ctr = ctr + 1
Next i
ctr = ctr + 2

'Box it.
mRange = "A6:H" & ctr
Call PlaceBox(xlSheet, xlApp1, mRange)

'ctr = ctr + 1
'mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
'mRange = "A" & ctr
'Call CustomizedHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 8, False, 25)

xlSheet.Range("A1").Select

ExitRtn2:

End Sub

Private Sub cmdRefresh_Click()

If FMSFlag = False Then
    ' Get DTR records for the selected range of date.
    If chkNoFilter.VALUE = 1 Then   ' get all employees
        If InvalidIDSw = False Then ' get all valid IDs
            StrQuery = "SELECT * FROM vwDTRRecords WHERE InvalidIDFlag = 0 AND DTR_Date BETWEEN '" & _
                Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "' ORDER BY DTRID "
        Else    ' get all invalid IDs
            StrQuery = "SELECT * FROM vwDTRRecords WHERE InvalidIDFlag = 1 AND DTR_Date BETWEEN '" & _
                Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "' ORDER BY DTRID "
        End If
    Else    ' get specific employee only
        If InvalidIDSw = False Then ' get all valid IDs
            StrQuery = "SELECT * FROM vwDTRRecords WHERE InvalidIDFlag = 0 AND DTRID = '" & mDTRID & "' AND DTR_Date BETWEEN '" & _
                Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "' ORDER BY DTRID "
        Else    ' get all invalid IDs for specific employee
            StrQuery = "SELECT * FROM vwDTRRecords WHERE InvalidIDFlag = 1 AND DTRID = '" & mDTRID & "' AND DTR_Date BETWEEN '" & _
                Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "' ORDER BY DTRID "
        End If
    End If
Else
    ' Get DTR records for the selected range of date.
    If chkNoFilter.VALUE = 1 Then   ' get all employees
        If InvalidIDSw = False Then ' get all valid IDs
            StrQuery = "SELECT * FROM vwDTRConductors WHERE InvalidIDFlag = 0 AND DTR_Date BETWEEN '" & _
                Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "' ORDER BY DTRID "
        Else    ' get all invalid IDs
            StrQuery = "SELECT * FROM vwDTRConductors WHERE InvalidIDFlag = 1 AND DTR_Date BETWEEN '" & _
                Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "' ORDER BY DTRID "
        End If
    Else    ' get specific employee only
        If InvalidIDSw = False Then ' get all valid IDs
            StrQuery = "SELECT * FROM vwDTRConductors WHERE InvalidIDFlag = 0 AND DTRID = '" & mDTRID & "' AND DTR_Date BETWEEN '" & _
                Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "' ORDER BY DTRID "
        Else    ' get all invalid IDs for specific employee
            StrQuery = "SELECT * FROM vwDTRConductors WHERE InvalidIDFlag = 1 AND DTRID = '" & mDTRID & "' AND DTR_Date BETWEEN '" & _
                Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "' ORDER BY DTRID "
        End If
    End If
End If

Set rsDTRRecords = ReturnRst(StrQuery, cnnstr, mstatic)

If ActiveCtr = 0 Then
    Call DisplayBasicData("", False)
Else
    If ListView1(0).ListItems.Count > 0 Then
        Call DisplayBasicData(ListView1(0).SelectedItem.Text, False)
    End If
End If

If firstRun = False Then
    If rsDTRRecords.RecordCount = 0 Then
        MsgBox "No DTR Records selected!", vbInformation, PayrollAlert
    ElseIf chkNoFilter.VALUE = 1 Then
        MsgBox "There were " & rsDTRRecords.RecordCount & " DTR records selected.", vbInformation, PayrollAlert
    End If
End If

End Sub

Private Sub UpdateDTRRecords()

If rsDTRRecords.State = 1 Then
    rsDTRRecords.Close
End If

' Get DTR records for the selected range of date.
If InvalidIDSw = False Then ' get all valid IDs
    StrQuery = "SELECT * FROM vwDTRRecords WHERE InvalidIDFlag = 0 AND DTR_Date BETWEEN '" & _
        Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "' ORDER BY DTRID "
Else    ' get all invalid IDs
    StrQuery = "SELECT * FROM vwDTRRecords WHERE InvalidIDFlag = 1 AND DTR_Date BETWEEN '" & _
        Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "' ORDER BY DTRID "
End If
Set rsDTRRecords = ReturnRst(StrQuery, cnnstr, mstatic)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Dim mComp As String
Dim mEmpN As String
Dim y As Integer

    firstRun = True
    Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    
    InvalidIDSw = False
    FirstHalfFlag = True

    'Assign Cutoff dates.
    DTRStart = Now - 20
    DTREnd = Now - 5

    Call GetEmployeeBasic
    If ActiveCtr = 0 Then
        Call DisplayBasicData("", True)
    Else
        Call DisplayBasicData(CurrEmployee, True)
    End If
    
    ' Populate month and year.
    y = 1
    cmbYear.Clear
    Do While y <= 10
        DoEvents
        cmbYear.AddItem 2003 + y
        cmbYear.ItemData(cmbYear.NewIndex) = y
        y = y + 1
    Loop
    'Populate Month.
    cmbMonth.AddItem "01-January"
    cmbMonth.AddItem "02-February"
    cmbMonth.AddItem "03-March"
    cmbMonth.AddItem "04-April"
    cmbMonth.AddItem "05-May"
    cmbMonth.AddItem "06-June"
    cmbMonth.AddItem "07-July"
    cmbMonth.AddItem "08-August"
    cmbMonth.AddItem "09-September"
    cmbMonth.AddItem "10-October"
    cmbMonth.AddItem "11-November"
    cmbMonth.AddItem "12-December"
    cmbYear.Text = Format(Now, "yyyy")
    cmbMonth.ListIndex = Format(Date, "mm") - 1
    cmbCompany.ListIndex = 0
    firstRun = False
    
End Sub

Private Sub GetEmployeeBasic()
Dim rsInvCode As ADODB.Recordset

'If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'    StrQuery = "SELECT * FROM clienttable WHERE CLIENTCODE2 = 'FMS' AND clientstatus = 1 and clientname is not null ORDER BY clientname"
'Else
'    StrQuery = "SELECT * FROM clienttable WHERE clientstatus = 1 and clientname is not null ORDER BY clientname"
'End If
' Get all Active Company.
StrQuery = GetClientRecord(0, True, " ORDER BY CLientName")
Set rsClientData = ReturnRst(StrQuery, cnnstr, mstatic)

If rsClientData.RecordCount <> 0 Then
    Call FillComboData(cmbCompany, rsClientData)
    rsClientData.MoveFirst
    txtClientCode = rsClientData!ClientCode
Else
    ' No client selected disable text boxes and command buttons.
    Exit Sub
End If

StrQuery = "SELECT * FROM qryHREmployeeHired WHERE EmpStatus = 1 ORDER BY EmpName "
Set rsEmployeeBasicData = ReturnRst(StrQuery, cnnstr, mstatic)
If rsEmployeeBasicData.RecordCount <> 0 Then
    Call FilterEmployee(rsClientData!ClientName)
End If

' Count for Invalid ID.
StrQuery = "SELECT DTRID FROM DTRRecords WHERE InvalidIDFlag = 1 AND ProcessFlag = 0"
Set rsInvCode = ReturnRst(StrQuery, cnnstr, mstatic)
InvalidIDCtr = rsInvCode.RecordCount
rsInvCode.Close

If InvalidIDCtr <> 0 Then
    chkInvalidID.Caption = "Invalid ID (" & InvalidIDCtr & ")"
Else
    chkInvalidID.Caption = "Invalid ID"
End If

'Get DTR records for the current month.
'StrQuery = "SELECT * FROM vwDTRRecords WHERE YEAR(DTR_Date) = " & _
'    IIf(cmbYear.Text = "", Format(Date, "yyyy"), cmbYear.Text) & _
'    " AND MONTH(DTR_Date) = " & Val(IIf(cmbMonth = "", Format(Date, "mm"), Left(cmbMonth.Text, 2))) & " ORDER BY DTRID "

'StrQuery = "SELECT * FROM vwDTRRecords WHERE DTR_Date BETWEEN '" & Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "'"
'Set rsDTRRecords = ReturnRst(StrQuery, cnnstr, mstatic)

Call cmdRefresh_Click

End Sub

Private Function FillComboData(ByRef mCmb As ComboBox, _
                        mRs As Recordset, Optional FilterFlag As Boolean, _
                        Optional ComboList As String = "")
    mCmb.Clear
    While Not (mRs.EOF Or mRs.BOF)
        If FilterFlag = False Then
            mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1)), "'", "''"))
            mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
        Else    ' need to filter the lists
            If InStr(1, ComboList, LTrim(RTrim(str(mRs.Fields(0)))), 1) = 0 Then
                mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1)), "'", "''"))
                mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
            End If
        End If
        mRs.MoveNext
    Wend

End Function

Private Sub FilterEmployee(ByVal mCompany As String)
    rsClientData.MoveFirst
    rsClientData.Find "Clientname = '" & Trim(mCompany) & "'"
    If rsClientData.EOF = True Then
        txtClientCode = 0
        Exit Sub
    Else
        FMSFlag = IIf(IsNull(rsClientData!ClientCode2) Or rsClientData!ClientCode2 <> "FMS", False, True) ' fmsflag
        txtClientCode = rsClientData!ClientCode
    End If
    With rsEmployeeBasicData
        .Filter = "ClientCode = " & Val(txtClientCode)
        If .RecordCount <> 0 Then
            'If Not .EOF And Not .BOF Then
            Call FillEmployeeList(ListView1(0), rsEmployeeBasicData)
            'rsEmployeeBasicData.MoveFirst
            lblclient(10).Caption = "Total:  " & ActiveCtr
             .Filter = "Clientcode = " & Val(txtClientCode)
        Else
            ListView1(0).ListItems.Clear
        End If
    End With
End Sub

Private Sub FillEmployeeList(ByRef mLsEmp As ListView, _
                        ByRef mRs As Recordset)
Dim litem1 As ListItem
Dim litem2 As ListItem

ActiveCtr = 0
mLsEmp.View = lvwReport
mLsEmp.ListItems.Clear
ListView1(0).ListItems.Clear
While Not (mRs.EOF Or mRs.BOF)
    'Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!EmpName), "Blank Name", mRs!EmpName & " (" & mRs!DTRID & ")"))
    Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!EmpName), "Blank Name", mRs!EmpName))
        litem1.SubItems(1) = mRs(0)
        ActiveCtr = ActiveCtr + 1
    mRs.MoveNext
Wend

'    'Get all Inactive employees regardless of employers.
'    mRs.MoveFirst
'    'mRs.Filter = "Employeestatus = false"
'    mRs.Filter = "Employeestatus = 0"
'    While Not (mRs.EOF Or mRs.BOF)
'        Set litem2 = ListView1(1).ListItems.Add(, , 1)
'        'Set litem2 = ListView1(1).ListItems.Add(, , IIf(IsNull(mRs!FullName, "Blank", mRs!FullName)))
'        litem2.SubItems(1) = IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName)
'        litem2.SubItems(2) = IIf(IsNull(mRs!ClientName), "Blank", mRs!ClientName)
'        'litem2.SubItems(3) = IIf(IsNull(mRs!EmployeeID), "Blank", mRs!EmployeeID)
'        litem2.SubItems(3) = IIf(IsNull(mRs!EmpNo2), "Blank", mRs!EmpNo2)
'        litem2.SubItems(4) = IIf(IsNull(mRs!StartContract) Or mRs!StartContract = #1/1/1900#, "__/__/____", Format(mRs!StartContract, "mm/dd/yyyy"))
'        litem2.SubItems(5) = IIf(IsNull(mRs!endcontract) Or mRs!endcontract = #1/1/1900#, "__/__/____", Format(mRs!endcontract, "mm/dd/yyyy"))
'        litem2.SubItems(6) = IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName)
'        litem2.SubItems(7) = IIf(IsNull(mRs!EmployeeID), 1, mRs!EmployeeID)
'        InActiveCtr = InActiveCtr + 1
'        mRs.MoveNext
'    Wend

End Sub

Private Sub FillHREmployeeList(ByRef mLsEmp As ListView, _
                        mRs As Recordset)
    Dim litem1 As ListItem

    mLsEmp.View = lvwReport
    mLsEmp.ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
        Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!Surname), "Blank Name", mRs!Surname & ", " & mRs!FirstName))
        litem1.SubItems(1) = IIf(IsNull(mRs!EmpNo), "", mRs!EmpNo) 'EmployeeID
        mRs.MoveNext
    Wend

End Sub

Private Sub Form_Unload(Cancel As Integer)
If TimeKeepFlag = True Then
    frmTimeKeep.txtCancelSw.Text = "1"
End If
End Sub

Private Sub lvwdisplay_DblClick()
    strRow = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).SubItems(1)
    txtSearch = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).Text
    If lvwdisplay.ListItems.Count = 0 Then
        MsgBox "List doesn't have any record from the table selected !", vbOKOnly + vbInformation, "Select..."
        Exit Sub
    End If
    Call Search(strRow)
    Unload Me
End Sub

Private Sub lvwdisplay_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtSearch = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).Text
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        strRow = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).SubItems(1)
        Call Search(strRow)
    End If
End Sub

Private Sub Search(ByVal strSearch As String)
Dim res, bkmrk

With rsHREmployees
    If .EOF Then
        .MoveLast
    ElseIf .BOF Then
        .MoveFirst
    End If
    
    bkmrk = .Bookmark
    .MoveFirst
    .Find "EmpNo= " & strSearch
    If Not .EOF Then
        Call Showdata(rsHREmployees)
    Else
        .Bookmark = bkmrk
    End If
    Unload Me
End With
End Sub

Private Sub DisplayBasicData(ByVal SelectedEmployee As String, _
    Optional recSw As Boolean)
Dim litem2 As ListItem
Dim LineNo As Long

LineNo = 0

If rsEmployeeBasicData.RecordCount = 0 Then Exit Sub
With rsEmployeeBasicData
    If recSw = False Then
        .MoveFirst
        .Find "Empname = '" & LTrim(RTrim(SelectedEmployee)) & "'"
    End If
    If Not .EOF And Not .BOF Then
        mDTRID = IIf(IsNull(!DTRID), "00000000", !DTRID)
        txtClient(0) = IIf(IsNull(!EmpNo), "", !EmpNo)
        txtClient(2) = LTrim(RTrim(SelectedEmployee))
        txtClient(8) = IIf(IsNull(!ShiftName), "", !ShiftName) ' get the shift Name
        txtClient(9) = IIf(IsNull(!DTRID), "", !DTRID) ' get DTR ID
        If rsDTRRecords.State = 1 And rsDTRRecords.RecordCount <> 0 Then
            'Get all DTR per employee.
            rsDTRRecords.Filter = ""
            rsDTRRecords.MoveFirst
            If InvalidIDSw = False Then
                If chkNoFilter.VALUE = 0 Then
                    rsDTRRecords.Filter = "DTRID = '" & mDTRID & "'"   ' look for dtr number
                End If
            End If
            lstDTRDisplay.ListItems.Clear
            While Not (rsDTRRecords.EOF Or rsDTRRecords.BOF)
                LineNo = LineNo + 1
                'Set litem2 = lstDTRDisplay.ListItems.Add(, , LineNo)
                Set litem2 = lstDTRDisplay.ListItems.Add(, , IIf(IsNull(rsDTRRecords!Machine_ID), "X", rsDTRRecords!Machine_ID))
                'litem2.SubItems(1) = IIf(IsNull(rsDTRRecords!Machine_ID), "X", rsDTRRecords!Machine_ID)
                'litem2.SubItems(1) = IIf(IsNull(rsDTRRecords!DTR_Mode), "X", rsDTRRecords!DTR_Mode)
                litem2.SubItems(1) = IIf(IsNull(rsDTRRecords!DTRID), "NO DTR ID", rsDTRRecords!DTRID)
                litem2.SubItems(2) = IIf(IsNull(rsDTRRecords!DTR_Date), "NO DTR DATE", Format(rsDTRRecords!DTR_Date, "mm/dd/yyyy"))
                litem2.SubItems(3) = IIf(IsNull(rsDTRRecords!TimeIN), "NO Time In", rsDTRRecords!TimeIN)
                litem2.SubItems(4) = IIf(IsNull(rsDTRRecords!TimeOut), "NO Time Out", rsDTRRecords!TimeOut)
                litem2.SubItems(5) = IIf(IsNull(rsDTRRecords!WorkHours), "0.00", rsDTRRecords!WorkHours)
                litem2.SubItems(6) = IIf(IsNull(rsDTRRecords!ImportDate), "", Format(rsDTRRecords!ImportDate, "mm/dd/yyyy"))
                litem2.SubItems(7) = IIf(IsNull(rsDTRRecords!Remarks), "", rsDTRRecords!Remarks)
                rsDTRRecords.MoveNext
            Wend
        End If
        RecordCount = rsDTRRecords.RecordCount
        rsDTRRecords.Filter = ""
    End If
End With
lblRecount.Caption = RecordCount

End Sub

Private Sub Showdata(ByRef rec As ADODB.Recordset)
    
If TimeKeepFlag = True Then
    frmTimeKeep.txtEmpNum = IIf(IsNull(rec!EmployeeID), 0, rec!EmployeeID)
    frmTimeKeep.txtClient(0) = IIf(IsNull(rec!EmpNo), "", rec!EmpNo)
    frmTimeKeep.txtClient(1) = IIf(IsNull(rec!EmploymentStatusDesc), "", rec!EmploymentStatusDesc)
    frmTimeKeep.txtClient(2) = IIf(IsNull(rec!Surname), "", rec!Surname & ", " & rec!FirstName)
    frmTimeKeep.txtRates(4) = IIf(IsNull(rec!EcolaAmount), 0, rec!EcolaAmount)
Else
    frmBasicSalary.txtEmpNum = IIf(IsNull(rec!EmployeeID), 0, rec!EmployeeID)
    frmBasicSalary.txtClient(0) = IIf(IsNull(rec!EmpNo), "", rec!EmpNo)
    frmBasicSalary.txtClient(1) = IIf(IsNull(rec!EmploymentStatusDesc), "", rec!EmploymentStatusDesc)
    frmBasicSalary.txtClient(2) = IIf(IsNull(rec!Surname), "", rec!Surname & ", " & rec!FirstName)
    frmBasicSalary.txtRates(4) = IIf(IsNull(rec!EcolaAmount), 0, rec!EcolaAmount)
End If
End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    InvalidIDSw = False
    Call UpdateDTRRecords
    Call DisplayBasicData(Item.Text, False)
End Sub

Private Sub lstDTRDisplay_DblClick()
Dim res
Dim invDTR As String

If chkInvalidID.VALUE = 1 Then
invDTR = lstDTRDisplay.SelectedItem.ListSubItems.Item(1).Text
'invDTR = lstDTRDisplay.ListItems.Item(3).ListSubItems.Item(1).Text
If MsgBox("Replace DTR ID " & invDTR & " with " & txtClient(9).Text & " of " & txtClient(2) & "?", vbYesNo + vbQuestion, "Replace DTR ID") = vbYes Then
    StrQuery = "EXEC ReplaceDTRID '" & invDTR & "', '" & txtClient(9).Text & "', ' " & LoginName & "'"
    If Action_Query(StrQuery, cnnstr) = True Then
        MsgBox "DTR ID of " & txtClient(2) & " was successfully replaced."
    End If
End If
End If

End Sub

