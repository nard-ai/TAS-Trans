VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPHealthReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Philhealth Report"
   ClientHeight    =   8520
   ClientLeft      =   2370
   ClientTop       =   2145
   ClientWidth     =   12600
   Icon            =   "frmPHealthReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   12600
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtRemittedAmount 
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
      Left            =   8760
      MaxLength       =   20
      TabIndex        =   18
      Text            =   "0.00"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdER2 
      Caption         =   "ER2 Report"
      Height          =   375
      Left            =   9480
      TabIndex        =   19
      ToolTipText     =   "ER2 Report"
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearchSBR 
      Caption         =   "Search"
      Height          =   255
      Left            =   5760
      TabIndex        =   60
      ToolTipText     =   "Search SBR Number"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox txtRemittedAmt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   64
      Text            =   "0.00"
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txtSBRNo 
      Alignment       =   2  'Center
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
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   63
      Text            =   "0"
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdateSBR 
      Caption         =   "Update"
      Height          =   255
      Left            =   5760
      TabIndex        =   61
      ToolTipText     =   "Update SBR Entries"
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelectGroupCodes 
      Caption         =   "Select"
      Height          =   255
      Left            =   11760
      TabIndex        =   57
      ToolTipText     =   "Select Group Payments"
      Top             =   1320
      Width           =   615
   End
   Begin VB.ComboBox cmbGroupCode 
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
      Height          =   315
      ItemData        =   "frmPHealthReport.frx":08CA
      Left            =   5040
      List            =   "frmPHealthReport.frx":08F2
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   6480
      Width           =   2085
   End
   Begin VB.CommandButton cmdClearView 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   240
      TabIndex        =   53
      ToolTipText     =   "process"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdExcelReport 
      Caption         =   "Generate Excel Report"
      Height          =   375
      Left            =   5880
      TabIndex        =   52
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton cmdCheckZero 
      Caption         =   "Check Zero Contributions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9840
      TabIndex        =   48
      Top             =   6590
      Width           =   2535
   End
   Begin VB.CommandButton cmdNoPHNo 
      Caption         =   "Check/Uncheck No PHealth No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9840
      TabIndex        =   47
      Top             =   6590
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10920
      TabIndex        =   46
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "Retrieve Payments"
      Height          =   375
      Left            =   1560
      TabIndex        =   45
      ToolTipText     =   "process"
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CheckBox chkIncluded 
      Caption         =   "Exclude Checked Records"
      Height          =   255
      Left            =   2520
      TabIndex        =   44
      Top             =   6600
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdGenerateRF1Report 
      Caption         =   "RF-1 Report"
      Height          =   375
      Left            =   8040
      TabIndex        =   41
      ToolTipText     =   "RF-1 Report"
      Top             =   7920
      Width           =   1335
   End
   Begin VB.ComboBox cmdAppYear 
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
      Height          =   315
      ItemData        =   "frmPHealthReport.frx":0958
      Left            =   10800
      List            =   "frmPHealthReport.frx":0971
      TabIndex        =   22
      Text            =   "cmdAppYear"
      Top             =   1320
      Width           =   885
   End
   Begin VB.TextBox txtDesignation 
      Alignment       =   2  'Center
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
      Left            =   10200
      MaxLength       =   30
      TabIndex        =   40
      Top             =   7440
      Width           =   2175
   End
   Begin VB.TextBox txtPrintedName 
      Alignment       =   2  'Center
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
      Left            =   10200
      MaxLength       =   30
      TabIndex        =   38
      Top             =   7125
      Width           =   2175
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Add/Delete Employees"
      Height          =   375
      Left            =   3720
      TabIndex        =   23
      ToolTipText     =   "process"
      Top             =   7920
      Width           =   2055
   End
   Begin VB.ComboBox cmdAppMonth 
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
      Height          =   315
      ItemData        =   "frmPHealthReport.frx":099F
      Left            =   10800
      List            =   "frmPHealthReport.frx":09C7
      TabIndex        =   21
      Text            =   "cmdAppMonth"
      Top             =   840
      Width           =   1245
   End
   Begin VB.TextBox txtSubTotal 
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
      Left            =   8760
      MaxLength       =   20
      TabIndex        =   36
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox txtAppMonthYear 
      Alignment       =   2  'Center
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
      Index           =   1
      Left            =   240
      MaxLength       =   20
      TabIndex        =   32
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txtReceivedBy 
      Alignment       =   2  'Center
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
      Left            =   10800
      MaxLength       =   30
      TabIndex        =   25
      Top             =   2160
      Width           =   1575
   End
   Begin VB.OptionButton optEmployerType 
      Caption         =   "HOUSEHOLD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6480
      TabIndex        =   17
      Top             =   1200
      Width           =   1335
   End
   Begin VB.OptionButton optEmployerType 
      Caption         =   "GOVERNMENT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   16
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdCompany 
      Caption         =   "..."
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      ToolTipText     =   "Select Company"
      Top             =   1750
      Width           =   375
   End
   Begin VB.OptionButton optEmployerType 
      Caption         =   "PRIVATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   14
      Top             =   480
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox txtTelephoneNo 
      Alignment       =   2  'Center
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
      Left            =   7440
      MaxLength       =   40
      TabIndex        =   12
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtMailingAddress 
      Alignment       =   2  'Center
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
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   9
      Top             =   2160
      Width           =   7935
   End
   Begin VB.TextBox txtEmployerName 
      Alignment       =   2  'Center
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox txtEmpTIN 
      Alignment       =   2  'Center
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
      Left            =   3120
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txtPhilHealthNo 
      Alignment       =   2  'Center
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
      Left            =   3120
      MaxLength       =   40
      TabIndex        =   5
      Top             =   960
      Width           =   3135
   End
   Begin MSMask.MaskEdBox mskDateReceived 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   11040
      TabIndex        =   24
      Top             =   1800
      Width           =   1305
      _ExtentX        =   2302
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Index           =   1
      Left            =   240
      TabIndex        =   42
      Top             =   2640
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   6165
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   24
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Client Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Surname"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Given Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Middle Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "PhilHealth No."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Basic Salary"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Bracket"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Personal Share"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Employer Share"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Personal PH"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Employer PH"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Difference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Client Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Index           =   0
      Left            =   240
      TabIndex        =   51
      Top             =   2640
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   6165
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Client Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Surname"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Given Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Middle Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "SSS No."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Basic Salary"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Bracket"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Employee Share"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Employer Share"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ECC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Employee SSS"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Employer SSS"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Difference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CheckBox chkAllPayments 
      Height          =   255
      Left            =   12120
      TabIndex        =   58
      ToolTipText     =   "All Payments"
      Top             =   840
      Width           =   255
   End
   Begin MSMask.MaskEdBox MaskDatePaid 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   65
      Top             =   7320
      Width           =   1065
      _ExtentX        =   1879
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
   Begin MSMask.MaskEdBox MaskDatePosted 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   66
      Top             =   7320
      Width           =   1065
      _ExtentX        =   1879
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
   Begin VB.TextBox txtNoOfEmp 
      Alignment       =   2  'Center
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
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   62
      Top             =   7320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtGrandTotal 
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
      Left            =   8760
      MaxLength       =   20
      TabIndex        =   37
      Top             =   7440
      Width           =   1215
   End
   Begin MSMask.MaskEdBox mskGenerationDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   8760
      TabIndex        =   20
      Top             =   7440
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
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
      Caption         =   "Excluded Employees:"
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
      Index           =   15
      Left            =   360
      TabIndex        =   72
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "No SSS No."
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
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   71
      Top             =   6240
      Width           =   1605
   End
   Begin VB.Label Label2 
      Caption         =   "Amount Paid"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   1440
      TabIndex        =   70
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "SBR# / OR#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   13
      Left            =   2400
      TabIndex        =   69
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Date Paid"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   3360
      TabIndex        =   68
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Date Posted"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   4440
      TabIndex        =   67
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Payments:"
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
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   59
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "No SSS No."
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
      Height          =   255
      Index           =   0
      Left            =   10680
      TabIndex        =   55
      Top             =   6240
      Width           =   1605
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "No SSS No."
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
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   54
      Top             =   6240
      Width           =   1605
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "Zero Contributions:"
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
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   50
      Top             =   6600
      Width           =   2565
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "No PHealth No."
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
      Height          =   255
      Index           =   7
      Left            =   5400
      TabIndex        =   49
      Top             =   6600
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label2 
      Caption         =   "Included Employees:"
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
      Index           =   19
      Left            =   360
      TabIndex        =   43
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   14
      X1              =   120
      X2              =   12480
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   10
      X1              =   120
      X2              =   12480
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   13
      X1              =   10080
      X2              =   10080
      Y1              =   6960
      Y2              =   7800
   End
   Begin VB.Label Label2 
      Caption         =   "Printed Name and Designation:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   18
      Left            =   10200
      TabIndex        =   39
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "GRANDTOTAL (PS+ES)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   17
      Left            =   7080
      TabIndex        =   35
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "SUBTOTAL (PS+ES)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   16
      Left            =   7080
      TabIndex        =   34
      Top             =   7150
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   12
      X1              =   6960
      X2              =   6960
      Y1              =   6960
      Y2              =   7800
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   11
      X1              =   120
      X2              =   12480
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Label Label2 
      Caption         =   "Applicable Period:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   33
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   10
      Left            =   10800
      TabIndex        =   31
      Top             =   1155
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Month:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   9
      Left            =   10800
      TabIndex        =   30
      Top             =   690
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   9
      X1              =   9600
      X2              =   9600
      Y1              =   1680
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   120
      X2              =   12480
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   8
      X1              =   120
      X2              =   12480
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label2 
      Caption         =   "RECEIVED BY:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   8
      Left            =   9720
      TabIndex        =   29
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "DATE RECEIVED:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   7
      Left            =   9720
      TabIndex        =   28
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "APPLICABLE PERIOD:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   10920
      TabIndex        =   27
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   7
      X1              =   12480
      X2              =   12480
      Y1              =   120
      Y2              =   8400
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   6
      X1              =   10680
      X2              =   10680
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "REPORT TYPE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   5
      Left            =   8040
      TabIndex        =   26
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   5
      X1              =   7920
      X2              =   7920
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "EMPLOYER TYPE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   4
      Left            =   6480
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   4
      X1              =   6360
      X2              =   6360
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "Telephone No:"
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
      Index           =   3
      Left            =   6240
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   1080
      X2              =   6360
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   12480
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label2 
      Caption         =   "Mailing Address:"
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
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Employer Name:"
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
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   3
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   8400
   End
   Begin VB.Label Label1 
      Caption         =   "EMPLOYER TIN:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "PHILHEALTH NO."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "EMPLOYER'S REMITTANCE REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "REVISED JAN 2008"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "RF-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   160
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   240
      Picture         =   "frmPHealthReport.frx":0A2D
      Top             =   240
      Width           =   765
   End
End
Attribute VB_Name = "frmPHealthReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsItems As ADODB.Recordset
Dim StrQuery As String
Dim mPage As Integer
Dim gMpage As Integer
Dim TotEmp As Integer
Dim SumTotal As String
Dim ctr2 As Integer, ctr3 As Integer
Dim LineCtr2 As Integer, LineCtr3 As Integer
Dim cPos2 As Integer
Dim PHGroupCode As String
Dim mTDiff As Double
Dim mTEE As Double
Dim mTER As Double
Dim mECC As Double
Dim WithSBRNo As Boolean

Private Sub cmdAppMonth_Click()
txtAppMonthYear(1).Text = cmdAppMonth.Text & " " & cmdAppYear.Text
End Sub

Private Sub cmdAppMonth_LostFocus()
txtAppMonthYear(1).Text = cmdAppMonth.Text & " " & cmdAppYear.Text
End Sub

Private Sub cmdAppYear_Click()
txtAppMonthYear(1).Text = cmdAppMonth.Text & " " & cmdAppYear.Text
End Sub

Private Sub cmdAppYear_LostFocus()
txtAppMonthYear(1).Text = cmdAppMonth.Text & " " & cmdAppYear.Text
End Sub

Private Sub cmdCheckZero_Click()
Dim mZeroDed As Double

ZeroDed = 0
If cmdCheckZero.Caption = "Check Zero Contributions" Then
    For i = 1 To ListView1(1).ListItems.Count
        'mZeroDed = Val(ListView1(1).ListItems.Item(i).SubItems(7)) ' get employee share
        mZeroDed = Val(ListView1(1).ListItems.Item(i).SubItems(9)) ' get adjusted employee share
        If mZeroDed = 0# Then
            'If ListView1(1).ListItems.Item(i).Checked = True Then
            '    ListView1(1).ListItems.Item(i).Checked = False
            'Else
                ListView1(1).ListItems.Item(i).Checked = True
                ListView1(1).ListItems.Item(i).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(1).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(2).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(3).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(4).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(5).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(6).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(7).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(8).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(9).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(10).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(11).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(12).ForeColor = vbRed
                ListView1(1).ListItems.Item(i).ListSubItems(13).ForeColor = vbRed
                'ListView1(1).ListItems.Item(i).ListSubItems(14).ForeColor = vbRed
            'End If
            ZeroDed = ZeroDed + 1
        End If
    Next i
    cmdCheckZero.Caption = "UnCheck Zero Contributions"
    lblclient(2).Caption = "Check Zero Contributions: " & ZeroDed
Else    ' Uncheck Zero Contributions
    For i = 1 To ListView1(1).ListItems.Count
        'mZeroDed = Val(ListView1(1).ListItems.Item(i).SubItems(7)) ' get employee share
        mZeroDed = Val(ListView1(1).ListItems.Item(i).SubItems(9)) ' get adjusted employee share
        If mZeroDed = 0# Then
            'If ListView1(1).ListItems.Item(i).Checked = True Then
                ListView1(1).ListItems.Item(i).Checked = False
                ListView1(1).ListItems.Item(i).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(1).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(2).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(3).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(4).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(5).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(6).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(7).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(8).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(9).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(10).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(11).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(12).ForeColor = vbBlack
                ListView1(1).ListItems.Item(i).ListSubItems(13).ForeColor = vbBlack
                'ListView1(1).ListItems.Item(i).ListSubItems(14).ForeColor = vbBlack
            'Else
            '    ListView1(1).ListItems.Item(i).Checked = True
            'End If
            ZeroDed = ZeroDed + 1
        End If
    Next i
    cmdCheckZero.Caption = "Check Zero Contributions"
    lblclient(2).Caption = "Uncheck Zero Contributions: " & ZeroDed
End If

End Sub

Private Sub cmdClearView_Click()
    ListView1(1).ListItems.Clear
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCompany_Click()
    ReportedFlag = 1
    SelectCompanyFlag = True
    frmCompany.Show
    frmCompany.SetFocus
End Sub

Private Sub cmdER2_Click()

If ListView1(1).ListItems.Count >= 1 Then
    Call GenerateER2Reports
Else
    MsgBox "No records to processed.", vbInformation, PayrollAlert
End If

End Sub

Private Sub cmdExcelReport_Click()
ListView1(1).SortKey = 0
ListView1(1).Sorted = True
ListView1(1).Refresh
If ListView1(1).ListItems.Count >= 1 Then
    Call GeneratePHealthReports
Else
    MsgBox "No records to process.", vbInformation, PayrollAlert
End If
ListView1(1).SortKey = 1
ListView1(1).Sorted = True
ListView1(1).Refresh
End Sub

Private Sub cmdGenerateRF1Report_Click()
'Dim X As Printer
'Dim y As PrinterOrientationConstants

If cmdGenerateRF1Report.Caption = "RF-1 Report" Then
    If ListView1(1).ListItems.Count >= 1 Then
        Call GenerateRF1Reports
        cmdGenerateRF1Report.Caption = "Copy to Disk"
    Else
        MsgBox "No records to processed.", vbInformation, PayrollAlert
    End If
Else
    ' Copy to disckette.
    Call CopyToDisk
    cmdGenerateRF1Report.Caption = "RF-1 Report"
End If

'strQuery = "SELECT * from clienttable WHERE clientstatus = true and clientname is not null ORDER BY clientname"
'Set rsItems = ReturnRst(strQuery, cnnStr, mstatic)
'
'For Each X In Printers
'   If X.Orientation = vbPRORPortrait Then
'      ' Set printer as system default.
'      Set Printer = X
'      ' Stop looking for a printer.
'      Exit For
'   End If
'Next
'
'Printer.Orientation = vbPRORLandscape
'
'    With rptPhilHealth
'        Set .DataSource = rsItems
'            .DataMember = rsItems.DataMember
'            .ReportWidth = 9090
'
'
'
'        .Show
'    End With
    
End Sub

Private Sub CopyToDisk()
Dim fs As New FileSystemObject
Dim mMsg As String
Dim ExcelFile As String
Dim ToDriveA As String

On Error GoTo ExitRtn2

CommonDialog1.Filter = "Excel Files (*.xls)"
CommonDialog1.Flags = cdlOFNFileMustExist
CommonDialog1.ShowOpen
ExcelFile = CommonDialog1.Filename

mMsg = "Copy " & ExcelFile & " to diskette? Insert diskette in Drive A: then select YES."
'If MsgBox(mMsg, vbQuestion + vbYesNo) = vbYes Then
    If MsgBox(mMsg, vbQuestion + vbYesNo) = vbYes Then
        If fs.FileExists(ExcelFile) Then
            ToDriveA = "a" & mID(ExcelFile, 2, Len(ExcelFile) - 1)
            On Error GoTo ExitRtn2
            fs.CopyFile ExcelFile, ToDriveA, True
            MsgBox ExcelFile & " was successfully copied to disk.", vbInformation, PayrollAlert
        Else
            MsgBox "Copy cancelled!", vbInformation, PayrollAlert
        End If
    End If
'Else
'    MsgBox "Copy cancelled.", vbInformation, PayrollAlert
'End If

ExitRtn1:

Exit Sub

ExitRtn2:

    MsgBox Err.Description & ". Cannot copy file.", vbInformation, PayrollAlert
    Err.Clear
        
End Sub


Private Sub GenerateRF1Reports()
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlBook2 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim ctr As Long
Dim i As Integer, j As Integer, k As Integer, l As Integer, m As Integer
Dim cPos As Integer
Dim prevClientCode As Integer
Dim prevControlNo As String
Dim prevCutOff As Variant
Dim StartOfLine As Integer
Dim EndOfLine As Integer
Dim ExcelTemplate As String
Dim mRange As String
Dim mRange1 As String
Dim mRange2 As String
Dim mRange3 As String
Dim HeaderFlag As Boolean
Dim MemberCtr As Integer
Dim EETotal As Double
Dim ERTotal As Double
Dim GEETotal As Double
Dim GERTotal As Double

ExcelTemplate = App.Path & "\Reports\RF-1 Template.xls"
'ExcelTemplate = App.Path & "\Reports\RF-1 New Template.xls"

Set xlApp1 = New Excel.Application
'Set xlBook1 = xlApp1.Workbooks.Open(ExcelTemplate)
Set xlBook1 = xlApp1.Workbooks.Add(ExcelTemplate)

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

'Set xlSheet = book1.Worksheets.Add
Set xlSheet = xlBook1.Worksheets(1)
'Set xlSheet.Name = "PayrollBankCopy"

cPos = 5
ctr = 7
TotalLoan = 0: LoanStart = 0
HeaderFlag = True
mPage = 1

'On Error GoTo Errh
' Print Header
'xlSheet.Range("E10") = mSapi  '" (SAPI)"

'xlSheet.Range("A2") = ReportHeader & " " & IIf(chkLoans.VALUE = 1, "(for Loans)", "") 'mSapi

'xlSheet.Range("A3") = ScopeRep & " " & cmbMonthof & " " & cmbYearof

ctr = ctr + 1
Call getTotalPages(15)
i = 1
j = 5 '18
MemberCtr = 1
With xlSheet
Do While ListView1(1).ListItems.Count >= i

If ListView1(1).ListItems.Item(i).Checked = True Then
    GoTo Next1
Else
    If HeaderFlag = True Then
        l = 1
        For k = Asc("C") To Asc("P") ' Philhealth no.
            If (Chr(k) <> "E" And Chr(k) <> "O") Then
                mRange = Chr(k) & j
                If mID(txtPhilHealthNo, l, 1) = "-" Then
                    l = l + 1
                End If
                .Range(mRange) = mID(txtPhilHealthNo, l, 1)
                l = l + 1
            End If
        Next
        j = j + 2
        l = 1
        For k = Asc("C") To Asc("Q") ' TIN no.
            If (Chr(k) <> "F" And Chr(k) <> "J" And Chr(k) <> "N") Then
                mRange = Chr(k) & j
                If mID(txtEmpTIN, l, 1) = "-" Then
                    l = l + 1
                End If
                .Range(mRange) = mID(txtEmpTIN, l, 1)
                l = l + 1
            End If
        Next
        j = j + 3
        mRange = "E" & j
        xlSheet.Range(mRange) = txtEmployerName
        j = j + 2
        mRange = "E" & j
        xlSheet.Range(mRange) = txtMailingAddress
        j = j + 2
        mRange = "E" & j
        xlSheet.Range(mRange) = txtTelephoneNo
        j = j + 4
        HeaderFlag = False
        MemberCtr = 0
        
        mRange1 = "AP" & j + 17 ' Signature 35
        .Range(mRange1) = txtPrintedName.Text
        mRange1 = "AP" & j + 18 ' Designation 36
        .Range(mRange1) = txtDesignation.Text
        
        mRange1 = "A" & j + 19 ' Applicable Date
        .Range(mRange1) = txtAppMonthYear(1).Text
        mRange1 = "C" & j + 19 ' Remitted Amount
        .Range(mRange1) = txtRemittedAmt.Text
        mRange1 = "H" & j + 19 ' Acknowledgement Receipt
        .Range(mRange1) = txtSBRNo.Text 'txtReceiptNo.Text
        mRange1 = "M" & j + 19 ' Transaction date
        .Range(mRange1) = MaskDatePaid.Text 'txtTransDate.Text
        mRange1 = "Q" & j + 19 ' No of Employees
        .Range(mRange1) = txtNoOfEmp.Text
    
    End If
    
    MemberCtr = MemberCtr + 1
    
'        litem1.SubItems(1) = IIf(IsNull(!Surname), "", !Surname)
'        litem1.SubItems(2) = IIf(IsNull(!Firstname), "", !Firstname)
'        litem1.SubItems(3) = IIf(IsNull(!Middlename), "", !Middlename)
'        litem1.SubItems(4) = IIf(IsNull(!PHealthSSSNo), "", !PHealthSSSNo)
'        litem1.SubItems(5) = Format(!BasicPay, "0.##")
'        litem1.SubItems(6) = !SalaryBracket
'        litem1.SubItems(7) = Format(!ActualEmployeeDed, "0.##")
'        litem1.SubItems(8) = Format(!ActualEmployerDed, "0.##")
'        litem1.SubItems(9) = Format(!AdjustedEmployeeDed, "0.##")
'        litem1.SubItems(10) = Format(!AdjustedEmployerDed, "0.##")
'        litem1.SubItems(11) = Format(!AmtDifference, "0.##")
    
    mRange1 = "B" & j
    mRange2 = "E" & j
    mRange3 = "M" & j
    
    .Range(mRange1) = ListView1(1).ListItems(i).SubItems(1) ' Surname
    .Range(mRange2) = ListView1(1).ListItems(i).SubItems(2) ' Firstname
    .Range(mRange3) = ListView1(1).ListItems(i).SubItems(3) ' Middlename
        
    l = 1
    For k = Asc("U") To Asc("Z")
        If Chr(k) <> "W" Then
            mRange = Chr(k) & j
            If mID(ListView1(1).ListItems(i).SubItems(4), l, 1) = "-" Then
                l = l + 1
            End If
            .Range(mRange) = mID(ListView1(1).ListItems(i).SubItems(4), l, 1)
            l = l + 1
        End If
    Next
    For k = Asc("A") To Asc("H")
        If Chr(k) <> "G" Then
            mRange = "A" & Chr(k) & j
            If mID(ListView1(1).ListItems(i).SubItems(4), l, 1) = "-" Then
                l = l + 1
            End If
            .Range(mRange) = mID(ListView1(1).ListItems(i).SubItems(4), l, 1)
            l = l + 1
        End If
    Next
    
    mRange1 = "AJ" & j
    .Range(mRange1) = ListView1(1).ListItems(i).SubItems(6) ' bracket
    mRange1 = "AM" & j
    .Range(mRange1) = ListView1(1).ListItems(i).SubItems(9) ' adjusted ee
    EETotal = EETotal + Val(ListView1(1).ListItems(i).SubItems(9))
    mRange1 = "AN" & j
    .Range(mRange1) = ListView1(1).ListItems(i).SubItems(10) ' adjusted er
    ERTotal = ERTotal + Val(ListView1(1).ListItems(i).SubItems(10))
    GEETotal = GEETotal + Val(ListView1(1).ListItems(i).SubItems(9))
    GERTotal = GERTotal + Val(ListView1(1).ListItems(i).SubItems(10))
    j = j + 1
    
    If MemberCtr = 15 Then
        ' Print Sub Total and Grand Total.
        mRange1 = "AM" & j + 1
        .Range(mRange1) = EETotal ' adjusted ee
        mRange1 = "AN" & j + 1
        .Range(mRange1) = ERTotal ' adjusted er
        mRange1 = "AM" & j + 2
        .Range(mRange1) = EETotal + ERTotal ' adjusted ee
        
        mRange1 = "AM" & j + 4
        .Range(mRange1) = GEETotal ' adjusted ee
        mRange1 = "AN" & j + 4
        .Range(mRange1) = GERTotal ' adjusted er
        mRange1 = "AM" & j + 5
        .Range(mRange1) = GEETotal + GERTotal ' adjusted ee
        EETotal = 0
        ERTotal = 0
        ' Print number of pages.
        mRange1 = "AP" & j + 6
        .Range(mRange1) = "        Pages " & mPage & " of " & gMpage
        
        HeaderFlag = True
        xlApp1.ActiveSheet.Range("A1:AP39").Select
        xlApp1.Selection.Copy
        If mPage > 1 Then
            j = j + 7
        Else
            j = j + 7
        End If
        mRange1 = "A" & j
        xlApp1.ActiveSheet.Range(mRange1).Select
        xlApp1.ActiveSheet.Paste
        xlApp1.ActiveSheet.Range(mRange1).Select
                
        ' Adjusted the RowHeight.
        mRange1 = j + 2 & ":" & j + 2
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 2.25
        mRange1 = j + 7 & ":" & j + 7
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 5.25
        mRange1 = j + 14 & ":" & j + 14
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 5.25
        mRange1 = j + 15 & ":" & j + 15
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 19.5
        mRange1 = j + 16 & ":" & j + 16
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 26.25


'    Rows("43:43").RowHeight = 5.25
'    Rows("47:47").RowHeight = 5.25
'    ActiveWindow.ScrollRow = 33
'    Rows("48:48").RowHeight = 6.75
'    Rows("49:49").RowHeight = 9
'    Rows("54:54").RowHeight = 4.5
'    Rows("56:56").RowHeight = 21
'    Rows("56:56").RowHeight = 24.75
'    ActiveWindow.SmallScroll Down:=9
'    Rows("72:72").RowHeight = 3.75
'    ActiveWindow.SelectedSheets.PrintPreview
'    ActiveWindow.ScrollRow = 39
'    Rows("54:54").RowHeight = 9.75
'    Rows("72:72").RowHeight = 11.25
'    Rows("48:48").RowHeight = 10.5
'    Rows("48:48").RowHeight = 8.25
        
        
        
        
        
        mPage = mPage + 1
        j = j + 4
    End If
    End If

Next1:
    
    i = i + 1

Loop

'Print the last page.
    If MemberCtr <> 15 Then
        For m = MemberCtr To 14
            mRange1 = "B" & j
            mRange2 = "E" & j
            mRange3 = "M" & j
    
            .Range(mRange1) = "" ' Surname
            .Range(mRange2) = "" ' Firstname
            .Range(mRange3) = "" ' Middlename
            
            l = 1
            For k = Asc("U") To Asc("Z")
                If Chr(k) <> "W" Then
                    mRange = Chr(k) & j
                    .Range(mRange) = ""
                    l = l + 1
                End If
            Next
            For k = Asc("A") To Asc("H")
                If Chr(k) <> "G" Then
                    mRange = "A" & Chr(k) & j
                    .Range(mRange) = ""
                    l = l + 1
                End If
            Next
        
            mRange1 = "AJ" & j
            .Range(mRange1) = "-" ' bracket
            mRange1 = "AM" & j
            .Range(mRange1) = "-" ' adjusted ee
            mRange1 = "AN" & j
            .Range(mRange1) = "-" ' adjusted er
            j = j + 1
        Next
        
        ' Print Sub Total and Grand Total.
        mRange1 = "AM" & j + 1
        .Range(mRange1) = EETotal ' adjusted ee
        mRange1 = "AN" & j + 1
        .Range(mRange1) = ERTotal ' adjusted er
        mRange1 = "AM" & j + 2
        .Range(mRange1) = EETotal + ERTotal ' adjusted ee
        
        'GEETotal = GEETotal + EETotal
        'GERTotal = GERTotal + ERTotal
        
        mRange1 = "AM" & j + 4
        .Range(mRange1) = GEETotal ' adjusted ee
        mRange1 = "AN" & j + 4
        .Range(mRange1) = GERTotal ' adjusted er
        mRange1 = "AM" & j + 5
        .Range(mRange1) = GEETotal + GERTotal ' adjusted ee

        ' Print number of pages.
        mRange1 = "AP" & j + 6
        .Range(mRange1) = "Pages " & mPage & " of " & gMpage
    End If

End With
mRange1 = "A1"
xlSheet.Range(mRange1).Select

End Sub

Private Sub GenerateER2Reports()
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlBook2 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim ctr As Long
Dim i As Integer, j As Integer, k As Integer, l As Integer, m As Integer
Dim cPos As Integer
Dim prevClientCode As Integer
Dim prevControlNo As String
Dim prevCutOff As Variant
Dim StartOfLine As Integer
Dim EndOfLine As Integer
Dim ExcelTemplate As String
Dim mRange As String
Dim mRange1 As String
Dim mRange2 As String
Dim mRange3 As String
Dim HeaderFlag As Boolean
Dim MemberCtr As Integer
Dim EETotal As Double
Dim ERTotal As Double
Dim GEETotal As Double
Dim GERTotal As Double

ExcelTemplate = App.Path & "\Reports\ER2 Template.xls"

Set xlApp1 = New Excel.Application
'Set xlBook1 = xlApp1.Workbooks.Open(ExcelTemplate)
Set xlBook1 = xlApp1.Workbooks.Add(ExcelTemplate)

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

'Set xlSheet = book1.Worksheets.Add
Set xlSheet = xlBook1.Worksheets(1)
'Set xlSheet.Name = "PayrollBankCopy"

cPos = 5
ctr = 7
TotalLoan = 0: LoanStart = 0
HeaderFlag = True
mPage = 1
ctr = ctr + 1
Call getTotalPages(30)
i = 1
j = 1 '18
MemberCtr = 1

With xlSheet

Do While ListView1(1).ListItems.Count >= i

If ListView1(1).ListItems.Item(i).Checked = True Then
    GoTo Next1
Else
    If HeaderFlag = True Then
'        l = 1
'        For k = Asc("C") To Asc("P") ' Philhealth no.
'            If (Chr(k) <> "E" And Chr(k) <> "O") Then
'                mRange = Chr(k) & j
'                If Mid(txtPhilHealthNo, l, 1) = "-" Then
'                    l = l + 1
'                End If
'                .Range(mRange) = Mid(txtPhilHealthNo, l, 1)
'                l = l + 1
'            End If
'        Next
'        j = j + 2
'        l = 1
'        For k = Asc("C") To Asc("Q") ' TIN no.
'            If (Chr(k) <> "F" And Chr(k) <> "J" And Chr(k) <> "N") Then
'                mRange = Chr(k) & j
'                If Mid(txtEmpTIN, l, 1) = "-" Then
'                    l = l + 1
'                End If
'                .Range(mRange) = Mid(txtEmpTIN, l, 1)
'                l = l + 1
'            End If
'        Next
        
        j = j + 3
        mRange = "C" & j
        xlSheet.Range(mRange) = txtEmployerName
        mRange = "H" & j
        xlSheet.Range(mRange) = txtPhilHealthNo
        
        j = j + 1
        mRange = "C" & j
        xlSheet.Range(mRange) = txtMailingAddress

'        j = j + 2
'        mRange = "E" & j
'        xlSheet.Range(mRange) = txtTelephoneNo
        
        j = j + 2   ' start of detail line
        HeaderFlag = False
        MemberCtr = 0
        
        mRange1 = "F" & j + 31 ' Signature 37
        .Range(mRange1) = txtPrintedName.Text
'        mRange1 = "AP" & j + 18 ' Designation 36
'        .Range(mRange1) = txtDesignation.Text
        
'        mRange1 = "A" & j + 19 ' Applicable Date
'        .Range(mRange1) = txtAppMonthYear(1).Text
'        mRange1 = "C" & j + 19 ' Remitted Amount
'        .Range(mRange1) = txtRemittedAmt.Text
'        mRange1 = "H" & j + 19 ' Acknowledgement Receipt
'        .Range(mRange1) = txtSBRNo.Text 'txtReceiptNo.Text
'        mRange1 = "M" & j + 19 ' Transaction date
'        .Range(mRange1) = MaskDatePaid.Text 'txtTransDate.Text
'        mRange1 = "Q" & j + 19 ' No of Employees
'        .Range(mRange1) = txtNoOfEmp.Text
    
    End If
    
    MemberCtr = MemberCtr + 1
    
'        litem1.SubItems(1) = IIf(IsNull(!Surname), "", !Surname)
'        litem1.SubItems(2) = IIf(IsNull(!Firstname), "", !Firstname)
'        litem1.SubItems(3) = IIf(IsNull(!Middlename), "", !Middlename)
'        litem1.SubItems(4) = IIf(IsNull(!PHealthSSSNo), "", !PHealthSSSNo)
'        litem1.SubItems(5) = Format(!BasicPay, "0.##")
'        litem1.SubItems(6) = !SalaryBracket
'        litem1.SubItems(7) = Format(!ActualEmployeeDed, "0.##")
'        litem1.SubItems(8) = Format(!ActualEmployerDed, "0.##")
'        litem1.SubItems(9) = Format(!AdjustedEmployeeDed, "0.##")
'        litem1.SubItems(10) = Format(!AdjustedEmployerDed, "0.##")
'        litem1.SubItems(11) = Format(!AmtDifference, "0.##")
    
    mRange1 = "B" & j
    mRange2 = "C" & j
    
    ' Philhealth No
    .Range(mRange1) = ListView1(1).ListItems(i).SubItems(4)
    ' Employee Name
    .Range(mRange2) = ListView1(1).ListItems(i).SubItems(1) & ", " & _
        ListView1(1).ListItems(i).SubItems(2) & " " & _
        IIf(ListView1(1).ListItems(i).SubItems(3) <> "", Left(ListView1(1).ListItems(i).SubItems(3), 1) & ".", "")
    ' Position
    mRange1 = "D" & j
    .Range(mRange1) = ListView1(1).ListItems(i).SubItems(22)
    ' Salary
    mRange1 = "E" & j
    .Range(mRange1) = ListView1(1).ListItems(i).SubItems(5)
    ' Employment Date
    mRange1 = "F" & j
    .Range(mRange1) = IIf(ListView1(1).ListItems(i).SubItems(21) = "", "__/__/____", Format(ListView1(1).ListItems(i).SubItems(21), "mm/dd/yyyy"))       ' get start contract
        
    j = j + 1
    
    If MemberCtr = 30 Then
        ' Print Total Employees
        mRange1 = "C" & j
        .Range(mRange1) = MemberCtr ' Total employees
        ' Print number of pages.
        mRange1 = "D" & j + 2
        .Range(mRange1) = "PAGE " & mPage & " OF " & gMpage & " SHEETS"
        
        HeaderFlag = True
        xlApp1.ActiveSheet.Range("A1:H39").Select
        xlApp1.Selection.Copy
        
        j = j + 3
        
        mRange1 = "A" & j
        xlApp1.ActiveSheet.Range(mRange1).Select
        xlApp1.ActiveSheet.Paste
        xlApp1.ActiveSheet.Range(mRange1).Select
                
        ' Adjusted the RowHeight.
'        mRange1 = j + 2 & ":" & j + 2
'        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 2.25
'        mRange1 = j + 7 & ":" & j + 7
'        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 5.25
'        mRange1 = j + 14 & ":" & j + 14
'        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 5.25
'        mRange1 = j + 15 & ":" & j + 15
'        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 19.5
'        mRange1 = j + 16 & ":" & j + 16
'        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 26.25

        
        mPage = mPage + 1
        'j = j + 3
    End If
    End If

Next1:
    
    i = i + 1

Loop

'Print the last page.
    If MemberCtr <> 30 Then
        For m = MemberCtr To 29
            ' Philhealth No
            mRange1 = "B" & j
            .Range(mRange1) = ""
            ' Employee Name
            mRange2 = "C" & j
            .Range(mRange2) = ""
            ' Position
            mRange1 = "D" & j
            .Range(mRange1) = ""
            ' Salary
            mRange1 = "E" & j
            .Range(mRange1) = ""
            ' Employment Date
            mRange1 = "F" & j
            .Range(mRange1) = ""
            j = j + 1
        Next
        
        ' Print Total Employees
        mRange1 = "C" & j
        .Range(mRange1) = MemberCtr ' Total employees
        ' Print number of pages.
        mRange1 = "D" & j + 2
        .Range(mRange1) = "PAGE " & mPage & " OF " & gMpage & " SHEETS"
        
    End If

End With

mRange1 = "A1"
xlSheet.Range(mRange1).Select

End Sub

Private Sub getTotalPages(ByVal RecPage As Integer)
Dim i As Integer

i = 1
TotEmp = 0
Do While ListView1(1).ListItems.Count >= i

    If i <> 0 And ListView1(1).ListItems.Item(i).Checked = False Then
        TotEmp = TotEmp + 1
    End If
    i = i + 1
Loop
gMpage = TotEmp \ RecPage
If TotEmp Mod RecPage <> 0 Then
    gMpage = gMpage + 1
End If
If gMpage = 0 Then gMpage = 1
txtNoOfEmp.Text = TotEmp
End Sub

Private Sub cmdNoPHNo_Click()
Dim NoPHNo As Integer

NoPHNo = 0
For i = 1 To ListView1(1).ListItems.Count
    If ListView1(1).ListItems.Item(i).SubItems(4) = "" Then ' get philhealth number
        If ListView1(1).ListItems.Item(i).Checked = True Then
            ListView1(1).ListItems.Item(i).Checked = False
        Else
            ListView1(1).ListItems.Item(i).Checked = True
        End If
        NoPHNo = NoPHNo + 1
    End If
Next i
lblclient(7).Caption = "No PHealth No.: " & NoPHNo

End Sub

Private Sub cmdRetrieve_Click()
Dim litem1 As ListItem
Dim Listctr As Integer
Dim rsPHList As ADODB.Recordset
Dim i As Integer

i = 1
Listctr = 0
If chkAllPayments.VALUE = 1 Then    ' get all payments regardless of users
    'strQuery = "SELECT * FROM PHContributions WHERE ApplicableMonth = '" & cmdAppMonth.Text & "' and ApplicableYear = " & cmdAppYear.Text & " AND ProcessedBy = '" & mUserID & "' AND GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "' ORDER BY Surname, Firstname"
    StrQuery = "SELECT * FROM PHContributions WHERE GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "' ORDER BY Surname, Firstname"
Else    ' get all payments for current user only
    StrQuery = "SELECT * FROM PHContributions WHERE ProcessedBy = '" & mUserID & "' AND GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "' ORDER BY Surname, Firstname"
End If
Set rsPHList = ReturnRst(StrQuery, cnnstr, mstatic)
If rsPHList.RecordCount = 0 Then
    MsgBox "No previous contributions retrieve.", vbInformation, PayrollAlert
    lblclient(0).Caption = "Difference: 0.00"
    lblclient(3).Caption = "Total EE: 0.00"
    lblclient(1).Caption = "Total ER: 0.00"
    txtRemittedAmt.Text = "0.00"
    GoTo Exit22
End If

WithSBRNo = False
With rsPHList
ListView1(1).ListItems.Clear
While Not (.EOF Or .BOF)
    Set litem1 = frmPHealthReport.ListView1(1).ListItems.Add(, , IIf(IsNull(!ClientName), "Blank", !ClientName))
        litem1.SubItems(1) = IIf(IsNull(!Surname), "", !Surname)
        litem1.SubItems(2) = IIf(IsNull(!FirstName), "", !FirstName)
        litem1.SubItems(3) = IIf(IsNull(!MiddleName), "", !MiddleName)
        litem1.SubItems(4) = IIf(IsNull(!PHealthSSSNo), "", !PHealthSSSNo)
        litem1.SubItems(5) = Format(!BasicPay, "0.##")
        litem1.SubItems(6) = IIf(IsNull(!SalaryBracket), "", !SalaryBracket)
        litem1.SubItems(7) = Format(!ActualEmployeeDed, "0.##")
        litem1.SubItems(8) = Format(!ActualEmployerDed, "0.##")
        litem1.SubItems(9) = Format(!AdjustedEmployeeDed, "0.##")
        litem1.SubItems(10) = Format(!AdjustedEmployerDed, "0.##")
        litem1.SubItems(11) = Format(!AmtDifference, "0.##")
        litem1.SubItems(12) = Format(!ClientName, "0.##")      ' please activate
        litem1.SubItems(13) = IIf(IsNull(!ProcessedBy), "", !ProcessedBy)      ' please activate
        litem1.SubItems(14) = IIf(IsNull(!ProcessedDate), "", !ProcessedDate)      ' please activate
        
        litem1.SubItems(16) = IIf(IsNull(!groupcode), "", !groupcode)
        litem1.SubItems(17) = Format(!AmountPaid, "0.00")
        litem1.SubItems(18) = IIf(IsNull(!SBRNo), "", !SBRNo)
        litem1.SubItems(19) = IIf(!DatePaid = #1/1/1900#, "__/__/____", Format(!DatePaid, "mm/dd/yyyy"))
        litem1.SubItems(20) = IIf(!DatePosted = #1/1/1900#, "__/__/____", Format(!DatePosted, "mm/dd/yyyy"))
        
        ' Employee Position.
        litem1.SubItems(21) = IIf(!StartContract = #1/1/1900#, "__/__/____", Format(!StartContract, "mm/dd/yyyy"))      ' get start contract
        litem1.SubItems(22) = IIf(IsNull(!empPosition), "", !empPosition)      ' get employee position
        litem1.SubItems(23) = !EmployeeID      ' get employee number
        
        If Not IsNull(!SBRNo) And !SBRNo <> "" And !SBRNo <> "0" Then
            WithSBRNo = True
        End If
        If !ExcludedFlag = True Then
            ListView1(1).ListItems.Item(i).Checked = True
        End If
    .MoveNext
    i = i + 1
Wend
End With

Set rsPHList = Nothing

Call GetTotalPHealth

Exit22:

End Sub

Private Sub GetTotalPHealth()
Dim mTDiff As Double
Dim mTEE As Double
Dim mTER As Double
Dim i As Integer
Dim ExcludedCtr As Integer

ExcludedCtr = 0
mTDiff = 0: mTEE = 0: mTER = 0: mECC = 0
mTotal = 0#
For i = 1 To ListView1(1).ListItems.Count
    If ListView1(1).ListItems.Item(i).Checked = False Then
        mTEE = mTEE + Val(ListView1(1).ListItems(i).SubItems(9)) ' employee adjust
        mTER = mTER + Val(ListView1(1).ListItems(i).SubItems(10)) ' employer adjust
        mTDiff = mTDiff + Val(ListView1(1).ListItems(i).SubItems(11)) ' difference
    Else
        ExcludedCtr = ExcludedCtr + 1
    End If
Next i
lblclient(0).Caption = "Difference: " & mTDiff
lblclient(3).Caption = "Adjusted PS: " & mTEE
lblclient(1).Caption = "Adjusted ES: " & mTER
txtRemittedAmt.Text = Format(mTEE + mTER, "0.00")

Label2(15).Caption = "Excluded Employees: " & ExcludedCtr
Label2(19).Caption = "Included Employees: " & ListView1(1).ListItems.Count - ExcludedCtr

End Sub

Private Sub ListView1_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
Dim mSurname As String
Dim mFirstName As String
Dim mMiddlename As String

mSurname = ListView1(Index).SelectedItem.ListSubItems(1).Text
mFirstName = ListView1(Index).SelectedItem.ListSubItems(2).Text
mMiddlename = ListView1(Index).SelectedItem.ListSubItems(3).Text

If ListView1(1).SelectedItem.Checked = True Then
    StrQuery = "UPDATE PHContributions SET ExcludedFlag = 1 " & _
    " WHERE Surname = '" & mSurname & "' AND " & _
    " Firstname = '" & mFirstName & "' AND " & _
    " Middlename = '" & mMiddlename & "' AND " & _
    " GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "'"
Else
    StrQuery = "UPDATE PHContributions SET ExcludedFlag = 0 " & _
    " WHERE Surname = '" & mSurname & "' AND " & _
    " Firstname = '" & mFirstName & "' AND " & _
    " Middlename = '" & mMiddlename & "' AND " & _
    " GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "'"
End If
Call Action_Query(StrQuery, cnnstr)

Call GetTotalPHealth
txtRemittedAmt.Text = Format(mTEE + mTER, "0.00")

End Sub

Private Sub cmdSearchSBR_Click()

Dim rsPHList As ADODB.Recordset

If chkAllPayments.VALUE = 1 Then ' all payments regardless of users
    StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy, SBRNo " & _
        " FROM PHContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy, SBRNo HAVING SBRNo = '" & txtSBRNo.Text & "' AND GroupCode <> ''" & _
        " ORDER BY ProcessedBy, GroupCode"
Else    ' for current login only
    StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy, SBRNo " & _
        " FROM PHContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy, SBRNo HAVING SBRNo = '" & txtSBRNo.Text & "' AND GroupCode <> ''" & _
        " AND ProcessedBy = '" & mUserID & "' ORDER BY ProcessedBy, GroupCode"
End If
Set rsPHList = ReturnRst(StrQuery, cnnstr, mstatic)
If rsPHList.RecordCount = 0 Then
    MsgBox "No payments found with SBR# " & txtSBRNo.Text, vbInformation, PayrollAlert
    rsPHList.Close
Else
    PHGroupCode = UCase(Left(mUserID, 5)) & "-" & Left(cmdAppMonth, 3) & cmdAppYear & "-01 New"
    Call FillGroupCode(cmbGroupCode, rsPHList, PHGroupCode)
    cmbGroupCode.ListIndex = 0
    rsPHList.Close
    Call cmdRetrieve_Click
    MsgBox "SBR# " & txtSBRNo.Text & " found!", vbInformation, PayrollAlert
End If

End Sub

Private Sub cmdSelect_Click()

If WithSBRNo = False Then
    'Call cmdRetrieve_Click
    PHealthFlag = True
    Call frmPHealthClients.Show
    'frmPHealthClients.SetFocus
    'Call FrmOtherReport.Show
    'FrmOtherReport.SetFocus
    Call GetTotalPHealth
Else
    MsgBox "This group of clients was already paid. Cannot update!", vbInformation, PayrollAlert
End If

End Sub

Private Sub cmdSelectGroupCodes_Click()
    Call UpdateGroupCodes
End Sub

Private Sub UpdateGroupCodes()
Dim rsPHList As ADODB.Recordset

If chkAllPayments.VALUE = 1 Then ' all payments regardless of users
    StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy " & _
        " FROM PHContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy HAVING ApplicableMonth = '" & cmdAppMonth.Text & "' and ApplicableYear = " & cmdAppYear.Text & " AND GroupCode <> ''" & _
        " ORDER BY ProcessedBy, GroupCode"
Else    ' for current login only
    StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy " & _
        " FROM PHContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy HAVING ApplicableMonth = '" & cmdAppMonth.Text & "' and ApplicableYear = " & cmdAppYear.Text & " AND GroupCode <> ''" & _
        " AND ProcessedBy = '" & mUserID & "' ORDER BY ProcessedBy, GroupCode"
End If
Set rsPHList = ReturnRst(StrQuery, cnnstr, mstatic)
PHGroupCode = UCase(Left(mUserID, 5)) & "-" & Left(cmdAppMonth, 3) & cmdAppYear & "-01 New"
Call FillGroupCode(cmbGroupCode, rsPHList, PHGroupCode)
cmbGroupCode.ListIndex = 0
rsPHList.Close

End Sub

Private Sub cmdUpdateSBR_Click()

If MsgBox("Update SBR#/OR# and other information?", vbQuestion + vbYesNo, "SBR Number Updates") = vbYes Then
    StrQuery = "UPDATE PHContributions SET SBRNO = '" & IIf(txtSBRNo.Text = "", "0", txtSBRNo.Text) & "', AmountPaid = " & IIf(txtRemittedAmt.Text = "", 0#, txtRemittedAmt.Text) & ", " & _
        " DatePaid = '" & IIf(MaskDatePaid.Text = "__/__/____", #1/1/1900#, Format(MaskDatePaid.Text, "mm/dd/yyyy")) & "', DatePosted = '" & IIf(MaskDatePosted.Text = "__/__/____", #1/1/1900#, Format(MaskDatePosted.Text, "mm/dd/yyyy")) & _
        "' WHERE GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "'"
    If Action_Query(StrQuery, cnnstr) = True Then
        Call cmdRetrieve_Click
        MsgBox "SBR/OR # and other information successfully updated!", vbInformation, PayrollAlert
    Else
        MsgBox "SBR/OR # and other information was not successfully updated!", vbInformation, PayrollAlert
    End If
End If

End Sub

Private Function FillGroupCode(ByRef mCmb As ComboBox, _
                        mRs As Recordset, Optional FirstGroup As String = "")
    mCmb.Clear
    While Not (mRs.EOF Or mRs.BOF)
        mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields("GroupCode")), FirstGroup, mRs.Fields("GroupCode")), "'", "''"))
        'mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
        'FirstGroup = UCase(Left(mRs.Fields("GroupCode"), 13)) & "-" & Right("00" & Right(LTrim(RTrim(mRs.Fields("GroupCode")))) + 1, 2) & " New"
        FirstGroup = UCase(Left(mRs.Fields("GroupCode"), 13)) & "-" & Right("00" & Right(LTrim(RTrim(mRs.Fields("GroupCode"))), 2) + 1, 2) & " New"
        mRs.MoveNext
    Wend
    mCmb.AddItem FirstGroup
    
End Function

Private Sub Form_Activate()
    SelectCompanyFlag = False
    PHealthFlag = True
    Call UpdateGroupCodes
End Sub

Private Sub Form_Load()

    PHealthFlag = True
    SelectCompanyFlag = False
    cmdAppMonth.Text = Format(Now, "mmmm")
    cmdAppYear.Text = Format(Now, "yyyy")
    FormLoadFlag = True
    MaskDatePaid = Format(Now, "mm/dd/yyyy")
    
    If PHealthFlag = False Then  ' for SSS
        Image1.Visible = False
        Label1(0).Caption = "SSS"
        Label1(2).Caption = "   SSS NO.    "
        ListView1(1).Visible = False
        ListView1(0).Visible = True
        lblclient(7).Caption = "No SSS No."
        cmdNoPHNo.Caption = "Check/Uncheck No SSS No."
        cmdGenerateReport.Visible = False
    Else        ' for PhilHealth
        Image1.Visible = True
        Label1(0).Caption = "RF-1"
        Label1(2).Caption = "PHILHEALTH NO."
        ListView1(1).Visible = True
        ListView1(0).Visible = False
        lblclient(7).Caption = "No PHealth No."
        cmdNoPHNo.Caption = "Check/Uncheck No PHealth No."
        cmdGenerateRF1Report.Visible = True
    End If
End Sub

Private Function PrintNoPHealthNumbers(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal mPCompany As String, ByVal ctr As Integer) As Integer

Dim j As Integer
Dim mRange As String, mRange1 As String
Dim LineCtr As Integer
Dim cPos As Integer
Dim mCo2 As String

LineCtr = 1
j = 1
ctr = ctr + 2
cPos = ctr + 1

mRange = "B" & ctr & ":D" & ctr
mCo2 = mPCompany
mPCompany = mPCompany & " (Employees with no PHealth#)"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mPCompany, "", xlLeft, , False, False, False, , , 9, True, 18)

Do While ListView1(1).ListItems.Count >= j
    If Trim(ListView1(1).ListItems(j).SubItems(12)) = Trim(mCo2) And ListView1(1).ListItems.Item(j).Checked = False And _
    (ListView1(1).ListItems(j).SubItems(4) = "__-_______-_" Or ListView1(1).ListItems(j).SubItems(4) = "") Then
        ctr = ctr + 1
        xlSheet.Cells(ctr, 1).VALUE = LineCtr
        xlSheet.Cells(ctr, 2).VALUE = ListView1(1).ListItems(j).SubItems(1) ' Surname
        xlSheet.Cells(ctr, 3).VALUE = ListView1(1).ListItems(j).SubItems(2) ' firstname
        xlSheet.Cells(ctr, 4).VALUE = ListView1(1).ListItems(j).SubItems(3) ' Middlename
        xlSheet.Cells(ctr, 5).VALUE = ListView1(1).ListItems(j).SubItems(4) ' sss no
        xlSheet.Cells(ctr, 6).VALUE = ListView1(1).ListItems(j).SubItems(7) ' employee
        xlSheet.Cells(ctr, 7).VALUE = ListView1(1).ListItems(j).SubItems(8) ' employer
        xlSheet.Cells(ctr, 8).VALUE = Val(ListView1(1).ListItems(j).SubItems(7)) + Val(ListView1(1).ListItems(j).SubItems(8)) ' PS + ES
        xlSheet.Cells(ctr, 9).VALUE = ListView1(1).ListItems(j).SubItems(9) ' employee adjust
        xlSheet.Cells(ctr, 10).VALUE = ListView1(1).ListItems(j).SubItems(10) ' employer adjust
        xlSheet.Cells(ctr, 11).VALUE = Val(ListView1(1).ListItems(j).SubItems(10)) + Val(ListView1(1).ListItems(j).SubItems(9)) ' Total PS + ES
        xlSheet.Cells(ctr, 12).VALUE = Val(ListView1(1).ListItems(j).SubItems(7)) - Val(ListView1(1).ListItems(j).SubItems(9))  ' difference PS
        xlSheet.Cells(ctr, 13).VALUE = Val(ListView1(1).ListItems(j).SubItems(8)) - Val(ListView1(1).ListItems(j).SubItems(10))  ' difference ES
        xlSheet.Cells(ctr, 14).VALUE = (Val(ListView1(1).ListItems(j).SubItems(8)) - Val(ListView1(1).ListItems(j).SubItems(10))) + (Val(ListView1(1).ListItems(j).SubItems(7)) - Val(ListView1(1).ListItems(j).SubItems(9)))
        LineCtr = LineCtr + 1
    End If
    j = j + 1
Loop

' Print Sub-Totals.
ctr = ctr + 1
ctr2 = ctr2 + 1
'mRange = "F" & cPos & ":N" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "F" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

    mRange = "F" & cPos + 1 & ":N" & ctr
    For l = 70 To 78 'F to N
        xlSheet.Range(mRange).Select
        mRange1 = Chr(l) & ctr
        xlSheet.Range(mRange1).Activate
        xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
    
        ' Copy to Sheet 2 the Total.
        xlSheet.Range(mRange1).Select
        xlApp1.Selection.Copy
        xlApp1.Sheets("Sheet2").Select
        mRange1 = Chr(l - 2) & ctr2
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        If l = 70 Then
            mRange1 = "A" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
            mRange1 = "B" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = mPCompany
            mRange1 = "C" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
        End If
        xlApp1.Sheets("Sheet1").Select
    Next l
'
'
'            ' Copy to Sheet 2 the Total.
'            ctr2 = ctr2 + 1
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "D" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            mRange1 = "A" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
'            mRange1 = "B" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.ActiveCell.FormulaR1C1 = mPCompany
'            mRange1 = "C" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
'            xlApp1.Sheets("Sheet1").Select
'
''mRange = "F" & cPos & ":K" & ctr
'mRange = "F" & cPos & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "G" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'            ' Copy to Sheet 2 the Total.
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "E" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            xlApp1.Sheets("Sheet1").Select
'
''mRange = "F" & cPos & ":K" & ctr
'mRange = "F" & cPos & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "H" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'            ' Copy to Sheet 2 the Total.
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "F" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            xlApp1.Sheets("Sheet1").Select
'
''mRange = "F" & cPos & ":K" & ctr
'mRange = "F" & cPos & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "I" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'            ' Copy to Sheet 2 the Total.
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "G" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            xlApp1.Sheets("Sheet1").Select
'
''mRange = "F" & cPos & ":K" & ctr
'mRange = "F" & cPos & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "J" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'            ' Copy to Sheet 2 the Total.
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "H" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            xlApp1.Sheets("Sheet1").Select
'
''mRange = "F" & cPos & ":K" & ctr
'mRange = "F" & cPos & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "K" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'            ' Copy to Sheet 2 the Total.
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "I" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            xlApp1.Sheets("Sheet1").Select
'
''Updated June 2009.
'mRange = "F" & cPos & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "L" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'            ' Copy to Sheet 2 the Total.
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "J" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            xlApp1.Sheets("Sheet1").Select

'mRange = "F" & cPos & ":N" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":N" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

If SumTotal = "" Then
    SumTotal = "=F" & ctr
Else
    SumTotal = SumTotal & "+F" & ctr
End If

PrintNoPHealthNumbers = ctr

End Function


Private Sub GeneratePHealthReports()
Dim rsDedRecords As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mRange As String, mRange1 As String
Dim ctr As Long
Dim i As Integer, l As Integer, t As Integer
Dim cPos As Integer
Dim mHead As String
Dim HeaderFlag As Boolean
Dim NoPHealthCtr As Integer
Dim ExcludedCtr As Integer
Dim LineCtr As Integer
Dim WithRecSw As Boolean
Dim GTotal As Double

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)
'Set xlSheet = xlBook1.Worksheets(2)

SumTotal = ""
NoPHealthCtr = 0: ExcludedCtr = 0
HeaderFlag = False
LineCtr = 1
LineCtr2 = 1
LineCtr3 = 1
ctr = 3
ctr2 = 7 '6    '7
ctr3 = 7 '6
cPos = 9
i = 1: l = 1
WithRecSw = False

'Print Header.
xlApp1.Sheets("Sheet1").Select
mHead = CompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 12, True)
mHead = "PhilHealth Contributions Report for the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)

'Print Header of Sheet2 for Remitted Clients Summary.
mHead = CompanyName & " (" & CompanyShort & ")"
xlApp1.Sheets("Sheet2").Select

xlApp1.Range("A2").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "PhilHealth Contributions Summary Report per Clients"
xlApp1.Range("A3").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "For the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text & "     Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
xlApp1.Range("A4").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead

For t = 65 To 80 ' A to P
    Call PrintHeadings(xlSheet, xlApp1, t, 6)
Next t

'xlApp1.Range("A6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "No."
'xlApp1.Range("B6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Client Name"
'xlApp1.Range("C6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Total Employees"
'xlApp1.Range("D6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "PS Share"
'xlApp1.Range("E6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "ES Share"
'xlApp1.Range("F6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "PS+ES Share"
'xlApp1.Range("G6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Adjusted PS"
'xlApp1.Range("H6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Adjusted ES"
'xlApp1.Range("I6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Total PS+ES"
'xlApp1.Range("J6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Difference PS"
'xlApp1.Range("K6").Select   ' updates
'xlApp1.ActiveCell.FormulaR1C1 = "Difference ES"
'xlApp1.Range("L6").Select   ' updates
'xlApp1.ActiveCell.FormulaR1C1 = "SBRTotal Difference"
'xlApp1.Range("K6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Remarks"

'xlApp1.Range("A6:R6").Select
'xlApp1.Selection.Interior.ColorIndex = 15
'xlApp1.Selection.Interior.Pattern = xlSolid

'Print Header of Sheet3 for Unremitted Clients Summary.
xlApp1.Sheets("Sheet3").Select

mHead = CompanyName & " (" & CompanyShort & ")"
xlApp1.Range("A2").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "PhilHealth Contributions Summary Report (NOT TO BE REMITTED CLIENTS)"
xlApp1.Range("A3").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "For the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text & "     Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
xlApp1.Range("A4").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead

For t = 65 To 80 ' A to P
    Call PrintHeadings(xlSheet, xlApp1, t, 6)
Next t

'xlApp1.Range("A6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "No."
'xlApp1.Range("B6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Client Name"
'xlApp1.Range("C6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Total Employees"
'xlApp1.Range("D6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "PS Share"
'xlApp1.Range("E6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "ES Share"
'xlApp1.Range("F6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "PS+ES Share"
'xlApp1.Range("G6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Adjusted PS"
'xlApp1.Range("H6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Adjusted ES"
'xlApp1.Range("I6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Total PS+ES"
'xlApp1.Range("J6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Difference"
'xlApp1.Range("K6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Remarks"

'xlApp1.Range("A6:R6").Select
'xlApp1.Selection.Interior.ColorIndex = 15
'xlApp1.Selection.Interior.Pattern = xlSolid

xlApp1.Sheets("Sheet1").Select

PrevCompany = ListView1(1).ListItems(i).SubItems(12)
Do While ListView1(1).ListItems.Count >= i

    If PrevCompany <> ListView1(1).ListItems(i).SubItems(12) Then
        If LineCtr > 1 Then
            ' Print Sub-Totals.
            ctr = ctr + 1
            'mRange = "F" & cPos + 1 & ":N" & ctr
            mRange = "F" & cPos & ":N" & ctr
'            xlSheet.Range(mRange).Select
'            mRange1 = "F" & ctr
'            xlSheet.Range(mRange1).Activate
'            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
            
            ' Copy to Sheet 2 the Total.
            ctr2 = ctr2 + 1
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "D" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            mRange1 = "A" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
'            mRange1 = "B" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
'            mRange1 = "C" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
'            xlApp1.Sheets("Sheet1").Select
            
           
            For l = 70 To 78 'F to N
                xlSheet.Range(mRange).Select
                mRange1 = Chr(l) & ctr
                xlSheet.Range(mRange1).Activate
                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
            
                ' Copy to Sheet 2 the Total.
                xlSheet.Range(mRange1).Select
                xlApp1.Selection.Copy
                xlApp1.Sheets("Sheet2").Select
                mRange1 = Chr(l - 2) & ctr2
                xlApp1.Range(mRange1).Select
                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=False
                If l = 70 Then
                    mRange1 = "A" & ctr2
                    xlApp1.Range(mRange1).Select
                    xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
                    mRange1 = "B" & ctr2
                    xlApp1.Range(mRange1).Select
                    xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
                    mRange1 = "C" & ctr2
                    xlApp1.Range(mRange1).Select
                    xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
                End If
                xlApp1.Sheets("Sheet1").Select
            Next l
            
'            xlSheet.Range(mRange).Select
'            mRange1 = "G" & ctr
'            xlSheet.Range(mRange1).Activate
'            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'            ' Copy to Sheet 2 the Total.
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "E" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            xlApp1.Sheets("Sheet1").Select
'
'            xlSheet.Range(mRange).Select
'            mRange1 = "H" & ctr
'            xlSheet.Range(mRange1).Activate
'            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'            ' Copy to Sheet 2 the Total.
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "F" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            xlApp1.Sheets("Sheet1").Select
'
'            xlSheet.Range(mRange).Select
'            mRange1 = "I" & ctr
'            xlSheet.Range(mRange1).Activate
'            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'            ' Copy to Sheet 2 the Total.
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "G" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            xlApp1.Sheets("Sheet1").Select
'
'            xlSheet.Range(mRange).Select
'            mRange1 = "J" & ctr
'            xlSheet.Range(mRange1).Activate
'            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'            ' Copy to Sheet 2 the Total.
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "H" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            xlApp1.Sheets("Sheet1").Select
'
'            xlSheet.Range(mRange).Select
'            mRange1 = "K" & ctr
'            xlSheet.Range(mRange1).Activate
'            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'            ' Copy to Sheet 2 the Total.
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "I" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            xlApp1.Sheets("Sheet1").Select
'
'            'Updated June 2009
'            xlSheet.Range(mRange).Select
'            mRange1 = "L" & ctr
'            xlSheet.Range(mRange1).Activate
'            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'            ' Copy to Sheet 2 the Total.
'            xlSheet.Range(mRange1).Select
'            xlApp1.Selection.Copy
'            xlApp1.Sheets("Sheet2").Select
'            mRange1 = "J" & ctr2
'            xlApp1.Range(mRange1).Select
'            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                False, Transpose:=False
'            xlApp1.Sheets("Sheet1").Select
            
            LineCtr2 = LineCtr2 + 1
            
            'mRange = "F" & cPos & ":L" & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            mRange = "F" & ctr & ":N" & ctr
            xlSheet.Range(mRange).Select
            xlApp1.Range(mRange).Font.Bold = True
            
            If SumTotal = "" Then
                SumTotal = "=F" & ctr
            Else
                SumTotal = SumTotal & "+F" & ctr
            End If
        End If
        
        If NoPHealthCtr > 0 Then
            ctr = PrintNoPHealthNumbers(xlSheet, xlApp1, PrevCompany, ctr)
            LineCtr2 = LineCtr2 + 1
        End If

        'ctr = ctr + 3
        PrevCompany = ListView1(1).ListItems(i).SubItems(12)
'        mRange = "A" & ctr
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
        HeaderFlag = False
        NoPHealthCtr = 0  ' set again to 0 after printing
        WithRecSw = False
    End If
    ' Print Header.
    If HeaderFlag = False Then
        LineCtr = 1
        ctr = ctr + 3
        mRange = "A" & ctr
        Call PrintReportHeader(xlSheet, xlApp1, mRange, PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
        ctr = ctr + 1
        mRange = "A" & ctr & ":A" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "No.", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 4.14)
        mRange = "B" & ctr & ":B" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Surname", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)
        mRange = "C" & ctr & ":C" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Firstname", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
        mRange = "D" & ctr & ":D" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Middlename", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
        mRange = "E" & ctr & ":E" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "PhilHealth#", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 11)
        
        For t = 68 To 80 ' F to P
            Call PrintHeadings(xlSheet, xlApp1, t, ctr, 2)
        Next t
        
'        mRange = "F" & ctr & ":F" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "PS Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'        mRange = "G" & ctr & ":G" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "ES Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'        mRange = "H" & ctr & ":H" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "PS+ES Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
'        mRange = "I" & ctr & ":I" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "PS Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'        mRange = "J" & ctr & ":J" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "ES Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'        mRange = "K" & ctr & ":K" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "Total PS+ES", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'
'        mRange = "L" & ctr & ":L" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'        mRange = "M" & ctr & ":M" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Remarks", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        
        HeaderFlag = True
        ctr = ctr + 1
        cPos = ctr + 1
    End If
    
    If ListView1(1).ListItems.Item(i).Checked = True Then
        ' Count those excluded records.
        ExcludedCtr = ExcludedCtr + 1
        GoTo Next1
    ElseIf PrevCompany = ListView1(1).ListItems(i).SubItems(12) And (ListView1(1).ListItems(i).SubItems(4) = "__-_______-_" Or ListView1(1).ListItems(i).SubItems(4) = "") Then
        ' Count those no sss number.
        NoPHealthCtr = NoPHealthCtr + 1
        GoTo Next1
    Else
        ctr = ctr + 1
        xlSheet.Cells(ctr, 1).VALUE = LineCtr
        xlSheet.Cells(ctr, 2).VALUE = ListView1(1).ListItems(i).SubItems(1) ' Surname
        xlSheet.Cells(ctr, 3).VALUE = ListView1(1).ListItems(i).SubItems(2) ' Firstname
        xlSheet.Cells(ctr, 4).VALUE = ListView1(1).ListItems(i).SubItems(3) ' Middlename
        xlSheet.Cells(ctr, 5).VALUE = ListView1(1).ListItems(i).SubItems(4) ' philhealth no
        xlSheet.Cells(ctr, 6).VALUE = ListView1(1).ListItems(i).SubItems(7) ' employee
        xlSheet.Cells(ctr, 7).VALUE = ListView1(1).ListItems(i).SubItems(8) ' employer
        xlSheet.Cells(ctr, 8).VALUE = Val(ListView1(1).ListItems(i).SubItems(7)) + Val(ListView1(1).ListItems(i).SubItems(8)) ' PS + ES Share
        xlSheet.Cells(ctr, 9).VALUE = ListView1(1).ListItems(i).SubItems(9) ' employee adjust
        xlSheet.Cells(ctr, 10).VALUE = ListView1(1).ListItems(i).SubItems(10) ' employer adjust
        xlSheet.Cells(ctr, 11).VALUE = Val(ListView1(1).ListItems(i).SubItems(10)) + Val(ListView1(1).ListItems(i).SubItems(9))  ' total PS + ES
        xlSheet.Cells(ctr, 12).VALUE = Val(ListView1(1).ListItems(i).SubItems(7)) - Val(ListView1(1).ListItems(i).SubItems(9))  ' difference PS
        xlSheet.Cells(ctr, 13).VALUE = Val(ListView1(1).ListItems(i).SubItems(8)) - Val(ListView1(1).ListItems(i).SubItems(10))  ' difference ES
        xlSheet.Cells(ctr, 14).VALUE = (Val(ListView1(1).ListItems(i).SubItems(8)) - Val(ListView1(1).ListItems(i).SubItems(10))) + (Val(ListView1(1).ListItems(i).SubItems(7)) - Val(ListView1(1).ListItems(i).SubItems(9)))
        LineCtr = LineCtr + 1
        WithRecSw = True
    End If
Next1:
    i = i + 1
Loop   ' End of Main Loop

If WithRecSw = True Then    ' check if there were details
    ' Print Sub-Totals.
    ctr = ctr + 1
    ctr2 = ctr2 + 1
    mRange = "F" & cPos & ":N" & ctr
    For l = 70 To 78 'F to N
        xlSheet.Range(mRange).Select
        mRange1 = Chr(l) & ctr
        xlSheet.Range(mRange1).Activate
        xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
    
        ' Copy to Sheet 2 the Total.
        xlSheet.Range(mRange1).Select
        xlApp1.Selection.Copy
        xlApp1.Sheets("Sheet2").Select
        mRange1 = Chr(l - 2) & ctr2
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        If l = 70 Then
            mRange1 = "A" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
            mRange1 = "B" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
            mRange1 = "C" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
        End If
        xlApp1.Sheets("Sheet1").Select
    Next l
    
'
'
'    xlSheet.Range(mRange).Select
'    mRange1 = "F" & ctr
'    xlSheet.Range(mRange1).Activate
'    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 2 the Total.
'                ctr2 = ctr2 + 1
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet2").Select
'                mRange1 = "D" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                mRange1 = "A" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
'                mRange1 = "B" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
'
'                mRange1 = "C" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
'                xlApp1.Sheets("Sheet1").Select
'
'    'mRange = "F" & cPos + 1 & ":K" & ctr
'    mRange = "F" & cPos + 1 & ":L" & ctr
'    xlSheet.Range(mRange).Select
'    mRange1 = "G" & ctr
'    xlSheet.Range(mRange1).Activate
'    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'    ' Copy to Sheet 2 the Total.
'    xlSheet.Range(mRange1).Select
'    xlApp1.Selection.Copy
'    xlApp1.Sheets("Sheet2").Select
'    mRange1 = "E" & ctr2
'    xlApp1.Range(mRange1).Select
'    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'        False, Transpose:=False
'    xlApp1.Sheets("Sheet1").Select
'
'    'mRange = "F" & cPos + 1 & ":K" & ctr
'    mRange = "F" & cPos + 1 & ":L" & ctr
'    xlSheet.Range(mRange).Select
'    mRange1 = "H" & ctr
'    xlSheet.Range(mRange1).Activate
'    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 2 the Total.
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet2").Select
'                mRange1 = "F" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                xlApp1.Sheets("Sheet1").Select
'
'    'mRange = "F" & cPos + 1 & ":K" & ctr
'    mRange = "F" & cPos + 1 & ":L" & ctr
'    xlSheet.Range(mRange).Select
'    mRange1 = "I" & ctr
'    xlSheet.Range(mRange1).Activate
'    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 2 the Total.
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet2").Select
'                mRange1 = "G" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                xlApp1.Sheets("Sheet1").Select
'
'    'mRange = "F" & cPos + 1 & ":K" & ctr
'    mRange = "F" & cPos + 1 & ":L" & ctr
'    xlSheet.Range(mRange).Select
'    mRange1 = "J" & ctr
'    xlSheet.Range(mRange1).Activate
'    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 2 the Total.
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet2").Select
'                mRange1 = "H" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                xlApp1.Sheets("Sheet1").Select
'
'    'mRange = "F" & cPos + 1 & ":K" & ctr
'    mRange = "F" & cPos + 1 & ":L" & ctr
'    xlSheet.Range(mRange).Select
'    mRange1 = "K" & ctr
'    xlSheet.Range(mRange1).Activate
'    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 2 the Total.
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet2").Select
'                mRange1 = "I" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                xlApp1.Sheets("Sheet1").Select
'
'    'mRange = "F" & cPos & ":K" & ctr
'    mRange = "F" & cPos & ":L" & ctr
'    xlSheet.Range(mRange).Select
'    mRange1 = "L" & ctr
'    xlSheet.Range(mRange1).Activate
'    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 2 the Total.
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet2").Select
'                mRange1 = "J" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                xlApp1.Sheets("Sheet1").Select
    
    'mRange = "F" & cPos & ":N" & ctr
    xlSheet.Range(mRange).Select
    xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
    mRange = "F" & ctr & ":N" & ctr
    xlSheet.Range(mRange).Select
    xlApp1.Range(mRange).Font.Bold = True
    
    If SumTotal = "" Then
        SumTotal = "=F" & ctr
    Else
        SumTotal = SumTotal & "+F" & ctr
    End If
Else
    ' Print Sub-Totals.
    ctr2 = ctr2 + 1
    xlApp1.Sheets("Sheet2").Select
    For l = 68 To 78 'D to N
        If l = 68 Then
            mRange1 = "A" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
            mRange1 = "B" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
            mRange1 = "C" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = 0
        End If
        mRange1 = Chr(l) & ctr2
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = 0#
        xlApp1.Sheets("Sheet1").Select
    Next l
'
'
'    mRange1 = "A" & ctr2
'    xlApp1.Range(mRange1).Select
'    xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
'    mRange1 = "B" & ctr2
'    xlApp1.Range(mRange1).Select
'    xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
'    mRange1 = "C" & ctr2
'    xlApp1.Range(mRange1).Select
'    xlApp1.ActiveCell.FormulaR1C1 = 0
'    mRange1 = "D" & ctr2
'    xlApp1.Range(mRange1).Select
'    xlApp1.ActiveCell.FormulaR1C1 = 0#
'    mRange1 = "E" & ctr2
'    xlApp1.Range(mRange1).Select
'    xlApp1.ActiveCell.FormulaR1C1 = 0#
'    mRange1 = "F" & ctr2
'    xlApp1.Range(mRange1).Select
'    xlApp1.ActiveCell.FormulaR1C1 = 0#
'    mRange1 = "G" & ctr2
'    xlApp1.Range(mRange1).Select
'    xlApp1.ActiveCell.FormulaR1C1 = 0#
'    mRange1 = "H" & ctr2
'    xlApp1.Range(mRange1).Select
'    xlApp1.ActiveCell.FormulaR1C1 = 0#
'    mRange1 = "I" & ctr2
'    xlApp1.Range(mRange1).Select
'    xlApp1.ActiveCell.FormulaR1C1 = 0#
'    mRange1 = "J" & ctr2
'    xlApp1.Range(mRange1).Select
'    xlApp1.ActiveCell.FormulaR1C1 = 0#
'    xlApp1.Sheets("Sheet1").Select
End If

If NoPHealthCtr > 0 Then
    LineCtr2 = LineCtr2 + 1
    ctr = PrintNoPHealthNumbers(xlSheet, xlApp1, PrevCompany, ctr)
End If
' Print Grand Totals.
mRange = "E9:E" & ctr
xlApp1.Range(mRange).Select
xlApp1.Selection.NumberFormat = "0"
'xlApp1.Selection.NumberFormat = "00-000000000-0"
'xlApp1.Selection.ColumnWidth = 15.29

' ************************************Print Totals ******************************************
ctr = ctr + 2
mRange = "B" & ctr & ":D" & ctr
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Totals", "", xlLeft, , False, False, False, , , 10, True)
GTotal = 0

For l = 70 To 78  '75 F to K (76) L up to N
    mRange = Chr(l) & ctr
    xlSheet.Range(mRange).Activate
    xlSheet.Cells(ctr, l - 64).VALUE = SumTotal
    If l = 73 Then
        xlSheet.Cells(ctr + 2, 9).VALUE = "Adjusted PS"
        xlSheet.Cells(ctr + 2, 10).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 2, 10).VALUE)
    ElseIf l = 74 Then
        xlSheet.Cells(ctr + 3, 9).VALUE = "Adjusted ES"
        xlSheet.Cells(ctr + 3, 10).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 3, 10).VALUE)
    End If
    SumTotal = Replace(SumTotal, Chr(l), Chr(l + 1))
Next l
mRange = "F" & cPos & ":N" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":N" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

ctr = ctr + 4
xlSheet.Cells(ctr, 9).VALUE = "Grand Total"
xlSheet.Cells(ctr, 10).VALUE = GTotal

mRange = "F" & cPos & ":N" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":N" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

' Print Unremitted PhilHealth Contributions.
If ExcludedCtr <> 0 Then
    ctr = UnremittedPHealthContributions(xlSheet, xlApp1, ctr)
End If

' ReFormat Sheet 2. Not yet done.
xlApp1.Sheets("Sheet2").Select
ctr2 = ctr2 + 1
GTotal = 0
For l = 68 To 76 '73 D to I, J up to L
    mRange1 = Chr(l) & ctr2
    xlApp1.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr2 - 6) * -1 & "]C:R[-1]C)"
    If l = 71 Then
        xlApp1.Range(mRange1).Activate
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.Copy
        mRange1 = "G" & ctr2 + 2
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = "Adjusted PS"
        mRange1 = "H" & ctr2 + 2
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        GTotal = GTotal + Val(xlApp1.ActiveCell.FormulaR1C1)
    ElseIf l = 72 Then
        xlApp1.Range(mRange1).Activate
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.Copy
        mRange1 = "G" & ctr2 + 3
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = "Adjusted ER"
        mRange1 = "H" & ctr2 + 3
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        GTotal = GTotal + Val(xlApp1.ActiveCell.FormulaR1C1)
    End If
Next l
mRange = "D7:L" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "C" & ctr2 & ":L" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

ctr2 = ctr2 + 4
mRange1 = "G" & ctr2
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = "Grand Total"
mRange1 = "H" & ctr2
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = GTotal

mRange1 = "D" & cPos & ":L" & ctr2
xlApp1.Range(mRange1).Select
xlApp1.Range(mRange1).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "D" & ctr2 & ":L" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

xlApp1.Sheets("Sheet2").Select
mHead = "Prepared by:"
mRange = "B" & ctr2 + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mHead = "Noted by:"
mRange = "C" & ctr2 + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mRange = "B" & ctr2 + 5
Call PrintReportHeader(xlSheet, xlApp1, mRange, mUserID, "", xlLeft, , False, False, False, , , 9, False)
mRange = "C" & ctr2 + 5
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Rhiza Capati", "", xlLeft, , False, False, False, , , 9, False)
mHead = "HR Officer"
mRange = "B" & ctr2 + 6
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mHead = "General Manager"
mRange = "C" & ctr2 + 6
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
ctr2 = ctr2 + 8

'Box it.
mRange = "A6:P" & ctr2
Call PlaceBox(xlSheet, xlApp1, mRange)
With xlApp1.Range(mRange)
    .Font.Name = "Arial"   'Font.Name
    .Font.Size = 8         'Font.size
End With

mRange = "D7:L" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "C" & ctr2 & ":L" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

mRange = "A6:L" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Selection.Columns.AutoFit
xlApp1.Range("A1").Select

' ReFormat Sheet 3.
xlApp1.Sheets("Sheet3").Select
ctr3 = ctr3 + 1
GTotal = 0
For l = 68 To 76 ' 73 D to I, J
    mRange1 = Chr(l) & ctr3
    xlApp1.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr3 - 6) * -1 & "]C:R[-1]C)"
    If l = 71 Then
        xlApp1.Range(mRange1).Activate
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.Copy
        mRange1 = "G" & ctr3 + 2
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = "Adjusted PS"
        mRange1 = "H" & ctr3 + 2
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        GTotal = GTotal + Val(xlApp1.ActiveCell.FormulaR1C1)
    ElseIf l = 72 Then
        xlApp1.Range(mRange1).Activate
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.Copy
        mRange1 = "G" & ctr3 + 3
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = "Adjusted ES"
        mRange1 = "H" & ctr3 + 3
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        GTotal = GTotal + Val(xlApp1.ActiveCell.FormulaR1C1)
    End If
Next l
mRange = "D7:L" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"

mRange = "C" & ctr3 & ":L" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

ctr3 = ctr3 + 4
mRange1 = "G" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = "Grand Total"
mRange1 = "H" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = GTotal


mRange1 = "D" & cPos & ":L" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.Range(mRange1).NumberFormat = "#,##0.00_);(#,##0.00)"

mRange = "D" & ctr2 & ":L" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

xlApp1.Sheets("Sheet3").Select
mHead = "Prepared by:"
mRange = "B" & ctr3 + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mHead = "Noted by:"
mRange = "C" & ctr3 + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mRange = "B" & ctr3 + 5
Call PrintReportHeader(xlSheet, xlApp1, mRange, mUserID, "", xlLeft, , False, False, False, , , 9, False)
mRange = "C" & ctr3 + 5
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Rhiza Capati", "", xlLeft, , False, False, False, , , 9, False)
mHead = "HR Officer"
mRange = "B" & ctr3 + 6
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mHead = "General Manager"
mRange = "C" & ctr3 + 6
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
ctr3 = ctr3 + 8

'Box it.
mRange = "A6:P" & ctr3
Call PlaceBox(xlSheet, xlApp1, mRange)
With xlApp1.Range(mRange)
    .Font.Name = "Arial"   'Font.Name
    .Font.Size = 8         'Font.size
End With

mRange = "D7:L" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "C" & ctr3 & ":L" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

mRange = "A6:L" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Selection.Columns.AutoFit
xlApp1.Range("A1").Select

'                   Date Transmitted:   January 2009
'   Company         Applicable Month:   October 2008
'   SSS# 03-9210478-2

'   SSS Payments    EC  Total       SBR NUMBER/OR#      Date Paid

xlApp1.Sheets("Sheet1").Select

mHead = "Prepared by:"
mRange = "B" & ctr + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mHead = "Noted by:"
mRange = "E" & ctr + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mRange = "B" & ctr + 5
Call PrintReportHeader(xlSheet, xlApp1, mRange, mUserID, "", xlLeft, , False, False, False, , , 9, True)
mRange = "E" & ctr + 5
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Rhiza Capati", "", xlLeft, , False, False, False, , , 9, True)
mHead = "HR Officer"
mRange = "B" & ctr + 6
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mHead = "General Manager"
mRange = "E" & ctr + 6
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
ctr = ctr + 8

'Box it.
'mRange = "A6:K" & ctr
mRange = "A6:R" & ctr
Call PlaceBox(xlSheet, xlApp1, mRange)

xlApp1.Sheets("Sheet1").Name = "PhilHealth Details"
xlApp1.Sheets("Sheet2").Name = "PhilHealth Summary"
xlApp1.Sheets("Sheet3").Name = "Not to be Remitted PhilHealth"

xlSheet.Range("A1").Select

ExitRtn2:

End Sub

Private Function UnremittedPHealthContributions(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal ctr As Integer) As Integer
Dim mRange As String, mRange1 As String
Dim LineCtr As Integer
Dim cPos As Integer
Dim i As Integer, l As Integer, t As Integer
Dim GTotal As Double

i = 1
LineCtr = 1
ctr = ctr + 2
cPos = ctr + 1

PrevCompany = ListView1(1).ListItems(i).SubItems(12)
mRange = "A" & ctr & ":D" & ctr
Call PrintReportHeader(xlSheet, xlApp1, mRange, "NOT TO BE REMITTED CLIENTS", "", xlLeft, , False, False, False, , , 10, True)

Do While ListView1(1).ListItems.Count >= i
    If ListView1(1).ListItems.Item(i).Checked = False Then
        GoTo Next3
    Else
        If PrevCompany <> ListView1(1).ListItems(i).SubItems(12) Then
            If LineCtr > 1 Then
                ' Print Sub-Totals.
                ctr = ctr + 1
                ctr3 = ctr3 + 1
                mRange = "F" & cPos + 1 & ":N" & ctr
                For l = 70 To 78 'F to N
                    xlSheet.Range(mRange).Select
                    mRange1 = Chr(l) & ctr
                    xlSheet.Range(mRange1).Activate
                    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
                
                    ' Copy to Sheet 2 the Total.
                    xlSheet.Range(mRange1).Select
                    xlApp1.Selection.Copy
                    xlApp1.Sheets("Sheet3").Select
                    mRange1 = Chr(l - 2) & ctr3
                    xlApp1.Range(mRange1).Select
                    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                        False, Transpose:=False
                    If l = 70 Then
                        mRange1 = "A" & ctr3
                        xlApp1.Range(mRange1).Select
                        xlApp1.ActiveCell.FormulaR1C1 = LineCtr3
                        mRange1 = "B" & ctr3
                        xlApp1.Range(mRange1).Select
                        xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
                        mRange1 = "C" & ctr3
                        xlApp1.Range(mRange1).Select
                        xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
                    End If
                    xlApp1.Sheets("Sheet1").Select
                Next l
'
'                xlSheet.Range(mRange).Select
'                mRange1 = "F" & ctr
'                xlSheet.Range(mRange1).Activate
'                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 3 the Total.
'                ctr3 = ctr3 + 1
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet3").Select
'                mRange1 = "D" & ctr3
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                mRange1 = "A" & ctr3
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = LineCtr3
'                mRange1 = "B" & ctr3
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
'                mRange1 = "C" & ctr3
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
'                xlApp1.Sheets("Sheet1").Select
'
'                'mRange = "F" & cPos + 1 & ":K" & ctr
'                mRange = "F" & cPos + 1 & ":L" & ctr
'                xlSheet.Range(mRange).Select
'                mRange1 = "G" & ctr
'                xlSheet.Range(mRange1).Activate
'                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 3 the Total.
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet3").Select
'                mRange1 = "E" & ctr3
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                xlApp1.Sheets("Sheet1").Select
'
'                'mRange = "F" & cPos + 1 & ":K" & ctr
'                mRange = "F" & cPos + 1 & ":L" & ctr
'                xlSheet.Range(mRange).Select
'                mRange1 = "H" & ctr
'                xlSheet.Range(mRange1).Activate
'                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 3 the Total.
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet3").Select
'                mRange1 = "F" & ctr3
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                xlApp1.Sheets("Sheet1").Select
'
'                'mRange = "F" & cPos + 1 & ":K" & ctr
'                mRange = "F" & cPos + 1 & ":L" & ctr
'                xlSheet.Range(mRange).Select
'                mRange1 = "I" & ctr
'                xlSheet.Range(mRange1).Activate
'                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 3 the Total.
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet3").Select
'                mRange1 = "G" & ctr3
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                xlApp1.Sheets("Sheet1").Select
'
'                'mRange = "F" & cPos + 1 & ":K" & ctr
'                mRange = "F" & cPos + 1 & ":L" & ctr
'                xlSheet.Range(mRange).Select
'                mRange1 = "J" & ctr
'                xlSheet.Range(mRange1).Activate
'                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 3 the Total.
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet3").Select
'                mRange1 = "H" & ctr3
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                xlApp1.Sheets("Sheet1").Select
'
'                'mRange = "F" & cPos + 1 & ":K" & ctr
'                mRange = "F" & cPos + 1 & ":L" & ctr
'                xlSheet.Range(mRange).Select
'                mRange1 = "K" & ctr
'                xlSheet.Range(mRange1).Activate
'                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 3 the Total.
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet3").Select
'                mRange1 = "I" & ctr3
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                xlApp1.Sheets("Sheet1").Select
'
'                'Updated June 2009.
'                mRange = "F" & cPos + 1 & ":L" & ctr
'                xlSheet.Range(mRange).Select
'                mRange1 = "L" & ctr
'                xlSheet.Range(mRange1).Activate
'                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'                ' Copy to Sheet 3 the Total.
'                xlSheet.Range(mRange1).Select
'                xlApp1.Selection.Copy
'                xlApp1.Sheets("Sheet3").Select
'                mRange1 = "J" & ctr3
'                xlApp1.Range(mRange1).Select
'                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'                    False, Transpose:=False
'                xlApp1.Sheets("Sheet1").Select
                
                LineCtr3 = LineCtr3 + 1
                
                'mRange = "F" & cPos & ":N" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                mRange = "F" & ctr & ":N" & ctr
                xlSheet.Range(mRange).Select
                xlApp1.Range(mRange).Font.Bold = True
            
                If SumTotal = "" Then
                    SumTotal = "=F" & ctr
                Else
                    SumTotal = SumTotal & "+F" & ctr
                End If
            
            End If
            PrevCompany = ListView1(1).ListItems(i).SubItems(12)
            HeaderFlag = False
            NoPHealthCtr = 0  ' set again to 0 after printing
        End If
        ' Print Header.
        If HeaderFlag = False Then
            LineCtr = 1
            ctr = ctr + 3
            mRange = "A" & ctr
            Call PrintReportHeader(xlSheet, xlApp1, mRange, PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
            ctr = ctr + 1
            mRange = "A" & ctr & ":A" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "No.", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 4.14)
            mRange = "B" & ctr & ":B" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Surname", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)
            mRange = "C" & ctr & ":C" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Firstname", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
            mRange = "D" & ctr & ":D" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Middlename", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
            mRange = "E" & ctr & ":E" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "PhilHealth#", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 11)
        
            For t = 68 To 80 ' F to P
                Call PrintHeadings(xlSheet, xlApp1, t, ctr, 2)
            Next t
            
'
'            mRange = "F" & ctr & ":F" & ctr + 1
'            'Call PrintReportHeader(xlSheet, xlApp1, "F7:F8", "Employee Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "PS Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            mRange = "G" & ctr & ":G" & ctr + 1
'            'Call PrintReportHeader(xlSheet, xlApp1, "G7:G8", "Employer Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "ES Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            mRange = "H" & ctr & ":H" & ctr + 1
'            'Call PrintReportHeader(xlSheet, xlApp1, "H7:H8", "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "PS+ES Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
'            mRange = "I" & ctr & ":I" & ctr + 1
'            'Call PrintReportHeader(xlSheet, xlApp1, "I7:I8", "Adjusted", "PS Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "PS Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            mRange = "J" & ctr & ":J" & ctr + 1
'            'Call PrintReportHeader(xlSheet, xlApp1, "J7:J8", "Adjusted", "ES Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "ES Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            mRange = "K" & ctr & ":K" & ctr + 1
'            'Call PrintReportHeader(xlSheet, xlApp1, "K7:K8", "Difference", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'
'
'            mRange = "L" & ctr & ":L" & ctr + 1
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "Adjusted PS+ES", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            mRange = "M" & ctr & ":M" & ctr + 1
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Remarks", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            
            HeaderFlag = True
            cPos = ctr + 2
            ctr = ctr + 1
        End If
        
            ctr = ctr + 1
            xlSheet.Cells(ctr, 1).VALUE = LineCtr
            xlSheet.Cells(ctr, 2).VALUE = ListView1(1).ListItems(i).SubItems(1) ' Surname
            xlSheet.Cells(ctr, 3).VALUE = ListView1(1).ListItems(i).SubItems(2) ' Firstname
            xlSheet.Cells(ctr, 4).VALUE = ListView1(1).ListItems(i).SubItems(3) ' Middlename
            xlSheet.Cells(ctr, 5).VALUE = ListView1(1).ListItems(i).SubItems(4) ' SSS no
            xlSheet.Cells(ctr, 6).VALUE = ListView1(1).ListItems(i).SubItems(7) ' Employee
            xlSheet.Cells(ctr, 7).VALUE = ListView1(1).ListItems(i).SubItems(8) ' Employer
            xlSheet.Cells(ctr, 8).VALUE = Val(ListView1(1).ListItems(i).SubItems(7)) + Val(ListView1(1).ListItems(i).SubItems(8)) ' PS + ES Share
            xlSheet.Cells(ctr, 9).VALUE = ListView1(1).ListItems(i).SubItems(9) ' employee adjust
            xlSheet.Cells(ctr, 10).VALUE = ListView1(1).ListItems(i).SubItems(10) ' employer adjust
            xlSheet.Cells(ctr, 11).VALUE = Val(ListView1(1).ListItems(i).SubItems(10)) + Val(ListView1(1).ListItems(i).SubItems(9))  ' Total PS + ES
            xlSheet.Cells(ctr, 12).VALUE = Val(ListView1(1).ListItems(i).SubItems(7)) - Val(ListView1(1).ListItems(i).SubItems(9))  ' difference PS
            xlSheet.Cells(ctr, 13).VALUE = Val(ListView1(1).ListItems(i).SubItems(8)) - Val(ListView1(1).ListItems(i).SubItems(10))  ' difference ES
            xlSheet.Cells(ctr, 14).VALUE = (Val(ListView1(1).ListItems(i).SubItems(8)) - Val(ListView1(1).ListItems(i).SubItems(10))) + (Val(ListView1(1).ListItems(i).SubItems(7)) - Val(ListView1(1).ListItems(i).SubItems(9)))
            LineCtr = LineCtr + 1
        End If
Next3:
    
    i = i + 1

Loop

' Print last Sub-Totals.
ctr = ctr + 1
ctr3 = ctr3 + 1
mRange = "F" & cPos + 1 & ":N" & ctr
For l = 70 To 78 'F to N
    xlSheet.Range(mRange).Select
    mRange1 = Chr(l) & ctr
    xlSheet.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

    ' Copy to Sheet 2 the Total.
    xlSheet.Range(mRange1).Select
    xlApp1.Selection.Copy
    xlApp1.Sheets("Sheet3").Select
    mRange1 = Chr(l - 2) & ctr3
    xlApp1.Range(mRange1).Select
    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    If l = 70 Then
        mRange1 = "A" & ctr3
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = LineCtr3
        mRange1 = "B" & ctr3
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
        mRange1 = "C" & ctr3
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
    End If
    xlApp1.Sheets("Sheet1").Select
Next l

'
'mRange = "F" & cPos + 1 & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "F" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'' Copy to Sheet 3 the Total.
'ctr3 = ctr3 + 1
'xlSheet.Range(mRange1).Select
'xlApp1.Selection.Copy
'xlApp1.Sheets("Sheet3").Select
'mRange1 = "D" & ctr3
'xlApp1.Range(mRange1).Select
'xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'    False, Transpose:=False
'mRange1 = "A" & ctr3
'xlApp1.Range(mRange1).Select
'xlApp1.ActiveCell.FormulaR1C1 = LineCtr3
'mRange1 = "B" & ctr3
'xlApp1.Range(mRange1).Select
'xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
'mRange1 = "C" & ctr3
'xlApp1.Range(mRange1).Select
'xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
'xlApp1.Sheets("Sheet1").Select
'''''''-----------
'
''mRange = "F" & cPos + 1 & ":K" & ctr
'mRange = "F" & cPos + 1 & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "G" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'' Copy to Sheet 3 the Total.
'xlSheet.Range(mRange1).Select
'xlApp1.Selection.Copy
'xlApp1.Sheets("Sheet3").Select
'mRange1 = "E" & ctr3
'xlApp1.Range(mRange1).Select
'xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'    False, Transpose:=False
'xlApp1.Sheets("Sheet1").Select
'
''mRange = "F" & cPos + 1 & ":K" & ctr
'mRange = "F" & cPos + 1 & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "H" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'' Copy to Sheet 3 the Total.
'xlSheet.Range(mRange1).Select
'xlApp1.Selection.Copy
'xlApp1.Sheets("Sheet3").Select
'mRange1 = "F" & ctr3
'xlApp1.Range(mRange1).Select
'xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'    False, Transpose:=False
'xlApp1.Sheets("Sheet1").Select
'
''mRange = "F" & cPos + 1 & ":K" & ctr
'mRange = "F" & cPos + 1 & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "I" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'' Copy to Sheet 3 the Total.
'xlSheet.Range(mRange1).Select
'xlApp1.Selection.Copy
'xlApp1.Sheets("Sheet3").Select
'mRange1 = "G" & ctr3
'xlApp1.Range(mRange1).Select
'xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'    False, Transpose:=False
'xlApp1.Sheets("Sheet1").Select
'
''mRange = "F" & cPos + 1 & ":K" & ctr
'mRange = "F" & cPos + 1 & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "J" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'' Copy to Sheet 3 the Total.
'xlSheet.Range(mRange1).Select
'xlApp1.Selection.Copy
'xlApp1.Sheets("Sheet3").Select
'mRange1 = "H" & ctr3
'xlApp1.Range(mRange1).Select
'xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'    False, Transpose:=False
'xlApp1.Sheets("Sheet1").Select
'
''mRange = "F" & cPos + 1 & ":K" & ctr
'mRange = "F" & cPos + 1 & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "K" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'' Copy to Sheet 3 the Total.
'xlSheet.Range(mRange1).Select
'xlApp1.Selection.Copy
'xlApp1.Sheets("Sheet3").Select
'mRange1 = "I" & ctr3
'xlApp1.Range(mRange1).Select
'xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'    False, Transpose:=False
'xlApp1.Sheets("Sheet1").Select
'
''Updated June 2009.
''mRange = "F" & cPos + 1 & ":K" & ctr
'mRange = "F" & cPos + 1 & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "L" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
'
'' Copy to Sheet 3 the Total.
'xlSheet.Range(mRange1).Select
'xlApp1.Selection.Copy
'xlApp1.Sheets("Sheet3").Select
'mRange1 = "J" & ctr3
'xlApp1.Range(mRange1).Select
'xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'    False, Transpose:=False
'xlApp1.Sheets("Sheet1").Select

'mRange = "F" & cPos & ":N" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":N" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

If SumTotal = "" Then
    SumTotal = "=F" & ctr
Else
    SumTotal = SumTotal & "+F" & ctr
End If

ctr = ctr + 2
mRange = "B" & ctr & ":D" & ctr
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Totals", "", xlLeft, , False, False, False, , , 10, True)
GTotal = 0

For l = 70 To 78  '75 F to K (76) L up to N
    mRange = Chr(l) & ctr
    xlSheet.Range(mRange).Activate
    xlSheet.Cells(ctr, l - 64).VALUE = SumTotal
    If l = 70 Then  ' 73
        xlSheet.Cells(ctr + 2, 9).VALUE = "Actual PS"
        xlSheet.Cells(ctr + 2, 10).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 2, 10).VALUE)
    ElseIf l = 71 Then ' 74
        xlSheet.Cells(ctr + 3, 9).VALUE = "Actual ES"
        xlSheet.Cells(ctr + 3, 10).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 3, 10).VALUE)
    End If
    SumTotal = Replace(SumTotal, Chr(l), Chr(l + 1))
Next l
mRange = "F" & cPos & ":N" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":N" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

ctr = ctr + 4
xlSheet.Cells(ctr, 9).VALUE = "Grand Total"
xlSheet.Cells(ctr, 10).VALUE = GTotal

mRange = "F" & cPos & ":N" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":N" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

UnremittedPHealthContributions = ctr

End Function

