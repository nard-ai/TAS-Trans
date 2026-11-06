VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSSSReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSS Report Form"
   ClientHeight    =   8520
   ClientLeft      =   2175
   ClientTop       =   2715
   ClientWidth     =   12600
   Icon            =   "frmSSSReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   12600
   Begin VB.CheckBox chkR1A 
      Caption         =   "R-1A SSS Report"
      Height          =   255
      Left            =   8040
      TabIndex        =   18
      ToolTipText     =   "All Payments"
      Top             =   1250
      Width           =   2055
   End
   Begin VB.CommandButton cmdSearchSBR 
      Caption         =   "Search"
      Height          =   255
      Left            =   5760
      TabIndex        =   73
      ToolTipText     =   "Search SBR Number"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdateSBR 
      Caption         =   "Update"
      Height          =   255
      Left            =   5760
      TabIndex        =   70
      ToolTipText     =   "Update SBR Entries"
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelectGroupCodes 
      Caption         =   "Select"
      Height          =   255
      Left            =   11760
      TabIndex        =   67
      ToolTipText     =   "Select Group Payments"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CheckBox chkAllPayments 
      Height          =   255
      Left            =   12120
      TabIndex        =   66
      ToolTipText     =   "All Payments"
      Top             =   840
      Width           =   255
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
      ItemData        =   "frmSSSReport.frx":08CA
      Left            =   4800
      List            =   "frmSSSReport.frx":08D1
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   6480
      Width           =   2205
   End
   Begin VB.CommandButton cmdSBRReport 
      Caption         =   "SSS w/ SBR Report"
      Height          =   375
      Left            =   9360
      TabIndex        =   64
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton cmdDeleteAll 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   240
      TabIndex        =   60
      ToolTipText     =   "process"
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdNRDK 
      Caption         =   "Copy NR3001DK to Disk"
      Height          =   375
      Left            =   7200
      TabIndex        =   59
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
      TabIndex        =   55
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
      TabIndex        =   54
      Top             =   6590
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   11400
      TabIndex        =   53
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "Retrieve Payments"
      Height          =   375
      Left            =   1320
      TabIndex        =   52
      ToolTipText     =   "process"
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CheckBox chkIncluded 
      Caption         =   "Exclude Checked Records"
      Height          =   255
      Left            =   3000
      TabIndex        =   51
      Top             =   6240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdGenerateReport 
      Caption         =   "Generate SSS Report"
      Height          =   375
      Left            =   7440
      TabIndex        =   48
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
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
      ItemData        =   "frmSSSReport.frx":08DF
      Left            =   10800
      List            =   "frmSSSReport.frx":08F8
      Style           =   2  'Dropdown List
      TabIndex        =   22
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
      MaxLength       =   40
      TabIndex        =   47
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
      MaxLength       =   40
      TabIndex        =   45
      Top             =   7125
      Width           =   2175
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Add/Delete Employees"
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      ToolTipText     =   "process"
      Top             =   7920
      Width           =   1935
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
      ItemData        =   "frmSSSReport.frx":0926
      Left            =   10800
      List            =   "frmSSSReport.frx":094E
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   840
      Width           =   1245
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
      TabIndex        =   44
      Top             =   7440
      Width           =   1215
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
      TabIndex        =   43
      Top             =   7080
      Width           =   1215
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
      TabIndex        =   40
      Top             =   7320
      Visible         =   0   'False
      Width           =   255
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
      TabIndex        =   37
      Text            =   "0"
      Top             =   7320
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
      TabIndex        =   34
      Text            =   "0.00"
      Top             =   7320
      Width           =   1095
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
      MaxLength       =   20
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
      Visible         =   0   'False
      Width           =   1335
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
      Caption         =   "REGULAR"
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
      MaxLength       =   60
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
      MaxLength       =   30
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
      MaxLength       =   30
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
      Index           =   0
      Left            =   240
      TabIndex        =   58
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
         Text            =   "Client Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Table Based EE"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Table Based ER"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Group Code"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Amount Paid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "SBR No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Date Paid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Date Posted"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Object.Width           =   882
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
      Index           =   1
      Left            =   240
      TabIndex        =   49
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
      NumItems        =   12
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
      TabIndex        =   71
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
      TabIndex        =   72
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
   Begin VB.CommandButton cmdR1AReport 
      Caption         =   "R-1A Report"
      Height          =   375
      Left            =   9360
      TabIndex        =   19
      Top             =   7920
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Index           =   2
      Left            =   240
      TabIndex        =   20
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
      BackColor       =   12648447
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
      NumItems        =   15
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
         Text            =   "Date of Birth"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Basic Salary"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Position"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Employment Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
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
   End
   Begin VB.Label Label2 
      Caption         =   "Total Employees:"
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
      Index           =   21
      Left            =   360
      TabIndex        =   75
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
      Index           =   5
      Left            =   9120
      TabIndex        =   74
      Top             =   6240
      Width           =   1605
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
      TabIndex        =   69
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
      Left            =   3480
      TabIndex        =   68
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
      Index           =   3
      Left            =   5520
      TabIndex        =   63
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
      Left            =   7320
      TabIndex        =   62
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
      Index           =   0
      Left            =   10800
      TabIndex        =   61
      Top             =   6240
      Width           =   1485
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
      TabIndex        =   57
      Top             =   6600
      Width           =   2445
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
      Index           =   7
      Left            =   5400
      TabIndex        =   56
      Top             =   6600
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label2 
      Caption         =   "Total Employees:"
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
      TabIndex        =   50
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
      TabIndex        =   46
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
      TabIndex        =   42
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
      TabIndex        =   41
      Top             =   7080
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
      Alignment       =   2  'Center
      Caption         =   "No. of Employees"
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
      Index           =   15
      Left            =   5640
      TabIndex        =   39
      Top             =   7080
      Visible         =   0   'False
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
      TabIndex        =   38
      Top             =   7080
      Width           =   1215
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
      TabIndex        =   36
      Top             =   7080
      Width           =   1215
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
      TabIndex        =   35
      Top             =   7080
      Width           =   855
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
      Caption         =   "SSS NUMBER:"
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
      Caption         =   "SSS"
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
      Height          =   1365
      Left            =   240
      Picture         =   "frmSSSReport.frx":09B4
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "frmSSSReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsItems As ADODB.Recordset
Dim StrQuery As String
Dim mPage As Integer
Dim gMpage As Integer
Dim TotEmp As Integer
Dim QtrNo As Integer
Dim DTRNo As Integer
Dim SumTotal As String
Dim ctr2 As Integer, ctr3 As Integer
Dim LineCtr2 As Integer, LineCtr3 As Integer
Dim cPos2 As Integer
Dim SSSGroupCode As String
Dim mTDiff As Double
Dim mTEE As Double
Dim mTER As Double
Dim mECC As Double
Dim WithSBRNo As Boolean

Private Sub chkR1A_Click()
If chkR1A.VALUE = 1 Then
    R1AFlag = True
    ListView1(2).Visible = True
    ListView1(0).Visible = False
    ListView1(1).Visible = False
    cmdNRDK.Visible = False
    cmdSBRReport.Visible = False
    cmdR1AReport.Visible = True
Else
    R1AFlag = False
    ListView1(2).Visible = False
    ListView1(0).Visible = True
    ListView1(1).Visible = True
    cmdNRDK.Visible = True
    cmdSBRReport.Visible = True
    cmdR1AReport.Visible = False
End If
End Sub

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
Dim unZeroDed As Double

ZeroDed = 0
If cmdCheckZero.Caption = "Check Zero Contributions" Then
    For i = 1 To ListView1(0).ListItems.Count
        mZeroDed = Val(ListView1(0).ListItems.Item(i).SubItems(10)) ' get adjusted employee share
        If mZeroDed = 0# Then
            'If ListView1(0).ListItems.Item(i).Checked = True Then
            '    ListView1(0).ListItems.Item(i).Checked = False
            '    unZeroDed = unZeroDed + 1
            'Else
                ListView1(0).ListItems.Item(i).Checked = True
                ListView1(0).ListItems.Item(i).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(1).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(2).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(3).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(4).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(5).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(6).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(7).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(8).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(9).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(10).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(11).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(12).ForeColor = vbRed
                ListView1(0).ListItems.Item(i).ListSubItems(13).ForeColor = vbRed
                ZeroDed = ZeroDed + 1
            'End If
        End If
    Next i
    lblclient(2).Caption = "Check Zero Contributions: " & ZeroDed
    cmdCheckZero.Caption = "Uncheck Zero Contributions"
Else        ' Uncheck zero contributions
    For i = 1 To ListView1(0).ListItems.Count
        mZeroDed = Val(ListView1(0).ListItems.Item(i).SubItems(10)) ' get adjusted employee share
        If mZeroDed = 0# Then
            'If ListView1(0).ListItems.Item(i).Checked = True Then
                ListView1(0).ListItems.Item(i).Checked = False
                ListView1(0).ListItems.Item(i).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(1).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(2).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(3).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(4).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(5).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(6).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(7).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(8).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(9).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(10).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(11).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(12).ForeColor = vbBlack
                ListView1(0).ListItems.Item(i).ListSubItems(13).ForeColor = vbBlack
            '    unZeroDed = unZeroDed + 1
            'Else
            '    ListView1(0).ListItems.Item(i).Checked = True
                ZeroDed = ZeroDed + 1
            'End If
        End If
    Next i
    lblclient(2).Caption = "Uncheck Zero Contributions: " & ZeroDed
    cmdCheckZero.Caption = "Check Zero Contributions"
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCompany_Click()
    ReportedFlag = 2
    SelectCompanyFlag = True
    PHealthFlag = False
    frmCompany.Show
    frmCompany.SetFocus
End Sub

Private Sub cmdDeleteAll_Click()

If chkR1A.VALUE = 1 Then
    ListView1(2).ListItems.Clear
    lblclient(7).Caption = "No SSS No.: 0"
    lblclient(0).Caption = "Difference: 0"
    lblclient(3).Caption = "Total EE: 0"
    lblclient(1).Caption = "Total ER: 0"
    lblclient(5).Caption = "Total ECC: 0"
    lblclient(2).Caption = "Zero Contributions: 0"
Else
    ListView1(0).ListItems.Clear
    lblclient(7).Caption = "No SSS No.: 0"
    lblclient(0).Caption = "Difference: 0"
    lblclient(3).Caption = "Total EE: 0"
    lblclient(1).Caption = "Total ER: 0"
    lblclient(5).Caption = "Total ECC: 0"
    lblclient(2).Caption = "Zero Contributions: 0"

End If
    MsgBox "Employee List was cleared.", vbInformation, PayrollAlert
End Sub

Private Sub cmdGenerateReport_Click()
Dim X As Printer
Dim y As PrinterOrientationConstants

ListView1(0).SortKey = 0
ListView1(0).Sorted = True
ListView1(0).Refresh
If ListView1(0).ListItems.Count >= 1 Then
    Call GenerateSSSReports
Else
    MsgBox "No SSS records to processed.", vbInformation, PayrollAlert
End If
ListView1(0).SortKey = 1
ListView1(0).Sorted = True
ListView1(0).Refresh

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

'Private Sub GenerateSSSReports2()
'Dim xlApp1 As Excel.Application
'Dim xlBook1 As Excel.Workbook
'Dim xlSheet As Excel.Worksheet
'Dim mRange As String
'Dim PrevCompany As String
'Dim ctr As Long
'Dim i As Integer
'
'Set xlApp1 = New Excel.Application
'Set xlBook1 = xlApp1.Workbooks.Add()
'
'xlApp1.WindowState = xlMaximized
'xlApp1.Visible = True
'xlApp1.DisplayAlerts = False
'
'Set xlSheet = xlBook1.Worksheets(1)
'
'ctr = 7
'i = 1
'
'On Error GoTo errH1
'
'' Print Header
'mHead = CompanyName & " (" & CompanyShort & ")"
'Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 12, True)
'mHead = "SSS Contributions Report"
'Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, False)
'mHead = "For the month " & cmdAppMonth.Text & " " & cmdAppYear.Text
'Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, False)
''mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
''Call PrintReportHeader(xlSheet, xlApp1, "A5", mHead, "", xlLeft, , False, False, False, , , 10, False)
'
'PrevCompany = ListView1(0).ListItems(i).SubItems(13)
'Call PrintReportHeader(xlSheet, xlApp1, "A6", PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
'
'Call PrintReportHeader(xlSheet, xlApp1, "A7", "No.", "", xlLeft, , False, False, False, , , 10, True, 5)
'Call PrintReportHeader(xlSheet, xlApp1, "B7", "Surname", "", xlLeft, , False, False, False, , , 10, True, 13)
'Call PrintReportHeader(xlSheet, xlApp1, "C7", "Firstname", "", xlLeft, , False, False, False, , , 10, True, 15)
'Call PrintReportHeader(xlSheet, xlApp1, "D7", "Middlename", "", xlLeft, , False, False, False, , , 10, True, 13)
'Call PrintReportHeader(xlSheet, xlApp1, "E7", "SSS Number", "", xlLeft, , False, False, False, , , 10, True, 12)
'Call PrintReportHeader(xlSheet, xlApp1, "F7", "EE Share", "", xlLeft, , False, False, False, , , 10, True)
'Call PrintReportHeader(xlSheet, xlApp1, "G7", "ER Share", "", xlLeft, , False, False, False, , , 10, True)
'Call PrintReportHeader(xlSheet, xlApp1, "H7", "ECC", "", xlLeft, , False, False, False, , , 10, True, 6)
'Call PrintReportHeader(xlSheet, xlApp1, "I7", "Adjusted EE", "", xlLeft, , False, False, False, , , 10, True, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "J7", "Adjusted ER", "", xlLeft, , False, False, False, , , 10, True, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "K7", "Difference", "", xlLeft, , False, False, False, , , 9, True, 9)
'
'Do While ListView1(0).ListItems.Count >= i
'
'If PrevCompany <> ListView1(0).ListItems(i).SubItems(13) Then
'    ctr = ctr + 3
'    PrevCompany = ListView1(0).ListItems(i).SubItems(13)
'    mRange = "A" & ctr
'    Call PrintReportHeader(xlSheet, xlApp1, mRange, PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
'End If
'
'If ListView1(0).ListItems.Item(i).Checked = True Or ListView1(0).ListItems(i).SubItems(4) = "__-_______-_" Then
'    GoTo Next1
'Else
'    ctr = ctr + 1
'    xlSheet.Cells(ctr, 1).VALUE = i
'    xlSheet.Cells(ctr, 2).VALUE = ListView1(0).ListItems(i).SubItems(1) ' Surname
'    xlSheet.Cells(ctr, 3).VALUE = ListView1(0).ListItems(i).SubItems(2) ' firstname
'    xlSheet.Cells(ctr, 4).VALUE = ListView1(0).ListItems(i).SubItems(3) ' Middlename
'    xlSheet.Cells(ctr, 5).VALUE = ListView1(0).ListItems(i).SubItems(4) ' sss no
'    xlSheet.Cells(ctr, 6).VALUE = ListView1(0).ListItems(i).SubItems(7) ' employee
'    xlSheet.Cells(ctr, 7).VALUE = ListView1(0).ListItems(i).SubItems(8) ' employer
'    xlSheet.Cells(ctr, 8).VALUE = ListView1(0).ListItems(i).SubItems(9) ' ecc
'    xlSheet.Cells(ctr, 9).VALUE = ListView1(0).ListItems(i).SubItems(10) ' employee adjust
'    xlSheet.Cells(ctr, 10).VALUE = ListView1(0).ListItems(i).SubItems(11) ' employer adjust
'    xlSheet.Cells(ctr, 11).VALUE = ListView1(0).ListItems(i).SubItems(12) ' difference
'
'End If
'
'Next1:
'
'    i = i + 1
'
'Loop
'
'errH1:
'
'End Sub

'Private Sub getTotalPages()
'Dim i As Integer
'
'i = 1
'TotEmp = 0
'Do While ListView1(1).ListItems.Count >= i
'
'    If i <> 0 And ListView1(1).ListItems.Item(i).Checked = False Then
'        TotEmp = TotEmp + 1
'    End If
'    i = i + 1
'Loop
'gMpage = TotEmp \ 15
'If TotEmp Mod 15 <> 0 Then
'    gMpage = gMpage + 1
'End If
'txtNoOfEmp.Text = TotEmp
'
'End Sub

Private Sub cmdNoPHNo_Click()
Dim NoPHNo As Integer

NoPHNo = 0
For i = 1 To ListView1(0).ListItems.Count
    If ListView1(0).ListItems.Item(i).SubItems(4) = "" Or ListView1(0).ListItems.Item(i).SubItems(4) = "__-_______-_" Then ' get philhealth number
        If ListView1(0).ListItems.Item(i).Checked = True Then
            ListView1(0).ListItems.Item(i).Checked = False
        Else
            ListView1(0).ListItems.Item(i).Checked = True
        End If
        NoPHNo = NoPHNo + 1
    End If
Next i

lblclient(7).Caption = "No SSS No.: " & NoPHNo

End Sub

Private Sub GetTotalSSS()
Dim ExcludedCtr As Integer

ExcludedCtr = 0
mTDiff = 0: mTEE = 0: mTER = 0: mECC = 0
mTotal = 0#
For i = 1 To ListView1(0).ListItems.Count
    If ListView1(0).ListItems.Item(i).Checked = False Then
        mTEE = mTEE + Val(ListView1(0).ListItems(i).SubItems(10)) ' employee adjust
        mTER = mTER + Val(ListView1(0).ListItems(i).SubItems(11)) ' employer adjust
        mECC = mECC + Val(ListView1(0).ListItems(i).SubItems(9)) ' ecc
        mTDiff = mTDiff + Val(ListView1(0).ListItems(i).SubItems(12)) ' difference
    Else
        ExcludedCtr = ExcludedCtr + 1
    End If
Next i
lblclient(0).Caption = "Difference: " & mTDiff
lblclient(3).Caption = "Total EE: " & mTEE
lblclient(1).Caption = "Total ER: " & mTER
lblclient(5).Caption = "Total ECC: " & mECC
txtRemittedAmt.Text = Format(mTEE + mTER + mECC, "0.00")

Label2(21).Caption = "Excluded Employees: " & ExcludedCtr
Label2(19).Caption = "Included Employees: " & ListView1(0).ListItems.Count - ExcludedCtr

End Sub

Private Sub cmdNRDK_Click()
Dim fs As New FileSystemObject
Dim mMsg As String

On Error GoTo ExitRtn2

If ListView1(0).ListItems.Count = 0 Then
    MsgBox "No SSS records to processed.", vbInformation, PayrollAlert
    GoTo ExitRtn1
End If

mMsg = "Create NR3001DK.TXT SSS file?"

If MsgBox(mMsg, vbQuestion + vbYesNo) = vbYes Then
    If CreateNR3001DKFile = True Then
        MsgBox "File NR3001DK.TXT was successfully created in your drive C:", vbInformation, PayrollAlert
    Else
        MsgBox "Cannot create file NR3001DK.TXT in drive C:", vbInformation, PayrollAlert
        GoTo ExitRtn1
    End If
    mMsg = "Do you want to copy file NR3001DK.TXT in diskette? Insert diskette in Drive A: then select YES."
    If MsgBox(mMsg, vbQuestion + vbYesNo) = vbYes Then
        If fs.FileExists("c:\NR3001DK.TXT") Then
            fs.CopyFile "c:\NR3001DK.TXT", "a:\NR3001DK.TXT", True
            MsgBox "NR3001DK.TXT was successfully copied to disk.", vbInformation, PayrollAlert
        Else
            MsgBox "Copy cancelled!", vbInformation, PayrollAlert
        End If
    End If
Else
    MsgBox "Copy cancelled.", vbInformation, PayrollAlert
End If

ExitRtn1:

Exit Sub

ExitRtn2:

    MsgBox Err.Description & ". Cannot copy file.", vbInformation, PayrollAlert
    Err.Clear
    
End Sub

Private Function CreateNR3001DKFile() As Boolean
Dim fs As FileSystemObject
Dim NRFile As Variant
Dim i As Integer, ctr As Integer
Dim mCoName As String
Dim AppliedMonth As String
Dim mCoSSS As String
Dim mName As String
Dim mSSSNo As String
Dim SSS1 As String
Dim SSS2 As String
Dim SSS3 As String
Dim MED1 As String
Dim MED2 As String
Dim MED3 As String
Dim ECC1 As String
Dim ECC2 As String
Dim ECC3 As String
Dim TSSS1 As Double
Dim TSSS2 As Double
Dim TSSS3 As Double
Dim TECC1 As Double
Dim TECC2 As Double
Dim TECC3 As Double

TSSS1 = TSSS2 = TSSS3 = 0#
TECC1 = TECC2 = TECC3 = 0#

On Error GoTo ErrRtn2

CreateNR3001DKFile = False
i = 1
ctr = 0
    
Set fs = CreateObject("Scripting.FileSystemObject")
Set NRFile = fs.CreateTextFile("c:\NR3001DK.TXT", True)

mCoName = Left(UCase(LTrim(RTrim(txtEmployerName.Text))) & Space(30), 30)
AppliedMonth = getMonthNo(cmdAppMonth.Text) & LTrim(RTrim(cmdAppYear.Text))
mCoSSS = Replace(txtPhilHealthNo.Text, "-", "")
NRFile.WriteLine ("00" & mCoName & AppliedMonth & mCoSSS)

Do While ListView1(0).ListItems.Count >= i
    If ListView1(0).ListItems.Item(i).Checked = True Then
        GoTo Next1
    Else
    'NRFile.WriteLine ("This is a test.")
    mName = Left(ListView1(0).ListItems(i).SubItems(1) + Space(15), 15)
    mName = mName & Left(ListView1(0).ListItems(i).SubItems(2) + Space(15), 15)
    mName = mName & IIf(ListView1(0).ListItems(i).SubItems(3) <> "", Left(ListView1(0).ListItems(i).SubItems(3), 1), " ")
    mSSSNo = Left(Replace(Replace(ListView1(0).ListItems(i).SubItems(4), "-", ""), "_", "1") & Space(10), 10)
    
    SSS1 = "   0.00"
    SSS2 = "   0.00"
    SSS3 = "   0.00"
    ECC1 = " 0.00"
    ECC2 = " 0.00"
    ECC3 = " 0.00"
    MED1 = " 0.00"
    MED2 = " 0.00"
    MED3 = " 0.00"
    
    Select Case DTRNo
        Case 1
            SSS1 = Right(Space(7) & Format(Val(ListView1(0).ListItems(i).SubItems(10)) + Val(ListView1(0).ListItems(i).SubItems(11)), "0.00"), 7)
            ECC1 = Right(Space(5) & Format(Val(ListView1(0).ListItems(i).SubItems(9)), "0.00"), 5)
            MED1 = " 0.00"
            TSSS1 = TSSS1 + Val(ListView1(0).ListItems(i).SubItems(10)) + Val(ListView1(0).ListItems(i).SubItems(11))
            TECC1 = TECC1 + Val(ListView1(0).ListItems(i).SubItems(9))
        Case 2
            SSS2 = Right(Space(7) & Format(Val(ListView1(0).ListItems(i).SubItems(10)) + Val(ListView1(0).ListItems(i).SubItems(11)), "0.00"), 7)
            ECC2 = Right(Space(5) & Format(ListView1(0).ListItems(i).SubItems(9), "0.00"), 5)
            MED2 = " 0.00"
            TSSS2 = TSSS2 + Val(ListView1(0).ListItems(i).SubItems(10)) + Val(ListView1(0).ListItems(i).SubItems(11))
            TECC2 = TECC2 + Val(ListView1(0).ListItems(i).SubItems(9))
        Case 3
            SSS3 = Right(Space(7) & Format(Val(ListView1(0).ListItems(i).SubItems(10)) + Val(ListView1(0).ListItems(i).SubItems(11)), "0.00"), 7)
            ECC3 = Right(Space(5) & Format(ListView1(0).ListItems(i).SubItems(9), "0.00"), 5)
            MED3 = " 0.00"
            TSSS2 = TSSS2 + Val(ListView1(0).ListItems(i).SubItems(10)) + Val(ListView1(0).ListItems(i).SubItems(11))
            TECC2 = TECC2 + Val(ListView1(0).ListItems(i).SubItems(9))
            ctr = ctr + 1
    End Select
                
'   if !empty(date_hired)
'      n_hired = '1'
'      d_hired = allt(Left(date_hired, 2) + subs(date_hired, 4, 2) + Right(date_hired, 2))
'   Else
'      n_hired = ' '
'      d_hired = spac(6)
'   End If
                
    NRFile.WriteLine "20" & mName & mSSSNo & Space(1) & SSS1 & Space(1) & _
        SSS2 & Space(1) & SSS3 & Space(1) & MED1 & Space(1) & MED2 & Space(1) & _
        MED3 & Space(1) & ECC1 & Space(1) & ECC2 & Space(1) & ECC3 & "      N" & Space(8)
    
    End If
    
'        Set litem1 = frmSSSReport.ListView1(0).ListItems.Add(, , IIf(IsNull(!ClientName), "Blank", !ClientName))
'        litem1.SubItems(1) = IIf(IsNull(!Surname), "", !Surname)
'        litem1.SubItems(2) = IIf(IsNull(!FirstName), "", !FirstName)
'        litem1.SubItems(3) = IIf(IsNull(!Middlename), "", !Middlename)
'        litem1.SubItems(4) = IIf(IsNull(!PHealthSSSNo), "", !PHealthSSSNo)
'        litem1.SubItems(5) = Format(!BasicPay, "0.##")
'        litem1.SubItems(6) = IIf(IsNull(!SalaryBracket), "", !SalaryBracket)
'        litem1.SubItems(7) = Format(!ActualEmployeeDed, "0.##")
'        litem1.SubItems(8) = Format(!ActualEmployerDed, "0.##")
'        litem1.SubItems(9) = Format(!AdjustedECC, "0.##")
'        litem1.SubItems(10) = Format(!AdjustedEmployeeDed, "0.##")
'        litem1.SubItems(11) = Format(!AdjustedEmployerDed, "0.##")
'        litem1.SubItems(12) = Format(!AmtDifference, "0.##")

Next1:
        i = i + 1

Loop
    
' Print the total.

SSS1 = Right(Space(11) & Format(TSSS1, "0.00"), 11)
SSS2 = Right(Space(11) & Format(TSSS2, "0.00"), 11)
SSS3 = Right(Space(11) & Format(TSSS3, "0.00"), 11)

ECC1 = Right(Space(11) & Format(TECC1, "0.00"), 9)
ECC2 = Right(Space(11) & Format(TECC2, "0.00"), 9)
ECC3 = Right(Space(11) & Format(TECC3, "0.00"), 9)

MED1 = Right(Space(11) & "0.00", 9)
MED2 = Right(Space(11) & "0.00", 9)
MED3 = Right(Space(11) & "0.00", 9)

NRFile.WriteLine "99" & Space(1) & SSS1 & Space(1) & _
    SSS2 & Space(1) & SSS3 & Space(1) & MED1 & Space(1) & MED2 & Space(1) & _
    MED3 & Space(1) & ECC1 & Space(1) & ECC2 & Space(1) & ECC3 '& Space(20)
    
'=FPUTS(outfile, "99"+spac(1)+t_sss1+spac(1)+t_sss2+spac(1)+t_sss3+ ;
'       spac(1)+t_med1+spac(1)+t_med2+spac(1)+t_med3+spac(1)+t_ec1+ ;
'       spac(1)+t_ec2+spac(1)+t_ec3+spac(20))
   
    
'00DATASCOPE COMM. (PHIL.), INC. 0720050387855347
'20DAYAPERA       ROSEMARIE      R0369301725  517.00    0.00    0.00  0.00  0.00  0.00 10.00  0.00  0.00072005N
'20KKKK           LLLLL          T0369301725  611.00    0.00    0.00  0.00  0.00  0.00 10.00  0.00  0.00072005N
'20PONFERRADA     ESTRELLA       A0394348258  470.00    0.00    0.00  0.00  0.00  0.00 10.00  0.00  0.00072005N
'20SAHAGUN        BERNARD        O0369301725  564.00    0.00    0.00  0.00  0.00  0.00 10.00  0.00  0.00072005N
'20TEST           TESTING        T0369301725  611.00    0.00    0.00  0.00  0.00  0.00 10.00  0.00  0.00072005N
'20TEST2          TESTING2       T0369301725  611.00    0.00    0.00  0.00  0.00  0.00 10.00  0.00  0.00072005N
'20VILLONCO       MONICA SR       0374427115 1008.00    0.00    0.00  0.00  0.00  0.00 10.00  0.00  0.00072005N
'99     4392.00        0.00        0.00      0.00      0.00      0.00     70.00      0.00      0.00
    
'=FPUTS(outfile, "20"+lname+spac(15-len(lname))+fname+spac(15-len(fname)) + ;
'           mname+trans(sssnum,repl('X',10))+spac(1)+c_sss1+spac(1)+c_sss2+spac(1)+ ;
'           c_sss3+spac(1)+c_med1+spac(1)+c_med2+spac(1)+c_med3+spac(1)+c_ec1+ ;
'           spac(1)+c_ec2+spac(1)+c_ec3+app_date+n_hired+d_hired+filler)
    
    
'   c_sss1 = str(sss, 9, 2)
'   c_sss1 = '000' + ltrim(strtran(c_sss1, "."))
'   c_sss1 = right(c_sss1,6)  && SSS contributions value for ex. "037800"
'
'   *** Convert Med Amount to '07500' ***
'   c_med1 = str(med, 9, 2)
'   c_med1 = '000' + ltrim(strtran(c_med1, "."))
'   c_med1 = right(c_med1,5)  && Med amount value for ex. "07500"
'
'   *** Convert EC Amount to '01000' ***
'   c_ec1 = str(ec, 7, 2)
'   c_ec1 = '000' + ltrim(strtran(c_ec1, "."))
'   c_ec1 = right(c_ec1,5)  && EC amount value for ex. "01000"
    
'**** For Third Month ****
't_sss3 = str(tot_sss3, 9, 2)
't_sss3 = '0000000000' + ltrim(strtran(t_sss3, "."))
't_sss3 = right(t_sss3,10)  && Total SSS contributions value for ex. "0001437800"
'
't_med3 = str(tot_med3, 9, 2)
't_med3 = '0000000000' + ltrim(strtran(t_med3, "."))
't_med3 = right(t_med3,9)  && Total Med amount value for ex. "000077500"
'
't_ec3 = str(tot_ec3, 7, 2)
't_ec3 = '0000000000' + ltrim(strtran(t_ec3, "."))
't_ec3 = right(t_ec3,9)  && Total EC amount value for ex. "000121000"
    
    
'=FPUTS(outfile, "99"+spac(1)+t_sss1+spac(1)+t_sss2+spac(1)+t_sss3+ ;
'       spac(1)+t_med1+spac(1)+t_med2+spac(1)+t_med3+spac(1)+t_ec1+ ;
'       spac(1)+t_ec2+spac(1)+t_ec3+spac(20))
'

    NRFile.Close

'If fs.FileExists("c:\nr3001dk.txt") Then
'    fs.CopyFile "c:\NR3001dk.TXT", "a:\NR3001DK.TXT", True
'Else
'    MsgBox "Cannot find file NR3001DK.TXT!", vbInformation, PayrollAlert
'End If

CreateNR3001DKFile = True

Exit Function

ErrRtn2:

    MsgBox Err.Description
    Err.Clear
'
''FileSystemObject.CopyFile "c:\mydocuments\letters\*.doc", "c:\tempfolder\"
'

End Function

Private Sub cmdR1AReport_Click()

If ListView1(2).ListItems.Count >= 1 Then
    Call GenerateR1AReports
Else
    MsgBox "No records to processed.", vbInformation, PayrollAlert
End If

End Sub

Private Sub cmdRetrieve_Click()
Dim litem1 As ListItem
Dim Listctr As Integer
Dim rsPHList As ADODB.Recordset
Dim i As Integer

Listctr = 0
If chkR1A.VALUE = 1 Then   ' R1A Report
    StrQuery = "SELECT * FROM R1AContributions WHERE GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "' ORDER BY Surname, Firstname"
    Set rsPHList = ReturnRst(StrQuery, cnnstr, mstatic)
    If rsPHList.RecordCount = 0 Then
        ListView1(0).ListItems.Clear
        MsgBox "No previous contributions retrieve.", vbInformation, PayrollAlert
        lblclient(0).Caption = "Difference: 0.00"
        lblclient(3).Caption = "Total EE: 0.00"
        lblclient(5).Caption = "Total ECC: 0.00"
        lblclient(1).Caption = "Total ER: 0.00"
        txtRemittedAmt.Text = "0.00"
        GoTo Exit22
    End If
    
    WithSBRNo = False
    i = 1
    With rsPHList
    ListView1(2).ListItems.Clear
    While Not (.EOF Or .BOF)
        Set litem1 = frmSSSReport.ListView1(2).ListItems.Add(, , IIf(IsNull(!ClientName), "Blank", !ClientName))
            litem1.SubItems(1) = IIf(IsNull(!Surname), "", !Surname)
            litem1.SubItems(2) = IIf(IsNull(!FirstName), "", !FirstName)
            litem1.SubItems(3) = IIf(IsNull(!MiddleName), "", !MiddleName)
            litem1.SubItems(4) = IIf(IsNull(!PHealthSSSNo), "", !PHealthSSSNo)
            
            litem1.SubItems(5) = IIf(!BIRTHDATE = #1/1/1900#, "__/__/____", Format(!BIRTHDATE, "mm/dd/yyyy"))
            litem1.SubItems(6) = Format(!MonthlyRate, "0.00")
            litem1.SubItems(7) = IIf(IsNull(!empPosition), "", !empPosition)
            litem1.SubItems(8) = IIf(!StartContract = #1/1/1900#, "__/__/____", Format(!StartContract, "mm/dd/yyyy"))
            litem1.SubItems(9) = IIf(!endcontract = #1/1/1900#, "__/__/____", Format(!endcontract, "mm/dd/yyyy"))
            litem1.SubItems(12) = IIf(IsNull(!ClientName), "Blank", !ClientName)
            litem1.SubItems(13) = IIf(IsNull(!groupcode), "", !groupcode)
            If !ExcludedFlag = True Then
                frmSSSReport.ListView1(2).ListItems.Item(i).Checked = True
            End If
        i = i + 1
        .MoveNext
    Wend
    End With
    Set rsPHList = Nothing
    Label2(19).Caption = "Total Employees: " & ListView1(0).ListItems.Count
Else
    'strQuery = "SELECT * FROM SSSContributions WHERE ApplicableMonth = '" & cmdAppMonth.Text & "' and ApplicableYear = " & cmdAppYear.Text & " AND ProcessedBy = '" & mUserID & "' AND GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "' ORDER BY Surname, Firstname"
    StrQuery = "SELECT * FROM SSSContributions WHERE GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "' ORDER BY Surname, Firstname"
    Set rsPHList = ReturnRst(StrQuery, cnnstr, mstatic)
    If rsPHList.RecordCount = 0 Then
        ListView1(0).ListItems.Clear
        MsgBox "No previous contributions retrieve.", vbInformation, PayrollAlert
        lblclient(0).Caption = "Difference: 0.00"
        lblclient(3).Caption = "Total EE: 0.00"
        lblclient(5).Caption = "Total ECC: 0.00"
        lblclient(1).Caption = "Total ER: 0.00"
        txtRemittedAmt.Text = "0.00"
        GoTo Exit22
    End If
    
    WithSBRNo = False
    i = 1
    With rsPHList
    ListView1(0).ListItems.Clear
    While Not (.EOF Or .BOF)
        Set litem1 = frmSSSReport.ListView1(0).ListItems.Add(, , IIf(IsNull(!ClientName), "Blank", !ClientName))
            litem1.SubItems(1) = IIf(IsNull(!Surname), "", !Surname)
            litem1.SubItems(2) = IIf(IsNull(!FirstName), "", !FirstName)
            litem1.SubItems(3) = IIf(IsNull(!MiddleName), "", !MiddleName)
            litem1.SubItems(4) = IIf(IsNull(!PHealthSSSNo), "", !PHealthSSSNo)
            litem1.SubItems(5) = Format(!BasicPay, "0.##")
            litem1.SubItems(6) = IIf(IsNull(!SalaryBracket), "", !SalaryBracket)
            litem1.SubItems(7) = Format(!ActualEmployeeDed, "0.##")
            litem1.SubItems(8) = Format(!ActualEmployerDed, "0.##")
            litem1.SubItems(9) = Format(!AdjustedECC, "0.##")
            litem1.SubItems(10) = Format(!AdjustedEmployeeDed, "0.##")
            litem1.SubItems(11) = Format(!AdjustedEmployerDed, "0.##")
            litem1.SubItems(12) = Format(!AmtDifference, "0.##")
            litem1.SubItems(13) = IIf(IsNull(!ClientName), "Blank", !ClientName)
            litem1.SubItems(14) = Format(!TableBasedEE, "0.##")
            litem1.SubItems(15) = Format(!TableBasedER, "0.##")
            litem1.SubItems(16) = IIf(IsNull(!groupcode), "", !groupcode)
            litem1.SubItems(17) = Format(!AmountPaid, "0.00")
            litem1.SubItems(18) = IIf(IsNull(!SBRNo), "", !SBRNo)
            litem1.SubItems(19) = IIf(!DatePaid = #1/1/1900#, "__/__/____", Format(!DatePaid, "mm/dd/yyyy"))
            litem1.SubItems(20) = IIf(!DatePosted = #1/1/1900#, "__/__/____", Format(!DatePosted, "mm/dd/yyyy"))
            If Not IsNull(!SBRNo) And !SBRNo <> "" And !SBRNo <> "0" Then
                WithSBRNo = True
            End If
            If !ExcludedFlag = True Then
                frmSSSReport.ListView1(0).ListItems.Item(i).Checked = True
                'frmPHealthReport.ListView1(1).ListItems.Checked = True
            End If
        i = i + 1
        .MoveNext
    Wend
    End With
    
    Set rsPHList = Nothing
    
    Label2(19).Caption = "Total Employees: " & ListView1(0).ListItems.Count
    Call GetTotalSSS

End If

Exit22:

End Sub

Private Function FillGroupCode(ByRef mCmb As ComboBox, _
                        mRs As Recordset, Optional FirstGroup As String = "")
    mCmb.Clear
    While Not (mRs.EOF Or mRs.BOF)
        mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields("GroupCode")), FirstGroup, mRs.Fields("GroupCode")), "'", "''"))
        'mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
        FirstGroup = UCase(Left(mRs.Fields("GroupCode"), 13)) & "-" & Right("00" & Right(LTrim(RTrim(mRs.Fields("GroupCode"))), 2) + 1, 2) & " New"
        mRs.MoveNext
    Wend
    mCmb.AddItem FirstGroup
    
End Function

Private Sub cmdSBRReport_Click()

ListView1(0).SortKey = 0
ListView1(0).Sorted = True
ListView1(0).Refresh
If ListView1(0).ListItems.Count >= 1 Then
    Call GenerateSBRReports
Else
    MsgBox "No SSS records to processed.", vbInformation, PayrollAlert
End If
ListView1(0).SortKey = 1
ListView1(0).Sorted = True
ListView1(0).Refresh

End Sub

Private Sub cmdSearchSBR_Click()
Dim rsPHList As ADODB.Recordset

If chkAllPayments.VALUE = 1 Then ' all payments regardless of users
    StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy, SBRNo " & _
        " FROM SSSContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy, SBRNo HAVING SBRNo = '" & txtSBRNo.Text & "' AND GroupCode <> ''" & _
        " ORDER BY ProcessedBy, GroupCode"
Else    ' for current login only
    StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy, SBRNo " & _
        " FROM SSSContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy, SBRNo HAVING SBRNo = '" & txtSBRNo.Text & "' AND GroupCode <> ''" & _
        " AND ProcessedBy = '" & mUserID & "' ORDER BY ProcessedBy, GroupCode"
End If
Set rsPHList = ReturnRst(StrQuery, cnnstr, mstatic)
If rsPHList.RecordCount = 0 Then
    MsgBox "No payments found for SBR#: " & txtSBRNo.Text, vbInformation, PayrollAlert
    rsPHList.Close
Else
    SSSGroupCode = UCase(Left(mUserID, 5)) & "-" & Left(cmdAppMonth, 3) & cmdAppYear & "-01 New"
    Call FillGroupCode(cmbGroupCode, rsPHList, SSSGroupCode)
    cmbGroupCode.ListIndex = 0
    rsPHList.Close
    Call cmdRetrieve_Click
    MsgBox "SBR# " & txtSBRNo.Text & " successfully found!", vbInformation, PayrollAlert
End If
End Sub

Private Sub cmdSelect_Click()
    'Call cmdRetrieve_Click
'    If Right(LTrim(cmbGroupCode.Text), 3) = "New" Then
'        If MsgBox("Create new Group of Payment?", vbQuestion + vbYesNo) = vbYes Then
'            ListView1(0).ListItems.Clear
'            PHealthFlag = False
'            Call frmPHealthClients.Show
'            Label2(19).Caption = "Total Employees: " & ListView1(0).ListItems.Count
'            Call GetTotalSSS
'        End If
'    Else
'            PHealthFlag = False
'            Call frmPHealthClients.Show
'            Label2(19).Caption = "Total Employees: " & ListView1(0).ListItems.Count
'            Call GetTotalSSS
'    End If
            
If chkR1A.VALUE = 1 Then
    'PHealthFlag = False
    Label2(19).Caption = "Total Employees: " & ListView1(0).ListItems.Count
    Call frmHDMFClients.Show
Else
    If WithSBRNo = False Then
        PHealthFlag = False
        Call frmPHealthClients.Show
        Label2(19).Caption = "Total Employees: " & ListView1(0).ListItems.Count
        Call GetTotalSSS
    Else
        MsgBox "This group of clients was already paid. Cannot update!", vbInformation, PayrollAlert
    End If
End If
End Sub

Private Sub cmdSelectGroupCodes_Click()
    Call UpdateGroupCodes
End Sub

Private Sub UpdateGroupCodes()
Dim rsPHList As ADODB.Recordset

If chkR1A.VALUE = 1 Then
    If chkAllPayments.VALUE = 1 Then ' all payments regardless of users
        StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy " & _
            " FROM R1AContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy HAVING ApplicableMonth = '" & cmdAppMonth.Text & "' and ApplicableYear = " & cmdAppYear.Text & " AND GroupCode <> ''" & _
            " ORDER BY ProcessedBy, GroupCode"
    Else    ' for current login only
        StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy " & _
            " FROM R1AContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy HAVING ApplicableMonth = '" & cmdAppMonth.Text & "' and ApplicableYear = " & cmdAppYear.Text & " AND GroupCode <> ''" & _
            " AND ProcessedBy = '" & mUserID & "' ORDER BY ProcessedBy, GroupCode"
    End If
Else
    If chkAllPayments.VALUE = 1 Then ' all payments regardless of users
        StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy " & _
            " FROM SSSContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy HAVING ApplicableMonth = '" & cmdAppMonth.Text & "' and ApplicableYear = " & cmdAppYear.Text & " AND GroupCode <> ''" & _
            " ORDER BY ProcessedBy, GroupCode"
    Else    ' for current login only
        StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy " & _
            " FROM SSSContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy HAVING ApplicableMonth = '" & cmdAppMonth.Text & "' and ApplicableYear = " & cmdAppYear.Text & " AND GroupCode <> ''" & _
            " AND ProcessedBy = '" & mUserID & "' ORDER BY ProcessedBy, GroupCode"
    End If
End If

Set rsPHList = ReturnRst(StrQuery, cnnstr, mstatic)
SSSGroupCode = UCase(Left(mUserID, 5)) & "-" & Left(cmdAppMonth, 3) & cmdAppYear & "-01 New"
Call FillGroupCode(cmbGroupCode, rsPHList, SSSGroupCode)
cmbGroupCode.ListIndex = 0
rsPHList.Close
End Sub

Private Sub cmdUpdateSBR_Click()

If MsgBox("Update SBR#/OR#?", vbQuestion + vbYesNo, "SBR Number Updates") = vbYes Then
    StrQuery = "UPDATE SSSContributions SET SBRNO = '" & IIf(txtSBRNo.Text = "", "0", txtSBRNo.Text) & "', AmountPaid = " & IIf(txtRemittedAmt.Text = "", 0#, txtRemittedAmt.Text) & ", " & _
        " DatePaid = '" & IIf(MaskDatePaid.Text = "__/__/____", #1/1/1900#, Format(MaskDatePaid.Text, "mm/dd/yyyy")) & "', DatePosted = '" & IIf(MaskDatePosted.Text = "__/__/____", #1/1/1900#, Format(MaskDatePosted.Text, "mm/dd/yyyy")) & "' WHERE GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "'"
    If Action_Query(StrQuery, cnnstr) = True Then
        Call cmdRetrieve_Click
        MsgBox "SBR/OR # successfully updated!", vbInformation, PayrollAlert
    Else
        MsgBox "SBR/OR # was not successfully updated!", vbInformation, PayrollAlert
    End If
End If

End Sub

Private Sub Form_Activate()
    SelectCompanyFlag = False
    ReportedFlag = 2
    PHealthFlag = False
    Call UpdateGroupCodes
End Sub

Private Sub Form_Load()

    PHealthFlag = False
    SelectCompanyFlag = False
    cmdAppMonth.Text = Format(Now, "mmmm")
    cmdAppYear.Text = Format(Now, "yyyy")
    FormLoadFlag = True
    MaskDatePaid = Format(Now, "mm/dd/yyyy")
    
    If PHealthFlag = False Then  ' for SSS
        'Image1.Visible = False
        Label1(0).Caption = "SSS"
        Label1(2).Caption = "  SSS NUMBER  "
        ListView1(1).Visible = False
        ListView1(0).Visible = True
        lblclient(7).Caption = "No SSS No."
        cmdNoPHNo.Caption = "Check/Uncheck No SSS No."
        'cmdGenerateReport.Visible = False
    Else        ' for PhilHealth
        'Image1.Visible = True
        Label1(0).Caption = "RF-1"
        Label1(2).Caption = "PHILHEALTH NO."
        ListView1(1).Visible = True
        ListView1(0).Visible = False
        lblclient(7).Caption = "No PHealth No."
        cmdNoPHNo.Caption = "Check/Uncheck No PHealth No."
        'cmdGenerateReport.Visible = True
    End If
End Sub

'Private Sub ListView1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'    ListView1(Index).SortKey = ColumnHeader.Index - 1
'    ListView1(Index).Sorted = True
'    If ListView1(0).ListItems.Item(i).Checked = True Then

'End Sub

Private Function PrintNoSSSNumbers(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
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
mPCompany = mPCompany & " (Employees with no SSS#)"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mPCompany, "", xlLeft, , False, False, False, , , 9, True, 18)
'Call PrintReportHeader(xlSheet, xlApp1, mRange, mPCompany, "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)

Do While ListView1(0).ListItems.Count >= j
    If ListView1(0).ListItems(j).SubItems(13) = mCo2 And ListView1(0).ListItems.Item(j).Checked = False And _
(ListView1(0).ListItems(j).SubItems(4) = "__-_______-_" Or ListView1(0).ListItems(j).SubItems(4) = "") Then
        ctr = ctr + 1
        xlSheet.Cells(ctr, 1).VALUE = LineCtr
        xlSheet.Cells(ctr, 2).VALUE = ListView1(0).ListItems(j).SubItems(1) ' Surname
        xlSheet.Cells(ctr, 3).VALUE = ListView1(0).ListItems(j).SubItems(2) ' firstname
        xlSheet.Cells(ctr, 4).VALUE = ListView1(0).ListItems(j).SubItems(3) ' Middlename
        xlSheet.Cells(ctr, 5).VALUE = ListView1(0).ListItems(j).SubItems(4) ' sss no
        xlSheet.Cells(ctr, 6).VALUE = ListView1(0).ListItems(j).SubItems(7) ' employee
        xlSheet.Cells(ctr, 7).VALUE = ListView1(0).ListItems(j).SubItems(8) ' employer
        xlSheet.Cells(ctr, 8).VALUE = ListView1(0).ListItems(j).SubItems(9) ' ecc
        xlSheet.Cells(ctr, 9).VALUE = ListView1(0).ListItems(j).SubItems(10) ' employee adjust
        xlSheet.Cells(ctr, 10).VALUE = ListView1(0).ListItems(j).SubItems(11) ' employer adjust
        xlSheet.Cells(ctr, 11).VALUE = ListView1(0).ListItems(j).SubItems(12) ' difference
        LineCtr = LineCtr + 1
    End If
    j = j + 1
Loop
' Print Sub-Totals.
ctr = ctr + 1
mRange = "F" & cPos & ":K" & ctr
xlSheet.Range(mRange).Select
mRange1 = "F" & ctr
xlSheet.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

            ' Copy to Sheet 2 the Total.
            ctr2 = ctr2 + 1
            xlSheet.Range(mRange1).Select
            xlApp1.Selection.Copy
            xlApp1.Sheets("Sheet2").Select
            mRange1 = "D" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            mRange1 = "A" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
            mRange1 = "B" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = mPCompany
            mRange1 = "C" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
            xlApp1.Sheets("Sheet1").Select


mRange = "F" & cPos & ":K" & ctr
xlSheet.Range(mRange).Select
mRange1 = "G" & ctr
xlSheet.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

            ' Copy to Sheet 2 the Total.
            xlSheet.Range(mRange1).Select
            xlApp1.Selection.Copy
            xlApp1.Sheets("Sheet2").Select
            mRange1 = "E" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            xlApp1.Sheets("Sheet1").Select


mRange = "F" & cPos & ":K" & ctr
xlSheet.Range(mRange).Select
mRange1 = "H" & ctr
xlSheet.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

            ' Copy to Sheet 2 the Total.
            xlSheet.Range(mRange1).Select
            xlApp1.Selection.Copy
            xlApp1.Sheets("Sheet2").Select
            mRange1 = "F" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            xlApp1.Sheets("Sheet1").Select

mRange = "F" & cPos & ":K" & ctr
xlSheet.Range(mRange).Select
mRange1 = "I" & ctr
xlSheet.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

            ' Copy to Sheet 2 the Total.
            xlSheet.Range(mRange1).Select
            xlApp1.Selection.Copy
            xlApp1.Sheets("Sheet2").Select
            mRange1 = "G" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            xlApp1.Sheets("Sheet1").Select

mRange = "F" & cPos & ":K" & ctr
xlSheet.Range(mRange).Select
mRange1 = "J" & ctr
xlSheet.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

            ' Copy to Sheet 2 the Total.
            xlSheet.Range(mRange1).Select
            xlApp1.Selection.Copy
            xlApp1.Sheets("Sheet2").Select
            mRange1 = "H" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            xlApp1.Sheets("Sheet1").Select

mRange = "F" & cPos & ":K" & ctr
xlSheet.Range(mRange).Select
mRange1 = "K" & ctr
xlSheet.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

            ' Copy to Sheet 2 the Total.
            xlSheet.Range(mRange1).Select
            xlApp1.Selection.Copy
            xlApp1.Sheets("Sheet2").Select
            mRange1 = "I" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            xlApp1.Sheets("Sheet1").Select

mRange = "F" & cPos & ":K" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":K" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

If SumTotal = "" Then
    SumTotal = "=F" & ctr
Else
    SumTotal = SumTotal & "+F" & ctr
End If

PrintNoSSSNumbers = ctr

End Function

Private Sub GenerateSSSReports()
Dim rsDedRecords As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mRange As String, mRange1 As String
Dim ctr As Long
Dim i As Integer
Dim cPos As Integer
Dim mHead As String
Dim HeaderFlag As Boolean
Dim NoSSSCtr As Integer
Dim ExcludedCtr As Integer
Dim LineCtr As Integer
Dim WithRecSw As Boolean

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)
'Set xlSheet = xlBook1.Worksheets(2)

WithRecSw = False
SumTotal = ""
NoSSSCtr = 0: ExcludedCtr = 0
HeaderFlag = False
LineCtr = 1
LineCtr2 = 1
LineCtr3 = 1
ctr = 3
ctr2 = 6    '7
ctr3 = 6
cPos = 9
i = 1

'Print Header.
xlApp1.Sheets("Sheet1").Select
mHead = CompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 12, True)
mHead = "SSS Contributions Report for the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)

'Print Header of Sheet2 for Remitted Clients Summary.
mHead = CompanyName & " (" & CompanyShort & ")"
xlApp1.Sheets("Sheet2").Select

xlApp1.Range("A2").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "SSS Contributions Summary Report per Clients"
xlApp1.Range("A3").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "For the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text
xlApp1.Range("A4").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead

xlApp1.Range("A6").Select
xlApp1.ActiveCell.FormulaR1C1 = "No."
xlApp1.Range("B6").Select
xlApp1.ActiveCell.FormulaR1C1 = "Client Name"
xlApp1.Range("C6").Select
xlApp1.ActiveCell.FormulaR1C1 = "Total Employees"
xlApp1.Range("D6").Select
xlApp1.ActiveCell.FormulaR1C1 = "EE Share"
xlApp1.Range("E6").Select
xlApp1.ActiveCell.FormulaR1C1 = "ER Share"
xlApp1.Range("F6").Select
xlApp1.ActiveCell.FormulaR1C1 = "ECC"
xlApp1.Range("G6").Select
xlApp1.ActiveCell.FormulaR1C1 = "Adjusted EE"
xlApp1.Range("H6").Select
xlApp1.ActiveCell.FormulaR1C1 = "Adjusted ER"
xlApp1.Range("I6").Select
xlApp1.ActiveCell.FormulaR1C1 = "Difference"

xlApp1.Range("A6:I6").Select
xlApp1.Selection.Interior.ColorIndex = 15
xlApp1.Selection.Interior.Pattern = xlSolid

'Print Header of Sheet3 for Unremitted Clients Summary.
xlApp1.Sheets("Sheet3").Select

mHead = CompanyName & " (" & CompanyShort & ")"
xlApp1.Range("A2").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "SSS Contributions Summary Report (NOT TO BE REMITTED CLIENTS)"
xlApp1.Range("A3").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "For the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text
xlApp1.Range("A4").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead

xlApp1.Range("A6").Select
xlApp1.ActiveCell.FormulaR1C1 = "No."
xlApp1.Range("B6").Select
xlApp1.ActiveCell.FormulaR1C1 = "Client Name"
xlApp1.Range("C6").Select
xlApp1.ActiveCell.FormulaR1C1 = "Total Employees"
xlApp1.Range("D6").Select
xlApp1.ActiveCell.FormulaR1C1 = "EE Share"
xlApp1.Range("E6").Select
xlApp1.ActiveCell.FormulaR1C1 = "ER Share"
xlApp1.Range("F6").Select
xlApp1.ActiveCell.FormulaR1C1 = "ECC"
xlApp1.Range("G6").Select
xlApp1.ActiveCell.FormulaR1C1 = "Adjusted EE"
xlApp1.Range("H6").Select
xlApp1.ActiveCell.FormulaR1C1 = "Adjusted ER"
xlApp1.Range("I6").Select
xlApp1.ActiveCell.FormulaR1C1 = "Difference"

xlApp1.Range("A6:I6").Select
xlApp1.Selection.Interior.ColorIndex = 15
xlApp1.Selection.Interior.Pattern = xlSolid

xlApp1.Sheets("Sheet1").Select

PrevCompany = ListView1(0).ListItems(i).SubItems(13)
Do While ListView1(0).ListItems.Count >= i

    If PrevCompany <> ListView1(0).ListItems(i).SubItems(13) Then
        If LineCtr > 1 Then
            ' Print Sub-Totals.
            ctr = ctr + 1
            mRange = "F" & cPos + 1 & ":K" & ctr
            xlSheet.Range(mRange).Select
            mRange1 = "F" & ctr
            xlSheet.Range(mRange1).Activate
            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
            
            ' Copy to Sheet 2 the Total.
            ctr2 = ctr2 + 1
            xlSheet.Range(mRange1).Select
            xlApp1.Selection.Copy
            xlApp1.Sheets("Sheet2").Select
            mRange1 = "D" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            mRange1 = "A" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
            mRange1 = "B" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
            mRange1 = "C" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
            xlApp1.Sheets("Sheet1").Select
            
            xlSheet.Range(mRange).Select
            mRange1 = "G" & ctr
            xlSheet.Range(mRange1).Activate
            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
            
            ' Copy to Sheet 2 the Total.
            xlSheet.Range(mRange1).Select
            xlApp1.Selection.Copy
            xlApp1.Sheets("Sheet2").Select
            mRange1 = "E" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            xlApp1.Sheets("Sheet1").Select
            
            xlSheet.Range(mRange).Select
            mRange1 = "H" & ctr
            xlSheet.Range(mRange1).Activate
            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
            
            ' Copy to Sheet 2 the Total.
            xlSheet.Range(mRange1).Select
            xlApp1.Selection.Copy
            xlApp1.Sheets("Sheet2").Select
            mRange1 = "F" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            xlApp1.Sheets("Sheet1").Select
            
            xlSheet.Range(mRange).Select
            mRange1 = "I" & ctr
            xlSheet.Range(mRange1).Activate
            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
            
            ' Copy to Sheet 2 the Total.
            xlSheet.Range(mRange1).Select
            xlApp1.Selection.Copy
            xlApp1.Sheets("Sheet2").Select
            mRange1 = "G" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            xlApp1.Sheets("Sheet1").Select
            
            xlSheet.Range(mRange).Select
            mRange1 = "J" & ctr
            xlSheet.Range(mRange1).Activate
            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
            
            ' Copy to Sheet 2 the Total.
            xlSheet.Range(mRange1).Select
            xlApp1.Selection.Copy
            xlApp1.Sheets("Sheet2").Select
            mRange1 = "H" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            xlApp1.Sheets("Sheet1").Select
            
            xlSheet.Range(mRange).Select
            mRange1 = "K" & ctr
            xlSheet.Range(mRange1).Activate
            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
            
            ' Copy to Sheet 2 the Total.
            xlSheet.Range(mRange1).Select
            xlApp1.Selection.Copy
            xlApp1.Sheets("Sheet2").Select
            mRange1 = "I" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            xlApp1.Sheets("Sheet1").Select
            LineCtr2 = LineCtr2 + 1
            
            mRange = "F" & cPos & ":K" & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            mRange = "F" & ctr & ":K" & ctr
            xlSheet.Range(mRange).Select
            xlApp1.Range(mRange).Font.Bold = True
            
            If SumTotal = "" Then
                SumTotal = "=F" & ctr
            Else
                SumTotal = SumTotal & "+F" & ctr
            End If
        End If
        
        If NoSSSCtr > 0 Then
            ctr = PrintNoSSSNumbers(xlSheet, xlApp1, PrevCompany, ctr)
            LineCtr2 = LineCtr2 + 1
        End If

        'ctr = ctr + 3
        PrevCompany = ListView1(0).ListItems(i).SubItems(13)
'        mRange = "A" & ctr
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
        HeaderFlag = False
        NoSSSCtr = 0  ' set again to 0 after printing
    End If
    ' Print Header.
    If HeaderFlag = False Then
        'PrevCompany = ListView1(0).ListItems(i).SubItems(13)
        LineCtr = 1
        ctr = ctr + 3
        mRange = "A" & ctr
        Call PrintReportHeader(xlSheet, xlApp1, mRange, PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
        'Call PrintReportHeader(xlSheet, xlApp1, "A6", PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
        ctr = ctr + 1
        mRange = "A" & ctr & ":A" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "A7:A8", "No.", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 4.14)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "No.", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 4.14)
        mRange = "B" & ctr & ":B" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "B7:B8", "Surname", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Surname", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)
        mRange = "C" & ctr & ":C" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "C7:C8", "Firstname", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Firstname", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
        mRange = "D" & ctr & ":D" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "D7:D8", "Middlename", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Middlename", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
        mRange = "E" & ctr & ":E" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "E7:E8", "SSS Number", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 11)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "SSS Number", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 11)
        mRange = "F" & ctr & ":F" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "F7:F8", "Employee Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Employee Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "G" & ctr & ":G" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "G7:G8", "Employer Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Employer Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "H" & ctr & ":H" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "H7:H8", "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
        mRange = "I" & ctr & ":I" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "I7:I8", "Adjusted", "PS Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "PS Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "J" & ctr & ":J" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "J7:J8", "Adjusted", "ES Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "ES Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "K" & ctr & ":K" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "K7:K8", "Difference", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        HeaderFlag = True
        ctr = ctr + 1
        cPos = ctr + 1
        WithRecSw = False
    End If
    
    If ListView1(0).ListItems.Item(i).Checked = True Then
        ' Count those excluded records.
        ExcludedCtr = ExcludedCtr + 1
        GoTo Next1
    ElseIf PrevCompany = ListView1(0).ListItems(i).SubItems(13) And (ListView1(0).ListItems(i).SubItems(4) = "__-_______-_" Or ListView1(0).ListItems(i).SubItems(4) = "") Then
        ' Count those no sss number.
        NoSSSCtr = NoSSSCtr + 1
        GoTo Next1
    Else
        ctr = ctr + 1
        xlSheet.Cells(ctr, 1).VALUE = LineCtr
        xlSheet.Cells(ctr, 2).VALUE = ListView1(0).ListItems(i).SubItems(1) ' Surname
        xlSheet.Cells(ctr, 3).VALUE = ListView1(0).ListItems(i).SubItems(2) ' firstname
        xlSheet.Cells(ctr, 4).VALUE = ListView1(0).ListItems(i).SubItems(3) ' Middlename
        xlSheet.Cells(ctr, 5).VALUE = ListView1(0).ListItems(i).SubItems(4) ' sss no
        xlSheet.Cells(ctr, 6).VALUE = ListView1(0).ListItems(i).SubItems(7) ' employee
        xlSheet.Cells(ctr, 7).VALUE = ListView1(0).ListItems(i).SubItems(8) ' employer
        xlSheet.Cells(ctr, 8).VALUE = ListView1(0).ListItems(i).SubItems(9) ' ecc
        xlSheet.Cells(ctr, 9).VALUE = ListView1(0).ListItems(i).SubItems(10) ' employee adjust
        xlSheet.Cells(ctr, 10).VALUE = ListView1(0).ListItems(i).SubItems(11) ' employer adjust
        xlSheet.Cells(ctr, 11).VALUE = ListView1(0).ListItems(i).SubItems(12) ' difference
        LineCtr = LineCtr + 1
        WithRecSw = True
    End If

Next1:
    
    i = i + 1

Loop

If WithRecSw = True Then
    ' Print Sub-Totals.
    ctr = ctr + 1
    mRange = "F" & cPos + 1 & ":K" & ctr
    xlSheet.Range(mRange).Select
    mRange1 = "F" & ctr
    xlSheet.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
    
    ' Copy to Sheet 2 the Total.
    ctr2 = ctr2 + 1
    xlSheet.Range(mRange1).Select
    xlApp1.Selection.Copy
    xlApp1.Sheets("Sheet2").Select
    mRange1 = "D" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    mRange1 = "A" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
    mRange1 = "B" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
    
    mRange1 = "C" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
    xlApp1.Sheets("Sheet1").Select
    
    mRange = "F" & cPos + 1 & ":K" & ctr
    xlSheet.Range(mRange).Select
    mRange1 = "G" & ctr
    xlSheet.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
    
    ' Copy to Sheet 2 the Total.
    xlSheet.Range(mRange1).Select
    xlApp1.Selection.Copy
    xlApp1.Sheets("Sheet2").Select
    mRange1 = "E" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False
    xlApp1.Sheets("Sheet1").Select
    
    mRange = "F" & cPos + 1 & ":K" & ctr
    xlSheet.Range(mRange).Select
    mRange1 = "H" & ctr
    xlSheet.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
    
    ' Copy to Sheet 2 the Total.
    xlSheet.Range(mRange1).Select
    xlApp1.Selection.Copy
    xlApp1.Sheets("Sheet2").Select
    mRange1 = "F" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    xlApp1.Sheets("Sheet1").Select
    
    mRange = "F" & cPos + 1 & ":K" & ctr
    xlSheet.Range(mRange).Select
    mRange1 = "I" & ctr
    xlSheet.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
    
    ' Copy to Sheet 2 the Total.
    xlSheet.Range(mRange1).Select
    xlApp1.Selection.Copy
    xlApp1.Sheets("Sheet2").Select
    mRange1 = "G" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    xlApp1.Sheets("Sheet1").Select
    
    mRange = "F" & cPos + 1 & ":K" & ctr
    xlSheet.Range(mRange).Select
    mRange1 = "J" & ctr
    xlSheet.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
    
    ' Copy to Sheet 2 the Total.
    xlSheet.Range(mRange1).Select
    xlApp1.Selection.Copy
    xlApp1.Sheets("Sheet2").Select
    mRange1 = "H" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    xlApp1.Sheets("Sheet1").Select
    
    mRange = "F" & cPos + 1 & ":K" & ctr
    xlSheet.Range(mRange).Select
    mRange1 = "K" & ctr
    xlSheet.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
    
    ' Copy to Sheet 2 the Total.
    xlSheet.Range(mRange1).Select
    xlApp1.Selection.Copy
    xlApp1.Sheets("Sheet2").Select
    mRange1 = "I" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    xlApp1.Sheets("Sheet1").Select
    
    
    mRange = "F" & cPos & ":K" & ctr
    xlSheet.Range(mRange).Select
    xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
    mRange = "F" & ctr & ":K" & ctr
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
    mRange1 = "A" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
    mRange1 = "B" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
    mRange1 = "C" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = 0
    mRange1 = "D" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = 0#
    mRange1 = "E" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = 0#
    mRange1 = "F" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = 0#
    mRange1 = "G" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = 0#
    mRange1 = "H" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = 0#
    mRange1 = "I" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = 0#
    xlApp1.Sheets("Sheet1").Select
End If
If NoSSSCtr > 0 Then
    LineCtr2 = LineCtr2 + 1
    ctr = PrintNoSSSNumbers(xlSheet, xlApp1, PrevCompany, ctr)
End If

ctr = ctr + 2
mRange = "B" & ctr & ":D" & ctr
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Grand Total:", "", xlLeft, , False, False, False, , , 10, True)
mRange = "F" & ctr
xlSheet.Range(mRange).Activate
xlSheet.Cells(ctr, 6).VALUE = SumTotal
mRange = "G" & ctr
xlSheet.Range(mRange).Activate
SumTotal = Replace(SumTotal, "F", "G")
xlSheet.Cells(ctr, 7).VALUE = SumTotal
mRange = "H" & ctr
xlSheet.Range(mRange).Activate
SumTotal = Replace(SumTotal, "G", "H")
xlSheet.Cells(ctr, 8).VALUE = SumTotal
mRange = "I" & ctr
xlSheet.Range(mRange).Activate
SumTotal = Replace(SumTotal, "H", "I")
xlSheet.Cells(ctr, 9).VALUE = SumTotal
mRange = "J" & ctr
xlSheet.Range(mRange).Activate
SumTotal = Replace(SumTotal, "I", "J")
xlSheet.Cells(ctr, 10).VALUE = SumTotal
mRange = "K" & ctr
xlSheet.Range(mRange).Activate
SumTotal = Replace(SumTotal, "J", "K")
xlSheet.Cells(ctr, 11).VALUE = SumTotal

mRange = "F" & cPos & ":K" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":K" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

'xlApp1.ActiveCell.FormulaR1C1 = "=R[-45]C+R[-38]C+R[-2]C"
'sumString = "=SUM(R[" & i * -1 & "]C:R[-2]C)"

' Print Unremitted SSS Contributions.
If ExcludedCtr <> 0 Then
    ctr = UnremittedSSSContributions(xlSheet, xlApp1, ctr)
End If

' ReFormat Sheet 2.
xlApp1.Sheets("Sheet2").Select
ctr2 = ctr2 + 1
mRange = "C7:I" & ctr2
xlApp1.Range(mRange).Select
mRange1 = "C" & ctr2
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr2 - 6) * -1 & "]C:R[-1]C)"
xlApp1.Range(mRange).Select
mRange1 = "D" & ctr2
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr2 - 6) * -1 & "]C:R[-1]C)"
xlApp1.Range(mRange).Select
mRange1 = "E" & ctr2
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr2 - 6) * -1 & "]C:R[-1]C)"
xlApp1.Range(mRange).Select
mRange1 = "F" & ctr2
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr2 - 6) * -1 & "]C:R[-1]C)"
xlApp1.Range(mRange).Select
mRange1 = "G" & ctr2
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr2 - 6) * -1 & "]C:R[-1]C)"
xlApp1.Range(mRange).Select
mRange1 = "H" & ctr2
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr2 - 6) * -1 & "]C:R[-1]C)"
xlApp1.Range(mRange).Select
mRange1 = "I" & ctr2
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr2 - 6) * -1 & "]C:R[-1]C)"

mRange = "A6:I" & ctr2
xlApp1.Range(mRange).Select
With xlApp1
    '.Selection.Columns.AutoFit
    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With xlApp1.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End With
With xlApp1.Range(mRange)
    .Font.Name = "Arial"   'Font.Name
    .Font.Size = 8         'Font.size
End With

mRange = "D7:I" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "C" & ctr2 & ":I" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

mRange = "A6:I" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Selection.Columns.AutoFit
xlApp1.Range("A1").Select

' ReFormat Sheet 3.
xlApp1.Sheets("Sheet3").Select
ctr3 = ctr3 + 1
mRange = "C7:I" & ctr3
xlApp1.Range(mRange).Select
mRange1 = "C" & ctr3
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr3 - 6) * -1 & "]C:R[-1]C)"
xlApp1.Range(mRange).Select
mRange1 = "D" & ctr3
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr3 - 6) * -1 & "]C:R[-1]C)"
xlApp1.Range(mRange).Select
mRange1 = "E" & ctr3
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr3 - 6) * -1 & "]C:R[-1]C)"
xlApp1.Range(mRange).Select
mRange1 = "F" & ctr3
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr3 - 6) * -1 & "]C:R[-1]C)"
xlApp1.Range(mRange).Select
mRange1 = "G" & ctr3
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr3 - 6) * -1 & "]C:R[-1]C)"
xlApp1.Range(mRange).Select
mRange1 = "H" & ctr3
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr3 - 6) * -1 & "]C:R[-1]C)"
xlApp1.Range(mRange).Select
mRange1 = "I" & ctr3
xlApp1.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr3 - 6) * -1 & "]C:R[-1]C)"

mRange = "A6:I" & ctr3
xlApp1.Range(mRange).Select
With xlApp1
    '.Selection.Columns.AutoFit
    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With xlApp1.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End With
With xlApp1.Range(mRange)
    .Font.Name = "Arial"   'Font.Name
    .Font.Size = 8         'Font.size
End With

mRange = "D7:I" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "C" & ctr3 & ":I" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

mRange = "A6:I" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Selection.Columns.AutoFit
xlApp1.Range("A1").Select

'                   Date Transmitted:   January 2009
'   Company         Applicable Month:   October 2008
'   SSS# 03-9210478-2

'   SSS Payments    EC  Total       SBR NUMBER/OR#      Date Paid

xlApp1.Sheets("Sheet1").Select

'Box it.
mRange = "A6:K" & ctr
Call PlaceBox(xlSheet, xlApp1, mRange)

xlApp1.Sheets("Sheet1").Name = "SSS Details"
xlApp1.Sheets("Sheet2").Name = "SSS Summary"
xlApp1.Sheets("Sheet3").Name = "Not to be Remitted SSS"

xlSheet.Range("A1").Select

ExitRtn2:

End Sub

Private Function UnremittedSSSContributions(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal ctr As Integer) As Integer
Dim mRange As String, mRange1 As String
Dim LineCtr As Integer
Dim cPos As Integer
Dim i As Integer

i = 1
LineCtr = 1
ctr = ctr + 2
cPos = ctr + 1

PrevCompany = ListView1(0).ListItems(i).SubItems(13)
mRange = "A" & ctr & ":D" & ctr
Call PrintReportHeader(xlSheet, xlApp1, mRange, "NOT TO BE REMITTED CLIENTS", "", xlLeft, , False, False, False, , , 10, True)
'Call PrintReportHeader(xlSheet, xlApp1, mRange, PrevCompany, "", xlLeft, , False, False, False, , , 9, True, 18)

Do While ListView1(0).ListItems.Count >= i
    If ListView1(0).ListItems.Item(i).Checked = False Then
        GoTo Next3
    Else
        If PrevCompany <> ListView1(0).ListItems(i).SubItems(13) Then
            If LineCtr > 1 Then
                ' Print Sub-Totals.
                ctr = ctr + 1
                mRange = "F" & cPos & ":K" & ctr
                xlSheet.Range(mRange).Select
                mRange1 = "F" & ctr
                xlSheet.Range(mRange1).Activate
                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
                
                ' Copy to Sheet 3 the Total.
                ctr3 = ctr3 + 1
                xlSheet.Range(mRange1).Select
                xlApp1.Selection.Copy
                xlApp1.Sheets("Sheet3").Select
                mRange1 = "D" & ctr3
                xlApp1.Range(mRange1).Select
                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=False
                mRange1 = "A" & ctr3
                xlApp1.Range(mRange1).Select
                xlApp1.ActiveCell.FormulaR1C1 = LineCtr3
                mRange1 = "B" & ctr3
                xlApp1.Range(mRange1).Select
                xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
                mRange1 = "C" & ctr3
                xlApp1.Range(mRange1).Select
                xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
                xlApp1.Sheets("Sheet1").Select
                
                mRange = "F" & cPos + 1 & ":K" & ctr
                xlSheet.Range(mRange).Select
                mRange1 = "G" & ctr
                xlSheet.Range(mRange1).Activate
                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
                
                ' Copy to Sheet 3 the Total.
                xlSheet.Range(mRange1).Select
                xlApp1.Selection.Copy
                xlApp1.Sheets("Sheet3").Select
                mRange1 = "E" & ctr3
                xlApp1.Range(mRange1).Select
                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=False
                xlApp1.Sheets("Sheet1").Select
                
                mRange = "F" & cPos + 1 & ":K" & ctr
                xlSheet.Range(mRange).Select
                mRange1 = "H" & ctr
                xlSheet.Range(mRange1).Activate
                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
                
                ' Copy to Sheet 3 the Total.
                xlSheet.Range(mRange1).Select
                xlApp1.Selection.Copy
                xlApp1.Sheets("Sheet3").Select
                mRange1 = "F" & ctr3
                xlApp1.Range(mRange1).Select
                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=False
                xlApp1.Sheets("Sheet1").Select
                
                mRange = "F" & cPos + 1 & ":K" & ctr
                xlSheet.Range(mRange).Select
                mRange1 = "I" & ctr
                xlSheet.Range(mRange1).Activate
                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
                
                ' Copy to Sheet 3 the Total.
                xlSheet.Range(mRange1).Select
                xlApp1.Selection.Copy
                xlApp1.Sheets("Sheet3").Select
                mRange1 = "G" & ctr3
                xlApp1.Range(mRange1).Select
                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=False
                xlApp1.Sheets("Sheet1").Select
                
                mRange = "F" & cPos + 1 & ":K" & ctr
                xlSheet.Range(mRange).Select
                mRange1 = "J" & ctr
                xlSheet.Range(mRange1).Activate
                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
                
                ' Copy to Sheet 3 the Total.
                xlSheet.Range(mRange1).Select
                xlApp1.Selection.Copy
                xlApp1.Sheets("Sheet3").Select
                mRange1 = "H" & ctr3
                xlApp1.Range(mRange1).Select
                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=False
                xlApp1.Sheets("Sheet1").Select
                
                mRange = "F" & cPos + 1 & ":K" & ctr
                xlSheet.Range(mRange).Select
                mRange1 = "K" & ctr
                xlSheet.Range(mRange1).Activate
                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
                
                ' Copy to Sheet 3 the Total.
                xlSheet.Range(mRange1).Select
                xlApp1.Selection.Copy
                xlApp1.Sheets("Sheet3").Select
                mRange1 = "I" & ctr3
                xlApp1.Range(mRange1).Select
                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=False
                xlApp1.Sheets("Sheet1").Select
                LineCtr3 = LineCtr3 + 1
                
                mRange = "F" & cPos & ":K" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                mRange = "F" & ctr & ":K" & ctr
                xlSheet.Range(mRange).Select
                xlApp1.Range(mRange).Font.Bold = True
                
            If SumTotal = "" Then
                SumTotal = "=F" & ctr
            Else
                SumTotal = SumTotal & "+F" & ctr
            End If
                
                
            End If
            PrevCompany = ListView1(0).ListItems(i).SubItems(13)
            HeaderFlag = False
            NoSSSCtr = 0  ' set again to 0 after printing
        End If
        ' Print Header.
        If HeaderFlag = False Then
            'PrevCompany = ListView1(0).ListItems(i).SubItems(13)
            LineCtr = 1
            ctr = ctr + 3
            mRange = "A" & ctr
            Call PrintReportHeader(xlSheet, xlApp1, mRange, PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
            'Call PrintReportHeader(xlSheet, xlApp1, "A6", PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
            ctr = ctr + 1
            mRange = "A" & ctr & ":A" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "A7:A8", "No.", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 4.14)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "No.", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 4.14)
            mRange = "B" & ctr & ":B" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "B7:B8", "Surname", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Surname", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)
            mRange = "C" & ctr & ":C" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "C7:C8", "Firstname", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Firstname", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
            mRange = "D" & ctr & ":D" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "D7:D8", "Middlename", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Middlename", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
            mRange = "E" & ctr & ":E" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "E7:E8", "SSS Number", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 11)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "SSS Number", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 11)
            mRange = "F" & ctr & ":F" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "F7:F8", "Employee Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Employee Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "G" & ctr & ":G" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "G7:G8", "Employer Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Employer Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "H" & ctr & ":H" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "H7:H8", "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
            mRange = "I" & ctr & ":I" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "I7:I8", "Adjusted", "PS Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "PS Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "J" & ctr & ":J" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "J7:J8", "Adjusted", "ES Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "ES Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "K" & ctr & ":K" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "K7:K8", "Difference", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            HeaderFlag = True
            cPos = ctr + 2
            ctr = ctr + 1
        End If
        
'        If ListView1(0).ListItems.Item(i).Checked = True Then
'            ' Count those excluded records.
'            ExcludedCtr = ExcludedCtr + 1
'            GoTo Next1
'        ElseIf ListView1(0).ListItems(i).SubItems(4) = "__-_______-_" Or ListView1(0).ListItems(i).SubItems(4) = "" Then
'            ' Count those no sss number.
'            NoSSSCtr = NoSSSCtr + 1
'            GoTo Next1
'        Else
            ctr = ctr + 1
            xlSheet.Cells(ctr, 1).VALUE = LineCtr
            xlSheet.Cells(ctr, 2).VALUE = ListView1(0).ListItems(i).SubItems(1) ' Surname
            xlSheet.Cells(ctr, 3).VALUE = ListView1(0).ListItems(i).SubItems(2) ' Firstname
            xlSheet.Cells(ctr, 4).VALUE = ListView1(0).ListItems(i).SubItems(3) ' Middlename
            xlSheet.Cells(ctr, 5).VALUE = ListView1(0).ListItems(i).SubItems(4) ' SSS no
            xlSheet.Cells(ctr, 6).VALUE = ListView1(0).ListItems(i).SubItems(7) ' Employee
            xlSheet.Cells(ctr, 7).VALUE = ListView1(0).ListItems(i).SubItems(8) ' Employer
            xlSheet.Cells(ctr, 8).VALUE = ListView1(0).ListItems(i).SubItems(9) ' Ecc
            xlSheet.Cells(ctr, 9).VALUE = ListView1(0).ListItems(i).SubItems(10) ' employee adjust
            xlSheet.Cells(ctr, 10).VALUE = ListView1(0).ListItems(i).SubItems(11) ' employer adjust
            xlSheet.Cells(ctr, 11).VALUE = ListView1(0).ListItems(i).SubItems(12) ' difference
            LineCtr = LineCtr + 1
'        End If
        End If
Next3:
    
    i = i + 1

Loop

' Print last Sub-Totals.
ctr = ctr + 1
mRange = "F" & cPos + 1 & ":K" & ctr
xlSheet.Range(mRange).Select
mRange1 = "F" & ctr
xlSheet.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

' Copy to Sheet 3 the Total.
ctr3 = ctr3 + 1
xlSheet.Range(mRange1).Select
xlApp1.Selection.Copy
xlApp1.Sheets("Sheet3").Select
mRange1 = "D" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False
mRange1 = "A" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = LineCtr3
mRange1 = "B" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
mRange1 = "C" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
xlApp1.Sheets("Sheet1").Select
''''''-----------

mRange = "F" & cPos + 1 & ":K" & ctr
xlSheet.Range(mRange).Select
mRange1 = "G" & ctr
xlSheet.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

' Copy to Sheet 3 the Total.
xlSheet.Range(mRange1).Select
xlApp1.Selection.Copy
xlApp1.Sheets("Sheet3").Select
mRange1 = "E" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False
xlApp1.Sheets("Sheet1").Select

mRange = "F" & cPos + 1 & ":K" & ctr
xlSheet.Range(mRange).Select
mRange1 = "H" & ctr
xlSheet.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

' Copy to Sheet 3 the Total.
xlSheet.Range(mRange1).Select
xlApp1.Selection.Copy
xlApp1.Sheets("Sheet3").Select
mRange1 = "F" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False
xlApp1.Sheets("Sheet1").Select

mRange = "F" & cPos + 1 & ":K" & ctr
xlSheet.Range(mRange).Select
mRange1 = "I" & ctr
xlSheet.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

' Copy to Sheet 3 the Total.
xlSheet.Range(mRange1).Select
xlApp1.Selection.Copy
xlApp1.Sheets("Sheet3").Select
mRange1 = "G" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False
xlApp1.Sheets("Sheet1").Select

mRange = "F" & cPos + 1 & ":K" & ctr
xlSheet.Range(mRange).Select
mRange1 = "J" & ctr
xlSheet.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

' Copy to Sheet 3 the Total.
xlSheet.Range(mRange1).Select
xlApp1.Selection.Copy
xlApp1.Sheets("Sheet3").Select
mRange1 = "H" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False
xlApp1.Sheets("Sheet1").Select

mRange = "F" & cPos + 1 & ":K" & ctr
xlSheet.Range(mRange).Select
mRange1 = "K" & ctr
xlSheet.Range(mRange1).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"

' Copy to Sheet 3 the Total.
xlSheet.Range(mRange1).Select
xlApp1.Selection.Copy
xlApp1.Sheets("Sheet3").Select
mRange1 = "I" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False
xlApp1.Sheets("Sheet1").Select

mRange = "F" & cPos & ":K" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":K" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

            If SumTotal = "" Then
                SumTotal = "=F" & ctr
            Else
                SumTotal = SumTotal & "+F" & ctr
            End If

' Print Total summary.
ctr = ctr + 2
mRange = "B" & ctr & ":D" & ctr
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Totals", "", xlLeft, , False, False, False, , , 10, True)
GTotal = 0

For l = 70 To 82  ' F to R
    mRange = Chr(l) & ctr
    xlSheet.Range(mRange).Activate
    xlSheet.Cells(ctr, l - 64).VALUE = SumTotal
    If l = 70 Then ' 70
        xlSheet.Cells(ctr + 2, 11).VALUE = "Actual EE Share"
        xlSheet.Cells(ctr + 2, 12).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 2, 12).VALUE)
    ElseIf l = 71 Then ' 77
        xlSheet.Cells(ctr + 3, 11).VALUE = "Actual ER Share"
        xlSheet.Cells(ctr + 3, 12).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 3, 12).VALUE)
        xlSheet.Cells(ctr + 4, 11).VALUE = "Sub-Total"
        xlSheet.Cells(ctr + 4, 12).VALUE = GTotal
        'GTotal = GTotal + Val(xlSheet.Cells(ctr + 3, 12).VALUE)
    ElseIf l = 74 Then ' 78
        xlSheet.Cells(ctr + 5, 11).VALUE = "ECC"
        xlSheet.Cells(ctr + 5, 12).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 5, 12).VALUE)
    End If
    SumTotal = Replace(SumTotal, Chr(l), Chr(l + 1))
Next l
mRange = "F" & cPos & ":R" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":R" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

ctr = ctr + 5 + 1
xlSheet.Cells(ctr, 11).VALUE = "Grand Total"
xlSheet.Cells(ctr, 12).VALUE = GTotal

mRange = "F" & cPos & ":R" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":R" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

UnremittedSSSContributions = ctr
End Function

Private Function UnremittedSSSWithSBR(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal ctr As Integer) As Integer
Dim mRange As String, mRange1 As String
Dim LineCtr As Integer
Dim cPos As Integer
Dim i As Integer, l As Integer

i = 1
l = 1
LineCtr = 1
ctr = ctr + 2
cPos = ctr + 1

PrevCompany = ListView1(0).ListItems(i).SubItems(13)
mRange = "A" & ctr & ":D" & ctr
Call PrintReportHeader(xlSheet, xlApp1, mRange, "NOT TO BE REMITTED CLIENTS", "", xlLeft, , False, False, False, , , 10, True)

Do While ListView1(0).ListItems.Count >= i
    If ListView1(0).ListItems.Item(i).Checked = False Then
        GoTo Next3
    Else
        If PrevCompany <> ListView1(0).ListItems(i).SubItems(13) Then
            If LineCtr > 1 Then
                ' Print Sub-Totals.
                ctr = ctr + 1
                xlApp1.Sheets("Sheet3").Select
                ctr3 = ctr3 + 1
                mRange1 = "A" & ctr3
                xlApp1.Range(mRange1).Select
                xlApp1.ActiveCell.FormulaR1C1 = LineCtr3
                mRange1 = "B" & ctr3
                xlApp1.Range(mRange1).Select
                xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
                mRange1 = "C" & ctr3
                xlApp1.Range(mRange1).Select
                xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
                
                For l = 70 To 82    ' from F to R
                    xlApp1.Sheets("Sheet1").Select
                    ' Compute for Total.
                    mRange1 = Chr(l) & ctr
                    xlSheet.Range(mRange1).Activate
                    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
                    ' Save the Total.
                    xlSheet.Range(mRange1).Select
                    xlApp1.Selection.Copy
                    ' Paste total to Sheet 2.
                    xlApp1.Sheets("Sheet3").Select
                    mRange1 = Chr(l - 2) & ctr3
                    xlApp1.Range(mRange1).Select
                    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                        False, Transpose:=False
                Next l
                
                xlApp1.Sheets("Sheet1").Select
                LineCtr3 = LineCtr3 + 1
                mRange = "F" & cPos & ":R" & ctr
                xlSheet.Range(mRange).Select
                xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
                mRange = "F" & ctr & ":R" & ctr
                xlSheet.Range(mRange).Select
                xlApp1.Range(mRange).Font.Bold = True
            
            If SumTotal = "" Then
                SumTotal = "=F" & ctr
            Else
                SumTotal = SumTotal & "+F" & ctr
            End If
            
            
            End If
            PrevCompany = ListView1(0).ListItems(i).SubItems(13)
            HeaderFlag = False
            NoSSSCtr = 0  ' set again to 0 after printing
        End If
        ' Print Header.
        If HeaderFlag = False Then
            'PrevCompany = ListView1(0).ListItems(i).SubItems(13)
            LineCtr = 1
            ctr = ctr + 3
            mRange = "A" & ctr
            Call PrintReportHeader(xlSheet, xlApp1, mRange, PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
            'Call PrintReportHeader(xlSheet, xlApp1, "A6", PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
            ctr = ctr + 1
            mRange = "A" & ctr & ":A" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "A7:A8", "No.", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 4.14)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "No.", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 4.14)
            mRange = "B" & ctr & ":B" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "B7:B8", "Surname", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Surname", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)
            mRange = "C" & ctr & ":C" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "C7:C8", "Firstname", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Firstname", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
            mRange = "D" & ctr & ":D" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "D7:D8", "Middlename", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Middlename", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
            mRange = "E" & ctr & ":E" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "E7:E8", "SSS Number", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 11)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "SSS Number", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 11)
            mRange = "F" & ctr & ":F" & ctr + 1
            'Call PrintReportHeader(xlSheet, xlApp1, "F7:F8", "Employee Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "G" & ctr & ":G" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "H" & ctr & ":H" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Table Based", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "I" & ctr & ":I" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Table Based", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "J" & ctr & ":J" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
            mRange = "K" & ctr & ":K" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
            mRange = "L" & ctr & ":L" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "M" & ctr & ":M" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "N" & ctr & ":N" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            
            mRange = "O" & ctr & ":O" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total On", "Adjusted", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "P" & ctr & ":P" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "EE", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "Q" & ctr & ":Q" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "ER", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "R" & ctr & ":R" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "Difference", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "S" & ctr & ":S" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "SBR#", "OR#", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "T" & ctr & ":T" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Date Paid", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            mRange = "U" & ctr & ":U" & ctr + 1
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Date Posted", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
            HeaderFlag = True
            cPos = ctr + 2
            ctr = ctr + 1
        End If
        
'        If ListView1(0).ListItems.Item(i).Checked = True Then
'            ' Count those excluded records.
'            ExcludedCtr = ExcludedCtr + 1
'            GoTo Next1
'        ElseIf ListView1(0).ListItems(i).SubItems(4) = "__-_______-_" Or ListView1(0).ListItems(i).SubItems(4) = "" Then
'            ' Count those no sss number.
'            NoSSSCtr = NoSSSCtr + 1
'            GoTo Next1
'        Else
            ctr = ctr + 1
            xlSheet.Cells(ctr, 1).VALUE = LineCtr
            xlSheet.Cells(ctr, 2).VALUE = ListView1(0).ListItems(i).SubItems(1) ' Surname
            xlSheet.Cells(ctr, 3).VALUE = ListView1(0).ListItems(i).SubItems(2) ' Firstname
            xlSheet.Cells(ctr, 4).VALUE = ListView1(0).ListItems(i).SubItems(3) ' Middlename
            xlSheet.Cells(ctr, 5).VALUE = ListView1(0).ListItems(i).SubItems(4) ' SSS no
            xlSheet.Cells(ctr, 6).VALUE = ListView1(0).ListItems(i).SubItems(7) ' Employee
            xlSheet.Cells(ctr, 7).VALUE = ListView1(0).ListItems(i).SubItems(8) ' Employer
'            xlSheet.Cells(ctr, 8).VALUE = ListView1(0).ListItems(i).SubItems(9) ' Ecc
'            xlSheet.Cells(ctr, 9).VALUE = ListView1(0).ListItems(i).SubItems(10) ' employee adjust
'            xlSheet.Cells(ctr, 10).VALUE = ListView1(0).ListItems(i).SubItems(11) ' employer adjust
'            xlSheet.Cells(ctr, 11).VALUE = ListView1(0).ListItems(i).SubItems(12) ' difference
            
            xlSheet.Cells(ctr, 8).VALUE = ListView1(0).ListItems(i).SubItems(14) ' table based ee
            xlSheet.Cells(ctr, 9).VALUE = ListView1(0).ListItems(i).SubItems(15) ' table based er
            xlSheet.Cells(ctr, 10).VALUE = ListView1(0).ListItems(i).SubItems(9) ' ecc
            ' Total table based
            xlSheet.Cells(ctr, 11).VALUE = Val(ListView1(0).ListItems(i).SubItems(14)) + Val(ListView1(0).ListItems(i).SubItems(15)) + Val(ListView1(0).ListItems(i).SubItems(9))   ' table based total
            
            xlSheet.Cells(ctr, 12).VALUE = ListView1(0).ListItems(i).SubItems(10) ' employee adjust
            xlSheet.Cells(ctr, 13).VALUE = ListView1(0).ListItems(i).SubItems(11) ' employer adjust
            xlSheet.Cells(ctr, 14).VALUE = ListView1(0).ListItems(i).SubItems(9) ' ecc
            ' Total Adjusted
            xlSheet.Cells(ctr, 15).VALUE = Val(ListView1(0).ListItems(i).SubItems(10)) + Val(ListView1(0).ListItems(i).SubItems(11)) + Val(ListView1(0).ListItems(i).SubItems(9))   ' table based total
            ' Difference EE
            xlSheet.Cells(ctr, 16).VALUE = Val(ListView1(0).ListItems(i).SubItems(7)) - Val(ListView1(0).ListItems(i).SubItems(10))
            ' Difference ER
            xlSheet.Cells(ctr, 17).VALUE = Val(ListView1(0).ListItems(i).SubItems(8)) - Val(ListView1(0).ListItems(i).SubItems(11))
            ' Total Difference
            xlSheet.Cells(ctr, 18).VALUE = (Val(ListView1(0).ListItems(i).SubItems(7)) - Val(ListView1(0).ListItems(i).SubItems(10))) + (Val(ListView1(0).ListItems(i).SubItems(8)) - Val(ListView1(0).ListItems(i).SubItems(11)))
            
            LineCtr = LineCtr + 1
'        End If
        End If
Next3:
    
    i = i + 1

Loop

' Print Sub-Totals.
ctr = ctr + 1
xlApp1.Sheets("Sheet3").Select
ctr3 = ctr3 + 1
mRange1 = "A" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = LineCtr3
mRange1 = "B" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
mRange1 = "C" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE

For l = 70 To 82    ' from F to R
    xlApp1.Sheets("Sheet1").Select
    ' Compute for Total.
    mRange1 = Chr(l) & ctr
    xlSheet.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
    ' Save the Total.
    xlSheet.Range(mRange1).Select
    xlApp1.Selection.Copy
    ' Paste total to Sheet 2.
    xlApp1.Sheets("Sheet3").Select
    mRange1 = Chr(l - 2) & ctr3
    xlApp1.Range(mRange1).Select
    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Next l

xlApp1.Sheets("Sheet1").Select
mRange = "F" & cPos & ":R" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":R" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

            If SumTotal = "" Then
                SumTotal = "=F" & ctr
            Else
                SumTotal = SumTotal & "+F" & ctr
            End If

' Print Total summary.
ctr = ctr + 2
mRange = "B" & ctr & ":D" & ctr
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Totals", "", xlLeft, , False, False, False, , , 10, True)
GTotal = 0

For l = 70 To 82  ' F to R
    mRange = Chr(l) & ctr
    xlSheet.Range(mRange).Activate
    xlSheet.Cells(ctr, l - 64).VALUE = SumTotal
    If l = 70 Then ' 70
        xlSheet.Cells(ctr + 2, 11).VALUE = "Actual EE Share"
        xlSheet.Cells(ctr + 2, 12).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 2, 12).VALUE)
    ElseIf l = 71 Then ' 77
        xlSheet.Cells(ctr + 3, 11).VALUE = "Actual ER Share"
        xlSheet.Cells(ctr + 3, 12).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 3, 12).VALUE)
        xlSheet.Cells(ctr + 4, 11).VALUE = "Sub-Total"
        xlSheet.Cells(ctr + 4, 12).VALUE = GTotal
        'GTotal = GTotal + Val(xlSheet.Cells(ctr + 3, 12).VALUE)
    ElseIf l = 74 Then ' 78
        xlSheet.Cells(ctr + 5, 11).VALUE = "ECC"
        xlSheet.Cells(ctr + 5, 12).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 5, 12).VALUE)
    End If
    SumTotal = Replace(SumTotal, Chr(l), Chr(l + 1))
Next l
mRange = "F" & cPos & ":R" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":R" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

ctr = ctr + 5 + 1
xlSheet.Cells(ctr, 11).VALUE = "Grand Total"
xlSheet.Cells(ctr, 12).VALUE = GTotal

mRange = "F" & cPos & ":R" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":R" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

UnremittedSSSWithSBR = ctr

End Function

Private Function PrintNoSSSWithSBR(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal mPCompany As String, ByVal ctr As Integer) As Integer

Dim j As Integer, l As Integer
Dim mRange As String, mRange1 As String
Dim LineCtr As Integer
Dim cPos As Integer
Dim mCo2 As String

LineCtr = 1
j = 1
l = 1
ctr = ctr + 2
cPos = ctr + 1

mRange = "B" & ctr & ":D" & ctr
mCo2 = mPCompany
mPCompany = mPCompany & " (Employees with no SSS#)"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mPCompany, "", xlLeft, , False, False, False, , , 9, True, 18)
'Call PrintReportHeader(xlSheet, xlApp1, mRange, mPCompany, "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)

Do While ListView1(0).ListItems.Count >= j
    If ListView1(0).ListItems(j).SubItems(13) = mCo2 And ListView1(0).ListItems.Item(j).Checked = False And _
(ListView1(0).ListItems(j).SubItems(4) = "__-_______-_" Or ListView1(0).ListItems(j).SubItems(4) = "") Then
        ctr = ctr + 1
        xlSheet.Cells(ctr, 1).VALUE = LineCtr
        xlSheet.Cells(ctr, 2).VALUE = ListView1(0).ListItems(j).SubItems(1) ' Surname
        xlSheet.Cells(ctr, 3).VALUE = ListView1(0).ListItems(j).SubItems(2) ' firstname
        xlSheet.Cells(ctr, 4).VALUE = ListView1(0).ListItems(j).SubItems(3) ' Middlename
        xlSheet.Cells(ctr, 5).VALUE = ListView1(0).ListItems(j).SubItems(4) ' sss no
        xlSheet.Cells(ctr, 6).VALUE = ListView1(0).ListItems(j).SubItems(7) ' employee
        xlSheet.Cells(ctr, 7).VALUE = ListView1(0).ListItems(j).SubItems(8) ' employer
        
        xlSheet.Cells(ctr, 8).VALUE = ListView1(0).ListItems(j).SubItems(14) ' table based ee
        xlSheet.Cells(ctr, 9).VALUE = ListView1(0).ListItems(j).SubItems(15) ' table based er
        xlSheet.Cells(ctr, 10).VALUE = ListView1(0).ListItems(j).SubItems(9) ' ecc
        ' Total table based
        xlSheet.Cells(ctr, 11).VALUE = Val(ListView1(0).ListItems(j).SubItems(14)) + Val(ListView1(0).ListItems(j).SubItems(15)) + Val(ListView1(0).ListItems(j).SubItems(9))   ' table based total
        
        xlSheet.Cells(ctr, 12).VALUE = ListView1(0).ListItems(j).SubItems(10) ' employee adjust
        xlSheet.Cells(ctr, 13).VALUE = ListView1(0).ListItems(j).SubItems(11) ' employer adjust
        xlSheet.Cells(ctr, 14).VALUE = ListView1(0).ListItems(j).SubItems(9) ' ecc
        ' Total Adjusted
        xlSheet.Cells(ctr, 15).VALUE = Val(ListView1(0).ListItems(j).SubItems(10)) + Val(ListView1(0).ListItems(j).SubItems(11)) + Val(ListView1(0).ListItems(j).SubItems(9))   ' table based total
        ' Difference EE
        xlSheet.Cells(ctr, 16).VALUE = Val(ListView1(0).ListItems(j).SubItems(7)) - Val(ListView1(0).ListItems(j).SubItems(10))
        ' Difference ER
        xlSheet.Cells(ctr, 17).VALUE = Val(ListView1(0).ListItems(j).SubItems(8)) - Val(ListView1(0).ListItems(j).SubItems(11))
        ' Total Difference
        xlSheet.Cells(ctr, 18).VALUE = (Val(ListView1(0).ListItems(j).SubItems(7)) - Val(ListView1(0).ListItems(j).SubItems(10))) + (Val(ListView1(0).ListItems(j).SubItems(8)) - Val(ListView1(0).ListItems(j).SubItems(11)))
        LineCtr = LineCtr + 1
    End If
    j = j + 1
Loop
' Print Sub-Totals.
ctr = ctr + 1
ctr2 = ctr2 + 1
xlApp1.Sheets("Sheet2").Select
mRange1 = "A" & ctr2
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
mRange1 = "B" & ctr2
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = mPCompany
mRange1 = "C" & ctr2
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE

For l = 70 To 82    ' from F to R
    xlApp1.Sheets("Sheet1").Select
    ' Compute for Total.
    mRange1 = Chr(l) & ctr
    xlSheet.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
    ' Save the Total.
    xlSheet.Range(mRange1).Select
    xlApp1.Selection.Copy
    ' Paste total to Sheet 2.
    xlApp1.Sheets("Sheet2").Select
    mRange1 = Chr(l - 2) & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Next l

xlApp1.Sheets("Sheet1").Select
mRange = "F" & cPos & ":R" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":R" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

If SumTotal = "" Then
    SumTotal = "=F" & ctr
Else
    SumTotal = SumTotal & "+F" & ctr
End If

PrintNoSSSWithSBR = ctr

End Function

Private Sub PrintSSSTransmittal(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application)
    
Dim LineCtr As Integer
Dim mHeading As String
Dim mRange As String

LineCtr = 7

xlApp1.Sheets("Sheet4").Select

mHeading = "DATE TRANSMITTED: " & MaskDatePosted.Text
mRange = "G7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeading, "", xlLeft, , False, False, False, , , 9, False)

mHeading = CompanyName
mRange = "A9"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeading, "", xlLeft, , False, False, False, , , 9, False)

mHeading = "APPLICABLE MONTH: " & txtAppMonthYear(1).Text
mRange = "G9"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeading, "", xlLeft, , False, False, False, , , 9, False)

mHeading = txtPhilHealthNo.Text ' get SSS #
mRange = "A11"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeading, "", xlLeft, , False, False, False, , , 9, False)

mHeading = "SSS"
mRange = "A13"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeading, "", xlCenter, , False, False, False, , , 9, False)

mHeading = "EC"
mRange = "C13"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeading, "", xlCenter, , False, False, False, , , 9, False)

mHeading = "TOTAL"
mRange = "E13"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeading, "", xlCenter, , False, False, False, , , 9, False)

mHeading = "SBR NUMBER/OR#"
mRange = "G13"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeading, "", xlLeft, , False, False, False, , , 9, False)

mHeading = "DATE PAID"
mRange = "I13"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeading, "", xlLeft, , False, False, False, , , 9, False)

xlApp1.Range("A15").Select
xlApp1.ActiveCell.FormulaR1C1 = "=RC[4]-RC[2]"

mRange = "A15:E15"
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"

mHeading = txtSBRNo.Text
mRange = "G15"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeading, "", xlLeft, , False, False, False, , , 9, False)

mHeading = MaskDatePaid.Text
mRange = "I15"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeading, "", xlLeft, , False, False, False, , , 9, False)

mHeading = "Prepared by:  JASMIN"
mRange = "B20"
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeading, "", xlLeft, , False, False, False, , , 9, False)

End Sub

Private Sub GenerateSBRReports()    ' SSS Reports with SBR #
Dim rsDedRecords As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mRange As String, mRange1 As String
Dim ctr As Long
Dim i As Integer, l As Integer
Dim cPos As Integer
Dim mHead As String
Dim HeaderFlag As Boolean
Dim NoSSSCtr As Integer
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

GTotal = 0#
xlSheetNo = 1
WithRecSw = False
SumTotal = ""
NoSSSCtr = 0: ExcludedCtr = 0
HeaderFlag = False
LineCtr = 1
LineCtr2 = 1
LineCtr3 = 1
ctr = 3
ctr2 = 7    '6
ctr3 = 7
cPos = 9
i = 1
l = 1

'Print Header.
xlApp1.Sheets("Sheet1").Select
mHead = CompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 12, True)
mHead = "SSS Contributions Report for the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)

'Print Header of Sheet2 for Remitted Clients Summary.
mHead = CompanyName & " (" & CompanyShort & ")"
xlApp1.Sheets("Sheet2").Select
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 12, True)

'xlApp1.Range("A2").Select
'xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "SSS Contributions Summary Report per Clients with SBR#"
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)

'xlApp1.Range("A3").Select
'xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "For the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)
'xlApp1.Range("A4").Select
'xlApp1.ActiveCell.FormulaR1C1 = mHead

'xlApp1.Range("A6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "No."
mRange = "A6:A7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "No.", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)
'xlApp1.Range("B6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Client Name"
mRange = "B6:B7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Client Name", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)

'xlApp1.Range("C6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Total Employees"
mRange = "C6:C7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "Employees", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)

'xlApp1.Range("D6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Actual Deductions"
mRange = "D6:D7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "E6:E7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)

'xlApp1.Range("E6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Employee Share"
mRange = "F6:F7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Table Based", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 9, False, 8.7)

'xlApp1.Range("E6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "ER Share"
mRange = "G6:G7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Table Based", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 9, False, 8.7)

'xlApp1.Range("F6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "ECC"
mRange = "H6:H7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 5)
mRange = "I6:I7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "J6:J7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "EE", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "K6:K7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "ER", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "L6:L7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 5)
mRange = "M6:M7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "Adjusted", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "N6:N7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "EE", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "O6:O7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "ER", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "P6:P7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "Difference", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)

mRange = "Q6:Q7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "SBR#", "OR#", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "R6:R7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Date", "Paid", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "S6:S7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Date", "Posted", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)


'Print Header of Sheet3 for Unremitted Clients Summary.
xlApp1.Sheets("Sheet3").Select
mHead = CompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 12, True)
mHead = "SSS Contributions Summary Report (NOT TO BE REMITTED CLIENTS)"
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "For the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)

'xlApp1.Range("A6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "No."
mRange = "A6:A7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "No.", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)
'xlApp1.Range("B6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Client Name"
mRange = "B6:B7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Client Name", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)

'xlApp1.Range("C6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Total Employees"
mRange = "C6:C7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "Employees", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)

'xlApp1.Range("D6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Actual Deductions"
mRange = "D6:D7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "E6:E7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)

'xlApp1.Range("E6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Employee Share"
mRange = "F6:F7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Table Based", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 9, False, 8.7)

'xlApp1.Range("E6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "ER Share"
mRange = "G6:G7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Table Based", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 9, False, 8.7)

'xlApp1.Range("F6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "ECC"
mRange = "H6:H7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 5)
mRange = "I6:I7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "J6:J7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "EE", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "K6:K7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "ER", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "L6:L7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 5)
mRange = "M6:M7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "Adjusted", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "N6:N7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "EE", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "O6:O7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "ER", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "P6:P7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "Difference", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)

mRange = "Q6:Q7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "SBR#", "OR#", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "R6:R7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Date", "Paid", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)
mRange = "S6:S7"
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Date", "Posted", xlCenter, xlCenter, True, False, True, 15, , 9, False, 7.5)

' Insert new Sheet for SSS Transmittal.
xlApp1.Sheets.Add
xlApp1.Sheets("Sheet4").Move After:=xlApp1.Sheets(4)
xlApp1.Sheets("Sheet4").Select
mHead = CompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 12, True)
mHead = "SSS Transmittal Report"
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "For the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)

xlApp1.Sheets("Sheet1").Select

PrevCompany = ListView1(0).ListItems(i).SubItems(13)
Do While ListView1(0).ListItems.Count >= i

    If PrevCompany <> ListView1(0).ListItems(i).SubItems(13) Then
        If LineCtr > 1 Then
            ' Print Sub-Totals.
            ctr = ctr + 1
            xlApp1.Sheets("Sheet2").Select
            ctr2 = ctr2 + 1
            mRange1 = "A" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
            mRange1 = "B" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
            mRange1 = "C" & ctr2
            xlApp1.Range(mRange1).Select
            xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
            
            For l = 70 To 82    ' from F to R
                xlApp1.Sheets("Sheet1").Select
                ' Compute for Total.
                mRange1 = Chr(l) & ctr
                xlSheet.Range(mRange1).Activate
                xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
                ' Save the Total.
                xlSheet.Range(mRange1).Select
                xlApp1.Selection.Copy
                ' Paste total to Sheet 2.
                xlApp1.Sheets("Sheet2").Select
                mRange1 = Chr(l - 2) & ctr2
                xlApp1.Range(mRange1).Select
                xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=False
            Next l

            xlApp1.Sheets("Sheet1").Select
            LineCtr2 = LineCtr2 + 1
                
            mRange = "F" & cPos & ":R" & ctr
            xlSheet.Range(mRange).Select
            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
            mRange = "F" & ctr & ":R" & ctr
            xlSheet.Range(mRange).Select
            xlApp1.Range(mRange).Font.Bold = True
            
            If SumTotal = "" Then
                SumTotal = "=F" & ctr
            Else
                SumTotal = SumTotal & "+F" & ctr
            End If
        End If
        
        If NoSSSCtr > 0 Then
            ctr = PrintNoSSSWithSBR(xlSheet, xlApp1, PrevCompany, ctr)
            LineCtr2 = LineCtr2 + 1
        End If

        'ctr = ctr + 3
        PrevCompany = ListView1(0).ListItems(i).SubItems(13)
'        mRange = "A" & ctr
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
        HeaderFlag = False
        NoSSSCtr = 0  ' set again to 0 after printing
    End If
    ' Print Header.
    If HeaderFlag = False Then
        'PrevCompany = ListView1(0).ListItems(i).SubItems(13)
        LineCtr = 1
        ctr = ctr + 3
        mRange = "A" & ctr
        Call PrintReportHeader(xlSheet, xlApp1, mRange, PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
        'Call PrintReportHeader(xlSheet, xlApp1, "A6", PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
        ctr = ctr + 1
        mRange = "A" & ctr & ":A" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "A7:A8", "No.", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 4.14)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "No.", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 4.14)
        mRange = "B" & ctr & ":B" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "B7:B8", "Surname", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Surname", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 18)
        mRange = "C" & ctr & ":C" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "C7:C8", "Firstname", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Firstname", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
        mRange = "D" & ctr & ":D" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "D7:D8", "Middlename", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Middlename", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 18)
        mRange = "E" & ctr & ":E" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "E7:E8", "SSS Number", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 11)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "SSS Number", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 11)
        mRange = "F" & ctr & ":F" & ctr + 1
        'Call PrintReportHeader(xlSheet, xlApp1, "F7:F8", "Employee Share", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "G" & ctr & ":G" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "H" & ctr & ":H" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Table Based", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "I" & ctr & ":I" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Table Based", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "J" & ctr & ":J" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
        mRange = "K" & ctr & ":K" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
        mRange = "L" & ctr & ":L" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "M" & ctr & ":M" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Adjusted", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "N" & ctr & ":N" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "ECC", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        
        mRange = "O" & ctr & ":O" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total On", "Adjusted", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "P" & ctr & ":P" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "EE", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "Q" & ctr & ":Q" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "ER", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "R" & ctr & ":R" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "Difference", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "S" & ctr & ":S" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "SBR#", "OR#", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "T" & ctr & ":T" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Date Paid", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        mRange = "U" & ctr & ":U" & ctr + 1
        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Date Posted", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
        HeaderFlag = True
        ctr = ctr + 1
        cPos = ctr + 1
        WithRecSw = False
    End If
    
    If ListView1(0).ListItems.Item(i).Checked = True Then
        ' Count those excluded records.
        ExcludedCtr = ExcludedCtr + 1
        GoTo Next1
    ElseIf PrevCompany = ListView1(0).ListItems(i).SubItems(13) And (ListView1(0).ListItems(i).SubItems(4) = "__-_______-_" Or ListView1(0).ListItems(i).SubItems(4) = "") Then
        ' Count those no sss number.
        NoSSSCtr = NoSSSCtr + 1
        GoTo Next1
    Else
        ctr = ctr + 1
        xlSheet.Cells(ctr, 1).VALUE = LineCtr
        xlSheet.Cells(ctr, 2).VALUE = ListView1(0).ListItems(i).SubItems(1) ' Surname
        xlSheet.Cells(ctr, 3).VALUE = ListView1(0).ListItems(i).SubItems(2) ' firstname
        xlSheet.Cells(ctr, 4).VALUE = ListView1(0).ListItems(i).SubItems(3) ' Middlename
        xlSheet.Cells(ctr, 5).VALUE = ListView1(0).ListItems(i).SubItems(4) ' sss no
        xlSheet.Cells(ctr, 6).VALUE = ListView1(0).ListItems(i).SubItems(7) ' actual employee
        xlSheet.Cells(ctr, 7).VALUE = ListView1(0).ListItems(i).SubItems(8) ' actual employer
        
        xlSheet.Cells(ctr, 8).VALUE = ListView1(0).ListItems(i).SubItems(14) ' table based ee
        xlSheet.Cells(ctr, 9).VALUE = ListView1(0).ListItems(i).SubItems(15) ' table based er
        xlSheet.Cells(ctr, 10).VALUE = ListView1(0).ListItems(i).SubItems(9) ' ecc
        ' Total table based
        xlSheet.Cells(ctr, 11).VALUE = Val(ListView1(0).ListItems(i).SubItems(14)) + Val(ListView1(0).ListItems(i).SubItems(15)) + Val(ListView1(0).ListItems(i).SubItems(9))   ' table based total
        
        xlSheet.Cells(ctr, 12).VALUE = ListView1(0).ListItems(i).SubItems(10) ' employee adjust
        xlSheet.Cells(ctr, 13).VALUE = ListView1(0).ListItems(i).SubItems(11) ' employer adjust
        xlSheet.Cells(ctr, 14).VALUE = ListView1(0).ListItems(i).SubItems(9) ' ecc
        ' Total Adjusted
        xlSheet.Cells(ctr, 15).VALUE = Val(ListView1(0).ListItems(i).SubItems(10)) + Val(ListView1(0).ListItems(i).SubItems(11)) + Val(ListView1(0).ListItems(i).SubItems(9))   ' table based total
        ' Difference EE
        xlSheet.Cells(ctr, 16).VALUE = Val(ListView1(0).ListItems(i).SubItems(7)) - Val(ListView1(0).ListItems(i).SubItems(10))
        ' Difference ER
        xlSheet.Cells(ctr, 17).VALUE = Val(ListView1(0).ListItems(i).SubItems(8)) - Val(ListView1(0).ListItems(i).SubItems(11))
        ' Total Difference
        xlSheet.Cells(ctr, 18).VALUE = (Val(ListView1(0).ListItems(i).SubItems(7)) - Val(ListView1(0).ListItems(i).SubItems(10))) + (Val(ListView1(0).ListItems(i).SubItems(8)) - Val(ListView1(0).ListItems(i).SubItems(11)))
        
        'xlSheet.Cells(ctr, 13).VALUE = ListView1(0).ListItems(i).SubItems(12) ' difference
        LineCtr = LineCtr + 1
        WithRecSw = True
    End If

Next1:
    
    i = i + 1

Loop    ' End of Main Loop

If WithRecSw = True Then
    ' Print Sub-Totals.
    ctr = ctr + 1
    xlApp1.Sheets("Sheet2").Select
    ctr2 = ctr2 + 1
    mRange1 = "A" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
    mRange1 = "B" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
    mRange1 = "C" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
    
    For l = 70 To 82    ' from F to R
        xlApp1.Sheets("Sheet1").Select
        ' Compute for Total.
        mRange1 = Chr(l) & ctr
        xlSheet.Range(mRange1).Activate
        xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
        ' Save the Total.
        xlSheet.Range(mRange1).Select
        xlApp1.Selection.Copy
        ' Paste total to Sheet 2.
        xlApp1.Sheets("Sheet2").Select
        mRange1 = Chr(l - 2) & ctr2
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
    Next l
    
    xlApp1.Sheets("Sheet1").Select
    mRange = "F" & cPos & ":R" & ctr
    xlSheet.Range(mRange).Select
    xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
    mRange = "F" & ctr & ":R" & ctr
    xlSheet.Range(mRange).Select
    xlApp1.Range(mRange).Font.Bold = True
    
    If SumTotal = "" Then
        SumTotal = "=F" & ctr
    Else
        SumTotal = SumTotal & "+F" & ctr
    End If
Else    ' in case no more details
    ' Print Sub-Totals.
    ctr2 = ctr2 + 1
    xlApp1.Sheets("Sheet2").Select
    mRange1 = "A" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
    mRange1 = "B" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
    mRange1 = "C" & ctr2
    xlApp1.Range(mRange1).Select
    xlApp1.ActiveCell.FormulaR1C1 = 0
    
    For l = 68 To 80    ' from D to P
        mRange1 = Chr(l) & ctr2
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = 0#
    Next l
    xlApp1.Sheets("Sheet1").Select
End If
If NoSSSCtr > 0 Then
    LineCtr2 = LineCtr2 + 1
    ctr = PrintNoSSSWithSBR(xlSheet, xlApp1, PrevCompany, ctr)
End If

ctr = ctr + 2
mRange = "B" & ctr & ":D" & ctr
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Totals", "", xlLeft, , False, False, False, , , 10, True)
GTotal = 0

For l = 70 To 82  ' F to R
    mRange = Chr(l) & ctr
    xlSheet.Range(mRange).Activate
    xlSheet.Cells(ctr, l - 64).VALUE = SumTotal
    If l = 76 Then
        xlSheet.Cells(ctr + 2, 11).VALUE = "Adjusted EE"
        xlSheet.Cells(ctr + 2, 12).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 2, 12).VALUE)
    ElseIf l = 77 Then
        xlSheet.Cells(ctr + 3, 11).VALUE = "Adjusted ER"
        xlSheet.Cells(ctr + 3, 12).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 3, 12).VALUE)
        xlSheet.Cells(ctr + 4, 11).VALUE = "Sub-Total"
        xlSheet.Cells(ctr + 4, 12).VALUE = GTotal
        'GTotal = GTotal + Val(xlSheet.Cells(ctr + 3, 12).VALUE)
    ElseIf l = 78 Then
        xlSheet.Cells(ctr + 5, 11).VALUE = "ECC"
        xlSheet.Cells(ctr + 5, 12).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 5, 12).VALUE)
    End If
    SumTotal = Replace(SumTotal, Chr(l), Chr(l + 1))
Next l
mRange = "F" & cPos & ":R" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":R" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

ctr = ctr + 5 + 1
xlSheet.Cells(ctr, 11).VALUE = "Grand Total"
xlSheet.Cells(ctr, 12).VALUE = GTotal

mRange = "F" & cPos & ":R" & ctr
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "F" & ctr & ":R" & ctr
xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

'xlApp1.ActiveCell.FormulaR1C1 = "=R[-45]C+R[-38]C+R[-2]C"
'sumString = "=SUM(R[" & i * -1 & "]C:R[-2]C)"

' Print Unremitted SSS Contributions.
If ExcludedCtr <> 0 Then
    ctr = UnremittedSSSWithSBR(xlSheet, xlApp1, ctr)
End If

' ReFormat Sheet 2.
xlApp1.Sheets("Sheet2").Select
ctr2 = ctr2 + 1
'mRange = "C7:I" & ctr2
'xlApp1.Range(mRange).Select
GTotal = 0
For l = 67 To 80 ' C to P
    mRange1 = Chr(l) & ctr2
    xlApp1.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr2 - 6) * -1 & "]C:R[-1]C)"
    If l = 74 Then
        xlApp1.Range(mRange1).Activate
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.Copy
        mRange1 = "J" & ctr2 + 2
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = "Adjusted EE"
        mRange1 = "K" & ctr2 + 2
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        GTotal = GTotal + Val(xlApp1.ActiveCell.FormulaR1C1)
    ElseIf l = 75 Then
        xlApp1.Range(mRange1).Activate
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.Copy
        mRange1 = "J" & ctr2 + 3
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = "Adjusted ER"
        mRange1 = "K" & ctr2 + 3
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        GTotal = GTotal + Val(xlApp1.ActiveCell.FormulaR1C1)
    ElseIf l = 76 Then
        xlApp1.Range(mRange1).Activate
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.Copy
        mRange1 = "J" & ctr2 + 4
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = "ECC"
        mRange1 = "K" & ctr2 + 4
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        GTotal = GTotal + Val(xlApp1.ActiveCell.FormulaR1C1)
        
        ' Copy it also in SSS Transmittal.
        xlApp1.Sheets("Sheet4").Select
        xlApp1.Range("C15").Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        xlApp1.Sheets("Sheet2").Select
    End If
Next l
mRange = "D7:P" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "C" & ctr2 & ":P" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

ctr2 = ctr2 + 5
mRange1 = "J" & ctr2
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = "Grand Total"
mRange1 = "K" & ctr2
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = GTotal

' Get the Grand Total for SSS Transmittal.
xlApp1.Selection.Copy
xlApp1.Sheets("Sheet4").Select
xlApp1.Range("E15").Select
xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=False
xlApp1.Sheets("Sheet2").Select

mRange1 = "D" & cPos & ":P" & ctr2
xlApp1.Range(mRange1).Select
xlApp1.Range(mRange1).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "D" & ctr2 & ":P" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

mHead = "Prepared by:"
mRange = "B" & ctr2 + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mHead = "Noted by:"
mRange = "H" & ctr2 + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)

mRange = "B" & ctr2 + 5
Call PrintReportHeader(xlSheet, xlApp1, mRange, mUserID, "", xlLeft, , False, False, False, , , 9, True)

mRange = "H" & ctr2 + 5
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Rhiza Capati", "", xlLeft, , False, False, False, , , 9, True)

mHead = "HR Officer"
mRange = "B" & ctr2 + 6
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mHead = "General Manager"
mRange = "H" & ctr2 + 6
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
ctr2 = ctr2 + 8

mRange = "A6:S" & ctr2
xlApp1.Range(mRange).Select
With xlApp1
    '.Selection.Columns.AutoFit
    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With xlApp1.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End With
With xlApp1.Range(mRange)
    .Font.Name = "Arial"   'Font.Name
    .Font.Size = 8         'Font.size
End With

mRange = "D7:P" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "C" & ctr2 & ":P" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

mRange = "A6:S" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Selection.Columns.AutoFit
xlApp1.Range("A1").Select

' ReFormat Sheet 3.
xlApp1.Sheets("Sheet3").Select
ctr3 = ctr3 + 1
'mRange = "C7:I" & ctr3
'xlApp1.Range(mRange).Select
GTotal = 0
For l = 67 To 80 ' C to P
    mRange1 = Chr(l) & ctr3
    xlApp1.Range(mRange1).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr3 - 6) * -1 & "]C:R[-1]C)"
    If l = 68 Then ' 74
        xlApp1.Range(mRange1).Activate
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.Copy
        mRange1 = "J" & ctr3 + 2
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = "Actual EE Share"
        mRange1 = "K" & ctr3 + 2
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        GTotal = GTotal + Val(xlApp1.ActiveCell.FormulaR1C1)
    ElseIf l = 69 Then ' 75
        xlApp1.Range(mRange1).Activate
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.Copy
        mRange1 = "J" & ctr3 + 3
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = "Actual ER Share"
        mRange1 = "K" & ctr3 + 3
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        GTotal = GTotal + Val(xlApp1.ActiveCell.FormulaR1C1)
    ElseIf l = 72 Then ' 76
        xlApp1.Range(mRange1).Activate
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.Copy
        mRange1 = "J" & ctr3 + 4
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = "ECC"
        mRange1 = "K" & ctr3 + 4
        xlApp1.Range(mRange1).Select
        xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        GTotal = GTotal + Val(xlApp1.ActiveCell.FormulaR1C1)
    End If
Next l
mRange = "D7:P" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "C" & ctr3 & ":P" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

ctr3 = ctr3 + 5
mRange1 = "J" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = "Grand Total"
mRange1 = "K" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.ActiveCell.FormulaR1C1 = GTotal

mRange1 = "D" & cPos & ":P" & ctr3
xlApp1.Range(mRange1).Select
xlApp1.Range(mRange1).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "D" & ctr2 & ":P" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

mHead = "Prepared by:"
mRange = "B" & ctr3 + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mHead = "Noted by:"
mRange = "H" & ctr3 + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)

mRange = "B" & ctr3 + 5
Call PrintReportHeader(xlSheet, xlApp1, mRange, mUserID, "", xlLeft, , False, False, False, , , 9, True)

mRange = "H" & ctr3 + 5
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Rhiza Capati", "", xlLeft, , False, False, False, , , 9, True)

mHead = "HR Officer"
mRange = "B" & ctr3 + 6
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mHead = "General Manager"
mRange = "H" & ctr3 + 6
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)

ctr3 = ctr3 + 8

mRange = "A6:S" & ctr3
xlApp1.Range(mRange).Select
With xlApp1
    '.Selection.Columns.AutoFit
    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With xlApp1.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End With
With xlApp1.Range(mRange)
    .Font.Name = "Arial"   'Font.Name
    .Font.Size = 8         'Font.size
End With

mRange = "D7:P" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
mRange = "C" & ctr3 & ":P" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Range(mRange).Font.Bold = True

mRange = "A6:S" & ctr3
xlApp1.Range(mRange).Select
xlApp1.Selection.Columns.AutoFit
xlApp1.Range("A1").Select

Call PrintSSSTransmittal(xlSheet, xlApp1)
'                   Date Transmitted:   January 2009
'   Company         Applicable Month:   October 2008
'   SSS# 03-9210478-2

'   SSS Payments    EC  Total       SBR NUMBER/OR#      Date Paid

xlApp1.Sheets("Sheet1").Select

mHead = "Prepared by:"
mRange = "B" & ctr + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mHead = "Noted by:"
mRange = "H" & ctr + 2
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)

mRange = "B" & ctr + 5
Call PrintReportHeader(xlSheet, xlApp1, mRange, mUserID, "", xlLeft, , False, False, False, , , 9, True)

mRange = "H" & ctr + 5
Call PrintReportHeader(xlSheet, xlApp1, mRange, "Rhiza Capati", "", xlLeft, , False, False, False, , , 9, True)

mHead = "HR Officer"
mRange = "B" & ctr + 6
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)
mHead = "General Manager"
mRange = "H" & ctr + 6
Call PrintReportHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 9, False)

ctr = ctr + 8
'Box it.
mRange = "A7:U" & ctr  ' A6
Call PlaceBox(xlSheet, xlApp1, mRange)

xlApp1.Sheets("Sheet1").Name = "SSS Details"
xlApp1.Sheets("Sheet2").Name = "SSS Summary"
xlApp1.Sheets("Sheet3").Name = "Not to be Remitted SSS"
xlApp1.Sheets("Sheet4").Name = "SSS Transmittal"

xlSheet.Range("A1").Select

ExitRtn2:

End Sub

Private Sub ListView1_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
Dim itmFound As ListItem   ' FoundItem variable.
'
'Set itmFound = ListView1(0).FindItem(ListView1(Index).SelectedItem.Text, lvwText, lvwPartial)
''
''   ' If no ListItem is found, then inform user and exit. If a
''   ' ListItem is found, scroll the control using the EnsureVisible
''   ' method, and select the ListItem.
''   If itmFound Is Nothing Then  ' If no match, inform user and exit.
''   Else
'       itmFound.EnsureVisible ' Scroll ListView to show found ListItem.
'       itmFound.SELECTED = True   ' Select the ListItem.
''      ' Return focus to the control to see selection.
'       ListView1(0).SetFocus
''   End If
'''mClientName = ListView1(0).SelectedItem.Text
    
    Call UpdateSSSContributions(Index)
End Sub

'Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
'    Call UpdateSSSContributions
'End Sub

Private Sub UpdateSSSContributions(ByVal mIndex As Integer)
Dim mSurname As String
Dim mFirstName As String
Dim mMiddlename As String

mSurname = ListView1(mIndex).SelectedItem.ListSubItems(1).Text
mFirstName = ListView1(mIndex).SelectedItem.ListSubItems(2).Text
mMiddlename = ListView1(mIndex).SelectedItem.ListSubItems(3).Text

'mSurname = mListView.SelectedItem.ListSubItems(1).Text
'mFirstName = mListView.SelectedItem.ListSubItems(2).Text
'mMiddlename = mListView.SelectedItem.ListSubItems(3).Text

'MsgBox ListView1(Index).SelectedItem.ListSubItems(5).Text

            
            'If ListView1(0).ListItems.Item(i).Checked = True Then
            
If chkR1A.VALUE = 1 Then
    If ListView1(2).SelectedItem.Checked = True Then
        StrQuery = "UPDATE R1AContributions SET ExcludedFlag = 1 " & _
        " WHERE Surname = '" & mSurname & "' AND " & _
        " Firstname = '" & mFirstName & "' AND " & _
        " Middlename = '" & mMiddlename & "' AND " & _
        " GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "'"
    Else
        StrQuery = "UPDATE R1AContributions SET ExcludedFlag = 0 " & _
        " WHERE Surname = '" & mSurname & "' AND " & _
        " Firstname = '" & mFirstName & "' AND " & _
        " Middlename = '" & mMiddlename & "' AND " & _
        " GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "'"
    End If
Else
    If ListView1(0).SelectedItem.Checked = True Then
        StrQuery = "UPDATE SSSContributions SET ExcludedFlag = 1 " & _
        " WHERE Surname = '" & mSurname & "' AND " & _
        " Firstname = '" & mFirstName & "' AND " & _
        " Middlename = '" & mMiddlename & "' AND " & _
        " GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "'"
    '    ListView1(0).SelectedItem.ForeColor = vbRed
    Else
        StrQuery = "UPDATE SSSContributions SET ExcludedFlag = 0 " & _
        " WHERE Surname = '" & mSurname & "' AND " & _
        " Firstname = '" & mFirstName & "' AND " & _
        " Middlename = '" & mMiddlename & "' AND " & _
        " GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "'"
    End If
End If
Call Action_Query(StrQuery, cnnstr)

If chkR1A.VALUE = 1 Then
    Call GetTotalSSS
    txtRemittedAmt.Text = Format(mTEE + mTER + mECC, "0.00")
End If
End Sub

Private Sub GenerateR1AReports()
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

ExcelTemplate = App.Path & "\Reports\R-1A Template.xls"

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
Call getTotalPages(20)
i = 1
j = 7 '18
MemberCtr = 1

With xlSheet

Do While ListView1(2).ListItems.Count >= i

If ListView1(2).ListItems.Item(i).Checked = True Then
    GoTo Next1
Else
    If HeaderFlag = True Then
        l = 1
        For k = Asc("A") To Asc("M") ' SSS no.
            If (Chr(k) <> "C" And Chr(k) <> "L") Then
                mRange = Chr(k) & j
                If mID(txtPhilHealthNo, l, 1) = "-" Then
                    l = l + 1
                End If
                .Range(mRange) = mID(txtPhilHealthNo, l, 1)
                l = l + 1
            End If
        Next
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
        
        'j = j - 3
        mRange = "N" & j
        xlSheet.Range(mRange) = txtEmployerName
        mRange = "AM" & j
        xlSheet.Range(mRange) = txtEmpTIN
'        mRange = "H" & j
'        xlSheet.Range(mRange) = TxtPhilHealthNo
        
        j = j + 2
        mRange = "N" & j
        xlSheet.Range(mRange) = txtMailingAddress

        mRange = "A" & j
        xlSheet.Range(mRange) = txtTelephoneNo
        
        j = j + 3   ' start of detail line
        HeaderFlag = False
        MemberCtr = 0
        
        mRange1 = "K" & j + 21 ' Signature 33
        .Range(mRange1) = txtPrintedName.Text
        mRange1 = "K" & j + 24 ' Designation 36
        .Range(mRange1) = txtDesignation.Text
        
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
    
    ' Philhealth No
    '.Range(mRange1) = ListView1(1).ListItems(i).SubItems(4)
    txtPhilHealthNo.Text = Replace(Replace(ListView1(2).ListItems(i).SubItems(4), "-", ""), "_", "")
        l = 1
        For k = Asc("A") To Asc("J") ' SSS no.
            mRange = Chr(k) & j
            .Range(mRange) = mID(txtPhilHealthNo, l, 1)
            l = l + 1
        Next
    
    
    ' Employee Name
    mRange2 = "L" & j
    .Range(mRange2) = ListView1(2).ListItems(i).SubItems(1) & ", " & _
        ListView1(2).ListItems(i).SubItems(2) & " " & _
        IIf(ListView1(2).ListItems(i).SubItems(3) <> "", Left(ListView1(2).ListItems(i).SubItems(3), 1) & ".", "")
    ' Position
    'mRange1 = "D" & j
    '.Range(mRange1) = Format(ListView1(2).ListItems(i).SubItems(5), "mmddyyyy")
    txtPhilHealthNo = Replace(Replace(Format(ListView1(2).ListItems(i).SubItems(5), "mmddyyyy"), "/", ""), "_", "")
        l = 1
        For k = Asc("R") To Asc("Y") ' SSS no.
            mRange = Chr(k) & j
            .Range(mRange) = mID(txtPhilHealthNo, l, 1)
            l = l + 1
        Next
    
    
    ' Position
    mRange1 = "Z" & j
    .Range(mRange1) = ListView1(2).ListItems(i).SubItems(7)
    ' Salary
    'mRange1 = "E" & j
    '.Range(mRange1) = ListView1(2).ListItems(i).SubItems(6)
    txtPhilHealthNo = Replace(Format(ListView1(2).ListItems(i).SubItems(6), "000000"), ".", "")
        l = 1
        For k = Asc("A") To Asc("F") ' Salary Rate
            mRange = "A" & Chr(k) & j
            .Range(mRange) = mID(txtPhilHealthNo, l, 1)
            l = l + 1
        Next
    
    ' Employment Date
    'mRange1 = "F" & j
    '.Range(mRange1) = IIf(ListView1(2).ListItems(i).SubItems(21) = "", "__/__/____", Format(ListView1(2).ListItems(i).SubItems(21), "mm/dd/yyyy"))       ' get start contract
        
    txtPhilHealthNo = Replace(Replace(Format(ListView1(2).ListItems(i).SubItems(8), "mmddyyyy"), "/", ""), "_", "")
        l = 1
        For k = Asc("G") To Asc("N") ' Employment Date
            mRange = "A" & Chr(k) & j
            .Range(mRange) = mID(txtPhilHealthNo, l, 1)
            l = l + 1
        Next
        
        
    j = j + 1
    
    If MemberCtr = 20 Then
        ' Print Total Employees
        mRange1 = "H" & j + 1
        .Range(mRange1) = MemberCtr ' Total employees
        ' Print number of pages.
        mRange1 = "A" & j + 4
        .Range(mRange1) = "Page " & mPage & " Of " & gMpage & " pages"
        
        HeaderFlag = True
        xlApp1.ActiveSheet.Range("A1:H39").Select
        xlApp1.Selection.Copy
        
        j = j + 7
        
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
        j = j + 6
    End If
    End If

Next1:
    
    i = i + 1

Loop

'Print the last page.
    If MemberCtr <> 20 Then
        For m = MemberCtr To 19
'            ' Philhealth No
'            mRange1 = "B" & j
'            .Range(mRange1) = ""
'            ' Employee Name
'            mRange2 = "C" & j
'            .Range(mRange2) = ""
'            ' Position
'            mRange1 = "D" & j
'            .Range(mRange1) = ""
'            ' Salary
'            mRange1 = "E" & j
'            .Range(mRange1) = ""
'            ' Employment Date
'            mRange1 = "F" & j
'            .Range(mRange1) = ""
'            j = j + 1
        
    ' SSS No
        l = 1
        For k = Asc("A") To Asc("J") ' SSS no.
            mRange = Chr(k) & j
            .Range(mRange) = ""
            l = l + 1
        Next
    
    ' Employee Name
    mRange2 = "K" & j
    .Range(mRange2) = ""
    mRange2 = "L" & j
    .Range(mRange2) = ""
    ' BirtDate
        l = 1
        For k = Asc("R") To Asc("Y")
            mRange = Chr(k) & j
            .Range(mRange) = ""
            l = l + 1
        Next
    
    ' Position
    mRange1 = "Z" & j
    .Range(mRange1) = ""
    ' Salary Range
        l = 1
        For k = Asc("A") To Asc("F")
            mRange = "A" & Chr(k) & j
            .Range(mRange) = ""
            l = l + 1
        Next
    
    ' Employment Date
        l = 1
        For k = Asc("G") To Asc("N") ' Employment Date
            mRange = "A" & Chr(k) & j
            .Range(mRange) = ""
            l = l + 1
        Next
        
        j = j + 1
        
        Next
        
        ' Print Total Employees
        mRange1 = "H" & j + 1
        .Range(mRange1) = MemberCtr ' Total employees
        ' Print number of pages.
        mRange1 = "A" & j + IIf(gMpage = 1, 4, 4)
        .Range(mRange1) = "Page " & mPage & " Of " & gMpage & " pages"
        
    End If

End With

mRange1 = "A1"
xlSheet.Range(mRange1).Select

End Sub

Private Sub getTotalPages(ByVal RecPage As Integer)
Dim i As Integer

i = 1
TotEmp = 0
Do While ListView1(2).ListItems.Count >= i
    If i <> 0 And ListView1(2).ListItems.Item(i).Checked = False Then
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

