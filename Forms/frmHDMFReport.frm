VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmHDMFReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PagIBIG Report"
   ClientHeight    =   8520
   ClientLeft      =   2565
   ClientTop       =   1395
   ClientWidth     =   12600
   Icon            =   "frmHDMFReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   12600
   Begin VB.TextBox txtZipCode 
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
      Left            =   8760
      MaxLength       =   40
      TabIndex        =   17
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtSSSNo 
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
      Left            =   3720
      MaxLength       =   40
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdSearchSBR 
      Caption         =   "Search"
      Height          =   255
      Left            =   5760
      TabIndex        =   59
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
      TabIndex        =   63
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
      TabIndex        =   62
      Text            =   "0"
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdateSBR 
      Caption         =   "Update"
      Height          =   255
      Left            =   5760
      TabIndex        =   60
      ToolTipText     =   "Update SBR Entries"
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelectGroupCodes 
      Caption         =   "Select"
      Height          =   255
      Left            =   11760
      TabIndex        =   56
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
      ItemData        =   "frmHDMFReport.frx":08CA
      Left            =   5040
      List            =   "frmHDMFReport.frx":08F2
      Style           =   2  'Dropdown List
      TabIndex        =   55
      Top             =   6480
      Width           =   2085
   End
   Begin VB.CommandButton cmdClearView 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   240
      TabIndex        =   52
      ToolTipText     =   "process"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdExcelReport 
      Caption         =   "Generate Excel Report"
      Height          =   375
      Left            =   5880
      TabIndex        =   51
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
      TabIndex        =   47
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
      TabIndex        =   46
      Top             =   6590
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10200
      TabIndex        =   45
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "Retrieve Payments"
      Height          =   375
      Left            =   1560
      TabIndex        =   44
      ToolTipText     =   "process"
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CheckBox chkIncluded 
      Caption         =   "Exclude Checked Records"
      Height          =   255
      Left            =   2520
      TabIndex        =   43
      Top             =   6600
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdGenerateReport 
      Caption         =   "Generate Report to Disk"
      Height          =   375
      Left            =   8040
      TabIndex        =   40
      Top             =   7920
      Width           =   2055
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
      ItemData        =   "frmHDMFReport.frx":0958
      Left            =   10800
      List            =   "frmHDMFReport.frx":0971
      Style           =   2  'Dropdown List
      TabIndex        =   21
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
      TabIndex        =   39
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
      TabIndex        =   37
      Top             =   7125
      Width           =   2175
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Add/Delete Employees"
      Height          =   375
      Left            =   3720
      TabIndex        =   22
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
      ItemData        =   "frmHDMFReport.frx":099F
      Left            =   10800
      List            =   "frmHDMFReport.frx":09C7
      Style           =   2  'Dropdown List
      TabIndex        =   20
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   31
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
      TabIndex        =   24
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdCompany 
      Caption         =   "..."
      Height          =   375
      Left            =   5760
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   11
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
      TabIndex        =   8
      Top             =   2160
      Width           =   6735
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
      TabIndex        =   7
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   23
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
      TabIndex        =   41
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
      NumItems        =   21
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
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Index           =   0
      Left            =   240
      TabIndex        =   50
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
      TabIndex        =   57
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
      TabIndex        =   64
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
      TabIndex        =   61
      Top             =   7320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "ZIP:"
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
      Index           =   0
      Left            =   8400
      TabIndex        =   19
      Top             =   2160
      Width           =   375
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
      TabIndex        =   71
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
      TabIndex        =   70
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
      TabIndex        =   69
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
      TabIndex        =   68
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
      TabIndex        =   67
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
      TabIndex        =   66
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
      TabIndex        =   58
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
      TabIndex        =   54
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
      TabIndex        =   53
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
      TabIndex        =   49
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
      TabIndex        =   48
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
      TabIndex        =   42
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
      TabIndex        =   38
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   32
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
      TabIndex        =   30
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   12
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
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   2040
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
      TabIndex        =   9
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
      TabIndex        =   6
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
      Caption         =   "TIN:"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "PAGIBIG NO."
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
      Left            =   1680
      TabIndex        =   2
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
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "HDMF"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   165
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   240
      Picture         =   "frmHDMFReport.frx":0A2D
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmHDMFReport"
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
    ReportedFlag = 3
    SelectCompanyFlag = True
    frmCompany.Show
    frmCompany.SetFocus
End Sub

Private Sub cmdExcelReport_Click()
ListView1(1).SortKey = 0
ListView1(1).Sorted = True
ListView1(1).Refresh
If ListView1(1).ListItems.Count >= 1 Then
    Call GenerateHDMFReports
Else
    MsgBox "No records to process.", vbInformation, PayrollAlert
End If
ListView1(1).SortKey = 1
ListView1(1).Sorted = True
ListView1(1).Refresh
End Sub

Private Sub cmdGenerateReport_Click()
Dim X As Printer
Dim y As PrinterOrientationConstants

If ListView1(1).ListItems.Count >= 1 Then
    Call GenerateHDMFReportInExcel
Else
    MsgBox "No records to processed.", vbInformation, PayrollAlert
End If
   
End Sub

Private Sub GenerateHDMFReportInExcel()
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
Dim TotMember As Integer

ExcelTemplate = App.Path & "\Reports\Pagibig Template.xls"
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

ctr = ctr + 1
Call getTotalPages
i = 1
j = 3 '18
MemberCtr = 1
TotMember = 0
With xlSheet
Do While ListView1(1).ListItems.Count >= i

If ListView1(1).ListItems.Item(i).Checked = True Then
    GoTo Next1
Else
    If HeaderFlag = True Then
        mRange = "A" & j    ' A3
        xlSheet.Range(mRange) = cmdAppMonth.Text
        mRange = "C" & j    ' C3
        xlSheet.Range(mRange) = cmdAppYear.Text
        
        mRange = "H" & j    ' H3
        xlSheet.Range(mRange) = txtPhilHealthNo.Text
        
        j = j + 2
        mRange = "A" & j    ' A5
        xlSheet.Range(mRange) = txtEmployerName
        mRange = "F" & j    ' F5
        xlSheet.Range(mRange) = txtSSSNo.Text
        j = j + 2
        mRange = "A" & j ' A7
        xlSheet.Range(mRange) = txtMailingAddress.Text
        mRange = "F" & j ' A7
        xlSheet.Range(mRange) = txtZipCode.Text
        mRange = "G" & j ' A7
        xlSheet.Range(mRange) = txtEmpTIN.Text
        mRange = "I" & j    ' I7
        xlSheet.Range(mRange) = txtTelephoneNo.Text
        j = j + 3
        HeaderFlag = False
        MemberCtr = 0
        
        mRange1 = "F" & j + 44 ' Signature 35
        .Range(mRange1) = txtPrintedName.Text
        mRange1 = "F" & j + 46 ' Designation 36
        .Range(mRange1) = txtDesignation.Text
        
'        mRange1 = "A" & j + 19 ' Applicable Date
'        .Range(mRange1) = txtAppMonthYear(1).Text
'        mRange1 = "C" & j + 19 ' Remitted Amount
'        .Range(mRange1) = txtRemittedAmt.Text
'        mRange1 = "H" & j + 19 ' Acknowledgement Receipt
'        .Range(mRange1) = txtSBRNo.Text 'txtReceiptNo.Text
        mRange1 = "I" & j + 44 ' Transaction date
        .Range(mRange1) = MaskDatePaid.Text 'txtTransDate.Text
        
    
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
    
    mRange1 = "C" & j
    mRange2 = "D" & j
    mRange3 = "F" & j
    
    .Range(mRange1) = ListView1(1).ListItems(i).SubItems(1) ' Surname
    .Range(mRange2) = ListView1(1).ListItems(i).SubItems(2) ' Firstname
    .Range(mRange3) = ListView1(1).ListItems(i).SubItems(3) ' Middlename
    mRange1 = "A" & j
    .Range(mRange1) = ListView1(1).ListItems(i).SubItems(4) ' Middlename
        
'    l = 1
'    For k = Asc("U") To Asc("Z")
'        If Chr(k) <> "W" Then
'            mRange = Chr(k) & j
'            If Mid(ListView1(1).ListItems(i).SubItems(4), l, 1) = "-" Then
'                l = l + 1
'            End If
'            .Range(mRange) = Mid(ListView1(1).ListItems(i).SubItems(4), l, 1)
'            l = l + 1
'        End If
'    Next
'    For k = Asc("A") To Asc("H")
'        If Chr(k) <> "G" Then
'            mRange = "A" & Chr(k) & j
'            If Mid(ListView1(1).ListItems(i).SubItems(4), l, 1) = "-" Then
'                l = l + 1
'            End If
'            .Range(mRange) = Mid(ListView1(1).ListItems(i).SubItems(4), l, 1)
'            l = l + 1
'        End If
'    Next
    
'    mRange1 = "AJ" & j
'    .Range(mRange1) = ListView1(1).ListItems(i).SubItems(6) ' bracket
    mRange1 = "G" & j
    .Range(mRange1) = ListView1(1).ListItems(i).SubItems(9) ' adjusted ee
    mRange1 = "H" & j
    .Range(mRange1) = ListView1(1).ListItems(i).SubItems(10) ' adjusted er
    mRange1 = "I" & j
    .Range(mRange1) = Format(Val(ListView1(1).ListItems(i).SubItems(9)) + Val(ListView1(1).ListItems(i).SubItems(10)), "0.00") ' ee + er
    
    EETotal = EETotal + Val(ListView1(1).ListItems(i).SubItems(9))
    ERTotal = ERTotal + Val(ListView1(1).ListItems(i).SubItems(10))
    GEETotal = GEETotal + Val(ListView1(1).ListItems(i).SubItems(9))
    GERTotal = GERTotal + Val(ListView1(1).ListItems(i).SubItems(10))
    j = j + 1
    
    If MemberCtr = 40 Then
        ' Print Sub Total and Grand Total.
        mRange1 = "G" & j
        .Range(mRange1) = EETotal ' adjusted ee
        mRange1 = "H" & j
        .Range(mRange1) = ERTotal ' adjusted er
        mRange1 = "I" & j
        .Range(mRange1) = EETotal + ERTotal ' adjusted ee
        
        mRange1 = "G" & j + 1
        .Range(mRange1) = GEETotal ' adjusted ee
        mRange1 = "H" & j + 1
        .Range(mRange1) = GERTotal ' adjusted er
        mRange1 = "I" & j + 1
        .Range(mRange1) = GEETotal + GERTotal ' adjusted ee
        EETotal = 0
        ERTotal = 0
        ' Print number of pages.
        mRange1 = "I" & j + 6
        .Range(mRange1) = mPage
        mRange1 = "J" & j + 6
        .Range(mRange1) = gMpage
        
        mRange1 = "C" & j  ' No of Employees per page
        .Range(mRange1) = MemberCtr
        TotMember = TotMember + MemberCtr
        mRange1 = "E" & j  ' Total No of Employees
        .Range(mRange1) = TotMember
        
        HeaderFlag = True
        xlApp1.ActiveSheet.Range("A1:J65").Select
        xlApp1.Selection.Copy
        j = j + 16
        mRange1 = j & ":" & j
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 47.25
        
        ' Adjusted the RowHeight.
        mRange1 = j + 1 & ":" & j + 1
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 18.75
        mRange1 = j + 2 & ":" & j + 2
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 19.5
        mRange1 = j + 3 & ":" & j + 3
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 18
        mRange1 = j + 4 & ":" & j + 4
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 19.5
        mRange1 = j + 5 & ":" & j + 5
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 15.75
        mRange1 = j + 6 & ":" & j + 6
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 26.25
        mRange1 = j + 7 & ":" & j + 7
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 15.75
        mRange1 = j + 8 & ":" & j + 8
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 21
        mRange1 = j + 40 & ":" & j + 40
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 15
        mRange1 = j + 41 & ":" & j + 41
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 15
        mRange1 = j + 42 & ":" & j + 42
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 16.5
        mRange1 = j + 43 & ":" & j + 43
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 12.75
        mRange1 = j + 41 & ":" & j + 41
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 16.5
        mRange1 = j + 41 & ":" & j + 41
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 17.25
        mRange1 = j + 41 & ":" & j + 41
        xlApp1.ActiveSheet.Rows(mRange1).RowHeight = 16.5
        
        mRange1 = "A" & j
        xlApp1.ActiveSheet.Range(mRange1).Select
        xlApp1.ActiveSheet.Paste
        xlApp1.ActiveSheet.Range(mRange1).Select
        
        
        
        mPage = mPage + 1
        j = j + 2  '4
    End If
    End If

Next1:
    
    i = i + 1

Loop

'Print the last page.
    If MemberCtr <> 40 Then ' clear excess
        For m = MemberCtr To 39
            mRange1 = "C" & j
            mRange2 = "D" & j
            mRange3 = "F" & j
            
            .Range(mRange1) = "" ' Surname
            .Range(mRange2) = "" ' Firstname
            .Range(mRange3) = "" ' Middlename
            mRange1 = "A" & j
            .Range(mRange1) = "" ' Middlename
            
            mRange1 = "G" & j
            .Range(mRange1) = "-" ' adjusted ee
            mRange1 = "H" & j
            .Range(mRange1) = "-" ' adjusted er
            mRange1 = "I" & j
            .Range(mRange1) = "-" ' ee + er
            
            j = j + 1
        Next
        
        ' Print Sub Total and Grand Total.
        mRange1 = "G" & j
        .Range(mRange1) = EETotal ' adjusted ee
        mRange1 = "H" & j
        .Range(mRange1) = ERTotal ' adjusted er
        mRange1 = "I" & j
        .Range(mRange1) = EETotal + ERTotal ' adjusted ee
        
        mRange1 = "G" & j + 1
        .Range(mRange1) = GEETotal ' adjusted ee
        mRange1 = "H" & j + 1
        .Range(mRange1) = GERTotal ' adjusted er
        mRange1 = "I" & j + 1
        .Range(mRange1) = GEETotal + GERTotal ' adjusted ee
        
        ' Print number of pages.
        mRange1 = "I" & j + 6
        .Range(mRange1) = mPage
        mRange1 = "J" & j + 6
        .Range(mRange1) = gMpage
    
        mRange1 = "C" & j  ' No of Employees per page
        .Range(mRange1) = MemberCtr
        TotMember = TotMember + MemberCtr
        mRange1 = "E" & j  ' Total No of Employees
        .Range(mRange1) = TotMember
    
    End If

End With

End Sub

Private Sub getTotalPages()
Dim i As Integer

i = 1
TotEmp = 0
Do While ListView1(1).ListItems.Count >= i

    If i <> 0 And ListView1(1).ListItems.Item(i).Checked = False Then
        TotEmp = TotEmp + 1
    End If
    i = i + 1
Loop
gMpage = TotEmp \ 40
If TotEmp Mod 40 <> 0 Then
    gMpage = gMpage + 1
End If
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
lblclient(7).Caption = "No Pagibig No.: " & NoPHNo

End Sub

Private Sub cmdRetrieve_Click()
Dim litem1 As ListItem
Dim Listctr As Integer
Dim rsPHList As ADODB.Recordset
Dim i As Integer

i = 1
Listctr = 0
If chkAllPayments.VALUE = 1 Then    ' get all payments regardless of users
    'strQuery = "SELECT * FROM HDMFContributions WHERE ApplicableMonth = '" & cmdAppMonth.Text & "' and ApplicableYear = " & cmdAppYear.Text & " AND ProcessedBy = '" & mUserID & "' AND GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "' ORDER BY Surname, Firstname"
    StrQuery = "SELECT * FROM HDMFContributions WHERE GroupCode = '" & RTrim(LTrim(Replace(cmbGroupCode.Text, "New", ""))) & "' ORDER BY Surname, Firstname"
Else    ' get all payments for current user only
    StrQuery = "SELECT * FROM HDMFContributions WHERE LTRIM(RTRIM(ProcessedBy)) = '" & mUserID & "' AND GroupCode = '" & RTrim(LTrim(Replace(cmbGroupCode.Text, "New", ""))) & "' ORDER BY Surname, Firstname"
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
    Set litem1 = frmHDMFReport.ListView1(1).ListItems.Add(, , IIf(IsNull(!ClientName), "Blank", !ClientName))
        litem1.SubItems(1) = IIf(IsNull(!Surname), "", !Surname)
        litem1.SubItems(2) = IIf(IsNull(!FirstName), "", !FirstName)
        litem1.SubItems(3) = IIf(IsNull(!MiddleName), "", !MiddleName)
        litem1.SubItems(4) = IIf(IsNull(!PHealthSSSNo), "", !PHealthSSSNo)
        litem1.SubItems(5) = Format(!BasicPay, "0.##")
        litem1.SubItems(6) = IIf(IsNull(!SalaryBracket), "", !SalaryBracket)
        litem1.SubItems(7) = Format(!ActualEmployeeDed, "0.##")
        litem1.SubItems(8) = Format(!ActualEmployerDed, "0.##")
        'litem1.SubItems(9) = Format(!AdjustedECC, "0.##")
        litem1.SubItems(9) = Format(!AdjustedEmployeeDed, "0.##")
        litem1.SubItems(10) = Format(!AdjustedEmployerDed, "0.##")
        litem1.SubItems(11) = Format(!AmtDifference, "0.##")
        litem1.SubItems(12) = IIf(IsNull(!ClientName), "Blank", !ClientName)
        
        litem1.SubItems(13) = IIf(IsNull(!ProcessedBy), "", !ProcessedBy)      ' please activate
        litem1.SubItems(14) = IIf(IsNull(!ProcessedDate), "", !ProcessedDate)      ' please activate
        
        litem1.SubItems(16) = IIf(IsNull(!groupcode), "", !groupcode)
        litem1.SubItems(17) = Format(!AmountPaid, "0.00")
        litem1.SubItems(18) = IIf(IsNull(!SBRNo), "", !SBRNo)
        litem1.SubItems(19) = IIf(!DatePaid = #1/1/1900#, "__/__/____", Format(!DatePaid, "mm/dd/yyyy"))
        litem1.SubItems(20) = IIf(!DatePosted = #1/1/1900#, "__/__/____", Format(!DatePosted, "mm/dd/yyyy"))
        If Not IsNull(!SBRNo) And !SBRNo <> "" And !SBRNo <> "0" Then
            WithSBRNo = True
        End If
        If !ExcludedFlag = True Then
            ListView1(1).ListItems.Item(i).Checked = True
        End If
'
'        If !ExcludedFlag = 1 Then
'            frmHDMFReport.ListView1(1).ListItems.Item(i).Checked = True
'            'frmHDMFReport.ListView1(1).ListItems.Checked = True
'        End If
    .MoveNext
    i = i + 1
    
Wend
End With

Set rsPHList = Nothing

Call GetTotalHDMF

Exit22:

End Sub

Private Sub GetTotalHDMF()
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
        mTEE = mTEE + Val(ListView1(1).ListItems(i).SubItems(9))        ' employee adjust
        mTER = mTER + Val(ListView1(1).ListItems(i).SubItems(10))       ' employer adjust
        mTDiff = mTDiff + Val(ListView1(1).ListItems(i).SubItems(11))   ' difference
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
    StrQuery = "UPDATE HDMFContributions SET ExcludedFlag = 1 " & _
    " WHERE Surname = '" & mSurname & "' AND " & _
    " Firstname = '" & mFirstName & "' AND " & _
    " Middlename = '" & mMiddlename & "' AND " & _
    " GroupCode = '" & LTrim(RTrim(Replace(cmbGroupCode.Text, "New", ""))) & "'"
Else
    StrQuery = "UPDATE HDMFContributions SET ExcludedFlag = 0 " & _
    " WHERE Surname = '" & mSurname & "' AND " & _
    " Firstname = '" & mFirstName & "' AND " & _
    " Middlename = '" & mMiddlename & "' AND " & _
    " GroupCode = '" & LTrim(RTrim(Replace(cmbGroupCode.Text, "New", ""))) & "'"
End If
Call Action_Query(StrQuery, cnnstr)

Call GetTotalHDMF
txtRemittedAmt.Text = Format(mTEE + mTER, "0.00")

End Sub

Private Sub cmdSearchSBR_Click()

Dim rsPHList As ADODB.Recordset

If chkAllPayments.VALUE = 1 Then ' all payments regardless of users
    StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy, SBRNo " & _
        " FROM HDMFContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy, SBRNo HAVING SBRNo = '" & txtSBRNo.Text & "' AND GroupCode <> ''" & _
        " ORDER BY ProcessedBy, GroupCode"
Else    ' for current login only
    StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy, SBRNo " & _
        " FROM HDMFContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy, SBRNo HAVING SBRNo = '" & txtSBRNo.Text & "' AND GroupCode <> ''" & _
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
    Call frmHDMFClients.Show
    'frmHDMFClients.SetFocus
    'Call FrmOtherReport.Show
    'FrmOtherReport.SetFocus
    Call GetTotalHDMF
Else
    MsgBox "This group of clients was already paid. Cannot update!", vbInformation, PayrollAlert
End If

End Sub

Private Sub cmdSelectGroupCodes_Click()
    Call UpdateGroupCodes
End Sub

Private Sub UpdateGroupCodes()
Dim rsPHList As ADODB.Recordset

If R1AFlag = True Then
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
            " FROM HDMFContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy HAVING ApplicableMonth = '" & cmdAppMonth.Text & "' and ApplicableYear = " & cmdAppYear.Text & " AND GroupCode <> ''" & _
            " ORDER BY ProcessedBy, GroupCode"
    Else    ' for current login only
        StrQuery = "SELECT ApplicableMonth, ApplicableYear, GroupCode, MAX(RIGHT(GroupCode,2)) as GroupCtr, ProcessedBy " & _
            " FROM HDMFContributions GROUP BY ApplicableMonth, ApplicableYear, GroupCode, ProcessedBy HAVING ApplicableMonth = '" & cmdAppMonth.Text & "' and ApplicableYear = " & cmdAppYear.Text & " AND GroupCode <> ''" & _
            " AND ProcessedBy = '" & mUserID & "' ORDER BY ProcessedBy, GroupCode"
    End If
End If
Set rsPHList = ReturnRst(StrQuery, cnnstr, mstatic)
PHGroupCode = UCase(Left(mUserID, 5)) & "-" & Left(cmdAppMonth, 3) & cmdAppYear & "-01 New"
Call FillGroupCode(cmbGroupCode, rsPHList, PHGroupCode)
cmbGroupCode.ListIndex = 0
rsPHList.Close

End Sub

Private Sub cmdUpdateSBR_Click()

If MsgBox("Update SBR#/OR#?", vbQuestion + vbYesNo, "SBR Number Updates") = vbYes Then
    StrQuery = "UPDATE HDMFContributions SET SBRNO = '" & IIf(txtSBRNo.Text = "", "0", txtSBRNo.Text) & "', AmountPaid = " & IIf(txtRemittedAmt.Text = "", 0#, txtRemittedAmt.Text) & ", " & _
        " DatePaid = '" & IIf(MaskDatePaid.Text = "__/__/____", #1/1/1900#, Format(MaskDatePaid.Text, "mm/dd/yyyy")) & "', DatePosted = '" & IIf(MaskDatePosted.Text = "__/__/____", #1/1/1900#, Format(MaskDatePosted.Text, "mm/dd/yyyy")) & "' WHERE GroupCode = '" & Replace(cmbGroupCode.Text, "New", "") & "'"
    If Action_Query(StrQuery, cnnstr) = True Then
        Call cmdRetrieve_Click
        MsgBox "SBR/OR # successfully updated!", vbInformation, PayrollAlert
    Else
        MsgBox "SBR/OR # was not successfully updated!", vbInformation, PayrollAlert
    End If
End If

End Sub

Private Function FillGroupCode(ByRef mCmb As ComboBox, _
                        mRs As Recordset, Optional FirstGroup As String = "")
    mCmb.Clear
    While Not (mRs.EOF Or mRs.BOF)
        mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields("GroupCode")), FirstGroup, mRs.Fields("GroupCode")), "'", "''"))
        'mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
        'FirstGroup = UCase(Left(mRs.Fields("GroupCode"), 13)) & "-" & Right("00" & mRs.Fields("GroupCode") + 1, 2) & " New"
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
        Label1(0).Caption = "HDMF"
        Label1(2).Caption = "PAGIBIG NO."
        ListView1(1).Visible = True
        ListView1(0).Visible = False
        lblclient(7).Caption = "No Pagibig No."
        cmdNoPHNo.Caption = "Check/Uncheck No HDMF No."
        cmdGenerateReport.Visible = True
    End If
End Sub

Private Function PrintNoHDMFNumbers(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
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
mPCompany = mPCompany & " (Employees with no Pagibig#)"
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
        xlSheet.Cells(ctr, 11).VALUE = Val(ListView1(1).ListItems(j).SubItems(10)) + Val(ListView1(1).ListItems(j).SubItems(9)) ' EE + ER
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
        xlApp1.ActiveCell.FormulaR1C1 = mPCompany
        mRange1 = "C" & ctr2
        xlApp1.Range(mRange1).Select
        xlApp1.ActiveCell.FormulaR1C1 = xlSheet.Cells(ctr - 1, 1).VALUE
    End If
    xlApp1.Sheets("Sheet1").Select
Next l

'mRange = "F" & cPos & ":L" & ctr
'xlSheet.Range(mRange).Select
'mRange1 = "F" & ctr
'xlSheet.Range(mRange1).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - cPos) * -1 & "]C:R[-1]C)"
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

PrintNoHDMFNumbers = ctr

End Function

Private Sub GenerateHDMFReports()
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
Dim NoHDMFCtr As Integer
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
NoHDMFCtr = 0: ExcludedCtr = 0
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
mHead = "Pagibig Contributions Report for the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)

'Print Header of Sheet2 for Remitted Clients Summary.
mHead = CompanyName & " (" & CompanyShort & ")"
xlApp1.Sheets("Sheet2").Select

xlApp1.Range("A2").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "Pagibig Contributions Summary Report per Clients"
xlApp1.Range("A3").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "For the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text & "     Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
xlApp1.Range("A4").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead

For t = 65 To 80 ' A to P
    Call PrintHeadings(xlSheet, xlApp1, t, 6, , False)
Next t

'xlApp1.Range("A6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "No."
'xlApp1.Range("B6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Client Name"
'xlApp1.Range("C6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Total Employees"
'xlApp1.Range("D6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Actual EE Share"
'xlApp1.Range("E6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Actual ER Share"
'xlApp1.Range("F6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Actual EE+ER Share"
'xlApp1.Range("G6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Remitted EE"
'xlApp1.Range("H6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Remitted ER"
'xlApp1.Range("I6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Total EE+ER"
'xlApp1.Range("J6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Difference"
'xlApp1.Range("K6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Remarks"

'xlApp1.Range("A6:K6").Select
'xlApp1.Selection.Interior.ColorIndex = 15
'xlApp1.Selection.Interior.Pattern = xlSolid

'Print Header of Sheet3 for Unremitted Clients Summary.
xlApp1.Sheets("Sheet3").Select

mHead = CompanyName & " (" & CompanyShort & ")"
xlApp1.Range("A2").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "Pagibig Contributions Summary Report (NOT TO BE REMITTED CLIENTS)"
xlApp1.Range("A3").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead
mHead = "For the month of " & cmdAppMonth.Text & " " & cmdAppYear.Text & "     Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
xlApp1.Range("A4").Select
xlApp1.ActiveCell.FormulaR1C1 = mHead

For t = 65 To 80 ' A to P
    Call PrintHeadings(xlSheet, xlApp1, t, 6, , False)
Next t

'xlApp1.Range("A6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "No."
'xlApp1.Range("B6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Client Name"
'xlApp1.Range("C6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Total Employees"
'xlApp1.Range("D6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Actual EE Share"
'xlApp1.Range("E6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Actual ER Share"
'xlApp1.Range("F6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Actual EE+ER Share"
'xlApp1.Range("G6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Remitted EE"
'xlApp1.Range("H6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Remitted ER"
'xlApp1.Range("I6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Total EE+ER"
'xlApp1.Range("J6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Difference"
'xlApp1.Range("K6").Select
'xlApp1.ActiveCell.FormulaR1C1 = "Remarks"

'xlApp1.Range("A6:K6").Select
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
'
'            ' Copy to Sheet 2 the Total.
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
'            'LineCtr2 = LineCtr2 + 1
'
'            'Updated June 2009.
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
        
        If NoHDMFCtr > 0 Then
            ctr = PrintNoHDMFNumbers(xlSheet, xlApp1, PrevCompany, ctr)
            LineCtr2 = LineCtr2 + 1
        End If

        PrevCompany = ListView1(1).ListItems(i).SubItems(12)
'        mRange = "A" & ctr
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, PrevCompany, "", xlLeft, , False, False, False, , , 10, True)
        HeaderFlag = False
        NoHDMFCtr = 0  ' set again to 0 after printing
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
        Call PrintReportHeader(xlSheet, xlApp1, mRange, " Pagibig # ", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 11)
        mRange = "F" & ctr & ":F" & ctr + 1
        
        For t = 68 To 80 ' F to P
            Call PrintHeadings(xlSheet, xlApp1, t, ctr, 2, False)
        Next t
        
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'        mRange = "G" & ctr & ":G" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'        mRange = "H" & ctr & ":H" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual EE+ER", "Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
'        mRange = "I" & ctr & ":I" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Remitted", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'        mRange = "J" & ctr & ":J" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Remitted", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'        mRange = "K" & ctr & ":K" & ctr + 1
'        Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "Remitted EE+ER", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
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
        NoHDMFCtr = NoHDMFCtr + 1
        GoTo Next1
    Else
        ctr = ctr + 1
        xlSheet.Cells(ctr, 1).VALUE = LineCtr
        xlSheet.Cells(ctr, 2).VALUE = ListView1(1).ListItems(i).SubItems(1) ' Surname
        xlSheet.Cells(ctr, 3).VALUE = ListView1(1).ListItems(i).SubItems(2) ' Firstname
        xlSheet.Cells(ctr, 4).VALUE = ListView1(1).ListItems(i).SubItems(3) ' Middlename
        xlSheet.Cells(ctr, 5).VALUE = ListView1(1).ListItems(i).SubItems(4) ' Pagibig no
        xlSheet.Cells(ctr, 6).VALUE = ListView1(1).ListItems(i).SubItems(7) ' employee
        xlSheet.Cells(ctr, 7).VALUE = ListView1(1).ListItems(i).SubItems(8) ' employer
        xlSheet.Cells(ctr, 8).VALUE = Val(ListView1(1).ListItems(i).SubItems(7)) + Val(ListView1(1).ListItems(i).SubItems(8)) ' PS + ES Share
        xlSheet.Cells(ctr, 9).VALUE = ListView1(1).ListItems(i).SubItems(9) ' employee adjust
        xlSheet.Cells(ctr, 10).VALUE = ListView1(1).ListItems(i).SubItems(10) ' employer adjust
        xlSheet.Cells(ctr, 11).VALUE = Val(ListView1(1).ListItems(i).SubItems(9)) + Val(ListView1(1).ListItems(i).SubItems(10))  ' Total EE + ER
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
    
'    mRange = "F" & cPos + 1 & ":N" & ctr
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
                
'                ctr2 = ctr2 + 1
'                xlApp1.Sheets("Sheet2").Select
'                mRange1 = "A" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = LineCtr2
'                mRange1 = "B" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = PrevCompany
'                mRange1 = "C" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = 0
'                mRange1 = "D" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = 0#
'                mRange1 = "E" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = 0#
'                mRange1 = "F" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = 0#
'                mRange1 = "G" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = 0#
'                mRange1 = "H" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = 0#
'                mRange1 = "I" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = 0#
'                mRange1 = "J" & ctr2
'                xlApp1.Range(mRange1).Select
'                xlApp1.ActiveCell.FormulaR1C1 = 0#
'                xlApp1.Sheets("Sheet1").Select

End If

If NoHDMFCtr > 0 Then
    LineCtr2 = LineCtr2 + 1
    ctr = PrintNoHDMFNumbers(xlSheet, xlApp1, PrevCompany, ctr)
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

For l = 70 To 78  '75 F to K up to N
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

' Print Unremitted Pagibig Contributions.
If ExcludedCtr <> 0 Then
    ctr = UnremittedHDMFContributions(xlSheet, xlApp1, ctr)
End If

' ReFormat Sheet 2.
xlApp1.Sheets("Sheet2").Select
ctr2 = ctr2 + 1
GTotal = 0
For l = 68 To 76 ' 73 D to I up to K
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

mRange = "A6:P" & ctr2
xlApp1.Range(mRange).Select
xlApp1.Selection.Columns.AutoFit
xlApp1.Range("A1").Select

' ReFormat Sheet 3.
xlApp1.Sheets("Sheet3").Select
ctr3 = ctr3 + 1
GTotal = 0
For l = 68 To 76 ' 73 D to I
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

mRange = "A6:P" & ctr3
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
mRange = "A6:R" & ctr
Call PlaceBox(xlSheet, xlApp1, mRange)

xlApp1.Sheets("Sheet1").Name = "Pagibig Details"
xlApp1.Sheets("Sheet2").Name = "Pagibig Summary"
xlApp1.Sheets("Sheet3").Name = "Not to be Remitted Pagibig"

xlSheet.Range("A1").Select

ExitRtn2:

End Sub

Private Function UnremittedHDMFContributions(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal ctr As Integer) As Integer
Dim mRange As String, mRange1 As String
Dim LineCtr As Integer
Dim cPos As Integer
Dim i As Integer, l As Integer
Dim GTotal As Double

GTotal = 0#
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
                'mRange = "F" & cPos + 1 & ":N" & ctr
                mRange = "F" & cPos & ":N" & ctr
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
'                mRange = "F" & cPos & ":L" & ctr
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
'                'LineCtr3 = LineCtr3 + 1
'
'                'Updated June 2009
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
            NoHDMFCtr = 0  ' set again to 0 after printing
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
            Call PrintReportHeader(xlSheet, xlApp1, mRange, " Pagibig # ", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 11)
            
            For t = 68 To 80 ' F to P
                Call PrintHeadings(xlSheet, xlApp1, t, ctr, 2, False)
            Next t
            
'            mRange = "F" & ctr & ":F" & ctr + 1
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            mRange = "G" & ctr & ":G" & ctr + 1
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            mRange = "H" & ctr & ":H" & ctr + 1
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Actual EE+ER", "Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
'            mRange = "I" & ctr & ":I" & ctr + 1
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Remitted", "EE Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            mRange = "J" & ctr & ":J" & ctr + 1
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Remitted", "ER Share", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            mRange = "K" & ctr & ":K" & ctr + 1
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total", "Remitted EE+ER", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
'            mRange = "L" & ctr & ":L" & ctr + 1
'            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Difference", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9)
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
            xlSheet.Cells(ctr, 11).VALUE = Val(ListView1(1).ListItems(i).SubItems(10)) + Val(ListView1(1).ListItems(i).SubItems(9)) ' Total EE + ER
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
ctr3 = ctr3 + 1  ' Updated by Bernard 9/19/2009
mRange = "F" & cPos & ":N" & ctr
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
''Update June 2009.
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

For l = 70 To 78  '75 F to K up to N
    mRange = Chr(l) & ctr
    xlSheet.Range(mRange).Activate
    xlSheet.Cells(ctr, l - 64).VALUE = SumTotal
    If l = 70 Then
        xlSheet.Cells(ctr + 2, 9).VALUE = "Actual PS"
        xlSheet.Cells(ctr + 2, 10).VALUE = SumTotal
        GTotal = GTotal + Val(xlSheet.Cells(ctr + 2, 10).VALUE)
    ElseIf l = 71 Then
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

UnremittedHDMFContributions = ctr

End Function

