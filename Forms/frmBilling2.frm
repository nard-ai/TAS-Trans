VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBilling 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Billing Form"
   ClientHeight    =   8550
   ClientLeft      =   1500
   ClientTop       =   1875
   ClientWidth     =   12600
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmBilling.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmBilling.frx":08CA
   ScaleHeight     =   8550
   ScaleWidth      =   12600
   Begin VB.CommandButton cmdUpdateTemplate 
      Caption         =   "Edit Record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   40
      Top             =   8040
      Width           =   1755
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12375
      Begin VB.CheckBox chkNewBilling 
         Caption         =   "New Billing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11040
         TabIndex        =   8
         Top             =   720
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.TextBox txtSOASeq 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         TabIndex        =   37
         Text            =   "001"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtSOAMonth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "XX09-01"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1320
         TabIndex        =   36
         Text            =   "JC09-01"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Select Billed SOA Number"
         Top             =   630
         Width           =   705
      End
      Begin VB.TextBox txtAppliedYear 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11520
         TabIndex        =   29
         Text            =   "2009"
         ToolTipText     =   "applicable year"
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbSelectedCutOff 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   3615
      End
      Begin VB.CommandButton cmdShowCutoff 
         Caption         =   "&Show Payroll Cutoff"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10200
         TabIndex        =   2
         ToolTipText     =   "Show Selected Payroll Cutoff"
         Top             =   240
         Width           =   1995
      End
      Begin VB.ComboBox cmbClient 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmBilling.frx":0C0C
         Left            =   1320
         List            =   "frmBilling.frx":0C19
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
      Begin MSMask.MaskEdBox mskStartContract 
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
         Left            =   3960
         TabIndex        =   11
         Top             =   630
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
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
      Begin MSMask.MaskEdBox txtCutoffDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6480
         TabIndex        =   18
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtCutoffDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   8880
         TabIndex        =   19
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
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
      Begin VB.ComboBox cmbYear 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmBilling.frx":0C36
         Left            =   9120
         List            =   "frmBilling.frx":0C61
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cmbMonth 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmBilling.frx":0CB3
         Left            =   6480
         List            =   "frmBilling.frx":0CDB
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtDash 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2040
         TabIndex        =   38
         Text            =   "-"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtSOANo 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Text            =   "JC09-01-014"
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbBilledSOA 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   8400
         TabIndex        =   34
         Top             =   660
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "End Cutoff:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   7800
         TabIndex        =   21
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label lblEntries 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start Cutoff:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   20
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   10
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Cut-Off:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   7
         Top             =   300
         Width           =   1365
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "SOA Ctrl #:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   5
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Client:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdClearTrans 
      Caption         =   "Clear Records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   27
      Top             =   8040
      Width           =   1755
   End
   Begin VB.CommandButton cmdPrintSOA 
      Caption         =   "&Generate Billing Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   8040
      Width           =   2115
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   13
      Top             =   8040
      Width           =   1755
   End
   Begin VB.CommandButton cmdAddCutoff 
      Caption         =   "&Add/Delete Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   8040
      Width           =   2355
   End
   Begin VB.Frame Frame2 
      Height          =   6855
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   12375
      Begin VB.CommandButton cmdDeleteRec 
         Caption         =   "Delete Record"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   41
         Top             =   6390
         Width           =   1755
      End
      Begin VB.TextBox txtClientCode 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4200
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   6360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkExcludedRecords 
         Caption         =   "Excluded Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7800
         TabIndex        =   35
         Top             =   6390
         Width           =   1605
      End
      Begin VB.CheckBox chkSuccess 
         Caption         =   "New Billing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   31
         Top             =   6480
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txtGrossBill 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10800
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   6360
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6015
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   10610
         SortKey         =   2
         View            =   3
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
         NumItems        =   68
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Control #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Employee Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DTR"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Regular Days"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Edit Reg Days"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Day Off"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Edit Day Off"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "DayOff (Spl)"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Edit Day Off (Spl)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Sunday OT"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Edit Sunday OT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "DayOff (Legal)"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Edit DayOff (Lgl)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "#Days Worked"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Edit Days Worked"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Tardi Hours"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Edit Tardi Hours"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Reg. O.T."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Edit Reg O.T."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Dayoff O.T."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Edit DayOff O.T."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Dayoff (Spl) O.T."
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Legal O.T."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Edit Legal O.T."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "Dayoff (Lgl) O.T."
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Edit DayOff (Lgl) OT"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "Night Diff'l"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "Edit Night Diff'l"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "Daily Rate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "Basic Pay"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "DayOff Pay"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "DayOff Spl Pay"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "Legal Pay"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "DayOff Lgl Pay"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "Tardi (UT) Pay"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "Reg OT Pay"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "DayOff Pay"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "DayOff Spl Pay"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "Legal Pay"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Text            =   "DayOff Lgl Pay"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Text            =   "Night Diff'l Pay"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "ECOLA"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   43
            Text            =   "Salary Adjust"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   44
            Text            =   "Edit Salary Adjust"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   45
            Text            =   "Gross Earnings"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   46
            Text            =   "13th Month Pay"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   47
            Text            =   "SSS/EC (ER)"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   48
            Text            =   "PhilHealth (ER)"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   49
            Text            =   "PagIBIG (ER)"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   50
            Text            =   "Total Pay & Benefits"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   51
            Text            =   "Agency Fee"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   52
            Text            =   "Gross Billing"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   53
            Text            =   "VAT 12%"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   54
            Text            =   "Other Ded/Allowance"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   55
            Text            =   "Edit Other Ded/Allow"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   56
            Text            =   "Total Amount Due"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   57
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   58
            Text            =   "Gross Payroll"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   59
            Text            =   "Ignore (Lates/Ecola)"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   60
            Text            =   "Diff'n (over/under)"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   61
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   62
            Text            =   "Start Cutoff"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   63
            Text            =   "End Cutoff"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   64
            Text            =   "SOA Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   65
            Text            =   "SSS Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   66
            Text            =   "PHealth Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   67
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3615
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   6376
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   44
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ctrl#"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Employee Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DTR"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Reg Days"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Worked D.O."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "DayOff (Spl)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Worked Lgl"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "DayOff Lgl"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Days Worked"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Tardi Hours"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Reg. O.T."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "DayOff O.T."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "DayOff Spl OT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Legal O.T."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "DayOff Lgl OT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Night Diff'l"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Daily Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Basic Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "DayOff Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "DayOff Spl Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Legal Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "DayOff Lgl Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Tardi (UT) Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Reg OT Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "DayOff Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "DayOff Spl Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "Legal Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "DayOff Lgl Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "Night Diff'l Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "Ecola"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "Salary Adjust"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "Gross Earn"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "13th Month Pay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "SSS/EC (ER)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "PhilHealth (ER)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "PagIBIG (ER)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "Total Pay & Benefits"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "Agency Fee"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "Gross Billing"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Text            =   "Gross Payroll"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "Ignore (Lates/Ecola)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   43
            Text            =   "Diff'n (over/under)"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Records: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   24
         Top             =   6360
         Width           =   1605
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Billing:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   9600
         TabIndex        =   23
         Top             =   6420
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdRetrieveTrans 
      Caption         =   "Recompute Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   8040
      Width           =   2355
   End
   Begin VB.TextBox txtEmpNo 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5280
      TabIndex        =   26
      Text            =   "BJC09-01-014"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCutOffBase 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5640
      TabIndex        =   16
      Text            =   "BJC09-01-014"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCutOffTable As ADODB.Recordset
Dim rsClientTable As ADODB.Recordset
Dim rsPayrollTrans As ADODB.Recordset
Dim rsBilledTrans As ADODB.Recordset
Dim rsBilledTrans2 As ADODB.Recordset
Dim rsEmpBranch As ADODB.Recordset

Dim OnProcessFlag As Boolean
Dim NoDaysperCutoff As Integer
Dim strQuery As String
Dim EditCutoffSw As Boolean
Dim mSOANo As String
Dim mClientCode As Integer
Dim ClientAddress As String
Dim ClientContact1 As String
Dim ClientContactPos1 As String
Dim ClientContact2 As String
Dim prevSOA As String

Dim DaysPresent As Double
Dim DaysPresentKey As Double
Dim RestDayOTHrs As Double
Dim NoSplHoliday As Double
Dim SundayOT As Double
Dim NoLglHoliday As Double
Dim DaysWorked As Double
Dim HoursTardiness As Double
Dim OTHours As Double
Dim ExcessSplOTHours As Double
Dim LglHolidayOTHours As Double
Dim ExcessOTSunLegalHrs As Double
Dim NightDiffOT As Double
Dim SalaryAdjust As Double
Dim IncentivePay As Double

'Public gGrossBill As Double ' declare in public variable function

'Public RestDayWorkedOTRate As Double   ' 30%
'Public SpecialWorkedDayOffRate As Double   ' 50%
'Public LegalHolidayWithPremiumRate As Double   ' 100%
'Public LegalHolidayWorkedDayoffRate As Double  ' 160%
'Public WorkHrsPerDay As Double
'Public OverTimeRate As Double
'Public ExcessSplOTRate As Double
'Public ExcessDayoffSplOTRate As Double
'Public ExcessLglHolOTRate As Double
'Public ExcessSunLglOTRate As Double
'Public NightDiffRate As Double
'Public AgencyRate As Double
'Public VatPercentRate As Double
'Public BillingTemplate As String

Dim RestDayOTPay As Double
Dim SplHolidayPay As Double
Dim SundayOTPay As Double
Dim LglHolidayPay As Double

Dim TardinessDeduction As Double ' deduction
Dim OTPay As Double
Dim ExcessSplOTPay As Double
Dim LglHolidayOTPay As Double
Dim ExcessSunLegalPay As Double
Dim NightDiffPay As Double
Dim BasicPay As Double
Dim GrossPay As Double
Dim GrossBill As Double
Dim Month13Pay As Double
Dim VATPayment As Double

Dim AgencyCharge As Double

Dim SSSEmployer As Double
Dim PHealthEmployer As Double
Dim PagibigEmployer As Double
Dim FirstCutoff1 As Variant
Dim FirstCutoff2 As Variant
Dim SecondCutoff1 As Variant
Dim SecondCutoff2 As Variant

Private Sub chkNewBilling_Click()

If chkNewBilling.VALUE = 0 Then ' get those billed transactions
    mSOANo = txtSOANo.Text  ' save original SOA #
    txtSOANo.Enabled = False
    txtSOANo.Visible = False
    txtSOAMonth.Visible = False
    txtSOASeq.Visible = False
    txtDash.Visible = False
    'Call GetBilledSOA
    cmbBilledSOA.Enabled = True
    cmbBilledSOA.Visible = True
    'txtCutoffDate(2).Enabled = True
    txtCutoffDate(2).Visible = False
    txtCutoffDate(3).Visible = False
    lblEntries(4).Visible = False
    lblEntries(3).Caption = "Payroll Month:"
    cmbMonth.Visible = True
    cmbYear.Visible = True
    lblEntries(8).Visible = True
    'txtCutoffDate(2).Text = "__/__/____"
    lblEntries(4).Visible = False
    txtCutoffDate(3).Visible = False
    ListView1(0).ListItems.Clear
    cmdrefresh.Visible = True
    cmbSelectedCutOff.Enabled = False
Else
    txtSOANo.Text = mSOANo
    txtSOANo.Enabled = True
    'txtSOANo.Visible = True
    txtSOAMonth.Visible = True
    txtSOASeq.Visible = True
    txtDash.Visible = True
    
    cmbBilledSOA.Enabled = False
    cmbBilledSOA.Visible = False
    txtCutoffDate(2).Visible = True
    txtCutoffDate(3).Visible = True
    lblEntries(4).Visible = True
    'txtCutoffDate(2).Enabled = False
    lblEntries(3).Caption = "Start Cutoff:"
    cmbMonth.Visible = False
    cmbYear.Visible = False
    lblEntries(8).Visible = False
    'txtCutoffDate(2).Text = "__/__/____"
    lblEntries(4).Visible = True
    txtCutoffDate(3).Visible = True
    cmdrefresh.Visible = False
    Call GetSelectedCutoff(False)
    cmbSelectedCutOff.Enabled = True
End If

End Sub

Private Sub GetBilledSOA()
Dim mFind As String
Dim pSOAno As String
Dim soactr As Integer

soactr = 0
'If txtCutoffDate(2).Text = "__/__/____" Then
If cmbMonth.Text = "" Then
    MsgBox "Please select Billed month and year for filtering of Billing records.", vbInformation, PayrollAlert
    Exit Sub
End If
pSOAno = ""
'mFind = " SearchKey = " & Format(txtCutoffDate(2), "yyyy") & Format(txtCutoffDate(2), "mm") & mClientCode  'goto selected client record
    rsBilledTrans.Requery
    If rsBilledTrans.RecordCount = 0 Then
        MsgBox "No records selected.", vbInformation, PayrollAlert
        GoTo ExitRtn1
    End If
    rsBilledTrans.MoveFirst
    'rsBilledTrans.Find " SearchKey = " & Format(txtCutoffDate(2), "yyyy") & Format(txtCutoffDate(2), "mm") & mClientCode, 0, adSearchForward ' goto selected client record
    'rsBilledTrans.Filter = " SearchKey = " & Format(txtCutoffDate(2), "yyyy") & Format(txtCutoffDate(2), "mm") & mClientCode ' filter the records
    rsBilledTrans.Filter = " SearchKey = " & cmbYear.Text & Left(cmbMonth.Text, 2) & mClientCode ' filter the records
    cmbBilledSOA.Clear
    If Not rsBilledTrans.EOF Then
        With rsBilledTrans
            While Not (.EOF Or .BOF)
                If pSOAno <> LTrim(RTrim(!soano)) Then
                    cmbBilledSOA.AddItem LTrim(RTrim(!soano))
                    pSOAno = LTrim(RTrim(!soano))
                    soactr = soactr + 1
                End If
                .MoveNext
            Wend
            If cmbBilledSOA.ListIndex = -1 Then
                cmbBilledSOA.ListIndex = 0
                Call GetSelectedCutoff(True)
            End If
            cmbBilledSOA.Refresh
            MsgBox "There were " & soactr & " SOA selected.", vbInformation, PayrollAlert
            cmbSelectedCutOff.Enabled = True
        End With
    Else
        MsgBox "No records selected.", vbInformation, PayrollAlert
        cmbSelectedCutOff.Enabled = False
    End If
    rsBilledTrans.Filter = ""

ExitRtn1:

End Sub

Private Sub cmbBilledSOA_Click()
txtSOANo.Text = cmbBilledSOA.Text
End Sub

Private Function GetBranchCode(ByVal EmpNm As String, ByVal RetSw As Integer) As Variant
Dim mReturn As Variant

    If rsEmpBranch.RecordCount <> 0 Then
        rsEmpBranch.MoveFirst
        rsEmpBranch.Find "Employees = '" & EmpNm & "'", 0, adSearchForward
        If Not rsEmpBranch.EOF Then
            If RetSw = 1 Then  ' return branch name
                mReturn = IIf(IsNull(rsEmpBranch!BrName), "No Area Specified.", rsEmpBranch!BrName)
            Else            ' return branch code
                mReturn = IIf(IsNull(rsEmpBranch!areacode), 0, rsEmpBranch!areacode)
            End If
        Else
            If RetSw = 1 Then  ' return branch name
                mReturn = "No Area Specified."
            Else                ' return branch code
                mReturn = 0
            End If
        End If
    End If
    GetBranchCode = mReturn
End Function

Private Sub cmbClient_Click()
    'Get Cutoff table based from selected client.
    mClientCode = cmbClient.ItemData(cmbClient.ListIndex)
    txtClientCode = mClientCode
    If chkSuccess.VALUE = 1 Then
        If frmBillingUpdate.txtClientCode.Text <> "" Then
            mClientCode = Val(frmBillingUpdate.txtClientCode.Text)
        End If
    End If
    rsClientTable.Requery   ' requery the data for any updates
    rsClientTable.MoveFirst
    rsClientTable.Find "ClientCode = " & mClientCode, 0, adSearchForward ' goto selected client record
    
    With rsClientTable
    
    If Not rsClientTable.EOF Then
        NoDaysperCutoff = IIf(IsNull(rsClientTable!ClientCutoffDates), 0, Val(rsClientTable!ClientCutoffDates))
        ClientAddress = IIf(IsNull(rsClientTable!ClientAddress), "No Address", rsClientTable!ClientAddress)
        ClientContact1 = IIf(IsNull(rsClientTable!ClientContactPerson), "No Contact Name", rsClientTable!ClientContactPerson)
        ClientContactPos1 = IIf(IsNull(rsClientTable!ClientContactPosition), "N/A", rsClientTable!ClientContactPosition)
        ClientContact2 = "N/A"
        'txtCutOffBase.Text = rsClientTable!ClientCutoff & " (" & Trim(NoDaysperCutoff) & ")"
        txtCutOffBase.Text = mClientCode
'        mSOANo = IIf(IsNull(rsClientTable!SOANoCode), "XX09-01", LTrim(RTrim(rsClientTable!SOANoCode))) & "-" & Right("000" & LTrim(RTrim(str(rsClientTable!soano + 1))), 3)
'        txtSOANo.Text = mSOANo
        txtSOAMonth = IIf(IsNull(rsClientTable!SOANoCode), "XX09-01", LTrim(RTrim(rsClientTable!SOANoCode)))
        If txtSOAMonth = "XX09-01" And prevSOA <> "" Then
            txtSOAMonth = prevSOA
        End If
        txtSOASeq = Right("000" & LTrim(RTrim(str(rsClientTable!soano + 1))), 3)
        mSOANo = txtSOAMonth & "-" & txtSOASeq
        txtSOANo.Text = mSOANo
        
        RestDayWorkedOTRate = Format(!RestDayRate, "0.00")  ' 0.3
        SpecialWorkedDayOffRate = Format(!SpecialHolidayRate, "0.00")  '0.5
        LegalHolidayWithPremiumRate = Format(!OTwoPremiumRate, "0.00") '1#
        LegalHolidayWorkedDayoffRate = Format(!legalholidayotrate, "0.00") '1.6
        WorkHrsPerDay = !WorkHrsPerDay '8
        
        OverTimeRate = Format(!OverTimeRate, "0.00") '1.25
        ExcessSplOTRate = Format(!ExcessOTSplSunRate, "0.00") '1.69
        ExcessDayoffSplOTRate = Format(!LegalHolidayRate, "0.00") '1.95
        ExcessLglHolOTRate = Format(!ExcessOTLegalRate, "0.00")  '2.6
        ExcessSunLglOTRate = Format(!ExcessOTSunLglRate, "0.00")   '3.38
        NightDiffRate = Format(!NDiffRate, "0.00")  '0.1
        
        AgencyRate = Format(!ClientAgencyRate, "0.00") ' 15
        VatPercentRate = Format(!VatPercentRate, "0.00")  '0.12
        BillingTemplate = IIf(IsNull(!BillingTemplate) Or !BillingTemplate = "", "Other Clients SOA Template.xls", !BillingTemplate)
        
        Call GetSelectedCutoff(False)
    Else
        MsgBox "Invalid client code selected!", vbInformation, PayrollAlert
    End If
    
    End With
End Sub

Private Sub cmbClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbCutoff_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbSelectedCutoff_Click()
Dim CutOffIndex As Integer
Dim mEndDate As Date

    If cmbSelectedCutOff.ListIndex = -1 Then
        Exit Sub
    Else
        CutOffIndex = IIf(cmbSelectedCutOff.ListIndex = -1, 0, cmbSelectedCutOff.ItemData(cmbSelectedCutOff.ListIndex))
    End If
    If cmbSelectedCutOff.Text <> "All Cutoff" Then
        txtCutoffDate(2).Text = Format(Left(cmbSelectedCutOff.Text, 10), "mm/dd/yyyy")
        txtCutoffDate(3).Text = Format(Mid(cmbSelectedCutOff.Text, 12, 10), "mm/dd/yyyy")
    End If
    
End Sub

Private Sub GetSelectedCutoff(ByVal BilledFlag As Boolean)
Dim rsSelectedCutoff As ADODB.Recordset
Dim prevCutOff As String

prevCutOff = ""
If BilledFlag = False Then
    ' Original Code
'    strQuery = "SELECT a.*, b.EntryDate FROM CutOffTable a LEFT JOIN qryPayrollCutoff b ON a.agencycode = b.agencycode and a.clientcode = b.clientcode and a.StartcutoffText = b.StartCutoffText and a.EndCutoffText = b.EndCutoffText WHERE a.AgencyCode = " & CompanyCode & " and a.ClientCode = " & Val(frmBilling.txtCutOffBase.Text) & _
'        " and a.SelectedFlag = True and a.ViewFlag = True AND a.SOANo = '" & txtSOANo.Text & "' ORDER BY StartCutOffDate"
    strQuery = "SELECT DISTINCT * FROM BillingCutoffTable a WHERE a.AgencyCode = " & CompanyCode & " and a.ClientCode = " & Val(frmBilling.txtCutOffBase.Text) & _
        " and a.SelectedFlag = True and a.ViewFlag = True AND a.SOANo = '" & txtSOANo.Text & "' ORDER BY StartCutOffDate"

'    strQuery = "SELECT * FROM CutOffTable WHERE AgencyCode = " & CompanyCode & " and ClientCode = " & Val(frmBilling.txtCutOffBase.Text) & _
'        " and SelectedFlag = True and ViewFlag = True AND SOANo = '" & txtSOANo.Text & "' ORDER BY StartCutOffDate"
    Set rsSelectedCutoff = ReturnRst(strQuery, cnnStr, mstatic)
    With rsSelectedCutoff
        cmbSelectedCutOff.Clear
        cmbSelectedCutOff.AddItem "All Cutoff"
        While Not (.EOF Or .BOF)
            'If prevCutOff <> Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & ":" & Format(!ProcessDate, "mmddyyyy") & " Sel" Then
            If prevCutOff <> Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & ":" & Format(!EntryDate, "mmddyyyy") & " Sel" Then
                'cmbSelectedCutOff.AddItem Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & ":" & Format(!ProcessDate, "mmddyyyy") & " Sel"
                cmbSelectedCutOff.AddItem Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & ":" & Format(!EntryDate, "mmddyyyy") & " Sel"
                'cmbSelectedCutOff.ItemData(cmbSelectedCutOff.NewIndex) = .Fields(0)
                'prevCutOff = Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & ":" & Format(!ProcessDate, "mmddyyyy") & " Sel"
                prevCutOff = Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & ":" & Format(!EntryDate, "mmddyyyy") & " Sel"
            End If
            .MoveNext
        Wend
        If cmbSelectedCutOff.ListIndex = -1 Then
            cmbSelectedCutOff.ListIndex = 0
        End If
    End With
    rsSelectedCutoff.Close
Else    ' for Billed records
    rsBilledTrans.MoveFirst
    rsBilledTrans.Filter = " SOANo = '" & cmbBilledSOA.Text & "'"
    With rsBilledTrans
        cmbSelectedCutOff.Clear
        cmbSelectedCutOff.AddItem "All Cutoff"
        While Not (.EOF Or .BOF)
            If prevCutOff <> Format(!StartPayDate, "mm/dd/yyyy") & "-" & Format(!EndPayDate, "mm/dd/yyyy") & ":" & LTrim(RTrim(!Paytext2)) & " Bil" Then
                cmbSelectedCutOff.AddItem Format(!StartPayDate, "mm/dd/yyyy") & "-" & Format(!EndPayDate, "mm/dd/yyyy") & ":" & LTrim(RTrim(!Paytext2)) & " Bil"
                'cmbSelectedCutOff.ItemData(cmbSelectedCutOff.NewIndex) = .Fields(0)
                prevCutOff = Format(!StartPayDate, "mm/dd/yyyy") & "-" & Format(!EndPayDate, "mm/dd/yyyy") & ":" & LTrim(RTrim(!Paytext2)) & " Bil"
            End If
            .MoveNext
        Wend
        If cmbSelectedCutOff.ListIndex = -1 Then
            cmbSelectedCutOff.ListIndex = 0
        End If
    End With
    rsBilledTrans.Filter = ""
End If
End Sub

Private Sub cmdAddCutoff_Click()
Dim mCurrIndex As Integer

    mCurrIndex = cmbClient.ListIndex
    Call frmBillingCutoff.Show(1)
    cmbClient.ListIndex = mCurrIndex
    
End Sub

Private Sub cmdClearTrans_Click()
Dim res As Variant

res = MsgBox("Clear all records from view?", vbQuestion + vbYesNo, PayrollAlert)
If res = vbYes Then      ' clear records
    ListView1(0).ListItems.Clear
    txtGrossBill.Text = "0.00"
    lblEntries(7).Caption = "Total Records: 0"
End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDeleteRec_Click()
Dim res As Variant
Dim mEmpNo As String
Dim mSOANo As String
Dim mClientCode As Integer

If Me.ListView1(0).ListItems.Count = 0 Then
    MsgBox "No records selected. Please add transactions first.", vbInformation, PayrollAlert
    GoTo ExitRtn2
End If
mEmpNo = Me.txtEmpNo
mSOANo = Me.txtSOANo
mClientCode = Me.cmbClient.ItemData(Me.cmbClient.ListIndex)

res = MsgBox("Do you want to delete selected record?", vbYesNo, PayrollAlert)
If res = vbYes Then
    ' Delete the record.
    If DeleteBilledTransTable(mEmpNo, mSOANo, mClientCode) = True Then
        MsgBox "Successfully deleted " & Me.ListView1(0).SelectedItem.ListSubItems(2).Text & ". Please Recompute Transactions to refresh data.", vbInformation, PayrollAlert
    Else
        MsgBox "Cannot delete current record.", vbInformation, PayrollAlert
    End If
End If

ExitRtn2:

End Sub

Private Function DeleteBilledTransTable(ByVal mEmpNo As String, ByVal mSOANo As String, _
    ByVal mClientCode As Integer) As Boolean

strQuery = "DELETE FROM BilledTrans WHERE ControlNo = '" & mEmpNo & "' and SOANo = '" & mSOANo & "' and Clientcode = " & mClientCode & " and dayspresent = " & DaysPresentKey
If Action_Query(strQuery, cnnStr) = True Then
    DeleteBilledTransTable = True
Else
    DeleteBilledTransTable = False
End If

End Function


Private Sub cmdPrintSOA_Click()
Dim AmountInWords As String
Dim res As Variant
Dim ClientIndex As Integer
Dim cmbctr As Integer
Dim OrderBy As String
Dim SelectString As String

'On Error GoTo ExitRtn2
If Mid(BillingTemplate, 1, 7) = "Carlson" Then
    'OrderBy = " ORDER BY Employees, StartPayDate, BRName"
    'OrderBy = " ORDER BY CutOffOrder, AreaCode, Employees" ' original
    OrderBy = " ORDER BY CutOffOrder, Employees, AreaCode"
    strQuery = "SELECT MIN(StartPayDate) FROM qryBilledTrans WHERE SOANo = '" & txtSOANo.Text & "'"
    Set rsBilledTrans2 = ReturnRst(strQuery, cnnStr, mstatic)
    FirstCutoff1 = Format(rsBilledTrans2.GetString, "mm/dd/yyyy")
    
    strQuery = "SELECT MAX(StartPayDate) FROM qryBilledTrans WHERE SOANo = '" & txtSOANo.Text & "'"
    Set rsBilledTrans2 = ReturnRst(strQuery, cnnStr, mstatic)
    FirstCutoff2 = Format(rsBilledTrans2.GetString, "mm/dd/yyyy")
    
'    strQuery = "DELETE FROM tmpCarlson"
'    Call Action_Query(strQuery, cnnStr)
    strQuery = "SELECT * FROM qryBilledEmployees WHERE SOANO = '" & txtSOANo.Text & "'"
    Set rsEmpBranch = ReturnRst(strQuery, cnnStr, mstatic)
    
    SelectString = "SELECT a.Employees, b.*, '" & FirstCutoff1 & "' AS CutoffOrder FROM qryBilledEmployees a LEFT JOIN (SELECT * FROM qryBilledTrans WHERE SOANO = '" & txtSOANo.Text & _
        "' AND StartPayDate = '" & FirstCutoff1 & _
        "') b ON LTRIM(RTRIM(a.Employees)) = LTRIM(RTRIM(b.EmployeeName )) WHERE a.SOANO = '" & txtSOANo.Text & "' UNION " & _
        "SELECT a.Employees, b.*, '" & FirstCutoff2 & "' AS CutoffOrder FROM qryBilledEmployees a LEFT JOIN (SELECT * FROM qryBilledTrans WHERE SOANO = '" & txtSOANo.Text & _
        "' AND StartPayDate = '" & FirstCutoff2 & _
        "') b ON LTRIM(RTRIM(a.Employees)) = LTRIM(RTRIM(b.EmployeeName )) WHERE a.SOANO = '" & txtSOANo.Text & "' "
    strQuery = SelectString & OrderBy
    Set rsBilledTrans2 = ReturnRst(strQuery, cnnStr, mstatic)
Else
    Call RetrieveTransactions(False)
End If

If rsBilledTrans2.RecordCount = 0 Then
    MsgBox "No records to process. Please recompute Transaction Records again.", vbInformation, PayrollAlert
    'rsEmpBranch.Close
    Exit Sub
End If

cmbctr = 0
ClientIndex = cmbClient.ListIndex   ' Save current client
AmountInWords = oCommonFunction.Word(Val(txtGrossBill.Text))

If GenerateSOAReport(AmountInWords) = False Then
    'rsEmpBranch.Close
    Exit Sub
End If

If chkExcludedRecords = False Or chkNewBilling = True Then
res = MsgBox("Do you want to mark these selected cutoff as billed and save SOA#?", vbYesNo, PayrollAlert)
If res = vbYes Then
'    strQuery = "UPDATE CutOfftable SET SOANo = '" & txtSOANo.Text & "', PaidFlag = 1, Selectedflag = 1, PaidDate = '" & _
'        Format(Now, "mm/dd/yyyy") & "' WHERE ClientCode = " & Val(txtCutOffBase.Text) & " and StartCutOffText = '" & _
'        Format(txtCutoffDate(2), "yyyymmdd") & "' and EndCutoffText = '" & _
'        Format(txtCutoffDate(3), "yyyymmdd") & "' and Format(ProcessDate, ""mmddyyyy"") = '" & Mid(cmbSelectedCutOff.Text, 23, 8) & "'"
            
    strQuery = "UPDATE CutOffTable a, (SELECT StartPayDate, EndPayDate, SOANO, PayText2 FROM BilledTrans WHERE ClientCode = " & mClientCode & " AND SOANo = '" & txtSOANo.Text & "') AS b SET a.SOANo = '" & txtSOANo.Text & "', a.PaidFlag = 1, a.Selectedflag = 1, a.PaidDate = '" & _
        Format(Now, "mm/dd/yyyy") & "' " & _
        " WHERE a.ClientCode = " & Val(txtCutOffBase.Text) & " and a.StartCutOffText = Format(b.StartPayDate, ""yyyymmdd"") and a.EndCutoffText = Format(b.EndPayDate, ""yyyymmdd"") and Format(a.ProcessDate, ""mmddyyyy"") = b.Paytext2 "
    Call Action_Query(strQuery, cnnStr)
           
    strQuery = "UPDATE ClientTable a SET a.SOANoCode = '" & txtSOAMonth.Text & "', a.SOANo = " & Val(txtSOASeq.Text) & _
        " WHERE a.ClientCode = " & Val(txtCutOffBase.Text)
    Call Action_Query(strQuery, cnnStr)
           
'    strQuery = "UPDATE a SET a.SOANo = '" & txtSOANo.Text & "', a.PaidFlag = 1, a.Selectedflag = 1, a.PaidDate = '" & _
'        Format(Now, "mm/dd/yyyy") & "' FROM CutOffTable a INNER JOIN BilledTrans b " & _
'        " ON a.ClientCode = b.ClientCode AND a.StartCutoffText = b.StartPayDate " & _
'        " and Format(a.ProcessDate, ""mmddyyyy"") = b.PayText2 " & _
'        " WHERE b.ClientCode = " & Val(txtCutOffBase.Text) & " and b.SOANo = '" & txtSOANo.Text & "'"

'    ' Update SOA Number.
'    strQuery = "UPDATE ClientBillTable SET SOANo = SOANo, SOANoCode = '" & Left(txtSOANo.Text, Len(txtSOANo.Text) - 4) & "'" & _
'        " WHERE ClientCode = " & Val(txtCutOffBase.Text)
'    Call Action_Query(strQuery, cnnStr)
    Call GetClientList
    cmbClient.ListIndex = ClientIndex
    MsgBox "Selected cutoff was successfully mark as billed with SOA# " & txtSOANo.Text, vbInformation, PayrollAlert
Else
    MsgBox "Selected cutoff was not mark as billed.", vbInformation, PayrollAlert
End If

End If

    'rsEmpBranch.Close

Exit Sub

ExitRtn2:
    
    MsgBox "No records to process. Please recompute Transaction Records again.", vbInformation, PayrollAlert
    Err.Clear

End Sub

Private Function GenerateSOAReport(ByVal AmtinWords As String) As Boolean
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
Dim FSO As New FileSystemObject
Dim LodgingAmt As Double
Dim PayrollBilling As Double
Dim PayrollBillHead As Double

On Error GoTo ExitRtn2

GenerateSOAReport = False
If rsBilledTrans2.RecordCount = 0 Then
    MsgBox "No records to process. Please recompute Transaction Records again.", vbInformation, PayrollAlert
    Exit Function
End If

'If Left(cmbClient.Text, 5) = "JAPAN" Then
'    ExcelTemplate = App.Path & "\Reports\Japan Home SOA Template.xls"
'Else
'    ExcelTemplate = App.Path & "\Reports\Other Clients SOA Template.xls"
'End If
ExcelTemplate = App.Path & "\Billing Templates\" & BillingTemplate

If FSO.FileExists(ExcelTemplate) = False Then
    MsgBox "Billing template " & ExcelTemplate & " does not exists!" & Chr(13) & _
        "Please check it first, cannot proceed.", vbInformation, PayrollAlert
    GoTo ExitRtn2
End If

Set xlApp1 = New Excel.Application
'Set xlBook1 = xlApp1.Workbooks.Open(ExcelTemplate)
Set xlBook1 = xlApp1.Workbooks.Add(ExcelTemplate)

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

'Set xlSheet = book1.Worksheets.Add
Set xlSheet = xlBook1.Worksheets("Data Source")
'Set xlSheet.Name = "PayrollBankCopy"

cPos = 5
ctr = 7
'On Error GoTo Errh
' Print Header
'xlSheet.Range("A2") = mSapi  '" (SAPI)"

'xlSheet.Range("A2") = ReportHeader & " " & IIf(chkLoans.VALUE = 1, "(for Loans)", "") 'mSapi

'xlSheet.Range("A3") = ScopeRep & " " & cmbMonthof & " " & cmbYearof

ctr = ctr + 1
i = 1
j = 5 '18
MemberCtr = 1

If Mid(BillingTemplate, 1, 7) <> "Carlson" Then
    With xlSheet
        .Range("A2") = " PESOS: " & AmtinWords & " "
        .Range("A4") = Format(txtCutoffDate(2), "mmm. dd") & " - " & Format(txtCutoffDate(3), "mmmm dd, yyyy") & "."
        .Range("A6") = mskStartContract.Text
    '    .Range("A7") = Format(gGrossBill, "0.00") ' total employees
        .Range("A8") = Format(txtGrossBill.Text, "0.00")
        .Range("A10") = txtSOANo.Text
        '.Range("A10") = Format(txtCutOffDate(2), "yyyy")
        .Range("A12") = "SOA Control #:" & txtSOANo.Text
        .Range("A13") = "Billing Period: " & .Range("A4")
        
        .Range("A15") = UCase(cmbClient.Text) ' Client Name
        .Range("A16") = UCase(ClientAddress) ' Client Address
    
        .Range("A18") = UCase(Replace(ClientContact1, Chr(13), "")) ' Contact 1
        .Range("A19") = "(" & ClientContactPos1 & ")" ' Contact position 1
    
        .Range("A22") = "NOTE:  Please  make all checks payable to " & UCase(CompanyName)
        
        .Range("A23") = UCase(CompanyName)
        .Range("A24") = UCase(cmbClient.Text)
    End With
End If

'Set xlSheet = xlBook1.Worksheets("Billing Report")
Call PopulateBillingRecords(xlApp1, xlSheet, xlBook1)

If Mid(BillingTemplate, 1, 7) = "Carlson" Then
    Set xlSheet = xlBook1.Worksheets("Data Source")
    With xlSheet
        .Range("A1") = UCase(CompanyName)
        ' First Cut-Off
        .Range("A2") = "Billing Period: " & Format(FirstCutoff1, "mmm. dd") & " - " & Format(FirstCutoff2, "mmmm dd, yyyy")
        .Range("A3") = "SOA Control #:" & txtSOANo.Text
        .Range("A4") = "Payroll Summary Report for Cut-Off " & Format(FirstCutoff1, "mmm. dd") & " - " & Format(FirstCutoff2, "mmmm dd, yyyy")
        .Range("A5") = Format(FirstCutoff1, "yyyy")
        .Range("B5") = "Billing Period: " & Format(FirstCutoff1, "mmm. dd") & " - " & Format(FirstCutoff2, "mmmm dd, yyyy")
        .Range("A6") = UCase(cmbClient.Text) & "(" & CompanyShort & ")"
        ' Second Cut-off
        .Range("A7") = "Billing Period: " & Format(SecondCutoff1, "mmm. dd") & " - " & Format(SecondCutoff2, "mmmm dd, yyyy")
        .Range("A8") = "SOA Control #:" & txtSOANo.Text
        .Range("A9") = "Payroll Summary Report for Cut-Off " & Format(SecondCutoff1, "mmm. dd") & " - " & Format(SecondCutoff2, "mmmm dd, yyyy")
        .Range("A10") = Format(SecondCutoff1, "yyyy")
        .Range("B10") = "Billing Period: " & Format(SecondCutoff1, "mmm. dd") & " - " & Format(SecondCutoff2, "mmmm dd, yyyy")
        
        .Range("A7") = "Billing Period: " & Format(SecondCutoff1, "mmm. dd") & " - " & Format(SecondCutoff2, "mmmm dd, yyyy")
        .Range("B5") = Format(SecondCutoff1, "yyyy")
        
        .Range("A12") = Format(FirstCutoff1, "yyyy")
        .Range("B12") = "ALLOWANCES & LODGING, INCENTIVE BILLING - " & Format(FirstCutoff1, "mmmm yyyy")
        .Range("A13") = Format(FirstCutoff1, "mm/dd/yyyy")
        
        LodgingAmt = Format(Val(.Range("A14")), "0.00") ' Board and lodging allow
        AmtinWords = oCommonFunction.Word(LodgingAmt)
        .Range("A15") = "     This is to bill you the amount in PESOS: "
        .Range("A16") = AmtinWords & " (PHP " & LodgingAmt & ")"
        .Range("A17") = "representing BOARD & LODGING ALLOWANCES, BONUS of our field staffers assigned to your company for the month as follows: "
        .Range("A20") = txtSOANo.Text
        
        .Range("A22") = Format(FirstCutoff1, "yyyy")
        .Range("B22") = "Billing for " & Format(FirstCutoff1, "yyyy") & " (HEAD OFFICE) - " & Format(FirstCutoff1, "mmmm yyyy")
        .Range("A23") = Format(FirstCutoff1, "mm/dd/yyyy")
        '.Range("A24") = PayrollBillHead
        PayrollBillHead = Format(Val(.Range("A24")), "0.00") ' Payroll Billing Head
        AmtinWords = oCommonFunction.Word(PayrollBillHead)
        .Range("A25") = "     This is to bill you the amount in PESOS: "
        .Range("A26") = AmtinWords   '& " (PHP " & PayrollBillHead & ")"
        .Range("A27") = " representing payroll for the services rendered by our field staffers for the period  for the MONTH of "
        .Range("A28") = Format(FirstCutoff1, "mmmm yyyy")
        
        .Range("A30") = Format(FirstCutoff1, "yyyy")
        .Range("B30") = "Billing for " & Format(FirstCutoff1, "yyyy") & " (PROMODISERS) - " & Format(FirstCutoff1, "mmm yyyy")
        .Range("A31") = Format(FirstCutoff1, "mm/dd/yyyy")
        PayrollBilling = Format(Val(.Range("A32")), "0.00") ' Payroll Billing Head
        AmtinWords = oCommonFunction.Word(PayrollBilling)
        .Range("A33") = "     This is to bill you the amount in PESOS: "
        .Range("A34") = UCase(AmtinWords) ' & " (PHP " & PayrollBilling & ")")
        .Range("A35") = " representing payroll for the services rendered by our field staffers for the period  for the MONTH of "
        .Range("A36") = Format(FirstCutoff1, "mmmm yyyy")
        
        .Range("B45") = Format(FirstCutoff1, "mmmm dd") & " - " & Format(FirstCutoff2, "dd, yyyy")
        .Range("B46") = Format(SecondCutoff1, "mmmm dd") & " - " & Format(SecondCutoff2, "dd, yyyy")
        
        .Range("B49") = Format(FirstCutoff1, "mmmm dd") & " - " & Format(FirstCutoff2, "dd, yyyy")
        .Range("B50") = Format(SecondCutoff1, "mmmm dd") & " - " & Format(SecondCutoff2, "dd, yyyy")
        
'        .Range("A4") = Format(txtCutoffDate(2), "mmm. dd") & " - " & Format(txtCutoffDate(3), "mmmm dd, yyyy") & "."
'        .Range("A6") = mskStartContract.Text
'    '    .Range("A7") = Format(gGrossBill, "0.00") ' total employees
'        .Range("A8") = Format(txtGrossBill.Text, "0.00")
'        '.Range("A10") = Format(txtCutOffDate(2), "yyyy")
'
'        .Range("A15") = UCase(cmbClient.Text) ' Client Name
'        .Range("A16") = UCase(ClientAddress) ' Client Address
'
'        .Range("A18") = UCase(Replace(ClientContact1, Chr(13), "")) ' Contact 1
'        .Range("A19") = "(" & ClientContactPos1 & ")" ' Contact position 1
'
'        .Range("A22") = "NOTE:  Please  make all checks payable to " & UCase(CompanyName)
'
'        .Range("A23") = UCase(CompanyName)
'        .Range("A24") = UCase(cmbClient.Text)
    End With
End If


'Set xlSheet = xlBook1.Worksheets("Payroll Summary")
'Call PayrollSummaryReport(xlApp1, xlSheet)

Set xlSheet = xlBook1.Worksheets("SOA")
xlApp1.Sheets("SOA").Select
mRange = "A1"
xlSheet.Range(mRange).Select

GenerateSOAReport = True

Exit Function

ExitRtn2:

MsgBox Err.Description
Err.Clear

End Function

'Private Sub PayrollSummaryReport(ByRef xlApp1 As Excel.Application, ByRef xlSheet As Excel.Worksheet)
'strQuery = "SELECT a.*, b.acctno " & _
'    "FROM PayrollMaster AS a LEFT JOIN hremployees AS b ON a.EmployeeNum = b.EmployeeNum " & _
'    " WHERE ClientCode = " & mClientCode & " and StartPayDate = '" & " "
'
'Set rsPayTrans = ReturnRst(strQuery, cnnStr, mstatic)
'
'End Sub

Private Sub cmdRefresh_Click()
    Call GetBilledSOA
    'Call chkNewBilling_Click
End Sub

Private Sub cmdRetrieveTrans_Click()
    Call RetrieveTransactions(False)
    Call GetSelectedCutoff(False)
End Sub

Private Sub cmdShowCutoff_Click()
    Call RetrieveTransactions(True)
End Sub

Private Sub cmdUpdateTemplate_Click()
Call ShowTemplateForm
End Sub

Private Sub Form_Activate()
If chkSuccess.VALUE = 1 Then
    ' Refresh the data.
    Call RetrieveTransactions(False)
    Call GetClientList
    cmbClient.Refresh
End If
End Sub

'Private Sub Form_Activate()
'    ' Get Company Info.
'    Call CompanyInfo
'    ' Get Client List.
'    Set rsClientTable = New ADODB.Recordset
' '  strQuery = "SELECT clientcode, clientname, clientcutoff, clientcutoffdates, clientflag1 FROM clienttable WHERE clientcode <> 0"
'    strQuery = "SELECT a.*, b.SOANo FROM clienttable a INNER JOIN ClientBillTable b On a.ClientCode = b.ClientCode WHERE a.clientcode <> 0 AND RTRIM(a.CLIENTNAME) <> '' ORDER BY a.clientname"
'    Set rsClientTable = ReturnRst(strQuery, cnnStr, mstatic)
'    Call oCommonFunction.FillCombo(cmbClient, rsClientTable)
'    If rsClientTable.RecordCount <> 0 Then cmbClient.ListIndex = 0
'
'    mskStartContract.Text = Format(Now, "mm/dd/yyyy")
'
'End Sub

Private Sub Form_Load()
Dim strQuery As String
    
 '  Center screen
 '  Move (mdiMain.Width - Width) \ 2, (mdiMain.Height - Height) \ 2
    Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    'txtDateToday = Format(Now, "mm/dd/yyyy")
    cmdrefresh.Visible = False
    
    ' Get Company Info.
    Call CompanyInfo
    ' Get Client List.
    Set rsClientTable = New ADODB.Recordset
 '  strQuery = "SELECT clientcode, clientname, clientcutoff, clientcutoffdates, clientflag1 FROM clienttable WHERE clientcode <> 0"
    Call GetClientList
'    strQuery = "SELECT a.*, b.SOANoCode, b.SOANo FROM clienttable a INNER JOIN ClientBillTable b On a.ClientCode = b.ClientCode WHERE a.clientcode <> 0 AND RTRIM(a.CLIENTNAME) <> '' ORDER BY a.clientname"
'    Set rsClientTable = ReturnRst(strQuery, cnnStr, mstatic)
'    Call oCommonFunction.FillCombo(cmbClient, rsClientTable)
'    If rsClientTable.RecordCount <> 0 Then cmbClient.ListIndex = 0

    mskStartContract.Text = Format(Now, "mm/dd/yyyy")
    
    ' Get all billed Cut-off.
    strQuery = "SELECT * FROM qryBilledSOA"
    Set rsBilledTrans = ReturnRst(strQuery, cnnStr, mstatic)
    cmbMonth.Text = Format(Now, "mm") & " - " & Format(Now, "mmmm")
    cmbYear.Text = Format(Now, "yyyy")
End Sub

Private Sub GetClientList()

    'strQuery = "SELECT a.*, b.SOANoCode, b.SOANo FROM clientTable a INNER JOIN CutoffTable b On a.ClientCode = b.ClientCode WHERE a.clientcode <> 0 AND RTRIM(a.CLIENTNAME) <> '' ORDER BY a.clientname"
    strQuery = "SELECT a.* FROM clientTable a WHERE a.clientcode <> 0 and ClientStatus = True AND RTRIM(a.CLIENTNAME) <> '' ORDER BY a.clientname"
    Set rsClientTable = ReturnRst(strQuery, cnnStr, mstatic)
    Call oCommonFunction.FillCombo(cmbClient, rsClientTable)
    If rsClientTable.RecordCount <> 0 Then cmbClient.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsCompanyTable = Nothing
Set rsClientTable = Nothing
If frmBillingUpdate.Visible = True Then
    Unload frmBillingUpdate
End If
End Sub

Private Sub ListView1_DblClick(Index As Integer)
    Call ShowTemplateForm
End Sub

Private Sub ShowTemplateForm()

chkSuccess.VALUE = 0
If ListView1(0).ListItems.Count <> 0 Then
    With ListView1(0)
    txtEmpNo.Text = .SelectedItem.ListSubItems(1).Text ' save employee number
    'frmBillingUpdate.Label1(14).Caption = "Edit Employee: " & .SelectedItem.ListSubItems(2).Text ' display name
    frmBillingUpdate.Frame1.Caption = "Employee Name: " & .SelectedItem.ListSubItems(2).Text ' display name
    frmBillingUpdate.txtEditFld(0).Text = .SelectedItem.ListSubItems(5).Text
    frmBillingUpdate.txtEditFld(1).Text = .SelectedItem.ListSubItems(7).Text
    frmBillingUpdate.txtEditFld(2).Text = .SelectedItem.ListSubItems(9).Text
    frmBillingUpdate.txtEditFld(3).Text = .SelectedItem.ListSubItems(11).Text
    frmBillingUpdate.txtEditFld(4).Text = .SelectedItem.ListSubItems(13).Text
    frmBillingUpdate.txtEditFld(5).Text = .SelectedItem.ListSubItems(15).Text
    frmBillingUpdate.txtEditFld(6).Text = .SelectedItem.ListSubItems(17).Text
    frmBillingUpdate.txtEditFld(7).Text = .SelectedItem.ListSubItems(19).Text
    frmBillingUpdate.txtEditFld(8).Text = .SelectedItem.ListSubItems(21).Text
    frmBillingUpdate.txtEditFld(9).Text = .SelectedItem.ListSubItems(24).Text
    frmBillingUpdate.txtEditFld(10).Text = .SelectedItem.ListSubItems(22).Text
    frmBillingUpdate.txtEditFld(15).Text = .SelectedItem.ListSubItems(26).Text
    frmBillingUpdate.txtEditFld(11).Text = .SelectedItem.ListSubItems(28).Text
    frmBillingUpdate.txtEditFld(14).Text = .SelectedItem.ListSubItems(29).Text
    frmBillingUpdate.txtEditFld(12).Text = .SelectedItem.ListSubItems(44).Text
    frmBillingUpdate.txtEditFld(13).Text = .SelectedItem.ListSubItems(55).Text
    
    frmBillingUpdate.txtEditFld(18).Text = Format(OverTimeRate, "0.00") ' ot rate
    frmBillingUpdate.txtEditFld(17).Text = Format(RestDayWorkedOTRate, "0.00") ' rest day rate
    frmBillingUpdate.txtEditFld(16).Text = Format(SpecialWorkedDayOffRate, "0.00") 'spl hol rate
    frmBillingUpdate.txtEditFld(21).Text = Format(LegalHolidayWithPremiumRate, "0.00") ' ot w/ premium
    frmBillingUpdate.txtEditFld(20).Text = Format(LegalHolidayWorkedDayoffRate, "0.00") ' legal holiday
    frmBillingUpdate.txtEditFld(19).Text = Format(ExcessSplOTRate, "0.00") ' excess OT spl rate
    frmBillingUpdate.txtEditFld(24).Text = Format(ExcessDayoffSplOTRate, "0.00") ' legal holiday rate
    frmBillingUpdate.txtEditFld(23).Text = Format(ExcessLglHolOTRate, "0.00") ' excess legal rate
    frmBillingUpdate.txtEditFld(22).Text = Format(ExcessSunLglOTRate, "0.00") ' excess sun legal rate
    frmBillingUpdate.txtEditFld(27).Text = Format(NightDiffRate, "0.00") ' night diff
    frmBillingUpdate.txtEditFld(26).Text = Format(AgencyRate, "0.00") ' client agency rate
    frmBillingUpdate.txtEditFld(25).Text = Format(VatPercentRate, "0.00") ' Vat rate
    
    frmBillingUpdate.txtEditFld(28).Text = BillingTemplate ' Billing Template
    
    End With

    frmBillingUpdate.txtClientCode.Text = ""
    Call frmBillingUpdate.Show
    frmBillingUpdate.SetFocus
Else
    MsgBox "No records, please select payroll transactions first.", vbInformation, PayrollAlert
End If

End Sub

Private Sub ListView1_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
Dim mControlno As String
Dim mName As String

mControlno = ListView1(Index).SelectedItem.ListSubItems(1).Text
mName = ListView1(Index).SelectedItem.ListSubItems(2).Text

If ListView1(0).SelectedItem.Checked = True Then
    strQuery = "UPDATE BilledTrans SET ExcludedFlag = 1 " & _
    " WHERE ClientCode = " & mClientCode & " and Employeename = '" & mName & "' AND " & _
    " ControlNo = '" & mControlno & "' AND " & _
    " SOANO = '" & txtSOANo.Text & "'"
Else
    strQuery = "UPDATE BilledTrans SET ExcludedFlag = 0 " & _
    " WHERE ClientCode = " & mClientCode & " and Employeename = '" & mName & "' AND " & _
    " ControlNo = '" & mControlno & "' AND " & _
    " SOANO = '" & txtSOANo.Text & "'"
End If
Call Action_Query(strQuery, cnnStr)
'Call RetrieveTransactions(False)

End Sub

'Private Sub ListView1_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
'
'With ListView1(0)
'    frmBillingUpdate.txtEditFld(0).Text = .SelectedItem.ListSubItems(5).Text
'    frmBillingUpdate.txtEditFld(1).Text = .SelectedItem.ListSubItems(7).Text
'
'End With
'
'End Sub

'Private Sub txtCutoffDate_GotFocus(Index As Integer)
'SendKeys "{HOME}+{END}"
'End Sub

Private Sub txtCutoffDate_GotFocus(Index As Integer)
txtCutoffDate(Index).BackColor = &HFFFFC0
'Call txtHLigths(txtCutoffDate(Index))
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCutOffDate_LostFocus(Index As Integer)
txtCutoffDate(Index).BackColor = vbWhite
End Sub

Private Sub txtCutoffDate_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Public Sub CompanyInfo()
    
    strQuery = "SELECT * FROM companytable WHERE agencycode = true"
    Set rsCompanyTable = ReturnRst(strQuery, CnnstrHR, mstatic)

End Sub

'Private Function PayrollMasterRec() As Boolean
'
'strQuery = "SELECT * FROM PayrollMaster WHERE clientcode = " & mClientCode & _
'        " and startpaydate = '" & txtCutoffDate(2) & "' and endpaydate = '" & txtCutoffDate(3) & "'"
'Set rsPayrollTrans = ReturnRst(strQuery, cnnStr, mstatic)
'
'End Function

'Private Sub DisplayTransactions()
'Dim litem1 As ListItem
'Dim Listctr As Integer
'Dim GrossBill As Double
'Dim VatRate As Double
'
'gGrossBill = 0#
'GrossBill = 0#
'VatRate = 0#
'Listctr = 0
'
'strQuery = "SELECT a.*, b.DailyRate, c.VATRate, c.ClientChargePerDay FROM ((PayrollMaster a LEFT JOIN HREmployees b ON a.EmployeeNum = b.EmployeeNum) " & _
'        " LEFT JOIN ClientBillTable c ON a.ClientCode = c.ClientCode) WHERE a.clientcode = " & mClientCode & _
'        " and a.startpaydate = '" & txtCutoffDate(2) & "' and a.endpaydate = '" & txtCutoffDate(3) & "'"
'Set rsPayrollTrans = ReturnRst(strQuery, cnnStr, mstatic)
'If rsPayrollTrans.RecordCount <= 0 Then
'    MsgBox "No payroll transaction records selected.", vbInformation, PayrollAlert
'    rsPayrollTrans.Close
'    GoTo ExitRtn1
'End If
'
'With rsPayrollTrans
'ListView1(0).ListItems.Clear
'While Not (.EOF Or .BOF)
'    Set litem1 = ListView1(0).ListItems.Add(, , !EmployeeNum)
'        litem1.SubItems(1) = IIf(IsNull(!ControlNo), "", !ControlNo)
'        litem1.SubItems(2) = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
'        litem1.SubItems(3) = Space(5) 'DTR
'        litem1.SubItems(4) = Format(!DaysPresent, "0")
'        litem1.SubItems(5) = Format(!RestDayOTHRs, "0") ' extra days
'        litem1.SubItems(6) = Format(!NoSplHoliday, "0") ' extra days
'        litem1.SubItems(7) = Format(!SundayOT, "0")     ' extra days
'        litem1.SubItems(8) = Format(!NoLglHoliday, "0") ' extra days
'        litem1.SubItems(9) = Format(!DaysPresent, "0")  ' # of days worked
'        litem1.SubItems(10) = Format(!HoursTardiness, "0.00")
'        litem1.SubItems(11) = Format(!OTHours, "0.00")
'        litem1.SubItems(12) = Format(!excesssplholothours, "0.00")
'        litem1.SubItems(13) = "-"
'        litem1.SubItems(14) = Format(!LglHolidayOTHours, "0.00")
'        litem1.SubItems(15) = Format(!ExcessOTSunLegalHrs, "0.00")
'        litem1.SubItems(16) = Format(!NightDiffOT, "0.00")
'        litem1.SubItems(17) = Format(!DailyRate, "0.00")    ' Daily Rate
'        litem1.SubItems(18) = Format(!BasicPay, "0.00")     ' Basic Pay
'        litem1.SubItems(19) = Format(!RestDayOTPay, "0.00")     ' Premium pay on extra days worked
'        litem1.SubItems(20) = Format(!SplHolidayPay, "0.00")    ' Premium pay on extra days worked
'        litem1.SubItems(21) = Format(!SundayOTPay, "0.00")      ' premium pay on extra days worked
'        litem1.SubItems(22) = Format(!LglHolidayPay, "0.00")    ' premium pay on extra days worked
'        litem1.SubItems(23) = Format(!TardinessDeduction, "0.00")
'        litem1.SubItems(24) = Format(!OTPay, "0.00")            ' ot
'        litem1.SubItems(25) = Format(!excesssplholotpay, "0.00") ' ot
'        litem1.SubItems(26) = "-"
'        litem1.SubItems(27) = Format(!LglHolidayOTPay, "0.00")  ' ot
'        litem1.SubItems(28) = Format(!excessotsunlegalpay, "0.00")  ' ot
'        litem1.SubItems(29) = Format(!NightDiffPay, "0.00") ' ot
'
'        litem1.SubItems(30) = Format(!EcolaPay, "0.00")
'        litem1.SubItems(31) = Format(!SalaryAdjust, "0.00")
'        litem1.SubItems(32) = Format(!GrossPay, "0.00")
'        litem1.SubItems(33) = Format(!Month13Payment, "0.00")
'
'        litem1.SubItems(34) = Format(!SSSEmployer + !ECCemployer, "0.00")
'        litem1.SubItems(35) = Format(!PHealthEmployer, "0.00")
'        litem1.SubItems(36) = Format(!PagibigEmployer, "0.00")
'        litem1.SubItems(37) = Format(!GrossPay + !SSSEmployer + !ECCemployer + !PHealthEmployer + !PagibigEmployer, "0.00") ' Total pay and benefits
'        litem1.SubItems(38) = Format(!DaysPresent * !clientchargeperday, "0.00") ' Agency Fee
'        GrossBill = !GrossPay + !SSSEmployer + !ECCemployer + !PHealthEmployer + !PagibigEmployer + (!DaysPresent * !clientchargeperday)
'        gGrossBill = gGrossBill + GrossBill
'        'litem1.SubItems(39) = Format(!Grosspay + !SSSemployer + !ECCemployer + !phealthemployer + !pagibigemployer + (!DaysPresent * !ClientChargePerDay), "0.00") ' Gross Billing
'        litem1.SubItems(39) = Format(GrossBill, "0.00") ' Gross Billing
'        VatRate = (GrossBill * !VatRate)
'        litem1.SubItems(40) = Format(VatRate, "0.00") 'Vat 12%
'        litem1.SubItems(41) = Format(!IncentivePay, "0.00") 'Gross Payroll
'        litem1.SubItems(42) = Format(GrossBill + !IncentivePay, "0.00") ' Total Amount Due
'        litem1.SubItems(43) = ""
'        litem1.SubItems(44) = Format(!GrossPay, "0.00") 'Gross Payroll
'        litem1.SubItems(45) = "-"   ' Ignore Lates
'        litem1.SubItems(46) = Format(!GrossPay - !GrossPay, "0.00") ' Diff'l
'
'    Listctr = Listctr + 1
'    .MoveNext
'Wend
'
'End With
'
'frmBilling.txtGrossBill.Text = Format(gGrossBill, "0.00")
'frmBilling.lblEntries(7).Caption = "Included Records: " & Listctr
'
'ExitRtn1:
'
'End Sub

Private Sub RetrieveTransactions(ByVal FilterFlag As Boolean)   ' Retrieve Transactions using SOA no
Dim litem1 As ListItem
Dim PHNo As String
Dim ActualEE As Double
Dim ActualER As Double
Dim AdjustedEE As Double, TableBasedEE As Double
Dim AdjustedER As Double, TableBasedER As Double
Dim AdjustedECC As Double
Dim mSurname As String
Dim mFirstName As String
Dim mMiddlename As String
Dim mBasicPay As Double
Dim mBracket As String
Dim mDiff As Double
Dim ExcludedFlag As Integer
Dim mListView As ListView
Dim i As Integer
Dim mZeroValue As Double
Dim Listctr As Integer
Dim GrossBill As Double
Dim VATPayment As Double
Dim ProcDate As String
Dim OrderBy As String
Dim SelectString As String

ProcDate = Mid(cmbSelectedCutOff.Text, 23, 8)
mZeroValue = 0#
GrossBill = 0#
Set mListView = ListView1(0)
DaysPresentKey = 0

Listctr = 0
gGrossBill = 0#

If Mid(BillingTemplate, 1, 7) = "Carlson" Then
    'OrderBy = " ORDER BY Employees, StartPayDate, BRName"
    'OrderBy = " ORDER BY CutOffOrder, AreaCode, Employees"
    OrderBy = " ORDER BY CutOffOrder, Employees, AreaCode"
    strQuery = "SELECT MIN(StartPayDate) FROM qryBilledTrans WHERE SOANo = '" & txtSOANo.Text & "'"
    Set rsBilledTrans2 = ReturnRst(strQuery, cnnStr, mstatic)
    FirstCutoff1 = Format(rsBilledTrans2.GetString, "mm/dd/yyyy")
    
    strQuery = "SELECT MAX(StartPayDate) FROM qryBilledTrans WHERE SOANo = '" & txtSOANo.Text & "'"
    Set rsBilledTrans2 = ReturnRst(strQuery, cnnStr, mstatic)
    FirstCutoff2 = Format(rsBilledTrans2.GetString, "mm/dd/yyyy")
    
    SelectString = "SELECT a.Employees, b.*, '" & FirstCutoff1 & "' AS CutoffOrder FROM qryBilledEmployees a LEFT JOIN (SELECT * FROM qryBilledTrans WHERE SOANO = '" & txtSOANo.Text & _
        "' AND StartPayDate = '" & FirstCutoff1 & _
        "') b ON LTRIM(RTRIM(a.Employees)) = LTRIM(RTRIM(b.EmployeeName )) WHERE a.SOANO = '" & txtSOANo.Text & "' UNION " & _
        "SELECT a.Employees, b.*, '" & FirstCutoff2 & "' AS CutoffOrder FROM qryBilledEmployees a LEFT JOIN (SELECT * FROM qryBilledTrans WHERE SOANO = '" & txtSOANo.Text & _
        "' AND StartPayDate = '" & FirstCutoff2 & _
        "') b ON LTRIM(RTRIM(a.Employees)) = LTRIM(RTRIM(b.EmployeeName )) WHERE a.SOANO = '" & txtSOANo.Text & "' "
    'Filtering is not applicable to carlson.
    strQuery = SelectString & OrderBy
Else
    OrderBy = " ORDER BY CutOffNo, format(StartPayDate, ""yyyymmdd""), DailyRate, EmployeeName"
'    SelectString = "SELECT EmployeeName AS Employees, * FROM qryBilledTrans a WHERE ClientCode = " & mClientCode & " and ExcludedFlag = 0 " & _
'            " and SOANo = '" & txtSOANo.Text & "' "
    'Get selected cut-off.
    If chkExcludedRecords.VALUE = False Then
        If FilterFlag = True And cmbSelectedCutOff.Text <> "All Cutoff" Then ' include start and end date
            strQuery = "SELECT EmployeeName AS Employees, * FROM qryBilledTrans a WHERE ClientCode = " & mClientCode & _
                " and SOANo = '" & txtSOANo.Text & "' AND StartPayDate = '" & Format(Left(cmbSelectedCutOff.Text, 10), "mm/dd/yyyy") & "'" & _
                " and EndPayDate = '" & Format(Mid(cmbSelectedCutOff.Text, 12, 10), "mm/dd/yyyy") & _
                "' and Ltrim(Rtrim(PayText2)) = '" & ProcDate & "' " & OrderBy ' ORDER BY CutOffNo, format(StartPayDate, ""yyyymmdd""), DailyRate, EmployeeName"
        Else    'Get all cut-off for soa #.
            strQuery = "SELECT EmployeeName AS Employees, * FROM qryBilledTrans a WHERE ClientCode = " & mClientCode & _
                " and SOANo = '" & txtSOANo.Text & "' " & OrderBy ' ORDER BY CutOffNo, Format(StartPayDate, ""yyyymmdd""), DailyRate, EmployeeName"
        End If
    Else        ' Print Excluded Only
        If FilterFlag = True And cmbSelectedCutOff.Text <> "All Cutoff" Then ' include start and end date
            strQuery = "SELECT EmployeeName AS Employees, * FROM qryBilledTrans a WHERE ClientCode = " & mClientCode & " and ExcludedFlag = True " & _
                " and SOANo = '" & txtSOANo.Text & "' AND StartPayDate = '" & Format(Left(cmbSelectedCutOff.Text, 10), "mm/dd/yyyy") & "'" & _
                " and EndPayDate = '" & Format(Mid(cmbSelectedCutOff.Text, 12, 10), "mm/dd/yyyy") & _
                "' and Ltrim(Rtrim(PayText2)) = '" & ProcDate & "' " & OrderBy ' ORDER BY CutOffNo, format(StartPayDate, ""yyyymmdd""), DailyRate, EmployeeName"
        Else    'Get all cut-off for soa #.
            strQuery = "SELECT EmployeeName AS Employees, * FROM qryBilledTrans a WHERE ClientCode = " & mClientCode & " and ExcludedFlag = True " & _
                " and SOANo = '" & txtSOANo.Text & "' " & OrderBy ' ORDER BY CutOffNo, Format(StartPayDate, ""yyyymmdd""), DailyRate, EmployeeName"
        End If
    End If
End If

Set rsBilledTrans2 = ReturnRst(strQuery, cnnStr, mstatic)
ListView1(0).ListItems.Clear
If rsBilledTrans2.RecordCount <= 0 Then
    frmBilling.txtGrossBill.Text = 0#
    frmBilling.lblEntries(7).Caption = "Total Records: 0"
    MsgBox "No records selected.", vbInformation, PayrollAlert
    GoTo ExitRtn2
End If

With rsBilledTrans2
While Not (.EOF Or .BOF)
    Listctr = Listctr + 1
    
    Set litem1 = ListView1(0).ListItems.Add(, , IIf(IsNull(!EmployeeNum), "0000", !EmployeeNum))
        litem1.SubItems(1) = IIf(IsNull(!ControlNo), "", !ControlNo)
        'litem1.SubItems(2) = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
        litem1.SubItems(2) = IIf(IsNull(!Employees), "No Name", !Employees)
        litem1.SubItems(3) = Space(5) 'DTR
        litem1.SubItems(4) = Format(!DaysPresent, "0.00")
        DaysPresentKey = Format(IIf(IsNull(!DaysPresent), 0, !DaysPresent), "0.00")
        DaysPresent = Format(IIf(IsNull(!DaysPresent2), 0, !DaysPresent2), "0.00")
        litem1.SubItems(5) = Format(!DaysPresent2, "0.00")     ' edit days present
        litem1.SubItems(6) = Format(!RestDayOTHrs, "0.00")     ' extra days
        RestDayOTHrs = Format(IIf(IsNull(!RestDayOTHrs2), 0, !RestDayOTHrs2), "0.00")
        litem1.SubItems(7) = Format(!RestDayOTHrs2, "0.00")    ' edit extra days
        litem1.SubItems(8) = Format(!NoSplHoliday, "0.00")     ' no spl holiday
        NoSplHoliday = Format(IIf(IsNull(!NoSplHoliday2), 0, !NoSplHoliday2), "0.00")
        litem1.SubItems(9) = Format(!NoSplHoliday2, "0.00")    ' edit no spl holiday
        
        litem1.SubItems(10) = Format(!SundayOT, "0.00")        ' sunday ot
        SundayOT = Format(IIf(IsNull(!SundayOT2), 0, !SundayOT2), "0.00")
        litem1.SubItems(11) = Format(!SundayOT2, "0.00")       ' edit sunday ot
        litem1.SubItems(12) = Format(!NoLglHoliday, "0.00")    ' no lgl holiday
        NoLglHoliday = Format(IIf(IsNull(!NoLglHoliday2), 0, !NoLglHoliday2), "0.00")
        litem1.SubItems(13) = Format(!NoLglHoliday2, "0.00")   ' edit no lgl holiday
        litem1.SubItems(14) = Format(!DaysPresent, "0.00")     ' # of days worked
        DaysPresentKey = IIf(IsNull(!DaysPresent), 0, Format(!DaysPresent, "0.00"))     ' # of days worked
        
        DaysWorked = Format(DaysPresent + RestDayOTHrs + NoSplHoliday + NoLglHoliday, "0.00")
        litem1.SubItems(15) = Format(DaysWorked, "0.00")           ' # of days worked
        
        litem1.SubItems(16) = Format(!HoursTardiness, "0.00")   ' hours tardi
        HoursTardiness = IIf(IsNull(!HoursTardiness2), 0, !HoursTardiness2)
        litem1.SubItems(17) = Format(!HoursTardiness2, "0.00")  ' edit hours tardi
        litem1.SubItems(18) = Format(!OTHours, "0.00")
        OTHours = Format(IIf(IsNull(!OTHours2), 0, !OTHours2), "0.00")
        litem1.SubItems(19) = Format(!OTHours2, "0.00")
        litem1.SubItems(20) = Format(!Excesssplholothours, "0.00")
        ExcessSplOTHours = Format(IIf(IsNull(!ExcessSplHolOTHours2), 0#, !ExcessSplHolOTHours2), "0.00")
        litem1.SubItems(21) = Format(ExcessSplOTHours, "0.00")
        'LglHolidayPay
        litem1.SubItems(22) = 0#   'Format(IIf(IsNull(!LglHolidayPay), 0#, !LglHolidayPay), "0.00")
        litem1.SubItems(23) = Format(!LglHolidayOTHours, "0.00")
        LglHolidayOTHours = IIf(IsNull(!LglHolidayOTHours2), 0, !LglHolidayOTHours2)
        litem1.SubItems(24) = Format(IIf(IsNull(!LglHolidayOTHours2), 0#, !LglHolidayOTHours2), "0.00")
        litem1.SubItems(25) = Format(!ExcessOTSunLegalHrs, "0.00")
        ExcessOTSunLegalHrs = IIf(IsNull(!ExcessOTSunLegalHrs2), 0, !ExcessOTSunLegalHrs2)
        litem1.SubItems(26) = Format(IIf(IsNull(!ExcessOTSunLegalHrs2), 0#, !ExcessOTSunLegalHrs2), "0.00")
        litem1.SubItems(27) = Format(!NightDiffOT, "0.00")
        NightDiffOT = IIf(IsNull(!NightDiffOT2), 0, !NightDiffOT2)
        litem1.SubItems(28) = Format(!NightDiffOT2, "0.00")
        
        litem1.SubItems(29) = Format(!DailyRate, "0.00")    ' Daily Rate
        'litem1.SubItems(30) = Format(!BasicPay, "0.00")     ' Basic Pay
        BasicPay = DaysWorked * IIf(IsNull(!DailyRate), 0, !DailyRate)
        litem1.SubItems(30) = Format(DaysWorked * !DailyRate, "0.00")     ' Basic Pay
        
'        RestDayWorkedOTRate = 0.3
'        SpecialWorkedDayOffRate = 0.5
'        LegalHolidayWithPremiumRate = 1#
'        LegalHolidayWorkedDayoffRate = 1.6
'        WorkHrsPerDay = 8
'
'        OverTimeRate = 1.25
'        ExcessSplOTRate = 1.69
'        ExcessLglHolOTRate = 2.6
'        ExcessSunLglOTRate = 3.38
'        NightDiffRate = 0.1
'
'        AgencyRate = 15
'        VatPercentRate = 0.12

        'litem1.SubItems(31) = Format(!RestDayOTPay, "0.00")     ' Premium pay on extra days worked
        RestDayOTPay = IIf(IsNull(!DailyRate), 0, !DailyRate) * RestDayWorkedOTRate * RestDayOTHrs
        litem1.SubItems(31) = Format(RestDayOTPay, "0.00")     ' Premium pay on ey on extra days worked
        LglHolidayPay = IIf(IsNull(!DailyRate), 0, !DailyRate) * LegalHolidayWorkedDayoffRate * NoLglHoliday
        litem1.SubItems(34) = Format(LglHolidayPay, "0.00")    ' premium pay on extra days worked
        
        'litem1.SubItems(35) = Format(!TardinessDeduction, "0.00")
        TardinessDeduction = HoursTardiness * (IIf(IsNull(!DailyRate), 0, !DailyRate) / WorkHrsPerDay)
        litem1.SubItems(35) = Format(TardinessDeduction, "0.00")
        
        'litem1.SubItems(36) = Format(!OTPay, "0.00")            ' ot
        OTPay = (IIf(IsNull(!DailyRate), 0, !DailyRate) / WorkHrsPerDay) * OverTimeRate * OTHours
        litem1.SubItems(36) = Format(OTPay, "0.00")            ' ot
        
        'litem1.SubItems(37) = Format(!excesssplholotpay, "0.00") ' ot
        ExcessSplOTPay = (IIf(IsNull(!DailyRate), 0, !DailyRate) / WorkHrsPerDay) * ExcessSplOTRate * ExcessSplOTHours
        litem1.SubItems(37) = Format(ExcessSplOTPay, "0.00") ' ot
        litem1.SubItems(38) = "-"
        'litem1.SubItems(39) = Format(!lglholidayotpay, "0.00")  ' ot
        LglHolidayOTPay = (IIf(IsNull(!DailyRate), 0, !DailyRate) / WorkHrsPerDay) * ExcessLglHolOTRate * LglHolidayOTHours
        litem1.SubItems(39) = Format(LglHolidayOTPay, "0.00") ' ot
        'litem1.SubItems(40) = Format(!excessotsunlegalpay, "0.00")  ' ot
        ExcessSunLegalPay = (IIf(IsNull(!DailyRate), 0, !DailyRate) / WorkHrsPerDay) * ExcessSunLglOTRate * ExcessOTSunLegalHrs
        litem1.SubItems(40) = Format(ExcessSunLegalPay, "0.00") ' ot

        'litem1.SubItems(41) = Format(!NightDiffPay, "0.00") ' ot
        NightDiffPay = (IIf(IsNull(!DailyRate), 0, !DailyRate) / WorkHrsPerDay) * NightDiffRate * NightDiffOT
        litem1.SubItems(41) = Format(NightDiffPay, "0.00") ' Night diff

        litem1.SubItems(42) = Format(!EcolaPay, "0.00")
        litem1.SubItems(43) = Format(!SalaryAdjust, "0.00")
        SalaryAdjust = IIf(IsNull(!SalaryAdjust2), 0, !SalaryAdjust2)
        litem1.SubItems(44) = Format(IIf(IsNull(!SalaryAdjust2), 0, !SalaryAdjust2), "0.00")
        'litem1.SubItems(45) = Format(!GrossPay, "0.00")
        GrossPay = BasicPay + RestDayOTPay + SplHolidayPay + SundayOTPay + LglHolidayPay + _
            OTPay + ExcessSplOTPay + LglHolidayOTPay + ExcessSunLegalPay + _
            ExcessSunLegalPay + NightDiffPay + SalaryAdjust - TardinessDeduction
        
        litem1.SubItems(45) = Format(GrossPay, "0.00")
        'litem1.SubItems(46) = Format(!Month13Payment, "0.00")
        Month13Pay = BasicPay / 12
        litem1.SubItems(46) = Format(Month13Pay, "0.00")

        'litem1.SubItems(47) = Format(!SSSEmployer + !ECCemployer, "0.00")
        SSSEmployer = DaysWorked * IIf(IsNull(!SSSRate), 0#, !SSSRate)
        litem1.SubItems(47) = Format(SSSEmployer, "0.00")
        'litem1.SubItems(48) = Format(!PHealthEmployer, "0.00")
        PHealthEmployer = DaysWorked * IIf(IsNull(!PHRate), 0#, !PHRate)
        litem1.SubItems(48) = Format(PHealthEmployer, "0.00")
        'litem1.SubItems(49) = Format(!PagibigEmployer, "0.00")
        If DaysWorked = 0 Then
            PagibigEmployer = 0#
        Else
            PagibigEmployer = (DaysWorked / DaysWorked) * 50
        End If
        litem1.SubItems(49) = Format(PagibigEmployer, "0.00")
        
        'litem1.SubItems(50) = Format(!GrossPay + !SSSEmployer + !ECCemployer + !PHealthEmployer + !PagibigEmployer, "0.00") ' Total pay and benefits
        litem1.SubItems(50) = Format(GrossPay + SSSEmployer + PHealthEmployer + PagibigEmployer, "0.00") ' Total pay and benefits
        'litem1.SubItems(51) = Format(!agencyCharge, "0.00") ' Agency Charge
        AgencyCharge = DaysWorked * AgencyRate
        litem1.SubItems(51) = Format(AgencyCharge, "0.00") ' Agency Charge
        
        GrossBill = GrossPay + SSSEmployer + PHealthEmployer + PagibigEmployer + AgencyCharge
        
        'litem1.SubItems(39) = Format(!Grosspay + !SSSemployer + !ECCemployer + !phealthemployer + !pagibigemployer + (!DaysPresent * !ClientChargePerDay), "0.00") ' Gross Billing
        litem1.SubItems(52) = Format(GrossBill, "0.00") ' Gross Billing
        VATPayment = (GrossBill * VatPercentRate)
        litem1.SubItems(53) = Format(VATPayment, "0.00") 'Vat 12%
        litem1.SubItems(54) = Format(!IncentivePay, "0.00") ' Incentive pay
        IncentivePay = IIf(IsNull(!IncentivePay2), 0, !IncentivePay2)
        litem1.SubItems(55) = Format(!IncentivePay2, "0.00") ' Edit Incentive pay
        litem1.SubItems(56) = Format(GrossBill + VATPayment + IncentivePay, "0.00") ' Total Amount Due
        
        gGrossBill = gGrossBill + GrossBill + VATPayment + IncentivePay
        
        litem1.SubItems(57) = "::"
        litem1.SubItems(58) = Format(!BasicPay, "0.00") 'Original BasicPay
        litem1.SubItems(59) = "-"   ' Ignore Lates
        litem1.SubItems(60) = Format(GrossPay - GrossPay, "0.00") ' Diff'l

        litem1.SubItems(61) = "==" 'Format((DaysWorked * AgencyRate), "0.00") ' Agency Fee
    
        litem1.SubItems(62) = Format(!StartPayDate, "mm/dd/yyyy") ' Start Cutoff
        litem1.SubItems(63) = Format(!EndPayDate, "mm/dd/yyyy")  ' End Cutoff
        litem1.SubItems(64) = IIf(IsNull(!soano), "", !soano) 'EmployeeNum
        litem1.SubItems(65) = IIf(IsNull(!SSSRate), 0, !SSSRate) 'SSS Rate
        litem1.SubItems(66) = IIf(IsNull(!PHRate), 0, !PHRate) 'PH Rate
        
        If !ExcludedFlag = True Then
            ListView1(0).ListItems.Item(Listctr).Checked = True
        End If

NextRecc:

    .MoveNext
Wend
End With
'If Listctr <> 0 Then
'    strQuery = "SELECT a.Clientcode, a.ClientName, b.Remarks, b.ApplicableMonth, b.ApplicableYear, b.ProcessedDate, b.ProcessedBy, a.ClientFlag3 AS EcolaFlag, b.GroupCode, b.SBRNo " & _
'        "FROM clienttable a LEFT JOIN (SELECT DISTINCT AgencyCode, ClientCode, ClientName, ContriCode, Remarks, ApplicableMonth, ApplicableYear, ProcessedDate, ProcessedBy, GroupCode, SBRNo, AmountPaid, DatePaid, DatePosted " & _
'        "FROM BilledTrans WHERE ApplicableMonth = '" & txtAppMonth.Text & "' AND ApplicableYear = " & txtAppYear.Text & ") AS b " & _
'        "ON a.ClientName = b.ClientName " & _
'        "WHERE a.clientstatus = true and a.clientname is not null ORDER BY a.clientname"
'
'    Set rsClientData = ReturnRst(strQuery, cnnStr, mstatic)
'    If rsClientData.RecordCount <> 0 Then
'        Call FillClientList2(ListView1(0), rsClientData)
'        rsClientData.MoveFirst
'    End If
'    AddToListFlag = False
'End If
'Call GetTotals
'MsgBox "There were " & Listctr & " records added to PhilHealth.", vbInformation, PayrollAlert
'cmdAddlist.Enabled = False  ' To prevent double adding of the same company

ExitRtn2:

frmBilling.txtGrossBill.Text = Format(gGrossBill, "0.00")
frmBilling.lblEntries(7).Caption = "Total Records: " & Listctr

End Sub

'Private Sub PopulateBillingRecords(ByRef xlApp1 As Excel.Application, ByRef xlSheet As Excel.Worksheet)
''Dim xlapp1 As Excel.Application
''Dim xlBook1 As Excel.Workbook
''Dim xlsheet As Excel.Worksheet
'Dim mTotalPay As Double
'Dim mTotalDeduction As Double
'Dim mRange As String, mRange1 As String
'Dim SaveTotal As String
'Dim ColString As String
'Dim sumString As String
'
'Dim ctr As Long
'Dim i As Integer, j As Integer, k As Integer
'Dim cPos As Integer
'Dim EndSw As Boolean
'Dim mHead As String
'Dim firstRun As Boolean
'
'Dim SubBDO(50) As Double
'Dim SubATM(50) As Double
'Dim SubCash(50) As Double
'Dim SubHold(50) As Double
'Dim bdoctr, atmctr, cashctr, holdctr
'Dim LineCtr As Integer
'Dim bdopos As Integer
'Dim atmpos As Integer
'Dim cashpos As Integer
'Dim holdpos As Integer
'Dim BillingTemplate As String
'Dim RecCtr As Integer
'Dim prevCutOff As Variant
'Dim PrevRate As Double
'
''Set xlApp1 = New Excel.Application
''Set xlBook1 = xlApp1.Workbooks.Add()
''
''xlApp1.WindowState = xlMaximized
''xlApp1.Visible = True
''xlApp1.DisplayAlerts = False
''
'''Set xlSheet = book1.Worksheets.Add
''Set xlSheet = xlBook1.Worksheets(1)
''
''BillingTemplate = App.Path & "\Reports\RF-1 Template.xls"
'''ExcelTemplate = App.Path & "\Reports\RF-1 New Template.xls"
''
''Set xlApp1 = New Excel.Application
'''Set xlBook1 = xlApp1.Workbooks.Open(ExcelTemplate)
''Set xlBook1 = xlApp1.Workbooks.Add(BillingTemplate)
''
''xlApp1.WindowState = xlMaximized
''xlApp1.Visible = True
''xlApp1.DisplayAlerts = False
''
'''Set xlSheet = book1.Worksheets.Add
''Set xlSheet = xlBook1.Worksheets(1)
'
'xlApp1.Sheets("Billing Report").Select
'SaveTotal = ""
'ColString = ""
'cPos = 10
'RecCtr = 1
'ctr = 7
'LineCtr = 10
'i = 1
'k = 10
'
''On Error GoTo Errh
''Print Header
''xlSheet.Range("A2") = mCompanyName & msapi  '" (SAPI)"
'
''' Updated by Bernard May 13, 2008.
''mHead = mCompanyName & " (" & CompanyShort & ")"
'''xlSheet.Range("A2") = mCompanyName & " (" & CompanyShort & ")"
''Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 14, True)
''If Format(rs!StartPayDate, "mmm") = Format(rs!EndPayDate, "mmm") Then
''    'xlSheet.Range("A3") = "Payroll Summary Report for Cut-Off " & Format(rs!startPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "dd, yyyy")
''    mHead = "Payroll Summary Report for Cut-Off " & Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "dd, yyyy")
''Else
''    'xlSheet.Range("A3") = "Payroll Summary Report for Cut-Off " & Format(rs!startPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd, yyyy")
''    mHead = "Payroll Summary Report for Cut-Off " & Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd, yyyy")
''End If
''Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, False)
''
'''xlSheet.Range("AO3") = mSeriesNo
''Call PrintReportHeader(xlSheet, xlApp1, "AS3", mSeriesNo, "", xlLeft, , False, False, False, , , 8, False)
'''xlSheet.Range("AO4") = Format(Now, "mm/dd/yyyy")
''Call PrintReportHeader(xlSheet, xlApp1, "AS4", Format(Now, "mm/dd/yyyy"), "", xlLeft, xlBottom, False, False, False, , , 8, False)
'
'If rsBilledTrans2.RecordCount = 0 Then
'    MsgBox "No records to process. Please recompute Transaction Records again.", vbInformation, PayrollAlert
'    GoTo ExitRtn1
'End If
'With rsBilledTrans2
'    .MoveFirst
'prevCutOff = Format(!StartPayDate, "mmddyyyy") & LTrim(RTrim(!PayText2))
'PrevRate = !DailyRate
'
'While Not (.EOF Or .BOF)
'
'    If !ExcludedFlag = True Then
'        GoTo NextRec
'    End If
'
'    If prevCutOff <> Format(!StartPayDate, "mmddyyyy") & LTrim(RTrim(!PayText2)) Then
'        If i > 1 Then
'            Call PrintTotals(xlSheet, xlApp1, LineCtr, k)
'            LineCtr = LineCtr + 1
'            'ctr = ctr + 1
'            'i = i + 1
'            k = LineCtr
'            SaveTotal = SaveTotal & LineCtr - 1 & ":"
'        End If
'        'prevCutOff = !StartPayDate
'        prevCutOff = Format(!StartPayDate, "mmddyyyy") & LTrim(RTrim(!PayText2))
'        RecCtr = 1
'    ElseIf PrevRate <> !DailyRate Then
'        If i > 1 Then
'            Call PrintTotals(xlSheet, xlApp1, LineCtr, k)
'            LineCtr = LineCtr + 1
'            'ctr = ctr + 1
'            'i = i + 1
'            k = LineCtr
'            SaveTotal = SaveTotal & LineCtr - 1 & ":"
'        End If
'        'prevCutOff = !StartPayDate
'        prevCutOff = Format(!StartPayDate, "mmddyyyy") & LTrim(RTrim(!PayText2))
'        PrevRate = !DailyRate
'        RecCtr = 1
'    End If
'
'        mRange = "A" & LineCtr
'        xlSheet.Range(mRange) = RecCtr  'IIf(IsNull(!ControlNo), "", !ControlNo)
'        mRange = "B" & LineCtr
'        xlSheet.Range(mRange) = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
'        mRange = "C" & LineCtr
'        xlSheet.Range(mRange) = Space(5)
'        DaysPresent = Format(!DaysPresent2, "0")    ' days present
'        mRange = "D" & LineCtr
'        xlSheet.Range(mRange) = Format(DaysPresent, "0")
'        RestDayOTHrs = Format(!RestDayOTHrs2, "0")  ' rest day ot hours
'        mRange = "E" & LineCtr
'        xlSheet.Range(mRange) = Format(RestDayOTHrs, "0")
'        NoSplHoliday = Format(!NoSplHoliday2, "0")  ' spl holiday
'        mRange = "F" & LineCtr
'        xlSheet.Range(mRange) = Format(NoSplHoliday, "0")
'        SundayOT = Format(!SundayOT2, "0")          ' sunday ot
'        NoLglHoliday = Format(!NoLglHoliday2, "0")  ' lgl holiday
'        mRange = "G" & LineCtr
'        xlSheet.Range(mRange) = Format(NoLglHoliday, "0.00") 'Format(SundayOT, "0")
'        mRange = "H" & LineCtr
'        xlSheet.Range(mRange) = Format(SundayOT, "0.00") 'Format(NoLglHoliday, "0")
'        mRange = "I" & LineCtr
'        DaysWorked = DaysPresent + RestDayOTHrs + NoSplHoliday + SundayOT
'        xlSheet.Range(mRange) = Format(DaysWorked, "0")  ' computed
'
'        HoursTardiness = Format(!HoursTardiness2, "0.00")
'        mRange = "J" & LineCtr
'        xlSheet.Range(mRange) = Format(HoursTardiness, "0.00")  ' tardi hours
'        OTHours = Format(!OTHours2, "0.00")
'        mRange = "K" & LineCtr
'        xlSheet.Range(mRange) = Format(OTHours, "0.00")         ' Reg ot
'        ExcessSplOTHours = Format(IIf(IsNull(!ExcessSplHolOTHours2), 0#, !ExcessSplHolOTHours2), "0.00")
'        mRange = "L" & LineCtr
'        xlSheet.Range(mRange) = Format(ExcessSplOTHours, "0.00")      ' day off spl
'        mRange = "M" & LineCtr
'        xlSheet.Range(mRange) = ""  ' -
'        LglHolidayOTHours = !LglHolidayOTHours2
'        mRange = "N" & LineCtr
'        xlSheet.Range(mRange) = Format(LglHolidayOTHours, "0.00")     ' Lgl holiday ot
'        ExcessOTSunLegalHrs = !ExcessOTSunLegalHrs2
'        mRange = "O" & LineCtr
'        xlSheet.Range(mRange) = Format(ExcessOTSunLegalHrs, "0.00")  ' Sun legal ot
'        NightDiffOT = !NightDiffOT2
'        mRange = "P" & LineCtr
'        xlSheet.Range(mRange) = Format(NightDiffOT, "0.00")           ' night diff ot
'        mRange = "Q" & LineCtr
'        xlSheet.Range(mRange) = Format(!DailyRate, "0.00")            ' daily rate
'        BasicPay = DaysWorked * !DailyRate
'        mRange = "R" & LineCtr
'        xlSheet.Range(mRange) = Format(BasicPay, "0.00")              ' Basic pay
'
'        RestDayWorkedOTRate = 0.3
'        SpecialWorkedDayOffRate = 0.5
'        LegalHolidayWithPremiumRate = 1#
'        LegalHolidayWorkedDayoffRate = 1.6
'        WorkHrsPerDay = 8
'
'        OverTimeRate = 1.25
'        ExcessSplOTRate = 1.69
'        ExcessLglHolOTRate = 2.6
'        ExcessSunLglOTRate = 3.38
'        NightDiffRate = 0.1
'
'        AgencyRate = 15
'        VatRate = 0.12
'
'        RestDayOTPay = !DailyRate * RestDayWorkedOTRate * RestDayOTHrs
'        mRange = "S" & LineCtr
'        xlSheet.Range(mRange) = Format(RestDayOTPay, "0.00")              ' rest day ot pay
'        SplHolidayPay = !DailyRate * SpecialWorkedDayOffRate * NoSplHoliday
'        mRange = "T" & LineCtr
'        xlSheet.Range(mRange) = Format(SplHolidayPay, "0.00")              ' spl holiday ot pay
'        SundayOTPay = !DailyRate * LegalHolidayWithPremiumRate * SundayOT
'        LglHolidayPay = !DailyRate * LegalHolidayWorkedDayoffRate * NoLglHoliday
'        mRange = "U" & LineCtr
'        xlSheet.Range(mRange) = Format(LglHolidayPay, "0.00") 'Format(SundayOTPay, "0.00")              ' sunday ot pay
'        mRange = "V" & LineCtr
'        xlSheet.Range(mRange) = Format(SundayOTPay, "0.00") 'Format(LglHolidayPay, "0.00")              ' lgl holiday pay
'        TardinessDeduction = HoursTardiness * (!DailyRate / WorkHrsPerDay)
'        mRange = "W" & LineCtr
'        xlSheet.Range(mRange) = Format(TardinessDeduction, "0.00")           ' tardi pay
'        OTPay = (!DailyRate / WorkHrsPerDay) * OverTimeRate * OTHours
'        mRange = "X" & LineCtr
'        xlSheet.Range(mRange) = Format(OTPay, "0.00")              ' OT Pay
'        ExcessSplOTPay = (!DailyRate / WorkHrsPerDay) * ExcessSplOTRate * ExcessSplOTHours
'        mRange = "Y" & LineCtr
'        xlSheet.Range(mRange) = Format(ExcessSplOTPay, "0.00")              ' excess spl ot pay
'        mRange = "Z" & LineCtr
'        xlSheet.Range(mRange) = "-"
'        LglHolidayOTPay = (!DailyRate / WorkHrsPerDay) * ExcessLglHolOTRate * LglHolidayOTHours
'        mRange = "AA" & LineCtr
'        xlSheet.Range(mRange) = Format(LglHolidayOTPay, "0.00")              ' lgl holiday ot pay
'        ExcessSunLegalPay = (!DailyRate / WorkHrsPerDay) * ExcessSunLglOTRate * ExcessOTSunLegalHrs
'        mRange = "AB" & LineCtr
'        xlSheet.Range(mRange) = Format(ExcessSunLegalPay, "0.00")             ' excess sun legal pay
'        NightDiffPay = (!DailyRate / WorkHrsPerDay) * NightDiffRate * NightDiffOT
'        mRange = "AC" & LineCtr
'        xlSheet.Range(mRange) = Format(NightDiffPay, "0.00")                   ' night diff pay
'
'        mRange = "AD" & LineCtr
'        xlSheet.Range(mRange) = Format(!EcolaPay, "0.00")                   ' ecola
'        SalaryAdjust = !SalaryAdjust2
'        mRange = "AE" & LineCtr
'        xlSheet.Range(mRange) = Format(SalaryAdjust, "0.00")                   ' salary adjust
'        GrossPay = BasicPay + RestDayOTPay + SplHolidayPay + SundayOTPay + LglHolidayPay + _
'            OTPay + ExcessSplOTPay + LglHolidayOTPay + ExcessSunLegalPay + _
'            ExcessSunLegalPay + NightDiffPay + SalaryAdjust - TardinessDeduction
'        mRange = "AF" & LineCtr
'        xlSheet.Range(mRange) = Format(GrossPay, "0.00")                   ' gross pay
'        Month13Pay = BasicPay / 12
'        mRange = "AG" & LineCtr
'        xlSheet.Range(mRange) = Format(Month13Pay, "0.00")                   ' month 13 pay
'
'        SSSEmployer = DaysWorked * IIf(IsNull(!SSSRate), 0#, !SSSRate)
'        mRange = "AH" & LineCtr
'        xlSheet.Range(mRange) = Format(SSSEmployer, "0.00")                   ' sss employer
'        PHealthEmployer = DaysWorked * IIf(IsNull(!PHRate), 0#, !PHRate)
'        mRange = "AI" & LineCtr
'        xlSheet.Range(mRange) = Format(PHealthEmployer, "0.00")                   ' phealth
'        PagibigEmployer = (DaysWorked / DaysWorked) * 50
'        mRange = "AJ" & LineCtr
'        xlSheet.Range(mRange) = Format(PagibigEmployer, "0.00")                   ' pagibig
'
'        mRange = "AK" & LineCtr
'        xlSheet.Range(mRange) = Format(GrossPay + SSSEmployer + PHealthEmployer + PagibigEmployer, "0.00") ' Total pay and benefits
'        AgencyCharge = DaysWorked * AgencyRate
'        mRange = "AL" & LineCtr
'        xlSheet.Range(mRange) = Format(AgencyCharge, "0.00")                   ' agency charge
'        GrossBill = GrossPay + SSSEmployer + PHealthEmployer + PagibigEmployer + AgencyCharge
'        mRange = "AM" & LineCtr
'        xlSheet.Range(mRange) = Format(GrossBill, "0.00")                   ' gross billing
'        VATPayment = (GrossBill * VatRate)
'        mRange = "AN" & LineCtr
'        xlSheet.Range(mRange) = Format(VATPayment, "0.00")                   ' VAT %
'        IncentivePay = !IncentivePay2
'        mRange = "AO" & LineCtr
'        xlSheet.Range(mRange) = Format(IncentivePay, "0.00")                   ' other ded
'        mRange = "AP" & LineCtr
'        xlSheet.Range(mRange) = Format(GrossBill + VATPayment + IncentivePay, "0.00") ' Total Amount Due
'        mRange = "AQ" & LineCtr
'        xlSheet.Range(mRange) = Space(5)
'        mRange = "AR" & LineCtr
'        xlSheet.Range(mRange) = Format(!BasicPay, "0.00")                   ' original basic pay
'
'        mRange = "AU" & LineCtr
'        xlSheet.Range(mRange) = Format(!StartPayDate, "mm/dd/yyyy")                   ' original basic pay
'        mRange = "AV" & LineCtr
'        xlSheet.Range(mRange) = Format(!EndPayDate, "mm/dd/yyyy")                   ' original basic pay
'        mRange = "AW" & LineCtr
'        xlSheet.Range(mRange) = Format(!SSSRate, "0.00")                   ' original basic pay
'        mRange = "AX" & LineCtr
'        xlSheet.Range(mRange) = Format(!PHRate, "0.00")                   ' original basic pay
'
'    LineCtr = LineCtr + 1
'    RecCtr = RecCtr + 1
'    i = i + 1
'
'NextRec:
'
'    .MoveNext
'
'Wend
'End With
'
'' Print Last record.
'Call PrintTotals(xlSheet, xlApp1, LineCtr, k)
'sumString = ""
'If Len(SaveTotal) = 0 Then
'    SaveTotal = LineCtr & ":"
'Else
'    SaveTotal = SaveTotal & LineCtr & ":"
'End If
'ColString = SaveTotal
'LineCtr = LineCtr + 1
'
''Compute for Grand Total.
'mRange = "B" & LineCtr & ":B" & LineCtr
'xlSheet.Range(mRange).Select
'xlApp1.ActiveCell.FormulaR1C1 = "  Grand Total"
'xlApp1.Range(mRange).Font.Bold = True
'firstRun = True
'
'Dim mGTotal As Double
'Dim lf, aa
'
'mGTotal = 0#
'For j = Asc("D") To Asc("Z")
''    If FirstRun = True Then
''        If Asc("H") = j Then
''            j = Asc("D")
''        Else
''            j = Asc("H")
''            FirstRun = False
''        End If
''    End If
'    SaveTotal = ColString
'    While Len(SaveTotal) <> 0
'        lf = InStr(SaveTotal, ":")
'        aa = LTrim(RTrim(Left(SaveTotal, lf - 1)))
'        'If sumString = "" Then ' orig code
'            sumString = "=R[" & ((LineCtr - aa) * -1) & "]C"
'            mRange = Chr(j) & aa
'            mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
'            'xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
'        'Else
'        '    sumString = sumString & "+R[" & ((LineCtr - aa) * -1) & "]C"
'        'End If
'        SaveTotal = Mid(SaveTotal, lf + 1, Len(SaveTotal) - lf)
'    Wend
'    mRange = Chr(j) & LineCtr
'    xlSheet.Range(mRange).Select
'    'xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)" orig code
'    xlApp1.ActiveCell.VALUE = mGTotal
'    'ActiveCell.FormulaR1C1 = "=R[-5]C+R[-3]C"
'    xlSheet.Range(mRange).Select
''    If xlApp1.ActiveCell.VALUE = 0 Then
''        xlApp1.Selection.EntireColumn.Hidden = True
''    End If
'    sumString = ""
'    mGTotal = 0#
'Next
'mRange = "D" & LineCtr & ":Z" & LineCtr
'xlApp1.Range(mRange).Font.Bold = True
'
'For j = Asc("A") To Asc("R")
'    SaveTotal = ColString
'    While Len(SaveTotal) <> 0
'        lf = InStr(SaveTotal, ":")
'        aa = LTrim(RTrim(Left(SaveTotal, lf - 1)))
'        'If sumString = "" Then
'            sumString = "=R[" & ((LineCtr - aa) * -1) & "]C"
'            mRange = "A" & Chr(j) & aa
'            mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
''        Else
''            sumString = sumString & "+R[" & ((LineCtr - aa) * -1) & "]C"
''        End If
'        SaveTotal = Mid(SaveTotal, lf + 1, Len(SaveTotal) - lf)
'    Wend
'    mRange = "A" & Chr(j) & LineCtr
'    xlSheet.Range(mRange).Select
'    'xlApp1.ActiveCell.FormulaR1C1 = sumString
'    xlApp1.ActiveCell.VALUE = mGTotal
'    ' Get the Total Amount Due.
'    If Asc("P") = j Then
'        xlApp1.Selection.Copy
'        Call CopyToSheetDataSource(xlSheet, xlApp1, mGTotal)
'    End If
'
'    xlSheet.Range(mRange).Select
''    If xlApp1.ActiveCell.VALUE = 0 Then
''        xlApp1.Selection.EntireColumn.Hidden = True
''    End If
'    sumString = ""
'    mGTotal = 0#
'Next
'mRange = "AA" & LineCtr & ":AR" & LineCtr
'xlApp1.Range(mRange).Font.Bold = True
'
'mRange = "A10:AX" & LineCtr + 2
'xlSheet.Range(mRange).Select
'With xlApp1.Range(mRange)
'    .Font.Name = "Arial"    'Font.Name
'    .Font.Size = 8          'Font.size
'    '.Font.Bold = False      'Font.Bold
'    .Font.Italic = False    'Font.Italic
'End With
'
''Format amount to 0.00.
'mRange = "E10:AR" & LineCtr + 2    ' E7:AP
'xlSheet.Range(mRange).Select
'xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'With xlApp1.Selection
'    .HorizontalAlignment = xlRight
'    .VerticalAlignment = xlBottom
'    .WrapText = False
'    .Orientation = 0
'    .AddIndent = False
'    .ShrinkToFit = False
'    .MergeCells = False
'    .Style = "Comma"
'End With
'
'mRange = "A10:AX" & LineCtr + 2
'Call PlaceBox(xlSheet, xlApp1, mRange)
'
'mRange = "A1"
'xlSheet.Range(mRange).Select
'
'Exit Sub
'
'ExitRtn1:
'
''    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
''    Err.Clear
'
''xlSheet.Name = "Overtime"
'
'End Sub

Private Sub PopulateBillingRecords(ByRef xlApp1 As Excel.Application, ByRef xlSheet As Excel.Worksheet, _
ByRef xlBook1 As Excel.Workbook)

If Mid(BillingTemplate, 1, 7) = "Carlson" Then
    Set xlSheet = xlBook1.Worksheets("15th(Payroll)")
    Call CarlsonBillingReports(xlApp1, xlSheet)
Else
    Set xlSheet = xlBook1.Worksheets("Billing Report")
    Call JapanAndOtherClientsBillingReports(xlApp1, xlSheet)
End If

End Sub

Private Sub JapanAndOtherClientsBillingReports(ByRef xlApp1 As Excel.Application, ByRef xlSheet As Excel.Worksheet)
'Dim xlapp1 As Excel.Application
'Dim xlBook1 As Excel.Workbook
'Dim xlsheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String, mRange1 As String
Dim SaveTotal As String
Dim ColString As String
Dim sumString As String

Dim ctr As Long
Dim i As Integer, j As Integer, k As Integer
Dim cPos As Integer
Dim EndSw As Boolean
Dim mHead As String
Dim firstRun As Boolean

Dim SubBDO(50) As Double
Dim SubATM(50) As Double
Dim SubCash(50) As Double
Dim SubHold(50) As Double
Dim bdoctr, atmctr, cashctr, holdctr
Dim LineCtr As Integer
Dim bdopos As Integer
Dim atmpos As Integer
Dim cashpos As Integer
Dim holdpos As Integer
'Dim BillingTemplate As String
Dim RecCtr As Integer
Dim prevCutOff As Variant
Dim PrevRate As Double

'Set xlApp1 = New Excel.Application
'Set xlBook1 = xlApp1.Workbooks.Add()
'
'xlApp1.WindowState = xlMaximized
'xlApp1.Visible = True
'xlApp1.DisplayAlerts = False
'
''Set xlSheet = book1.Worksheets.Add
'Set xlSheet = xlBook1.Worksheets(1)
'
'BillingTemplate = App.Path & "\Reports\RF-1 Template.xls"
''ExcelTemplate = App.Path & "\Reports\RF-1 New Template.xls"
'
'Set xlApp1 = New Excel.Application
''Set xlBook1 = xlApp1.Workbooks.Open(ExcelTemplate)
'Set xlBook1 = xlApp1.Workbooks.Add(BillingTemplate)
'
'xlApp1.WindowState = xlMaximized
'xlApp1.Visible = True
'xlApp1.DisplayAlerts = False
'
''Set xlSheet = book1.Worksheets.Add
'Set xlSheet = xlBook1.Worksheets(1)

xlApp1.Sheets("Billing Report").Select
SaveTotal = ""
ColString = ""
cPos = 10
RecCtr = 1
ctr = 7
LineCtr = 12
i = 1
k = 12

'On Error GoTo Errh
'Print Header
'xlSheet.Range("A2") = mCompanyName & msapi  '" (SAPI)"

'' Updated by Bernard May 13, 2008.
'mHead = mCompanyName & " (" & CompanyShort & ")"
''xlSheet.Range("A2") = mCompanyName & " (" & CompanyShort & ")"
'Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 14, True)
'If Format(rs!StartPayDate, "mmm") = Format(rs!EndPayDate, "mmm") Then
'    'xlSheet.Range("A3") = "Payroll Summary Report for Cut-Off " & Format(rs!startPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "dd, yyyy")
'    mHead = "Payroll Summary Report for Cut-Off " & Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "dd, yyyy")
'Else
'    'xlSheet.Range("A3") = "Payroll Summary Report for Cut-Off " & Format(rs!startPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd, yyyy")
'    mHead = "Payroll Summary Report for Cut-Off " & Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd, yyyy")
'End If
'Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, False)
'
''xlSheet.Range("AO3") = mSeriesNo
'Call PrintReportHeader(xlSheet, xlApp1, "AS3", mSeriesNo, "", xlLeft, , False, False, False, , , 8, False)
''xlSheet.Range("AO4") = Format(Now, "mm/dd/yyyy")
'Call PrintReportHeader(xlSheet, xlApp1, "AS4", Format(Now, "mm/dd/yyyy"), "", xlLeft, xlBottom, False, False, False, , , 8, False)

If rsBilledTrans2.RecordCount = 0 Then
    MsgBox "No records to process. Please recompute Transaction Records again.", vbInformation, PayrollAlert
    GoTo ExitRtn1
End If
With rsBilledTrans2
    .MoveFirst
    
prevCutOff = Format(!StartPayDate, "mmddyyyy") & LTrim(RTrim(!Paytext2))
PrevRate = !DailyRate

While Not (.EOF Or .BOF)

    If chkExcludedRecords.VALUE = False Then
        If !ExcludedFlag = True Then
            GoTo NextRec
        End If
    End If
    
    If prevCutOff <> Format(!StartPayDate, "mmddyyyy") & LTrim(RTrim(!Paytext2)) Then
        If i > 1 Then
            ctr = LineCtr   ' save the sub - total formulation
            Call PrintTotals(xlSheet, xlApp1, LineCtr, k)
            LineCtr = LineCtr + 2
            
            mRange = "A" & ctr & ":AV" & ctr
            xlSheet.Range(mRange).Select
            xlApp1.Selection.Copy
            mRange = "A" & LineCtr
            xlApp1.ActiveSheet.Range(mRange).Select
            xlApp1.Selection.Insert Shift:=xlDown
            'ctr = ctr + 1
            'i = i + 1
            k = LineCtr
            'SaveTotal = SaveTotal & LineCtr - 1 & ":"
            SaveTotal = SaveTotal & ctr & ":"
        End If
        'prevCutOff = !StartPayDate
        prevCutOff = Format(!StartPayDate, "mmddyyyy") & LTrim(RTrim(!Paytext2))
        'RecCtr = 1
    ElseIf PrevRate <> !DailyRate Then
        If i > 1 Then
            ctr = LineCtr
            Call PrintTotals(xlSheet, xlApp1, LineCtr, k)
            LineCtr = LineCtr + 2
            
            mRange = "A" & ctr & ":AV" & ctr
            xlSheet.Range(mRange).Select
            xlApp1.Selection.Copy
            mRange = "A" & LineCtr
            xlApp1.ActiveSheet.Range(mRange).Select
            xlApp1.Selection.Insert Shift:=xlDown
            
            'ctr = ctr + 1
            'i = i + 1
            k = LineCtr
            'SaveTotal = SaveTotal & LineCtr - 1 & ":"
            SaveTotal = SaveTotal & ctr & ":"
        End If
        'prevCutOff = !StartPayDate
        prevCutOff = Format(!StartPayDate, "mmddyyyy") & LTrim(RTrim(!Paytext2))
        PrevRate = !DailyRate
        'RecCtr = 1
    End If
    
    If LineCtr > 12 Then    ' paste the template
        mRange = "A12:AV12"
        xlSheet.Range(mRange).Select
        xlApp1.Selection.Copy
        mRange = "A" & LineCtr
        xlApp1.ActiveSheet.Range(mRange).Select
        'xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone , SkipBlanks:=
            'False, Transpose:=False
        xlApp1.Selection.Insert Shift:=xlDown
    'Rows("1:1").Select
    'Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    End If
    
        mRange = "A" & LineCtr
        xlSheet.Range(mRange) = RecCtr  'IIf(IsNull(!ControlNo), "", !ControlNo)
        mRange = "B" & LineCtr
        'xlSheet.Range(mRange) = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
        xlSheet.Range(mRange) = IIf(IsNull(!Employees), "No Name", !Employees)
        'mRange = "C" & LineCtr
        'xlSheet.Range(mRange) = Space(5)
        DaysPresentKey = Format(!DaysPresent, "0.00")  ' days present
        DaysPresent = Format(!DaysPresent2, "0.00")    ' days present
        RestDayOTHrs = Format(!RestDayOTHrs2, "0.00")  ' rest day ot hours
        SundayOT = Format(!SundayOT2, "0.00")          ' sunday ot
        NoSplHoliday = Format(!NoSplHoliday2, "0.00")  ' spl holiday
        'DaysPresent = DaysPresent - (RestDayOTHrs + SundayOT + NoSplHoliday)
        DaysPresent = DaysPresent - (RestDayOTHrs + NoSplHoliday)
        mRange = "D" & LineCtr
        xlSheet.Range(mRange) = Format(DaysPresent, "0.00")
        mRange = "E" & LineCtr
        xlSheet.Range(mRange) = Format(RestDayOTHrs + NoSplHoliday, "0.00")
        mRange = "F" & LineCtr
        xlSheet.Range(mRange) = 0# 'Format(NoSplHoliday, "0.00")
        NoLglHoliday = Format(!NoLglHoliday2, "0.00")  ' lgl holiday
        mRange = "G" & LineCtr
        xlSheet.Range(mRange) = Format(NoLglHoliday, "0.00") 'Format(SundayOT, "0.00")
        mRange = "H" & LineCtr
        xlSheet.Range(mRange) = 0#   'Format(SundayOT, "0.00") 'Format(NoLglHoliday, "0.00")
        mRange = "I" & LineCtr
        DaysWorked = DaysPresent + RestDayOTHrs + NoSplHoliday '+ SundayOT
        xlSheet.Range(mRange) = Format(DaysWorked, "0.00")  ' computed

        HoursTardiness = Format(!HoursTardiness2, "0.00")
        mRange = "J" & LineCtr
        xlSheet.Range(mRange) = Format(HoursTardiness, "0.00")  ' tardi hours
        OTHours = Format(!OTHours2, "0.00")
        mRange = "K" & LineCtr
        xlSheet.Range(mRange) = Format(OTHours, "0.00")         ' Reg ot
        ExcessSplOTHours = Format(SundayOT, "0.00")     ' sunday ot
        mRange = "L" & LineCtr
        xlSheet.Range(mRange) = Format(ExcessSplOTHours, "0.00")      ' day off spl
        mRange = "M" & LineCtr
        xlSheet.Range(mRange) = 0#    ' -
        LglHolidayOTHours = !LglHolidayOTHours2
        mRange = "N" & LineCtr
        xlSheet.Range(mRange) = Format(LglHolidayOTHours, "0.00")     ' Lgl holiday ot
        ExcessOTSunLegalHrs = !ExcessOTSunLegalHrs2
        mRange = "O" & LineCtr
        xlSheet.Range(mRange) = Format(ExcessOTSunLegalHrs, "0.00")  ' Sun legal ot
        NightDiffOT = !NightDiffOT2
        mRange = "P" & LineCtr
        xlSheet.Range(mRange) = Format(NightDiffOT, "0.00")           ' night diff ot
        mRange = "Q" & LineCtr
        xlSheet.Range(mRange) = Format(!DailyRate, "0.00")            ' daily rate
        BasicPay = DaysWorked * !DailyRate
        'mRange = "R" & LineCtr
        'xlSheet.Range(mRange) = Format(BasicPay, "0.00")              ' Basic pay
        
'        RestDayWorkedOTRate = 0.3
'        SpecialWorkedDayOffRate = 0.5
'        LegalHolidayWithPremiumRate = 1#
'        LegalHolidayWorkedDayoffRate = 1.6
'        WorkHrsPerDay = 8
'
'        OverTimeRate = 1.25
'        ExcessSplOTRate = 1.69
'        ExcessLglHolOTRate = 2.6
'        ExcessSunLglOTRate = 3.38
'        NightDiffRate = 0.1
'
'        AgencyRate = 15
'        VatRate = 0.12
        
'        RestDayOTPay = !DailyRate * RestDayWorkedOTRate * RestDayOTHrs
'        mRange = "S" & LineCtr
'        xlSheet.Range(mRange) = Format(RestDayOTPay, "0.00")              ' rest day ot pay
'        SplHolidayPay = !DailyRate * SpecialWorkedDayOffRate * NoSplHoliday
'        mRange = "T" & LineCtr
'        xlSheet.Range(mRange) = Format(SplHolidayPay, "0.00")              ' spl holiday ot pay
'        SundayOTPay = !DailyRate * LegalHolidayWithPremiumRate * SundayOT
'        LglHolidayPay = !DailyRate * LegalHolidayWorkedDayoffRate * NoLglHoliday
'        mRange = "U" & LineCtr
'        xlSheet.Range(mRange) = Format(LglHolidayPay, "0.00") 'Format(SundayOTPay, "0.00")              ' sunday ot pay
'        mRange = "V" & LineCtr
'        xlSheet.Range(mRange) = Format(SundayOTPay, "0.00") 'Format(LglHolidayPay, "0.00")              ' lgl holiday pay
'        TardinessDeduction = HoursTardiness * (!DailyRate / WorkHrsPerDay)
'        mRange = "W" & LineCtr
'        xlSheet.Range(mRange) = Format(TardinessDeduction, "0.00")           ' tardi pay
'        OTPay = (!DailyRate / WorkHrsPerDay) * OverTimeRate * OTHours
'        mRange = "X" & LineCtr
'        xlSheet.Range(mRange) = Format(OTPay, "0.00")              ' OT Pay
'        ExcessSplOTPay = (!DailyRate / WorkHrsPerDay) * ExcessSplOTRate * ExcessSplOTHours
'        mRange = "Y" & LineCtr
'        xlSheet.Range(mRange) = Format(ExcessSplOTPay, "0.00")              ' excess spl ot pay
'        mRange = "Z" & LineCtr
'        xlSheet.Range(mRange) = "-"
'        LglHolidayOTPay = (!DailyRate / WorkHrsPerDay) * ExcessLglHolOTRate * LglHolidayOTHours
'        mRange = "AA" & LineCtr
'        xlSheet.Range(mRange) = Format(LglHolidayOTPay, "0.00")              ' lgl holiday ot pay
'        ExcessSunLegalPay = (!DailyRate / WorkHrsPerDay) * ExcessSunLglOTRate * ExcessOTSunLegalHrs
'        mRange = "AB" & LineCtr
'        xlSheet.Range(mRange) = Format(ExcessSunLegalPay, "0.00")             ' excess sun legal pay
'        NightDiffPay = (!DailyRate / WorkHrsPerDay) * NightDiffRate * NightDiffOT
'        mRange = "AC" & LineCtr
'        xlSheet.Range(mRange) = Format(NightDiffPay, "0.00")                   ' night diff pay

        mRange = "AD" & LineCtr
        xlSheet.Range(mRange) = Format(!EcolaPay, "0.00")                   ' ecola
        SalaryAdjust = !SalaryAdjust2
        mRange = "AE" & LineCtr
        xlSheet.Range(mRange) = Format(SalaryAdjust, "0.00")                   ' salary adjust
        GrossPay = BasicPay + RestDayOTPay + SplHolidayPay + SundayOTPay + LglHolidayPay + _
            OTPay + ExcessSplOTPay + LglHolidayOTPay + ExcessSunLegalPay + _
            ExcessSunLegalPay + NightDiffPay + SalaryAdjust - TardinessDeduction
'        mRange = "AF" & LineCtr
'        xlSheet.Range(mRange) = Format(GrossPay, "0.00")                   ' gross pay
        Month13Pay = BasicPay / 12
'        mRange = "AG" & LineCtr
'        xlSheet.Range(mRange) = Format(Month13Pay, "0.00")                   ' month 13 pay
        
        SSSEmployer = DaysWorked * IIf(IsNull(!SSSRate), 0#, !SSSRate)
'        mRange = "AH" & LineCtr
'        xlSheet.Range(mRange) = Format(SSSEmployer, "0.00")                   ' sss employer
        PHealthEmployer = DaysWorked * IIf(IsNull(!PHRate), 0#, !PHRate)
'        mRange = "AI" & LineCtr
'        xlSheet.Range(mRange) = Format(PHealthEmployer, "0.00")                   ' phealth
        PagibigEmployer = (DaysWorked / DaysWorked) * 50
'        mRange = "AJ" & LineCtr
'        xlSheet.Range(mRange) = Format(PagibigEmployer, "0.00")                   ' pagibig
        
'        mRange = "AK" & LineCtr
'        xlSheet.Range(mRange) = Format(GrossPay + SSSEmployer + PHealthEmployer + PagibigEmployer, "0.00") ' Total pay and benefits
'        AgencyCharge = DaysWorked * AgencyRate
'        mRange = "AL" & LineCtr
'        xlSheet.Range(mRange) = Format(AgencyCharge, "0.00")                   ' agency charge
'        GrossBill = GrossPay + SSSEmployer + PHealthEmployer + PagibigEmployer + AgencyCharge
'        mRange = "AM" & LineCtr
'        xlSheet.Range(mRange) = Format(GrossBill, "0.00")                   ' gross billing
'        VATPayment = (GrossBill * VatRate)
'        mRange = "AN" & LineCtr
'        xlSheet.Range(mRange) = Format(VATPayment, "0.00")                   ' VAT %
'        IncentivePay = !IncentivePay2
'        mRange = "AO" & LineCtr
'        xlSheet.Range(mRange) = Format(IncentivePay, "0.00")                   ' other ded
'        mRange = "AP" & LineCtr
'        xlSheet.Range(mRange) = Format(GrossBill + VATPayment + IncentivePay, "0.00") ' Total Amount Due
'        mRange = "AQ" & LineCtr
'        xlSheet.Range(mRange) = Space(5)
        mRange = "AR" & LineCtr
        xlSheet.Range(mRange) = Format(!GrossPay, "0.00")                   ' original gross pay
        
        mRange = "AU" & LineCtr
        xlSheet.Range(mRange) = Format(!StartPayDate, "mm/dd/yyyy")                   ' original basic pay
        mRange = "AV" & LineCtr
        xlSheet.Range(mRange) = Format(!EndPayDate, "mm/dd/yyyy")                   ' original basic pay
        mRange = "AW" & LineCtr
        xlSheet.Range(mRange) = Format(!SSSRate, "0.00")                   ' sss rate
        mRange = "AX" & LineCtr
        xlSheet.Range(mRange) = Format(!PHRate, "0.00")                   ' phealth rate
        '
        mRange = "AY" & LineCtr
        xlSheet.Range(mRange) = Format(!deptname, "0.00")                   ' Dept Name
        mRange = "AZ" & LineCtr
        xlSheet.Range(mRange) = Format(!BrName, "0.00")                     ' BR Name
        mRange = "BA" & LineCtr
        xlSheet.Range(mRange) = Format(!EmpPosition, "0.00")                ' Employee Position
        mRange = "BB" & LineCtr
        xlSheet.Range(mRange) = Format(!ValeWithClient, "0.00")                   ' Vale with client
        mRange = "BC" & LineCtr
        xlSheet.Range(mRange) = Format(!PlacementFee, "0.00")             ' Placement Fee
        mRange = "BD" & LineCtr
        xlSheet.Range(mRange) = Format(!MiscFee, "0.00")                   ' Placement fee
    
    
    LineCtr = LineCtr + 1
    RecCtr = RecCtr + 1
    i = i + 1

NextRec:

    .MoveNext

Wend
End With

' Print Last record.
Call PrintTotals(xlSheet, xlApp1, LineCtr, k)
sumString = ""
If Len(SaveTotal) = 0 Then
    SaveTotal = LineCtr & ":"
Else
    SaveTotal = SaveTotal & LineCtr & ":"
End If
ColString = SaveTotal
LineCtr = LineCtr + 1

'Compute for Grand Total.
mRange = "B" & LineCtr & ":B" & LineCtr
xlSheet.Range(mRange).Select
xlApp1.ActiveCell.FormulaR1C1 = "  GRAND TOTAL"
xlApp1.Range(mRange).Font.Bold = True
firstRun = True

Dim mGTotal As Double
Dim lf, aa

mGTotal = 0#
'For j = Asc("D") To Asc("Z")
'    If FirstRun = True Then
'        If Asc("H") = j Then
'            j = Asc("D")
'        Else
'            j = Asc("H")
'            FirstRun = False
'        End If
'    End If
    SaveTotal = ColString
    While Len(SaveTotal) <> 0
        lf = InStr(SaveTotal, ":")
        aa = LTrim(RTrim(Left(SaveTotal, lf - 1)))
        If sumString = "" Then ' orig code
            sumString = "=R[" & ((LineCtr - aa) * -1) & "]C"
            'mRange = Chr(j) & aa
            'mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
            'xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
        Else
            sumString = sumString & "+R[" & ((LineCtr - aa) * -1) & "]C"
        End If
        SaveTotal = Mid(SaveTotal, lf + 1, Len(SaveTotal) - lf)
    Wend
    'mRange = Chr(j) & LineCtr
    'xlSheet.Range(mRange).Select
    'xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)" orig code
    'xlApp1.ActiveCell.VALUE = mGTotal
    'ActiveCell.FormulaR1C1 = "=R[-5]C+R[-3]C"
    Call PrintTotals(xlSheet, xlApp1, LineCtr, k, 0, sumString, True)
            
    xlSheet.Range(mRange).Select
'    If xlApp1.ActiveCell.VALUE = 0 Then
'        xlApp1.Selection.EntireColumn.Hidden = True
'    End If
    sumString = ""
    mGTotal = 0#
'Next

mRange = "D" & LineCtr & ":Z" & LineCtr
xlApp1.Range(mRange).Font.Bold = True

'For j = Asc("A") To Asc("R")
    SaveTotal = ColString
    While Len(SaveTotal) <> 0
        lf = InStr(SaveTotal, ":")
        aa = LTrim(RTrim(Left(SaveTotal, lf - 1)))
        If sumString = "" Then
            sumString = "=R[" & ((LineCtr - aa) * -1) & "]C"
            'mRange = "A" & Chr(j) & aa
            'mGTotal = mGTotal + xlSheet.Range(mRange).VALUE
        Else
            sumString = sumString & "+R[" & ((LineCtr - aa) * -1) & "]C"
        End If
        SaveTotal = Mid(SaveTotal, lf + 1, Len(SaveTotal) - lf)
    Wend
    'mRange = "A" & Chr(j) & LineCtr
    'xlSheet.Range(mRange).Select
    'xlApp1.ActiveCell.FormulaR1C1 = sumString
    'xlApp1.ActiveCell.VALUE = mGTotal
    ' Get the Total Amount Due.
'    If Asc("P") = j Then
'        xlApp1.Selection.Copy
'        Call CopyToSheetDataSource(xlSheet, xlApp1, mGTotal)
'    End If
    
    Call PrintTotals(xlSheet, xlApp1, LineCtr, k, 99, sumString, True)
    
    'xlSheet.Range(mRange).Select
'    If xlApp1.ActiveCell.VALUE = 0 Then
'        xlApp1.Selection.EntireColumn.Hidden = True
'    End If
    sumString = ""
    mGTotal = 0#
'Next

mRange = "AA" & LineCtr & ":AR" & LineCtr
xlApp1.Range(mRange).Font.Bold = True

'mRange = "A10:AX" & LineCtr + 2

'mRange = "A" & LineCtr & ":AX" & LineCtr + 2
mRange = "A" & LineCtr & ":BD" & LineCtr + 2
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    '.Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
End With

'Format amount to 0.00.
'mRange = "E10:AR" & LineCtr + 2    ' E7:AP
mRange = "E10:AR" & LineCtr + 2    ' E7:AP

xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
With xlApp1.Selection
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .MergeCells = False
    .Style = "Comma"
End With

'mRange = "A10:AX" & LineCtr + 2
mRange = "A10:BD" & LineCtr + 2
Call PlaceBox(xlSheet, xlApp1, mRange)

mRange = "A1"
xlSheet.Range(mRange).Select

Exit Sub
   
ExitRtn1:

'    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
'    Err.Clear
    
'xlSheet.Name = "Overtime"

End Sub

'Private Sub CarlsonBillingReports2(ByRef xlApp1 As Excel.Application, ByRef xlSheet As Excel.Worksheet)
''Dim xlApp1 As Excel.Application
''Dim xlBook1 As Excel.Workbook
''Dim xlSheet As Excel.Worksheet
'Dim mTotalPay As Double
'Dim mTotalDeduction As Double
'Dim mRange As String
'Dim sumString As String
'Dim ctr As Long
'Dim i As Integer, j As Integer
'Dim cPos As Integer
'Dim EndSw As Boolean
'Dim mHead As String
'
'Dim SubBDO(50) As Double
'Dim SubATM(50) As Double
'Dim SubCash(50) As Double
'Dim SubHold(50) As Double
'Dim bdoctr, atmctr, cashctr, holdctr
'Dim LineCtr As Integer
'Dim bdopos As Integer
'Dim atmpos As Integer
'Dim cashpos As Integer
'Dim holdpos As Integer
'
''Set xlApp1 = New Excel.Application
''Set xlBook1 = xlApp1.Workbooks.Add()
'
'xlApp1.Sheets("15th(Payroll)").Select
'
'xlApp1.WindowState = xlMaximized
'xlApp1.Visible = True
'xlApp1.DisplayAlerts = False
'
''Set xlSheet = book1.Worksheets.Add
'Set xlSheet = xlBook1.Worksheets(1)
'
'cPos = 5
'ctr = 7
'LineCtr = 0
'For i = 1 To 50
'    SubBDO(i) = 0
'    SubATM(i) = 0
'    SubCash(i) = 0
'    SubHold(i) = 0
'    bdoctr = 0
'    atmctr = 0
'    cashctr = 0
'    holdctr = 0
'Next
'i = 1
'
''On Error GoTo Errh
''Print Header
''xlSheet.Range("A2") = mCompanyName & msapi  '" (SAPI)"
'
'' Updated by Bernard May 13, 2008.
'mHead = mCompanyName & " (" & CompanyShort & ")"
''xlSheet.Range("A2") = mCompanyName & " (" & CompanyShort & ")"
'Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 14, True)
'If Format(rs!StartPayDate, "mmm") = Format(rs!EndPayDate, "mmm") Then
'    'xlSheet.Range("A3") = "Payroll Summary Report for Cut-Off " & Format(rs!startPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "dd, yyyy")
'    mHead = "Payroll Summary Report for Cut-Off " & Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "dd, yyyy")
'Else
'    'xlSheet.Range("A3") = "Payroll Summary Report for Cut-Off " & Format(rs!startPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd, yyyy")
'    mHead = "Payroll Summary Report for Cut-Off " & Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd, yyyy")
'End If
'Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, False)
'
''xlSheet.Range("AO3") = mSeriesNo
'Call PrintReportHeader(xlSheet, xlApp1, "AS3", mSeriesNo, "", xlLeft, , False, False, False, , , 8, False)
''xlSheet.Range("AO4") = Format(Now, "mm/dd/yyyy")
'Call PrintReportHeader(xlSheet, xlApp1, "AS4", Format(Now, "mm/dd/yyyy"), "", xlLeft, xlBottom, False, False, False, , , 8, False)
'
'Do While Not rs.EOF
'    ' DoEvents
'    xlSheet.Cells(ctr, 1).VALUE = rs!ControlNum
'    xlSheet.Cells(ctr, 2).VALUE = rs!EmployeeName
'    xlSheet.Cells(ctr, 3).VALUE = rs!EmpTaxCode
'    xlSheet.Cells(ctr, 4).VALUE = rs!DailyRate
'    xlSheet.Cells(ctr, 5).VALUE = rs!DaysPresent
'
'    If rs!PaymentSystem = "BDO" Then
'        bdoctr = bdoctr + 1
'        SubBDO(1) = SubBDO(1) + rs!DaysPresent
'        SubBDO(2) = SubBDO(2) + rs!BasicPay
'        SubBDO(3) = SubBDO(3) + rs!AbsentDeduction
'        SubBDO(4) = SubBDO(4) + rs!HoursTardiness   ' Add
'        SubBDO(5) = SubBDO(5) + rs!TardinessDeduction
'        SubBDO(6) = SubBDO(6) + rs!OTHours
'        SubBDO(7) = SubBDO(7) + rs!OTPay
'        SubBDO(8) = SubBDO(8) + rs!NightDiffOT
'        SubBDO(9) = SubBDO(9) + rs!NightDiffPay
'        SubBDO(10) = SubBDO(10) + rs!RestDayOTHrs
'        SubBDO(11) = SubBDO(11) + rs!RestDayOTPay
'        SubBDO(12) = SubBDO(12) + rs!SundayOT
'        SubBDO(13) = SubBDO(13) + rs!SundayOTPay
'
'        SubBDO(14) = SubBDO(14) + rs!NoSplHoliday
'        SubBDO(15) = SubBDO(15) + rs!SplHolidayPay
'        SubBDO(16) = SubBDO(16) + rs!NoLglHoliday
'        SubBDO(17) = SubBDO(17) + rs!LglHolidayPay
''        SubBDO(16) = SubBDO(16) + rs!NoSickLeave
''        SubBDO(18) = SubBDO(18) + rs!SickLeavePay
''        SubBDO(19) = SubBDO(19) + rs!NoVacationLeave
''        SubBDO(20) = SubBDO(20) + rs!VacationLeavePay
'        SubBDO(18) = SubBDO(18) + rs!SalaryAdjust
'        SubBDO(19) = SubBDO(19) + rs!IncentivePay
'        SubBDO(20) = SubBDO(20) + rs!EcolaPay
'        SubBDO(21) = SubBDO(21) + rs!GrossPay
'        SubBDO(22) = SubBDO(22) + rs!WTaxAmount
'
'        SubBDO(23) = SubBDO(23) + rs!SSSEmployee
'        SubBDO(24) = SubBDO(24) + rs!PHealthEmployee
'        SubBDO(25) = SubBDO(25) + rs!PagibigEmployee
'        SubBDO(26) = SubBDO(26) + rs!CashBond
'        SubBDO(27) = SubBDO(27) + rs!PlacementFee
'        SubBDO(28) = SubBDO(28) + rs!MiscFee
'        SubBDO(29) = SubBDO(29) + rs!CardTransFee
'
'        SubBDO(30) = SubBDO(30) + rs!SalaryLoan
'        SubBDO(31) = SubBDO(31) + rs!Vale_IOU
'        SubBDO(32) = SubBDO(32) + rs!Uniform
'        SubBDO(33) = SubBDO(33) + rs!Medical
'        SubBDO(34) = SubBDO(34) + rs!Variances
'        SubBDO(35) = SubBDO(35) + rs!OverUnderPay
'
'        SubBDO(36) = SubBDO(36) + rs!SSSLoan
'        SubBDO(37) = SubBDO(37) + rs!PagibigLoan
'        SubBDO(38) = SubBDO(38) + rs!ValeWithClient
'        SubBDO(39) = SubBDO(39) + rs!Deductions
'        SubBDO(40) = SubBDO(40) + rs!NetPay
'        'SubBDO(33) = SubBDO(33) + rs!AcctNo    Account No
'        SubBDO(42) = SubBDO(42) + rs!Month13Payment
'        If rs!SSSEmployer <> 0 Then
'            SubBDO(43) = SubBDO(43) + rs!SSSEmployer - 10
'            SubBDO(44) = SubBDO(44) + 10
'        Else
'            SubBDO(43) = SubBDO(43) + 0
'            SubBDO(44) = SubBDO(44) + 0
'        End If
'        SubBDO(45) = SubBDO(45) + rs!PHealthEmployer
'        SubBDO(46) = SubBDO(46) + rs!PagibigEmployer
'        SubBDO(47) = SubBDO(47) + rs!AgencyCharge
'    ElseIf rs!PaymentSystem = "ATM" Then
'        atmctr = atmctr + 1
'        SubATM(1) = SubATM(1) + rs!DaysPresent
'        SubATM(2) = SubATM(2) + rs!BasicPay
'        SubATM(3) = SubATM(3) + rs!AbsentDeduction
'        SubATM(4) = SubATM(4) + rs!HoursTardiness   ' Add
'        SubATM(5) = SubATM(5) + rs!TardinessDeduction
'        SubATM(6) = SubATM(6) + rs!OTHours
'        SubATM(7) = SubATM(7) + rs!OTPay
'        SubATM(8) = SubATM(8) + rs!NightDiffOT
'        SubATM(9) = SubATM(9) + rs!NightDiffPay
'        SubATM(10) = SubATM(10) + rs!RestDayOTHrs
'        SubATM(11) = SubATM(11) + rs!RestDayOTPay
'        SubATM(12) = SubATM(12) + rs!SundayOT
'        SubATM(13) = SubATM(13) + rs!SundayOTPay
'
'        SubATM(14) = SubATM(14) + rs!NoSplHoliday
'        SubATM(15) = SubATM(15) + rs!SplHolidayPay
'        SubATM(16) = SubATM(16) + rs!NoLglHoliday
'        SubATM(17) = SubATM(17) + rs!LglHolidayPay
''        subatm(16) = subatm(16) + rs!NoSickLeave
''        subatm(18) = subatm(18) + rs!SickLeavePay
''        subatm(19) = subatm(19) + rs!NoVacationLeave
''        subatm(20) = subatm(20) + rs!VacationLeavePay
'        SubATM(18) = SubATM(18) + rs!SalaryAdjust
'        SubATM(19) = SubATM(19) + rs!IncentivePay
'        SubATM(20) = SubATM(20) + rs!EcolaPay
'        SubATM(21) = SubATM(21) + rs!GrossPay
'        SubATM(22) = SubATM(22) + rs!WTaxAmount
'
'        SubATM(23) = SubATM(23) + rs!SSSEmployee
'        SubATM(24) = SubATM(24) + rs!PHealthEmployee
'        SubATM(25) = SubATM(25) + rs!PagibigEmployee
'        SubATM(26) = SubATM(26) + rs!CashBond
'        SubATM(27) = SubATM(27) + rs!PlacementFee
'        SubATM(28) = SubATM(28) + rs!MiscFee
'        SubATM(29) = SubATM(29) + rs!CardTransFee
'
'        SubATM(30) = SubATM(30) + rs!SalaryLoan
'        SubATM(31) = SubATM(31) + rs!Vale_IOU
'        SubATM(32) = SubATM(32) + rs!Uniform
'        SubATM(33) = SubATM(33) + rs!Medical
'        SubATM(34) = SubATM(34) + rs!Variances
'        SubATM(35) = SubATM(35) + rs!OverUnderPay
'
'        SubATM(36) = SubATM(36) + rs!SSSLoan
'        SubATM(37) = SubATM(37) + rs!PagibigLoan
'        SubATM(38) = SubATM(38) + rs!ValeWithClient
'        SubATM(39) = SubATM(39) + rs!Deductions
'        SubATM(40) = SubATM(40) + rs!NetPay
'        'subatm(33) = subatm(33) + rs!AcctNo    Account No
'        SubATM(42) = SubATM(42) + rs!Month13Payment
'        If rs!SSSEmployer <> 0 Then
'            SubATM(43) = SubATM(43) + rs!SSSEmployer - 10
'            SubATM(44) = SubATM(44) + 10
'        Else
'            SubATM(43) = SubATM(43) + 0
'            SubATM(44) = SubATM(44) + 0
'        End If
'        SubATM(45) = SubATM(45) + rs!PHealthEmployer
'        SubATM(46) = SubATM(46) + rs!PagibigEmployer
'        SubATM(47) = SubATM(47) + rs!AgencyCharge
'    ElseIf rs!PaymentSystem = "Cash" Then
'        cashctr = cashctr + 1
'        SubCash(1) = SubCash(1) + rs!DaysPresent
'        SubCash(2) = SubCash(2) + rs!BasicPay
'        SubCash(3) = SubCash(3) + rs!AbsentDeduction
'        SubCash(4) = SubCash(4) + rs!HoursTardiness   ' Add
'        SubCash(5) = SubCash(5) + rs!TardinessDeduction
'        SubCash(6) = SubCash(6) + rs!OTHours
'        SubCash(7) = SubCash(7) + rs!OTPay
'        SubCash(8) = SubCash(8) + rs!NightDiffOT
'        SubCash(9) = SubCash(9) + rs!NightDiffPay
'        SubCash(10) = SubCash(10) + rs!RestDayOTHrs
'        SubCash(11) = SubCash(11) + rs!RestDayOTPay
'        SubCash(12) = SubCash(12) + rs!SundayOT
'        SubCash(13) = SubCash(13) + rs!SundayOTPay
'
'        SubCash(14) = SubCash(14) + rs!NoSplHoliday
'        SubCash(15) = SubCash(15) + rs!SplHolidayPay
'        SubCash(16) = SubCash(16) + rs!NoLglHoliday
'        SubCash(17) = SubCash(17) + rs!LglHolidayPay
''        subcash(16) = subcash(16) + rs!NoSickLeave
''        subcash(18) = subcash(18) + rs!SickLeavePay
''        subcash(19) = subcash(19) + rs!NoVacationLeave
''        subcash(20) = subcash(20) + rs!VacationLeavePay
'        SubCash(18) = SubCash(18) + rs!SalaryAdjust
'        SubCash(19) = SubCash(19) + rs!IncentivePay
'        SubCash(20) = SubCash(20) + rs!EcolaPay
'        SubCash(21) = SubCash(21) + rs!GrossPay
'        SubCash(22) = SubCash(22) + rs!WTaxAmount
'
'        SubCash(23) = SubCash(23) + rs!SSSEmployee
'        SubCash(24) = SubCash(24) + rs!PHealthEmployee
'        SubCash(25) = SubCash(25) + rs!PagibigEmployee
'        SubCash(26) = SubCash(26) + rs!CashBond
'        SubCash(27) = SubCash(27) + rs!PlacementFee
'        SubCash(28) = SubCash(28) + rs!MiscFee
'        SubCash(29) = SubCash(29) + rs!CardTransFee
'
'        SubCash(30) = SubCash(30) + rs!SalaryLoan
'        SubCash(31) = SubCash(31) + rs!Vale_IOU
'        SubCash(32) = SubCash(32) + rs!Uniform
'        SubCash(33) = SubCash(33) + rs!Medical
'        SubCash(34) = SubCash(34) + rs!Variances
'        SubCash(35) = SubCash(35) + rs!OverUnderPay
'
'        SubCash(36) = SubCash(36) + rs!SSSLoan
'        SubCash(37) = SubCash(37) + rs!PagibigLoan
'        SubCash(38) = SubCash(38) + rs!ValeWithClient
'        SubCash(39) = SubCash(39) + rs!Deductions
'        SubCash(40) = SubCash(40) + rs!NetPay
'        'subcash(33) = subcash(33) + rs!AcctNo    Account No
'        SubCash(42) = SubCash(42) + rs!Month13Payment
'        If rs!SSSEmployer <> 0 Then
'            SubCash(43) = SubCash(43) + rs!SSSEmployer - 10
'            SubCash(44) = SubCash(44) + 10
'        Else
'            SubCash(43) = SubCash(43) + 0
'            SubCash(44) = SubCash(44) + 0
'        End If
'        SubCash(45) = SubCash(45) + rs!PHealthEmployer
'        SubCash(46) = SubCash(46) + rs!PagibigEmployer
'        SubCash(47) = SubCash(47) + rs!AgencyCharge
'    Else 'If rs!PaymentSystem = "HOLD" Then
'        holdctr = holdctr + 1
'        SubHold(1) = SubHold(1) + rs!DaysPresent
'        SubHold(2) = SubHold(2) + rs!BasicPay
'        SubHold(3) = SubHold(3) + rs!AbsentDeduction
'        SubHold(4) = SubHold(4) + rs!HoursTardiness   ' Add
'        SubHold(5) = SubHold(5) + rs!TardinessDeduction
'        SubHold(6) = SubHold(6) + rs!OTHours
'        SubHold(7) = SubHold(7) + rs!OTPay
'        SubHold(8) = SubHold(8) + rs!NightDiffOT
'        SubHold(9) = SubHold(9) + rs!NightDiffPay
'        SubHold(10) = SubHold(10) + rs!RestDayOTHrs
'        SubHold(11) = SubHold(11) + rs!RestDayOTPay
'        SubHold(12) = SubHold(12) + rs!SundayOT
'        SubHold(13) = SubHold(13) + rs!SundayOTPay
'
'        SubHold(14) = SubHold(14) + rs!NoSplHoliday
'        SubHold(15) = SubHold(15) + rs!SplHolidayPay
'        SubHold(16) = SubHold(16) + rs!NoLglHoliday
'        SubHold(17) = SubHold(17) + rs!LglHolidayPay
''        subhold(16) = subhold(16) + rs!NoSickLeave
''        subhold(18) = subhold(18) + rs!SickLeavePay
''        subhold(19) = subhold(19) + rs!NoVacationLeave
''        subhold(20) = subhold(20) + rs!VacationLeavePay
'        SubHold(18) = SubHold(18) + rs!SalaryAdjust
'        SubHold(19) = SubHold(19) + rs!IncentivePay
'        SubHold(20) = SubHold(20) + rs!EcolaPay
'        SubHold(21) = SubHold(21) + rs!GrossPay
'        SubHold(22) = SubHold(22) + rs!WTaxAmount
'
'        SubHold(23) = SubHold(23) + rs!SSSEmployee
'        SubHold(24) = SubHold(24) + rs!PHealthEmployee
'        SubHold(25) = SubHold(25) + rs!PagibigEmployee
'        SubHold(26) = SubHold(26) + rs!CashBond
'        SubHold(27) = SubHold(27) + rs!PlacementFee
'        SubHold(28) = SubHold(28) + rs!MiscFee
'        SubHold(29) = SubHold(29) + rs!CardTransFee
'
'        SubHold(30) = SubHold(30) + rs!SalaryLoan
'        SubHold(31) = SubHold(31) + rs!Vale_IOU
'        SubHold(32) = SubHold(32) + rs!Uniform
'        SubHold(33) = SubHold(33) + rs!Medical
'        SubHold(34) = SubHold(34) + rs!Variances
'        SubHold(35) = SubHold(35) + rs!OverUnderPay
'
'        SubHold(36) = SubHold(36) + rs!SSSLoan
'        SubHold(37) = SubHold(37) + rs!PagibigLoan
'        SubHold(38) = SubHold(38) + rs!ValeWithClient
'        SubHold(39) = SubHold(39) + rs!Deductions
'        SubHold(40) = SubHold(40) + rs!NetPay
'        'subhold(33) = subhold(33) + rs!AcctNo    Account No
'        SubHold(42) = SubHold(42) + rs!Month13Payment
'        If rs!SSSEmployer <> 0 Then
'            SubHold(43) = SubHold(43) + rs!SSSEmployer - 10
'            SubHold(44) = SubHold(44) + 10
'        Else
'            SubHold(43) = SubHold(43) + 0
'            SubHold(44) = SubHold(44) + 0
'        End If
'        SubHold(45) = SubHold(45) + rs!PHealthEmployer
'        SubHold(46) = SubHold(46) + rs!PagibigEmployer
'        SubHold(47) = SubHold(47) + rs!AgencyCharge
'    End If
'    xlSheet.Cells(ctr, 6).VALUE = Format(rs!BasicPay, "##,##0.00")
'    xlSheet.Cells(ctr, 7).VALUE = Format(rs!AbsentDeduction, "##,#0.00")
'    xlSheet.Cells(ctr, 8).VALUE = Format(rs!HoursTardiness, "##,#0.00")
'    xlSheet.Cells(ctr, 9).VALUE = Format(rs!TardinessDeduction, "##,#0.00")
'
'    xlSheet.Cells(ctr, 10).VALUE = Format(rs!OTHours, "##,##0.00")
'    xlSheet.Cells(ctr, 11).VALUE = Format(rs!OTPay, "##,##0.00")
'    xlSheet.Cells(ctr, 12).VALUE = Format(rs!NightDiffOT, "##,##0.00")
'    xlSheet.Cells(ctr, 13).VALUE = Format(rs!NightDiffPay, "##,##0.00")
'
'    xlSheet.Cells(ctr, 14).VALUE = Format(rs!RestDayOTHrs, "##,##0.00")
'    xlSheet.Cells(ctr, 15).VALUE = Format(rs!RestDayOTPay, "##,##0.00")
'
'    xlSheet.Cells(ctr, 16).VALUE = Format(rs!SundayOT, "##,##0.00")
'    xlSheet.Cells(ctr, 17).VALUE = Format(rs!SundayOTPay, "##,##0.00")
'
'    xlSheet.Cells(ctr, 18).VALUE = Format(rs!NoSplHoliday, "##,##0.00")
'    xlSheet.Cells(ctr, 19).VALUE = Format(rs!SplHolidayPay, "##,##0.00")
'    xlSheet.Cells(ctr, 20).VALUE = Format(rs!NoLglHoliday, "##,##0.00")
'    xlSheet.Cells(ctr, 21).VALUE = Format(rs!LglHolidayPay, "##,##0.00")
'
''    xlSheet.Cells(ctr, 21).VALUE = Format(rs!NoSickLeave, "##,##0.00")
''    xlSheet.Cells(ctr, 22).VALUE = Format(rs!SickLeavePay, "##,##0.00")
''    xlSheet.Cells(ctr, 23).VALUE = Format(rs!NoVacationLeave, "##,##0.00")
''    xlSheet.Cells(ctr, 24).VALUE = Format(rs!VacationLeavePay, "##,##0.00")
'
'    xlSheet.Cells(ctr, 22).VALUE = Format(rs!SalaryAdjust, "##,##0.00")
'    xlSheet.Cells(ctr, 23).VALUE = Format(rs!IncentivePay, "##,##0.00")
'    xlSheet.Cells(ctr, 24).VALUE = Format(rs!EcolaPay, "##,##0.00")
'
'    xlSheet.Cells(ctr, 25).VALUE = Format(rs!GrossPay, "##,##0.00")
'    xlSheet.Cells(ctr, 26).VALUE = Format(rs!WTaxAmount, "##,##0.00")
'    xlSheet.Cells(ctr, 27).VALUE = Format(rs!SSSEmployee, "##,##0.00")
'    xlSheet.Cells(ctr, 28).VALUE = Format(rs!PHealthEmployee, "##,##0.00")
'    xlSheet.Cells(ctr, 29).VALUE = Format(rs!PagibigEmployee, "##,##0.00")
'
'    xlSheet.Cells(ctr, 30).VALUE = Format(rs!CashBond, "##,##0.00")
'    xlSheet.Cells(ctr, 31).VALUE = Format(rs!PlacementFee, "##,##0.00")
'    xlSheet.Cells(ctr, 32).VALUE = Format(rs!MiscFee, "##,##0.00")
'    xlSheet.Cells(ctr, 33).VALUE = Format(rs!CardTransFee, "##,##0.00")
'
'    xlSheet.Cells(ctr, 34).VALUE = Format(rs!SalaryLoan, "##,##0.00")
'    xlSheet.Cells(ctr, 35).VALUE = Format(rs!Vale_IOU, "##,##0.00")
'    xlSheet.Cells(ctr, 36).VALUE = Format(rs!Uniform, "##,##0.00")
'    xlSheet.Cells(ctr, 37).VALUE = Format(rs!Medical, "##,##0.00")
'    xlSheet.Cells(ctr, 38).VALUE = Format(rs!Variances, "##,##0.00")
'    xlSheet.Cells(ctr, 39).VALUE = Format(rs!OverUnderPay, "##,##0.00")
'
'    xlSheet.Cells(ctr, 40).VALUE = Format(rs!SSSLoan, "##,##0.00")
'    xlSheet.Cells(ctr, 41).VALUE = Format(rs!PagibigLoan, "##,##0.00")
'    xlSheet.Cells(ctr, 42).VALUE = Format(rs!ValeWithClient, "##,##0.00")
'    xlSheet.Cells(ctr, 43).VALUE = Format(rs!Deductions, "##,##0.00")
'    xlSheet.Cells(ctr, 44).VALUE = Format(rs!NetPay, "##,##0.00")
'    xlSheet.Cells(ctr, 45).VALUE = IIf(rs!AcctNo = "", "n/a", rs!AcctNo)
'
'    xlSheet.Cells(ctr, 46).VALUE = Format(rs!Month13Payment, "##,##0.00")
'    If rs!SSSEmployer <> 0 Then
'        xlSheet.Cells(ctr, 47).VALUE = Format(rs!SSSEmployer - 10, "##,##0.00")
'        xlSheet.Cells(ctr, 48).VALUE = Format(10, "##,##0.00")
'    Else
'        xlSheet.Cells(ctr, 47).VALUE = Format(0, "##,##0.00")
'        xlSheet.Cells(ctr, 48).VALUE = Format(0, "##,##0.00")
'    End If
'    xlSheet.Cells(ctr, 49).VALUE = Format(rs!PHealthEmployer, "##,##0.00")
'    xlSheet.Cells(ctr, 50).VALUE = Format(rs!PagibigEmployer, "##,##0.00")
'    xlSheet.Cells(ctr, 51).VALUE = Format(rs!AgencyCharge, "##,##0.00")
'    xlSheet.Cells(ctr, 52).FormulaR1C1 = "=RC[-21]+RC[-5]+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
'    xlSheet.Cells(ctr, 53).FormulaR1C1 = "=RC[-2]+RC[-1]"
'
'    'xlsheetActiveCell.FormulaR1C1 = "=RC[-21]+RC[-5]+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
'
'    'xlSheet.Cells.EntireColumn.AutoFit
'    'xlSheet.Cells.EntireRow.Font.Name = "Arial"
'    'xlSheet.Cells.EntireRow.Font.Size = 8
'
'    i = i + 1
'    ctr = ctr + 1
'    rs.MoveNext
'
'Loop
'End Sub

Private Sub CarlsonBillingReports(ByRef xlApp1 As Excel.Application, ByRef xlSheet As Excel.Worksheet)
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String, mRange1 As String
Dim SaveTotal As String
Dim ColString As String
Dim sumString As String

Dim ctr As Long
Dim i As Integer, j As Integer, k As Integer
Dim cPos As Integer
Dim EndSw As Boolean
Dim mHead As String
Dim FirstHalf As Boolean

Dim LineCtr As Integer, LineCtr2 As Integer, LineCtr3 As Integer
Dim RecCtr As Integer, RecCtr1 As Integer, RecCtr2 As Integer, RecCtr3 As Integer, RecCtr4 As Integer, RecCtr5 As Integer
Dim prevCutOff As Variant
Dim prevBranch As String
Dim PrevRate As Double
Dim firstRun As Boolean
Dim BillFlag15 As Boolean
Dim PayrollFlag As Boolean

Dim HeadStart As Integer
Dim AceStart As Integer
Dim HandyStart As Integer
Dim WorkShopStart As Integer
Dim OthersStart As Integer
Dim HeadStartBill As Integer
Dim AceStartBill As Integer
Dim HandyStartBill As Integer
Dim WorkShopStartBill As Integer
Dim OthersStartBill As Integer
Dim Headsw As Boolean
Dim Acesw As Boolean
Dim Handysw As Boolean
Dim Worksw As Boolean
Dim OtherSw As Boolean
Dim mBranchCD As Integer
Dim AreaCodeSw As Boolean
Dim StartPos As Integer
Dim mGrossPay As Double

Headsw = False: Acesw = False: Handysw = False: Worksw = False: OtherSw = False

HeadStart = 8
AceStart = 10
HandyStart = 12
WorkShopStart = 14
OthersStart = 16
StartPos = 8

HeadStartBill = 14
AceStartBill = 24
HandyStartBill = 29
WorkShopStartBill = 34
OthersStartBill = 39

'AreaCode AreaName
'1       Head Office
'2       Ace
'3       Handyman
'4       Workshop
'5       Others

RecCtr1 = 0: RecCtr2 = 0: RecCtr3 = 0: RecCtr4 = 0: RecCtr5 = 0
AreaCodeSw = False
BillFlag15 = True
PayrollFlag = True
firstRun = True
FirstHalf = True
SaveTotal = ""
ColString = ""
cPos = 10
RecCtr = 1
ctr = 7
LineCtr = 8 ' 12
LineCtr2 = 8    ' Payroll
LineCtr3 = 14   ' Billing
i = 1
k = 12

'xlApp1.Sheets("Billing Report").Select
xlApp1.Sheets("15th(Payroll)").Select

On Error GoTo ExitRtn3

If rsBilledTrans2.RecordCount = 0 Then
    MsgBox "No records to process. Please recompute Transaction Records again.", vbInformation, PayrollAlert
    GoTo ExitRtn3
End If

With rsBilledTrans2
    .MoveFirst
    
'prevCutOff = Format(!StartPayDate, "mmddyyyy") & LTrim(RTrim(!Paytext2))
'FirstCutoff1 = Format(!StartPayDate, "mm/dd/yyyy")
'FirstCutoff2 = Format(!EndPayDate, "mm/dd/yyyy")
SecondCutoff1 = #1/1/1900#
SecondCutoff2 = #1/1/1900#
'prevCutOff = Format(!StartPayDate, "mmddyyyy")
prevCutOff = Format(!CutOffOrder, "mmddyyyy")

prevBranch = IIf(IsNull(!BrName), "No Area Specified", LTrim(RTrim(!BrName)))
PrevRate = IIf(IsNull(!DailyRate), 0, !DailyRate)

While Not (.EOF Or .BOF)

    If chkExcludedRecords.VALUE = False Then
        If !ExcludedFlag = True Then
            GoTo NextRec
        End If
    End If
    
    ' Get the branch code.
    mBranchCD = IIf(IsNull(!BrName), GetBranchCode(!Employees, 0), !areacode)
    
    If FirstCutoff1 = Format(!CutOffOrder, "mm/dd/yyyy") Then
        xlApp1.Sheets("15th(Payroll)").Select
        FirstHalf = True
        BillFlag15 = True
        
'        If FirstRun = False Then
'            LineCtr = LineCtr + 1
'            RecCtr = RecCtr + 1
'            Select Case mBranchCD '!AreaCode
'                Case 1      ' Head Office
'                    RecCtr1 = RecCtr1 + 1
'                Case 2      ' Ace
'                    RecCtr2 = RecCtr2 + 1
'                Case 3      ' Handyman
'                    RecCtr3 = RecCtr3 + 1
'                Case 4      ' Workshop
'                    RecCtr4 = RecCtr4 + 1
'                Case Else   ' Others
'                    RecCtr5 = RecCtr5 + 1
'            End Select
'        End If
        'FirstRun = False
    Else    ' Process second half
        If firstRun = False And FirstHalf = True Then
            HeadStart = 8
            AceStart = 10
            HandyStart = 12
            WorkShopStart = 14
            OthersStart = 16
            
            HeadStartBill = 14
            AceStartBill = 24
            HandyStartBill = 29
            WorkShopStartBill = 34
            OthersStartBill = 39
            RecCtr = 1
            LineCtr = 8
            LineCtr2 = 8    ' Payroll
            LineCtr3 = 14   ' Billing
            RecCtr1 = 0: RecCtr2 = 0: RecCtr3 = 0: RecCtr4 = 0: RecCtr5 = 0
            firstRun = True
        End If
        SecondCutoff1 = IIf(IsNull(!StartPayDate), #1/1/1900#, Format(!StartPayDate, "mm/dd/yyyy"))
        SecondCutoff2 = IIf(IsNull(!EndPayDate), #1/1/1900#, Format(!EndPayDate, "mm/dd/yyyy"))
'        If FirstRun = False Then
'            LineCtr = LineCtr + 1
'            RecCtr = RecCtr + 1
'            Select Case mBranchCD '!AreaCode
'                Case 1      ' Head Office
'                    RecCtr1 = RecCtr1 + 1
'                Case 2      ' Ace
'                    RecCtr2 = RecCtr2 + 1
'                Case 3      ' Handyman
'                    RecCtr3 = RecCtr3 + 1
'                Case 4      ' Workshop
'                    RecCtr4 = RecCtr4 + 1
'                Case Else   ' Others
'                    RecCtr5 = RecCtr5 + 1
'            End Select
'        End If
        xlApp1.Sheets("30th(Payroll)").Select
        If FirstHalf = False Then
            firstRun = False
        End If
        BillFlag15 = False
        FirstHalf = False
    End If
    
    ' Get the Target location for copy.
    Select Case mBranchCD '!AreaCode
        Case 1  ' Head Office
            LineCtr2 = RecCtr1 + HeadStart
            StartPos = HeadStart
            If RecCtr1 = 1 And firstRun = False Then
                AreaCodeSw = True
            Else
                AreaCodeSw = False
            End If
        Case 2  ' Ace
            LineCtr2 = RecCtr2 + AceStart
            StartPos = AceStart
            If RecCtr2 = 1 And firstRun = False Then
                AreaCodeSw = True
            Else
                AreaCodeSw = False
            End If
        Case 3  ' Handyman
            LineCtr2 = RecCtr3 + HandyStart
            StartPos = HandyStart
            If RecCtr3 = 1 And firstRun = False Then
                AreaCodeSw = True
            Else
                AreaCodeSw = False
            End If
        Case 4      ' Workshop
            LineCtr2 = RecCtr4 + WorkShopStart
            StartPos = WorkShopStart
            If RecCtr4 = 1 And firstRun = False Then
                AreaCodeSw = True
            Else
                AreaCodeSw = False
            End If
        Case Else  ' Others
            LineCtr2 = RecCtr5 + OthersStart
            StartPos = OthersStart
            If RecCtr5 = 1 And firstRun = False Then
                AreaCodeSw = True
            Else
                AreaCodeSw = False
            End If
    End Select
    
    'If LineCtr > 8 And FirstRun = False And Headsw = False And Acesw = False And Handysw = False And Worksw = False And OtherSw = False Then    ' Paste the template
    'If LineCtr > 8 And FirstRun = False Then      ' Paste the template
    If StartPos <> LineCtr2 Then      ' Paste the template
        'mRange = "A8:AW8"
        mRange = "A" & LineCtr2 - 1 & ":AW" & LineCtr2 - 1
        'xlSheet.Range(mRange).Select
        xlApp1.ActiveSheet.Range(mRange).Select
        xlApp1.Selection.Copy
        
        mRange = "A" & LineCtr2
        'mRange = "A" & LineCtr
        xlApp1.ActiveSheet.Range(mRange).Select
        'xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone , SkipBlanks:=
            'False, Transpose:=False
        xlApp1.Selection.Insert Shift:=xlDown
        'Rows("1:1").Select
        'Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

        If BillFlag15 = True Then
            xlApp1.Sheets("15th Billing").Select
        Else
            xlApp1.Sheets("30th Billing").Select
        End If
        
        ' Get the Target location for copy.
        Select Case mBranchCD '!AreaCode
            Case 1      ' Head Office
                LineCtr3 = RecCtr1 + HeadStartBill
                'RecCtr1 = RecCtr1 + 1
            Case 2      ' Ace
                LineCtr3 = RecCtr2 + AceStartBill
                'RecCtr2 = RecCtr2 + 1
            Case 3      ' Handyman
                LineCtr3 = RecCtr3 + HandyStartBill
                'RecCtr3 = RecCtr3 + 1
            Case 4      ' Workshop
                LineCtr3 = RecCtr4 + WorkShopStartBill
                'RecCtr4 = RecCtr4 + 1
            Case Else   ' Others
                LineCtr3 = RecCtr5 + OthersStartBill
                'RecCtr5 = RecCtr5 + 1
        End Select
        
        'mRange = "B14:AX14"
        mRange = "A" & LineCtr3 - 1 & ":AX" & LineCtr3 - 1
        xlApp1.ActiveSheet.Range(mRange).Select
        xlApp1.Selection.Copy
        mRange = "A" & LineCtr3
        'mRange = "B" & LineCtr15Bil
        xlApp1.ActiveSheet.Range(mRange).Select
        xlApp1.Selection.Insert Shift:=xlDown
        
        If BillFlag15 = False Then
        ' Get the equivalent SSS rate.
        mRange = "AK" & LineCtr3
        xlApp1.ActiveSheet.Range(mRange).Select
        mGrossPay = xlApp1.ActiveSheet.Range(mRange).VALUE
        
        mRange = "AJ" & LineCtr3 ' SSS
        xlApp1.ActiveSheet.Range(mRange).Select
        xlApp1.ActiveSheet.Range(mRange).VALUE = GetMandatoryRates(xlApp1, xlSheet, mGrossPay, True)
        
        ' Get the equivalent phealth rate.
        mRange = "AI" & LineCtr3 ' phealth
        xlApp1.ActiveSheet.Range(mRange).Select
        xlApp1.ActiveSheet.Range(mRange).VALUE = GetMandatoryRates(xlApp1, xlSheet, mGrossPay, False)
        
        End If
        
        If BillFlag15 = True Then
            xlApp1.Sheets("15th(Payroll)").Select
        Else
            xlApp1.Sheets("30th(Payroll)").Select
        End If

    End If
    
    firstRun = False
    
        mRange = "A" & LineCtr2
        'xlApp1.ActiveSheet.Range(mRange).VALUE = RecCtr
        'xlSheet.Range(mRange) = RecCtr  'IIf(IsNull(!ControlNo), "", !ControlNo)
        mRange = "B" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = CompanyShort
        'xlSheet.Range(mRange) = CompanyShort  'IIf(IsNull(!ControlNo), "", !ControlNo)
        
        mRange = "C" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(IsNull(!BrName), GetBranchCode(!Employees, 1), !BrName)
        'xlSheet.Range(mRange) = IIf(IsNull(!Brname), "No Area Specified", !Brname)
        
'        mRange = "D" & linectr2
'        xlSheet.Range(mRange) = !deptname                   ' Dept Name
''        mRange = "AZ" & linectr2
''        xlSheet.Range(mRange) = Format(!brname, "0.00")                     ' BR Name
''        mRange = "BA" & linectr2
''        xlSheet.Range(mRange) = Format(!EmpPosition, "0.00")                ' Employee Position
        
        mRange = "D" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(IsNull(!Employees), "No Name", !Employees)
        ' Highlight no payroll.
        If IsNull(!EmployeeName) = True Then
'            xlApp1.ActiveCell.Characters(1, 50).ColorIndex = 36
            With xlApp1.Selection.Interior
                .ColorIndex = 36
                .Pattern = xlSolid
            End With
        Else
            'xlApp1.ActiveCell.Characters(1, 50).ColorIndex = xlNone
            With xlApp1.Selection.Interior
                .ColorIndex = xlNone
                .Pattern = xlSolid
            End With
        End If
        
        'xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
        'xlSheet.Range(mRange) = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
        mRange = "E" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(IsNull(!EmpTaxCode), "X", !EmpTaxCode)
        'xlSheet.Range(mRange) = IIf(IsNull(!EmpTaxCode), "X", !EmpTaxCode)
        mRange = "F" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!DailyRate, "0.00")
        'xlSheet.Range(mRange) = Format(!DailyRate, "0.00")
        mRange = "G" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!DaysPresent2, "0.00")
        'xlSheet.Range(mRange) = Format(!DaysPresent2, "0.00")
        mRange = "H" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!BasicPay, "0.00")
        'xlSheet.Range(mRange) = Format(!BasicPay, "0.00")
        mRange = "I" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!AbsentDeduction, "0.00")
        'xlSheet.Range(mRange) = Format(!AbsentDeduction, "0.00")
        mRange = "J" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!HoursTardiness2, "0.00")
        'xlSheet.Range(mRange) = Format(!HoursTardiness2, "0.00")
        mRange = "K" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!TardinessDeduction, "0.00")
        'xlSheet.Range(mRange) = Format(!TardinessDeduction, "0.00")
        mRange = "L" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!OTHours, "0.00")
        'xlSheet.Range(mRange) = Format(!OTHours, "0.00")
        mRange = "M" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!OTPay, "0.00")
        'xlSheet.Range(mRange) = Format(!OTPay, "0.00")
        mRange = "N" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!NightDiffOT, "0.00")
        'xlSheet.Range(mRange) = Format(!NightDiffOT, "0.00")
        mRange = "O" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!NightDiffPay, "0.00")
        'xlSheet.Range(mRange) = Format(!NightDiffPay, "0.00")
        mRange = "P" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!RestDayOTHrs, "0.00")
        'xlSheet.Range(mRange) = Format(!RestDayOTHrs, "0.00")
        mRange = "Q" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!RestDayOTPay, "0.00")
        'xlSheet.Range(mRange) = Format(!RestDayOTPay, "0.00")
        mRange = "R" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!SundayOT, "0.00")
        'xlSheet.Range(mRange) = Format(!SundayOT, "0.00")
        mRange = "S" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!SundayOTPay, "0.00")
        'xlSheet.Range(mRange) = Format(!SundayOTPay, "0.00")
        mRange = "T" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!NoSplHoliday, "0.00")
        'xlSheet.Range(mRange) = Format(!NoSplHoliday, "0.00")
        mRange = "U" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!SplHolidayPay, "0.00")
        'xlSheet.Range(mRange) = Format(!SplHolidayPay, "0.00")
        mRange = "V" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!NoLglHoliday, "0.00")
        'xlSheet.Range(mRange) = Format(!NoLglHoliday, "0.00")
        mRange = "W" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!LglHolidayPay, "0.00")
        'xlSheet.Range(mRange) = Format(!LglHolidayPay, "0.00")
        mRange = "X" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!SalaryAdjust, "0.00")
        'xlSheet.Range(mRange) = Format(!SalaryAdjust, "0.00")
        mRange = "Y" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!IncentivePay, "0.00")
        'xlSheet.Range(mRange) = Format(!IncentivePay, "0.00")
        mRange = "Z" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!EcolaPay, "0.00")
        'xlSheet.Range(mRange) = Format(!EcolaPay, "0.00")
        mRange = "AA" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!GrossPay, "0.00")
        'xlSheet.Range(mRange) = Format(!GrossPay, "0.00")
        mRange = "AB" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!WTaxAmount, "0.00")
        'xlSheet.Range(mRange) = Format(!WTaxAmount, "0.00")
        mRange = "AC" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!SSSEmployee, "0.00")
        'xlSheet.Range(mRange) = Format(!SSSEmployee, "0.00")
        mRange = "AD" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!PHealthEmployee, "0.00")
        'xlSheet.Range(mRange) = Format(!PHealthEmployee, "0.00")
        mRange = "AE" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!PagibigEmployee, "0.00")
        'xlSheet.Range(mRange) = Format(!PagibigEmployee, "0.00")
        mRange = "AF" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!CashBond, "0.00")
        'xlSheet.Range(mRange) = Format(!CashBond, "0.00")
        mRange = "AG" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!PlacementFee, "0.00")
        'xlSheet.Range(mRange) = Format(!PlacementFee, "0.00")
        mRange = "AH" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!MiscFee, "0.00")
        'xlSheet.Range(mRange) = Format(!MiscFee, "0.00")
        mRange = "AI" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!CardTransFee, "0.00")
        'xlSheet.Range(mRange) = Format(!CardTransFee, "0.00")
        mRange = "AJ" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!SalaryLoan, "0.00")
        'xlSheet.Range(mRange) = Format(!SalaryLoan, "0.00")
        mRange = "AK" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!Vale_IOU, "0.00")
        'xlSheet.Range(mRange) = Format(!Vale_IOU, "0.00")
        mRange = "AL" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!Uniform, "0.00")
        'xlSheet.Range(mRange) = Format(!Uniform, "0.00")
        mRange = "AM" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!Medical, "0.00")
        'xlSheet.Range(mRange) = Format(!Medical, "0.00")
        mRange = "AN" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!Variances, "0.00")
        'xlSheet.Range(mRange) = Format(!Variances, "0.00")
        mRange = "AO" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!OverUnderPay, "0.00")
        'xlSheet.Range(mRange) = Format(!OverUnderPay, "0.00")
        mRange = "AP" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!SSSLoan, "0.00")
        'xlSheet.Range(mRange) = Format(!SSSLoan, "0.00")
        mRange = "AQ" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!PagibigLoan, "0.00")
        'xlSheet.Range(mRange) = Format(!PagibigLoan, "0.00")
        mRange = "AR" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!ValeWithClient, "0.00")
        'xlSheet.Range(mRange) = Format(!ValeWithClient, "0.00")
        mRange = "AS" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!Deductions, "0.00")
        'xlSheet.Range(mRange) = Format(!Deductions, "0.00")
        mRange = "AT" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!NetPay, "0.00")
        'xlSheet.Range(mRange) = Format(!NetPay, "0.00")
        mRange = "AU" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(!AcctNo = "", "n/a", !AcctNo)
        'xlSheet.Range(mRange) = IIf(!AcctNo = "", "n/a", !AcctNo)
        
        mRange = "AX" & LineCtr2
        xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(IsNull(!BillingRate), 0#, Format(!BillingRate, "0.00"))
        
            LineCtr = LineCtr + 1
            RecCtr = RecCtr + 1
            Select Case mBranchCD '!AreaCode
                Case 1      ' Head Office
                    RecCtr1 = RecCtr1 + 1
                Case 2      ' Ace
                    RecCtr2 = RecCtr2 + 1
                Case 3      ' Handyman
                    RecCtr3 = RecCtr3 + 1
                Case 4      ' Workshop
                    RecCtr4 = RecCtr4 + 1
                Case Else   ' Others
                    RecCtr5 = RecCtr5 + 1
            End Select
    
    i = i + 1

NextRec:

    .MoveNext

Wend

End With

mRange = "A1"
xlApp1.ActiveSheet.Range(mRange).Select

Exit Sub
   
ExitRtn3:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
    
End Sub

Private Function GetMandatoryRates(ByRef xlApp1 As Excel.Application, ByRef xlSheet As Excel.Worksheet, _
    ByVal mGross As Double, ByVal mFlag As Boolean) As Integer

Dim mStart As Double
Dim mEnd As Double
Dim mRate As Integer
Dim mRange As String
Dim t As Integer

GetMandatoryRates = 0#
mRate = 1

If mFlag = True Then ' Get SSS rates
    xlApp1.Sheets("SSS.EC").Select
    
    For t = 6 To 34
        mRange = "B" & t
        mStart = xlApp1.ActiveSheet.Range(mRange).VALUE
        mRange = "D" & t
        mEnd = xlApp1.ActiveSheet.Range(mRange).VALUE
        
        
        If mGross > mStart Then
            GoTo ExitRtn
        End If

        mRange = "A" & t
        If mGross >= mStart And mGross <= mEnd Then
            mRate = xlApp1.ActiveSheet.Range(mRange).VALUE
            GoTo ExitRtn
        End If
        mRate = xlApp1.ActiveSheet.Range(mRange).VALUE
    Next t
Else    ' Get PhilHealth
    xlApp1.Sheets("PHIL.HEALTH").Select
    
    For t = 8 To 34
        mRange = "B" & t
        mStart = xlApp1.ActiveSheet.Range(mRange).VALUE
        mRange = "D" & t
        mEnd = xlApp1.ActiveSheet.Range(mRange).VALUE

        If mGross > mStart Then
            GoTo ExitRtn
        End If
        
        mRange = "A" & t
        If mGross >= mStart And mGross <= mEnd Then
            mRate = xlApp1.ActiveSheet.Range(mRange).VALUE
            GoTo ExitRtn
        End If
        mRate = xlApp1.ActiveSheet.Range(mRange).VALUE
    Next t

End If

ExitRtn:

xlApp1.Sheets("30th Billing").Select


GetMandatoryRates = mRate
    
End Function

Private Sub PrintTotals(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal i As Integer, ByVal k As Integer, Optional TotalAmt As Double, _
    Optional sumString As String, Optional GFlag As Boolean = False)

Dim EndSw As Boolean
Dim mRange As String
Dim j As Integer
Dim addno As Integer
Dim mRange2 As String

'addno = 6   '7
'
''Compute for Grand Total.
'EndSw = False
'sumString = "=SUM(R[" & i * -1 & "]C:R[-2]C)"
If GFlag = False Then
    sumString = "=SUM(R[" & (i - k) * -1 & "]C:R[-1]C)"
End If
'mRange = "E" & i + addno & ":E" & i + addno

'mRange = "B" & i & ":B" & i
'xlSheet.Range(mRange).Select
'xlApp1.ActiveCell.FormulaR1C1 = "         Sub-Total"
'xlApp1.Range(mRange).Font.Bold = True

'Print Total Employee.
'mRange = "D" & i & ":D" & i
'xlSheet.Range(mRange).Select
'xlApp1.ActiveCell.FormulaR1C1 = IIf(i = k, i - 1, i - k)
'xlApp1.Range(mRange).Font.Bold = True

For j = Asc("D") To Asc("Z")
    mRange = Chr(j) & k & ":" & Chr(j) & i
    xlSheet.Range(mRange).Select
    mRange2 = Chr(j) & i
    xlSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
    'xlApp1.Range(mRange2).Font.Bold = True
    xlSheet.Range(mRange).Select
Next
For j = Asc("A") To Asc("R")    ' K
    If Chr(j) <> "Q" Then
        mRange = "A" & Chr(j) & k & ":A" & Chr(j) & i
        xlSheet.Range(mRange).Select
        mRange2 = "A" & Chr(j) & i
        xlSheet.Range(mRange2).Activate
        xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
        'xlApp1.Range(mRange2).Font.Bold = True
        If Asc("P") = j And GFlag = True And TotalAmt = 99 Then
           mRange2 = "A" & Chr(j) & i
            xlSheet.Range(mRange2).Select
            xlApp1.Selection.Copy
            TotalAmt = Format(xlApp1.ActiveCell.VALUE, "0.00")
            Call CopyToSheetDataSource(xlSheet, xlApp1, TotalAmt)
        End If
        
        xlSheet.Range(mRange).Select
    End If
Next

End Sub

Private Sub PrintTotalsForCarlson(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal i As Integer, ByVal k As Integer, Optional TotalAmt As Double, _
    Optional sumString As String, Optional GFlag As Boolean = False, Optional TotalEmp As Integer = 0)

Dim EndSw As Boolean
Dim mRange As String
Dim j As Integer
Dim addno As Integer
Dim mRange2 As String

'addno = 6   '7
'
''Compute for Grand Total.
'EndSw = False
'sumString = "=SUM(R[" & i * -1 & "]C:R[-2]C)"
If GFlag = False Then
    sumString = "=SUM(R[" & (i - k) * -1 & "]C:R[-1]C)"
End If
'mRange = "E" & i + addno & ":E" & i + addno

'mRange = "B" & i & ":B" & i
'xlSheet.Range(mRange).Select
'xlApp1.ActiveCell.FormulaR1C1 = "         Sub-Total"
'xlApp1.Range(mRange).Font.Bold = True

'Print Total Employee.
'mRange = "D" & i & ":D" & i
'xlSheet.Range(mRange).Select
'xlApp1.ActiveCell.FormulaR1C1 = IIf(i = k, i - 1, i - k)
'xlApp1.Range(mRange).Font.Bold = True

For j = Asc("J") To Asc("Z")
    mRange = Chr(j) & k & ":" & Chr(j) & i
    xlSheet.Range(mRange).Select
    mRange2 = Chr(j) & i
    xlSheet.Range(mRange2).Activate
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
    'xlApp1.Range(mRange2).Font.Bold = True
    xlSheet.Range(mRange).Select
Next
For j = Asc("A") To Asc("S")    ' K
    If Chr(j) <> "R" Then
        mRange = "A" & Chr(j) & k & ":A" & Chr(j) & i
        xlSheet.Range(mRange).Select
        mRange2 = "A" & Chr(j) & i
        xlSheet.Range(mRange2).Activate
        xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
        'xlApp1.Range(mRange2).Font.Bold = True
        If Asc("R") = j And GFlag = True And TotalAmt = 99 Then
           mRange2 = "A" & Chr(j) & i
            xlSheet.Range(mRange2).Select
            xlApp1.Selection.Copy
            TotalAmt = Format(xlApp1.ActiveCell.VALUE, "0.00")
            Call CopyToSheetDataSource(xlSheet, xlApp1, TotalAmt, TotalEmp)
        End If
        
        xlSheet.Range(mRange).Select
    End If
Next

End Sub


Private Sub CopyToSheetDataSource(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal TAmount As Double, Optional TotalEmp As Integer = 0)
    
Dim mRange As String
Dim AmtinWords As String

    xlApp1.Sheets("Data Source").Select
    mRange = "A8"
    xlApp1.ActiveSheet.Range(mRange).Select
    xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    AmtinWords = oCommonFunction.Word(Val(TAmount))
    mRange = "A2"
    xlApp1.ActiveSheet.Range(mRange).Select
    xlApp1.ActiveCell.VALUE = " PESOS: " & AmtinWords & " "
    
    mRange = "A7"
    xlApp1.ActiveSheet.Range(mRange).Select
    xlApp1.ActiveCell.VALUE = TotalEmp
    
    xlApp1.Sheets("Billing Report").Select

End Sub

Private Sub txtSOAMonth_GotFocus()
txtSOAMonth.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtSOAMonth
Call txtHLigths(txtSOAMonth)
End Sub

Private Sub txtSOAMonth_LostFocus()
txtSOAMonth.BackColor = vbWhite
mSOANo = txtSOAMonth & "-" & txtSOASeq
txtSOANo.Text = mSOANo
prevSOA = txtSOAMonth.Text
End Sub

Private Sub txtSOASeq_GotFocus()
txtSOASeq.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtSOASeq
Call txtHLigths(txtSOASeq)
End Sub

Private Sub txtSOASeq_LostFocus()
txtSOASeq.BackColor = vbWhite
txtSOASeq.Text = IIf(Len(txtSOASeq.Text) <> 3, Right("000" & RTrim(LTrim(txtSOASeq.Text)), 3), txtSOASeq.Text)
mSOANo = txtSOAMonth & "-" & txtSOASeq
txtSOANo.Text = mSOANo
End Sub

