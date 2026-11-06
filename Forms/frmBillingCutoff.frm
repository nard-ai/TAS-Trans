VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmBillingCutoff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Cutoff for Billing"
   ClientHeight    =   5220
   ClientLeft      =   1920
   ClientTop       =   1875
   ClientWidth     =   11685
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmBillingCutoff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   11685
   Begin VB.Frame Frame2 
      Caption         =   "Add Cutoff for SOA#:"
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   11415
      Begin VB.TextBox txtSeqNo 
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
         Left            =   7200
         TabIndex        =   18
         Text            =   "1"
         Top             =   2955
         Width           =   615
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
         Left            =   9840
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   2955
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4683
         SortKey         =   2
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
         NumItems        =   54
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
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
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Regular Days"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Day Off"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "DayOff (Spl)"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Sunday OT"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "DayOff (Legal)"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "#Days Worked"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Tardi Hours"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Reg. O.T."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Dayoff O.T."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Dayoff (Spl) O.T."
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Legal O.T."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Dayoff (Lgl) O.T."
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Night Diff'l"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Daily Rate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Basic Pay"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "DayOff Pay"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "DayOff Spl Pay"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Legal Pay"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "DayOff Lgl Pay"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Tardi (UT) Pay"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Reg OT Pay"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "DayOff Pay"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "DayOff Spl Pay"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "Legal Pay"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "DayOff Lgl Pay"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "Night Diff'l Pay"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "ECOLA"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "Salary Adjust"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "Gross Earnings"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "13th Month Pay"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "SSS/EC (ER)"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "PhilHealth (ER)"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "PagIBIG (ER)"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "Total Pay & Benefits"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "Agency Fee"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "Gross Billing"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Text            =   "VAT 12%"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Text            =   "Other Deductions"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "Total Amount Due"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   43
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   44
            Text            =   "Gross Payroll"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   45
            Text            =   "Ignore (Lates/Ecola)"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   46
            Text            =   "Diff'n (over/under)"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   47
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   48
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   49
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   50
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   51
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   52
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   53
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Sequence#:"
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
         Left            =   5880
         TabIndex        =   19
         Top             =   3000
         Width           =   1125
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Excluded Records: 0"
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
         Left            =   2280
         TabIndex        =   16
         Top             =   3000
         Width           =   2085
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Included Records:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   1845
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount:"
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
         Left            =   8520
         TabIndex        =   11
         Top             =   3000
         Width           =   1365
      End
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
      Left            =   6960
      TabIndex        =   9
      Text            =   "BJC09-01-014"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdDeleteCutoff 
      Caption         =   "&Exclude Selected Cutoff"
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
      Left            =   3000
      TabIndex        =   7
      Top             =   4680
      Width           =   2475
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Return"
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
      Left            =   9000
      TabIndex        =   6
      Top             =   4680
      Width           =   2475
   End
   Begin VB.CommandButton cmdAddCutOff 
      Caption         =   "Include &Cutoff for Billing"
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
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   2475
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11415
      Begin VB.CommandButton cmdShowTrans 
         Caption         =   "Show Records"
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
         Left            =   9720
         TabIndex        =   13
         Top             =   240
         Width           =   1515
      End
      Begin VB.ComboBox cmbCutOff 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   5295
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
         ItemData        =   "frmBillingCutoff.frx":08CA
         Left            =   10200
         List            =   "frmBillingCutoff.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   855
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
         Left            =   6960
         TabIndex        =   14
         Top             =   270
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
         Left            =   8400
         TabIndex        =   15
         Top             =   270
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
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   1365
      End
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmBillingCutoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCutOffTable As ADODB.Recordset
Dim rsClientTable As ADODB.Recordset
Dim rsPayrollTrans As ADODB.Recordset
Dim OnProcessFlag As Boolean
Dim NoDaysperCutoff As Integer
Dim StrQuery As String
Dim AddCutoffSw As Boolean
Dim mSOANo As String
Dim mClientCode As Integer

Dim Listctr As Integer
Dim ExcludedCtr As Integer
Dim IncludedCtr As Integer
Dim CutOffIndex1 As Integer

Dim DaysPresent As Double
Dim TaxRefundHours As Double
Dim NoRegHoliday As Double
Dim SundayOT As Double
Dim NoSplHoliday As Double
Dim DaysWorked As Double
Dim HoursTardiness As Double
Dim OTHours As Double
Dim ExcessSplOTHours As Double
Dim LglHolidayOTHours As Double
Dim ExcessOTSunLegalHrs As Double
Dim NightDiffOT As Double
Dim OtherIncome As Double
Dim Allowance As Double

'Dim RestDayWorkedOTRate As Double   ' 30%
'Dim SpecialWorkedDayOffRate As Double   ' 50%
'Dim LegalHolidayWithPremiumRate As Double   ' 100%
'Dim LegalHolidayWorkedDayoffRate As Double  ' 160%

'Dim WorkHrsPerDay As Double
'Dim OverTimeRate As Double
'Dim BillingTemplate As String

'Dim ExcessSplOTRate As Double
'Dim ExcessLglHolOTRate As Double
'Dim ExcessSunLglOTRate As Double
'Dim NightDiffRate As Double
'Dim VatPercentRate As Double
'Dim AgencyRate As Double

Dim TaxRefundPay As Double
Dim RegHolidayPay As Double
Dim SundayOTPay As Double
Dim SplHolidayPay As Double

Dim TardinessDeduction As Double ' Deduction
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

Dim CalamityLoan As Double

Dim SSSEmployer As Double
Dim PHealthEmployer As Double
Dim PagibigEmployer As Double

Private Sub GetClientCutoff()
Dim mYearNo As Integer

    If Format(Now, "mm") < 6 Then
        mYearNo = Format(Now, "yyyy") - 1
    Else
        mYearNo = Format(Now, "yyyy")
    End If
    ' Get Cutoff table based from selected client.
'    mClientCode = cmbClient.ItemData(cmbClient.ListIndex)
'    rsClientTable.MoveFirst
'    rsClientTable.Find "clientcode = " & mClientCode, 0, adSearchForward ' goto selected client record
'    If Not rsClientTable.EOF Then
'        NoDaysperCutoff = IIf(IsNull(rsClientTable!ClientCutoffDates), 0, Val(rsClientTable!ClientCutoffDates))
'        txtCutOffBase.Text = rsClientTable!ClientCutoff & " (" & Trim(NoDaysperCutoff) & ")"
'        mSOANo = IIf(IsNull(rsClientTable!SOANo), 0, Val(rsClientTable!SOANo))
'        mSOANo = mSOANo + 1 ' series number + 1
''        If rsClientTable!clientflag1 = True Then
''            Image3.Visible = True ' show the on process flag
''        Else
''            Image3.Visible = False
''        End If
''        Image2.Visible = False
        ' Get only those cutoff period with viewflag = true
        
        Set rsCutOffTable = New ADODB.Recordset
        'Set rsCutOffTable = ReturnRst("SELECT * FROM CutOffTable where clientcode = " & frmBilling.txtCutOffBase.Text & " and viewflag = true ORDER BY RecNo", cnnStr, mstatic)
        'Set rsCutOffTable = ReturnRst("SELECT a.*, b.EntryDate FROM CutOffTable a LEFT JOIN qryPayrollCutoff b ON a.agencycode = b.agencycode and a.clientcode = b.clientcode and a.StartcutoffText = b.StartCutoffText and a.EndCutoffText = b.EndCutoffText Where a.clientcode = " & frmBilling.txtCutOffBase.Text & " and a.viewflag = true ORDER BY RecNo", cnnStr, mstatic)
        'Updated July 10, 2009.
        'Set rsCutOffTable = ReturnRst("SELECT DISTINCT * FROM BillingCutoffTable WHERE Format(EndCutoffDate, ""yyyy"") = Format(now, ""yyyy"") and ClientCode = " & frmBilling.txtCutOffBase.Text & " and ViewFlag = true ORDER BY StartCutOffDate", cnnStr, mstatic)
        'Updated Jan. 5, 2010.
        Set rsCutOffTable = ReturnRst("SELECT DISTINCT * FROM BillingCutoffTable WHERE Format(EndCutoffDate, ""yyyy"") >= " & mYearNo & " and ClientCode = " & frmBilling.txtCutOffBase.Text & " and ViewFlag = true ORDER BY StartCutOffDate", cnnstr, mstatic)
        
        ' Populate cutoff list.
        'chkNewCutoff.Visible = True
        'chkNewCutoff.Enabled = True
        With rsCutOffTable
            cmbCutOff.Clear
            While Not (.EOF Or .BOF)
                'cmbCutoff.AddItem Format(.Fields(5), "mm/dd/yyyy") & "-" & Format(.Fields(6), "mm/dd/yyyy") & "  " & IIf(!billedflag = 1, "Close", "Open")
                'cmbCutOff.AddItem Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & ":" & Format(!ProcessDate, "mmddyyyy") & " " & IIf(!PaidFlag = True, "Bil (" & !soano & ")", IIf(!Selectedflag = True, "Sel (" & !soano & ")", "Clo")) & " " & IIf(IsNull(!CutOffSeqNo), 0, !CutOffSeqNo)
                cmbCutOff.AddItem Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & ":" & Format(!EntryDate, "mmddyyyy") & " " & IIf(!PaidFlag = True, "Bil (" & !soano & ")", IIf(!Selectedflag = True, "Sel (" & !soano & ")", "Clo")) & " " & IIf(IsNull(!CutOffSeqNo), 0, !CutOffSeqNo)
                cmbCutOff.ItemData(cmbCutOff.NewIndex) = .Fields(0)
                If frmBilling.txtSOANo.Text = !soano Then
                    If Val(txtSeqNo.Text) <= !CutOffSeqNo Then
                        txtSeqNo.Text = !CutOffSeqNo + 1
                    End If
                ElseIf Val(txtSeqNo.Text) = 0 Then
                        txtSeqNo.Text = "1"
                End If
'                If !BilledFlag = False Then
'                    chkNewCutoff.Visible = False
'                End If
                .MoveNext
            Wend
            cmbCutOff.Refresh
            If cmbCutOff.ListCount <> 0 Then
                cmbCutOff.ListIndex = 0
            End If
        End With
    'txtSeriesNo.Text = CompanyShort & Format(Now, "yy") & "-" & Right("0000" & SOANo, 4)
    'txtSeriesNo.Refresh
'    Else
'        MsgBox "Invalid client code selected!", vbInformation, PayrollAlert
'    End If
End Sub

Private Sub GetSelectedCutoff()
Dim rsSelectedCutoff As ADODB.Recordset

'strQuery = "SELECT * FROM CutOffTable WHERE AgencyCode = " & CompanyCode & " and ClientCode = " & frmBilling.txtCutOffBase.Text & _
'    " and SelectedFlag = True and Viewflag = True and SOANo = '" & frmBilling.txtSOANo.Text & "' ORDER BY RecNo"

'Original Code.
'strQuery = "SELECT a.*, b.EntryDate FROM CutOffTable a LEFT JOIN qryPayrollCutoff b ON a.agencycode = b.agencycode and a.clientcode = b.clientcode and a.StartcutoffText = b.StartcutoffText and a.EndCutoffText = b.EndCutoffText WHERE a.AgencyCode = " & CompanyCode & " and a.ClientCode = " & frmBilling.txtCutOffBase.Text & _
'    " and a.SelectedFlag = True and a.Viewflag = True and a.SOANo = '" & frmBilling.txtSOANo.Text & "' ORDER BY RecNo"

StrQuery = "SELECT DISTINCT * FROM BillingCutoffTable a WHERE a.ClientCode = " & frmBilling.txtCutOffBase.Text & _
    " and a.SelectedFlag = True and a.Viewflag = True and a.SOANo = '" & frmBilling.txtSOANo.Text & "' ORDER BY StartCutoffDate"

Set rsSelectedCutoff = ReturnRst(StrQuery, cnnstr, mstatic)
With rsSelectedCutoff
    frmBilling.cmbSelectedCutOff.Clear
    frmBilling.cmbSelectedCutOff.AddItem "All Cutoff"
    While Not (.EOF Or .BOF)
        'frmBilling.cmbSelectedCutOff.AddItem Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & ":" & Format(!ProcessDate, "mmddyyyy") & " Sel"
        frmBilling.cmbSelectedCutOff.AddItem Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & ":" & Format(!EntryDate, "mmddyyyy") & " Sel"
        'frmBilling.cmbSelectedCutOff.ItemData(cmbCutOff.NewIndex) = .Fields(0)
        .MoveNext
    Wend
End With
rsSelectedCutoff.Close
End Sub

Private Sub cmbClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

'Private Sub cmbClient_LostFocus()
'    ' Get Cutoff table based from selected client.
'    mClientCode = cmbClient.ItemData(cmbClient.ListIndex)
'    mCompanyName = Trim(cmbClient.Text)
'    Set rsCutOffTable = ReturnRst("SELECT * FROM cutofftable where clientcode = " & mClientCode & " and viewflag = true ORDER BY RecNo", cnnStr, mstatic)
''    If rsCutOffTable.RecordCount = 0 Then ' no payroll cut-off selected, enable new cutoff
''        chkNewCutoff.Enabled = True
''    Else
''        chkNewCutoff.Enabled = False
''        chkNewCutoff.VALUE = 0
''    End If
'    ' Populate cutoff list.
'    'chkNewCutoff.Visible = True
'    'chkNewCutoff.Enabled = True
'    With rsCutOffTable
'        cmbCutOff.Clear
'        While Not (.EOF Or .BOF)
'            cmbCutOff.AddItem Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & "  " & IIf(!BilledFlag = True, "Close", "Open")
'            cmbCutOff.ItemData(cmbCutOff.NewIndex) = .Fields(0)
''            If !BilledFlag = False Then
''                chkNewCutoff.Visible = False
''            End If
'            .MoveNext
'        Wend
'        cmbCutOff.Refresh
'    End With
'    'If rsCutOffTable.State = adStateOpen Then rsCutOffTable.Close
'End Sub

Private Sub cmbCutoff_Click()
Dim CutOffIndex As Integer
Dim mEndDate As Date

    'On Error GoTo Exit Sub.
    If cmbCutOff.ListIndex = -1 Then
        Exit Sub
    Else
        CutOffIndex = IIf(cmbCutOff.ListIndex = -1, 0, cmbCutOff.ItemData(cmbCutOff.ListIndex))
        mEndDate = Format(Left(cmbCutOff.Text, 10), "mm/dd/yyyy")
    End If
    rsCutOffTable.MoveFirst
    ' Original Code.
    'rsCutOffTable.Find "recno = " & CutOffIndex, 0, adSearchForward ' goto selected client record
    'rsCutOffTable.Find "EndCutOffDate = '" & mEndDate & "', 0, adSearchForward ' goto selected client record"
                
    txtCutoffDate(2).Text = Format(Left(cmbCutOff.Text, 10), "mm/dd/yyyy")
    txtCutoffDate(3).Text = Format(Mid(cmbCutOff.Text, 12, 10), "mm/dd/yyyy")
    
'    txtSeqNo.Text = IIf(Mid(cmbCutOff.Text, 50, 2) = "0", "1", Val(Mid(cmbCutOff.Text, 50, 2)) + 1)
    
    frmBilling.txtCutoffDate(2).Text = Format(Left(cmbCutOff.Text, 10), "mm/dd/yyyy")
    frmBilling.txtCutoffDate(3).Text = Format(Mid(cmbCutOff.Text, 12, 10), "mm/dd/yyyy")
    
    If InStr(1, cmbCutOff.Text, "Sel") <> 0 Then
        AddCutoffSw = False
    Else
        AddCutoffSw = True
    End If
    
    If InStr(1, cmbCutOff.Text, "Bil") <> 0 Then
        AddCutoffSw = False
    End If
    
    ' Get the old client series no.
    'txtSeriesNo.Text = IIf(IsNull(rsCutOffTable!SOANo), txtSeriesNo.Text, rsCutOffTable!SOANo)
    'txtSeriesNo.Refresh

'    If rsCutOffTable!BilledFlag = False Then
'        Image2.Visible = True
'        chkNewCutoff.Enabled = False
'        chkNewCutoff.VALUE = 0
'    Else
'        Image2.Visible = False
'        chkNewCutoff.Enabled = True
'        'chkNewCutoff.Visible = True
'    End If
'    If chkNewCutoff = 1 Then ' get new cut-off w/ accordance to the last cut-off
'        rsCutOffTable.MoveLast  ' go to last record and get the last cut-off
'        mEndDate = Format(rsCutOffTable!EndCutoffDate + 1, "mm/dd/yyyy")
'        txtCutOffDate(0).Text = Format(mEndDate, "mm/dd/yyyy")
'        txtCutOffDate(1).Text = Format((mEndDate + NoDaysperCutoff), "mm/dd/yyyy")
'        txtCutOffDate(0).Enabled = True
'        txtCutOffDate(1).Enabled = True
'    Else
'        txtCutOffDate(0).Enabled = False
'        txtCutOffDate(1).Enabled = False
'    End If
    Exit Sub

ExitSub:

    'MsgBox "Error: In Cut-off Selection!", vbInformation, PayrollAlert
    
    Err.Clear

End Sub

Private Sub cmbCutoff_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdAddCutoff_Click()
    Call IncludeCutOffInBilling
    Call UpdateCutoffTable(True)
    Call GetClientCutoff
    cmdAddCutoff.Enabled = False
    cmbCutOff.ListIndex = CutOffIndex1
    'txtSeqNo.Text = Val(txtSeqNo.Text) + 1
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub UpdateCutoffTable(ByVal UpdateSw As Boolean) ' if true means add
Dim mProcessDate As String
Dim res As Variant

mProcessDate = Mid(cmbCutOff.Text, 23, 8)

If InStr(1, cmbCutOff.Text, " bil ", vbTextCompare) <> 0 Then
    res = MsgBox("Cutoff already marked Billed. Do you really want to delete it?", vbYesNo, PayrollAlert)
    If res = vbNo Then
        Exit Sub
    End If
ElseIf UpdateSw = False Then
    res = MsgBox("Delete selected payroll cutoff?", vbYesNo, PayrollAlert)
    If res = vbNo Then
        Exit Sub
    End If
End If
If UpdateSw = True Then
'    strQuery = "UPDATE CutoffTable SET SOANo = '"
    StrQuery = "UPDATE BillingCutoffTable SET SOANo = '" & _
        frmBilling.txtSOANo.Text & "', Selectedflag = 1, PaidDate = '" & _
        Format(Now, "mm/dd/yyyy") & "', CutoffSeqNo = " & txtSeqNo.Text & " WHERE ClientCode = " & _
        Val(frmBilling.txtCutOffBase.Text) & " and StartCutOffText = '" & _
        Format(txtCutoffDate(2), "yyyymmdd") & "' and EndCutoffText = '" & _
        Format(txtCutoffDate(3), "yyyymmdd") & "' and Format(EntryDate, ""mmddyyyy"") = '" & mProcessDate & "'"
        'Format(txtCutoffDate(3), "yyyymmdd") & "' and format(Processdate, ""mmddyyyy"") = '" & mProcessDate
    Call Action_Query(StrQuery, cnnstr)
    ' Add new selected cutoff to billing Form for filtering.
    frmBilling.cmbSelectedCutOff.AddItem Format(txtCutoffDate(2), "mm/dd/yyyy") & "-" & Format(txtCutoffDate(3), "mm/dd/yyyy") & ":" & mProcessDate & " Sel"
Else ' Exclude cutoff in billing list.
    StrQuery = "DELETE * FROM BilledTrans WHERE ClientCode = " & _
        Val(frmBilling.txtCutOffBase.Text) & " and StartPayDate = '" & _
        Format(txtCutoffDate(2), "mm/dd/yyyy") & "' and EndPayDate = '" & _
        Format(txtCutoffDate(3), "mm/dd/yyyy") & "' and Paytext2 = '" & mProcessDate & "'"
    If Action_Query(StrQuery, cnnstr) = True Then
        'strQuery = "UPDATE CutoffTable SET SOANo = '', Selectedflag = 0, PaidDate = #01/01/1900#, PaidFlag = 0, CutOffSeqNo = 0 WHERE ClientCode = "
        StrQuery = "UPDATE BillingCutoffTable SET SOANo = '', Selectedflag = 0, PaidDate = #01/01/1900#, PaidFlag = 0, CutOffSeqNo = 0 WHERE ClientCode = " & _
            Val(frmBilling.txtCutOffBase.Text) & " and StartCutOffText = '" & _
            Format(txtCutoffDate(2), "yyyymmdd") & "' and EndCutoffText = '" & _
            Format(txtCutoffDate(3), "yyyymmdd") & "' and Format(EntryDate, ""mmddyyyy"") = '" & mProcessDate & "'"
            'Format(txtCutoffDate(3), "yyyymmdd") & "' and format(Processdate, ""mmddyyyy"") = '" & mProcessDate
        Call Action_Query(StrQuery, cnnstr)
        MsgBox "Payroll cutoff " & txtCutoffDate(2) & " to " & txtCutoffDate(3) & " was successfully deleted.", vbInformation, PayrollAlert
        ListView1(0).ListItems.Clear
        cmdDeleteCutoff.Enabled = False
        txtGrossBill.Text = "0.00"
        lblEntries(0).Caption = "Excluded Records: 0"
        lblEntries(7).Caption = "Included Records: 0"
        UpdateBillingTransactions
    End If
End If

End Sub

Private Sub UpdateBillingTransactions()  ' Update Billing Transactions
Dim rsBilledTrans As ADODB.Recordset
Dim litem1 As ListItem
Dim mBasicPay As Double
Dim ExcludedFlag As Integer
Dim mListView As ListView
Dim i As Integer
Dim mZeroValue As Double
Dim Listctr As Integer
Dim GrossBill As Double

mZeroValue = 0#
GrossBill = 0#
Set mListView = frmBilling.ListView1(0)

Listctr = 0
gGrossBill = 0#

StrQuery = "SELECT * FROM BilledTrans WHERE ClientCode = " & mClientCode & _
    " and SOANo = '" & frmBilling.txtSOANo.Text & "'"
Set rsBilledTrans = ReturnRst(StrQuery, cnnstr, mstatic)

With rsBilledTrans
mListView.ListItems.Clear
While Not (.EOF Or .BOF)
    Listctr = Listctr + 1
    Set litem1 = mListView.ListItems.Add(, , !EmployeeID)
        litem1.SubItems(1) = IIf(IsNull(!EmpNo), "", !EmpNo)
        litem1.SubItems(2) = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
        litem1.SubItems(3) = Space(5) 'DTR
        litem1.SubItems(4) = Format(!DaysPresent, "0")
        litem1.SubItems(5) = Format(!DaysPresent2, "0") ' edit days present
        litem1.SubItems(6) = Format(!TaxRefundHours, "0") ' extra days
        litem1.SubItems(7) = Format(!TaxRefundHours2, "0") ' edit extra days
        litem1.SubItems(8) = Format(!NoRegHoliday, "0") ' no spl holiday
        litem1.SubItems(9) = Format(!NoRegHoliday2, "0") ' edit no spl holiday
        
        litem1.SubItems(10) = Format(!SundayOT, "0")     ' sunday ot
        litem1.SubItems(11) = Format(!SundayOT2, "0")     ' edit sunday ot
        litem1.SubItems(12) = Format(!NoSplHoliday, "0") ' no lgl holiday
        litem1.SubItems(13) = Format(!NoSplHoliday2, "0") ' edit no lgl holiday
        litem1.SubItems(14) = Format(!DaysPresent, "0")  ' # of days worked
        litem1.SubItems(15) = Format(!DaysPresent2, "0")  ' # of days worked
        litem1.SubItems(16) = Format(!HoursTardiness, "0.00") ' hours tardi
        litem1.SubItems(17) = Format(!HoursTardiness2, "0.00")  ' edit hours tardi
        litem1.SubItems(18) = Format(!OTHours, "0.00")
        litem1.SubItems(19) = Format(!OTHours2, "0.00")
        litem1.SubItems(20) = Format(!Excesssplholothours, "0.00")
        litem1.SubItems(21) = Format(!ExcessSplHolOTHours2, "0.00")
        litem1.SubItems(22) = "0.00"
        litem1.SubItems(23) = Format(!LglHolidayOTHours, "0.00")
        litem1.SubItems(24) = Format(!LglHolidayOTHours2, "0.00")
        litem1.SubItems(25) = Format(!ExcessOTSunLegalHrs, "0.00")
        litem1.SubItems(26) = Format(!ExcessOTSunLegalHrs2, "0.00")
        litem1.SubItems(27) = Format(!NightDiffOT, "0.00")
        litem1.SubItems(28) = Format(!NightDiffOT2, "0.00")
        
        litem1.SubItems(29) = Format(!DailyRate, "0.00")    ' Daily Rate
        litem1.SubItems(30) = Format(!BasicPay, "0.00")     ' Basic Pay
        litem1.SubItems(31) = Format(!TaxRefundPay, "0.00")     ' Premium pay on extra days worked
        litem1.SubItems(32) = Format(!RegHolidayPay, "0.00")    ' Premium pay on extra days worked
        litem1.SubItems(33) = Format(!SundayOTPay, "0.00")      ' premium pay on extra days worked
        litem1.SubItems(34) = Format(!SplHolidayPay, "0.00")    ' premium pay on extra days worked
        litem1.SubItems(35) = Format(!TardinessDeduction, "0.00")
        litem1.SubItems(36) = Format(!OTPay, "0.00")            ' ot
        litem1.SubItems(37) = Format(!excesssplholotpay, "0.00") ' ot
        litem1.SubItems(38) = "-"
        litem1.SubItems(39) = Format(!LglHolidayOTPay, "0.00")  ' ot
        litem1.SubItems(40) = Format(!excessotsunlegalpay, "0.00")  ' ot
        litem1.SubItems(41) = Format(!NightDiffPay, "0.00") ' ot

        litem1.SubItems(42) = Format(!EcolaPay, "0.00")
        litem1.SubItems(43) = Format(!OtherIncome, "0.00")
        litem1.SubItems(44) = Format(!OtherIncome2, "0.00")
        litem1.SubItems(45) = Format(!GrossPay, "0.00")
        litem1.SubItems(46) = Format(!Month13Payment, "0.00")

        'litem1.SubItems(47) = Format(!SSSEmployer + !ECCEmployer, "0.00")  ' ECC was already included in SSSEmployer payroll entry
        litem1.SubItems(47) = Format(!SSSEmployer, "0.00")
        litem1.SubItems(48) = Format(!PHealthEmployer, "0.00")
        litem1.SubItems(49) = Format(!PagibigEmployer, "0.00")
        ' ECC already included in SSS Employer computation.
        'litem1.SubItems(50) = Format(!GrossPay + !SSSEmployer + !ECCEmployer + !PHealthEmployer + !PagibigEmployer, "0.00") ' Total pay and benefits
        litem1.SubItems(50) = Format(!GrossPay + !SSSEmployer + !PHealthEmployer + !PagibigEmployer, "0.00")  ' Total pay and benefits
        litem1.SubItems(51) = Format(!CalamityLoan, "0.00") ' Agency Charge
        'GrossBill = !GrossPay + !SSSEmployer + !ECCEmployer + !PHealthEmployer + !PagibigEmployer + !CalamityLoan
        GrossBill = !GrossPay + !SSSEmployer + !PHealthEmployer + !PagibigEmployer + !CalamityLoan
        gGrossBill = gGrossBill + GrossBill
        'litem1.SubItems(39) = Format(!Grosspay + !SSSemployer + !ECCemployer + !phealthemployer + !pagibigemployer + (!DaysPresent * !ClientChargePerDay), "0.00") ' Gross Billing
        litem1.SubItems(52) = Format(!GrossBill, "0.00") ' Gross Billing
        'VatRate = (GrossBill * !VatRate)
        litem1.SubItems(53) = Format(!VatRate, "0.00") 'Vat 12%
        litem1.SubItems(54) = Format(!Allowance, "0.00") ' Incentive pay
        litem1.SubItems(55) = Format(!Allowance2, "0.00") ' edit Incentive pay
        litem1.SubItems(56) = Format(!GrossBill + !Allowance, "0.00") ' Total Amount Due
        litem1.SubItems(57) = "::"
        litem1.SubItems(58) = Format(!GrossPay, "0.00") 'Gross Payroll
        litem1.SubItems(59) = "-"   ' Ignore Lates
        litem1.SubItems(60) = Format(!GrossPay - !GrossPay, "0.00") ' Diff'l

        litem1.SubItems(61) = "==" 'Format((!DailyRate * 15), "0.00") ' Agency Fee
    
        litem1.SubItems(62) = Format(!StartPayDate, "mm/dd/yyyy") ' Start Cutoff
        litem1.SubItems(63) = Format(!EndPayDate, "mm/dd/yyyy")  ' End Cutoff
        litem1.SubItems(64) = !soano  'EmployeeID
        litem1.SubItems(65) = IIf(IsNull(!SSSRate), 0#, !SSSRate)    'SSS Rate
        litem1.SubItems(66) = IIf(IsNull(!PHRate), 0#, !PHRate)    'PH Rate
        
        If !ExcludedFlag = 1 Then
            mListView.ListItems.Item(Listctr).Checked = True
        End If

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
    
rsBilledTrans.Close

frmBilling.txtGrossBill.Text = Format(gGrossBill, "0.00")
frmBilling.lblEntries(7).Caption = "Total Records: " & Listctr

End Sub

'Private Sub cmdPrintSOA_Click()
'Dim amountinwords As String
'
'amountinwords = oCommonFunction.Word(gGrossBill)
'
'Call GenerateSOAReport(amountinwords)
'
'End Sub

Private Sub cmdShowPayrollEntries_Click()
    Call DisplayTransactions
End Sub

Private Sub cmdDeleteCutoff_Click()
    Call UpdateCutoffTable(False)
    Call GetClientCutoff
    Call GetSelectedCutoff
    'cmbCutOff.ListIndex = CutOffIndex1 ' for checking of error
    txtSeqNo.Text = IIf(Val(txtSeqNo.Text) > 0, Val(txtSeqNo.Text) - 1, 0)
End Sub

Private Sub cmdShowTrans_Click()
    CutOffIndex1 = cmbCutOff.ListIndex
    Call DisplayTransactions
    If AddCutoffSw = True Then
        cmdAddCutoff.Enabled = True
        cmdDeleteCutoff.Enabled = False
    Else
        cmdAddCutoff.Enabled = False
        cmdDeleteCutoff.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    ' Get Company Info.
    Call GetClientCutoff
    lblEntries(1).Caption = "Client: " & UCase(frmBilling.cmbClient.Text)
    ExcludedCtr = 0
    IncludedCtr = 0
    CutOffIndex1 = 0
    mClientCode = frmBilling.txtCutOffBase.Text
    Frame2.Caption = "Add Cutoff for SOA#: " & frmBilling.txtSOANo
    'Call CompanyInfo
    ' Get Client List.
'    Set rsClientTable = New ADODB.Recordset
' '  strQuery = "SELECT clientcode, clientname, clientcutoff, clientcutoffdates, clientflag1 FROM clienttable WHERE clientcode <> 0"
'    strQuery = "SELECT a.*, b.SOANo FROM Clienttable a INNER JOIN ClientBillTable b On a.ClientCode = b.ClientCode WHERE a.clientcode <> 0 AND RTRIM(a.CLIENTNAME) <> '' ORDER BY a.clientname"
'    Set rsClientTable = ReturnRst(strQuery, cnnStr, mstatic)
'    Call oCommonFunction.FillCombo(cmbClient, rsClientTable)
'    If rsClientTable.RecordCount <> 0 Then cmbClient.ListIndex = 0
'
'    mskStartContract.Text = Format(Now, "mm/dd/yyyy")
    
End Sub

Private Sub Form_Load()
Dim StrQuery As String
    
 '  Center screen
    Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    cmdAddCutoff.Enabled = False
    cmdDeleteCutoff.Enabled = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsCompanyTable = Nothing
Set rsClientTable = Nothing
End Sub

Private Sub ListView1_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)

If ListView1(0).SelectedItem.Checked = True Then
    ExcludedCtr = ExcludedCtr + 1
    If IncludedCtr > 0 Then
        IncludedCtr = IncludedCtr - 1
    End If
Else
    If ExcludedCtr > 0 Then
        ExcludedCtr = ExcludedCtr - 1
    End If
    IncludedCtr = IncludedCtr + 1
End If

lblEntries(7).Caption = "Included Records: " & IncludedCtr
lblEntries(0).Caption = "Excluded Records: " & ExcludedCtr

End Sub

Private Sub txtCutoffDate_GotFocus(Index As Integer)
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCutoffDate_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Public Sub CompanyInfo()
    
    StrQuery = "SELECT * FROM companytable WHERE agencycode = true"
    Set rsCompanyTable = ReturnRst(StrQuery, CnnstrHR, mstatic)

End Sub

'Private Function PayrollMasterRec() As Boolean
'
'strQuery = "SELECT * FROM PayrollMaster WHERE clientcode = " & mClientCode & _
'        " and startpaydate = '" & txtCutoffDate(2) & "' and endpaydate = '" & txtCutoffDate(3) & "'"
'Set rsPayrollTrans = ReturnRst(strQuery, cnnStr, mstatic)
'
'End Function

Private Sub DisplayTransactions()
Dim rsSSSTable As ADODB.Recordset
Dim rsPHealthTable As ADODB.Recordset
Dim litem1 As ListItem
Dim Listctr1 As Integer
Dim GrossBill As Double
Dim LgGrossBill As Double
Dim VatRate As Double
Dim mProcessDate As String
Dim GetName2 As String

LgGrossBill = 0#
GrossBill = 0#
VatRate = 0#
Listctr1 = 0
IncludedCtr = 0
ExcludedCtr = 0

mProcessDate = Mid(cmbCutOff.Text, 23, 8)

StrQuery = "SELECT minwage, maxwage, er_share, ec_share, ee_share FROM SSStable order by id"
Set rsSSSTable = ReturnRst(StrQuery, cnnstr, mstatic)
StrQuery = "SELECT minwage, maxwage, ph_er, ph_ee FROM PhilHealthtable order by id"
Set rsPHealthTable = ReturnRst(StrQuery, cnnstr, mstatic)

'strQuery = "SELECT a.*, b.DailyRate, b.MonthlyRate, b.EcolaAmount, c.NoOfDaysPerMonth, c.VATPercentRate, c.ClientChargePerDay, c.ClientEcolaFlag AS WithEcola FROM ((PayrollMaster a LEFT JOIN HREmployees b ON a.EmployeeID = b.EmployeeID) " & _
'        " LEFT JOIN CutoffTable c ON a.ClientCode = c.ClientCode) WHERE a.clientcode = " & mClientCode & _
'        " and a.startpaydate = '" & txtCutoffDate(2) & "' and a.endpaydate = '" & txtCutoffDate(3) & _
'        "' and format(a.EntryDate, ""mmddyyyy"") = '" & mProcessDate & "'"
        
GetName2 = " b.Surname + " & """, """ & " + b.Firstname + " & """ """ & " + IIF(b.Middlename = """", """", MID(b.Middlename,1,1) + ""."")  AS EmployeeName2 "
        
StrQuery = "SELECT " + GetName2 + ", a.*, b.DailyRate, b.MonthlyRate, b.EcolaAmount, c.* FROM ((PayrollMaster a LEFT JOIN HREmployees b ON a.EmployeeID = b.EmployeeID) " & _
        " LEFT JOIN ClientTable c ON a.ClientCode = c.ClientCode) WHERE a.clientcode = " & mClientCode & _
        " and a.startpaydate = '" & txtCutoffDate(2) & "' and a.endpaydate = '" & txtCutoffDate(3) & _
        "' and format(a.EntryDate, ""mmddyyyy"") = '" & mProcessDate & "'"
        
Set rsPayrollTrans = ReturnRst(StrQuery, cnnstr, mstatic)
If rsPayrollTrans.RecordCount <= 0 Then
    MsgBox "No payroll transaction records selected.", vbInformation, PayrollAlert
    rsPayrollTrans.Close
    GoTo ExitRtn1
End If

With rsPayrollTrans
ListView1(0).ListItems.Clear
While Not (.EOF Or .BOF)
'    Set litem1 = ListView1(0).ListItems.Add(, , !EmployeeID)
'        litem1.SubItems(1) = IIf(IsNull(!EmpNo), "", !EmpNo)
'        litem1.SubItems(2) = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
'        litem1.SubItems(3) = Space(5) 'DTR
'        litem1.SubItems(4) = Format(!DaysPresent, "0")
'        DaysPresent = Format(!DaysPresent, "0")
'
'        litem1.SubItems(5) = Format(IIf(!TaxRefundHours <> 0, (!TaxRefundHours / 8), 0), "0") ' extra days
'        TaxRefundHours = Format(IIf(!TaxRefundHours <> 0, (!TaxRefundHours / 8), 0), "0")
'        litem1.SubItems(6) = Format(IIf(!NoRegHoliday <> 0, (!NoRegHoliday / 8), 0), "0") ' extra days
'        NoRegHoliday = Format(IIf(!NoRegHoliday <> 0, (!NoRegHoliday / 8), 0), "0")
'        litem1.SubItems(7) = Format(IIf(!SundayOT <> 0, (!SundayOT / 8), 0), "0")   ' extra days
'        SundayOT = Format(IIf(!SundayOT <> 0, (!SundayOT / 8), 0), "0")
'        NoSplHoliday = Format(IIf(!NoSplHoliday <> 0, (!NoSplHoliday / 8), 0), "0")
'        litem1.SubItems(8) = Format(NoSplHoliday, "0") ' extra days
'
'        'DaysWorked = Format(DaysPresent + TaxRefundHours + NoRegHoliday + NoSplHoliday, "0")
'        DaysWorked = Format(DaysPresent + TaxRefundHours, "0")
'
'        litem1.SubItems(9) = Format(DaysWorked, "0")  ' # of days worked
'        HoursTardiness = !HoursTardiness
'        litem1.SubItems(10) = Format(HoursTardiness, "0.00")
'        OTHours = Format(!OTHours, "0.00")
'        litem1.SubItems(11) = Format(OTHours, "0.00")
'        ExcessSplOTHours = Format(IIf(IsNull(!excesssplholothours), 0#, !excesssplholothours), "0.00")
'        litem1.SubItems(12) = Format(ExcessSplOTHours, "0.00")
'
'        litem1.SubItems(13) = "-"
'
'        LglHolidayOTHours = !LglHolidayOTHours
'        litem1.SubItems(14) = Format(LglHolidayOTHours, "0.00")
'        ExcessOTSunLegalHrs = !ExcessOTSunLegalHrs
'        litem1.SubItems(15) = Format(ExcessOTSunLegalHrs, "0.00")
'        NightDiffOT = !NightDiffOT
'        litem1.SubItems(16) = Format(NightDiffOT, "0.00")
'        litem1.SubItems(17) = Format(!DailyRate, "0.00")        ' Daily Rate
'        BasicPay = DaysWorked * !DailyRate
'        litem1.SubItems(18) = Format(BasicPay, "0.00")          ' Basic Pay
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
'        TaxRefundPay = !DailyRate * RestDayWorkedOTRate * TaxRefundHours
'        litem1.SubItems(19) = Format(TaxRefundPay, "0.00")     ' Premium pay on extra days worked
'        RegHolidayPay = !DailyRate * SpecialWorkedDayOffRate * NoRegHoliday
'        litem1.SubItems(20) = Format(RegHolidayPay, "0.00")    ' Premium pay on extra days worked
'        SundayOTPay = !DailyRate * LegalHolidayWithPremiumRate * SundayOT
'        litem1.SubItems(21) = Format(SundayOTPay, "0.00")      ' premium pay on extra days worked
'        SplHolidayPay = !DailyRate * LegalHolidayWorkedDayoffRate * NoSplHoliday
'        litem1.SubItems(22) = Format(SplHolidayPay, "0.00")    ' premium pay on extra days worked
'        TardinessDeduction = HoursTardiness * (!DailyRate / WorkHrsPerDay)
'        litem1.SubItems(23) = Format(TardinessDeduction, "0.00")
'
'        OTPay = (!DailyRate / WorkHrsPerDay) * OverTimeRate * OTHours
'        litem1.SubItems(24) = Format(OTPay, "0.00")            ' ot
'        ExcessSplOTPay = (!DailyRate / WorkHrsPerDay) * ExcessSplOTRate * ExcessSplOTHours
'        litem1.SubItems(25) = Format(ExcessSplOTPay, "0.00") ' ot
'        litem1.SubItems(26) = "-"
'        LglHolidayOTPay = (!DailyRate / WorkHrsPerDay) * ExcessLglHolOTRate * LglHolidayOTHours
'        litem1.SubItems(27) = Format(LglHolidayOTPay, "0.00")  ' ot
'        ExcessSunLegalPay = (!DailyRate / WorkHrsPerDay) * ExcessSunLglOTRate * ExcessOTSunLegalHrs
'        litem1.SubItems(28) = Format(ExcessSunLegalPay, "0.00")  ' ot
'        NightDiffPay = (!DailyRate / WorkHrsPerDay) * NightDiffRate * NightDiffOT
'        litem1.SubItems(29) = Format(NightDiffPay, "0.00") ' OT
'
'
'        litem1.SubItems(30) = Format(!EcolaPay, "0.00")
'        OtherIncome = !OtherIncome
'        litem1.SubItems(31) = Format(OtherIncome, "0.00")
'        GrossPay = BasicPay + TaxRefundPay + RegHolidayPay + SundayOTPay + SplHolidayPay + _
'            OTPay + ExcessSplOTPay + LglHolidayOTPay + ExcessSunLegalPay + _
'            ExcessSunLegalPay + NightDiffPay + OtherIncome - TardinessDeduction
'        litem1.SubItems(32) = Format(GrossPay, "0.00")
'        Month13Pay = BasicPay / 12
'        litem1.SubItems(33) = Format(Month13Pay, "0.00")
'
'        SSSEmployer = Format(DaysWorked * EmployerSSSRate(rsSSSTable, (!MonthlyRate + !EcolaAmount), !NoofDaysperMonth), "0.00")
'        litem1.SubItems(34) = Format(SSSEmployer, "0.00")
'        PHealthEmployer = Format(DaysWorked * EmployerPHealthRate(rsPHealthTable, (!MonthlyRate + !EcolaAmount), !NoofDaysperMonth), "0.00")
'        litem1.SubItems(35) = Format(PHealthEmployer, "0.00")
'        PagibigEmployer = (DaysWorked / DaysWorked) * 50
'        litem1.SubItems(36) = Format(PagibigEmployer, "0.00")
'
'        litem1.SubItems(37) = Format(GrossPay + SSSEmployer + PHealthEmployer + PagibigEmployer, "0.00") ' Total pay and benefits
'        CalamityLoan = DaysWorked * AgencyRate
'        litem1.SubItems(38) = Format(CalamityLoan, "0.00") ' Agency charge
'
'        GrossBill = GrossPay + SSSEmployer + PHealthEmployer + PagibigEmployer + CalamityLoan
'        gGrossBill = gGrossBill + GrossBill
'        LgGrossBill = LgGrossBill + GrossBill
'
'        litem1.SubItems(39) = Format(GrossBill, "0.00") ' Gross Billing
'        VatRate = (GrossBill * IIf(IsNull(!VatRate), 0.12, !VatRate))
'        litem1.SubItems(40) = Format(VatRate, "0.00") 'Vat 12%
'        litem1.SubItems(41) = Format(!Allowance, "0.00") 'Gross Payroll
'
'        litem1.SubItems(42) = Format(GrossBill + VATPayment + Allowance, "0.00") ' Total Amount Due
'        litem1.SubItems(43) = ""
'        litem1.SubItems(44) = Format(GrossPay, "0.00") 'Gross Payroll
'        litem1.SubItems(45) = "-"   ' Ignore Lates
'        litem1.SubItems(46) = Format(GrossPay - GrossPay, "0.00") ' Diff'l
'
'        litem1.SubItems(47) = Format(!StartPayDate, "mm/dd/yyyy") ' Start Cutoff
'        litem1.SubItems(48) = Format(!EndPayDate, "mm/dd/yyyy")  ' End Cutoff
'        litem1.SubItems(49) = !EmployeeID
'        ' SSS Rate
'        litem1.SubItems(50) = EmployerSSSRate(rsSSSTable, (!MonthlyRate + !EcolaAmount), !NoofDaysperMonth)
'        ' PhealthRate
'        litem1.SubItems(51) = EmployerPHealthRate(rsPHealthTable, (!MonthlyRate + !EcolaAmount), !NoofDaysperMonth)


    Set litem1 = ListView1(0).ListItems.Add(, , !EmployeeID)
        litem1.SubItems(1) = IIf(IsNull(!EmpNo), "", !EmpNo)
        litem1.SubItems(2) = IIf(IsNull(!EmployeeName2), !EmployeeName, !EmployeeName2)
        'litem1.SubItems(2) = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
        litem1.SubItems(3) = Space(5) 'DTR
        If LTrim(RTrim(!ClientCutoff)) = "Weekly" Then  ' entries is in hours
            litem1.SubItems(4) = Format(!DaysPresent / 8, "0.00")
            DaysPresent = !DaysPresent / 8
        Else
            litem1.SubItems(4) = Format(!DaysPresent, "0.00")
            DaysPresent = !DaysPresent
        End If

        litem1.SubItems(5) = Format(IIf(!TaxRefundHours <> 0, (!TaxRefundHours / 8), 0), "0.00") ' extra days
        TaxRefundHours = Format(IIf(!TaxRefundHours <> 0, (!TaxRefundHours / 8), 0), "0.00")
        litem1.SubItems(6) = Format(IIf(!NoRegHoliday <> 0, (!NoRegHoliday / 8), 0), "0.00") ' extra days
        litem1.SubItems(7) = Format(IIf(!SundayOT <> 0, !SundayOT, 0), "0.00")   ' Excess OT days
        litem1.SubItems(8) = Format(IIf(!NoSplHoliday <> 0, (!NoSplHoliday / 8), 0), "0.00") ' extra days
        
        litem1.SubItems(9) = Format(DaysPresent, "0.00")  ' # of days worked
        litem1.SubItems(10) = Format(!HoursTardiness, "0.00")
        litem1.SubItems(11) = Format(!OTHours, "0.00")
        
        litem1.SubItems(12) = Format(!excessothours, "0.00")
        litem1.SubItems(13) = Format(!Excesssplholothours, "0.00")  ' Previously subitems(12)
        litem1.SubItems(14) = Format(!LglHolidayOTHours, "0.00")
        litem1.SubItems(15) = Format(!ExcessOTSunLegalHrs, "0.00")
        
        litem1.SubItems(16) = Format(!NightDiffOT, "0.00")
        litem1.SubItems(17) = Format(!DailyRate, "0.00")        ' Daily Rate
        litem1.SubItems(18) = Format(!BasicPay, "0.00")         ' Basic Pay
        
        litem1.SubItems(19) = Format(!TaxRefundPay, "0.00")     ' Premium pay on extra days worked
        litem1.SubItems(20) = Format(!RegHolidayPay, "0.00")    ' Premium pay on extra days worked
        litem1.SubItems(21) = Format(!SundayOTPay, "0.00")      ' premium pay on extra days worked
        litem1.SubItems(22) = Format(!SplHolidayPay, "0.00")    ' premium pay on extra days worked
        
        litem1.SubItems(23) = Format(!TardinessDeduction, "0.00")
        litem1.SubItems(24) = Format(!OTPay, "0.00")            ' OT
        
        litem1.SubItems(25) = Format(!excessotpay, "0.00") ' OT
        litem1.SubItems(26) = Format(!excesssplholotpay, "0.00") ' OT previously subitems(25)
        litem1.SubItems(27) = Format(!LglHolidayOTPay, "0.00")  ' OT
        litem1.SubItems(28) = Format(!excessotsunlegalpay, "0.00")  ' ot
        
        litem1.SubItems(29) = Format(!NightDiffPay, "0.00") ' ot

        litem1.SubItems(30) = Format(!EcolaPay, "0.00")
        litem1.SubItems(31) = Format(!OtherIncome, "0.00")
        litem1.SubItems(32) = Format(!GrossPay, "0.00")
        litem1.SubItems(33) = Format(!Month13Payment, "0.00")

        'litem1.SubItems(34) = Format(!SSSEmployer + !ECCemployer, "0.00")
        litem1.SubItems(34) = Format(DaysPresent * EmployerSSSRate(rsSSSTable, (!MonthlyRate + !EcolaAmount), !NoofDaysperMonth), "0.00")
        'litem1.SubItems(35) = Format(!PHealthEmployer, "0.00")
        litem1.SubItems(35) = Format(DaysPresent * EmployerPHealthRate(rsPHealthTable, (!MonthlyRate + !EcolaAmount), !NoofDaysperMonth), "0.00")
        litem1.SubItems(36) = Format(!PagibigEmployer, "0.00")
        ' ECC already included in SSS Employer.
        'litem1.SubItems(37) = Format(!GrossPay + !SSSEmployer + !ECCEmployer + !PHealthEmployer + !PagibigEmployer, "0.00") ' Total pay and benefits
        litem1.SubItems(37) = Format(!GrossPay + !SSSEmployer + !PHealthEmployer + !PagibigEmployer, "0.00")  ' Total pay and benefits
        litem1.SubItems(38) = Format(DaysPresent * !clientchargeperday, "0.00") ' Agency Fee
        'GrossBill = !GrossPay + !SSSEmployer + !ECCEmployer + !PHealthEmployer + !PagibigEmployer + (!DaysPresent * !clientchargeperday)
        GrossBill = !GrossPay + !SSSEmployer + !PHealthEmployer + !PagibigEmployer + (DaysPresent * !clientchargeperday)
        gGrossBill = gGrossBill + GrossBill
        LgGrossBill = LgGrossBill + GrossBill
        'litem1.SubItems(39) = Format(!Grosspay + !SSSemployer + !ECCemployer + !phealthemployer + !pagibigemployer + (!DaysPresent * !ClientChargePerDay), "0.00") ' Gross Billing
        litem1.SubItems(39) = Format(GrossBill, "0.00") ' Gross Billing
        VatRate = (GrossBill * IIf(IsNull(!VatPercentRate), 0.12, !VatPercentRate))
        litem1.SubItems(40) = Format(VatRate, "0.00") 'Vat 12%
        litem1.SubItems(41) = Format(!Allowance, "0.00") 'Gross Payroll
        litem1.SubItems(42) = Format(GrossBill + !Allowance, "0.00") ' Total Amount Due
        litem1.SubItems(43) = ""
        litem1.SubItems(44) = Format(!GrossPay, "0.00") 'Gross Payroll
        litem1.SubItems(45) = "-"   ' Ignore Lates
        litem1.SubItems(46) = Format(!GrossPay - !GrossPay, "0.00") ' Diff'l

        litem1.SubItems(47) = Format(!StartPayDate, "mm/dd/yyyy") ' Start Cutoff
        litem1.SubItems(48) = Format(!EndPayDate, "mm/dd/yyyy")  ' End Cutoff
        litem1.SubItems(49) = !EmployeeID
        ' SSS Rate
        litem1.SubItems(50) = EmployerSSSRate(rsSSSTable, (!MonthlyRate + !EcolaAmount), !NoofDaysperMonth)
        ' PhealthRate
        litem1.SubItems(51) = EmployerPHealthRate(rsPHealthTable, (!MonthlyRate + !EcolaAmount), !NoofDaysperMonth)
        
       
'        strQuery = "UPDATE PayrollTrans SET employeename = '" & EmployeeName & "', agencycode = " & CompanyCode & ", clientcode = " & mClientCode & ", dayspresent = " & ChkVal(txtDaysWork) & ", basicpay = " & ChkVal(TxtAmount(0)) & _
'            ", othours = " & ChkVal(txtRegOT) & ", otpay = " & ChkVal(TxtAmount(1)) & _
'            ", excessothours = " & ChkVal(txtExcessOTRegDay) & ", excessotpay = " & ChkVal(TxtAmount(12)) & _
'            ", nightdiffot = " & ChkVal(txtNightDiff) & ", nightdiffpay = " & ChkVal(TxtAmount(2)) & _
'            ", month13rate = " & ChkVal(txt13thMonth) & ", month13payment = " & ChkVal(TxtAmount(11)) & _
'            ", TaxRefundHours = " & ChkVal(txtRestDay) & ", TaxRefundPay = " & ChkVal(TxtAmount(3)) & _
'            ", sundayot = " & ChkVal(txtSundayOT) & ", sundayotpay = " & ChkVal(TxtAmount(4)) & _
'            ", NoRegHoliday = " & ChkVal(txtSplHoliday) & ", RegHolidayPay = " & ChkVal(TxtAmount(5)) & _
'            ", splholidayothours = " & ChkVal(txtSplOtHours) & ", splholidayotpay = " & ChkVal(TxtAmount(7)) & _
'            ", excesssplholothours = " & ChkVal(txtExcessOTSplSun) & ", excesssplholotpay = " & ChkVal(TxtAmount(13)) & _
'            ", NoSplHoliday = " & ChkVal(txtLglHoliday) & ", SplHolidayPay = " & ChkVal(TxtAmount(6)) & _
'            ", lglholidayothours = " & ChkVal(txtLglHolidayOT) & ", lglholidayotpay = " & ChkVal(TxtAmount(8)) & _
'            ", excesslglholothours = " & ChkVal(txtExcessOTLegal) & ", excesslglholotpay = " & ChkVal(TxtAmount(15)) & _
'            ", excessotsunlegalhrs = " & ChkVal(txtExcessOTSunLgl) & ", excessotsunlegalpay = " & ChkVal(TxtAmount(14)) & _
'            ", otwithoutpremium = " & ChkVal(txtOTwoPremium) & ", otwithoutpremiumpay = " & ChkVal(TxtAmount(16)) & _
'            ", novacationleave = " & ChkVal(txtVacLeave) & _
'            ", vacationleavepay = " & ChkVal(TxtAmount(8)) & ", nosickleave = " & ChkVal(txtSickLeave) & _
'            ", sickleavepay = " & ChkVal(TxtAmount(7)) & ", OtherIncome = " & ChkVal(txtSalAdjustment) & _
'            ", Allowance = " & ChkVal(txtIncentive) & ", incentiveremarks = '" & Trim(txtIncentiveRemarks) & _
'            "', ecolapay = " & ChkVal(txtECOLA) & ", grosspay = " & ChkVal(txtGrossPay) & ", daysabsent = " & _
'            " ChkVal(txtDaysAbsent) & ", absentdeduction = " & ChkVal(txtDeductions(0)) & ", hourstardiness = " & ChkVal(txtHoursLate) & ", tardinessdeduction = " & ChkVal(txtDeductions(1)) & _
'            ", TASAesla = " & ChkVal(txtCashBond) & ", LMCDed = " & ChkVal(txtBDOCard) & ", valewithclient = " & ChkVal(txtVale) & ", CashAdvance = " & ChkVal(txtPlacementFee) & ", miscfee = " & ChkVal(txtMiscFee) & ", overunderpay = " & ChkVal(txtOverPay) & _
'            ", sssemployee = " & ChkVal(txtEmpyeeSSS) & ", sssemployer = " & ChkVal(txtEmpyerSSS) & ", eccemployer = " & Format(mECCAmt, "0.00") & ", phealthemployee = " & ChkVal(txtPhilHealthYee) & ", phealthemployer = " & ChkVal(txtPHealthYer) & ", pagibigemployee = " & ChkVal(txtEmpyeePagibig) & _
'            ", pagibigemployer = " & ChkVal(txtEmpyerPagibig) & ", wtaxamount = " & ChkVal(txtTAX) & ", CalamityLoan = " & ChkVal(txtAgencyFee) & ", sssloan = " & ChkVal(txtSSSLoan) & ", pagibigloan = " & ChkVal(txtPagibigLoan) & ", otherloan = " & Format(mOtherLoan, "0.00") & ", deductions = " & ChkVal(txtDeductions(2)) & ", netpay = " & ChkVal(txtNetPay) & IIf(EdtField = "", "", ", " & EdtField) & _
'            ", startpaydate = '" & Format(frmSelectCutoff.txtCutoffDate(0).Text, "mm/dd/yyyy") & "', endpaydate = '" & Format(frmSelectCutoff.txtCutoffDate(1).Text, "mm/dd/yyyy") & "', monthno = " & Format(frmSelectCutoff.txtCutoffDate(0).Text, "m") & ", cutoffno = " & mCutOffNo & ", payflag1 = 0, payflag2 = 0, paytext1 = '', paytext2 = '', entryby = '" & mUserID & "', entrydate = '" & Format(Now, "mm/dd/yyyy") & "' WHERE EmployeeID = " & EmployeeNo
        
'        'litem1.SubItems(9) = Format(!AdjustedECC, "0.##")
'        litem1.SubItems(9) = Format(!AdjustedEmployeeDed, "0.##")
'        litem1.SubItems(10) = Format(!AdjustedEmployerDed, "0.##")
'        litem1.SubItems(11) = Format(!AmtDifference, "0.##")
'        litem1.SubItems(12) = IIf(IsNull(!ClientName), "Blank", !ClientName)
'
'        litem1.SubItems(13) = IIf(IsNull(!ProcessedBy), "", !ProcessedBy)      ' please activate
'        litem1.SubItems(14) = IIf(IsNull(!ProcessedDate), "", !ProcessedDate)      ' please activate
'
'        litem1.SubItems(16) = IIf(IsNull(!GroupCode), "", !GroupCode)
'        litem1.SubItems(17) = Format(!AmountPaid, "0.00")
'        litem1.SubItems(18) = IIf(IsNull(!SBRNo), "", !SBRNo)
'        litem1.SubItems(19) = IIf(!DatePaid = #1/1/1900#, "__/__/____", Format(!DatePaid, "mm/dd/yyyy"))
'        litem1.SubItems(20) = IIf(!DatePosted = #1/1/1900#, "__/__/____", Format(!DatePosted, "mm/dd/yyyy"))
'        If Not IsNull(!SBRNo) And !SBRNo <> "" And !SBRNo <> "0" Then
'            WithSBRNo = True
'        End If
'        If !ExcludedFlag = True Then
'            ListView1(1).ListItems.Item(i).Checked = True
'        End If
'
'        If !ExcludedFlag = 1 Then
'            frmHDMFReport.ListView1(1).ListItems.Item(i).Checked = True
'            'frmHDMFReport.ListView1(1).ListItems.Checked = True
'        End If
    Listctr1 = Listctr1 + 1
    IncludedCtr = IncludedCtr + 1
    .MoveNext
Wend

End With

txtGrossBill.Text = Format(LgGrossBill, "0.00")
lblEntries(7).Caption = "Total Records: " & IncludedCtr

rsSSSTable.Close
rsPHealthTable.Close
MsgBox "There were " & IncludedCtr & " records selected.", vbInformation, PayrollAlert

ExitRtn1:

End Sub

Private Function EmployerSSSRate(ByRef rs As ADODB.Recordset, _
    ByVal mBasicPay As Double, ByVal NoofDaysperMonth As Integer) As Double

With rs
    .MoveFirst
    Do While Not .EOF
        If mBasicPay >= !minwage And mBasicPay <= !maxwage Then
            EmployerSSSRate = Format((!er_share + !ec_share) / NoofDaysperMonth, "0.00")
            Exit Function
        End If
        .MoveNext
    Loop
End With
EmployerSSSRate = 0#
End Function

Private Function EmployerPHealthRate(ByRef rs As ADODB.Recordset, _
    ByVal mBasicPay As Double, ByVal NoofDaysperMonth As Integer) As Double

With rs
    .MoveFirst
    Do While Not .EOF
        If mBasicPay >= !minwage And mBasicPay <= !maxwage Then
            EmployerPHealthRate = Format((!ph_er / NoofDaysperMonth), "0.00")
            Exit Function
        End If
        .MoveNext
    Loop
End With
EmployerPHealthRate = 0#
End Function

Private Sub IncludeCutOffInBilling()  ' add cutoff for billing
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
Dim ProcDate As String

ProcDate = Mid(cmbCutOff.Text, 23, 8)
mZeroValue = 0#
Set mListView = ListView1(0)
If mListView.ListItems.Count = 0 Then
    GoTo ExitRtn2
End If
If frmBilling.ListView1(0).ListItems.Count = 0 Then
    Listctr = 0
    gGrossBill = 0#
End If

For i = 1 To mListView.ListItems.Count
        Set litem1 = frmBilling.ListView1(0).ListItems.Add(, , mListView.ListItems.Item(i).SubItems(49))
        If mListView.ListItems.Item(i).Checked = True Then  ' get only those without check
            ExcludedFlag = 1
        Else
            ExcludedFlag = 0
        End If
   
            'Set litem1 = ListView1(0).ListItems.Add(, , !EmployeeID)
            litem1.SubItems(1) = mListView.ListItems.Item(i).SubItems(1)    ' Control #
            litem1.SubItems(2) = mListView.ListItems.Item(i).SubItems(2)    ' Employee name
            litem1.SubItems(3) = Space(5) 'DTR
            litem1.SubItems(4) = mListView.ListItems.Item(i).SubItems(4)    ' days present
            litem1.SubItems(5) = mListView.ListItems.Item(i).SubItems(4)    ' edit days present
            'litem1.SubItems(5) = Format(!TaxRefundHours, "0") ' extra days
            litem1.SubItems(6) = mListView.ListItems.Item(i).SubItems(5)    ' rest day ot hours
            litem1.SubItems(7) = mListView.ListItems.Item(i).SubItems(5)    ' edit rest day ot hours
            
            'litem1.SubItems(6) = Format(!NoRegHoliday, "0") ' extra days
            litem1.SubItems(8) = mListView.ListItems.Item(i).SubItems(6)    ' no spl holiday
            litem1.SubItems(9) = mListView.ListItems.Item(i).SubItems(6)    ' edit no spl holiday
            'litem1.SubItems(7) = Format(!SundayOT, "0")     ' extra days
            litem1.SubItems(10) = mListView.ListItems.Item(i).SubItems(7)   ' sunday ot
            litem1.SubItems(11) = mListView.ListItems.Item(i).SubItems(7)   ' edit sunday ot
            'litem1.SubItems(8) = Format(!NoSplHoliday, "0") ' extra days
            
            litem1.SubItems(12) = mListView.ListItems.Item(i).SubItems(8)    ' no lgl holiday
            litem1.SubItems(13) = mListView.ListItems.Item(i).SubItems(8)    ' edit no lgl holiday
            litem1.SubItems(14) = mListView.ListItems.Item(i).SubItems(9)      ' # of days worked
            litem1.SubItems(15) = mListView.ListItems.Item(i).SubItems(9)      ' edit # of days worked
            
            'litem1.SubItems(10) = Format(!HoursTardiness, "0.00")
            litem1.SubItems(16) = mListView.ListItems.Item(i).SubItems(10)       ' hours tardiness
            litem1.SubItems(17) = mListView.ListItems.Item(i).SubItems(10)       ' edit hours tardiness
            'litem1.SubItems(11) = Format(!OTHours, "0.00")
            litem1.SubItems(18) = mListView.ListItems.Item(i).SubItems(11)      ' OT Hours
            litem1.SubItems(19) = mListView.ListItems.Item(i).SubItems(11)      ' edit ot hours
            
            litem1.SubItems(20) = mListView.ListItems.Item(i).SubItems(12)      ' excess spl ot hours
            litem1.SubItems(21) = mListView.ListItems.Item(i).SubItems(12)      ' edit excess spl ot hours
            litem1.SubItems(22) = mListView.ListItems.Item(i).SubItems(13)      ' excess rest day spl
            litem1.SubItems(23) = mListView.ListItems.Item(i).SubItems(14)      ' lgl holiday ot hours
            litem1.SubItems(24) = mListView.ListItems.Item(i).SubItems(14)      ' edit lgl holiday ot hours
            litem1.SubItems(25) = mListView.ListItems.Item(i).SubItems(15)      ' excess ot sun legal hrs
            litem1.SubItems(26) = mListView.ListItems.Item(i).SubItems(15)      ' edit excess ot sun legal hrs
            
            'litem1.SubItems(16) = Format(!NightDiffOT, "0.00")
            litem1.SubItems(27) = mListView.ListItems.Item(i).SubItems(16)      ' night diff ot
            litem1.SubItems(28) = mListView.ListItems.Item(i).SubItems(16)      ' edit night diff ot
            'litem1.SubItems(17) = Format(!DailyRate, "0.00")        ' DailyRate
            litem1.SubItems(29) = mListView.ListItems.Item(i).SubItems(17)        ' DailyRate
            'litem1.SubItems(18) = Format(!BasicPay, "0.00")                        ' BasicPay
            litem1.SubItems(30) = mListView.ListItems.Item(i).SubItems(18)         ' BasicPay
            
            'litem1.SubItems(19) = Format(!TaxRefundPay, "0.00")     ' Premium pay on extra days worked
            litem1.SubItems(31) = mListView.ListItems.Item(i).SubItems(19)     ' Premium pay on extra days worked
            'litem1.SubItems(20) = Format(!RegHolidayPay, "0.00")    ' Premium pay on extra days worked
            litem1.SubItems(32) = mListView.ListItems.Item(i).SubItems(20)    ' Premium pay on extra days worked
            'litem1.SubItems(21) = Format(!SundayOTPay, "0.00")      ' premium pay on extra days worked
            litem1.SubItems(33) = mListView.ListItems.Item(i).SubItems(21)      ' premium pay on extra days worked
            'litem1.SubItems(22) = Format(!SplHolidayPay, "0.00")    ' premium pay on extra days worked
            litem1.SubItems(34) = mListView.ListItems.Item(i).SubItems(22)    ' premium pay on extra days worked
            
            'litem1.SubItems(23) = Format(!TardinessDeduction, "0.00")
            litem1.SubItems(35) = mListView.ListItems.Item(i).SubItems(23)
            'litem1.SubItems(24) = Format(!OTPay, "0.00")            ' reg ot
            litem1.SubItems(36) = mListView.ListItems.Item(i).SubItems(24)
            'litem1.SubItems(25) = Format(!excesssplholotpay, "0.00") ' ot
            litem1.SubItems(37) = mListView.ListItems.Item(i).SubItems(25)
            'litem1.SubItems(26) = "-"
            litem1.SubItems(38) = mListView.ListItems.Item(i).SubItems(26)
            'litem1.SubItems(27) = Format(!lglholidayotpay, "0.00")  ' ot
            litem1.SubItems(39) = mListView.ListItems.Item(i).SubItems(27)
            'litem1.SubItems(28) = Format(!excessotsunlegalpay, "0.00")  ' ot
            litem1.SubItems(40) = mListView.ListItems.Item(i).SubItems(28)
            'litem1.SubItems(29) = Format(!NightDiffPay, "0.00") ' ot
            litem1.SubItems(41) = mListView.ListItems.Item(i).SubItems(29)
            
            'litem1.SubItems(30) = Format(!EcolaPay, "0.00")
            litem1.SubItems(42) = mListView.ListItems.Item(i).SubItems(30)
            'litem1.SubItems(31) = Format(!OtherIncome, "0.00")
            litem1.SubItems(43) = mListView.ListItems.Item(i).SubItems(31)  ' salary adjust
            litem1.SubItems(44) = mListView.ListItems.Item(i).SubItems(31)  ' edit salary adjust
            'litem1.SubItems(32) = Format(!GrossPay, "0.00")
            litem1.SubItems(45) = mListView.ListItems.Item(i).SubItems(32)
            'litem1.SubItems(33) = Format(!Month13Payment, "0.00")
            litem1.SubItems(46) = mListView.ListItems.Item(i).SubItems(33)
            
            'litem1.SubItems(34) = Format(!sssemployer + !ECCemployer, "0.00")
            litem1.SubItems(47) = mListView.ListItems.Item(i).SubItems(34)
            'litem1.SubItems(35) = Format(!PHealthEmployer, "0.00")
            litem1.SubItems(48) = mListView.ListItems.Item(i).SubItems(35)
            'litem1.SubItems(36) = Format(!PagibigEmployer, "0.00")
            litem1.SubItems(49) = mListView.ListItems.Item(i).SubItems(36)
            'litem1.SubItems(37) = Format(!GrossPay + !sssemployer + !ECCemployer + !PHealthEmployer + !PagibigEmployer, "0.00") ' Total pay and benefits
            litem1.SubItems(50) = mListView.ListItems.Item(i).SubItems(37)
            'litem1.SubItems(38) = Format(!DaysPresent * !clientchargeperday, "0.00") ' Agency Fee
            litem1.SubItems(51) = mListView.ListItems.Item(i).SubItems(38)
            'GrossBill = !GrossPay + !sssemployer + !ECCemployer + !PHealthEmployer + !PagibigEmployer + (!DaysPresent * !clientchargeperday)
            'gGrossBill = gGrossBill + GrossBill
            'litem1.SubItems(39) = Format(GrossBill, "0.00") ' Gross Billing
            litem1.SubItems(52) = mListView.ListItems.Item(i).SubItems(39)
            'VatRate = (GrossBill * !VatRate)
            'litem1.SubItems(40) = Format(VatRate, "0.00") 'Vat 12%
            litem1.SubItems(53) = mListView.ListItems.Item(i).SubItems(40)
            'litem1.SubItems(41) = Format(!Allowance, "0.00") ' Incentive
            litem1.SubItems(54) = mListView.ListItems.Item(i).SubItems(41) ' Other Deductions/Incentive
            litem1.SubItems(55) = mListView.ListItems.Item(i).SubItems(41) ' Edit Other Deductions/Allowance
            'litem1.SubItems(42) = Format(GrossBill + !Allowance, "0.00") ' Total Amount Due
            litem1.SubItems(56) = mListView.ListItems.Item(i).SubItems(42)
            gGrossBill = gGrossBill + Val(mListView.ListItems.Item(i).SubItems(42))
            'litem1.SubItems(43) = ""
            litem1.SubItems(57) = ""
            'litem1.SubItems(44) = Format(!GrossPay, "0.00") ' Net Pay
            litem1.SubItems(58) = mListView.ListItems.Item(i).SubItems(44) ' Net Pay
            'litem1.SubItems(45) = "-"   ' Ignore Lates
            litem1.SubItems(59) = "-"   ' Ignore Lates
            'litem1.SubItems(46) = Format(!GrossPay - !GrossPay, "0.00") ' Diff'l
            litem1.SubItems(60) = mListView.ListItems.Item(i).SubItems(46)
        
            litem1.SubItems(61) = mListView.ListItems.Item(i).SubItems(9) * 15 ' Agency Fee
        
            'litem1.SubItems(47) = Format(!StartCutoff, "mm/dd/yyyy") ' Start Cutoff
            litem1.SubItems(62) = mListView.ListItems.Item(i).SubItems(47) ' Start Cutoff
            'litem1.SubItems(48) = Format(!EndCutOff, "mm/dd/yyyy")  ' End Cutoff
            litem1.SubItems(63) = mListView.ListItems.Item(i).SubItems(48)  ' End Cutoff
            'litem1.SubItems(49) = !EmployeeID
            litem1.SubItems(64) = frmBilling.txtSOANo
            
            litem1.SubItems(65) = mListView.ListItems.Item(i).SubItems(50) ' sss rate
            
            litem1.SubItems(66) = mListView.ListItems.Item(i).SubItems(51) ' phealth rate
        
        If ExcludedFlag = 1 Then
            frmBilling.ListView1(0).ListItems.Item(i).Checked = True
        End If
    
        StrQuery = "INSERT INTO BilledTrans (EmployeeID, EmpNo, employeename, agencycode, clientcode, dayspresent, dayspresent2, dailyrate, dailyrate2, basicpay, " & _
            "othours, othours2, otpay, excessothours, excessothours2, excessotpay, nightdiffot, nightdiffot2, nightdiffpay, month13rate, month13payment, " & _
            "TaxRefundHours, TaxRefundHours2, TaxRefundPay, sundayot, sundayot2, sundayotpay, NoRegHoliday, NoRegHoliday2, RegHolidayPay, " & _
            "splholidayothours, splholidayotpay, excesssplholothours, excesssplholotpay, " & _
            "NoSplHoliday, NoSplHoliday2, SplHolidayPay, lglholidayothours, lglholidayothours2, lglholidayotpay, excesslglholothours, excesslglholotpay, excessotsunlegalhrs, excessotsunlegalhrs2, excessotsunlegalpay, otwithoutpremium, otwithoutpremiumpay, novacationleave, " & _
            "vacationleavepay, nosickleave, sickleavepay, OtherIncome, OtherIncome2, Allowance, Allowance2, incentiveremarks, ecolapay, grosspay, daysabsent, absentdeduction, hourstardiness, hourstardiness2, tardinessdeduction, " & _
            "TASAesla, LMCDed, valewithclient, CashAdvance, miscfee, overunderpay, " & _
            "sssemployee, sssemployer, eccemployer, phealthemployee, phealthemployer, pagibigemployee, " & _
            "pagibigemployer, wtaxamount, CalamityLoan, sssloan, pagibigloan, otherloan, deductions, netpay, " & _
            "startpaydate, endpaydate, monthno, entryby, entrydate, ExcludedFlag, SOANo, PayText2, SSSRate, PHRate, CutOffNo)" & _
            " VALUES (" & mListView.ListItems.Item(i).SubItems(49) & ", '" & litem1.SubItems(1) & "', '" & litem1.SubItems(2) & "'," & CompanyCode & ", " & mClientCode & ", " & litem1.SubItems(4) & ", " & litem1.SubItems(5) & ", " & litem1.SubItems(29) & ", " & litem1.SubItems(29) & ", " & litem1.SubItems(30) & _
            ", " & litem1.SubItems(18) & ", " & litem1.SubItems(19) & ", " & litem1.SubItems(36) & ", " & litem1.SubItems(20) & ", " & litem1.SubItems(21) & ", " & litem1.SubItems(37) & ", " & litem1.SubItems(27) & ", " & litem1.SubItems(28) & ", " & litem1.SubItems(41) & ", " & Format(litem1.SubItems(30) / 12, "0.##") & ", " & litem1.SubItems(46) & _
            ", " & litem1.SubItems(6) & ", " & litem1.SubItems(7) & ", " & litem1.SubItems(31) & ", " & litem1.SubItems(10) & ", " & litem1.SubItems(11) & ", " & litem1.SubItems(33) & ", " & litem1.SubItems(8) & ", " & litem1.SubItems(9) & ", " & litem1.SubItems(32) & _
            ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & _
            ", " & litem1.SubItems(12) & ", " & litem1.SubItems(13) & ", " & litem1.SubItems(34) & ", " & litem1.SubItems(23) & ", " & litem1.SubItems(24) & ", " & litem1.SubItems(39) & ", " & mZeroValue & ", " & mZeroValue & ", " & litem1.SubItems(25) & ", " & litem1.SubItems(26) & ", " & litem1.SubItems(40) & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & _
            ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & litem1.SubItems(43) & ", " & litem1.SubItems(44) & ", " & litem1.SubItems(54) & ", " & litem1.SubItems(55) & ", '" & "N/A" & "', " & litem1.SubItems(42) & ", " & litem1.SubItems(45) & ", " & mZeroValue & ", " & mZeroValue & ", " & litem1.SubItems(16) & ", " & litem1.SubItems(17) & ", " & litem1.SubItems(35) & _
            ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & _
            ", " & mZeroValue & ", " & litem1.SubItems(47) & ", " & mZeroValue & ", " & mZeroValue & ", " & litem1.SubItems(48) & ", " & mZeroValue & _
            ", " & litem1.SubItems(49) & ", " & litem1.SubItems(53) & ", " & litem1.SubItems(61) & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & litem1.SubItems(58) & _
            ", '" & Format(litem1.SubItems(62), "mm/dd/yyyy") & "', '" & Format(litem1.SubItems(63), "mm/dd/yyyy") & "', " & Format(litem1.SubItems(62), "m") & ", '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & _
            "', " & ExcludedFlag & ", '" & frmBilling.txtSOANo.Text & "', '" & ProcDate & "', " & litem1.SubItems(65) & ", " & litem1.SubItems(66) & ", " & txtSeqNo.Text & ")"
    
        Call Action_Query(StrQuery, cnnstr)
   
    
'            ", " & ChkVal(txtSplOtHours) & ", " & ChkVal(TxtAmount(7)) & ", " & ChkVal(txtExcessOTSplSun) & ", " & ChkVal(TxtAmount(13))
'            ", " & ChkVal(litem1.SubItems(13)) & ", " & ChkVal(litem1.SubItems(34)) & ", " & ChkVal(litem1.SubItems(24)) & ", " & ChkVal(litem1.SubItems(39)) & ", " & ChkVal(txtExcessOTLegal) & ", " & ChkVal(TxtAmount(15)) & ", " & ChkVal(txtExcessOTSunLgl) & ", " & ChkVal(TxtAmount(14)) & ", " & ChkVal(txtOTwoPremium) & ", " & ChkVal(TxtAmount(16)) & ", " & ChkVal(txtVacLeave)
'            ", " & ChkVal(txtCashBond) & ", " & ChkVal(txtBDOCard) & ", " & ChkVal(txtVale) & ", " & ChkVal(txtPlacementFee) & ", " & ChkVal(txtMiscFee) & ", " & ChkVal(txtOverPay)
'            ", " & ChkVal(txtEmpyeeSSS) & ", " & ChkVal(txtEmpyerSSS) & ", " & Format(mECCAmt, "0.00") & ", " & ChkVal(txtPhilHealthYee) & ", " & ChkVal(txtPHealthYer) & ", " & ChkVal(txtEmpyeePagibig)


'        strQuery = "INSERT INTO BilledTrans (EmployeeID, EmpNo, employeename, agencycode, clientcode, dayspresent, basicpay, " & _
'            "othours, otpay, excessothours, excessotpay, nightdiffot, nightdiffpay, month13rate, month13payment, " & _
'            "TaxRefundHours, TaxRefundPay, sundayot, sundayotpay, NoRegHoliday, RegHolidayPay, " & _
'            "splholidayothours, splholidayotpay, excesssplholothours, excesssplholotpay, " & _
'            "NoSplHoliday, SplHolidayPay, lglholidayothours, lglholidayotpay, excesslglholothours, excesslglholotpay, excessotsunlegalhrs, excessotsunlegalpay, otwithoutpremium, otwithoutpremiumpay, novacationleave, " & _
'            "vacationleavepay, nosickleave, sickleavepay, OtherIncome, Allowance, incentiveremarks, ecolapay, grosspay, daysabsent, absentdeduction, hourstardiness, tardinessdeduction, " & _
'            "TASAesla, LMCDed, valewithclient, CashAdvance, miscfee, overunderpay, " & _
'            "sssemployee, sssemployer, eccemployer, phealthemployee, phealthemployer, pagibigemployee, " & _
'            "pagibigemployer, wtaxamount, CalamityLoan, sssloan, pagibigloan, otherloan, deductions, netpay, " & _
'            "startpaydate, endpaydate, monthno, entryby, entrydate )" & _
'            " VALUES (" & litem1.SubItems(62) & ", '" & litem1.SubItems(1) & "', '" & litem1.SubItems(2) & "'," & CompanyCode & ", " & mClientCode & ", " & ChkVal(litem1.SubItems(5)) & ", " & ChkVal(litem1.SubItems(30)) & _
'            ", " & ChkVal(litem1.SubItems(19)) & ", " & ChkVal(litem1.SubItems(36)) & ", " & ChkVal(litem1.SubItems(21)) & ", " & ChkVal(litem1.SubItems(37)) & ", " & ChkVal(litem1.SubItems(28)) & ", " & ChkVal(litem1.SubItems(41)) & ", " & ChkVal((litem1.SubItems(30) / 12)) & ", " & ChkVal(litem1.SubItems(46)) & _
'            ", " & ChkVal(litem1.SubItems(7)) & ", " & ChkVal(litem1.SubItems(31)) & ", " & ChkVal(litem1.SubItems(11)) & ", " & ChkVal(litem1.SubItems(33)) & ", " & ChkVal(litem1.SubItems(9)) & ", " & ChkVal(litem1.SubItems(32)) & _
'            ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & _
'            ", " & ChkVal(litem1.SubItems(13)) & ", " & ChkVal(litem1.SubItems(34)) & ", " & ChkVal(litem1.SubItems(24)) & ", " & ChkVal(litem1.SubItems(39)) & ", " & mZeroValue & ", " & mZeroValue & ", " & ChkVal(litem1.SubItems(26)) & ", " & ChkVal(litem1.SubItems(40)) & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & _
'            ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & ChkVal(litem1.SubItems(44)) & ", " & ChkVal(litem1.SubItems(54)) & ", '" & "N/A" & "', " & ChkVal(litem1.SubItems(42)) & ", " & ChkVal(litem1.SubItems(45)) & ", " & ChkVal(litem1.SubItems(5)) & ", " & ChkVal((litem1.SubItems(29) * litem1.SubItems(5))) & ", " & ChkVal(litem1.SubItems(17)) & ", " & ChkVal(litem1.SubItems(35)) & _
'            ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & _
'            ", " & mZeroValue & ", " & ChkVal(litem1.SubItems(47)) & ", " & mZeroValue & ", " & mZeroValue & ", " & ChkVal(litem1.SubItems(48)) & ", " & mZeroValue & _
'            ", " & ChkVal(litem1.SubItems(49)) & ", " & ChkVal(litem1.SubItems(53)) & ", " & ChkVal(litem1.SubItems(60)) & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & mZeroValue & ", " & ChkVal(litem1.SubItems(57)) & _
'            ", '" & Format(litem1.SubItems(61), "mm/dd/yyyy") & "', '" & Format(litem1.SubItems(62), "mm/dd/yyyy") & "', " & Format(litem1.SubItems(61), "m") & ", '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & "')"
    
    Listctr = Listctr + 1
Next i
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
  
frmBilling.txtGrossBill.Text = gGrossBill
frmBilling.lblEntries(7).Caption = "Total Records: " & Listctr

MsgBox "There were " & Listctr & " records added for Billing.", vbInformation, PayrollAlert
cmdAddCutoff.Enabled = False
'ListView1(0).ListItems.Clear

End Sub


