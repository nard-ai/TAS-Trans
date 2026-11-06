VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHRMainEmployee 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Details"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   Icon            =   "frmHRMainEmployee.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MousePointer    =   1  'Arrow
   ScaleHeight     =   5895
   ScaleWidth      =   9720
   Begin VB.TextBox cmdCommand 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Height          =   285
      Left            =   8010
      MouseIcon       =   "frmHRMainEmployee.frx":058A
      MousePointer    =   99  'Custom
      TabIndex        =   33
      TabStop         =   0   'False
      Text            =   "Movement entry"
      Top             =   5400
      Width           =   1395
   End
   Begin TabDlg.SSTab SSTabEmpMenu 
      Height          =   3645
      Left            =   90
      TabIndex        =   34
      Top             =   2160
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   6429
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   -2147483638
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Benefits"
      TabPicture(0)   =   "frmHRMainEmployee.frx":0894
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lvwBenefits"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Assignment"
      TabPicture(1)   =   "frmHRMainEmployee.frx":08B0
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lvwAssignment"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Training"
      TabPicture(2)   =   "frmHRMainEmployee.frx":08CC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvwTraining"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Offenses"
      TabPicture(3)   =   "frmHRMainEmployee.frx":08E8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvwOffenses"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Leaves"
      TabPicture(4)   =   "frmHRMainEmployee.frx":0904
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lvwLeaves"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Appraisal"
      TabPicture(5)   =   "frmHRMainEmployee.frx":0920
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lvwAppraisal"
      Tab(5).ControlCount=   1
      Begin MSComctlLib.ListView lvwAssignment 
         Height          =   2655
         Left            =   225
         TabIndex        =   35
         Top             =   540
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Assign. from"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Assign to"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Assignment"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Shift"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Remarks"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComctlLib.ListView lvwTraining 
         Height          =   2655
         Left            =   -74775
         TabIndex        =   36
         Top             =   495
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483639
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Duration"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Training"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Facilitator"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Venue"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Cost"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Start"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "End"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lvwOffenses 
         Height          =   2655
         Left            =   -74775
         TabIndex        =   37
         Top             =   540
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483639
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Offense Made"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nature"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Action"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Remarks"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lvwLeaves 
         Height          =   2610
         Left            =   -74730
         TabIndex        =   38
         Top             =   585
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4604
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483639
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Start date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "End date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Reason"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Remarks"
            Object.Width           =   4057
         EndProperty
      End
      Begin MSComctlLib.ListView lvwAppraisal 
         Height          =   2655
         Left            =   -74730
         TabIndex        =   39
         Top             =   540
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483639
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Evaluated by"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Schedule"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Performance Factor"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Personal Factor"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Total"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Action taken/Remarks"
            Object.Width           =   4057
         EndProperty
      End
      Begin MSComctlLib.ListView lvwBenefits 
         Height          =   2655
         Left            =   -74685
         TabIndex        =   40
         Top             =   540
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483639
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Benefit No"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Benefit"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Acquire date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Remarks"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   420
      Left            =   225
      TabIndex        =   30
      Top             =   5310
      Width           =   6270
      Begin VB.Label Label4 
         Caption         =   "Select then press ""Delete button"" to remove item."
         Height          =   240
         Left            =   180
         TabIndex        =   31
         Top             =   135
         Width           =   4650
      End
   End
   Begin VB.CommandButton btnTop 
      Height          =   255
      Left            =   8595
      Picture         =   "frmHRMainEmployee.frx":093C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Top record"
      Top             =   630
      Width           =   255
   End
   Begin VB.CommandButton btnPrev 
      Height          =   255
      Left            =   8835
      Picture         =   "frmHRMainEmployee.frx":0A86
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Previous record"
      Top             =   630
      Width           =   255
   End
   Begin VB.CommandButton btnNext 
      Height          =   255
      Left            =   9075
      Picture         =   "frmHRMainEmployee.frx":0BD0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Next record"
      Top             =   630
      Width           =   255
   End
   Begin VB.CommandButton btnLast 
      Height          =   255
      Left            =   9315
      Picture         =   "frmHRMainEmployee.frx":0D1A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Last record"
      Top             =   630
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9480
      Top             =   8160
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton btnEmployeeNew 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   0
         Picture         =   "frmHRMainEmployee.frx":0E64
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "New"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton btnEmployeeEdit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   630
         Picture         =   "frmHRMainEmployee.frx":1396
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton btnEmployeeDelete 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   1215
         Picture         =   "frmHRMainEmployee.frx":14E0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton btnEmployeeSearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   2385
         Picture         =   "frmHRMainEmployee.frx":15E2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Search"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton btnEmpRefresh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         Picture         =   "frmHRMainEmployee.frx":16E4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Refresh"
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
   Begin VB.TextBox txtHide 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Height          =   285
      Left            =   8400
      MouseIcon       =   "frmHRMainEmployee.frx":182E
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Text            =   "Hide"
      Top             =   5445
      Visible         =   0   'False
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2040
      Left            =   90
      TabIndex        =   17
      Top             =   720
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   3598
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmHRMainEmployee.frx":1B38
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblEmpStat"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDtHired"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblEmpNm"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblEmpNo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LblPaymentbasis"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "mskDtHired"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtEmpStat"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtClient"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtEmpName"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtControlNo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtRate"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtstartContract"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtEndContract"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxtEmployeeNum"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      Begin VB.TextBox TxtEmployeeNum 
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
         Left            =   3465
         TabIndex        =   32
         Top             =   225
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.TextBox txtEndContract 
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
         Left            =   8010
         TabIndex        =   28
         Top             =   945
         Width           =   1215
      End
      Begin VB.TextBox TxtstartContract 
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
         Left            =   5985
         TabIndex        =   26
         Top             =   945
         Width           =   1215
      End
      Begin VB.TextBox TxtRate 
         Alignment       =   1  'Right Justify
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
         Left            =   8010
         TabIndex        =   24
         Top             =   585
         Width           =   1215
      End
      Begin VB.TextBox txtControlNo 
         Appearance      =   0  'Flat
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
         Left            =   1620
         TabIndex        =   0
         Top             =   225
         Width           =   1575
      End
      Begin VB.TextBox txtEmpName 
         Appearance      =   0  'Flat
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
         Left            =   1620
         TabIndex        =   1
         Top             =   585
         Width           =   3495
      End
      Begin VB.TextBox txtClient 
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
         Left            =   1620
         TabIndex        =   2
         Top             =   945
         Width           =   3495
      End
      Begin VB.TextBox txtEmpStat 
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
         Left            =   5985
         TabIndex        =   4
         Top             =   585
         Width           =   405
      End
      Begin MSMask.MaskEdBox mskDtHired 
         Height          =   285
         Left            =   8010
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   -2147483635
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
         Format          =   "mm/dd/yyYY"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End date"
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
         Left            =   7290
         TabIndex        =   29
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Start date"
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
         Left            =   5175
         TabIndex        =   27
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Label LblPaymentbasis 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Rate"
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
         Left            =   6645
         TabIndex        =   25
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label2 
         Height          =   240
         Left            =   4230
         TabIndex        =   23
         Top             =   45
         Width           =   1500
      End
      Begin VB.Label lblEmpNo 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   225
         TabIndex        =   22
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label lblEmpNm 
         BackColor       =   &H80000005&
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
         Height          =   255
         Left            =   225
         TabIndex        =   21
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label lblDtHired 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   6930
         TabIndex        =   20
         Top             =   315
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
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
         Left            =   180
         TabIndex        =   19
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label lblEmpStat 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   5445
         TabIndex        =   18
         Top             =   630
         Width           =   1095
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   -120
      X2              =   9720
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label lblFiltEmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1875
      TabIndex        =   5
      Top             =   3570
      Visible         =   0   'False
      Width           =   7575
   End
End
Attribute VB_Name = "frmHRMainEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public rs As ADODB.Recordset

Dim listmove As ListItem
Dim ListAssgn As ListItem
Dim ListTrng As ListItem
Dim ListOff As ListItem
Dim ListLeave As ListItem
Dim ListAtnd As ListItem
Dim ListBen As ListItem
Dim ListAppr As ListItem
Dim ListAtndc As ListItem
Dim rsEB  As ADODB.Recordset
Dim rsM  As ADODB.Recordset
Dim rsA  As ADODB.Recordset
Dim rst  As ADODB.Recordset
Dim rsO  As ADODB.Recordset
Dim rsL  As ADODB.Recordset
Dim rsD  As ADODB.Recordset
Dim rsB  As ADODB.Recordset
Dim rsP  As ADODB.Recordset
Dim str, str1, str2, str3, str4, str5, str6, str7 As String
Dim EmpNo As String

Private Sub btnEmployeeDelete_Click(Index As Integer)

Dim rsDelete As ADODB.Recordset

'Set rsDelete = New ADODB.Recordset
'StrQuery = "Select * from employeebasicsalary where EmployeeID = " & GetEmployeeID() & "  and EmployeeStatus = 1 and agencycode=" & CompanyCode
'Set rsDelete = ReturnRst(StrQuery, cnnstr, mstatic)
''rsDelete.Open "Select * from employeebasicsalary where EmployeeID=" & GetEmployeeID() & "  and employeestatus=true and agencycode=" & CompanyCode, cnn, adOpenKeyset, adLockOptimistic
'
'If rsDelete.EOF = False Then
'    MsgBox "This employee is included in Payroll List, cannot be deleted. Mark it inactive in Employee Basic Salary maintenance first.", vbInformation, PayrollAlert
'Exit Sub
'Else
'Call DELEMP ' do not allow deletion here

'End If

End Sub

Private Function GetEmployeeID() As Integer

Dim rsemp As ADODB.Recordset
Set rsemp = New ADODB.Recordset
StrQuery = "Select EmployeeID from hremployees where EmpNo='" & Me.TxtControlNo & "'"
Set rsemp = ReturnRst(StrQuery, cnnstr, mstatic)
'rsemp.Open "Select EmployeeID from hremployees where EmpNo='" & Me.txtControlNo & "'", cnn, adOpenKeyset, adLockOptimistic
If rsemp.EOF = False Then
GetEmployeeID = rsemp.Fields(0)
Else
GetEmployeeID = 0
End If
End Function

Private Sub btnEmployeeEdit_Click(Index As Integer)
    EditFlag = True
    AddFlag = False
    addflagline = False
    addflagbatch = False
    frmHRFinalEmployees.Show
    frmHRFinalEmployees.SetFocus

End Sub




Private Sub btnEmployeeHelp_Click(Index As Integer)

End Sub

Private Sub btnEmployeeNew_Click(Index As Integer)
'frmHREmployees.Show vbModal

AddFlag = True
EditFlag = False
addflagline = False
addflagbatch = False

frmHRFinalEmployees.Show
Call frmHRFinalEmployees.CheckUser1(operatorID)
    
End Sub

Private Sub btnEmployeeSearch_Click(Index As Integer)
MMEmp = True
frmHRSearch.Show ' vbModal
frmHRSearch.SetFocus
End Sub

Private Sub btnEmpRefresh_Click()
'Showdata rs
Dim rsout As New ADODB.Recordset

    If rsout.State = 1 Then rsout.Close
    StrQuery = "Select * from HREmployees where agencycode=" & CompanyCode & " order by surname, firstname"
    Set rsout = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsout.Open "Select * from HREmployees where agencycode=" & CompanyCode & " order by surname, firstname", cnnHR, adOpenKeyset, adLockOptimistic
    If rsout.EOF = False Then
         Showdata rsout
    Else
        MsgBox "No records in HR Database left. Cannot proceed.", vbInformation
        'Unload Me
        Exit Sub
    End If
End Sub

Private Sub btnLast_Click()
    If rs.State = 1 Then
        rs.MoveLast
        Showdata rs
    End If
End Sub

Private Sub btnNext_Click()

If rs.State = 1 Then
    rs.MoveNext

    If rs.EOF Then
        rs.MoveLast
       ' MsgBox "End of file encountered", vbInformation, CompanyName
    End If

    Showdata rs
    
End If

End Sub

Private Sub btnPrev_Click()

If rs.State = 1 Then
     rs.MovePrevious
    
    If rs.BOF Then
        rs.MoveFirst
      '  MsgBox "Beginning of file encountered", vbInformation, CompanyName
    End If
   
    Showdata rs
End If

End Sub

Private Sub btnTop_Click()
    If rs.State = 1 Then
        rs.MoveFirst
        Showdata rs
    End If
End Sub

'Private Sub Command1_Click()
'frmHRFinalEmployees.Show vbModal
'End Sub

Private Sub Form_Load()

MMEmp = True
' On Error GoTo Errh
       Top = 0
       Left = 0
      ' frmHRMainEmployee.Height = 3105
    
 '  Me.Top = 0 '(mdiMain.Height - Me.Height) / 8
  '  Me.Left = (mdiMain.Width - Me.Width) / 2
  '  CheckUser
  
    Set rs = New ADODB.Recordset
    Set rsM = New ADODB.Recordset
    Set rsA = New ADODB.Recordset
    Set rst = New ADODB.Recordset
    Set rsO = New ADODB.Recordset
    Set rsL = New ADODB.Recordset
    Set rsD = New ADODB.Recordset
    Set rsB = New ADODB.Recordset
    Set rsP = New ADODB.Recordset
    
    If frmHRFilter.Filter_var = "" Then
        StrQuery = "Select * from HREmployees where agencycode= " & CompanyCode & " order by surname, firstname "
        Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
        'rs.Open m_str, cnnHR, adOpenKeyset, adLockOptimistic
    Else
        StrQuery = "Select * from HREmployees where " & frmHRFilter.Filter_var & " AND agencycode= " & CompanyCode & " order by surname, firstname "
        Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
        'rs.Open m_str, cnnHR, adOpenKeyset, adLockOptimistic
   End If
    
    If rs.RecordCount <= 0 Then
        MsgBox "No record found!", vbInformation, CompanyName
        'Unload Me
        'additional code jeff 010403
        btnEmployeeNew(1).Enabled = True
        btnEmployeeEdit(1).Enabled = False
        btnEmployeeDelete(0).Enabled = False
        btnEmployeeSearch(1).Enabled = False
        btnEmployeeRefresh.Enabled = False
        btnEmpRefresh.Enabled = True
       ' btnEmployeeFilter(3).Enabled = False
        btnTop.Enabled = False
        btnPrev.Enabled = False
        btnNext.Enabled = False
        btnLast.Enabled = False
        Exit Sub
    End If
    
    Showdata rs
    
    EmpNo = TxtControlNo
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Set rs = Nothing
    Set rsM = Nothing
    Set rsA = Nothing
    Set rst = Nothing
    Set rsO = Nothing
    Set rsL = Nothing
    Set rsD = Nothing
    Set rsB = Nothing
    Set rsP = Nothing
    MMEmp = False
End Sub

Private Sub lvwAppraisal_DblClick()
frmHRAppraisal.Show
frmHRAppraisal.ShowAppraisal (lvwAppraisal.SelectedItem.Text)
End Sub

Private Sub lvwAppraisal_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete Then
        If MsgBox("Are you sure you want to remove the selected row?", vbYesNo + vbQuestion) = vbYes Then
                Call Action_Query("Delete from hremployeesappraisal where appraisalnum= " & lvwAppraisal.SelectedItem.Text & " ", cnnstr)
                lvwAppraisal.ListItems.Remove (lvwAppraisal.SelectedItem.Index)
           
        End If
    End If
End Sub

Private Sub lvwAssignment_DblClick()
frmHRProjAssign.Show
frmHRProjAssign.ShowAssignment (Me.lvwAssignment.SelectedItem.Text)
End Sub

Private Sub lvwAssignment_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete Then
        If MsgBox("Are you sure you want to remove the selected row?", vbYesNo + vbQuestion) = vbYes Then
            Call Action_Query("Delete from hremployeesassignment where assignmentnum= " & lvwAssignment.SelectedItem.Text & " ", cnnstr)
    
            lvwAssignment.ListItems.Remove (lvwAssignment.SelectedItem.Index)
               End If
    End If
End Sub

Private Sub lvwBenefits_DblClick()
frmHRBenefits.Show
frmHRBenefits.showBenefits (lvwBenefits.SelectedItem.Text)

End Sub

Private Sub lvwBenefits_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete Then
        If MsgBox("Are you sure you want to remove the selected row?", vbYesNo + vbQuestion) = vbYes Then
                 Call Action_Query("Delete from hremployeebenefits where id= " & lvwBenefits.SelectedItem.Text & " ", cnnstr)
                     lvwBenefits.ListItems.Remove (lvwBenefits.SelectedItem.Index)
      
        End If
    End If
End Sub

Private Sub lvwLeaves_DblClick()
frmHRLeave.Show
frmHRLeave.ShowLeaves (lvwLeaves.SelectedItem.Text)
End Sub

Private Sub lvwLeaves_ItemClick(ByVal Item As MSComctlLib.ListItem)
'Call Form_Load
End Sub

Private Sub lvwLeaves_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
        If MsgBox("Are you sure you want to remove the selected row?", vbYesNo + vbQuestion) = vbYes Then
               Call Action_Query("Delete from hremployeesLeaves where leavenum= " & lvwLeaves.SelectedItem.Text & " ", cnnstr)
                     lvwLeaves.ListItems.Remove (lvwLeaves.SelectedItem.Index)
      
        End If
    End If

End Sub

Private Sub lvwOffenses_DblClick()
frmHROffense.Show
frmHROffense.ShowOffenses (lvwOffenses.SelectedItem.Text)
End Sub

Private Sub lvwOffenses_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
        If MsgBox("Are you sure you want to remove the selected row?", vbYesNo + vbQuestion) = vbYes Then
               Call Action_Query("Delete from hremployeesOffenses where offensenum= " & lvwOffenses.SelectedItem.Text & " ", cnnstr)
                    lvwOffenses.ListItems.Remove (lvwOffenses.SelectedItem.Index)
          ' lvwOffenses.ListItems.Remove (lvwOffenses.SelectedItem.Index)
      
        End If
    End If

End Sub

Private Sub lvwTraining_DblClick()
frmHRTraining.Show
frmHRTraining.ShowTraining (Me.lvwTraining.SelectedItem.Text)

End Sub

Private Sub lvwTraining_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
        If MsgBox("Are you sure you want to remove the selected row?", vbYesNo + vbQuestion) = vbYes Then
               Call Action_Query("Delete from hremployeesTraining where trainingnum= " & lvwTraining.SelectedItem.Text & " ", cnnstr)
                    lvwTraining.ListItems.Remove (lvwTraining.SelectedItem.Index)
      
        End If
    End If

End Sub

Private Sub SSTabEmpMenu_Click(PreviousTab As Integer)
'If frmHRMainEmployee.Height = 3105 Then frmHRMainEmployee.Height = 6450
If SSTabEmpMenu.Tab = 0 Then
cmdCommand.Text = "Benefits"
ElseIf SSTabEmpMenu.Tab = 1 Then
cmdCommand.Text = "Assignment entry"
ElseIf SSTabEmpMenu.Tab = 2 Then
cmdCommand.Text = "Training"
ElseIf SSTabEmpMenu.Tab = 3 Then
cmdCommand.Text = "Offenses"
ElseIf SSTabEmpMenu.Tab = 4 Then
cmdCommand.Text = "Add Leaves"
ElseIf SSTabEmpMenu.Tab = 5 Then
cmdCommand.Text = "Appraisal"
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_Change()


    
End Sub

Private Sub cmdCommand_Click()
 If SSTabEmpMenu.Tab = 0 Then
       frmHRBenefits.Show
       frmHRBenefits.SetFocus
       Showdata rs
 ElseIf SSTabEmpMenu.Tab = 1 Then
       frmHRProjAssign.Show
       frmHRProjAssign.SetFocus
       Showdata rs
  ElseIf SSTabEmpMenu.Tab = 2 Then
       frmHRTraining.Show
       frmHRTraining.SetFocus
       Showdata rs
  ElseIf SSTabEmpMenu.Tab = 3 Then
       frmHROffense.Show
       frmHROffense.SetFocus
       Showdata rs
  ElseIf SSTabEmpMenu.Tab = 4 Then
       frmHRLeave.Show
       frmHRLeave.SetFocus
       Showdata rs
  ElseIf SSTabEmpMenu.Tab = 5 Then
       frmHRAppraisal.Show
       frmHRAppraisal.SetFocus
       Showdata rs
  End If
End Sub

Public Sub Showdata(ByRef rec As ADODB.Recordset)
    Dim rs As ADODB.Recordset
    
'**************
'show record
'**************
    If rec.State = 0 Then MsgBox "No records Found!", vbInformation, PayrollAlert: Exit Sub
    If rec.EOF = True Then

    Else
        TxtControlNo.Text = IIf(IsNull(rec!EmpNo) = True, "", rec!EmpNo)
        TxtEmployeeNum.Text = rec!EmployeeID
        txtEmpName.Text = rec!Surname & ", " & rec!FirstName
    '   LblUpdateBy.Caption = UCase(GETuSERNAME(IIf(IsNull(rec!USERID), "0", rec!USERID)))
        Me.Caption = "EMPLOYEES DETAILS   [Updated by: " + UCase(GETuSERNAME(IIf(IsNull(rec!USERID), "0", rec!USERID))) + "]"
             
        mskDtHired.Text = IIf(IsDate(rec!datehired), Format(rec!datehired, "mm/dd/yyyy"), "__/__/____")
        txtEmpStat.Text = IIf(IsNull(rec!employmentstatus), "", rec!employmentstatus)
    End If
        
'    If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'        StrQuery = "Select * From Clienttable Where CLIENTCODE2 = 'FMS' AND clientCode = " & IIf(IsNull(rec!ClientCode), 0, rec!ClientCode)
'    Else
'        StrQuery = "Select * From Clienttable Where clientCode = " & IIf(IsNull(rec!ClientCode), 0, rec!ClientCode)
'    End If
    StrQuery = GetClientRecord(IIf(IsNull(rec!ClientCode), 0, rec!ClientCode))
    Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
        
    'Set rs = cnn.Execute("Select * From Clienttable Where clientCode =" & IIf(IsNull(rec!ClientCode), 0, rec!ClientCode))
    If Not rs.EOF Then
        txtClient.Text = rs.Fields(1)
    Else
        txtClient.Text = ""
    End If
    rs.Close
               
  Info rec!EmployeeID
    
    Set rs = Nothing
    Set rsM = Nothing
    Set rsA = Nothing
    Set rst = Nothing
    Set rsO = Nothing
    Set rsL = Nothing
    Set rsD = Nothing
    Set rsB = Nothing
    Set rsP = Nothing
    
End Sub

Private Sub Info(EmpNo As String)

'**************
'show employee information
'**************
   'start here
    StrQuery = "Select * From HREMPLOYEES Where EmployeeID=" & EmpNo
    Set rsEB = ReturnRst(StrQuery, cnnstr, mstatic)
    'Set rsEB = cnn.Execute("Select * From HREMPLOYEES Where EmployeeID=" & EmpNo & "")
        
        If Not rsEB.EOF Then
        LblPaymentbasis = rsEB.Fields("paymentbasis") & "Rate"
          If LCase(rsEB.Fields("paymentbasis")) = "hourly" Then
          txtRate.Text = Format(rsEB.Fields("hourlyrate"), "# ##0.00")
          ElseIf LCase(rsEB.Fields("paymentbasis")) = "daily" Then
          txtRate.Text = Format(rsEB.Fields("dailyrate"), "# ##0.00")
          ElseIf LCase(rsEB.Fields("paymentbasis")) = "monthly" Then
            txtRate.Text = Format(rsEB.Fields("monthlyrate"), "# ##0.00")
          End If
            Me.TxtstartContract = IIf(IsNull(rsEB.Fields("startcontract")), "", rsEB.Fields("startcontract"))
            Me.txtEndContract = IIf(IsNull(rsEB.Fields("endcontract")), "", rsEB.Fields("endcontract"))
        Else
            txtRate.Text = ""
            Me.TxtstartContract = ""
           Me.txtEndContract = ""
        End If
        rsEB.Close
    'end here

    'Assignment
    str1 = "Select * From HREmployeesAssignment where EmployeeID='" & EmpNo & "'"
    Set rsA = ReturnRst(str1, cnnstr, mstatic)
        'Set rsA = cnn.Execute(str1)
        lvwAssignment.ListItems.Clear
        Do While Not rsA.EOF
            Set ListAssgn = lvwAssignment.ListItems.Add(, , rsA!assignmentnum)
             ListAssgn.SubItems(1) = IIf(IsNull(rsA!assigndate), "", rsA!assigndate)
            ListAssgn.SubItems(2) = IIf(IsNull(rsA!endassigndate), "", rsA!endassigndate)
            ListAssgn.SubItems(3) = IIf(IsNull(rsA!AssignType), "", rsA!AssignType)
            ListAssgn.SubItems(4) = IIf(IsNull(rsA!shiftsched), "", rsA!shiftsched)
            ListAssgn.SubItems(5) = IIf(IsNull(rsA!assignremarks), "", rsA!assignremarks)
  
        rsA.MoveNext
        Loop
        rsA.Close

    'Training
    str2 = "Select * From HREmployeesTraining where EmployeeID='" & EmpNo & "'"
        Set rst = cnn.Execute(str2)
        lvwTraining.ListItems.Clear
        Do While Not rst.EOF
            Set ListTrng = lvwTraining.ListItems.Add(, , rst!trainingnum)
            ListTrng.SubItems(1) = IIf(IsNull(rst!trainingtype), "", rst!trainingtype)
            ListTrng.SubItems(2) = IIf(IsNull(rst!trainingduration), "", rst!trainingduration)
            ListTrng.SubItems(3) = IIf(IsNull(rst!training), "", rst!training)
            ListTrng.SubItems(4) = IIf(IsNull(rst!facilitator), "", rst!facilitator)
            ListTrng.SubItems(5) = IIf(IsNull(rst!trainingvenue), "", rst!trainingvenue)
            ListTrng.SubItems(6) = IIf(IsNull(rst!trainingcost), "", rst!trainingcost)
            ListTrng.SubItems(7) = IIf(IsNull(rst!starttrainingcontract), "", rst!starttrainingcontract)
            ListTrng.SubItems(8) = IIf(IsNull(rst!endtrainingcontract), "", rst!endtrainingcontract)
        rst.MoveNext
        Loop
        rst.Close
 
    'Offenses
    str3 = "Select * From HREmployeesOffenses where EmployeeID='" & EmpNo & "'"
        Set rsO = cnn.Execute(str3)
        lvwOffenses.ListItems.Clear
        Do While Not rsO.EOF
            Set ListOff = lvwOffenses.ListItems.Add(, , rsO!offensenum)
            ListOff.SubItems(1) = IIf(IsNull(rsO!offensemade), "", rsO!offensemade)
            ListOff.SubItems(2) = IIf(IsNull(rsO!offensetype), "", rsO!offensetype)
            ListOff.SubItems(3) = IIf(IsNull(rsO!offensedate), "", rsO!offensedate)
            ListOff.SubItems(4) = IIf(IsNull(rsO!actiontype), "", rsO!actiontype)
            ListOff.SubItems(5) = IIf(IsNull(rsO!Remarks), "", rsO!Remarks)
             rsO.MoveNext
        Loop
        rsO.Close

    'Leaves
    str4 = "Select * From HREmployeesLeaves where EmployeeID='" & EmpNo & "'"
        Set rsL = cnn.Execute(str4)
        lvwLeaves.ListItems.Clear
        Do While Not rsL.EOF
            Set ListLeave = lvwLeaves.ListItems.Add(, , rsL!leavenum)
            ListLeave.SubItems(1) = IIf(IsNull(rsL!leavetype), "", rsL!leavetype)
           ' ListLeave.SubItems(2) = IIf(IsNull(rsL!leaveserved), "", rsL!leaveserved)
            ListLeave.SubItems(2) = IIf(IsNull(rsL!leavedatefr), "", rsL!leavedatefr)
            ListLeave.SubItems(3) = IIf(IsNull(rsL!leavedateto), "", rsL!leavedateto)
            ListLeave.SubItems(4) = IIf(IsNull(rsL!leaversn), "", rsL!leaversn)
            'ListLeave.SubItems(6) = IIf(IsNull(rsL!applied), "", rsL!applied)
            ListLeave.SubItems(5) = IIf(IsNull(rsL!Remarks), "", rsL!Remarks)
            'ListLeave.SubItems(8) = IIf(IsNull(rsL!noted), "", rsL!noted)
           'ListLeave.SubItems(5) = IIf(IsNull(rsL!leavedapprovedby), "", IIf(rsL!leavedapprovedby, "Yes", "No"))
        rsL.MoveNext
        Loop
        rsL.Close
    'Benefits
       str6 = "Select * From qryHRBenefits where EmployeeID='" & EmpNo & "'"
        Set rsB = cnn.Execute(str6)
        lvwBenefits.ListItems.Clear
        Do While Not rsB.EOF
            Set ListBen = lvwBenefits.ListItems.Add(, , rsB!id)
           'ListBen.SubItems(1) = IIf(IsNull(rsB!id), "", rsB!id)
            ListBen.SubItems(1) = IIf(IsNull(rsB!benefitnum), "", rsB!benefitnum)
            ListBen.SubItems(2) = IIf(IsNull(rsB!benefit), "", rsB!benefit)
            ListBen.SubItems(3) = IIf(IsNull(rsB!benefitacquiredate), "", rsB!benefitacquiredate)
            ListBen.SubItems(4) = IIf(IsNull(rsB!benefitremarks), "", rsB!benefitremarks)

        rsB.MoveNext
        Loop
        rsB.Close

    'Appraisal
        str7 = "Select * From HREmployeesAppraisal where EmployeeID='" & EmpNo & "'"
        Set rsP = cnn.Execute(str7)
        lvwAppraisal.ListItems.Clear
        Do While Not rsP.EOF
            Set ListAppr = lvwAppraisal.ListItems.Add(, , rsP!appraisalnum)
            ListAppr.SubItems(1) = IIf(IsNull(rsP!appraisaldate), "", rsP!appraisaldate)
            ListAppr.SubItems(2) = IIf(IsNull(rsP!evalby), "", rsP!evalby)
            ListAppr.SubItems(3) = IIf(IsNull(rsP!appraisalsched), "", rsP!appraisalsched)
            ListAppr.SubItems(4) = IIf(IsNull(rsP!performancefactor), "", rsP!performancefactor)
            ListAppr.SubItems(5) = IIf(IsNull(rsP!personalfactor), "", rsP!personalfactor)
            ListAppr.SubItems(6) = IIf(IsNull(rsP!Total), "", rsP!Total)
            ListAppr.SubItems(7) = IIf(IsNull(rsP!actiontaken), "", rsP!actiontaken)
            rsP.MoveNext
        Loop
        rsP.Close

End Sub

Private Sub DELEMP()
 Dim sqlString As String
 res = MsgBox("You are about to delete 1 record." & vbCrLf & " " & vbCrLf & "If you click Yes, You won't able to undo this delete operation. " & vbCrLf & "Are you sure you want to delete the Employee No: " & TxtControlNo & "?", vbQuestion + vbYesNo, CompanyName)
  If res = vbYes Then
    sqlString = "Delete from HREmployees where EmployeeID=" & GetEmployeeID() & "":        Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HRDetails where EmployeeID=" & GetEmployeeID() & "":       Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HRCharReference where EmployeeID='" & GetEmployeeID() & "'":         Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HRDialects where EmployeeID='" & GetEmployeeID() & "'":          Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HREmploymentRecord where EmployeeID='" & GetEmployeeID() & "'":      Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HRSeminar where EmployeeID='" & TxtControlNo & "'":         Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HREmployeesAssignment where EmployeeID='" & GetEmployeeID() & "'":        Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HREmployeesLeaves where EmployeeID='" & GetEmployeeID() & "'":         Call Action_Query(sqlString, cnnstr)
    'sqlString = "Delete from HREmployeesMovement where EmployeeID='" & GetEmployeeID() & "'":      Call Action_Query(sqlString, Cnnstr)
    sqlString = "Delete from HREmployeesOffenses where EmployeeID='" & GetEmployeeID() & "'":       Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HREmployeesTraining where EmployeeID='" & GetEmployeeID() & "'":      Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HREmployeeBenefits where EmployeeID='" & GetEmployeeID() & "'":       Call Action_Query(sqlString, cnnstr)
     sqlString = "Delete from HREmployeesAppraisal where EmployeeID='" & GetEmployeeID() & "'":       Call Action_Query(sqlString, cnnstr)
  End If


   MsgBox "The record of [" & txtEmpName.Text & "] was deleted.", vbInformation
      Call btnEmpRefresh_Click
   '     Showdata rs
End Sub

Private Sub txtHide_Click()
If frmHRMainEmployee.Height = 6450 Then frmHRMainEmployee.Height = 3105
End Sub
