VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmReportCustomized 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Asia Report Manager"
   ClientHeight    =   5010
   ClientLeft      =   2730
   ClientTop       =   2565
   ClientWidth     =   10110
   Icon            =   "frmReportHired.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   10110
   Begin TabDlg.SSTab SSTab1 
      Height          =   4785
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   8440
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   -2147483638
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "  "
      TabPicture(0)   =   "frmReportHired.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSort(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmFilter"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmHistory"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SSTab2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdShow"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtTitle"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   3660
         TabIndex        =   39
         Top             =   3870
         Width           =   4515
         Begin VB.CheckBox chkSort 
            Caption         =   "Descending"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   2835
            TabIndex        =   0
            Top             =   270
            Width           =   1230
         End
         Begin VB.ComboBox cboSort 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   1080
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   225
            Width           =   1755
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Sort by"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   1
            Top             =   240
            Width           =   1230
         End
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4755
         TabIndex        =   10
         Text            =   "S E A R C H   A S I A   R E P O R T "
         Top             =   465
         Width           =   4650
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Sho&w Report"
         Height          =   315
         Left            =   8295
         TabIndex        =   9
         Top             =   4095
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "SELECT NAME"
         Height          =   900
         Left            =   3600
         TabIndex        =   7
         Top             =   4545
         Visible         =   0   'False
         Width           =   5865
         Begin VB.ComboBox CboName 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   990
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   270
            Width           =   4455
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4470
         Left            =   225
         TabIndex        =   2
         Top             =   180
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   7885
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Report List"
         TabPicture(0)   =   "frmReportHired.frx":05A6
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstVw"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Selection"
         TabPicture(1)   =   "frmReportHired.frx":05C2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Lv"
         Tab(1).ControlCount=   1
         Begin VB.ListBox lstVw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3900
            ItemData        =   "frmReportHired.frx":05DE
            Left            =   135
            List            =   "frmReportHired.frx":05E0
            TabIndex        =   3
            Top             =   450
            Width           =   2850
         End
         Begin MSComctlLib.ListView Lv 
            Height          =   3855
            Left            =   -74910
            TabIndex        =   4
            Top             =   450
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   6800
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   3528
            EndProperty
         End
      End
      Begin VB.Frame frmHistory 
         Height          =   3030
         Left            =   3600
         TabIndex        =   41
         Top             =   840
         Visible         =   0   'False
         Width           =   5865
         Begin VB.CheckBox chkStatus2 
            Caption         =   "AWOL"
            Height          =   285
            Index           =   8
            Left            =   3960
            TabIndex        =   60
            Top             =   2520
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus2 
            Caption         =   "EOC"
            Height          =   285
            Index           =   7
            Left            =   3960
            TabIndex        =   58
            Top             =   2280
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus2 
            Caption         =   "Early EOC"
            Height          =   285
            Index           =   6
            Left            =   3960
            TabIndex        =   57
            Top             =   2040
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus2 
            Caption         =   "Resigned"
            Height          =   285
            Index           =   5
            Left            =   2280
            TabIndex        =   56
            Top             =   2520
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus2 
            Caption         =   "Terminated"
            Height          =   285
            Index           =   4
            Left            =   2280
            TabIndex        =   55
            Top             =   2280
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus2 
            Caption         =   "Regular"
            Height          =   285
            Index           =   3
            Left            =   2280
            TabIndex        =   54
            Top             =   2040
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus2 
            Caption         =   "Probationary"
            Height          =   285
            Index           =   2
            Left            =   600
            TabIndex        =   53
            Top             =   2520
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus2 
            Caption         =   "Contractual"
            Height          =   285
            Index           =   1
            Left            =   600
            TabIndex        =   52
            Top             =   2280
            Value           =   1  'Checked
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus2 
            Caption         =   "Project Base"
            Height          =   285
            Index           =   0
            Left            =   600
            TabIndex        =   51
            Top             =   2040
            Width           =   1230
         End
         Begin VB.ComboBox cmbYearof 
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmReportHired.frx":05E2
            Left            =   4260
            List            =   "frmReportHired.frx":0607
            TabIndex        =   48
            Text            =   "cmbYearof"
            Top             =   840
            Width           =   915
         End
         Begin VB.ComboBox cmbMonthof 
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmReportHired.frx":064D
            Left            =   1800
            List            =   "frmReportHired.frx":0675
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   840
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.ComboBox cmbClient 
            Height          =   315
            ItemData        =   "frmReportHired.frx":06DB
            Left            =   1740
            List            =   "frmReportHired.frx":06DD
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   360
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.ComboBox cboHistoryName 
            Height          =   315
            ItemData        =   "frmReportHired.frx":06DF
            Left            =   2040
            List            =   "frmReportHired.frx":06E1
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   1320
            Width           =   3375
         End
         Begin MSMask.MaskEdBox mskStartCutoff 
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
            Left            =   1440
            TabIndex        =   49
            Top             =   1560
            Visible         =   0   'False
            Width           =   1305
            _ExtentX        =   2302
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
         Begin MSMask.MaskEdBox mskEndCutoff 
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
            Left            =   4080
            TabIndex        =   59
            Top             =   1560
            Visible         =   0   'False
            Width           =   1305
            _ExtentX        =   2302
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
         Begin VB.Label Label2 
            Caption         =   "End Cutoff:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   62
            Top             =   1560
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Start Cutoff:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   61
            Top             =   1560
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblSort 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Year:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   3600
            TabIndex        =   50
            Top             =   960
            Width           =   465
         End
         Begin VB.Label lblSort 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payroll Month:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   47
            Top             =   960
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label lblSort 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   44
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblSort 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   360
            TabIndex        =   43
            Top             =   1440
            Width           =   1425
         End
      End
      Begin VB.Frame frmFilter 
         Height          =   3030
         Left            =   3600
         TabIndex        =   11
         Top             =   855
         Width           =   5865
         Begin VB.ComboBox cboEmpStatus 
            Height          =   315
            Left            =   3330
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1080
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.TextBox lblCon 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   3
            Left            =   2295
            TabIndex        =   33
            Text            =   "Equal to"
            Top             =   1575
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox lblCon 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   2
            Left            =   2295
            TabIndex        =   32
            Text            =   "Equal to"
            Top             =   1170
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox lblCon 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   1
            Left            =   2295
            TabIndex        =   31
            Text            =   "Equal to"
            Top             =   765
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.ComboBox cboFilterby 
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            ItemData        =   "frmReportHired.frx":06E3
            Left            =   360
            List            =   "frmReportHired.frx":06E5
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1485
            Width           =   1950
         End
         Begin VB.ComboBox cboFilterby 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            ItemData        =   "frmReportHired.frx":06E7
            Left            =   360
            List            =   "frmReportHired.frx":06E9
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1080
            Width           =   1950
         End
         Begin VB.ComboBox cboFilterby 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "frmReportHired.frx":06EB
            Left            =   360
            List            =   "frmReportHired.frx":06ED
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   675
            Width           =   1950
         End
         Begin VB.TextBox lblCon 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   0
            Left            =   2295
            TabIndex        =   24
            Text            =   "Equal to"
            Top             =   315
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.ComboBox cboFilterby 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "frmReportHired.frx":06EF
            Left            =   360
            List            =   "frmReportHired.frx":06F1
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   270
            Width           =   1950
         End
         Begin VB.ComboBox CboCompany 
            Height          =   315
            Left            =   3330
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   270
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Contractual"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   20
            Top             =   2385
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Project Base"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   19
            Top             =   2115
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Probationary"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   18
            Top             =   2655
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Regular"
            Height          =   255
            Index           =   3
            Left            =   1980
            TabIndex        =   17
            Top             =   2115
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Terminated"
            Height          =   255
            Index           =   4
            Left            =   1980
            TabIndex        =   16
            Top             =   2385
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Resigned"
            Height          =   255
            Index           =   5
            Left            =   1980
            TabIndex        =   15
            Top             =   2655
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Early EOC"
            Height          =   255
            Index           =   6
            Left            =   3690
            TabIndex        =   14
            Top             =   2115
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "EOC"
            Height          =   255
            Index           =   7
            Left            =   3690
            TabIndex        =   13
            Top             =   2385
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "AWOL"
            Height          =   255
            Index           =   8
            Left            =   3690
            TabIndex        =   12
            Top             =   2655
            Width           =   1590
         End
         Begin MSMask.MaskEdBox mskDate1 
            Height          =   285
            Index           =   0
            Left            =   3330
            TabIndex        =   35
            Top             =   675
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
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
         Begin MSMask.MaskEdBox mskDate2 
            Height          =   285
            Index           =   0
            Left            =   3330
            TabIndex        =   36
            Top             =   1485
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   4210752
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
         Begin MSMask.MaskEdBox mskDate1 
            Height          =   285
            Index           =   1
            Left            =   4590
            TabIndex        =   37
            Top             =   675
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   4210752
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
         Begin MSMask.MaskEdBox mskDate2 
            Height          =   285
            Index           =   1
            Left            =   4590
            TabIndex        =   38
            Top             =   1485
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   4210752
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
         Begin VB.CheckBox Check1 
            Caption         =   "(1) Filter by"
            Height          =   285
            Index           =   0
            Left            =   135
            TabIndex        =   23
            Top             =   270
            Width           =   1185
         End
         Begin VB.CheckBox Check1 
            Caption         =   "(3) Filter by"
            Height          =   285
            Index           =   1
            Left            =   135
            TabIndex        =   28
            Top             =   675
            Width           =   1230
         End
         Begin VB.CheckBox Check1 
            Caption         =   "(2) Filter by"
            Height          =   285
            Index           =   2
            Left            =   135
            TabIndex        =   29
            Top             =   1080
            Width           =   1230
         End
         Begin VB.CheckBox Check1 
            Caption         =   "(4) Filter by"
            Height          =   285
            Index           =   3
            Left            =   135
            TabIndex        =   30
            Top             =   1485
            Width           =   1230
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   5850
            Y1              =   1935
            Y2              =   1935
         End
      End
      Begin VB.Label lblSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Title :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3540
         TabIndex        =   5
         Top             =   510
         Width           =   1140
      End
   End
End
Attribute VB_Name = "FrmReportCustomized"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
      ' cmbdepartment.AddItem rstClient!clientname
      ' cmbdepartment.ItemData(cmbdepartment.NewIndex) = rstClient!clientcode
        
Option Explicit
Dim i As Integer
Public rsItems As ADODB.Recordset
Public SELECT1, SELECT2, SELECT3, SELECT4, SELECT5, SELECTED, whereCon As String
Public SELECT1TITLE, SELECT2TITLE, SELECT3TITLE, SELECT4TITLE, SELECT5TITLE As String
Dim mClientCode As Integer
Dim StrQuery As String
Dim DTRFlag As Boolean
Dim EmpStatus As String
Dim ForWhat As String

Public Function SearchFilterby(ByVal ndx As Integer) As Boolean
Dim X As Integer
For X = 0 To cboFilterby.Count - 1
If ndx <> X And cboFilterby(ndx).Text <> "" Then
 If cboFilterby(ndx).Text = cboFilterby(X).Text Then cboFilterby(ndx).ListIndex = -1: SearchFilterby = False: Exit Function
End If
Next X

SearchFilterby = True
End Function

Private Sub cboFilterby_Click(Index As Integer)
    If SearchFilterby(Index) = False Then MsgBox "Filter field is already selected.", vbInformation: Exit Sub
    Hidefield (Index)
    If cboFilterby(Index).ListIndex = 0 Then 'CLIENT NAME
             CboCompany.Clear
             
             Dim rstclient As ADODB.Recordset
             
'            If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'                 StrQuery = "select distinct clientname,clientcode from clienttable WHERE CLIENTCODE2 = 'FMS' AND clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
'            Else
'                 StrQuery = "select distinct clientname,clientcode from clienttable WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
'            End If
            StrQuery = GetClientRecord(0, False, " ORDER BY ClientName", " distinct clientname, clientcode ")
            Set rstclient = ReturnRst(StrQuery, cnnstr, mstatic)
             
             'Set rstclient = cnn.Execute("select distinct clientname,clientcode from clienttable  WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname")
             Do While Not rstclient.EOF
                    CboCompany.AddItem rstclient!ClientName
                      CboCompany.ItemData(CboCompany.NewIndex) = rstclient!ClientCode
                    rstclient.MoveNext
             Loop
             rstclient.Close
             CboCompany.Top = cboFilterby(Index).Top: lblCon(Index).Visible = True: CboCompany.Visible = True
             lblCon(Index).Visible = True: lblCon(Index).Text = "equal"
             
    ElseIf cboFilterby(Index).ListIndex = 1 Then 'DATEHIRED
         mskDate1(0).Visible = True: mskDate1(0).Top = cboFilterby(Index).Top
         mskDate1(1).Visible = True: mskDate1(1).Top = cboFilterby(Index).Top
         lblCon(Index).Visible = True: lblCon(Index).Text = "between"
    
    ElseIf cboFilterby(Index).ListIndex = 2 Then 'END CONTRACT
         mskDate2(0).Visible = True: mskDate2(0).Top = cboFilterby(Index).Top
         mskDate2(1).Visible = True: mskDate2(1).Top = cboFilterby(Index).Top
         lblCon(Index).Visible = True: lblCon(Index).Text = "between"
    
    ElseIf cboFilterby(Index).ListIndex = 3 Then 'STATUS
         Me.cboEmpStatus.Clear

        
         Dim rstStatus As ADODB.Recordset
         Set rstStatus = cnn.Execute("select distinct TaxExemptioncode,DESCRIPTION from TaxExemptioncode")
         Do While Not rstStatus.EOF
                cboEmpStatus.AddItem rstStatus!TaxExemptionCode & "   - " & rstStatus!Description
                 ' cboEmpStatus.ItemData(cboEmpStatus.NewIndex) = rstClient!employmentstatus
                rstStatus.MoveNext
         Loop
         rstStatus.Close
         cboEmpStatus.Top = cboFilterby(Index).Top: lblCon(Index).Visible = True: cboEmpStatus.Visible = True
         lblCon(Index).Visible = True: lblCon(Index).Text = "equal"
    End If
End Sub

    Private Function Hidefield(ByVal ndx As Integer)
    If mskDate1(0).Visible = True And mskDate1(0).Top = cboFilterby(ndx).Top Then mskDate1(0).Visible = False
    If mskDate1(1).Visible = True And mskDate1(1).Top = cboFilterby(ndx).Top Then mskDate1(1).Visible = False
    If mskDate2(0).Visible = True And mskDate2(0).Top = cboFilterby(ndx).Top Then mskDate2(0).Visible = False
    If mskDate2(1).Visible = True And mskDate2(1).Top = cboFilterby(ndx).Top Then mskDate2(1).Visible = False
    If CboCompany.Visible = True And CboCompany.Top = cboFilterby(ndx).Top Then CboCompany.Visible = False
    If cboEmpStatus.Visible = True And cboEmpStatus.Top = cboFilterby(ndx).Top Then cboEmpStatus.Visible = False
    End Function
    
    Private Sub HIDEALL()
         mskDate1(0).Visible = False
         mskDate1(1).Visible = False
         mskDate2(0).Visible = False
         mskDate2(1).Visible = False
         Dim X As Integer
         For X = 0 To lblCon.Count - 1
         lblCon(X).Visible = False
         Next X
         CboCompany.Visible = False
         cboEmpStatus.Visible = False
    End Sub

Private Sub CboName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
Call FindMatch(CboName, KeyAscii)
End Sub

Private Sub Check1_Click(Index As Integer)
    If cboFilterby(Index).Enabled = True Then
        cboFilterby(Index).Enabled = False
        cboFilterby(Index).BackColor = &H8000000A
        If Index = 0 Then CHECKINDEX (Index)
        If Index = 1 Then CHECKINDEX (Index)
        If Index = 2 Then CHECKINDEX (Index)
        If Index = 3 Then CHECKINDEX (Index)
        cboFilterby(Index).ListIndex = -1
    Else
        If Check1(Index).VALUE = 1 Then cboFilterby(Index).Enabled = True
        If Check1(Index).VALUE = 1 Then cboFilterby(Index).BackColor = vbWhite
    End If
End Sub

Private Sub CHECKINDEX(ByVal Index As Integer)
    If cboFilterby(Index).ListIndex = 0 Then
        Me.CboCompany.Visible = False: lblCon(Index).Visible = False
    ElseIf cboFilterby(Index).ListIndex = 1 Then
        Me.mskDate1(0).Visible = False: Me.mskDate1(1).Visible = False: lblCon(Index).Visible = False
    ElseIf cboFilterby(Index).ListIndex = 2 Then
        Me.mskDate2(0).Visible = False: Me.mskDate2(1).Visible = False: lblCon(Index).Visible = False
    ElseIf cboFilterby(Index).ListIndex = 3 Then
        Me.cboEmpStatus.Visible = False: lblCon(Index).Visible = False
    End If
End Sub

Private Sub Check2_Click()
   If cboSort(0).Enabled = True Then
        cboSort(0).Enabled = False
        cboSort(0).BackColor = &H8000000A
        cboSort(0).ListIndex = -1
        chkSort(0).Enabled = False
    Else
        cboSort(0).Enabled = True
        cboSort(0).BackColor = vbWhite
        chkSort(0).Enabled = True
    End If
End Sub

Private Sub chkStatus_Click(Index As Integer)
If Index = 9 Then
Dim X As Integer
For X = 0 To 8
    If chkStatus(Index).VALUE = 1 Then
        chkStatus(X).VALUE = 1
    Else
        chkStatus(X).VALUE = 0
    End If
Next X
End If
End Sub

Private Sub Form_Load()
Me.Caption = UCase(CompanyName) & " REPORT MANAGER"
txtTitle.Text = UCase(CompanyName)

    whereCon = ""
    Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Dim rsAssgnList As ADODB.Recordset
    Call LoadListReport
    Set rsAssgnList = New ADODB.Recordset
       sqlString = "SELECT FIELDNAME, FIELDOBJECT1 FROM tblCustomized"
       rsAssgnList.Open sqlString, cnnstr, adOpenDynamic, adLockOptimistic
       PopulateList LV, rsAssgnList
       LV.ColumnHeaders(1).Width = 2500  '2000
       LV.ColumnHeaders(2).Width = 0
       LV.ListItems.Item(1).Checked = True
       LV.ListItems.Item(2).Checked = True
     Dim ndx As Integer
     For ndx = 0 To cboFilterby.Count - 1
       cboFilterby(ndx).BackColor = &H8000000A
     Next ndx

'lstVw.SetFocus
'QryHrHiredREport = "SELECT DISTINCT HREmployees.EmpNo, HREmployees.EmployeeID, HREmployees.EmploymentStatus, HREmploymentStatus.EmploymentStatusDesc, HREmployees.Surname+', '+HREmployees.FirstName +' '+IIf(IsNull(HREmployees.middlename) Or HREmployees.middlename='','',Left(HREmployees.middlename,1)+'.') AS EmpName, ClientTable.CutOffPeriod, HREmployees.Age, HREmployees.Height, HREmployees.Birthplace, HREmployees.TelNum1, " & _
'                    " HREmployees.TelNum2 , HRDetails.SpouseName, HREmployees.TaxCode, HREmployees.Religion, HREmployees.Address1, HREmployees.Philhealth, HREmployees.TIN, HREmployees.SSS, HREmployees.Pagibig, ClientTable.ClientName, HREmployees.EmpPosition, HREmployees.DateHired, HREmployees.Contact, HREmployees.CAddress, HREmployees.Cphone, HREmployees.MonthlyRate, HREmployees.DailyRate, HREmployees.HourlyRate, HREmployees.StartContract, " & _
'                    " HREmployees.EndContract, HREmployees.AcctNo, ClientTable.ClientCutoffDates, ClientTable.PickUpofDTR, ClientTable.SalaryRelease, ClientTable.BillingSchedule, EmployeeBasicSalary.ClientwServiceContract " & _
'                    ",EmployeeBasicSalary.ClientwCompPackage, EmployeeBasicSalary.IfNoCompensationPackage " & _
'                    " FROM (((HREmployees LEFT JOIN EmployeeBasicSalary ON HREmployees.EmployeeID = EmployeeBasicSalary.EmployeeID) LEFT JOIN ClientTable ON HREmployees.ClientCode = ClientTable.ClientCode) LEFT JOIN HRDetails ON HREmployees.EmpNo = HRDetails.EmployeeID) LEFT JOIN HREmploymentStatus ON HREmployees.EmploymentStatus = HREmploymentStatus.EmploymentStatus "

QryHrHiredREport = "SELECT * FROM qryHREmployeeHired"

  Dim strHistory As String
    strHistory = "select (surname + ', ' + firstname + ' ' + Left(middlename,1) + '.') as NAME,EmployeeID,EmpNo AS [CONTROL NO] ,STARTCONTRACT,ENDCONTRACT from HREMPLOYEES where agencycode=" & CompanyCode & " order by surname, firstname"
             cboHistoryName.Clear
             Dim rstName As ADODB.Recordset
             Set rstName = ReturnRst(strHistory, CnnstrHR, mstatic)
             
             Do While Not rstName.EOF
                    cboHistoryName.AddItem IIf(IsNull(rstName!Name), "No Name", rstName!Name)
                    cboHistoryName.ItemData(cboHistoryName.NewIndex) = rstName!EmployeeID
                    rstName.MoveNext
             Loop
             rstName.Close
             
'If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'    StrQuery = "SELECT * FROM Clienttable WHERE CLIENTCODE2 = 'FMS' AND clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
'Else
'    StrQuery = "SELECT * FROM Clienttable WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
'End If
StrQuery = GetClientRecord(0, False, " ORDER BY ClientName")
'StrQuery = "SELECT * FROM Clienttable WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
Set rstName = ReturnRst(StrQuery, cnnstr, mstatic)

Call oCommonFunction.FillCombo(cmbClient, rstName, 9) ' 9 means add "All Clients"
If rstName.RecordCount <> 0 Then cmbClient.ListIndex = 0
rstName.Close
    
cmbMonthof.Text = Format(Date, "mmmm")
cmbYearof.Text = Format(Date, "yyyy")

mskStartCutoff = Format(Date - 15, "mm/dd/yyyy")
mskEndCutoff = Format(Date, "mm/dd/yyyy")

End Sub

'Private Sub cboFilterby_Click()
'    If cboFilterby.Text = "ClientName" Then
'        Text1.Text = "is Equal to"
'        CboCompany.Visible = True
'        mskDate1.Visible = False
'        mskDate2.Visible = False
'        LoadCompany
'    ElseIf cboFilterby.Text = "DateHired" Then
'        Text1.Text = "is Between"
'        CboCompany.Visible = False
'        mskDate1.Visible = True
'        mskDate2.Visible = True
'    ElseIf cboFilterby.Text = "EndContract" Then
'        Text1.Text = "is Between"
'        CboCompany.Visible = False
'        mskDate1.Visible = True
'        mskDate2.Visible = True
'    End If
'End Sub

Private Sub cboSort_Click(Index As Integer)
  chkSort(Index).Enabled = cboSort(Index).ListIndex > 0
End Sub

Private Sub cboSort_GotFocus(Index As Integer)
 cboSort(0).Clear
If SSTab2.Tab = 0 Then
    With cboSort(0)
    .AddItem "(None)"
    If SELECT1 = "" Then
       SELECT1 = "EmpNo"
    End If
    
       .AddItem SELECT1
       .ItemData(.NewIndex) = 1
    
    If SELECT2 = "" Then
         SELECT2 = "Empname"
    End If
         
         .AddItem SELECT2
         .ItemData(.NewIndex) = 2
    
    If SELECT3 <> "" Then
         .AddItem SELECT3
         .ItemData(.NewIndex) = 3
    End If
    
    If SELECT4 <> "" Then
         .AddItem SELECT4
         .ItemData(.NewIndex) = 4
    End If
      
    If SELECT5 <> "" Then
         .AddItem SELECT5
         .ItemData(.NewIndex) = 5
    End If
    End With
Else

Call CheckField
  
     For i = 0 To cboSort.Count - 1
        With cboSort(i)
          .AddItem "(None)"
          .AddItem SELECT1
         .ItemData(.NewIndex) = 1
          .AddItem SELECT2
         .ItemData(.NewIndex) = 2
          .AddItem SELECT3
          .ItemData(.NewIndex) = 3
          .AddItem SELECT4
          .ItemData(.NewIndex) = 4
          .AddItem SELECT5
          .ItemData(.NewIndex) = 5
          .ListIndex = 0
        End With
      Next
End If

End Sub

'section 1 is detail
'section2 is header

Public Sub LoadListReport()
    lstVw.Clear
    lstVw.AddItem "Information Sheet"
    lstVw.AddItem "Employee List(Company, Branch, Dep't)"
    lstVw.AddItem "List of Pagibig Members"
    lstVw.AddItem "List of SSS Members"
    lstVw.AddItem "List of Philhealth Members"
    lstVw.AddItem "List of TIN Members"
    lstVw.AddItem "End of Contract Report"
    lstVw.AddItem "Resigned Employees"
    lstVw.AddItem "Employees Account Nos."
    lstVw.AddItem "Employees Birthday"
    lstVw.AddItem "Employees Rate"
    lstVw.AddItem "Tax Exception"
    lstVw.AddItem "Masterlist Excel Report "
'    lstVw.AddItem "DTR Summary"
'    lstVw.AddItem "3 Months Summary"
    lstVw.AddItem "Employees Head Count Report"
    lstVw.AddItem "Employment History"
    lstVw.AddItem "Summary OT and ND Report"
    lstVw.AddItem "Employee OT and ND Report"

 For i = 0 To cboFilterby.Count - 1
           cboFilterby(i).AddItem "Client name"
           cboFilterby(i).AddItem "Start of contract"
           cboFilterby(i).AddItem "End of contract"
           cboFilterby(i).AddItem "Civil status"
           '   CboCompany.ItemData(CboCompany.NewIndex) = rstClient!clientcode
 Next
End Sub

Private Function WhereConn() As String
Dim StrQuery As String
Dim str As String
 
whereCon = ""
 
If CboCompany.Visible = True Then
    whereCon = " clientname= '" & CboCompany.Text & "'"
Else

End If

'not include the confidential data of company with ClientFlag2=true
'start
If (SSTab2.Tab = 1 And LV.ListItems.Item(7).Checked = True) Or (SSTab2.Tab = 1 And LV.ListItems.Item(8).Checked = True) Or (SSTab2.Tab = 1 And LV.ListItems.Item(9).Checked = True) Or (SSTab2.Tab = 1 And LV.ListItems.Item(10).Checked = True) Or (SSTab2.Tab = 0 And Me.lstVw.ListIndex = 0) Or (SSTab2.Tab = 0 And Me.lstVw.ListIndex = 10) Or (SSTab2.Tab = 0 And Me.lstVw.ListIndex = 12) Then
If whereCon = "" Then
    Set rsClientTable = New ADODB.Recordset
    
    If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
        StrQuery = "SELECT * FROM clienttable WHERE CLIENTCODE2 = 'FMS' AND clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' and ClientFlag2=1 ORDER BY clientname "
    Else
        StrQuery = "SELECT * FROM clienttable WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' and ClientFlag2=1 ORDER BY clientname "
    End If
    Set rsClientTable = ReturnRst(StrQuery, cnnstr, mstatic)
    
    Do While rsClientTable.EOF = False
    If str <> "" Then str = str & ",'" & rsClientTable.Fields("clientname") & "'" Else str = "'" & rsClientTable.Fields("clientname") & "'"
        rsClientTable.MoveNext
    Loop
    If str <> "" Then whereCon = " clientname not IN " & "(" & str & ")"
    If whereCon <> "" Then
    If MsgBox("Cannot include the company " & str & " in the report because this is tagged as confidential(hidden salary)." & vbCrLf & "If you want to see this report you must be super user." & vbCrLf & "" & vbCrLf & "Do you want to continue without this company?", vbYesNo + vbQuestion) = vbNo Then
    whereCon = "EXIT"
    Exit Function
    End If
    End If
Else
    Set rsClientTable = New ADODB.Recordset
    If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
        StrQuery = "SELECT * FROM clienttable WHERE CLIENTCODE2 = 'FMS' AND clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' and  clientname= '" & CboCompany.Text & "' and ClientFlag2=1 ORDER BY clientname "
    Else
        StrQuery = "SELECT * FROM clienttable WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' and  clientname= '" & CboCompany.Text & "' and ClientFlag2=1 ORDER BY clientname "
    End If
    'StrQuery = "SELECT * FROM clienttable WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' and  clientname= '" & CboCompany.Text & "' and ClientFlag2=1 ORDER BY clientname "
    Set rsClientTable = ReturnRst(StrQuery, cnnstr, mstatic)
    
    If rsClientTable.EOF = False Then
    'str = "'" & rsClientTable.Fields("clientname") & "'"
    MsgBox "Unable to display the report because the company name [" & CboCompany.Text & "] tagged as confidential(hidden salary). " & vbCrLf & "If you want to see this report you must be super user.", vbInformation
    whereCon = "EXIT"
    Exit Function
    End If
    'If str <> "" Then whereCon = " clientname <> " & "(" & str & ")"
End If
End If

'end

 If cboEmpStatus.Visible = True Then
    If whereCon = "" Then
        whereCon = "TaxCode= '" & mID(cboEmpStatus.Text, 1, 3) & "'"
    Else
        whereCon = whereCon & " and TaxCode ='" & Trim(mID(cboEmpStatus.Text, 1, 3)) & "'"
    End If
 End If

 If mskDate1(0).Visible = True And mskDate1(1).Visible = True Then
     If whereCon = "" Then
        whereCon = "  StartContract between  '" & mskDate1(0).Text & "' and '" & mskDate1(1).Text & "'"
     Else
        whereCon = whereCon & " and  startcontract between '" & mskDate1(0).Text & "' and '" & mskDate1(1).Text & "'"
     End If
  End If
  
 If mskDate2(0).Visible = True And mskDate2(1).Visible = True Then
      
    If whereCon = "" Then
    whereCon = " endcontract between  '" & mskDate2(0).Text & "' and '" & mskDate2(1).Text & "'"
    Else
    whereCon = whereCon & " and endcontract between  '" & mskDate2(0).Text & "' and '" & mskDate2(1).Text & "'"
    End If
 End If

ForWhat = ""
str = ""
If chkStatus(0).VALUE = 1 Or chkStatus2(0).VALUE = 1 Then
    str = "'B'"
    ForWhat = ForWhat & "For Project Based"
End If

If chkStatus(1).VALUE = 1 Or chkStatus2(1).VALUE = 1 Then
    If str <> "" Then str = str & ",'C'" Else str = "'C'"
    If ForWhat <> "" Then ForWhat = ForWhat & ", Contractual" Else ForWhat = "For Contractual"
End If

If chkStatus(2).VALUE = 1 Or chkStatus2(2).VALUE = 1 Then
    If str <> "" Then str = str & ",'P'" Else str = "'P'"
    If ForWhat <> "" Then ForWhat = ForWhat & ", Probationary" Else ForWhat = "For Probationary"
End If

If chkStatus(3).VALUE = 1 Or chkStatus2(3).VALUE = 1 Then
    If str <> "" Then str = str & ",'R'" Else str = "'R'"
    If ForWhat <> "" Then ForWhat = ForWhat & ", Regular" Else ForWhat = "For Regular"
End If

If chkStatus(4).VALUE = 1 Or chkStatus2(4).VALUE = 1 Then
    If str <> "" Then str = str & ",'T'" Else str = "'T'"
    If ForWhat <> "" Then ForWhat = ForWhat & ", Terminated" Else ForWhat = "For Terminated"
End If

If chkStatus(5).VALUE = 1 Or chkStatus2(5).VALUE = 1 Then
    If str <> "" Then str = str & ",'X'" Else str = "'X'"
    If ForWhat <> "" Then ForWhat = ForWhat & ", Resigned" Else ForWhat = "For Resigned"
End If

If chkStatus(6).VALUE = 1 Or chkStatus2(6).VALUE = 1 Then
    If str <> "" Then str = str & ",'Y'" Else str = "'Y'"
    If ForWhat <> "" Then ForWhat = ForWhat & ", Early EOC" Else ForWhat = "For Early EOC"
End If

If chkStatus(7).VALUE = 1 Or chkStatus2(7).VALUE = 1 Then
    If str <> "" Then str = str & ",'Z'" Else str = "'Z'"
    If ForWhat <> "" Then ForWhat = ForWhat & ", EOC" Else ForWhat = "For End of Contract"
End If

If chkStatus(8).VALUE = 1 Or chkStatus2(8).VALUE = 1 Then
    If str <> "" Then str = str & ",'W'" Else str = "'W'"
    If ForWhat <> "" Then ForWhat = ForWhat & ", AWOL" Else ForWhat = "For AWOL"
End If

If lstVw.ListIndex <> 7 And lstVw.ListIndex <> 16 Then
    If str = "" Then MsgBox "Please select employment status.", vbInformation: whereCon = "EXIT": Exit Function
End If

EmpStatus = str ' for DTR where condition

If whereCon <> "" Then
    'whereCon = whereCon & " AND HREmployees.EMPLOYMENTSTATUS IN " & "(" & str & ")"
    whereCon = whereCon & " AND EMPLOYMENTSTATUS IN " & "(" & str & ")"
Else
    'whereCon = " HREmployees.EMPLOYMENTSTATUS IN " & "(" & str & ")"
    whereCon = " EMPLOYMENTSTATUS IN " & "(" & str & ")"
End If

If whereCon <> "" Then: WhereConn = " where " & whereCon: whereCon = " where " & whereCon
 
End Function

Public Function CheckFieldDAte() As Boolean
If mskDate1(0).Visible = True And mskDate1(0).Text = "__/__/____" Then GoTo exx
If mskDate1(1).Visible = True And mskDate1(1).Text = "__/__/____" Then GoTo exx
If mskDate2(0).Visible = True And mskDate2(0).Text = "__/__/____" Then GoTo exx
If mskDate2(1).Visible = True And mskDate2(1).Text = "__/__/____" Then GoTo exx
CheckFieldDAte = True
Exit Function
exx:
MsgBox "Fill up the blank field.", vbInformation: CheckFieldDAte = False:

End Function

Private Function ChangeChkStatus(ByVal what As Boolean)
Dim X As Integer
For X = 0 To 8 '9
chkStatus(X).Enabled = what
Next X
End Function

Private Function HideCheckStatus(ByVal what As Boolean)
Dim X As Integer

For X = 0 To 8 '9
    chkStatus2(X).Visible = what
Next X

End Function

Private Sub cmdShow_Click()

'On Error Resume Next
If CheckFieldDAte = False Then Exit Sub
    
If frmFilter.Visible = True And lstVw.ListIndex <> 13 Then ' and not equal to DTR summary (listindex 13)
If cboFilterby(0).Enabled = False Or cboFilterby(0).ListIndex = -1 Then
    If MsgBox("Are you sure you want generate report without filtering?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
End If
Else
If cboHistoryName.ListIndex = -1 And lstVw.ListIndex <> 13 And lstVw.ListIndex <> 16 And lstVw.ListIndex <> 15 Then
    If MsgBox("Are you sure you want generate report without filtering?", vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
    End If
End If
End If

Call WhereConn
If whereCon = "EXIT" Then Exit Sub
    If SSTab2.Tab = 1 Then
        If isCheckExceed = False Then MsgBox "No. of fields selected maybe less than 5 or exceeded . 5 fields are required.", vbInformation: Exit Sub
        Call CheckField
    Else
        If lstVw.ListIndex = -1 Then MsgBox "Select Type of Report!", vbInformation: Exit Sub
        SELECT1 = "EmpNo"
        SELECT2 = "Empname"
        SELECT3 = "Empposition"
        SELECT4 = "ClientName"
        
        SELECT1TITLE = "Emp. No."
        SELECT2TITLE = "Name"
        SELECT3TITLE = "Position"
        SELECT4TITLE = "Client Name"
    
    If lstVw.ListIndex = 0 Then 'Information Sheet
    Dim Sortby, ByDirection As String
        ByDirection = IIf(chkSort(0).VALUE = 1, "Desc", "Asc")
  
  If cboSort(0).Text <> "(None)" And Trim(cboSort(0).Text) <> "" Then Sortby = " order by " & cboSort(0).Text & " " & ByDirection
        StrQuery = "Select * FROM qryHrEmployeeHired " & whereCon & Sortby
        Set rsItems = ReturnRst(StrQuery, cnnstr, mstatic)
        'Set rsItems = ReturnRst(QryHrHiredREport & whereCon & Sortby, cnnstr, mstatic)
        If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
         
         Dim tmpimg As RptImage
        Set tmpimg = rptReportHired.Sections(2).Controls("ImgLogo")
        On Error GoTo b:
        'Set tmpimg.Picture = LoadPicture(App.Path & "\Logos\" & CompanyCode & " logo.gif")
        If LoadPictureFlag = True Then
            Set tmpimg.Picture = LoadPicture(logoPath)
        End If
b:
        
        With rptReportHired
        Set .DataSource = rsItems
             .DataMember = rsItems.DataMember
            .Sections("Section2").Controls("lblAddress").Caption = getAgencyAddress(CompanyCode) ' "Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
             .Show
        End With
 
        Exit Sub
    ElseIf lstVw.ListIndex = 1 Then 'Employees List(Company, Branch, Dep't)
        SELECT3 = "ClientName"
        SELECT3TITLE = "Client Name"
        SELECT4 = "Brname"
        SELECT4TITLE = "Branch"
        SELECT5 = "Deptname"
        SELECT5TITLE = "Department"
    '    WhereCon = ""
    ElseIf lstVw.ListIndex = 2 Then 'List of Pagibig Members
        SELECT5 = "Pagibig"
        SELECT5TITLE = "Pagibig"
     '   WhereCon = ""
    ElseIf lstVw.ListIndex = 3 Then 'List of SSS Members
        SELECT5 = "SSS"
        SELECT5TITLE = "SSS"
      '  WhereCon = ""
  ElseIf lstVw.ListIndex = 4 Then 'List of Philhealth Members
        SELECT5 = "PHILHEALTH"
        SELECT5TITLE = "Philhealth"
       ' WhereCon = ""
    ElseIf lstVw.ListIndex = 5 Then 'List of TIN Members
        SELECT5 = "TIN"
        SELECT5TITLE = "T.I.N"
        'WhereCon = ""
    ElseIf lstVw.ListIndex = 6 Then 'End of Contract Report
        SELECT3 = "ClientName"
        SELECT3TITLE = "Client Name"
        SELECT4 = "StartContract"
        SELECT4TITLE = "Start of Contract"
        SELECT5 = "Endcontract"
        SELECT5TITLE = "End of Contract"
        'WhereCon = ""
    ElseIf lstVw.ListIndex = 7 Then 'Resigned Employees
        SELECT3 = "ClientName"
        SELECT3TITLE = "Client Name"
        SELECT4 = "Dateresign"
        SELECT4TITLE = "Date resigned"
        SELECT5 = "ReasonRes"
        SELECT5TITLE = "Reason"
        chkStatus(5).VALUE = 1
       
        If whereCon = "" Then
         whereCon = " where employmentstatus= 'X' "
         Else
         whereCon = whereCon & " and employmentstatus= 'X' "
         End If
    ElseIf lstVw.ListIndex = 8 Then 'Employees Account nos.
        SELECT5 = "Acctno"
        SELECT5TITLE = "Account no."
        'WhereCon = ""
    ElseIf lstVw.ListIndex = 9 Then 'Employees Birthday
        SELECT5 = "Birthdate"
        SELECT5TITLE = "Birthday"
       ' WhereCon = ""
     ElseIf lstVw.ListIndex = 10 Then 'Employees Rate
        SELECT3 = "ClientName"
        SELECT3TITLE = "Client Name"
        SELECT4 = "Dailyrate"
        SELECT4TITLE = "Daily Rate"
        SELECT5 = "MonthlyRate"
        SELECT5TITLE = "Monthly Rate"
    ElseIf lstVw.ListIndex = 11 Then 'Tax Exception
        SELECT3 = "EmpCivilStatus"
        SELECT3TITLE = "Exemption code"
        SELECT4 = "Description"
        SELECT4TITLE = "Description"
        SELECT5 = "ExemptionAmount"
        SELECT5TITLE = "Exemption Amount"
    ElseIf lstVw.ListIndex = 12 Then 'Masterlist Excel
       Call CreateMasterList
       Exit Sub
    ElseIf lstVw.ListIndex = 15 Or lstVw.ListIndex = 16 Then 'DTR Smmary updated by Bernard ' 13
        If lstVw.ListIndex = 15 Then
            Call OvertimeReport(False)
        Else
            Call OvertimeReport(True)
        End If
        Exit Sub
        'Call CreateDTRReport 'DTR Smmary updated by Bernard
        'Exit Sub
    ElseIf lstVw.ListIndex = 14 Then '3 months summary
        SELECT3 = "ClientName"
        SELECT3TITLE = "Client Name"
        SELECT4 = "StartContract"
        SELECT4TITLE = "Start of Contract"
        SELECT5 = "Endcontract"
        SELECT5TITLE = "End of Contract"
        
        Dim startC, endC As String
        startC = Date - 92
        startC = month(startC) & "/1/" & Year(startC)
        endC = month(startC) & "/31/" & Year(startC)
        If IsDate(endC) = False Then
        endC = month(startC) & "/30/" & Year(startC)
        End If
                
        If IsDate(endC) = False Then
        endC = month(startC) & "/29/" & Year(startC)
        End If
        
        
        If IsDate(endC) = False Then
        endC = month(startC) & "/28/" & Year(startC)
        End If
        'endC = startC + 30
    
       If whereCon = "" Then
            whereCon = " where startcontract between  '" & startC & "' and '" & endC & "'"
         Else
            whereCon = whereCon & " and startcontract between '" & startC & "' and '" & endC & "'"
         End If
       ElseIf lstVw.ListIndex = 13 Then 'Create Employee Count
            Call CreateEmployeesCount
            Exit Sub
       ElseIf lstVw.ListIndex = 15 Then 'Create OT and ND
            Call OvertimeReport(False)
            Exit Sub
       ElseIf lstVw.ListIndex = 16 Then 'Create OT and ND per employee
            Call OvertimeReport(True)
            Exit Sub
 ElseIf lstVw.ListIndex = 14 Then 'Masterlist Excel
       Call CreateEmploymentHistory
       Exit Sub
    End If
    
    End If
    SELECTED = SELECT1 & "," & SELECT2 & "," & SELECT3 & "," & SELECT4 & "," & SELECT5
    SELECTED = Replace(SELECTED, "controlno", "empno")
    ' If CompanyCode = 1 Then
    ' Call ShowReport_sapi
     Call ShowCustomizedReport
'     ElseIf CompanyCode = 2 Then
'     Call ShowReport_jci
'     ElseIf CompanyCode = 3 Then
'     Call ShowReport_cspi
'    End If
     
End Sub

' Updated by Bernard 10/27/2008

Private Function CreateDTRReport()

Dim str As String
Dim rsTemp As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet1 As Excel.Worksheet
Dim ctr As Long
Dim i As Long
Dim StrQuery As String
Dim Sortby, ByDirection As String

Dim CutOffPeriod As String
Dim SalaryRelease As String

'  ByDirection = IIf(chkSort(0).VALUE = 1, "Desc", "Asc")
'
'  If cboSort(0).Text <> "(None)" And Trim(cboSort(0).Text) <> "" Then Sortby = " order by " & cboSort(0).Text & " " & ByDirection
'
'  If Sortby = "" Then Sortby = " order by Empname"
'    'SELECTED = " Empname,EmpNo,Brname,Empposition,Telnum1,Telnum2,startcontract,endcontract,sss,philhealth,TIN,birthdate,TaxCode,Dailyrate,ecolaamount,employmentstatusdesc "
'    'sqlString = "SELECT " & _
'              "" & SELECTED & "" & _
'               " From vwEmployeeCustomInfo " & whereCon & Sortby
'               sqlString = "sELECT * FROM vwEmployeeCustomInfo " & whereCon & Sortby
'Sortby = ""
'
'Set rsTemp = ReturnRst(sqlString, cnn)
'If rsTemp.EOF = False Then
'    'Set xlApp1 = New Excel.Application
'    'Set xlBook1 = xlApp1.Workbooks.Add()
'
'    'xlApp1.Visible = True
'    Call ShowDtrReport(rsTemp, 4)
'    'xlApp1.DisplayAlerts = False
'
'Else
'    MsgBox "No Records!!!", vbCritical, ""
'End If

' New DTR Report Layout Bernard. For Payroll checklist.
mClientCode = cmbClient.ItemData(cmbClient.ListIndex)

'If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'    StrQuery = "SELECT * FROM ClientTable WHERE CLIENTCODE2 = 'FMS' AND ClientCode = " & mClientCode
'Else
'    StrQuery = "SELECT * FROM ClientTable WHERE ClientCode = " & mClientCode
'End If
StrQuery = GetClientRecord(mClientCode)
Set rsTemp = ReturnRst(StrQuery, cnnstr, mstatic)

If rsTemp.EOF = False Then
    CutOffPeriod = rsTemp!CutOffPeriod
    SalaryRelease = rsTemp!SalaryRelease
Else
    CutOffPeriod = "NOT SPECIFIED"
    SalaryRelease = "NOT SPECIFIED"
End If
rsTemp.Close

If mskStartCutoff = "__/__/____" Or mskEndCutoff = "__/__/____" Then
    MsgBox "Please enter Start and End Cutoff Date for validation of Deductions Entries date.", vbInformation, PayrollAlert
    GoTo ExitRtn1
ElseIf Val(Format(mskEndCutoff, "yyyymmdd")) < Val(Format(mskStartCutoff, "yyyymmdd")) Then
    MsgBox "Start cutoff should be greater than End Cutoff! Please correct first.", vbInformation, PayrollAlert
    GoTo ExitRtn1
End If

mUserID = Trim(UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID))))
'strQuery = "DELETE FROM tmpPayrollChecklist WHERE useriD like '%" & mUserID & "%'"
StrQuery = "DELETE FROM tmpPayrollChecklist WHERE LTRIM(RTRIM(useriD)) = '" & mUserID & "'"
Call Action_Query(StrQuery, cnnstr)
StrQuery = "INSERT INTO tmpPayrollChecklist (AgencyCode, ClientCode, EmployeeID, AgencyBDOCardFee, SSSLoanAmount, SSSLoanPayment, StartSSSPayment, LastSSSPayment, PbigLoanAmount, PbigLoanPayment, StartPbigPayment, LastPbigPayment, ValeLoanAmount, ValeLoanPayment, StartValePayment, LastValePayment, UpdatedBy, DateUpdated, EmpNo, EmpName, DailyRate, StartContract, EndContract, CBAmount, CBPayment, CBStartPayment, CBEndPayment, MiscFee, MFPayment, MFStartPayment, MFEndPayment, CashAdvance, PFPayment, PFStartPayment, PFENdPayment, LastUpdate, AccountNo, CutOffPeriod, SalaryRelease, Remarks) " & _
        "SELECT AgencyCode, ClientCode, EmployeeID, AgencyBDOCardFee, SSSLoanAmount, SSSLoanPayment, StartSSSPayment, LastSSSPayment, PbigLoanAmount, PbigLoanPayment, StartPbigPayment, LastPbigPayment, ValeLoanAmount, ValeLoanPayment, StartValePayment, LastValePayment, UpdatedBy, DateUpdated, EmpNo, EmpName, DailyRate, StartContract, EndContract, CBAmount, CBPayment, CBStartPayment, CBEndPayment, MiscFee, MFPayment, MFStartPayment, MFEndPayment, CashAdvance, PFPayment, PFStartPayment, PFENdPayment, LastUpdate, AccountNo, CutOffPeriod, SalaryRelease, Remarks " & _
        " From qryDeductionChecklist WHERE ClientCode = " & mClientCode & " AND EmploymentStatus IN " & "(" & EmpStatus & ")"
Call Action_Query(StrQuery, cnnstr)
StrQuery = "UPDATE tmpPayrollChecklist SET UserID = '" & mUserID & "' WHERE userID IS NULL"
Call Action_Query(StrQuery, cnnstr)
StrQuery = "SELECT DISTINCT * FROM tmpPayrollChecklist WHERE ClientCode = " & mClientCode & " AND userID = '" & mUserID & "' ORDER BY EmpName"
Set rsTemp = ReturnRst(StrQuery, cnnstr, mstatic)

'Check if there were records.
If rsTemp.RecordCount = 0 Then
    MsgBox "No records gathered from master database. Cannot proceed!", vbInformation, PayrollAlert
    GoTo ExitRtn1
End If

Call PrintDTRChecklistReport(rsTemp, mClientCode, CutOffPeriod, SalaryRelease)

ExitRtn1:

End Function

Private Sub PrintDTRChecklistReport(ByRef rs As ADODB.Recordset, ByVal mClientCode As Integer, _
    ByVal mCutOffPeriod As String, ByVal mSalaryRelease As String)
    
Dim rsDedRecords As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String, mRange2 As String
Dim sumString As String
Dim ctr As Long
Dim i As Integer
Dim cPos As Integer
Dim LineCtr As Integer
Dim mHead As String
Dim mDeduc As Double
Dim GetHRDedFlag As Boolean
Dim StartCutoffDate As Variant
Dim EndCutoffDate As Variant
Dim StartDate As Variant
Dim LastDate As Variant

'' Get all deduction records for specific client.
'strQuery = "SELECT * FROM qryDeductionRecords WHERE ClientCode = " & mClientCode & " AND year(StartPaydate) > " & (cmbYearof.Text - 2) & " order by clientCode"
'Set rsDedRecords = ReturnRst(strQuery, cnnStr, mstatic)
'If rsDedRecords.EOF = True Then
'    MsgBox "No payroll deduction records selected! Cannot proceed.", vbInformation
'    Set rsDedRecords = Nothing
'    GoTo ExitRtn2
'End If

'Get all Deduction Entries from HR Master database.
'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE ClientCode = " & mClientCode & " AND year(DedStartPayment) >= " & (cmbYearof.Text - 2) & " order by clientCode"
StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE year(DedStartPayment) >= " & (cmbYearof.Text - 2) & " order by clientCode"
Set rsDedRecords = ReturnRst(StrQuery, CnnstrHR, mstatic)
If rsDedRecords.EOF = True Then ' No deduction entries for this client
    GetHRDedFlag = False
Else
    GetHRDedFlag = True
End If

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)

StartCutoffDate = Val(Format(mskStartCutoff, "yyyymmdd"))
EndCutoffDate = Val(Format(mskEndCutoff, "yyyymmdd"))

ctr = 8
i = 1
rs.MoveFirst

'Updated by Bernard May 13, 2008.
mHead = CompanyName & " (" & CompanyShort & ")"
Call CustomizedHeader(xlSheet, xlApp1, "A1", mHead, "", xlLeft, , False, False, False, , "Tahoma", 12, True)
'mHead = "PAYROLL CHECKLIST FOR THE MONTH OF " & UCase(cmbMonthof.Text) & " " & cmbYearof.Text & " EVERY " & mCutOffPeriod
mHead = "FOR THE PAYROLL PERIOD OF: " & Space(20) & cmbMonthof.Text & "   ," & cmbYearof.Text & Space(10) & "(Cutoff Date: " & mskStartCutoff & " - " & mskEndCutoff & ")"
Call CustomizedHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 10, True)
'mHead = "PAYROLL RELEASE DATE EVERY " & mSalaryRelease & " OF " & UCase(cmbMonthof.Text) & " " & cmbYearof.Text
mHead = "PAYROLL RELEASE DATE:      " & Space(20) & cmbMonthof.Text & "   ," & cmbYearof.Text
Call CustomizedHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "COMPANY NAME:              " & Space(20) & cmbClient.Text & Space(10) & "(" & ForWhat & " Employees) "
Call CustomizedHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)

Call CustomizedHeader(xlSheet, xlApp1, "A6:A7", "No.", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 3.14)
Call CustomizedHeader(xlSheet, xlApp1, "B6:B7", "Name", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 22)
Call CustomizedHeader(xlSheet, xlApp1, "C6:C7", "Daily", "Rate", xlCenter, xlBottom, True, True, True, 15, , 9, False, 4.3)
Call CustomizedHeader(xlSheet, xlApp1, "D6:D7", "Staff", "Status", xlCenter, xlBottom, True, True, True, 15, , 9, False, 6)
Call CustomizedHeader(xlSheet, xlApp1, "E6:H6", "Original DTR", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)
Call CustomizedHeader(xlSheet, xlApp1, "E7:E7", "Bandy Clock", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 4.7)
Call CustomizedHeader(xlSheet, xlApp1, "F7:F7", "Chrono Type", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 5)
Call CustomizedHeader(xlSheet, xlApp1, "G7:G7", "Hand Written", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 4.85)
Call CustomizedHeader(xlSheet, xlApp1, "H7:H7", "Xerox", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 4.56)
Call CustomizedHeader(xlSheet, xlApp1, "I6:Q6", "DEDUCTIONS", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)
Call CustomizedHeader(xlSheet, xlApp1, "I7:I7", "C.B. Balance", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.71)
Call CustomizedHeader(xlSheet, xlApp1, "J7:J7", "P/F Balance", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.71)
Call CustomizedHeader(xlSheet, xlApp1, "K7:K7", "M/F Balance", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.71)
Call CustomizedHeader(xlSheet, xlApp1, "L7:L7", "Trans. Fee", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 4.57)
Call CustomizedHeader(xlSheet, xlApp1, "M7:M7", "VALE C/A", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5)
Call CustomizedHeader(xlSheet, xlApp1, "N7:N7", "SLOAN Amount", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
Call CustomizedHeader(xlSheet, xlApp1, "O7:O7", "Variance w/ Client", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 6.57)
Call CustomizedHeader(xlSheet, xlApp1, "P7:P7", "Medical Ded.", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
Call CustomizedHeader(xlSheet, xlApp1, "Q7:Q7", "Uniform Ded.", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
Call CustomizedHeader(xlSheet, xlApp1, "R6:R7", "Pag-ibig", "Loan", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
Call CustomizedHeader(xlSheet, xlApp1, "S6:S7", "SSS", "Loan", xlCenter, xlBottom, True, True, True, 15, , 9, False, 5.57)
'Call CustomizedHeader(xlSheet, xlApp1, "T6:T7", "Daily", "Rate", xlCenter, xlBottom, True, True, True, 15, , 9, False, 10)
Call CustomizedHeader(xlSheet, xlApp1, "T6:T7", "Start Contract", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 7.72)
Call CustomizedHeader(xlSheet, xlApp1, "U6:U7", "End Contract", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 7.72)
Call CustomizedHeader(xlSheet, xlApp1, "V6:V7", "Account Number", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 14.14)
Call CustomizedHeader(xlSheet, xlApp1, "W6:W7", "Remarks", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 20)
rs.MoveFirst

Do While Not rs.EOF
    ' DoEvents
    xlSheet.Cells(ctr, 1).VALUE = i
    xlSheet.Cells(ctr, 2).VALUE = rs!EmpName
    mDeduc = 0#
'    If rs!CbAmount <> 0 Then
'        mDeduc = RunningBalance(rsDedRecords, 1, rs!EmployeeID, rs!cbstartpayment, rs!cbendpayment)
'        xlSheet.Cells(ctr, 8).VALUE = mDeduc
'    End If

    If GetHRDedFlag = True Then ' get HR deduction entries from HR Master
        'Get Vale IOU.
        mDeduc = GetDeductionBalance(rsDedRecords, "vale_iou", rs!EmployeeID)
        xlSheet.Cells(ctr, 13).VALUE = Format(mDeduc, "0.##")  ' Vale/IOU
        'Get Salary Loan.
        mDeduc = GetDeductionBalance(rsDedRecords, "salaryloan", rs!EmployeeID)
        xlSheet.Cells(ctr, 14).VALUE = Format(mDeduc, "0.##")  ' Salary Loan
        'Get Medical deduction.
        mDeduc = GetDeductionBalance(rsDedRecords, "medical", rs!EmployeeID)
        xlSheet.Cells(ctr, 16).VALUE = Format(mDeduc, "0.##")  ' medical
        'Get Uniform Loan.
        mDeduc = GetDeductionBalance(rsDedRecords, "uniform", rs!EmployeeID)
        xlSheet.Cells(ctr, 17).VALUE = Format(mDeduc, "0.##")  ' Uniform
'        'Get Variances or Other deductions.
'        mDeduc = GetDeductionBalance(rsDedRecords, "variances", rs!EmployeeID)
'        xlSheet.Cells(ctr, 17).VALUE = Format(mDeduc, "0.##")  ' variances or other ded
    End If
    xlSheet.Cells(ctr, 3).VALUE = Format(rs!DailyRate, "0.##")
    
    xlSheet.Cells(ctr, 12).VALUE = Format(rs!AgencyBDOCardFee, "0.##")  ' Trans Fee
    
     ' Check if last Vale Payment date elapsed.
     StartDate = Val(Format(rs!StartValePayment, "yyyymmdd"))
     If StartDate = 19000101 Then StartDate = 99999999
     LastDate = Val(Format(rs!LastValePayment, "yyyymmdd"))
     If LastDate = 19000101 Then LastDate = 99999999
    
     If (StartCutoffDate >= StartDate And StartCutoffDate <= LastDate) Or _
         (EndCutoffDate >= StartDate And EndCutoffDate <= LastDate) Then
        xlSheet.Cells(ctr, 15).VALUE = Format(IIf(IsNull(rs!ValeLoanPayment), 0#, rs!ValeLoanPayment), "##,##0.00")   ' Variance with Client
     Else
         xlSheet.Cells(ctr, 15).VALUE = 0#
     End If
    
    'xlSheet.Cells(ctr, 15).VALUE = Format(rs!ValeLoanPayment, "0.##")   ' Variance with Client
    xlSheet.Cells(ctr, 18).VALUE = Format(rs!PbigLoanPayment, "0.##")
    xlSheet.Cells(ctr, 19).VALUE = Format(rs!SSSLoanPayment, "0.##")
    xlSheet.Cells(ctr, 20).VALUE = Format(rs!StartContract, "mm/dd/yyyy")
    xlSheet.Cells(ctr, 21).VALUE = Format(rs!endcontract, "mm/dd/yyyy")
    xlSheet.Cells(ctr, 22).VALUE = rs!AccountNo
    xlSheet.Cells(ctr, 23).VALUE = rs!Remarks
    ctr = ctr + 1
    i = i + 1
    rs.MoveNext
Loop

mRange = "I" & 8 & ":I" & ctr + 1
xlSheet.Range(mRange).Select
mRange = "I" & ctr + 1
xlSheet.Range(mRange).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"

mRange = "J" & 8 & ":J" & ctr + 1
xlSheet.Range(mRange).Select
mRange = "J" & ctr + 1
xlSheet.Range(mRange).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"

mRange = "K" & 8 & ":K" & ctr + 1
xlSheet.Range(mRange).Select
mRange = "K" & ctr + 1
xlSheet.Range(mRange).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"

mRange = "L" & 8 & ":L" & ctr + 1
xlSheet.Range(mRange).Select
mRange = "L" & ctr + 1
xlSheet.Range(mRange).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"

mRange = "M" & 8 & ":M" & ctr + 1
xlSheet.Range(mRange).Select
mRange = "M" & ctr + 1
xlSheet.Range(mRange).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"

mRange = "N" & 8 & ":N" & ctr + 1
xlSheet.Range(mRange).Select
mRange = "N" & ctr + 1
xlSheet.Range(mRange).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"

mRange = "O" & 8 & ":O" & ctr + 1
xlSheet.Range(mRange).Select
mRange = "O" & ctr + 1
xlSheet.Range(mRange).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"

mRange = "P" & 8 & ":P" & ctr + 1
xlSheet.Range(mRange).Select
mRange = "P" & ctr + 1
xlSheet.Range(mRange).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"

mRange = "Q" & 8 & ":Q" & ctr + 1
xlSheet.Range(mRange).Select
mRange = "Q" & ctr + 1
xlSheet.Range(mRange).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"

mRange = "R" & 8 & ":R" & ctr + 1
xlSheet.Range(mRange).Select
mRange = "R" & ctr + 1
xlSheet.Range(mRange).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"

mRange = "S" & 8 & ":S" & ctr + 1
xlSheet.Range(mRange).Select
mRange = "S" & ctr + 1
xlSheet.Range(mRange).Activate
xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"

mRange = "B" & ctr + 1 & ":B" & ctr + 1
Call CustomizedHeader(xlSheet, xlApp1, mRange, "sub-total", "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 22)
mRange = "B" & ctr + 2 & ":B" & ctr + 2
Call CustomizedHeader(xlSheet, xlApp1, mRange, "Total of Staff", "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 22)
mRange = "D" & ctr + 2 & ":D" & ctr + 2
Call CustomizedHeader(xlSheet, xlApp1, mRange, i - 1, "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 5)
mRange = "B" & ctr + 3 & ":B" & ctr + 3
Call CustomizedHeader(xlSheet, xlApp1, mRange, "No. of lacking", "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 22)
ctr = ctr + 3

'Box it.
mRange = "A6:W" & ctr + 3
Call PlaceBox(xlSheet, xlApp1, mRange)

ctr = ctr + 4
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
mRange = "B" & ctr
Call CustomizedHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 8, False, 25)

ctr = ctr + 3
xlSheet.Cells(ctr, 2).VALUE = "1st Batched"
xlSheet.Cells(ctr + 1, 2).VALUE = "No. of DTR Submitted: _______"
xlSheet.Cells(ctr + 2, 2).VALUE = "No. of DTR Lacking: _______"
xlSheet.Cells(ctr + 3, 2).VALUE = "Count of Hold Salary: _______"
xlSheet.Cells(ctr + 4, 2).VALUE = "Count of Cash Salary: _______"

xlSheet.Cells(ctr, 6).VALUE = "2st Batched"
xlSheet.Cells(ctr + 1, 6).VALUE = "No. of DTR Submitted: _______"
xlSheet.Cells(ctr + 2, 6).VALUE = "No. of DTR Lacking: _______"
xlSheet.Cells(ctr + 3, 6).VALUE = "Count of Hold Salary: _______"
xlSheet.Cells(ctr + 4, 6).VALUE = "Count of Cash Salary: _______"

xlSheet.Cells(ctr, 14).VALUE = "3rd Batched"
xlSheet.Cells(ctr + 1, 14).VALUE = "No. of DTR Submitted: _______"
xlSheet.Cells(ctr + 2, 14).VALUE = "No. of DTR Lacking: _______"
xlSheet.Cells(ctr + 3, 14).VALUE = "Count of Hold Salary: _______"
xlSheet.Cells(ctr + 4, 14).VALUE = "Count of Cash Salary: _______"

ctr = ctr + 3 + 4
xlSheet.Cells(ctr, 2).VALUE = "Prepared by: ______________"
xlSheet.Cells(ctr, 7).VALUE = "Note by: ______________"
xlSheet.Cells(ctr, 11).VALUE = "Received by: ______________"
xlSheet.Range("A1").Select

ExitRtn2:
    Set rsDedRecords = Nothing

End Sub

Private Function GetDeductionBalance(ByRef rsDed As ADODB.Recordset, ByVal mKeyField As String, ByVal EmpNo As Integer, _
    Optional ByVal Startpay As Variant, Optional ByVal EndPay As Variant) As Double
Dim mBal As Double
Dim StartDate As Variant
Dim EndDate As Variant

mBal = 0#
With rsDed
    'StartDate = Val(Format(Startpay, "yyyymmdd"))
    'EndDate = Val(Format(EndPay, "yyyymmdd"))
    .Filter = "EmployeeID" & " = " & EmpNo & " AND DeductionFieldName = '" & mKeyField & "'"

'    .Filter = "EmployeeID" & " = " & EmpNo '& " AND (VAL(format(StartPaydate, ""yyyymmdd"")) between " & StartDate & " and " & EndDate & _
        '" or val(format(EndPaydate, ""yyyymmdd"")) between " & StartDate & " and " & EndDate & ")"

'            If StartDate = 19000101 Then StartDate = 99999999
'            If LastDate = 19000101 Then LastDate = 99999999
'            ' Validate whether entry dates fall on cut-off dates.
'            If (StartCutOffDate >= StartDate And StartCutOffDate <= LastDate) Or _
'            (EndCutOffDate >= StartDate And EndCutOffDate <= LastDate) Then
'                IncludeFlag = True
'            Else
'                IncludeFlag = False
'                mRem = "Elpsd."
'            End If

    If .EOF Then
        mBal = 0#
    Else
        mBal = 0#   '!DeductionPayment
'        .MoveFirst
'        Do While Not .EOF
'            Select Case mFieldNo
'                Case 1: ' CB Payments
'                    mBal = mBal + !TASAesla
'                Case 1: ' CB Payments
'                    mBal = mBal + !TASAesla
'                Case 1: ' CB Payments
'                    mBal = mBal + !TASAesla
'                Case 1: ' CB Payments
'                    mBal = mBal + !TASAesla
'                Case 1: ' CB Payments
'                    mBal = mBal + !TASAesla
'            End Select
'            .MoveNext
'        Loop
    End If
    .Filter = ""
End With
GetDeductionBalance = mBal
End Function

Private Sub ShowDtrReport(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim sumString As String
Dim ctr As Long
Dim i As Long, j As Long
Dim cPos As Integer

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)
cPos = 7
ctr = 9
i = 1

' On Error GoTo Errh
' Print Header
xlSheet.Range("A1") = CompanyName

xlSheet.Range("A2") = "FOR THE PAYROLL PERIOD OF:"
xlSheet.Range("A3") = "PAYROLL RELEASE DATE:"
xlSheet.Range("A4") = "COMPANY NAME:"

xlSheet.Range("D2") = rs!CutOffPeriod 'Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd yyyy")
xlSheet.Range("D3") = rs!SalaryRelease  'Format(Date, "mmm dd yyyy")
xlSheet.Range("D4") = rs!ClientName

'With xlSheet.Range("A2")
'    .Font.Name = "Arial"  'Font.Name
'    .Font.Size = 12 'font.size
'    .Font.Bold = True ' Font.Bold
'    .Font.Italic = False 'font.Italic
'End With

xlSheet.Cells(cPos - 1, 1).VALUE = "No."
xlSheet.Cells(cPos - 1, 2).VALUE = "NAME"
xlSheet.Cells(cPos - 1, 3).VALUE = "BRANCH"
xlSheet.Cells(cPos - 1, 4).VALUE = "DTR"
xlSheet.Cells(cPos - 1, 5).VALUE = "OT"
xlSheet.Cells(cPos, 6).VALUE = "C/B"
xlSheet.Cells(cPos, 7).VALUE = "P/F"
xlSheet.Cells(cPos, 8).VALUE = "M/F"
xlSheet.Cells(cPos, 9).VALUE = "C/A"
xlSheet.Cells(cPos - 1, 10).VALUE = "ADD-ON"
xlSheet.Cells(cPos, 10).VALUE = "OTHERS"
xlSheet.Cells(cPos - 1, 11).VALUE = "ACCOUNT"
xlSheet.Cells(cPos, 11).VALUE = "NUMBER"
xlSheet.Cells(cPos, 12).VALUE = "REMARKS"
cPos = cPos + 1

Do While Not rs.EOF
    'DoEvents
    xlSheet.Cells(ctr, 1).VALUE = i
    xlSheet.Cells(ctr, 2).VALUE = rs!EmpName
    xlSheet.Cells(ctr, 3).VALUE = rs!BrName
    xlSheet.Cells(ctr, 6).VALUE = Format(rs!Cbamount, "##,##0.00")
    xlSheet.Cells(ctr, 7).VALUE = Format(rs!CashAdvance, "##,#0.00")
    xlSheet.Cells(ctr, 8).VALUE = Format(rs!MiscFee, "##,##0.00")
    xlSheet.Cells(ctr, 11).VALUE = IIf(rs!AcctNo = "", "n/a", rs!AcctNo)
 
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Cells.EntireRow.AutoFit
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext
Loop

 xlSheet.Cells(ctr + 3, 2).VALUE = "Prepared by :"
 xlSheet.Cells(ctr + 5, 2).VALUE = UCase(GETuSERNAME(operatorID))
 
 xlSheet.Cells(ctr + 3, 5).VALUE = "Noted by :"
 
 xlSheet.Cells(ctr + 3, 10).VALUE = "Received by :"
 
    xlSheet.Columns("A:A").ColumnWidth = 3.43
    xlSheet.Columns("B:B").ColumnWidth = 24.57
    xlSheet.Range("F6:I6").Select
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
     xlApp1.Selection.Merge
    xlApp1.ActiveCell.FormulaR1C1 = "DEDUCTIONS"
    Dim R
    R = "A6:L" & ctr
    xlSheet.Range("" & R & "").Select
     xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
     xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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
        xlSheet.Range("A6:L7").Select
     xlApp1.Selection.Font.Bold = True
    With xlApp1.Selection.Font
        .Name = "Tahoma"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    xlSheet.Columns("L:L").ColumnWidth = 17.86
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
    End With
    xlSheet.Columns("F:I").Select
     xlApp1.Selection.ColumnWidth = 7
    xlSheet.Columns("E:E").ColumnWidth = 4.29
    xlSheet.Columns("D:D").ColumnWidth = 4.29
    
    xlSheet.Columns("K:K").Select
     xlApp1.Selection.ColumnWidth = 14.43
     xlApp1.Selection.ColumnWidth = 14.29
    
      xlApp1.Range("A6:L7").Select
    With xlApp1.Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
    End With
    xlApp1.Range("A1:G4").Select
    xlApp1.Selection.Font.Bold = True
    'xlApp1.Range("F9:H12").Select
    R = "F9:H" & ctr - 1
    xlApp1.Range("" & R & "").Select
    xlApp1.Selection.NumberFormat = "0.0"
    xlApp1.Selection.NumberFormat = "0.00"

    Set xlApp1 = Nothing
    Exit Sub
Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
     Set xlApp1 = Nothing
End Sub

Private Function CreateEmployeesCount()
Dim str As String
Dim rsTemp As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet1 As Excel.Worksheet
Dim ctr As Long
Dim i As Long
              
whereCon = WhereConn

sqlString = "SELECT * FROM vwEmployeeCountPerCompany " & WhereConn
Set rsTemp = ReturnRst(sqlString, cnnstr, mstatic)
If rsTemp.EOF = False Then
    Call ShowEmployeesCount(rsTemp, 4)
Else
    MsgBox "No Records!!!", vbCritical, ""
End If
Set rsTemp = Nothing

End Function

Private Sub OvertimeReport(Optional byEmp As Boolean = False)
Dim rsTemp As ADODB.Recordset
Dim getYear As Integer

'getYear = InputBox("Enter Payroll Year", "Payroll OT & ND Report", Year(Now))
'If Not IsNumeric(getYear) Then
'    MsgBox "Year should be in numeric format!", vbInformation, PayrollAlert
'    Exit Sub
'End If

getYear = Val(cmbYearof.Text)

If byEmp = False Then ' Summary
sqlString = "SELECT max(b.clientname) as ClientName, max(datename(month,a.Startpaydate)) AS PayMonth, Year(a.startpaydate) as PayYear, Max(a.OTHours) AS OTHour, Max(a.OTPay) as OTAmount, " & _
    " MAX(a.Nightdiffot) as NDHours, MAX(a.Nightdiffpay) as NDPay, COUNT(a.employeeid) as EmpCount from " & _
    " PayrollMaster a Inner Join ClientTable b on a.ClientCode = b.ClientCode " & _
    " Where (a.OTHours <> 0 Or a.NightDiffPay <> 0) And Datename(month, a.Startpaydate) = '" & cmbMonthof.Text & "' and Year(a.startpaydate) = " & getYear & _
    " GROUP BY a.ClientCode, MONTH(a.StartPayDate), YEAR(a.StartPaydate) ORDER BY b.ClientName "
Else    ' Per Employee
sqlString = "SELECT max(a.Employeename) as ClientName, max(datename(month,a.Startpaydate)) AS PayMonth, Year(a.startpaydate) as PayYear, Max(a.OTHours) AS OTHour, Max(a.OTPay) as OTAmount, " & _
    " MAX(a.Nightdiffot) as NDHours, MAX(a.Nightdiffpay) as NDPay, COUNT(a.employeeid) as EmpCount from " & _
    " PayrollMaster a Inner Join ClientTable b on a.ClientCode = b.ClientCode " & _
    " Where (a.OTHours <> 0 Or a.NightDiffPay <> 0) And Datename(month, a.Startpaydate) = '" & cmbMonthof.Text & "' and Year(a.startpaydate) = " & getYear & _
    " GROUP BY a.EmployeeName, MONTH(a.StartPayDate), YEAR(a.StartPaydate) ORDER BY a.EmployeeName "
End If
Set rsTemp = ReturnRst(sqlString, cnnstr, mstatic)
If rsTemp.EOF = False Then
    Call ShowEmployeesOT(rsTemp, 4)
Else
    MsgBox "No Records!!!", vbCritical, ""
End If
Set rsTemp = Nothing
End Sub

Private Sub ShowEmployeesOT(ByRef rs As ADODB.Recordset, _
    Optional mDetail As Boolean = False)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim sumString As String
Dim ctr As Long
Dim i As Long, j As Long
Dim cPos As Integer

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)
cPos = 6
ctr = 7
i = 1

'On Error GoTo Errh
'Print Header
xlSheet.Range("A2") = Me.txtTitle.Text & " for " & UCase(cmbMonthof.Text) & " " & cmbYearof.Text
xlSheet.Range("A3") = "COMPANY :"
xlSheet.Range("A4") = "AS OF DATE : "

xlSheet.Range("D3") = CompanyName 'Format(Date, "mmmm dd, yyyy ") 'rs!SalaryRelease  'Format(Date, "mmm dd yyyy")
xlSheet.Range("D4") = "" & Format(Date, "mmmm dd, yyyy") 'rs!ClientName

xlSheet.Cells(cPos, 1).VALUE = "No."
If mDetail = False Then
    xlSheet.Cells(cPos, 2).VALUE = "COMPANY NAME"
Else
    xlSheet.Cells(cPos, 2).VALUE = "EMPLOYEE NAME"
End If
xlSheet.Cells(cPos, 3).VALUE = "PAY MONTH"
xlSheet.Cells(cPos, 4).VALUE = "PAY YEAR"
xlSheet.Cells(cPos, 5).VALUE = "OT HOURS"
xlSheet.Cells(cPos, 6).VALUE = "OT AMOUNT"
xlSheet.Cells(cPos, 7).VALUE = "ND HOURS"
xlSheet.Cells(cPos, 8).VALUE = "ND AMOUNT"
If mDetail = True Then
    xlSheet.Cells(cPos, 9).VALUE = "EMP COUNT"
End If
cPos = cPos + 1

Do While Not rs.EOF
    'DoEvents
    xlSheet.Cells(ctr, 1).VALUE = i
    xlSheet.Cells(ctr, 2).VALUE = rs!ClientName
    xlSheet.Cells(ctr, 3).VALUE = rs!PayMonth
    xlSheet.Cells(ctr, 4).VALUE = rs!PayYear
    xlSheet.Cells(ctr, 5).VALUE = rs!OTHour
    xlSheet.Cells(ctr, 6).VALUE = rs!OTAmount
    xlSheet.Cells(ctr, 7).VALUE = rs!NDHours
    xlSheet.Cells(ctr, 8).VALUE = rs!NDPay
    If mDetail = True Then
        xlSheet.Cells(ctr, 9).VALUE = rs!EmpCount
    End If
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Cells.EntireRow.AutoFit
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext
Loop
   
Dim R
    R = "A6:I" & ctr
    xlSheet.Range("" & R & "").Select
     xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
     xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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
    
    xlSheet.Range("A6:I6").Select
    xlApp1.Selection.Font.Bold = True
    With xlApp1.Selection.Font
        .Name = "Tahoma"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
      
    xlApp1.Range("A6:I6").Select
    With xlApp1.Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
    End With
    
    xlApp1.Columns("A:A").ColumnWidth = 3.57
    xlApp1.Columns("D:D").ColumnWidth = 10.86
    xlApp1.Range("A2:D4").Select
    xlApp1.Selection.Font.Bold = True
'    With xlApp1.Selection
'        .HorizontalAlignment = xlLeft
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .MergeCells = False
'    End With
'    xlApp1.ActiveWindow.ScrollRow = 28
'    xlApp1.ActiveWindow.ScrollRow = 1
'    xlApp1.Range("A6").Select
'    xlApp1.Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(5), _
'        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
        
    Set xlApp1 = Nothing
    Exit Sub

Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
     Set xlApp1 = Nothing
End Sub

Private Sub ShowEmployeesCount(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim sumString As String
Dim ctr As Long
Dim i As Long, j As Long
Dim cPos As Integer

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)
cPos = 6
ctr = 7
i = 1

'On Error GoTo Errh
'Print Header
xlSheet.Range("A2") = Me.txtTitle.Text
xlSheet.Range("A3") = "COMPANY :"
xlSheet.Range("A4") = "AS OF DATE : "

xlSheet.Range("D3") = CompanyName 'Format(Date, "mmmm dd, yyyy ") 'rs!SalaryRelease  'Format(Date, "mmm dd yyyy")
xlSheet.Range("D4") = "" & Format(Date, "mmmm dd, yyyy") 'rs!ClientName


xlSheet.Cells(cPos, 1).VALUE = "No."
xlSheet.Cells(cPos, 2).VALUE = "COMPANY NAME"
xlSheet.Cells(cPos, 3).VALUE = "CODE"
xlSheet.Cells(cPos, 4).VALUE = "EMPLOYMENT STATUS"
xlSheet.Cells(cPos, 5).VALUE = "COUNTS"
cPos = cPos + 1

Do While Not rs.EOF
    'DoEvents
    xlSheet.Cells(ctr, 1).VALUE = i
    xlSheet.Cells(ctr, 2).VALUE = rs!ClientName
    xlSheet.Cells(ctr, 3).VALUE = rs!employmentstatus
    xlSheet.Cells(ctr, 4).VALUE = rs!EmploymentStatusDesc
    xlSheet.Cells(ctr, 5).VALUE = rs!COUNTOFEMPLOYMENTSTATUS
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Cells.EntireRow.AutoFit
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext
Loop
   
    Dim R
    R = "A6:E" & ctr
    xlSheet.Range("" & R & "").Select
     xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
     xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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
        xlSheet.Range("A6:E6").Select
     xlApp1.Selection.Font.Bold = True
    With xlApp1.Selection.Font
        .Name = "Tahoma"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
      
    xlApp1.Range("A6:E6").Select
    With xlApp1.Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
    End With
    
    
    xlApp1.Columns("A:A").ColumnWidth = 3.57
    xlApp1.Range("A2:D4").Select
    xlApp1.Selection.Font.Bold = True
    With xlApp1.Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlApp1.ActiveWindow.ScrollRow = 28
    xlApp1.ActiveWindow.ScrollRow = 1
    xlApp1.Range("A6").Select
    xlApp1.Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(5), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
        
    Set xlApp1 = Nothing
    Exit Sub

Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
     Set xlApp1 = Nothing
End Sub

Public Sub ShowReport_benefits(ByVal valRpt As Integer)
If valRpt = 1 Then
Set rsItems = CreateRecords1("hrrptbenefits")
If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
   With rptBenefits
     Set .DataSource = rsItems
     .DataMember = rsItems.DataMember
     .Sections("Section2").Controls("lblAddress").Caption = getAgencyAddress(CompanyCode) 'Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
       If CompanyCode = 1 Then
            .Sections("Section2").Controls("image1").Visible = True
            .Sections("Section2").Controls("image2").Visible = False
            .Sections("Section2").Controls("image3").Visible = False
       ElseIf CompanyCode = 2 Then
            .Sections("Section2").Controls("image2").Visible = True
            .Sections("Section2").Controls("image1").Visible = False
            .Sections("Section2").Controls("image3").Visible = False
          ElseIf CompanyCode = 3 Then 'to change
          .Sections("Section2").Controls("image3").Visible = True
            .Sections("Section2").Controls("image2").Visible = False
            .Sections("Section2").Controls("image1").Visible = False
        End If

    SELECT1 = Replace(SELECT1, "controlno", "empno")
    .Sections("Section1").Controls("teXT1").DataField = SELECT1
    .Sections("Section1").Controls("Text2").DataField = SELECT2
    .Sections("Section1").Controls("Text3").DataField = SELECT3
    .Sections("Section1").Controls("Text4").DataField = SELECT4
  '  .Sections("Section1").Controls("Text5").DataField = SELECT5
    .Sections("Section2").Controls("lblname").Caption = CboName.Text
    .Sections("Section2").Controls("lblidno").Caption = GetDataName("hrrptbenefits where employeename = '" & CboName.Text & "'", "EmployeeID")
    .Sections("Section2").Controls("lblTitle").Caption = txtTitle
  DoEvents
 .Show
   End With
 
ElseIf valRpt = 2 Then
        Set rsItems = CreateRecords1("hremployeesassignment")
        If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
           With rptAssignment
             Set .DataSource = rsItems
             .DataMember = rsItems.DataMember
             .Sections("Section2").Controls("lblAddress").Caption = getAgencyAddress(CompanyCode) '"Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
               If CompanyCode = 1 Then
                    .Sections("Section2").Controls("image1").Visible = True
                    .Sections("Section2").Controls("image2").Visible = False
               ElseIf CompanyCode = 2 Then
                    .Sections("Section2").Controls("image2").Visible = True
                    .Sections("Section2").Controls("image1").Visible = False
                   ElseIf CompanyCode = 3 Then
                    .Sections("Section2").Controls("image2").Visible = True
                    .Sections("Section2").Controls("image1").Visible = False
                End If
        SELECT1 = Replace(SELECT1, "controlno", "empno")
            .Sections("Section1").Controls("teXT1").DataField = SELECT1
            .Sections("Section1").Controls("Text2").DataField = SELECT2
            .Sections("Section1").Controls("Text3").DataField = SELECT3
            .Sections("Section1").Controls("Text4").DataField = SELECT4
           .Sections("Section1").Controls("Text5").DataField = SELECT5
            .Sections("Section2").Controls("lblname").Caption = CboName.Text
            .Sections("Section2").Controls("lblidno").Caption = GetDataName("hrrptbenefits where employeename = '" & CboName.Text & "'", "EmployeeID")
            .Sections("Section2").Controls("lblTitle").Caption = txtTitle
          DoEvents
         .Show
           End With
 
End If
 
  Exit Sub
End Sub

Public Sub ShowCustomizedReport()
   ' Get a recordset to use with the report
   
    Dim tmpimg As RptImage
    Set tmpimg = rptCustomizedReport.Sections(2).Controls("imgLogo")
    On Error GoTo b:
    'Set tmpimg.Picture = LoadPicture(App.Path & "..\Logos\" & CompanyCode & " logo.gif")
    If LoadPictureFlag = True Then
        Set tmpimg.Picture = LoadPicture(logoPath)
    End If
b:

   
   Set rsItems = CreateRecords
   If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
   With rptCustomizedReport
    ' Bind the Recordset to the Report
     Set .DataSource = rsItems
    .DataMember = rsItems.DataMember
     .Sections("Section2").Controls("lblAddress").Caption = getAgencyAddress(CompanyCode) '"Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
     '.Sections("Section2").Controls("imgLogo").Picture = LoadPicture("D:\myFILES\SYSTEMS\SIDELINE\SEARCHASIA\SEARCH\Logos\1 logo.gif")
     
    SELECT1 = Replace(SELECT1, "controlno", "empno")
    .Sections("Section2").Controls("lblTitle").Caption = txtTitle
    .Sections("Section2").Controls("L1").Caption = SELECT1TITLE
    .Sections("Section2").Controls("L2").Caption = SELECT2TITLE
    .Sections("Section2").Controls("L3").Caption = SELECT3TITLE
    .Sections("Section2").Controls("L4").Caption = SELECT4TITLE
    .Sections("Section2").Controls("L5").Caption = SELECT5TITLE
    .Sections("Section1").Controls("Text1").DataField = SELECT1
    .Sections("Section1").Controls("Text2").DataField = SELECT2
    .Sections("Section1").Controls("Text3").DataField = SELECT3
    .Sections("Section1").Controls("Text4").DataField = SELECT4
    .Sections("Section1").Controls("Text5").DataField = SELECT5
    '.Sections("Section5").Controls("lblRecCount").Caption = CreateRecords.RecordCount & " items listed"
    DoEvents
    .Show
  End With
End Sub

Public Function CreateRecords1(ByVal tblName As String) As ADODB.Recordset
  Dim rsTemp As New ADODB.Recordset, i As Integer
  Dim Sortby, ByDirection As String
  ByDirection = IIf(chkSort(0).VALUE = 1, "Desc", "Asc")
  If cboSort(0).Text <> "(None)" And Trim(cboSort(0).Text) <> "" Then Sortby = " order by " & cboSort(0).Text & " " & ByDirection
  sqlString = "SELECT " & _
              "" & Replace(SELECTED, "controlno", "empno") & "" & _
               " From " & tblName & "  where employeename = '" & CboName.Text & "'" & whereCon & Sortby
    Set CreateRecords1 = ReturnRst(sqlString, cnnstr, mstatic)
End Function

Public Function CreateRecords() As ADODB.Recordset
  Dim rsTemp As New ADODB.Recordset, i As Integer
  Dim Sortby, ByDirection As String
  
  ByDirection = IIf(chkSort(0).VALUE = 1, "Desc", "Asc")
  
  If cboSort(0).Text <> "(None)" And Trim(cboSort(0).Text) <> "" Then Sortby = " order by " & cboSort(0).Text & " " & ByDirection
   
  sqlString = "SELECT " & _
              "" & Replace(SELECTED, "controlno", "empno") & "" & _
               " From vwEmployeeCustomInfo " & whereCon & Sortby
 Set CreateRecords = ReturnRst(sqlString, cnnstr, mstatic)
 'Set CreateRecords = ReturnRst(sqlString, cnnstr, mread)

End Function

Private Sub Form_Unload(Cancel As Integer)
    Set rsItems = Nothing
End Sub

Private Sub lstVw_Click()
Dim X As Integer
    
    frmFilter.Visible = True
    frmHistory.Visible = False
    
    Check2.Enabled = True
    DTRFlag = False
    txtTitle = UCase(lstVw.Text)
    For X = 0 To 3
        Check1(X).Enabled = True
    Next X
    If cboFilterby(0).ListCount <= 1 Then
        For i = 0 To cboFilterby.Count - 1
                  cboFilterby(0).Clear
                  cboFilterby(0).AddItem "Client name"
                  cboFilterby(0).AddItem "Start of contract"
                  cboFilterby(0).AddItem "End of contract"
                  cboFilterby(0).AddItem "Civil status"
        Next
     End If
 
    If lstVw.ListIndex = 7 Then
        Call ChangeChkStatus(False)
        Call ChkStatusAS(0)
         'chkStatus(1).VALUE = 0
        chkStatus(5).VALUE = 1
   ElseIf lstVw.ListIndex = 13 Then ' Head Count Report
      Call ChangeChkStatus(True)
        chkStatus(1).VALUE = 1
       If cboFilterby(0).ListCount > 1 Then
           Check1(1).Enabled = False
           Check1(2).Enabled = False
           Check1(3).Enabled = False
           cboFilterby(0).RemoveItem (2)
           cboFilterby(0).RemoveItem (2)
           cboFilterby(0).RemoveItem (1)
           Check2.Enabled = False
       End If
       ' Do not allow selection per company
        For X = 0 To 3
            Check1(X).Enabled = False
            cboFilterby(X).Enabled = False
        Next X
       Exit Sub
  '  ElseIf lstVw.ListIndex = 16 Or lstVw.ListIndex = 13 Then
    ElseIf lstVw.ListIndex = 16 Or lstVw.ListIndex = 15 Then
            If lstVw.ListIndex = 15 Or lstVw.ListIndex = 16 Then    ' if DTR summary
                Call ChkStatusAS(0, True)
                chkStatus2(1).VALUE = 1
                
                cboHistoryName.Visible = False
                lblSort(0).Visible = False
                Check2.Enabled = False
                HideCheckStatus (True)
                
                'lblSort(2).Visible = True
                lblSort(3).Visible = True
                lblSort(4).Visible = True
                'cmbClient.Visible = True
                cmbMonthof.Visible = True
                cmbYearof.Visible = True
                'mskStartCutoff.Visible = True
                'mskEndCutoff.Visible = True
                'Label1.Visible = True
                'Label2.Visible = True
                DTRFlag = True
            Else                            ' Employment History
                cboHistoryName.Visible = True
                lblSort(0).Visible = True
                Check2.Enabled = True
                HideCheckStatus (False)
                
                lblSort(2).Visible = False
                lblSort(3).Visible = False
                lblSort(4).Visible = False
                cmbClient.Visible = False
                cmbMonthof.Visible = False
                cmbYearof.Visible = False
                mskStartCutoff.Visible = False
                mskEndCutoff.Visible = False
                Label1.Visible = False
                Label2.Visible = False
            End If
            frmFilter.Visible = False
            frmHistory.Visible = True
   Else
        Call ChkStatusAS(0)
        Call ChangeChkStatus(True)
       'chkStatus(1).VALUE = 1  ' Comment by Bernard 11/14/2008
       
   End If
     
    For X = 0 To 3
         If Me.Check1(X).VALUE = 1 Then
             cboFilterby(X).Enabled = True
         Else
             cboFilterby(X).Enabled = False
         End If
    Next X
    
End Sub

Private Sub ChkStatusAS(ByVal VALUE As Integer, Optional checkFlag As Boolean = False)
Dim X As Integer
For X = 0 To 8 '9
    If checkFlag = True Then
        chkStatus(X).VALUE = VALUE
    Else
        chkStatus2(X).VALUE = VALUE
    End If
Next X
End Sub

Private Sub ListBenefitName()
Dim rsemp As New ADODB.Recordset
Set rsemp = cnn.Execute("Select distinct employeename From hrrptbenefits")
CboName.Clear
Do While Not rsemp.EOF
        CboName.AddItem rsemp(0)
        rsemp.MoveNext
    Loop
    rsemp.Close
    Set rsemp = Nothing
End Sub

Private Function GetDataName(ByVal tblWhere As String, ByVal returnFld As String) As String
Dim rsemp As New ADODB.Recordset
Set rsemp = cnn.Execute("Select distinct " & returnFld & " From " & tblWhere)
CboName.Clear
If rsemp.EOF = False Then
        GetDataName = rsemp(0)
        rsemp.MoveNext
  End If
    rsemp.Close
    Set rsemp = Nothing
End Function

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Index = 1 Or Item.Index = 2 Then
        LV.ListItems(Item.Index).Checked = True
        MsgBox "Field cannot be remove from the report", vbInformation
        Exit Sub
    Else
    Dim X As Integer
    If isCheckExceed(Item.Index) = True Then
    For X = 0 To 3
       
         Me.Check1(X).Enabled = True
        cboFilterby(X).Enabled = False
      
    Next X
  Call ChangeChkStatus(True)
  
  'chkStatus(5).Enabled = True
  'chkStatus(7).Enabled = True
    Else
  Call ChangeChkStatus(False)
    
 ' chkStatus(5).Enabled = False
 ' chkStatus(7).Enabled = False
     For X = 0 To 3
    cboFilterby(X).Enabled = False
    Me.Check1(X).Enabled = False
    Me.Check1(X).VALUE = 0
          cboFilterby(X).BackColor = &H8000000A
          cboFilterby(X).ListIndex = -1
          lblCon(X).Visible = False
  
   'If Me.Check1(X).Value = 1 Then cboFilterby(X).Enabled = True
   'If Me.Check1(X).Value = 0 Then cboFilterby(X).Enabled = False
    Next X
    End If
    End If
End Sub

Private Sub mskDate1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then mskDate1(1).SetFocus
End Sub

Private Sub mskDate1_Validate(Index As Integer, Cancel As Boolean)
    If mskDate1(Index) <> "__/__/____" And IsDate(mskDate1(Index)) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskDate1(Index).SetFocus
End Sub

Private Sub mskDate2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then mskDate2(1).SetFocus
End Sub

Private Sub mskDate2_Validate(Index As Integer, Cancel As Boolean)
    If mskDate2(Index) <> "__/__/____" And IsDate(mskDate2(Index)) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskDate2(Index).SetFocus
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)

If PreviousTab = 2 Then
    Dim X As Integer
    
    frmFilter.Visible = True 'False
    Frame1.Visible = False

    If isCheckExceed(1) = True Then
        For X = 0 To 3
            Me.Check1(X).Enabled = False
            cboFilterby(X).Enabled = False
        Next X
  '   chkStatus(5).Enabled = False
 ' chkStatus(7).Enabled = False
    Else
        frmFilter.Visible = True
        Frame1.Visible = True
        Frame3.Visible = False
        lstVw.ListIndex = 0
    End If
Else
    frmFilter.Visible = True
    frmHistory.Visible = False
End If
If SSTab2.Tab = 1 Then
    frmFilter.Visible = True
    frmHistory.Visible = False
ElseIf DTRFlag = True Then
        frmHistory.Visible = True
        frmFilter.Visible = False
Else
        frmHistory.Visible = False
        frmFilter.Visible = True
End If
    
End Sub

Private Sub txtTitle_GotFocus()
    Set oCommonFunction.ValidAlpha = txtTitle
     Call txtHLigths(txtTitle)
End Sub

Public Sub CheckField()
    SELECT1 = ""
    SELECT2 = ""
    SELECT3 = ""
    SELECT4 = ""
    SELECT5 = ""
    SELECTED = ""
    i = 1
    Do While Not i = 38
     If LV.ListItems.Item(i).Checked = True Then
     If SELECTED <> "" Then SELECTED = SELECTED & ","
        SELECTED = SELECTED & LV.ListItems.Item(i).SubItems(1)
       If SELECT1 = "" Then
        SELECT1 = LV.ListItems.Item(i).SubItems(1)
        SELECT1TITLE = LV.ListItems.Item(i).Text
       ElseIf SELECT2 = "" Then
        SELECT2 = LV.ListItems.Item(i).SubItems(1)
        SELECT2TITLE = LV.ListItems.Item(i).Text
       ElseIf SELECT3 = "" Then
        SELECT3 = LV.ListItems.Item(i).SubItems(1)
        SELECT3TITLE = LV.ListItems.Item(i).Text
       ElseIf SELECT4 = "" Then
        SELECT4 = LV.ListItems.Item(i).SubItems(1)
        SELECT4TITLE = LV.ListItems.Item(i).Text
       ElseIf SELECT5 = "" Then
        SELECT5 = LV.ListItems.Item(i).SubItems(1)
        SELECT5TITLE = LV.ListItems.Item(i).Text
       End If
    End If
    i = i + 1
    Loop
    SELECTED = Replace(SELECTED, "controlno", "empno")
End Sub

Public Function isCheckExceed(Optional ByVal ndx As Integer) As Boolean
i = 1
Dim iss As Integer
  'If Lv.ListItems.Item(i).Checked = True Then cboFilterby(0).AddItem Lv.ListItems.Item(ndx).Text
  'If Lv.ListItems.Item(i).Checked = false Then cboFilterby(0) AddItem Lv.ListItems.Item(ndx).Text
  
    Do While Not i = 38
     If LV.ListItems.Item(i).Checked = True Then
        iss = iss + 1
     End If
     i = i + 1
    Loop
   If iss = 5 Then
        isCheckExceed = True
   Else
        isCheckExceed = False
   End If

End Function

Private Sub CreateEmploymentHistory()
Dim str As String
Dim rsTemp As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet1 As Excel.Worksheet
Dim ctr As Long
Dim i As Long

Dim Sortby, ByDirection As String

ByDirection = IIf(chkSort(0).VALUE = 1, "Desc", "Asc")
  
  If cboSort(0).Text <> "(None)" And Trim(cboSort(0).Text) <> "" Then Sortby = " order by " & cboSort(0).Text & " " & ByDirection
   
    If Sortby = "" Then Sortby = " order by Empname"
    SELECTED = " SURNAME + ', ' + FIRSTNAME +' '+  IIf(IsNull(middlename) Or middlename='','',Left(middlename,1)+'.')   AS EMPNAME,EmpNo,clientname,Brname,Deptname ,Empposition,startcontract,endcontract,EMPLOYMENTSTATUS,COMPANYID"
  
  Dim whereCon2 As String
  If cboHistoryName.ListIndex = -1 Then
  whereCon = ""
  whereCon2 = ""
  Else
    whereCon = " where EmployeeID='" & cboHistoryName.ItemData(cboHistoryName.ListIndex) & "'"
    whereCon2 = " where EmployeeID=" & cboHistoryName.ItemData(cboHistoryName.ListIndex) & ""
  End If
 
  
'sqlString = "SELECT " & SELECTED & " From hremployees_history " & whereCon & " ORDER BY SURNAME,FIRSTNAME,startcontract "
 
' sqlString = " SELECT SURNAME+', '+FIRSTNAME+' '+IIf(IsNull(middlename) Or middlename='','',Left(middlename,1)+'.') AS EMPNAME, hremployees_history.EmpNo, ClientTable.ClientName, HRBranch.BrName, HRDept.DeptName, hremployees_history.EmpPosition, hremployees_history.StartContract, hremployees_history.EndContract," & _
'             " hremployees_history.EmploymentStatus, CompanyTable.AgencyName AS COMPANYNAME, CompanyTable.AgencyCode AS COMPANYID" & _
'              " FROM (((hremployees_history INNER JOIN ClientTable ON hremployees_history.ClientCode = ClientTable.ClientCode) INNER JOIN CompanyTable ON hremployees_history.AgencyCode = CompanyTable.AgencyCode) LEFT JOIN HRBranch ON hremployees_history.BrCode = HRBranch.BRCode) LEFT JOIN HRDept ON hremployees_history.DeptCode = HRDept.DeptCode "
' sqlString = sqlString & "" & whereCon & " ORDER BY hremployees_history.SURNAME,hremployees_history.FIRSTNAME,hremployees_history.startcontract "
''
'sqlString = "SELECT CompanyTable.AgencyCode, SURNAME+', '+FIRSTNAME+' '+IIf(IsNull(middlename) Or middlename='','',Left(middlename,1)+'.') AS EMPNAME, hremployees_history.EmpNo,hremployees_history.EmployeeID,hremployees_history.ClientCode, HRBranch.BrName, HRDept.DeptName, hremployees_history.EmpPosition, hremployees_history.StartContract, hremployees_history.EndContract, hremployees_history.EmploymentStatus, CompanyTable.AgencyName AS COMPANYNAME, CompanyTable.AgencyCode AS COMPANYID " & _
'        " FROM ((hremployees_history INNER JOIN CompanyTable ON hremployees_history.AgencyCode = CompanyTable.AgencyCode) LEFT JOIN HRBranch ON hremployees_history.BrCode = HRBranch.BRCode) LEFT JOIN HRDept ON hremployees_history.DeptCode = HRDept.DeptCode "
'sqlString = sqlString & "" & whereCon & " order by startcontract "
'
'sqlString = sqlString & " UNION SELECT CompanyTable.AgencyCode, SURNAME+', '+FIRSTNAME+' '+IIf(IsNull(middlename) Or middlename='','',Left(middlename,1)+'.') AS EMPNAME, hremployees.EmpNo,hremployees.EmployeeID,hremployees.ClientCode, HRBranch.BrName, HRDept.DeptName, hremployees.EmpPosition, hremployees.StartContract, hremployees.EndContract, hremployees.EmploymentStatus, CompanyTable.AgencyName AS COMPANYNAME, CompanyTable.AgencyCode AS COMPANYID " & _
'        " FROM ((hremployees INNER JOIN CompanyTable ON hremployees.AgencyCode = CompanyTable.AgencyCode) LEFT JOIN HRBranch ON hremployees.BrCode = HRBranch.BRCode) LEFT JOIN HRDept ON hremployees.DeptCode = HRDept.DeptCode "
'sqlString = sqlString & "" & whereCon2 & " order by startcontract "

sqlString = "SELECT * FROM vwHREmploymentHistory " & whereCon & " ORDER BY StartContract "


Sortby = ""
               
Set rsTemp = ReturnRst(sqlString, CnnstrHR, mstatic)
If rsTemp.EOF = False Then
    Set xlApp1 = New Excel.Application
    Set xlBook1 = xlApp1.Workbooks.Add()
    xlApp1.Visible = True
    CreateExcelEmploymentHistory rsTemp, xlBook1, xlApp1
    xlApp1.DisplayAlerts = False
 Else
    MsgBox "No Records!!!", vbCritical, "HR Leave OT System"
End If
End Sub

Private Sub CreateMasterList()
Dim str As String
Dim rsTemp As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet1 As Excel.Worksheet
Dim ctr As Long
Dim i As Long

Dim Sortby, ByDirection As String

ByDirection = IIf(chkSort(0).VALUE = 1, "Desc", "Asc")
 
If cboSort(0).Text <> "(None)" And Trim(cboSort(0).Text) <> "" Then Sortby = " order by " & cboSort(0).Text & " " & ByDirection
  
If Sortby = "" Then Sortby = " Order by Empname"
'SELECTED = " Empname,EmpNo,Brname,Deptname, Empposition,Telnum1,Telnum2,startcontract,endcontract,sss,philhealth,TIN,birthdate,TaxCode,Dailyrate,ecolaamount,reasonres,employmentstatusdesc "
SELECTED = " Empname, ClientName, EmpNo, DTRID, Deptname, Empposition, PaymentBasis, DateHired, StartContract, EndContract, SSS, Philhealth, TIN, Pagibig, BirthDate, TaxCode, MonthlyRate, Dailyrate, HourlyRate, EcolaAmount, PFPayment, AcctNo, Supervisor, EmploymentStatusDesc, CommissionRate "

whereCon = Replace(whereCon, "where  clientname", "where  dbo.Clienttable.clientname")
whereCon = Replace(whereCon, "AND EMPLOYMENTSTATUS", "AND dbo.HREmployees.EMPLOYMENTSTATUS")

'' Actual Code.
'sqlString = "SELECT " & _
'             "" & Replace(SELECTED, "controlno", "empno") & "" & _
'              " From vwEmployeeCustomInfo " & whereCon & Sortby

'MonthlyRate, Dailyrate, HourlyRate, EcolaAmount, PFPayment, AcctNo, Supervisor, EmploymentStatusDesc, CommissionRate

sqlString = "SELECT DISTINCT dbo.HREmployees.EmpNo AS ControlNo, CASE WHEN (dbo.HREmployees.middlename IS NULL) OR LTRIM(RTRIM(dbo.HREmployees.middlename)) = '' THEN RTRIM(LTRIM(dbo.HREmployees.Surname)) + ', ' + RTRIM(LTRIM(dbo.HREmployees.FirstName))  ELSE RTRIM(LTRIM(dbo.HREmployees.Surname)) + ', ' + RTRIM(LTRIM(dbo.HREmployees.FirstName))  + ' ' + LEFT(dbo.HREmployees.Middlename, 1) + '.' END  AS EmpName, " & _
        "dbo.ClientTable.ClientName, dbo.HREmployees.DTRID, dbo.HRDept.DeptName, dbo.HRPositions.PositionName AS EmpPosition, dbo.HREmployees.PaymentBasis, dbo.HREmployees.DateHired, dbo.HREmployees.StartContract, " & _
        "dbo.HREmployees.EndContract, dbo.HREmployees.SSS, dbo.HREmployees.Philhealth, dbo.HREmployees.TIN, dbo.HREmployees.Pagibig, dbo.HREmployees.BirthDate, " & _
        "dbo.HREmployees.TaxCode, dbo.HREmployees.MonthlyRate, dbo.HREmployees.DailyRate, dbo.HREmployees.HourlyRate, dbo.EmployeeBasicSalary.EcolaAmount, dbo.HREmployees.PFPayment, dbo.HREmployees.AcctNo, " & _
        "dbo.HREmployees.Supervisor, dbo.HREmploymentStatus.EmploymentStatusDesc, dbo.HREmployees.CommissionRate " & _
    "FROM dbo.HRDetails RIGHT OUTER JOIN " & _
        "dbo.HREmployees LEFT OUTER JOIN " & _
        "dbo.EmployeeBasicSalary ON dbo.HREmployees.EmployeeID = dbo.EmployeeBasicSalary.EmployeeID LEFT OUTER JOIN " & _
        "dbo.HRDept ON dbo.HREmployees.DeptCode = dbo.HRDept.DeptCode LEFT OUTER JOIN " & _
        "dbo.HRPositions ON dbo.HREmployees.PositionCode = dbo.HRPositions.PositionCode LEFT OUTER JOIN " & _
        "dbo.HRBranch ON dbo.HREmployees.BrCode = dbo.HRBranch.BrCode LEFT OUTER JOIN " & _
        "dbo.ClientTable ON dbo.HREmployees.ClientCode = dbo.ClientTable.ClientCode LEFT OUTER JOIN " & _
        "dbo.HREmploymentStatus ON dbo.HREmployees.EmploymentStatus = dbo.HREmploymentStatus.EmploymentStatus LEFT OUTER JOIN " & _
        "dbo.TaxExemptionCode ON dbo.EmployeeBasicSalary.EmpTaxCode = dbo.TaxExemptionCode.TaxExemptionCode ON " & _
        "dbo.HRDetails.EmployeeID = dbo.HREmployees.EmployeeID " & whereCon & Sortby

Sortby = ""
              
Set rsTemp = ReturnRst(sqlString, cnnstr, mstatic)
If rsTemp.EOF = False Then
       Set xlApp1 = New Excel.Application
   Set xlBook1 = xlApp1.Workbooks.Add()
   xlApp1.Visible = True
   CreateExcel rsTemp, xlBook1, xlApp1
   xlApp1.DisplayAlerts = False
Else
   MsgBox "No Records Found!!!", vbCritical, PayrollAlert
End If

End Sub

Function CreateExcelEmploymentHistory(rs As ADODB.Recordset, book1 As Excel.Workbook, ByVal xlApp1 As Excel.Application)
    Dim xlSheet As Excel.Worksheet
    Dim ctr As Long
    Dim i As Long
    Set xlSheet = book1.Worksheets.Add
    ctr = 7
    i = 1
    xlSheet.Cells(1, 1).VALUE = "EMPLOYMENT HISTORY " & IIf(CboCompany.ListIndex >= 0, " ASSIGNED TO " & CboCompany.Text, "")
    xlSheet.Cells(2, 1).VALUE = "As of:"
    xlSheet.Cells(2, 3).VALUE = Format(Date, "mmmm dd, yyyy")
    xlSheet.Cells(3, 1).VALUE = "Agency:"
    xlSheet.Cells(3, 3).VALUE = CompanyName
    xlSheet.Cells(4, 1).VALUE = "Prepared by:"
    xlSheet.Cells(4, 3).VALUE = UCase(GETuSERNAME(operatorID))
    'SELECTED = " Empname,EmpNo,Brname,Deptname, Empposition,Telnum1,Telnum2,startcontract,endcontract,sss,philhealth,TIN,birthdate,TaxCode,Dailyrate,ecolaamount,reasonres,employmentstatusdesc "
'SELECTED = " SURNAME + ', ' + FIRSTNAME + ' ' + MIDDLE AS EMPNAME,EmpNo,CLIENTname,Deptname,BRNAME ,Empposition,startcontract,endcontract,EMPLOYMENTSTATUS AS LASTEMPLOYMENTSTATUS "

    xlSheet.Cells(6, 1).VALUE = "No."
    xlSheet.Cells(6, 2).VALUE = "NAME"
    xlSheet.Cells(6, 3).VALUE = "ID No."
     xlSheet.Cells(6, 4).VALUE = "Client name"
    xlSheet.Cells(6, 5).VALUE = "Branch"
    xlSheet.Cells(6, 6).VALUE = "Department"
    xlSheet.Cells(6, 7).VALUE = "Position"
    xlSheet.Cells(6, 8).VALUE = "Start contract"
    xlSheet.Cells(6, 9).VALUE = "End Contract"
    xlSheet.Cells(6, 10).VALUE = "STATUS"
    xlSheet.Cells(6, 11).VALUE = "COMPANY"
   
    
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Cells.EntireRow.AutoFit
    xlSheet.Cells.Range("A1:k6").Font.Bold = True
    xlSheet.Cells.Range("A1:a5").HorizontalAlignment = xlLeft
    xlSheet.Cells.Range("c1:c5").HorizontalAlignment = xlLeft
    xlSheet.Cells.Range("A6:K6").Interior.ColorIndex = 15
    xlSheet.Columns("A:A").ColumnWidth = 5.71
    Do While Not rs.EOF
       ' DoEvents
        xlSheet.Cells(ctr, 1).VALUE = i
        xlSheet.Cells(ctr, 2).VALUE = rs!EmpName
        xlSheet.Cells(ctr, 3).VALUE = rs!EmpNo
        xlSheet.Cells(ctr, 4).VALUE = rs!ClientName
        'xlSheet.Cells(ctr, 4).VALUE = getClientNamebyID(rs!ClientCode, getTargetCnnstr(rs!AgencyCode)) ' rs!clientName
        xlSheet.Cells(ctr, 5).VALUE = IIf(IsNull(rs!BrName) = True, "", rs!BrName)
        xlSheet.Cells(ctr, 6).VALUE = IIf(IsNull(rs!deptname) = True, "", rs!deptname)
        xlSheet.Cells(ctr, 7).VALUE = IIf(IsNull(rs!empPosition) = True, "", rs!empPosition)
             
         xlSheet.Cells(ctr, 8).VALUE = rs!StartContract
        xlSheet.Cells(ctr, 8).NumberFormat = "mm/dd/yyyy"
        xlSheet.Cells(ctr, 9).VALUE = rs!endcontract
        xlSheet.Cells(ctr, 9).NumberFormat = "mm/dd/yyyy"
    xlSheet.Cells(ctr, 10).VALUE = rs!employmentstatus
    
    xlSheet.Cells(ctr, 11).VALUE = rs!CompanyName ' IIf(rs!companyID = 1, "Jobs and Careers Inc.", "SAPI Manpower & Mgt. Services")
        
        i = i + 1
        ctr = ctr + 1
        rs.MoveNext
    Loop
        xlSheet.Cells.EntireColumn.AutoFit
        xlSheet.Cells.EntireRow.AutoFit
       xlSheet.Columns("A:A").ColumnWidth = 5.71
    
       xlSheet.Columns("C:C").ColumnWidth = 10
    
    'center
    
        xlSheet.Range("A6:K6").Select
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    
    
    'center id
    
      xlSheet.Range("A7:A" & ctr).Select
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
     
       'line
    'Range("A6:Q14").Select
    xlSheet.Range("A6:K" & ctr).Select
    xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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

    xlSheet.Name = "Employment history"
End Function

Function CreateExcel(rs As ADODB.Recordset, book1 As Excel.Workbook, ByVal xlApp1 As Excel.Application)
    Dim xlSheet As Excel.Worksheet
    Dim ctr As Long
    Dim i As Long
    Set xlSheet = book1.Worksheets.Add
    ctr = 7
    i = 1
    xlSheet.Cells(1, 1).VALUE = "EMPLOYEE MASTERLIST " & IIf(CboCompany.ListIndex >= 0, " ASSIGNED TO " & CboCompany.Text, "")
    xlSheet.Cells(2, 1).VALUE = "As of:"
    xlSheet.Cells(2, 3).VALUE = Format(Date, "mmmm dd, yyyy")
    xlSheet.Cells(3, 1).VALUE = "Company:"
    xlSheet.Cells(3, 3).VALUE = CompanyName
    xlSheet.Cells(4, 1).VALUE = "Prepared by:"
    xlSheet.Cells(4, 3).VALUE = UCase(GETuSERNAME(operatorID))
    
    xlSheet.Cells(6, 1).VALUE = "No."
    xlSheet.Cells(6, 2).VALUE = "Name"
    xlSheet.Cells(6, 3).VALUE = "Company Name"
    xlSheet.Cells(6, 4).VALUE = "ID No."
    xlSheet.Cells(6, 5).VALUE = "DTR ID"
    xlSheet.Cells(6, 6).VALUE = "Department"
    xlSheet.Cells(6, 7).VALUE = "Position"
    xlSheet.Cells(6, 8).VALUE = "Payment Base"
    xlSheet.Cells(6, 9).VALUE = "Date Hired"
    xlSheet.Cells(6, 10).VALUE = "Years of Service" ' updated 8/30/2012
    xlSheet.Cells(6, 11).VALUE = "Date From"
    xlSheet.Cells(6, 12).VALUE = "Date To"
    
    xlSheet.Cells(6, 13).VALUE = "SSS Number"
    xlSheet.Cells(6, 14).VALUE = "PhilHealth"
    xlSheet.Cells(6, 15).VALUE = "TIN Number"
    xlSheet.Cells(6, 16).VALUE = "Pagibig"
    xlSheet.Cells(6, 17).VALUE = "Birthday"
    xlSheet.Cells(6, 18).VALUE = "Age"              ' updated 8/30/2012
    xlSheet.Cells(6, 19).VALUE = "CS"
    xlSheet.Cells(6, 20).VALUE = "Monthly Rate"
    xlSheet.Cells(6, 21).VALUE = "Daily Rate"
    xlSheet.Cells(6, 22).VALUE = "Hourly Rate"
    xlSheet.Cells(6, 23).VALUE = "Ecola Rate"
    xlSheet.Cells(6, 24).VALUE = "Allowance"
    xlSheet.Cells(6, 25).VALUE = "Account No"
    xlSheet.Cells(6, 26).VALUE = "Employment Status"
    
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Cells.EntireRow.AutoFit
    xlSheet.Cells.Range("A1:Z6").Font.Bold = True
    xlSheet.Cells.Range("A1:A5").HorizontalAlignment = xlLeft
    xlSheet.Cells.Range("C1:D5").HorizontalAlignment = xlLeft
    xlSheet.Cells.Range("A6:Z6").Interior.ColorIndex = 15
    xlSheet.Columns("A:A").ColumnWidth = 5.71
    
    Do While Not rs.EOF
        DoEvents
        xlSheet.Cells(ctr, 1).VALUE = i
        xlSheet.Cells(ctr, 2).VALUE = rs!EmpName
        xlSheet.Cells(ctr, 3).VALUE = rs!ClientName
        'xlSheet.Cells(ctr, 4).VALUE = rs!EmpNo
        xlSheet.Cells(ctr, 4).VALUE = rs!ControlNo
        xlSheet.Cells(ctr, 5).VALUE = rs!DTRID
        xlSheet.Cells(ctr, 6).VALUE = rs!deptname
        xlSheet.Cells(ctr, 7).VALUE = rs!empPosition
        xlSheet.Cells(ctr, 8).VALUE = rs!PaymentBasis
        xlSheet.Cells(ctr, 9).VALUE = rs!datehired
        xlSheet.Cells(ctr, 9).NumberFormat = "mm/dd/yyyy"
        If IsNull(rs!datehired) = True Then
            xlSheet.Cells(ctr, 10).VALUE = 0
        Else
            xlSheet.Cells(ctr, 10).VALUE = Year(Now) - Year(rs!datehired)
        End If
        xlSheet.Cells(ctr, 11).VALUE = rs!StartContract
        xlSheet.Cells(ctr, 11).NumberFormat = "mm/dd/yyyy"
        xlSheet.Cells(ctr, 12).VALUE = rs!endcontract
        xlSheet.Cells(ctr, 12).NumberFormat = "mm/dd/yyyy"
    
        xlSheet.Cells(ctr, 13).VALUE = IIf(rs!sss = "__-_______-_", "", rs!sss)
        xlSheet.Cells(ctr, 14).VALUE = IIf(rs!PhilHealth = "__-_________-_", "", rs!PhilHealth)
        xlSheet.Cells(ctr, 15).VALUE = IIf(rs!TIN = "___-___-___", "", rs!TIN)
        xlSheet.Cells(ctr, 16).VALUE = rs!Pagibig
        xlSheet.Cells(ctr, 17).VALUE = rs!BIRTHDATE
        xlSheet.Cells(ctr, 17).NumberFormat = "mm/dd/yyyy"
        If IsNull(rs!BIRTHDATE) = True Then
            xlSheet.Cells(ctr, 18).VALUE = 0
        Else
            xlSheet.Cells(ctr, 18).VALUE = Year(Now) - Year(rs!BIRTHDATE)
        End If
        xlSheet.Cells(ctr, 19).VALUE = rs!TaxCode
        
    'If AccountType = "Accounting" Or AccountType = "Administrator" Then
    If ForPayroll(AccountType) = True Then
        If rs!MonthlyRate <> "" Or IsNull(rs!MonthlyRate) = False Then
            If rs!PaymentBasis = "Commission" Then
                xlSheet.Cells(ctr, 20).VALUE = Format(rs!CommissionRate, "##0.00") & "% Com."
            Else
                'xlSheet.Cells(ctr, 20).VALUE = "Php " & Format(rs!MonthlyRate, "#,##0.00")
                xlSheet.Cells(ctr, 20).VALUE = Format(rs!MonthlyRate, "#,##0.00")
            End If
        End If
        If rs!DailyRate <> "" Or IsNull(rs!DailyRate) = False Then
            'xlSheet.Cells(ctr, 21).VALUE = "Php " & Format(rs!DailyRate, "#,##0.00")
            xlSheet.Cells(ctr, 21).VALUE = Format(rs!DailyRate, "#,##0.00")
        End If
        If rs!HourlyRate <> "" Or IsNull(rs!HourlyRate) = False Then
            'xlSheet.Cells(ctr, 22).VALUE = "Php " & Format(rs!HourlyRate, "#,##0.00")
            xlSheet.Cells(ctr, 22).VALUE = Format(rs!HourlyRate, "#,##0.00")
        End If
    End If
        If rs!EcolaAmount <> "" Or IsNull(rs!EcolaAmount) = False Then
            xlSheet.Cells(ctr, 23).VALUE = Format(rs!EcolaAmount, "#,##0.00")
        Else
            xlSheet.Cells(ctr, 23).VALUE = Format(0#, "#,##0.00")
        End If
        
        'xlSheet.Cells(ctr, 21).VALUE = rs!EcolaAmount
        xlSheet.Cells(ctr, 24).VALUE = Format(rs!PFPayment, "#,##0.00")
        xlSheet.Cells(ctr, 25).VALUE = rs!AcctNo
        'xlSheet.Cells(ctr, 25).VALUE = rs!Supervisor ' commented by Bernard 6/12/2018
        xlSheet.Cells(ctr, 26).VALUE = rs!EmploymentStatusDesc
        
        i = i + 1
        ctr = ctr + 1
        rs.MoveNext
    Loop
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Cells.EntireRow.AutoFit
    xlSheet.Columns("A:A").ColumnWidth = 5.71
    xlSheet.Columns("D:D").ColumnWidth = 10
    'Center
    xlSheet.Range("A6:Z6").Select
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    
    'center id
    
      xlSheet.Range("A7:A" & ctr).Select
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
     
       'line
    'Range("A6:Q14").Select
    xlSheet.Range("A6:Z" & ctr).Select
    xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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

    xlSheet.Name = "Masterlist"
   
End Function



