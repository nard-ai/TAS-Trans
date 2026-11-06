VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmLeaveDayoff 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave and Dayoff Application Form"
   ClientHeight    =   6195
   ClientLeft      =   4755
   ClientTop       =   2490
   ClientWidth     =   10965
   Icon            =   "frmLeaveDayoff.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10965
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Leave Remarks"
      Height          =   975
      Left            =   120
      TabIndex        =   152
      Top             =   5160
      Width           =   2895
      Begin VB.CommandButton cmdSaveRem 
         Caption         =   "Save"
         Height          =   315
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   154
         ToolTipText     =   "save remarks"
         Top             =   0
         Width           =   585
      End
      Begin VB.TextBox txtLeaveRem 
         Height          =   525
         Left            =   120
         TabIndex        =   153
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdExportToExcel 
      Appearance      =   0  'Flat
      Caption         =   "Export Schedule to Excel"
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
      ToolTipText     =   "Export Leave/Dayoff Schedule to Excel"
      Top             =   5160
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.CheckBox chkNoFilter 
      BackColor       =   &H8000000E&
      Caption         =   "      All Employees"
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   6360
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Height          =   5535
      Left            =   3120
      TabIndex        =   8
      Top             =   600
      Width           =   7740
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   4335
         Left            =   120
         TabIndex        =   135
         Top             =   1200
         Visible         =   0   'False
         Width           =   7695
         Begin MSComctlLib.ListView ListView1 
            Height          =   2895
            Index           =   1
            Left            =   120
            TabIndex        =   136
            Top             =   240
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   5106
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "EmpNo"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Start OT Date"
               Object.Width           =   1939
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Start Time"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "End Time"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "OT Hours"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Remarks"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Updated By"
               Object.Width           =   2117
            EndProperty
         End
         Begin VB.TextBox txtRemarks 
            Height          =   285
            Left            =   1080
            TabIndex        =   151
            Top             =   3960
            Width           =   4935
         End
         Begin VB.CommandButton cmdDeleteOT 
            Caption         =   "Delete OT"
            Height          =   315
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   150
            Top             =   3480
            Width           =   945
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Start Time ------ End Time ---- Total Time"
            Height          =   615
            Left            =   1800
            TabIndex        =   138
            Top             =   3240
            Width           =   3150
            Begin VB.TextBox txtWorkHours 
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
               Left            =   2280
               TabIndex        =   141
               Text            =   "0.00"
               Top             =   240
               Width           =   630
            End
            Begin MSMask.MaskEdBox mskStartOT 
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "00:00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   139
               Top             =   240
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               ForeColor       =   12582912
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "0"
            End
            Begin MSMask.MaskEdBox mskEndOT 
               Height          =   285
               Left            =   1200
               TabIndex        =   140
               Top             =   240
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               ForeColor       =   12582912
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "0"
            End
         End
         Begin MSComCtl2.DTPicker dtStartOT 
            Height          =   375
            Left            =   240
            TabIndex        =   137
            Top             =   3480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   61145089
            CurrentDate     =   40359
         End
         Begin MSComCtl2.DTPicker dtEndOT 
            Height          =   375
            Left            =   1560
            TabIndex        =   143
            Top             =   3480
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   61145089
            CurrentDate     =   40359
         End
         Begin VB.Label lblclient 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks:"
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
            Index           =   11
            Left            =   120
            TabIndex        =   149
            Top             =   3960
            Width           =   855
         End
         Begin VB.Label lblclient 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
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
            Index           =   9
            Left            =   6480
            TabIndex        =   146
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Overtime:"
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
            Index           =   8
            Left            =   6240
            TabIndex        =   145
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "End OT Date:"
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
            Index           =   7
            Left            =   1680
            TabIndex        =   144
            Top             =   3240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Overtime Date:"
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
            Index           =   5
            Left            =   240
            TabIndex        =   142
            Top             =   3240
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkAllMonths 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Apply to All Months"
         Height          =   255
         Left            =   3480
         TabIndex        =   134
         Top             =   5040
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdSaveEntries 
         Caption         =   "Save Entries"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   720
         Width           =   1185
      End
      Begin VB.CommandButton cmdEditEntries 
         Caption         =   "Edit Entries"
         Height          =   315
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   720
         Width           =   1185
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   41
         ItemData        =   "frmLeaveDayoff.frx":058A
         Left            =   6600
         List            =   "frmLeaveDayoff.frx":058C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   4800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   40
         ItemData        =   "frmLeaveDayoff.frx":058E
         Left            =   5520
         List            =   "frmLeaveDayoff.frx":0590
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   4800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   39
         ItemData        =   "frmLeaveDayoff.frx":0592
         Left            =   4440
         List            =   "frmLeaveDayoff.frx":0594
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   4800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   38
         ItemData        =   "frmLeaveDayoff.frx":0596
         Left            =   3360
         List            =   "frmLeaveDayoff.frx":0598
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   4800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   37
         ItemData        =   "frmLeaveDayoff.frx":059A
         Left            =   2280
         List            =   "frmLeaveDayoff.frx":059C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   4800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   36
         ItemData        =   "frmLeaveDayoff.frx":059E
         Left            =   1200
         List            =   "frmLeaveDayoff.frx":05A0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   4800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   35
         ItemData        =   "frmLeaveDayoff.frx":05A2
         Left            =   120
         List            =   "frmLeaveDayoff.frx":05A4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   4800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   6
         Left            =   6600
         TabIndex        =   78
         Text            =   "Saturday"
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   5
         Left            =   5520
         TabIndex        =   77
         Text            =   "Friday"
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   4
         Left            =   4440
         TabIndex        =   76
         Text            =   "Thursday"
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   75
         Text            =   "Wednesday"
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   74
         Text            =   "Tuesday"
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   73
         Text            =   "Monday"
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   72
         Text            =   "Sunday"
         Top             =   1200
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   34
         ItemData        =   "frmLeaveDayoff.frx":05A6
         Left            =   6600
         List            =   "frmLeaveDayoff.frx":05A8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   33
         ItemData        =   "frmLeaveDayoff.frx":05AA
         Left            =   5520
         List            =   "frmLeaveDayoff.frx":05AC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   32
         ItemData        =   "frmLeaveDayoff.frx":05AE
         Left            =   4440
         List            =   "frmLeaveDayoff.frx":05B0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   31
         ItemData        =   "frmLeaveDayoff.frx":05B2
         Left            =   3360
         List            =   "frmLeaveDayoff.frx":05B4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   30
         ItemData        =   "frmLeaveDayoff.frx":05B6
         Left            =   2280
         List            =   "frmLeaveDayoff.frx":05B8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   29
         ItemData        =   "frmLeaveDayoff.frx":05BA
         Left            =   1200
         List            =   "frmLeaveDayoff.frx":05BC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   28
         ItemData        =   "frmLeaveDayoff.frx":05BE
         Left            =   120
         List            =   "frmLeaveDayoff.frx":05C0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   27
         ItemData        =   "frmLeaveDayoff.frx":05C2
         Left            =   6600
         List            =   "frmLeaveDayoff.frx":05C4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   26
         ItemData        =   "frmLeaveDayoff.frx":05C6
         Left            =   5520
         List            =   "frmLeaveDayoff.frx":05C8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   25
         ItemData        =   "frmLeaveDayoff.frx":05CA
         Left            =   4440
         List            =   "frmLeaveDayoff.frx":05CC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   24
         ItemData        =   "frmLeaveDayoff.frx":05CE
         Left            =   3360
         List            =   "frmLeaveDayoff.frx":05D0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   23
         ItemData        =   "frmLeaveDayoff.frx":05D2
         Left            =   2280
         List            =   "frmLeaveDayoff.frx":05D4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   22
         ItemData        =   "frmLeaveDayoff.frx":05D6
         Left            =   1200
         List            =   "frmLeaveDayoff.frx":05D8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   21
         ItemData        =   "frmLeaveDayoff.frx":05DA
         Left            =   120
         List            =   "frmLeaveDayoff.frx":05DC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   20
         ItemData        =   "frmLeaveDayoff.frx":05DE
         Left            =   6600
         List            =   "frmLeaveDayoff.frx":05E0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   19
         ItemData        =   "frmLeaveDayoff.frx":05E2
         Left            =   5520
         List            =   "frmLeaveDayoff.frx":05E4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   18
         ItemData        =   "frmLeaveDayoff.frx":05E6
         Left            =   4440
         List            =   "frmLeaveDayoff.frx":05E8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   17
         ItemData        =   "frmLeaveDayoff.frx":05EA
         Left            =   3360
         List            =   "frmLeaveDayoff.frx":05EC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   16
         ItemData        =   "frmLeaveDayoff.frx":05EE
         Left            =   2280
         List            =   "frmLeaveDayoff.frx":05F0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   15
         ItemData        =   "frmLeaveDayoff.frx":05F2
         Left            =   1200
         List            =   "frmLeaveDayoff.frx":05F4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   14
         ItemData        =   "frmLeaveDayoff.frx":05F6
         Left            =   120
         List            =   "frmLeaveDayoff.frx":05F8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   13
         ItemData        =   "frmLeaveDayoff.frx":05FA
         Left            =   6600
         List            =   "frmLeaveDayoff.frx":05FC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   12
         ItemData        =   "frmLeaveDayoff.frx":05FE
         Left            =   5520
         List            =   "frmLeaveDayoff.frx":0600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   11
         ItemData        =   "frmLeaveDayoff.frx":0602
         Left            =   4440
         List            =   "frmLeaveDayoff.frx":0604
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   10
         ItemData        =   "frmLeaveDayoff.frx":0606
         Left            =   3360
         List            =   "frmLeaveDayoff.frx":0608
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   9
         ItemData        =   "frmLeaveDayoff.frx":060A
         Left            =   2280
         List            =   "frmLeaveDayoff.frx":060C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   8
         ItemData        =   "frmLeaveDayoff.frx":060E
         Left            =   1200
         List            =   "frmLeaveDayoff.frx":0610
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   7
         ItemData        =   "frmLeaveDayoff.frx":0612
         Left            =   120
         List            =   "frmLeaveDayoff.frx":0614
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   6
         ItemData        =   "frmLeaveDayoff.frx":0616
         Left            =   6600
         List            =   "frmLeaveDayoff.frx":0618
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   5
         ItemData        =   "frmLeaveDayoff.frx":061A
         Left            =   5520
         List            =   "frmLeaveDayoff.frx":061C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   4
         ItemData        =   "frmLeaveDayoff.frx":061E
         Left            =   4440
         List            =   "frmLeaveDayoff.frx":0620
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   3
         ItemData        =   "frmLeaveDayoff.frx":0622
         Left            =   3360
         List            =   "frmLeaveDayoff.frx":0624
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   2
         ItemData        =   "frmLeaveDayoff.frx":0626
         Left            =   2280
         List            =   "frmLeaveDayoff.frx":0628
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   1
         ItemData        =   "frmLeaveDayoff.frx":062A
         Left            =   1200
         List            =   "frmLeaveDayoff.frx":062C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbDays 
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
         Index           =   0
         ItemData        =   "frmLeaveDayoff.frx":062E
         Left            =   120
         List            =   "frmLeaveDayoff.frx":0630
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1050
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
         Text            =   "frmLeaveDayoff.frx":0632
         Top             =   345
         Width           =   1875
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
         Text            =   "frmLeaveDayoff.frx":063E
         Top             =   345
         Width           =   1275
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Display Leaves"
         Height          =   315
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Width           =   1305
      End
      Begin MSComctlLib.ListView lstDTRDisplay 
         Height          =   330
         Left            =   360
         TabIndex        =   9
         Top             =   6000
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
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
         Left            =   1200
         TabIndex        =   17
         Top             =   5760
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
         Left            =   1320
         TabIndex        =   26
         Top             =   6000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61145089
         CurrentDate     =   40262
      End
      Begin MSComCtl2.DTPicker DTREnd 
         Height          =   375
         Left            =   3480
         TabIndex        =   27
         Top             =   6000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61145089
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
         ItemData        =   "frmLeaveDayoff.frx":064A
         Left            =   2760
         List            =   "frmLeaveDayoff.frx":064C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
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
         ItemData        =   "frmLeaveDayoff.frx":064E
         Left            =   960
         List            =   "frmLeaveDayoff.frx":0650
         Style           =   2  'Dropdown List
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   720
         Width           =   1245
      End
      Begin VB.CommandButton cmdPrintLeave 
         Caption         =   "Print"
         Height          =   195
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   156
         ToolTipText     =   "Print Leaves"
         Top             =   5280
         Width           =   585
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Used VL:"
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
         Index           =   12
         Left            =   120
         TabIndex        =   155
         Top             =   5280
         Width           =   2775
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
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
         Index           =   4
         Left            =   5940
         TabIndex        =   133
         Top             =   5280
         Width           =   1905
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Updated By:"
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
         Index           =   3
         Left            =   3600
         TabIndex        =   132
         Top             =   5280
         Width           =   2295
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   0
         X2              =   7680
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   41
         Left            =   6720
         TabIndex        =   129
         Top             =   4560
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   40
         Left            =   5640
         TabIndex        =   128
         Top             =   4560
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   4560
         TabIndex        =   127
         Top             =   4560
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   38
         Left            =   3480
         TabIndex        =   126
         Top             =   4560
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   2400
         TabIndex        =   125
         Top             =   4560
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   1320
         TabIndex        =   124
         Top             =   4560
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   240
         TabIndex        =   123
         Top             =   4560
         Width           =   840
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
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
         Left            =   2280
         TabIndex        =   115
         Top             =   780
         Width           =   360
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   6720
         TabIndex        =   114
         Top             =   3960
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   5640
         TabIndex        =   113
         Top             =   3960
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   4560
         TabIndex        =   112
         Top             =   3960
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   3480
         TabIndex        =   111
         Top             =   3960
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   2400
         TabIndex        =   110
         Top             =   3960
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   1320
         TabIndex        =   109
         Top             =   3960
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   108
         Top             =   3960
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   6720
         TabIndex        =   107
         Top             =   3360
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   5640
         TabIndex        =   106
         Top             =   3360
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   4560
         TabIndex        =   105
         Top             =   3360
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   3480
         TabIndex        =   104
         Top             =   3360
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   2400
         TabIndex        =   103
         Top             =   3360
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   1320
         TabIndex        =   102
         Top             =   3360
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   101
         Top             =   3360
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   6720
         TabIndex        =   100
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   5640
         TabIndex        =   99
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   4560
         TabIndex        =   98
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   3480
         TabIndex        =   97
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   2400
         TabIndex        =   96
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   1320
         TabIndex        =   95
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   94
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   6720
         TabIndex        =   93
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   5640
         TabIndex        =   92
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   4560
         TabIndex        =   91
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3480
         TabIndex        =   90
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   2400
         TabIndex        =   89
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   88
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   87
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   6720
         TabIndex        =   86
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5640
         TabIndex        =   85
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4560
         TabIndex        =   84
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   83
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   82
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   81
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   80
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Month of:"
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
         Left            =   120
         TabIndex        =   79
         Top             =   780
         Width           =   840
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
         Caption         =   "End Date:"
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
         Left            =   2760
         TabIndex        =   28
         Top             =   6075
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
         Left            =   4080
         TabIndex        =   25
         Top             =   6240
         Width           =   960
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date:"
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
         Left            =   480
         TabIndex        =   15
         Top             =   6075
         Width           =   915
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
         Left            =   4200
         TabIndex        =   10
         Top             =   5880
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
      ItemData        =   "frmLeaveDayoff.frx":0652
      Left            =   120
      List            =   "frmLeaveDayoff.frx":065F
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
      Width           =   10965
      _ExtentX        =   19341
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
         Begin VB.OptionButton optLeave 
            Caption         =   "Overtime"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   148
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton optLeave 
            Caption         =   "Leave/Dayoff"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   147
            Top             =   120
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Employee Leave and Dayoff Entries"
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
            Width           =   4575
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
      Top             =   1440
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
Attribute VB_Name = "frmLeaveDayoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsClientData As ADODB.Recordset
Dim rsEmployeeBasicData As ADODB.Recordset
Dim rsDTRRecords As ADODB.Recordset
Dim rsHREmployees As ADODB.Recordset
Dim rsCalendar As ADODB.Recordset
Dim rsLeaveType As ADODB.Recordset
Dim rsFiledOT As ADODB.Recordset
Dim rsLeaveRemarks As ADODB.Recordset
Dim rsLeaveYear As ADODB.Recordset

Dim InvalidIDCtr As Integer
Dim ActiveCtr As Integer
Dim TotalCount As Integer
Dim InvalidIDSw As Boolean
Dim DuplicateDTR As Integer
Dim mDTRRec As Integer

Dim mDTRID As String
Dim firstRun As Boolean
Dim CalStart As Integer
Dim CalEnd As Integer
Dim AddOTFlag As Boolean
Dim LeaveRem(0 To 31) As String
Dim CurIndexNo As Integer
Dim DayIndex As Integer

Private Sub cmbCompany_Click()
AutoSizeDropDownWidth cmbCompany
    Call FilterEmployee(cmbCompany.Text)
    If ActiveCtr = 0 Then
        Call DisplayBasicData("", False)
    Else
        If ListView1(0).ListItems.Count > 0 Then
            Call DisplayBasicData(ListView1(0).SelectedItem.Text, False)
        End If
    End If
End Sub

Private Sub EditFiledOvertime()
If cmdEditEntries.Caption = "Edit OT" Then
    cmdEditEntries.Caption = "Cancel"
    cmdSaveEntries.Caption = "Save"
    cmdDeleteOT.Enabled = False
    EditOTFlag = False
    dtStartOT.Enabled = False
    dtEndOT.Enabled = False
    Frame6.Enabled = True
    txtRemarks.Enabled = True
Else    ' Cancel edits
    cmdEditEntries.Caption = "Edit OT"
    cmdSaveEntries.Caption = "Add OT"
    dtStartOT.Enabled = False
    dtEndOT.Enabled = False
    Frame6.Enabled = False
    txtRemarks.Enabled = False
    cmdDeleteOT.Enabled = True
End If
    
End Sub

Private Sub cmbDays_Click(Index As Integer)
AutoSizeDropDownWidth cmbDays(Index)
End Sub

Private Sub cmbDays_GotFocus(Index As Integer)
Frame5.Caption = Mid(cmbMonth, 4, 3) & " " & lblNo(Index).Caption & ", " & cmbYear & " Remarks"
CurIndexNo = Val(lblNo(Index).Caption)
DayIndex = Index

If cmbDays(Index).Text <> "" Then
    txtLeaveRem.Text = GetLeaveRemarks(rsLeaveRemarks)
Else
    txtLeaveRem.Text = ""
End If
End Sub

Private Sub cmdDeleteOT_Click()
If MsgBox("Delete selected overtime?", vbYesNo + vbQuestion, "Overtime Entry") = vbYes Then
    StrQuery = "DELETE HRFiledOvertime WHERE DTRID = '" & mDTRID & "' AND Year(StartOTDate) = " & cmbYear.Text & " and Month(StartOTDate) = " & Val(Left(cmbMonth.Text, 2)) & " AND day(startotdate) = " & Format(dtStartOT.VALUE, "dd")
    Action_Query StrQuery, cnnstr
Else
    Exit Sub
End If
StrQuery = "SELECT * FROM HRFiledOvertime WHERE DTRID = '" & mDTRID & "' AND Year(StartOTDate) = " & cmbYear.Text & " and Month(StartOTDate) = " & Val(Left(cmbMonth.Text, 2)) & " ORDER BY EmpNo, StartOTDate"
Set rsFiledOT = ReturnRst(StrQuery, cnnstr, mstatic)
Call DisplayFiledOT(rsFiledOT)
End Sub

Private Sub cmdEditEntries_Click()
Dim i As Integer

If optLeave(0).VALUE = False Then ' edit OT
    AddOTFlag = False
    Call EditFiledOvertime
    Exit Sub
End If

If cmdEditEntries.Caption = "Cancel" Then
    If MsgBox("Cancel changes to Leaves and Dayoff entries?", vbYesNo + vbQuestion, "DTR Leave and Dayoff Entry") = vbYes Then
        cmdEditEntries.Caption = "Edit Entries"
        cmdSaveEntries.Enabled = False
        cmbCompany.Enabled = True
        ListView1(0).Enabled = True
        For i = 0 To 41
            cmbDays(i).Enabled = False
        Next i
        Exit Sub
    End If
End If
cmdEditEntries.Caption = "Cancel"
cmdSaveEntries.Enabled = True
cmbCompany.Enabled = False
ListView1(0).Enabled = False
For i = 0 To 41
    cmbDays(i).Enabled = True
Next i
End Sub

Private Sub cmdExportToExcel_Click()
If rsCalendar.RecordCount <> 0 Then
    Call PrintScheduleToExcel
Else
    MsgBox "No records found!", vbInformation, PayrollAlert
End If
End Sub

Private Sub PrintScheduleToExcel()
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
Dim StartDeduction As Variant
Dim EndDeduction As Variant
Dim DeductionYear As Integer

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
mHead = "Schedule of Leaves and Dayoff for All Employees"
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy hh:mm:ss")
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

Private Sub cmdPrintLeave_Click()
Call PrintLeaveBalance
End Sub

Private Sub cmdRefresh_Click()

If optLeave(0).VALUE = False Then ' for OT
    rsFiledOT.Close
    StrQuery = "SELECT * FROM HRFiledOvertime WHERE DTRID = '" & mDTRID & "' AND Year(StartOTDate) = " & cmbYear.Text & " and Month(StartOTDate) = " & Val(Left(cmbMonth.Text, 2)) & " ORDER BY EmpNo, StartOTDate"
    Set rsFiledOT = ReturnRst(StrQuery, cnnstr, mstatic)

    'Call DisplayBasicData(ListView1(1).SelectedItem.Text, False)
    Call DisplayFiledOT(rsFiledOT)
    Exit Sub
End If

Call DisplayCalendar

rsCalendar.Close
StrQuery = "SELECT * FROM HRDayoffCalendar WHERE calyearmonth = " & LTrim(RTrim(cmbYear.Text)) & Left(cmbMonth.Text, 2)
Set rsCalendar = ReturnRst(StrQuery, cnnstr, mstatic)

rsLeaveYear.Close
StrQuery = "SELECT a.*, (LTRIM(LTRIM(b.Surname)) + ', ' + LTRIM(RTRIM(b.Firstname)) + ' ' + LEFT(b.Middlename,1)) AS Fullname " & _
    "FROM HRDayoffCalendar a LEFT JOIN HREmployees b ON a.EmpNo = b.EmpNo WHERE b.ClientCode = " & Val(txtClientCode) & " and a.CalYear = " & LTrim(RTrim(cmbYear.Text)) & " ORDER BY b.Surname, b.Firstname, a.CalyearMonth, a.Day_ID"
Set rsLeaveYear = ReturnRst(StrQuery, cnnstr, mstatic)

rsLeaveRemarks.Close
StrQuery = "SELECT * FROM HRDayoffRemarks WHERE calyearmonth = " & LTrim(RTrim(cmbYear.Text)) & Left(cmbMonth.Text, 2)
Set rsLeaveRemarks = ReturnRst(StrQuery, cnnstr, mstatic)

Call DisplayBasicData(ListView1(0).SelectedItem.Text, False)
If firstRun = True Then
    MsgBox "Leave and Dayoff Schedule was updated!", vbInformation, PayrollAlert
End If

End Sub

Private Sub DisplayCalendar()
Dim i As Integer
Dim j As Integer
Dim mDate As Date
Dim CalStartFlag As Boolean

On Error GoTo ExitRtn1

j = 0
CalStart = 0
CalEnd = 0
CalStartFlag = False

For i = 0 To 41
    cmbDays(i).Visible = False
    lblNo(i).Visible = False
Next i

If cmbYear.Text = "" Then
    Exit Sub
End If

'mDate = Format(Date, Left(cmbMonth.Text, 2) & "/" & Right("00" & i, 2) & "/" & cmbYear.Text)

For i = 0 To 41
    j = i + 1
    mDate = CDate(Left(cmbMonth.Text, 2) & "/" & Right("00" & j, 2) & "/" & cmbYear.Text)
    
    Select Case Format(mDate, "dddd")
       ' for Sundays
        Case "Sunday"
            'If (i = 0 Or i = 7 Or i = 14 Or i = 21 Or i = 28 or i = 35) Then
            If i <= 0 Then
                If CalStart = 0 And CalStartFlag = False Then
                    CalStart = 0: CalStartFlag = True
                End If
                lblNo(0).Caption = j
                cmbDays(0).Visible = True
                lblNo(0).Visible = True
            ElseIf i <= 7 Then
                If CalStart = 0 And CalStartFlag = False Then
                    CalStart = 7: CalStartFlag = True
                End If
                lblNo(7).Caption = j
                cmbDays(7).Visible = True
                lblNo(7).Visible = True
            ElseIf i <= 14 Then
                lblNo(14).Caption = j
                cmbDays(14).Visible = True
                lblNo(14).Visible = True
            ElseIf i <= 21 Then
                lblNo(21).Caption = j
                cmbDays(21).Visible = True
                lblNo(21).Visible = True
            ElseIf i <= 28 Then
                CalEnd = j
                lblNo(28).Caption = j
                cmbDays(28).Visible = True
                lblNo(28).Visible = True
            Else
                CalEnd = j
                lblNo(35).Caption = j
                cmbDays(35).Visible = True
                lblNo(35).Visible = True
            End If
            'End If
        Case "Monday"
            'If (i = 1 Or i = 8 Or i = 15 Or i = 22 Or i = 29) Then
            If i <= 1 Then
                If CalStart = 0 And CalStartFlag = False Then
                    CalStart = 1: CalStartFlag = True
                End If
                lblNo(1).Caption = j
                cmbDays(1).Visible = True
                lblNo(1).Visible = True
            ElseIf i <= 8 Then
                lblNo(8).Caption = j
                cmbDays(8).Visible = True
                lblNo(8).Visible = True
            ElseIf i <= 15 Then
                lblNo(15).Caption = j
                cmbDays(15).Visible = True
                lblNo(15).Visible = True
            ElseIf i <= 22 Then
                lblNo(22).Caption = j
                cmbDays(22).Visible = True
                lblNo(22).Visible = True
            ElseIf i <= 29 Then
                CalEnd = j
                lblNo(29).Caption = j
                cmbDays(29).Visible = True
                lblNo(29).Visible = True
            Else
                CalEnd = j
                lblNo(36).Caption = j
                cmbDays(36).Visible = True
                lblNo(36).Visible = True
            End If
            'End If
        Case "Tuesday"
            'If (i = 2 Or i = 9 Or i = 16 Or i = 23 Or i = 30) Then
            If i <= 2 Then
                If CalStart = 0 And CalStartFlag = False Then
                    CalStart = 2: CalStartFlag = True
                End If
                lblNo(2).Caption = j
                cmbDays(2).Visible = True
                lblNo(2).Visible = True
            ElseIf i <= 9 Then
                lblNo(9).Caption = j
                cmbDays(9).Visible = True
                lblNo(9).Visible = True
            ElseIf i <= 16 Then
                lblNo(16).Caption = j
                cmbDays(16).Visible = True
                lblNo(16).Visible = True
            ElseIf i <= 23 Then
                lblNo(23).Caption = j
                cmbDays(23).Visible = True
                lblNo(23).Visible = True
            ElseIf i <= 30 Then
                CalEnd = j
                lblNo(30).Caption = j
                cmbDays(30).Visible = True
                lblNo(30).Visible = True
            Else
                CalEnd = j
                lblNo(37).Caption = j
                cmbDays(37).Visible = True
                lblNo(37).Visible = True
            End If
            'End If
        Case "Wednesday"
            'If (i = 3 Or i = 10 Or i = 17 Or i = 24 Or i = 31) Then
            If i <= 3 Then
                If CalStart = 0 And CalStartFlag = False Then
                    CalStart = 3: CalStartFlag = True
                End If
                lblNo(3).Caption = j
                cmbDays(3).Visible = True
                lblNo(3).Visible = True
            ElseIf i <= 10 Then
                lblNo(10).Caption = j
                cmbDays(10).Visible = True
                lblNo(10).Visible = True
            ElseIf i <= 17 Then
                lblNo(17).Caption = j
                cmbDays(17).Visible = True
                lblNo(17).Visible = True
            ElseIf i <= 24 Then
                lblNo(24).Caption = j
                cmbDays(24).Visible = True
                lblNo(24).Visible = True
            ElseIf i <= 31 Then
                CalEnd = j
                lblNo(31).Caption = j
                cmbDays(31).Visible = True
                lblNo(31).Visible = True
            Else
                CalEnd = j
                lblNo(38).Caption = j
                cmbDays(38).Visible = True
                lblNo(38).Visible = True
            End If
            'End If
        Case "Thursday"
            'If (i = 4 Or i = 11 Or i = 18 Or i = 25 Or i = 32) Then
            If i <= 4 Then
                If CalStart = 0 And CalStartFlag = False Then
                    CalStart = 4: CalStartFlag = True
                End If
                lblNo(4).Caption = j
                cmbDays(4).Visible = True
                lblNo(4).Visible = True
            ElseIf i <= 11 Then
                lblNo(11).Caption = j
                cmbDays(11).Visible = True
                lblNo(11).Visible = True
            ElseIf i <= 18 Then
                lblNo(18).Caption = j
                cmbDays(18).Visible = True
                lblNo(18).Visible = True
            ElseIf i <= 25 Then
                lblNo(25).Caption = j
                cmbDays(25).Visible = True
                lblNo(25).Visible = True
            ElseIf i <= 32 Then
                CalEnd = j
                lblNo(32).Caption = j
                cmbDays(32).Visible = True
                lblNo(32).Visible = True
            Else
                CalEnd = j
                lblNo(39).Caption = j
                cmbDays(39).Visible = True
                lblNo(39).Visible = True
            End If
        Case "Friday"
            'If (i = 5 Or i = 12 Or i = 19 Or i = 26 Or i = 33) Then
            If i <= 5 Then
                If CalStart = 0 And CalStartFlag = False Then
                    CalStart = 5: CalStartFlag = True
                End If
                lblNo(5).Caption = j
                cmbDays(5).Visible = True
                lblNo(5).Visible = True
            ElseIf i <= 12 Then
                lblNo(12).Caption = j
                cmbDays(12).Visible = True
                lblNo(12).Visible = True
            ElseIf i <= 19 Then
                lblNo(19).Caption = j
                cmbDays(19).Visible = True
                lblNo(19).Visible = True
            ElseIf i <= 26 Then
                lblNo(26).Caption = j
                cmbDays(26).Visible = True
                lblNo(26).Visible = True
            ElseIf i <= 33 Then
                CalEnd = j
                lblNo(33).Caption = j
                cmbDays(33).Visible = True
                lblNo(33).Visible = True
            Else
                CalEnd = j
                lblNo(40).Caption = j
                cmbDays(40).Visible = True
                lblNo(40).Visible = True
            End If
            'End If
        Case "Saturday"
            'If (i = 6 Or i = 13 Or i = 20 Or i = 27 Or i = 34) Then
            If i <= 6 Then
                If CalStart = 0 And CalStartFlag = False Then
                    CalStart = 6: CalStartFlag = True
                End If
                lblNo(6).Caption = j
                cmbDays(6).Visible = True
                lblNo(6).Visible = True
            ElseIf i <= 13 Then
                lblNo(13).Caption = j
                cmbDays(13).Visible = True
                lblNo(13).Visible = True
            ElseIf i <= 20 Then
                lblNo(20).Caption = j
                cmbDays(20).Visible = True
                lblNo(20).Visible = True
            ElseIf i <= 27 Then
                lblNo(27).Caption = j
                cmbDays(27).Visible = True
                lblNo(27).Visible = True
            ElseIf i <= 34 Then
                CalEnd = j
                lblNo(34).Caption = j
                cmbDays(34).Visible = True
                lblNo(34).Visible = True
            Else
                CalEnd = j
                lblNo(41).Caption = j
                cmbDays(41).Visible = True
                lblNo(41).Visible = True
            End If
            'End If
     End Select
Next i

ExitRtn1:
    
    'MsgBox Err.Description, vbInformation, PayrollAlert

End Sub

Private Sub UpdateDTRRecords()

'If rsDTRRecords.State = 1 Then
'    rsDTRRecords.Close
'End If
'
'' Get DTR records for the selected range of date.
'If InvalidIDSw = False Then ' get all valid IDs
'    StrQuery = "SELECT * FROM vwDTRRecords WHERE InvalidIDFlag = 0 AND DTR_Date BETWEEN '" & _
'        Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "' ORDER BY DTRID "
'Else    ' get all invalid IDs
'    StrQuery = "SELECT * FROM vwDTRRecords WHERE InvalidIDFlag = 1 AND DTR_Date BETWEEN '" & _
'        Format(DTRStart, "mm/dd/yyyy") & "' AND '" & Format(DTREnd, "mm/dd/yyyy") & "' ORDER BY DTRID "
'End If
'Set rsDTRRecords = ReturnRst(StrQuery, cnnstr, mstatic)

End Sub

Private Sub DisplayFiledOT(ByRef rsOT As ADODB.Recordset, Optional Flag1 As Boolean = False)
Dim litem2 As ListItem
Dim Dtrctr As Integer
Dim WithDTR As Integer
Dim Ttime As Double
Dim TDays As Double
Dim TAbsent As Double
Dim TLates As Double
Dim TUndertime As Double
Dim TOvertime As Double
Dim tottime As Double

'If rsOT.RecordCount = 0 Then Exit Sub
tottime = 0

With rsOT
    .Filter = "DTRID = '" & mDTRID & "'"   ' look for dtr number
    If .EOF = True Then
        .Filter = ""
        ListView1(1).ListItems.Clear
        lblclient(9).Caption = Format(tottime, "0.00")
        Exit Sub
    End If
    
    .MoveFirst
    ListView1(1).ListItems.Clear
    While Not (rsOT.EOF Or rsOT.BOF)
        Dtrctr = Dtrctr + 1
        Set litem2 = ListView1(1).ListItems.Add(, , Dtrctr)
        litem2.SubItems(1) = IIf(IsNull(!EmpNo), "NO Emp NO", !EmpNo)
        litem2.SubItems(2) = IIf(IsNull(!StartOTDate), "No start OT Date", (!StartOTDate))
        'litem2.SubItems(3) = IIf(IsNull(!EndOTDate), "NO Start Time", !EndOTDate)
        litem2.SubItems(3) = IIf(IsNull(!StartOTTime), "NO Start OT Time", !StartOTTime)
        litem2.SubItems(4) = IIf(IsNull(!EndOTTime), "NO End OT Time", !EndOTTime)
        litem2.SubItems(5) = IIf(IsNull(!OTHours), "0.00", Format(!OTHours, "0.##"))
        litem2.SubItems(6) = IIf(IsNull(!Remarks), "", !Remarks)
        litem2.SubItems(7) = IIf(IsNull(!Updatedby), "", !Updatedby)
        tottime = tottime + Val(!OTHours)
        rsOT.MoveNext
    Wend

    .Filter = ""
End With

lblclient(9).Caption = Format(tottime, "0.00")

End Sub

Private Sub AddEditOvertime()
Dim StartOT As String
Dim mHr As Integer
Dim mIn As Integer
Dim OTHrs As Double

If cmdSaveEntries.Caption = "Add OT" Then
    AddOTFlag = True
    cmdDeleteOT.Enabled = False
    dtStartOT.Enabled = True
    dtEndOT.Enabled = True
    Frame6.Enabled = True
    txtRemarks.Enabled = True
    cmdSaveEntries.Caption = "Save"
    cmdEditEntries.Caption = "Cancel"
    mskStartOT.Text = "00:00"
    mskEndOT.Text = "00:00"
    txtWorkHours.Text = "0"
    dtStartOT.SetFocus
    
    Exit Sub
End If

StartOT = IIf(IsNull(mskStartOT), "00:00", mskStartOT)
mHr = Val(Left(StartOT, 2))
mMin = Val(Right(StartOT, 2))
If mHr > 23 Or mMin > 59 Then
    MsgBox "Invalid Start Time of Overtime, should be equal or less than 23 hours and maximum of 59 minutes only!", vbInformation, PayrollAlert
    GoTo ExitRtn
End If

EndOT = IIf(IsNull(mskEndOT), "00:00", mskEndOT)
mHr = Val(Left(EndOT, 2))
mMin = Val(Right(EndOT, 2))
If mHr > 23 Or mMin > 59 Then
    MsgBox "Invalid End Time of Overtime, should be equal or less than 23 hours and maximum of 59 minutes only!", vbInformation, PayrollAlert
    GoTo ExitRtn
End If

OTHrs = IIf(IsNull(txtWorkHours.Text), "0", txtWorkHours.Text)

If AddOTFlag = True Then     ' Add New record
StrQuery = "INSERT INTO HRFiledOvertime (EmpNO, DTRID, StartOTDate, EndOTDate, StartOTTime, EndOTTime, OTHours, Remarks, UpdatedBy) " & _
    " VALUES ('" & txtClient(0).Text & "', '" & txtClient(9).Text & "', '" & dtStartOT.VALUE & "', '" & dtEndOT.VALUE & "', '" & StartOT & "', '" & EndOT & "', " & OTHrs & ", '" & txtRemarks.Text & "', '" & LoginName & "')"
Else
StrQuery = "UPDATE HRFiledOvertime SET EmpNO = '" & txtClient(0).Text & "', " & _
    "DTRID = '" & txtClient(9).Text & "', StartOTDate = '" & dtStartOT.VALUE & "', " & _
    "EndOTDate = '" & dtEndOT.VALUE & "', StartOTTime = '" & StartOT & "', " & _
    "EndOTTime = '" & EndOT & "', OTHours = " & OTHrs & ", Remarks = '" & txtRemarks.Text & "', " & _
    "UpdatedBy = '" & LoginName & "' WHERE EmpNO = '" & txtClient(0).Text & "' AND YEAR(StartOTDate) = " & Format(dtStartOT.VALUE, "yyyy") & " and MONTH(StartOTDate) = " & Val(Format(dtStartOT.VALUE, "mm")) & " and day(startotdate) = " & Val(Format(dtStartOT.VALUE, "dd"))
End If

Action_Query StrQuery, cnnstr
rsFiledOT.Requery
If AddOTFlag = True Then
    MsgBox "Overtime was successfully saved!", vbInformation, PayrollAlert
Else
    MsgBox "Overtime was successfully updated!", vbInformation, PayrollAlert
End If
        
AddOTFlag = False

dtStartOT.Enabled = False
dtEndOT.Enabled = False
Frame6.Enabled = False
txtRemarks.Enabled = False
cmdSaveEntries.Caption = "Add OT"
cmdEditEntries.Caption = "Edit OT"
cmdDeleteOT.Enabled = True

Call DisplayFiledOT(rsFiledOT)
Call rsShowFiledOT

ExitRtn:

End Sub

Private Sub rsShowFiledOT()

End Sub

Private Sub cmdSaveEntries_Click()
Dim i As Integer

If optLeave(0).VALUE = False Then ' for Overtime
    Call AddEditOvertime
    Exit Sub
End If

Call UpdateLeaveCalendarEntries

cmbCompany.Enabled = True
ListView1(0).Enabled = True
cmdSaveEntries.Enabled = False

For i = 0 To 41
    cmbDays(i).Enabled = False
Next i
cmdEditEntries.Caption = "Edit Entries"

End Sub

Private Sub UpdateLeaveCalendarEntries()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim rsDayoff As ADODB.Recordset
Dim mString As String
Dim mDayOff As String
Dim mOffDay As Date

On Error GoTo ExitRtn1

mDayOff = ""
If MsgBox("Save changes to Leaves and Dayoff entries?", vbYesNo + vbQuestion, "DTR Leave and Dayoff Entry") <> vbYes Then
    Exit Sub
End If

mString = ""
j = 0
For i = CalStart To IIf(CalStart = 0, CalEnd - 1, IIf((CalEnd + CalStart) > 31, 31 + ((CalEnd + CalStart) - 32), (CalEnd + CalStart)))
    j = j + 1
    mString = LTrim(RTrim(mString)) & " Days" & j & " = '" & Left(cmbDays(i).Text, 1) & "', "
    If Left(cmbDays(i).Text, 1) = "D" And mDayOff = "" Then
        mOffDay = Format(Left(cmbMonth, 2) & "/" & Right("00" & str(j), 2) & "/" & cmbYear, "mm/dd/yyyy")
        mDayOff = Format(mOffDay, "dddd")
    End If
Next i

StrQuery = "SELECT EMPNO from HRDayoffCalendar WHERE ltrim(rtrim(empno)) = '" & Trim$(txtClient(0).Text) & "' and calyearmonth = " & LTrim(RTrim(cmbYear.Text)) & Left(cmbMonth.Text, 2)
Set rsDayoff = ReturnRst(StrQuery, cnnstr, mstatic)

If rsDayoff.RecordCount = 0 Then
    StrQuery = "INSERT INTO HRDayoffCalendar (EmpNo, DTRID, CalYear, CalMonth, CalYearMonth, UpdatedBy) VALUES " & _
        "('" & LTrim(RTrim(txtClient(0).Text)) & "', '" & LTrim(RTrim(txtClient(9).Text)) & "', " & cmbYear.Text & ", " & Val(Left(cmbMonth, 2)) & _
        ", " & LTrim(RTrim(cmbYear.Text)) & Left(cmbMonth.Text, 2) & ", '" & LoginName & "')"
    Action_Query StrQuery, cnnstr
End If
rsDayoff.Close

'Update the schedules.
StrQuery = "UPDATE HRDayoffCalendar SET Empno = LTRIM(RTRIM(EmpNo)), DTRID = '" & LTrim(RTrim(txtClient(9).Text)) & "', " & mString & " UpdatedBy = '" & LoginName & "', Updated = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'" & _
    " WHERE LTRIM(RTRIM(empno)) = '" & LTrim(RTrim(txtClient(0).Text)) & "' and calyearmonth = " & LTrim(RTrim(cmbYear.Text)) & Left(cmbMonth.Text, 2)
Action_Query StrQuery, cnnstr

rsCalendar.Close
StrQuery = "SELECT * FROM HRDayoffCalendar WHERE calyearmonth = " & LTrim(RTrim(cmbYear.Text)) & Left(cmbMonth.Text, 2)
Set rsCalendar = ReturnRst(StrQuery, cnnstr, mstatic)

'Update the HR Employees Dayoff.
StrQuery = "UPDATE HREmployees SET Dayoff = '" & mDayOff & "'" & _
    " WHERE LTRIM(RTRIM(empno)) = '" & txtClient(0).Text & "'"
Action_Query StrQuery, cnnstr

' Display Leave Balance Used
lblclient(12).Caption = "Used VL: " & GetUsedLeave(True) & " SL: " & GetUsedLeave(False)

Exit Sub

ExitRtn1:

End Sub

Private Function GetLeaveRemarks(rsLeave As Recordset, Optional LFlag As Boolean = False, Optional CalYrMo As Long, _
    Optional CalDays As Integer, Optional CalEmpno As String) As String
With rsLeave
    .Filter = ""
    If .RecordCount <> 0 Then
        .MoveFirst
        If LFlag = True Then    ' Change the filter key for Leave Balance Reporting
            .Filter = "calyearmonth = " & CalYrMo & " and EMPNo = '" & CalEmpno & "' and CalDays = " & CalDays
        Else
            .Filter = "calyearmonth = " & Val(LTrim(RTrim(cmbYear.Text)) & LTrim(Left(cmbMonth.Text, 2))) & " and EMPNo LIKE '" & Trim(txtClient(0).Text) & "' and CalDays = " & CurIndexNo
        End If
        If .EOF = False Then
            GetLeaveRemarks = !LeaveRemarks
        Else
            GetLeaveRemarks = ""
        End If
    End If
    If LFlag = True And GetLeaveRemarks = "" Then    ' Change the filter key for Leave Balance Reporting
        GetLeaveRemarks = "No reason."
    End If
End With
End Function

Private Sub DisplayFiledDayoff()
Dim i As Integer
Dim j As Integer
Dim mLDesc As String
Dim mList As String

On Error GoTo Exit1

If rsCalendar.RecordCount = 0 Then
    Exit Sub
End If

rsCalendar.MoveFirst
With rsCalendar
    .Filter = ""
    .Filter = "calyearmonth = " & Val(LTrim(RTrim(cmbYear.Text)) & LTrim(Left(cmbMonth.Text, 2))) & " and EMPNo LIKE '" & Trim(txtClient(0).Text) & "'"
    j = 5
    If .RecordCount <> 0 Then
        For i = CalStart To 41 'CalEnd + 10
            j = j + 1
            If IsNull(.Fields(j)) Then
                mLDesc = ""
            Else
                mLDesc = Trim$(.Fields(j))
            End If
            mLDesc = GetLeaveType(mLDesc)
            If mLDesc = "" Then
                cmbDays(i).ListIndex = 0
            Else
                cmbDays(i).Text = mLDesc
            End If
        Next i
        lblclient(3).Caption = "Updated by: " & .Fields(38)
        lblclient(4).Caption = "Date: " & Format(.Fields(37), "mm/dd/yyyy hh:mm:ss")
    Else
        For i = CalStart To 41 'CalEnd + 10
            j = j + 1
            cmbDays(i).ListIndex = 0
        Next i
    End If
    .Filter = ""
End With

Exit Sub

Exit1:

'MsgBox Err.Description
Resume Next

End Sub

Private Sub cmdSaveRem_Click()
Dim rsDayoff As ADODB.Recordset

On Error GoTo Err1

'LeaveRem(Val(lblNo(Index).Caption)) = txtLeaveRem.Text
If txtLeaveRem.Text = "" Then
    If MsgBox("Delete remarks?", vbYesNo + vbQuestion, "Remarks Updates") = vbNo Then Exit Sub
    txtLeaveRem.Text = "DELETED"
End If
StrQuery = "SELECT EMPNO from HRDayoffRemarks WHERE ltrim(rtrim(EmpNo)) = '" & _
    Trim$(txtClient(0).Text) & "' and calyearmonth = " & LTrim(RTrim(cmbYear.Text)) & _
    Left(cmbMonth.Text, 2) & " and calDays = " & CurIndexNo
Set rsDayoff = ReturnRst(StrQuery, cnnstr, mstatic)

If rsDayoff.RecordCount = 0 Then
    StrQuery = "INSERT INTO HRDayoffRemarks (EmpNo, DTRID, CalYear, CalMonth, CalYearMonth, UpdatedBy, LeaveRemarks, CalDays, LeaveType) VALUES " & _
        "('" & LTrim(RTrim(txtClient(0).Text)) & "', '" & LTrim(RTrim(txtClient(9).Text)) & "', " & cmbYear.Text & ", " & Val(Left(cmbMonth, 2)) & _
        ", " & LTrim(RTrim(cmbYear.Text)) & Left(cmbMonth.Text, 2) & ", '" & LoginName & "', '" & txtLeaveRem.Text & "', " & CurIndexNo & ", '" & Left(cmbDays(DayIndex).Text, 1) & "')"
    Action_Query StrQuery, cnnstr
End If
rsDayoff.Close

'Update the schedules.
StrQuery = "UPDATE HRDayoffRemarks SET Empno = LTRIM(RTRIM(EmpNo)), DTRID = '" & LTrim(RTrim(txtClient(9).Text)) & "', LeaveRemarks = '" & txtLeaveRem.Text & "', UpdatedBy = '" & LoginName & _
    "', Updated = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "', LeaveType = '" & Left(cmbDays(DayIndex).Text, 1) & _
    "' WHERE LTRIM(RTRIM(empno)) = '" & LTrim(RTrim(txtClient(0).Text)) & "' and calyearmonth = " & LTrim(RTrim(cmbYear.Text)) & Left(cmbMonth.Text, 2) & " and calDays = " & CurIndexNo
Action_Query StrQuery, cnnstr

rsLeaveRemarks.Close
StrQuery = "SELECT * FROM HRDayoffRemarks WHERE calyearmonth = " & LTrim(RTrim(cmbYear.Text)) & Left(cmbMonth.Text, 2)
Set rsLeaveRemarks = ReturnRst(StrQuery, cnnstr, mstatic)

MsgBox "Remarks was save!", vbInformation, PayrollAlert

Exit Sub

Err1:

    MsgBox Err.Description, vbInformation, PayrollAlert
    Err.Clear

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
    
    CurIndexNo = 1  ' default to 1
    DayIndex = 1

    InvalidIDSw = False
    FirstHalfFlag = True

    AddOTFlag = False

    dtStartOT.Enabled = False
    dtEndOT.Enabled = False
    Frame6.Enabled = False
    txtRemarks.Enabled = False

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
    
    ' Populate Month.
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
    
'Get all Filed OT.
'StrQuery = "SELECT * FROM HRFiledOvertime ORDER BY EmpNo, StartOTDate"
StrQuery = "SELECT * FROM HRFiledOvertime WHERE Year(StartOTDate) = " & cmbYear.Text & " and Month(StartOTDate) = " & Val(Left(cmbMonth.Text, 2)) & " ORDER BY EmpNo, StartOTDate"
Set rsFiledOT = ReturnRst(StrQuery, cnnstr, mstatic)
        
' Get all Dayoff schedules.
StrQuery = "SELECT * FROM HRDayoffCalendar WHERE calyearmonth = " & LTrim(RTrim(cmbYear.Text)) & LTrim(RTrim(Left(cmbMonth.Text, 2)))
Set rsCalendar = ReturnRst(StrQuery, cnnstr, mstatic)

' Get all filed leaves for the year.
StrQuery = "SELECT a.*, (LTRIM(LTRIM(b.Surname)) + ', ' + LTRIM(RTRIM(b.Firstname)) + ' ' + LEFT(b.Middlename,1)) AS Fullname " & _
    "FROM HRDayoffCalendar a LEFT JOIN HREmployees b ON a.EmpNo = b.EmpNo WHERE b.ClientCode = " & Val(txtClientCode) & " and a.CalYear = " & LTrim(RTrim(cmbYear.Text)) & " ORDER BY b.Surname, b.Firstname, a.CalyearMonth, a.Day_ID "
Set rsLeaveYear = ReturnRst(StrQuery, cnnstr, mstatic)

' Get all Dayoff remarks.
StrQuery = "SELECT * FROM HRDayoffRemarks WHERE calyearmonth = " & LTrim(RTrim(cmbYear.Text)) & LTrim(RTrim(Left(cmbMonth.Text, 2)))
Set rsLeaveRemarks = ReturnRst(StrQuery, cnnstr, mstatic)

Call cmdRefresh_Click

End Sub

Private Function GetUsedLeave(Optional isVLType As Boolean = True) As Double
Dim j As Integer, sFld As Integer
Dim mString As String
Dim TotUseLeave As Double

GetUsedLeave = 0

On Error GoTo ErrLeave1

If rsLeaveYear.RecordCount = 0 Then Exit Function

sFld = 6    ' start of Days fields
TotUseLeave = 0

With rsLeaveYear
    .Filter = ""
    .MoveFirst
    .Filter = "EmpNo = '" & Trim$(txtClient(0).Text) & "'"
    If .EOF = False Then
        Do While Not .EOF
            For j = 6 To 36
                If isVLType = True Then ' Count for VL leave
                    If InStr(1, HalfWithPay, Left(.Fields(j).VALUE, 1)) <> 0 Then
                        If Left(.Fields(j).VALUE, 1) = "Y" Then
                            TotUseLeave = TotUseLeave + 0.5
                        End If
                    ElseIf InStr(1, VLWithPay, Left(.Fields(j).VALUE, 1)) <> 0 Then
                        TotUseLeave = TotUseLeave + 1
                    End If
                Else ' Count for SL leave
                    If InStr(1, HalfWithPay, Left(.Fields(j).VALUE, 1)) <> 0 Then
                        If Left(.Fields(j).VALUE, 1) = "F" Then
                            TotUseLeave = TotUseLeave + 0.5
                        End If
                    ElseIf InStr(1, SLWithPay, Left(.Fields(j).VALUE, 1)) <> 0 Then
                        TotUseLeave = TotUseLeave + 1
                    End If
                End If
            Next j
            .MoveNext
        Loop
    End If
    .Filter = ""
End With

GetUsedLeave = TotUseLeave

Exit Function

ErrLeave1:

    MsgBox Err.Description, vbInformation, PayrollAlert
    Err.Clear
    
End Function

Private Sub PrintLeaveBalance()
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim j As Integer, i As Integer, k As Integer
Dim prevName As String, mLeaveDate As String, mFld As String, mReason As String
Dim WLeave As Boolean, LType As String, LDays As Double, NameFlag As Boolean
Dim TotSL As Double, TotVL As Double
Dim rsLeaveRemarksYear As Recordset

On Error GoTo ErrLeave1

If rsLeaveYear.RecordCount = 0 Then MsgBox "No SL/VL records to print.", vbInformation, PayrollAlert: Exit Sub

StrQuery = "SELECT * FROM HRDayoffRemarks WHERE calyear = " & cmbYear.Text
Set rsLeaveRemarksYear = ReturnRst(StrQuery, cnnstr, mstatic)

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)

ctr = 8
i = 1

mHead = CompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Schedule of Sickness and Vacation Leaves for All Employees for the Year " & cmbYear.Text
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy hh:mm:ss")
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "COMPANY NAME: " & cmbCompany.Text
Call PrintReportHeader(xlSheet, xlApp1, "A5", mHead, "", xlLeft, , False, False, False, , , 10, True)

Call PrintReportHeader(xlSheet, xlApp1, "A6:A7", "Employee Name", "", xlCenter, xlCenter, True, False, True, 15, , 10, False, 30)
Call PrintReportHeader(xlSheet, xlApp1, "B6:B7", "Leave Date", "", xlCenter, xlCenter, True, False, True, 15, , 10, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "C6:C7", "Leave Days", "", xlCenter, xlCenter, True, True, True, 15, , 10, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "D6:D7", "Leave Type", "", xlCenter, xlCenter, True, True, True, 15, , 10, False, 15)
Call PrintReportHeader(xlSheet, xlApp1, "E6:E7", "Reason", "", xlCenter, xlCenter, True, False, True, 15, , 10, False, 30)
Call PrintReportHeader(xlSheet, xlApp1, "F6:F7", "Filed By", "", xlCenter, xlCenter, True, False, True, 15, , 10, False, 30)

prevName = ""
TotSL = 0: TotVL = 0
With rsLeaveYear
    .Filter = ""
    .MoveFirst
    Do While Not .EOF
        ' Check for filed leave.
        k = 0: NameFlag = False
        WLeave = False: LDays = 0: LType = "Vacation Leave"
        
'        ' Print Leave Totals.
'        If prevName <> "" And prevName <> !FullName Then
'            xlSheet.Cells(ctr, 1).VALUE = "Total VL: " & TotVL & "  Total SL: " & TotSL
'            TotSL = 0: TotVL = 0
'            ctr = ctr + 1
'        End If
        
        For j = 6 To 36
            k = k + 1
            mLeaveDate = Right("00" & LTrim(str(!CalMonth)), 2) & "/" & Right("00" & LTrim(str(k)), 2) & "/" & LTrim(!CalYear)
            ' Count for Halfday leave.
            If InStr(1, HalfWithPay, Left(.Fields(j).VALUE, 1)) <> 0 Then
                LDays = 0.5
                mLDesc = Trim$(.Fields(j))
                If Left(.Fields(j).VALUE, 1) = "F" Then
                    TotSL = TotSL + 0.5
                Else
                    TotVL = TotVL + 0.5
                End If
                LType = GetLeaveType(Trim$(Left(.Fields(j).VALUE, 1)))
                mReason = GetLeaveRemarks(rsLeaveRemarksYear, True, LTrim(!CalYear) & Right("00" & LTrim(str(!CalMonth)), 2), Right("00" & LTrim(str(k)), 2), !EmpNo)
                'LType = "Halfday Leave"
                WLeave = True
            ' Count for Vacation Leave
            ElseIf InStr(1, VLWithPay, Left(.Fields(j).VALUE, 1)) <> 0 Then
                LDays = 1#
                TotVL = TotVL + 1
                LType = GetLeaveType(Trim$(Left(.Fields(j).VALUE, 1)))
                mReason = GetLeaveRemarks(rsLeaveRemarksYear, True, LTrim(!CalYear) & Right("00" & LTrim(str(!CalMonth)), 2), Right("00" & LTrim(str(k)), 2), !EmpNo)
                'LType = "Vacation Leave"
                WLeave = True
            ' Count for Sickleave
            Else
                If InStr(1, SLWithPay, Left(.Fields(j).VALUE, 1)) <> 0 Then
                    LDays = 1#
                    TotSL = TotSL + 1
                    'LType = "Sick Leave"
                    LType = GetLeaveType(Trim$(Left(.Fields(j).VALUE, 1)))
                    mReason = GetLeaveRemarks(rsLeaveRemarksYear, True, LTrim(!CalYear) & Right("00" & LTrim(str(!CalMonth)), 2), Right("00" & LTrim(str(k)), 2), !EmpNo)
                    WLeave = True
                End If
            End If
            'Print Leave Details.
            If WLeave = True Then
                ' Print Employee Name.
                If prevName <> !FullName Then
                    If NameFlag = False Then
                        ctr = ctr + 1
                        i = 1
                        xlSheet.Cells(ctr, 1).VALUE = !FullName
                        ctr = ctr + 1
                        prevName = !FullName
                        NameFlag = True
                    End If
                End If
                xlSheet.Cells(ctr, 1).VALUE = i
                xlSheet.Cells(ctr, 2).VALUE = mLeaveDate
                xlSheet.Cells(ctr, 3).VALUE = LDays
                xlSheet.Cells(ctr, 4).VALUE = LType
                xlSheet.Cells(ctr, 5).VALUE = mReason '"Reason"
                xlSheet.Cells(ctr, 6).VALUE = !Updatedby & " (" & !Updated & ")"
                ctr = ctr + 1
                i = i + 1
            End If
            WLeave = False
        Next j
        .MoveNext
    Loop
'    xlSheet.Cells(ctr, 1).VALUE = "Total VL: " & TotVL & "  Total SL: " & TotSL
    .Filter = ""
End With
ctr = ctr + 2

'Box it.
mRange = "A6:F" & ctr
Call PlaceBox(xlSheet, xlApp1, mRange)
xlSheet.Range("A1").Select

Exit Sub

ErrLeave1:

    MsgBox Err.Description, vbInformation, PayrollAlert
    Err.Clear

End Sub

Private Sub GetEmployeeBasic()
Dim rsInvCode As ADODB.Recordset
Dim i As Integer

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

StrQuery = "SELECT * FROM vwEmployeeBasicSalary ORDER BY Fullname"
'StrQuery = "SELECT * FROM qryHREmployeeHired WHERE EmpStatus = 1 ORDER BY EmpName "
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

StrQuery = "Select * from HRLeaveType"
Set rsLeaveType = ReturnRst(StrQuery, cnnstr, mstatic)

For i = 0 To 41
    cmbDays(i).AddItem ""
    cmbDays(i).Enabled = False
Next i

Do While Not rsLeaveType.EOF
    For i = 0 To 41
        cmbDays(i).AddItem rsLeaveType!leavetype & "-" & rsLeaveType!leavetypedesc
    Next i
    rsLeaveType.MoveNext
Loop
rsLeaveType.MoveFirst

End Sub

Private Function GetLeaveType(ByVal LType As String) As String
If LType = "" Then
    GetLeaveType = ""
Else
With rsLeaveType
    .MoveFirst
    .Find "LeaveType = '" & LType & "'"
    If .EOF = False Then
        GetLeaveType = LType & "-" & LTrim(RTrim(!leavetypedesc))
    Else
        GetLeaveType = ""
    End If
End With
End If
End Function

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
            
            ' Refresh Leave Records.
            StrQuery = "SELECT a.*, (LTRIM(LTRIM(b.Surname)) + ', ' + LTRIM(RTRIM(b.Firstname)) + ' ' + LEFT(b.Middlename,1)) AS Fullname " & _
                "FROM HRDayoffCalendar a LEFT JOIN HREmployees b ON a.EmpNo = b.EmpNo WHERE b.ClientCode = " & Val(txtClientCode) & " and a.CalYear = " & IIf(cmbYear.Text = "", Year(Now), LTrim(RTrim(cmbYear.Text))) & " ORDER BY b.Surname, b.Firstname, a.CalyearMonth, a.Day_ID"
            Set rsLeaveYear = ReturnRst(StrQuery, cnnstr, mstatic)
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
    'Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!EmpName), "Blank Name", mRs!EmpName))
    Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName))
        litem1.SubItems(1) = mRs(0)
        ActiveCtr = ActiveCtr + 1
    mRs.MoveNext
Wend

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

Private Sub DisplayBasicData(ByVal SelectedEmployee As String, _
    Optional recSw As Boolean)
Dim litem2 As ListItem
Dim LineNo As Integer

LineNo = 0

If optLeave(0).VALUE = True Then
    If rsEmployeeBasicData.RecordCount = 0 Then Exit Sub
End If
With rsEmployeeBasicData
    If recSw = False Then
        .MoveFirst
        '.Find "Empname = '" & LTrim(RTrim(SelectedEmployee)) & "'"
        .Find "Fullname = '" & LTrim(RTrim(SelectedEmployee)) & "'"
    End If
    If Not .EOF And Not .BOF Then
        mDTRID = IIf(IsNull(!DTRID), "00000000", !DTRID)
        txtClient(0) = IIf(IsNull(!EmpNo), "", !EmpNo)
        txtClient(2) = LTrim(RTrim(SelectedEmployee))
        txtClient(8) = IIf(IsNull(!ShiftName), "", !ShiftName) ' get the shift Name
        txtClient(9) = IIf(IsNull(!DTRID), "", !DTRID) ' get DTR ID
'        If rsDTRRecords.State = 1 And rsDTRRecords.RecordCount <> 0 Then
'            'Get all DTR per employee.
'            rsDTRRecords.Filter = ""
'            rsDTRRecords.MoveFirst
'            If InvalidIDSw = False Then
'                If chkNoFilter.VALUE = 0 Then
'                    rsDTRRecords.Filter = "DTRID = '" & mDTRID & "'"   ' look for dtr number
'                End If
'            End If
'            lstDTRDisplay.ListItems.Clear
'            While Not (rsDTRRecords.EOF Or rsDTRRecords.BOF)
'                LineNo = LineNo + 1
'                'Set litem2 = lstDTRDisplay.ListItems.Add(, , LineNo)
'                Set litem2 = lstDTRDisplay.ListItems.Add(, , IIf(IsNull(rsDTRRecords!Machine_ID), "X", rsDTRRecords!Machine_ID))
'                'litem2.SubItems(1) = IIf(IsNull(rsDTRRecords!Machine_ID), "X", rsDTRRecords!Machine_ID)
'                'litem2.SubItems(1) = IIf(IsNull(rsDTRRecords!DTR_Mode), "X", rsDTRRecords!DTR_Mode)
'                litem2.SubItems(1) = IIf(IsNull(rsDTRRecords!DTRID), "NO DTR ID", rsDTRRecords!DTRID)
'                litem2.SubItems(2) = IIf(IsNull(rsDTRRecords!DTR_Date), "NO DTR DATE", Format(rsDTRRecords!DTR_Date, "mm/dd/yyyy"))
'                litem2.SubItems(3) = IIf(IsNull(rsDTRRecords!TimeIN), "NO Time In", rsDTRRecords!TimeIN)
'                litem2.SubItems(4) = IIf(IsNull(rsDTRRecords!TimeOut), "NO Time Out", rsDTRRecords!TimeOut)
'                litem2.SubItems(5) = IIf(IsNull(rsDTRRecords!WorkHours), "0.00", rsDTRRecords!WorkHours)
'                litem2.SubItems(6) = IIf(IsNull(rsDTRRecords!ImportDate), "", Format(rsDTRRecords!ImportDate, "mm/dd/yyyy"))
'                litem2.SubItems(7) = IIf(IsNull(rsDTRRecords!Remarks), "", rsDTRRecords!Remarks)
'                rsDTRRecords.MoveNext
'            Wend
        'End If
        'RecordCount = rsDTRRecords.RecordCount
        'rsDTRRecords.Filter = ""
    End If
End With

If optLeave(0).VALUE = False Then ' OT
    Call DisplayFiledOT(rsFiledOT)
Else
    Call DisplayFiledDayoff
End If

'lblRecount.Caption = RecordCount

End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    If Index = 1 Then
        If ListView1(1).ListItems.Count <> 0 Then
            dtStartOT.VALUE = ListView1(1).ListItems(ListView1(1).SelectedItem.Index).ListSubItems(2).Text
            mskStartOT.Text = Trim$(ListView1(1).ListItems(ListView1(1).SelectedItem.Index).ListSubItems(3).Text)
            mskEndOT.Text = Trim$(ListView1(1).ListItems(ListView1(1).SelectedItem.Index).ListSubItems(4).Text)
            txtWorkHours.Text = Trim$(ListView1(1).ListItems(ListView1(1).SelectedItem.Index).ListSubItems(5).Text)
            txtRemarks.Text = ListView1(1).ListItems(ListView1(1).SelectedItem.Index).ListSubItems(6).Text
        End If
    Else
        InvalidIDSw = False
        'Call UpdateDTRRecords
        Call DisplayBasicData(Item.Text, False)
        ' Display Leave Balance Used
        lblclient(12).Caption = "USED VL:  " & GetUsedLeave(True) & "   USED SL:  " & GetUsedLeave(False) & " (" & cmbYear.Text & ")"
    End If
End Sub

Private Sub optLeave_Click(Index As Integer)

If optLeave(0).VALUE = True Then ' Leave
    Frame4.Visible = False
    cmdrefresh.Caption = "Display Leaves"
    cmdEditEntries.Caption = "Edit Entries"
    cmdSaveEntries.Caption = "Save Entries"
    cmdSaveEntries.Enabled = False
    Label1.Caption = "Employee Leave and Dayoff Entries"
    Frame5.Visible = True
Else                            ' Overtime
    Frame4.Visible = True
    cmdrefresh.Caption = "Display OT"
    cmdEditEntries.Caption = "Edit OT"
    cmdSaveEntries.Caption = "Add OT"
    cmdSaveEntries.Enabled = True
    Label1.Caption = "Employee Overtime Entries"
    Frame5.Visible = False
End If

End Sub

Private Sub mskStartOT_GotFocus()
mskStartOT.BackColor = &HFFFFC0
Call mskHLigths(mskStartOT)
End Sub

'Private Sub mskStartOT_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub

Private Sub mskStartOT_LostFocus()
mskStartOT.BackColor = vbWhite
End Sub

Private Sub mskEndOT_GotFocus()
mskEndOT.BackColor = &HFFFFC0
Call mskHLigths(mskEndOT)
End Sub

'Private Sub mskEndOT_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub

Private Sub mskEndOT_LostFocus()
mskEndOT.BackColor = vbWhite
End Sub

Private Sub txtWorkHours_GotFocus()
txtWorkHours.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtWorkHours
Call txtHLigths(txtWorkHours)
End Sub

Private Sub txtWorkHours_LostFocus()
txtWorkHours.BackColor = vbWhite
End Sub

