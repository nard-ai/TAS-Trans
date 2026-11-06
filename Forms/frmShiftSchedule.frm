VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmShiftSchedule 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shift Schedule"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmShiftSchedule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   6960
   Begin MSComctlLib.ListView lstSchedule 
      Height          =   2250
      Left            =   120
      TabIndex        =   65
      Top             =   4440
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   3969
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Shift Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Hours"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Start "
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "End "
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Mid"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Remarks"
         Object.Width           =   4057
      EndProperty
   End
   Begin VB.CommandButton cmdlast 
      Height          =   255
      Left            =   6480
      Picture         =   "frmShiftSchedule.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Last record"
      Top             =   6720
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Height          =   255
      Left            =   6240
      Picture         =   "frmShiftSchedule.frx":0A14
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Next record"
      Top             =   6720
      Width           =   255
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   255
      Left            =   6000
      Picture         =   "frmShiftSchedule.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Previous record"
      Top             =   6720
      Width           =   255
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   255
      Left            =   5760
      Picture         =   "frmShiftSchedule.frx":0CA8
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Top record"
      Top             =   6720
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      ForeColor       =   &H8000000F&
      Height          =   3855
      Left            =   120
      TabIndex        =   33
      Top             =   600
      Width           =   6735
      Begin VB.TextBox txtUpdatedBy 
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
         Left            =   4500
         TabIndex        =   56
         Top             =   3360
         Width           =   1950
      End
      Begin VB.CheckBox chkFlexible 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Flexible Time"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         ToolTipText     =   "Flexible Schedule"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtSCode 
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
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "MS001"
         Top             =   240
         Width           =   630
      End
      Begin VB.TextBox txtWorkHours 
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
         Left            =   5880
         TabIndex        =   49
         Text            =   "8.00"
         Top             =   240
         Width           =   630
      End
      Begin VB.TextBox txtShiftCode 
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
         Left            =   0
         MaxLength       =   15
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.TextBox txtShiftName 
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
         Left            =   2880
         TabIndex        =   2
         Text            =   "MS001 [8:30 - 5:30]"
         Top             =   225
         Width           =   1830
      End
      Begin VB.TextBox txtLastUpdate 
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
         Left            =   1140
         TabIndex        =   26
         Top             =   3360
         Width           =   1950
      End
      Begin VB.CheckBox chkUndertime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "With UDT Grace Period"
         Height          =   375
         Left            =   4200
         TabIndex        =   16
         ToolTipText     =   "With UDT Grace Period"
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox txtUndertime 
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
         Left            =   5760
         TabIndex        =   7
         Text            =   "5"
         Top             =   1680
         Width           =   510
      End
      Begin VB.CheckBox chkGracePeriod 
         BackColor       =   &H00FFFFFF&
         Caption         =   "With Grace Period"
         Height          =   435
         Left            =   1080
         TabIndex        =   15
         ToolTipText     =   "with Grace Period"
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtGracePeriod 
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
         Left            =   2340
         TabIndex        =   6
         Text            =   "10"
         Top             =   1680
         Width           =   510
      End
      Begin VB.CheckBox chkNoNightDiff 
         BackColor       =   &H00FFFFFF&
         Caption         =   "With Night Differential"
         Height          =   375
         Left            =   3240
         TabIndex        =   23
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "End Time"
         Enabled         =   0   'False
         Height          =   615
         Left            =   8760
         TabIndex        =   45
         Top             =   2280
         Width           =   2070
         Begin VB.CheckBox chkNoOutNight 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Out"
            Height          =   255
            Left            =   1080
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start Time ------ End Time"
         Height          =   615
         Left            =   4380
         TabIndex        =   44
         Top             =   2280
         Width           =   2070
         Begin VB.CheckBox chkNoInNight 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No In"
            Height          =   255
            Left            =   1920
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSMask.MaskEdBox mskStartNightDiff 
            Height          =   285
            Left            =   120
            TabIndex        =   11
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
         Begin MSMask.MaskEdBox mskEndNightDiff 
            Height          =   285
            Left            =   1200
            TabIndex        =   59
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
      Begin VB.CheckBox chkNoOvertime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "With OT"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "End Time"
         Enabled         =   0   'False
         Height          =   615
         Left            =   8640
         TabIndex        =   43
         Top             =   1560
         Width           =   2070
         Begin VB.CheckBox chkNoOutOT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Out"
            Height          =   255
            Left            =   1080
            TabIndex        =   22
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start Time ------ End Time"
         Height          =   615
         Left            =   1020
         TabIndex        =   42
         Top             =   2280
         Width           =   2070
         Begin VB.CheckBox chkNoInOT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No In"
            Height          =   255
            Left            =   1920
            TabIndex        =   21
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSMask.MaskEdBox mskStartOT 
            Height          =   285
            Left            =   120
            TabIndex        =   10
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
            TabIndex        =   58
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
      Begin VB.CheckBox chkNoBreaktime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No Breaktime"
         Height          =   375
         Left            =   7200
         TabIndex        =   17
         Top             =   840
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "End Time"
         Enabled         =   0   'False
         Height          =   615
         Left            =   10800
         TabIndex        =   41
         Top             =   720
         Width           =   2070
         Begin VB.CheckBox chkNoOutBreak 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Out"
            Height          =   255
            Left            =   1080
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin MSMask.MaskEdBox mskEndBreak 
            Height          =   285
            Left            =   120
            TabIndex        =   9
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start Time"
         Enabled         =   0   'False
         Height          =   615
         Left            =   8340
         TabIndex        =   40
         Top             =   720
         Width           =   2070
         Begin VB.CheckBox chkNoInBreak 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No In"
            Height          =   255
            Left            =   1140
            TabIndex        =   18
            Top             =   240
            Width           =   735
         End
         Begin MSMask.MaskEdBox mskStartBreak 
            Height          =   285
            Left            =   120
            TabIndex        =   8
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
      Begin VB.CheckBox chkNoInOut 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No In/Out"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   960
         Width           =   855
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "End Time"
         Height          =   735
         Left            =   3960
         TabIndex        =   38
         Top             =   720
         Width           =   2600
         Begin VB.CheckBox chkNoUDT 
            BackColor       =   &H00FFFFFF&
            Caption         =   " No UDT"
            Height          =   495
            Left            =   1080
            TabIndex        =   14
            ToolTipText     =   "No Undertime"
            Top             =   120
            Value           =   1  'Checked
            Width           =   735
         End
         Begin MSMask.MaskEdBox mskEndTime 
            Height          =   285
            Left            =   120
            TabIndex        =   5
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
         Begin VB.CheckBox chkMidNight 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mid Night"
            Height          =   495
            Left            =   1800
            TabIndex        =   60
            ToolTipText     =   "Mid Night schedule"
            Top             =   120
            Value           =   1  'Checked
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start Time"
         Height          =   735
         Left            =   1020
         TabIndex        =   37
         Top             =   720
         Width           =   1950
         Begin VB.CheckBox chkNoLate 
            BackColor       =   &H00FFFFFF&
            Caption         =   " No Late"
            Height          =   495
            Left            =   1140
            TabIndex        =   13
            ToolTipText     =   "No Late"
            Top             =   120
            Value           =   1  'Checked
            Width           =   735
         End
         Begin MSMask.MaskEdBox mskStartTime 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            AllowPrompt     =   -1  'True
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
      Begin VB.TextBox txtRemarks 
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
         Height          =   285
         Left            =   1140
         TabIndex        =   12
         Top             =   2985
         Width           =   5310
      End
      Begin MSMask.MaskEdBox mskHalfDay 
         Height          =   285
         Left            =   3075
         TabIndex        =   4
         Top             =   1080
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   12582912
         AllowPrompt     =   -1  'True
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3480
         TabIndex        =   57
         Top             =   3360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Undertime:"
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
         Height          =   195
         Index           =   7
         Left            =   7440
         TabIndex        =   54
         Top             =   2520
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Grace Period:"
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
         Height          =   195
         Index           =   3
         Left            =   7440
         TabIndex        =   53
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Code (5 chars)"
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
         Height          =   390
         Index           =   6
         Left            =   240
         TabIndex        =   52
         Top             =   240
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Mid - Time"
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
         Height          =   195
         Index           =   5
         Left            =   3120
         TabIndex        =   51
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours / day:"
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
         Height          =   195
         Index           =   4
         Left            =   4920
         TabIndex        =   50
         Top             =   240
         Width           =   900
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   120
         X2              =   6480
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Mins."
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
         Height          =   195
         Index           =   1
         Left            =   5880
         TabIndex        =   47
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Mins."
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
         Height          =   195
         Index           =   2
         Left            =   2475
         TabIndex        =   46
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Update:"
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
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   39
         Top             =   3360
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Name"
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
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   35
         Top             =   240
         Width           =   990
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   2985
         Width           =   675
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   600
         Picture         =   "frmShiftSchedule.frx":0DF2
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   2340
         Picture         =   "frmShiftSchedule.frx":0F3C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1170
         Picture         =   "frmShiftSchedule.frx":146E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmShiftSchedule.frx":19A0
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "New"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   2940
         Picture         =   "frmShiftSchedule.frx":1ED2
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   1755
         Picture         =   "frmShiftSchedule.frx":201C
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Delete"
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
Attribute VB_Name = "frmShiftSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsShiftSched As ADODB.Recordset
Dim mShiftCode As Integer
Dim EditFlag As Boolean
Dim AddFlag As Boolean

Private Sub chkGracePeriod_Click()
If chkGracePeriod.VALUE = 1 Then
    txtGracePeriod.Enabled = False
Else
    If AddFlag = True Or EditFlag = True Then
        txtGracePeriod.Enabled = True
    End If
End If
End Sub

'Private Sub chkNoBreaktime_Click()
'If chkNoBreaktime.VALUE = 1 Then
'    Frame4.Enabled = False
'    Frame5.Enabled = False
'Else
'    Frame4.Enabled = True
'    Frame5.Enabled = True
'End If
'End Sub

'Private Sub chkNoIn_Click()
'If chkNoIn.VALUE = 1 Then
'    mskStartTime.Enabled = False
'Else
'    mskStartTime.Enabled = True
'End If
'End Sub

'Private Sub chkNoInOut_Click()
'If chkNoInOut.VALUE = 1 Then
'    Frame2.Enabled = False
'    Frame3.Enabled = False
'Else
'    Frame2.Enabled = True
'    Frame3.Enabled = True
'End If
'End Sub

'Private Sub chkNoNightDiff_Click()
'If chkNoNightDiff.VALUE = 1 Then
'    Frame8.Enabled = False
'    Frame9.Enabled = False
'Else
'    Frame8.Enabled = True
'    Frame9.Enabled = True
'End If
'End Sub

'Private Sub chkNoOvertime_Click()
'If chkNoOvertime.VALUE = 1 Then
'    Frame6.Enabled = False
'    Frame7.Enabled = False
'Else
'    Frame6.Enabled = True
'    Frame7.Enabled = True
'End If
'End Sub

'Private Sub chkUndertime_Click()
'If chkUndertime.VALUE = 1 Then
'    txtUndertime.Enabled = False
'Else
'    If AddFlag = True Or EditFlag = True Then
'        txtUndertime.Enabled = True
'    End If
'End If
'End Sub

Private Sub cmdRefresh_Click()

If rsShiftSched.State = 1 Then
    rsShiftSched.Close
End If
StrQuery = "SELECT * FROM HRShiftSchedule WHERE ShiftStatus = 1 ORDER BY ShiftName"
Set rsShiftSched = ReturnRst(StrQuery, cnnstr, mstatic)

Call DisplaySchedules(rsShiftSched)
Showdata rsShiftSched

Call EnableDisableButtons(True)
Call EnableDisableFields(False)

MsgBox "Shift Schedule was updated!", vbInformation, PayrollAlert

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsShiftSched = Nothing
End Sub

Private Sub mskEndBreak_GotFocus()
mskEndBreak.BackColor = &HFFFFC0
Call mskHLigths(mskEndBreak)
End Sub

Private Sub mskEndBreak_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskEndBreak_LostFocus()
mskEndBreak.BackColor = vbWhite
End Sub

Private Sub mskEndNightDiff_GotFocus()
mskEndNightDiff.BackColor = &HFFFFC0
Call mskHLigths(mskEndNightDiff)
End Sub

Private Sub mskEndNightDiff_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskEndNightDiff_LostFocus()
mskEndNightDiff.BackColor = vbWhite
End Sub

Private Sub mskEndOT_GotFocus()
mskEndOT.BackColor = &HFFFFC0
Call mskHLigths(mskEndOT)
End Sub

Private Sub mskEndOT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskEndOT_LostFocus()
mskEndOT.BackColor = vbWhite
End Sub

Private Sub mskEndTime_GotFocus()
mskEndTime.BackColor = &HFFFFC0
Call mskHLigths(mskEndTime)
End Sub

Private Sub mskEndTime_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskEndTime_LostFocus()
mskEndTime.BackColor = vbWhite
End Sub

Private Sub mskStartBreak_GotFocus()
mskStartBreak.BackColor = &HFFFFC0
Call mskHLigths(mskStartBreak)
End Sub

Private Sub mskStartBreak_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskStartBreak_LostFocus()
mskStartBreak.BackColor = vbWhite
End Sub

Private Sub mskStartNightDiff_GotFocus()
mskStartNightDiff.BackColor = &HFFFFC0
Call mskHLigths(mskStartNightDiff)
End Sub

Private Sub mskStartNightDiff_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskStartNightDiff_LostFocus()
mskStartNightDiff.BackColor = vbWhite
End Sub

Private Sub mskStartOT_GotFocus()
mskStartOT.BackColor = &HFFFFC0
Call mskHLigths(mskStartOT)
End Sub

Private Sub mskStartOT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskStartOT_LostFocus()
mskStartOT.BackColor = vbWhite
End Sub

Private Sub mskStartTime_GotFocus()
mskStartTime.BackColor = &HFFFFC0
Call mskHLigths(mskStartTime)
End Sub

Private Sub mskHalfDay_GotFocus()
mskHalfDay.BackColor = &HFFFFC0
Call mskHLigths(mskHalfDay)
End Sub

Private Sub mskHalfDay_LostFocus()
mskHalfDay.BackColor = vbWhite
End Sub

Private Sub mskHalfDay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskStartTime_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskStartTime_LostFocus()
mskStartTime.BackColor = vbWhite
End Sub

Private Sub txtGracePeriod_GotFocus()
txtGracePeriod.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtGracePeriod
End Sub

'Private Sub txtGracePeriod_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub

Private Sub txtGracePeriod_LostFocus()
txtGracePeriod.BackColor = vbWhite
End Sub

Private Sub txtShiftName_GotFocus()
txtShiftName.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtShiftName
Call txtHLigths(txtShiftName)
End Sub

Private Sub txtSCode_GotFocus()
txtSCode.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtSCode
Call txtHLigths(txtSCode)
End Sub

Private Sub txtWorkHours_GotFocus()
txtWorkHours.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtWorkHours
Call txtHLigths(txtWorkHours)
End Sub

'Private Sub txtShiftName_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub

Private Sub txtWorkHours_LostFocus()
txtWorkHours.BackColor = vbWhite
End Sub

Private Sub txtShiftName_LostFocus()
txtShiftName.BackColor = vbWhite
End Sub

Private Sub txtSCode_LostFocus()
txtSCode.BackColor = vbWhite
End Sub

Private Sub cmdAdd_Click()

'On Error GoTo Errh

AddFlag = True
EditFlag = False
Frame1.Enabled = True

Call ClearScreen
Call EnableDisableButtons(False)
Call EnableDisableFields(True)
txtSCode.SetFocus

End Sub

Private Sub cmdCancel_Click()

AddFlag = False
EditFlag = False
Frame1.Enabled = False

MsgBox "Updates was cancelled!", vbInformation, PayrollAlert

If rsShiftSched.EOF = False Then
    Showdata rsShiftSched
End If

Call EnableDisableButtons(True)

End Sub

Private Sub cmdDelete_Click()
Dim rsshift As ADODB.Recordset
Dim res

res = MsgBox("Delete Shift " & txtShiftName.Text & "?", vbQuestion + vbYesNo, PayrollAlert)
If res = vbYes Then
    ' Check first whether shift code not in used by any employee.
    StrQuery = "SELECT Surname, Firstname, shiftcode FROM HREmployees WHERE EmpStatus = 1 And ShiftCode = '" & txtShiftCode.Text & "'"
    Set rsshift = ReturnRst(StrQuery, cnnstr, mstatic)
    If rsshift.RecordCount <> 0 Then
        MsgBox "Shift Schedule was still used by " & rsshift!Surname & ", " & rsshift!FirstName & ". Cannot be deleted.", vbInformation, PayrollAlert
    Else
        StrQuery = "Update HRShiftSchedule SET ShiftStatus = 0 WHERE ShiftCode = '" & txtShiftCode.Text & "'"
        Call Action_Query(StrQuery, cnnstr)
        Call Form_Load
        MsgBox "Shift " & txtShiftName.Text & " was successfully deleted.", vbInformation, PayrollAlert
    End If
    Set rsshift = Nothing
End If
End Sub

Private Sub cmdEdit_Click()

EditFlag = True
AddFlag = False
Frame1.Enabled = True

Call EnableDisableButtons(False)
Call EnableDisableFields(True)
txtSCode.SetFocus

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
rsShiftSched.MoveFirst
cmdfirst.Enabled = False
cmdPrev.Enabled = False
cmdlast.Enabled = True
cmdNext.Enabled = True
Showdata rsShiftSched
End Sub

Private Sub cmdlast_Click()
rsShiftSched.MoveLast
cmdlast.Enabled = False
cmdNext.Enabled = False
cmdPrev.Enabled = True
cmdfirst.Enabled = True
Showdata rsShiftSched
End Sub

Private Sub cmdNext_Click()
On Error GoTo X
 
rsShiftSched.MoveNext
cmdPrev.Enabled = True
cmdfirst.Enabled = True
If rsShiftSched.EOF Then
    rsShiftSched.MoveLast
    cmdNext.Enabled = False
    cmdlast.Enabled = False
    MsgBox "End of file encountered", vbInformation, CompanyName
End If
Showdata rsShiftSched

X:
    Exit Sub
 
End Sub

Private Sub cmdPrev_Click()

On Error GoTo X

rsShiftSched.MovePrevious
cmdNext.Enabled = True
cmdlast.Enabled = True
If rsShiftSched.BOF Then
    rsShiftSched.MoveFirst
    cmdPrev.Enabled = False
    cmdfirst.Enabled = False
    MsgBox "Beginning of file encountered", vbInformation, CompanyName
End If
    
assignmentnum = rsShiftSched!ShiftCode
Showdata rsShiftSched

X:
    Exit Sub
    
End Sub

Private Sub cmdSave_Click()

If txtSCode.Text = "" Or txtSCode.Text = "<No Code>" Then MsgBox "Shift Code should not be blank. ", vbInformation: txtSCode.SetFocus: Exit Sub
If txtWorkHours.Text = "" Then MsgBox "Work Hours should not be blank. ", vbInformation: txtWorkHours.SetFocus: Exit Sub

'Add new record.
Call SaveNewShiftCode

Call EnableDisableButtons(True)
Call EnableDisableFields(False)

End Sub

Private Sub save()

cmdadd.Enabled = True
cmdEdit.Enabled = True
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
txtEmpNo.Enabled = False
mskassigndate.Enabled = False
txtAssignType.Enabled = False
txtRemarks.Enabled = False
mskendassigndate.Enabled = False
'cmbcountycode.Enabled = False
cmbempname.Enabled = False
'cmbprojectcode.Enabled = False
'cmbclientcode.Enabled = False
cmbshiftsched.Enabled = False
End Sub

Private Sub DTmovementdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

Me.Top = 0 '(mdiMain.Height - Me.Height) / 8
Me.Left = (mdiMain.Width - Me.Width) / 2

AddFlag = False
EditFlag = False
Frame1.Enabled = False

Call EnableDisableFields(False)
Call EnableDisableButtons(True)

StrQuery = "SELECT * FROM HRShiftSchedule WHERE ShiftStatus = 1 ORDER BY ShiftName"
Set rsShiftSched = ReturnRst(StrQuery, cnnstr, mstatic)

If rsShiftSched.RecordCount > 0 Then
    Showdata rsShiftSched
Else
   Call ClearScreen
   Call EnableDisableButtons(False, True)
End If
Call DisplaySchedules(rsShiftSched)

End Sub

Private Sub EnableDisableButtons(ByVal mFalse As Boolean, _
    Optional mFlag As Boolean = False)

If mFlag = True Then
    cmdadd.Enabled = True
Else
    cmdadd.Enabled = mFalse
End If
cmdEdit.Enabled = mFalse
If AddFlag = True Or EditFlag = True Then
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
Else
    cmdsave.Enabled = False
    cmdcancel.Enabled = False
End If
cmdrefresh.Enabled = mFalse
cmddelete.Enabled = mFalse

End Sub

Private Sub SaveNewShiftCode()
Dim mShiftName As String
Dim mSCode As String
Dim mWorkHours As String
Dim StartTime As String
Dim EndTime As String
Dim HalfTime As String
Dim StartBreak As String
Dim EndBreak As String
Dim GracePeriod As String
Dim UnderPeriod As String
Dim StartOT As String
Dim EndOT As String
Dim StartNight As String
Dim EndNight As String
Dim mHr As Integer, mMin As Integer
Dim res

On Error GoTo err1

mShiftCode = txtShiftCode.Text
mSCode = IIf(IsNull(txtSCode), "<No Code>", txtSCode)

mWorkHours = IIf(IsNull(txtWorkHours), "8.00", txtWorkHours)

StartTime = IIf(IsNull(mskStartTime), "00:00", mskStartTime)
mHr = Val(Left(StartTime, 2))
mMin = Val(Right(StartTime, 2))
If mHr > 23 Or mMin > 59 Then
    MsgBox "Invalid Start Time, should be equal or less than 23 hours and maximum of 59 minutes only!", vbInformation, PayrollAlert
    mskStartTime.SetFocus
    GoTo ExitRtn
End If

HalfTime = IIf(IsNull(mskStartTime), "12:00", mskHalfDay)
mHr = Val(Left(HalfTime, 2))
mMin = Val(Right(HalfTime, 2))
If mHr > 23 Or mMin > 59 Then
    MsgBox "Invalid Half Day Time, should be equal or less than 23 hours and maximum of 59 minutes only!", vbInformation, PayrollAlert
    mskHalfDay.SetFocus
    GoTo ExitRtn
End If

EndTime = IIf(IsNull(mskEndTime), "00:00", mskEndTime)
mHr = Val(Left(EndTime, 2))
mMin = Val(Right(EndTime, 2))
If mHr > 23 Or mMin > 59 Then
    MsgBox "Invalid End Time, should be equal or less than 23 hours and maximum of 59 minutes only!", vbInformation, PayrollAlert
    mskEndTime.SetFocus
    GoTo ExitRtn
End If

'mShiftName = IIf(IsNull(txtShiftName), "<Shift Name>", txtShiftName)
mShiftName = Trim$(mSCode) & " [" & StartTime & " - " & EndTime & "]"
txtShiftName.Text = mShiftName

StartBreak = IIf(IsNull(mskStartBreak), "00:00", mskStartBreak)
mHr = Val(Left(StartBreak, 2))
mMin = Val(Right(StartBreak, 2))
If mHr > 23 Or mMin > 59 Then
    MsgBox "Invalid Start of Break Time, should be equal or less than 23 hours and maximum of 59 minutes only!", vbInformation, PayrollAlert
    GoTo ExitRtn
End If

EndBreak = IIf(IsNull(mskEndBreak), "00:00", mskEndBreak)
mHr = Val(Left(EndBreak, 2))
mMin = Val(Right(EndBreak, 2))
If mHr > 23 Or mMin > 59 Then
    MsgBox "Invalid End of Break Time, should be equal or less than 23 hours and maximum of 59 minutes only!", vbInformation, PayrollAlert
    GoTo ExitRtn
End If

GracePeriod = IIf(IsNull(txtGracePeriod), "0", txtGracePeriod)
UnderPeriod = IIf(IsNull(txtUndertime), "0", txtUndertime)

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

StartNight = IIf(IsNull(mskStartNightDiff), "00:00", mskStartNightDiff)
mHr = Val(Left(StartNight, 2))
mMin = Val(Right(StartNight, 2))
If mHr > 23 Or mMin > 59 Then
    MsgBox "Invalid Start of Night Diff, should be equal or less than 23 hours and maximum of 59 minutes only!", vbInformation, PayrollAlert
    GoTo ExitRtn
End If

EndNight = IIf(IsNull(mskEndNightDiff), "00:00", mskEndNightDiff)
mHr = Val(Left(EndNight, 2))
mMin = Val(Right(EndNight, 2))
If mHr > 23 Or mMin > 59 Then
    MsgBox "Invalid End of Night Diff, should be equal or less than 23 hours and maximum of 59 minutes only!", vbInformation, PayrollAlert
    GoTo ExitRtn
End If

If AddFlag = True Or addflagship = True Then    ' Add New record
StrQuery = "INSERT INTO HRShiftSchedule (ShiftName, ShiftShort, WorkHours, StartTime, MidTime, NoLate, EndTime, NoUDT, NoInOut, " & _
    "StartBreakTime, EndBreakTime, NoBreakTime, GracePeriod, NoGracePeriod, UnderTime, NoUnderTime, " & _
    "StartRegOT, EndRegOT, NoOT, StartNightDiff, EndNightDiff, NoNightDiff, ShiftStatus, " & _
    "FlexiTime, MidNightOut, ShiftDate, UpdatedBy, Remarks) " & _
    " VALUES ('" & mShiftName & "', '" & mSCode & "', " & Val(mWorkHours) & ", '" & StartTime & "', '" & HalfTime & "', " & chkNoLate & ", '" & EndTime & "', " & chkNoUDT & ", " & chkNoInOut & _
    ", '" & StartBreak & "', '" & EndBreak & "', " & chkNoBreaktime & ", '" & GracePeriod & "', " & chkGracePeriod & ", '" & UnderPeriod & "', " & chkUndertime & _
    ", '" & StartOT & "', '" & EndOT & "', " & chkNoOvertime & ", '" & StartNight & "', '" & EndNight & "', " & chkNoNightDiff & ", 1" & _
    ", " & chkFlexible & ", " & chkMidNight & ", '" & Format(Now, "mm/dd/yyyy") & "', '" & LoginName & "', '" & txtRemarks & "')"

Else
StrQuery = "UPDATE HRShiftSchedule SET ShiftName = '" & mShiftName & "', ShiftShort = '" & mSCode & "', WorkHours = " & Val(mWorkHours) & ", StartTime = '" & StartTime & "', MidTime = '" & HalfTime & "', NoLate = " & chkNoLate & ", EndTime = '" & EndTime & "', NoUDT = " & chkNoUDT & ", NoInOut = " & chkNoInOut & _
    ", StartBreakTime = '" & StartBreak & "', EndBreakTime = '" & EndBreak & "', NoBreakTime = " & chkNoBreaktime & ", GracePeriod = '" & GracePeriod & "', NoGracePeriod = " & chkGracePeriod & ", UnderTime = '" & UnderPeriod & "', NoUnderTime = " & chkUndertime & _
    ", StartRegOT = '" & StartOT & "', EndRegOT = '" & EndOT & "', NoOT = " & chkNoOvertime & ", StartNightDiff = '" & StartNight & "', EndNightDiff = '" & EndNight & "', NoNightDiff = " & chkNoNightDiff & _
    ", ShiftDate = '" & Format(Now, "mm/dd/yyyy") & "', FlexiTime = " & chkFlexible & ", MidNightOut = " & chkMidNight & ", UpdatedBy = '" & LoginName & "', Remarks = '" & txtRemarks & "' WHERE ShiftCode = " & mShiftCode
End If

Action_Query StrQuery, cnnstr

rsShiftSched.Requery
If EditFlag = True Then
    rsShiftSched.Find "ShiftCode = " & mShiftCode
End If

If AddFlag = True Or addflagship = True Then
    MsgBox "New Shift Schedule was successfully saved!", vbInformation, PayrollAlert
Else
    MsgBox "Shift Schedule was successfully updated!", vbInformation, PayrollAlert
End If
        
AddFlag = False
EditFlag = False
Frame1.Enabled = False

Call EnableDisableButtons(True)
Call EnableDisableFields(False)

Call DisplaySchedules(rsShiftSched)
Showdata rsShiftSched

ExitRtn:

Exit Sub


err1:

    MsgBox Err.Description, vbInformation, PayrollAlert
    Err.Clear
    
End Sub

Function Getmax() As Integer
    Dim rs As ADODB.Recordset
    Dim str As String
    str = "Select max(assignmentnum) From HREmployeesAssignment"
        Set rs = cnn.Execute(str)
    If Not rs.EOF Then
         Getmax = rs.Fields(0)
    End If
End Function

Private Sub ClearScreen()

txtShiftName.Text = ""
txtSCode.Text = "XX001"
chkFlexible.VALUE = 0
chkMidNight.VALUE = 0
txtWorkHours.Text = "8.00"
chkNoInOut.VALUE = 0
mskStartTime.Text = "00:00"
mskEndTime.Text = "00:00"
mskHalfDay.Text = "12:00"

chkNoLate.VALUE = 0
chkNoUDT.VALUE = 0

chkGracePeriod.VALUE = 1
chkUndertime.VALUE = 1
txtGracePeriod.Text = "10"
txtUndertime.Text = "5"

chkNoBreaktime.VALUE = 1
mskStartBreak.Text = "00:00"
mskEndBreak.Text = "00:00"
chkNoInBreak.VALUE = 0
chkNoOutBreak.VALUE = 0

chkNoOvertime.VALUE = 1
mskStartOT.Text = "00:00"
mskEndOT.Text = "00:00"
chkNoInOT.VALUE = 0
chkNoOutOT.VALUE = 0

chkNoNightDiff.VALUE = 1
mskStartNightDiff.Text = "00:00"
mskEndNightDiff.Text = "00:00"
chkNoInNight.VALUE = 0
chkNoOutNight.VALUE = 0

txtRemarks = ""
txtLastUpdate = Format(Now, "mm/dd/yyyy hh:mm:ss")

Call EnableDisableFields(True)

End Sub

Private Sub EnableDisableFields(ByVal mFalse As Boolean)

'txtShiftName.Enabled = mFalse
txtSCode.Enabled = mFalse
txtWorkHours.Enabled = mFalse
chkNoInOut.Enabled = mFalse
chkFlexible.Enabled = mFalse
chkMidNight.Enabled = mFalse
mskStartTime.Enabled = mFalse
mskHalfDay.Enabled = mFalse

mskEndTime.Enabled = mFalse
chkNoLate.Enabled = mFalse
chkNoUDT.Enabled = mFalse

txtGracePeriod.Enabled = mFalse
txtUndertime.Enabled = mFalse
chkUndertime.Enabled = mFalse
chkGracePeriod.Enabled = mFalse

chkNoBreaktime.Enabled = mFalse
mskStartBreak.Enabled = mFalse
mskEndBreak.Enabled = mFalse
chkNoInBreak.Enabled = mFalse
chkNoOutBreak.Enabled = mFalse

chkNoOvertime.Enabled = mFalse
mskStartOT.Enabled = mFalse
mskEndOT.Enabled = mFalse
chkNoInOT.Enabled = mFalse
chkNoOutOT.Enabled = mFalse

chkNoNightDiff.Enabled = mFalse
mskStartNightDiff.Enabled = mFalse
mskEndNightDiff.Enabled = mFalse
chkNoInNight.Enabled = mFalse
chkNoOutNight.Enabled = mFalse

txtRemarks.Enabled = mFalse
'txtLastUpdate = Format(Now, "mm/dd/yyyy hh:mm:ss AM")

End Sub

Private Sub Showdata(ByRef rec As ADODB.Recordset)
Dim rs As ADODB.Recordset

If rec.EOF = True Then
    Call ClearScreen
    Call EnableDisableButtons(False, True)
Else
    txtShiftCode.Text = rec!ShiftCode
    txtSCode.Text = IIf(IsNull(rec!ShiftShort), "<No Code>", Trim$(rec!ShiftShort))
    txtShiftName.Text = IIf(IsNull(rec!ShiftName), "<No Shift>", Trim$(rec!ShiftName))
    txtWorkHours.Text = IIf(IsNull(rec!WorkHours), "8.00", rec!WorkHours)
    chkFlexible.VALUE = IIf(rec!FlexiTime = True, 1, 0)
    chkNoInOut.VALUE = IIf(rec!NoInOut = True, 1, 0)
    mskStartTime.Text = IIf(IsNull(rec!StartTime), "00:00", Format(rec!StartTime, "hh:mm"))
    mskHalfDay.Text = IIf(IsNull(rec!MidTime), "00:00", Format(rec!MidTime, "hh:mm"))
    mskEndTime.Text = IIf(IsNull(rec!EndTime), "00:00", Format(rec!EndTime, "hh:mm"))
    chkNoLate.VALUE = IIf(rec!Nolate = True, 1, 0)
    chkNoUDT.VALUE = IIf(rec!NoUDT = True, 1, 0)
    chkMidNight.VALUE = IIf(rec!MidNightOut = True, 1, 0)
    
    txtGracePeriod.Text = IIf(IsNull(rec!GracePeriod), "0", rec!GracePeriod)
    txtUndertime.Text = IIf(IsNull(rec!Undertime), "0", rec!Undertime)
    
    chkGracePeriod.VALUE = IIf(rec!NoGracePeriod = True, 1, 0)
    chkUndertime.VALUE = IIf(rec!NoUndertime = True, 1, 0)
    
    chkNoBreaktime.VALUE = IIf(rec!NoBreakTime = True, 1, 0)
    mskStartBreak.Text = IIf(IsNull(rec!StartBreakTime), "00:00", Format(rec!StartBreakTime, "hh:mm"))
    mskEndBreak.Text = IIf(IsNull(rec!EndBreakTime), "00:00", Format(rec!EndBreakTime, "hh:mm"))
    'chkNoInBreak.VALUE = IIf(rec!NeedIn = True, 1, 0)
    'chkNoOutBreak.VALUE = IIf(rec!NeedOut = True, 1, 0)
    
    chkNoOvertime.VALUE = IIf(rec!NoOT = True, 1, 0)
    mskStartOT.Text = IIf(IsNull(rec!StartRegOT), "00:00", Format(rec!StartRegOT, "hh:mm"))
    mskEndOT.Text = IIf(IsNull(rec!EndRegOT), "23:59", Format(rec!EndRegOT, "hh:mm"))
    'chkNoInOT.VALUE = IIf(rec!NeedIn = True, 1, 0)
    'chkNoOutOT.VALUE = IIf(rec!NeedOut = True, 1, 0)
    
    chkNoNightDiff.VALUE = IIf(rec!NoNightDiff = True, 1, 0)
    mskStartNightDiff.Text = IIf(IsNull(rec!StartNightDiff), "00:00", Format(rec!StartNightDiff, "hh:mm"))
    mskEndNightDiff.Text = IIf(IsNull(rec!EndNightDiff), "06:00", Format(rec!EndNightDiff, "hh:mm"))
    'chkNoInOT.VALUE = IIf(rec!NeedIn = True, 1, 0)
    'chkNoOutOT.VALUE = IIf(rec!NeedOut = True, 1, 0)
    
    txtRemarks.Text = IIf(IsNull(rec!Remarks), "", rec!Remarks)
    txtLastUpdate.Text = IIf(IsNull(rec!ShiftDate), "__/__/____", Format(rec!ShiftDate, "mm/dd/yyyy hh:mm:ss"))
    txtUpdatedBy.Text = IIf(IsNull(rec!Updatedby), "", rec!Updatedby)
    
End If

End Sub

Private Sub txtremarks_GotFocus()
txtRemarks.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtRemarks
End Sub

Private Sub txtremarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtremarks_LostFocus()
txtRemarks.BackColor = vbWhite
End Sub

Private Sub txtUndertime_GotFocus()
txtUndertime.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtUndertime
End Sub

'Private Sub txtUndertime_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub

Private Sub txtUndertime_LostFocus()
txtUndertime.BackColor = vbWhite
End Sub

Private Sub DisplaySchedules(ByRef rsshift As ADODB.Recordset)
Dim litem2 As ListItem
Dim Dtrctr As Integer
Dim WithDTR As Integer
Dim Ttime As Double
Dim TDays As Double
Dim TAbsent As Double
Dim TLates As Double
Dim TUndertime As Double
Dim TOvertime As Double

If rsshift.RecordCount = 0 Then Exit Sub
    
With rsshift
    .MoveFirst
    lstSchedule.ListItems.Clear
    While Not (rsshift.EOF Or rsshift.BOF)
        Dtrctr = Dtrctr + 1
        Set litem2 = lstSchedule.ListItems.Add(, , Dtrctr)
        litem2.SubItems(1) = IIf(IsNull(rsshift!ShiftName), "NO Shift Name", Trim$(rsshift!ShiftName))
        litem2.SubItems(2) = IIf(IsNull(rsshift!WorkHours), "No Hours/Day", Trim$(rsshift!WorkHours))
        litem2.SubItems(3) = IIf(IsNull(rsshift!StartTime), "NO Start Time", Trim$(rsshift!StartTime))
        litem2.SubItems(4) = IIf(IsNull(rsshift!EndTime), "NO End Time", Trim$(rsshift!EndTime))
        litem2.SubItems(5) = IIf(IsNull(rsshift!MidTime), "NO Mid-Time", Trim$(rsshift!MidTime))
        'litem2.SubItems(6) = IIf(rsShift!Nolate = True, "No Late", "With Late")
        'litem2.SubItems(7) = IIf(rsShift!NoUDT = True, "No UDT", "With UDT")
        litem2.SubItems(6) = IIf(IsNull(rsshift!Remarks), "", Trim$(rsshift!Remarks))
        rsshift.MoveNext
    Wend
End With

End Sub

'Private Sub lstSchedule_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
'    Showdata rsShiftSched
'End Sub

Private Sub lstSchedule_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtShiftName = Item.ListSubItems(1).Text
    rsShiftSched.MoveFirst
    rsShiftSched.Find "ShiftName = '" & txtShiftName & "'"
    Showdata rsShiftSched
End Sub

Private Sub lstSchedule_Click()
    Showdata rsShiftSched
End Sub

