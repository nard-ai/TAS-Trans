VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTimeKeep 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Time Keep"
   ClientHeight    =   8295
   ClientLeft      =   3750
   ClientTop       =   3540
   ClientWidth     =   12270
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmTimeKeep.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   12270
   Begin TabDlg.SSTab SSTab1 
      Height          =   7725
      Index           =   1
      Left            =   90
      TabIndex        =   36
      Top             =   540
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   13626
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
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
      TabCaption(0)   =   "Employee Time Keeping"
      TabPicture(0)   =   "frmTimeKeep.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdNavigation(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdNavigation(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdNavigation(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdNavigation(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   3
         Left            =   10725
         Picture         =   "frmTimeKeep.frx":05A6
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   8430
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   2
         Left            =   10440
         Picture         =   "frmTimeKeep.frx":0AD8
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "First"
         Top             =   8430
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   1
         Left            =   10155
         Picture         =   "frmTimeKeep.frx":0BD2
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "First"
         Top             =   8430
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   0
         Left            =   9870
         Picture         =   "frmTimeKeep.frx":0CCC
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "First"
         Top             =   8430
         Width           =   300
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   7305
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   11895
         Begin VB.TextBox mText1 
            Height          =   285
            Left            =   0
            TabIndex        =   230
            Top             =   0
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Frame Frame6 
            Caption         =   "Change Cutoff Date"
            ForeColor       =   &H00800000&
            Height          =   1815
            Left            =   2880
            TabIndex        =   161
            Top             =   960
            Visible         =   0   'False
            Width           =   8895
            Begin VB.CommandButton cmdSaveCutoff 
               Height          =   405
               Left            =   5520
               Picture         =   "frmTimeKeep.frx":11FE
               Style           =   1  'Graphical
               TabIndex        =   164
               ToolTipText     =   "Save Cutoff"
               Top             =   480
               Width           =   465
            End
            Begin VB.CommandButton cmdCancelCutoff 
               Height          =   405
               Left            =   6360
               Picture         =   "frmTimeKeep.frx":1730
               Style           =   1  'Graphical
               TabIndex        =   163
               ToolTipText     =   "Cancel"
               Top             =   480
               Width           =   465
            End
            Begin VB.CheckBox chkInitCutoff 
               Caption         =   "Use initial Cutoff"
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   5400
               TabIndex        =   162
               Top             =   1200
               Width           =   1935
            End
            Begin MSComCtl2.DTPicker DTStartFirstCutoff 
               Height          =   375
               Left            =   1920
               TabIndex        =   165
               Top             =   315
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Format          =   143785985
               CurrentDate     =   40262
            End
            Begin MSComCtl2.DTPicker DTEndFirstCutoff 
               Height          =   375
               Left            =   3720
               TabIndex        =   166
               Top             =   315
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Format          =   143785985
               CurrentDate     =   40262
            End
            Begin MSComCtl2.DTPicker DTStartSecondCutoff 
               Height          =   375
               Left            =   1920
               TabIndex        =   167
               Top             =   720
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Format          =   143785985
               CurrentDate     =   40262
            End
            Begin MSComCtl2.DTPicker DTEndSecondCutoff 
               Height          =   375
               Left            =   3720
               TabIndex        =   168
               Top             =   720
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Format          =   143785985
               CurrentDate     =   40262
            End
            Begin MSComCtl2.DTPicker DTIniStartCutoff 
               Height          =   375
               Left            =   1920
               TabIndex        =   169
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Format          =   143785985
               CurrentDate     =   40262
            End
            Begin MSComCtl2.DTPicker DTIniEndCutoff 
               Height          =   375
               Left            =   3720
               TabIndex        =   170
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Format          =   143785985
               CurrentDate     =   40262
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Second Cutoff from:"
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
               Index           =   49
               Left            =   240
               TabIndex        =   176
               Top             =   840
               Width           =   1635
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "First Cutoff from:"
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
               Index           =   47
               Left            =   240
               TabIndex        =   175
               Top             =   465
               Width           =   1275
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "To:"
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
               Index           =   50
               Left            =   3360
               TabIndex        =   174
               Top             =   480
               Width           =   315
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "To:"
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
               Index           =   51
               Left            =   3360
               TabIndex        =   173
               Top             =   840
               Width           =   315
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "To:"
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
               Index           =   66
               Left            =   3360
               TabIndex        =   172
               Top             =   1320
               Width           =   315
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Set Initial Cutoff"
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
               Index           =   67
               Left            =   240
               TabIndex        =   171
               Top             =   1320
               Width           =   1635
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   2880
            TabIndex        =   212
            Top             =   1920
            Width           =   8895
            Begin VB.CheckBox chkAllDTR 
               BackColor       =   &H00C0FFFF&
               Caption         =   "All"
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   5400
               TabIndex        =   222
               ToolTipText     =   "Preview All DTR Records"
               Top             =   150
               Width           =   495
            End
            Begin VB.CommandButton cmdPreview 
               Caption         =   "Preview Process DTR Records"
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
               Left            =   2760
               TabIndex        =   221
               Top             =   120
               Width           =   2535
            End
            Begin VB.CommandButton cmdLeaveApplication 
               BackColor       =   &H00FF0000&
               Caption         =   "Print for Payroll"
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
               Left            =   7320
               MaskColor       =   &H00FF0000&
               TabIndex        =   216
               ToolTipText     =   "Leave/Dayoff Application Form"
               Top             =   120
               Width           =   1455
            End
            Begin VB.CommandButton cmdReprocess 
               Caption         =   "Re- Process DTR Records"
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
               Left            =   120
               TabIndex        =   215
               Top             =   120
               Width           =   2535
            End
            Begin VB.CommandButton cmdDTRPayroll 
               Caption         =   "Import DTR for Payroll Processing"
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
               Left            =   120
               TabIndex        =   214
               Top             =   480
               Width           =   8655
            End
            Begin VB.CommandButton cmdDTRManual 
               Caption         =   "Save to Excel"
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
               Left            =   5880
               TabIndex        =   213
               ToolTipText     =   "Manual DTR Entry or Updates"
               Top             =   120
               Width           =   1455
            End
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   1080
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   2778
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
         Begin VB.TextBox txtCancelSw 
            Height          =   285
            Left            =   2160
            TabIndex        =   147
            Text            =   "0"
            Top             =   120
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Other Deductions"
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
            TabIndex        =   108
            Top             =   8040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtClient 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Index           =   7
            Left            =   7740
            TabIndex        =   102
            Text            =   "Client Name"
            Top             =   7815
            Width           =   1125
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Cal/SSS Loans"
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
            Left            =   3360
            TabIndex        =   101
            Top             =   8040
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Ecola/SSS Deductions"
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
            Left            =   4800
            TabIndex        =   100
            Top             =   8040
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.TextBox txtEmpNum 
            Height          =   285
            Left            =   990
            TabIndex        =   75
            Text            =   "txtclientcode"
            Top             =   180
            Visible         =   0   'False
            Width           =   645
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
            Index           =   5
            Left            =   2160
            TabIndex        =   29
            Text            =   "Client Name"
            Top             =   8040
            Width           =   6615
         End
         Begin VB.CheckBox chkRates 
            Caption         =   "Client with Compensation Package"
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
            Index           =   7
            Left            =   1920
            TabIndex        =   28
            Top             =   8040
            Value           =   1  'Checked
            Width           =   2985
         End
         Begin VB.TextBox txtClientCode 
            Height          =   285
            Left            =   1800
            TabIndex        =   66
            Text            =   "txtclientcode"
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkRates 
            Caption         =   "Client with Service Contract"
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
            Index           =   6
            Left            =   360
            TabIndex        =   27
            Top             =   8040
            Value           =   1  'Checked
            Width           =   2535
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
            Index           =   6
            Left            =   2640
            TabIndex        =   30
            Text            =   "Client Name"
            Top             =   7920
            Width           =   1935
         End
         Begin VB.TextBox txtClient 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Index           =   4
            Left            =   9870
            TabIndex        =   34
            Text            =   "Client Name"
            Top             =   7815
            Width           =   1005
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
            Index           =   3
            Left            =   4320
            TabIndex        =   31
            Text            =   "Client Name"
            Top             =   8040
            Width           =   1365
         End
         Begin VB.Frame Frame1 
            Height          =   1455
            Left            =   12480
            TabIndex        =   44
            Top             =   3720
            Visible         =   0   'False
            Width           =   5415
            Begin VB.TextBox txtRates 
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
               Index           =   0
               Left            =   1800
               TabIndex        =   132
               Text            =   "15"
               Top             =   360
               Width           =   1005
            End
            Begin VB.TextBox txtRates 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
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
               Index           =   3
               Left            =   4320
               TabIndex        =   14
               Text            =   "15"
               Top             =   1035
               Width           =   825
            End
            Begin VB.TextBox txtRates 
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
               Index           =   2
               Left            =   4140
               TabIndex        =   10
               Text            =   "15"
               Top             =   405
               Width           =   1005
            End
            Begin VB.TextBox txtRates 
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
               Left            =   2950
               TabIndex        =   9
               Text            =   "15"
               Top             =   405
               Width           =   1005
            End
            Begin VB.ComboBox cmbPayBasis 
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
               ItemData        =   "frmTimeKeep.frx":1C62
               Left            =   180
               List            =   "frmTimeKeep.frx":1C72
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   405
               Width           =   1485
            End
            Begin VB.ComboBox cmbPaySystem 
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
               ItemData        =   "frmTimeKeep.frx":1C9A
               Left            =   180
               List            =   "frmTimeKeep.frx":1CAD
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   1035
               Width           =   1485
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
               Left            =   1800
               TabIndex        =   12
               Top             =   1035
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
            Begin MSMask.MaskEdBox mskEndContract 
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
               Left            =   3060
               TabIndex        =   13
               Top             =   1035
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
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Start Contract:"
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
               Left            =   1800
               TabIndex        =   52
               Top             =   765
               Width           =   1275
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Payment Basis:"
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
               Left            =   180
               TabIndex        =   51
               Top             =   180
               Width           =   1455
            End
            Begin VB.Label lblclient 
               Alignment       =   2  'Center
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   41
               Left            =   1680
               TabIndex        =   50
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Daily Rate"
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
               Index           =   42
               Left            =   3090
               TabIndex        =   49
               Top             =   180
               Width           =   915
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Hourly Rate"
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
               Index           =   43
               Left            =   4140
               TabIndex        =   48
               Top             =   180
               Width           =   1035
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Payment System:"
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
               Index           =   44
               Left            =   180
               TabIndex        =   47
               Top             =   765
               Width           =   1455
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "End Contract:"
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
               Left            =   3060
               TabIndex        =   46
               Top             =   765
               Width           =   1275
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Month No."
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
               Index           =   45
               Left            =   4320
               TabIndex        =   45
               Top             =   765
               Width           =   915
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
            ItemData        =   "frmTimeKeep.frx":1CCE
            Left            =   180
            List            =   "frmTimeKeep.frx":1CDB
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   450
            Width           =   2565
         End
         Begin VB.Frame Frame5 
            Height          =   1695
            Index           =   1
            Left            =   12480
            TabIndex        =   80
            Top             =   2040
            Visible         =   0   'False
            Width           =   5415
            Begin VB.CheckBox chkRates 
               Caption         =   "SSS Ded is Monthly"
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
               Index           =   10
               Left            =   3620
               TabIndex        =   106
               Top             =   1320
               Width           =   1725
            End
            Begin VB.TextBox txtRates 
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
               Index           =   10
               Left            =   2790
               TabIndex        =   85
               Text            =   "0.00"
               Top             =   1350
               Width           =   735
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "Include ECOLA in ROT Basis"
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
               Index           =   9
               Left            =   2730
               TabIndex        =   105
               Top             =   1050
               Width           =   2565
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "No Tax OT,Ecola,Vale,Ince..."
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
               Index           =   8
               Left            =   90
               TabIndex        =   104
               Top             =   1050
               Width           =   2565
            End
            Begin VB.TextBox txtRates 
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
               Index           =   4
               Left            =   1830
               TabIndex        =   94
               Text            =   "50.00"
               Top             =   140
               Width           =   735
            End
            Begin VB.TextBox txtRates 
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
               Index           =   5
               Left            =   1830
               TabIndex        =   93
               Text            =   "20.83"
               Top             =   420
               Width           =   735
            End
            Begin VB.TextBox txtRates 
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
               Index           =   6
               Left            =   1035
               TabIndex        =   86
               Text            =   "0.00"
               Top             =   1350
               Width           =   735
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "W/ 13th Month Pay"
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
               Index           =   2
               Left            =   90
               TabIndex        =   97
               Top             =   450
               Value           =   1  'Checked
               Width           =   1725
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "ECOLA Payment"
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
               Index           =   1
               Left            =   90
               TabIndex        =   96
               Top             =   150
               Value           =   1  'Checked
               Width           =   1725
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "W/ Tax Deductions"
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
               Index           =   0
               Left            =   90
               TabIndex        =   95
               Top             =   750
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.TextBox txtRates 
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
               Index           =   7
               Left            =   4530
               TabIndex        =   92
               Text            =   "0.00"
               Top             =   140
               Width           =   735
            End
            Begin VB.TextBox txtRates 
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
               Index           =   8
               Left            =   4530
               TabIndex        =   91
               Text            =   "0.00"
               Top             =   420
               Width           =   735
            End
            Begin VB.TextBox txtRates 
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
               Index           =   9
               Left            =   4530
               TabIndex        =   90
               Text            =   "0.00"
               Top             =   710
               Width           =   735
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "W/ SSS Deductions"
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
               Index           =   3
               Left            =   2730
               TabIndex        =   89
               Top             =   150
               Value           =   1  'Checked
               Width           =   1725
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "W/ PHealth Dedtns"
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
               Index           =   4
               Left            =   2730
               TabIndex        =   88
               Top             =   450
               Value           =   1  'Checked
               Width           =   1725
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "Pag-ibig Deductions"
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
               Index           =   5
               Left            =   2730
               TabIndex        =   87
               Top             =   750
               Value           =   1  'Checked
               Width           =   1725
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Uniform:"
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
               Left            =   120
               TabIndex        =   99
               Top             =   1395
               Width           =   915
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "TAS/Aesla:"
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
               Left            =   1935
               TabIndex        =   98
               Top             =   1395
               Width           =   855
            End
         End
         Begin VB.Frame Frame5 
            Height          =   1695
            Index           =   2
            Left            =   12480
            TabIndex        =   109
            Top             =   120
            Visible         =   0   'False
            Width           =   5655
            Begin VB.CommandButton cmdDelDed 
               Caption         =   "Del"
               Height          =   255
               Left            =   4680
               TabIndex        =   130
               Top             =   1320
               Width           =   615
            End
            Begin VB.CommandButton cmdEditDed 
               Caption         =   "Edit"
               Height          =   255
               Left            =   4080
               TabIndex        =   129
               Top             =   1320
               Width           =   615
            End
            Begin VB.CommandButton cmdAddDed 
               Caption         =   "Add"
               Height          =   255
               Left            =   3360
               TabIndex        =   128
               Top             =   1320
               Width           =   735
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "Count Payments"
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
               Height          =   195
               Index           =   12
               Left            =   3480
               TabIndex        =   127
               Top             =   1200
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.TextBox txtRates 
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
               Index           =   21
               Left            =   2400
               TabIndex        =   126
               Text            =   "1"
               Top             =   1300
               Width           =   795
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "Active Deduction"
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
               Height          =   195
               Index           =   11
               Left            =   3720
               TabIndex        =   118
               Top             =   285
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.ComboBox cmbTypes 
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
               ItemData        =   "frmTimeKeep.frx":1CFE
               Left            =   1470
               List            =   "frmTimeKeep.frx":1D0B
               Style           =   2  'Dropdown List
               TabIndex        =   117
               Top             =   180
               Width           =   2205
            End
            Begin VB.TextBox txtRates 
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
               Index           =   18
               Left            =   1350
               TabIndex        =   121
               Text            =   "0.00"
               Top             =   720
               Width           =   915
            End
            Begin VB.TextBox txtRates 
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
               Index           =   17
               Left            =   150
               TabIndex        =   119
               Text            =   "0.00"
               Top             =   720
               Width           =   915
            End
            Begin VB.TextBox txtRates 
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
               Index           =   16
               Left            =   1350
               TabIndex        =   125
               Text            =   "0.00"
               Top             =   1300
               Width           =   915
            End
            Begin VB.TextBox txtRates 
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
               Index           =   15
               Left            =   120
               TabIndex        =   124
               Text            =   "0.00"
               Top             =   1300
               Width           =   915
            End
            Begin MSMask.MaskEdBox MaskEdBox5 
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
               Left            =   2610
               TabIndex        =   122
               Top             =   705
               Width           =   1110
               _ExtentX        =   1958
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
            Begin MSMask.MaskEdBox MaskEdBox6 
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
               TabIndex        =   123
               Top             =   705
               Width           =   1110
               _ExtentX        =   1958
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
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Payments"
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
               Index           =   36
               Left            =   3360
               TabIndex        =   131
               Top             =   1080
               Width           =   1785
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "No of Pays"
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
               Index           =   31
               Left            =   2400
               TabIndex        =   120
               Top             =   1125
               Width           =   1035
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Payments"
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
               Index           =   30
               Left            =   1410
               TabIndex        =   115
               Top             =   540
               Width           =   915
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Amt"
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
               Index           =   29
               Left            =   240
               TabIndex        =   114
               Top             =   540
               Width           =   825
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
               Index           =   28
               Left            =   4110
               TabIndex        =   113
               Top             =   525
               Width           =   735
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
               Index           =   27
               Left            =   2730
               TabIndex        =   112
               Top             =   525
               Width           =   825
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Overall Total"
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
               Index           =   26
               Left            =   1380
               TabIndex        =   111
               Top             =   1125
               Width           =   1035
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Interest"
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
               Index           =   25
               Left            =   330
               TabIndex        =   110
               Top             =   1125
               Width           =   825
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Deduction Types"
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
               Index           =   48
               Left            =   150
               TabIndex        =   116
               Top             =   240
               Width           =   1545
            End
         End
         Begin VB.Frame Frame5 
            Height          =   1695
            Index           =   0
            Left            =   12480
            TabIndex        =   61
            Top             =   1440
            Visible         =   0   'False
            Width           =   5415
            Begin VB.TextBox txtRates 
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
               Index           =   20
               Left            =   120
               TabIndex        =   23
               Text            =   "0.00"
               Top             =   1300
               Width           =   915
            End
            Begin VB.TextBox txtRates 
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
               Index           =   19
               Left            =   1350
               TabIndex        =   24
               Text            =   "0.00"
               Top             =   1300
               Width           =   915
            End
            Begin VB.TextBox txtRates 
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
               Index           =   14
               Left            =   150
               TabIndex        =   19
               Text            =   "0.00"
               Top             =   840
               Width           =   915
            End
            Begin VB.TextBox txtRates 
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
               Index           =   13
               Left            =   1350
               TabIndex        =   20
               Text            =   "0.00"
               Top             =   840
               Width           =   915
            End
            Begin VB.TextBox txtRates 
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
               Index           =   12
               Left            =   1350
               TabIndex        =   16
               Text            =   "0.00"
               Top             =   360
               Width           =   915
            End
            Begin VB.TextBox txtRates 
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
               Index           =   11
               Left            =   150
               TabIndex        =   15
               Text            =   "0.00"
               Top             =   360
               Width           =   915
            End
            Begin MSMask.MaskEdBox mskStartSSS 
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
               Left            =   2610
               TabIndex        =   17
               Top             =   360
               Width           =   1110
               _ExtentX        =   1958
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
            Begin MSMask.MaskEdBox mskEndSSS 
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
               Left            =   3930
               TabIndex        =   18
               Top             =   360
               Width           =   1110
               _ExtentX        =   1958
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
            Begin MSMask.MaskEdBox mskStartPbig 
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
               Left            =   2610
               TabIndex        =   21
               Top             =   840
               Width           =   1110
               _ExtentX        =   1958
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
            Begin MSMask.MaskEdBox mskEndPbig 
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
               Left            =   3930
               TabIndex        =   22
               Top             =   840
               Width           =   1110
               _ExtentX        =   1958
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
            Begin MSMask.MaskEdBox mskStartVale 
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
               Left            =   2610
               TabIndex        =   25
               Top             =   1305
               Width           =   1110
               _ExtentX        =   1958
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
            Begin MSMask.MaskEdBox mskEndVale 
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
               TabIndex        =   26
               Top             =   1305
               Width           =   1110
               _ExtentX        =   1958
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
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Tot Calamity"
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
               Index           =   35
               Left            =   210
               TabIndex        =   84
               Top             =   1125
               Width           =   945
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Cal Payment"
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
               Index           =   34
               Left            =   1380
               TabIndex        =   83
               Top             =   1125
               Width           =   1035
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
               Index           =   33
               Left            =   2730
               TabIndex        =   82
               Top             =   1125
               Width           =   825
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
               Index           =   32
               Left            =   4110
               TabIndex        =   81
               Top             =   1125
               Width           =   735
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "PBig Loan"
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
               Index           =   22
               Left            =   240
               TabIndex        =   74
               Top             =   660
               Width           =   825
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "PBig Payment"
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
               Index           =   21
               Left            =   1410
               TabIndex        =   73
               Top             =   660
               Width           =   1035
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
               Index           =   20
               Left            =   2730
               TabIndex        =   72
               Top             =   660
               Width           =   825
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
               Index           =   19
               Left            =   4110
               TabIndex        =   71
               Top             =   660
               Width           =   735
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
               Index           =   18
               Left            =   4110
               TabIndex        =   70
               Top             =   180
               Width           =   735
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
               Index           =   16
               Left            =   2730
               TabIndex        =   69
               Top             =   180
               Width           =   825
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "SSS Payment"
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
               Index           =   14
               Left            =   1350
               TabIndex        =   68
               Top             =   135
               Width           =   1035
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "SSS Loan"
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
               Left            =   300
               TabIndex        =   67
               Top             =   165
               Width           =   825
            End
         End
         Begin VB.Frame Frame4 
            Height          =   4680
            Left            =   0
            TabIndex        =   133
            Top             =   2640
            Width           =   11895
            Begin VB.Frame Frame8 
               BorderStyle     =   0  'None
               Height          =   735
               Left            =   3000
               TabIndex        =   155
               Top             =   1800
               Visible         =   0   'False
               Width           =   5415
               Begin VB.Label lblclient 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Display"
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
                  Height          =   375
                  Index           =   65
                  Left            =   240
                  TabIndex        =   156
                  Top             =   240
                  Width           =   5025
               End
            End
            Begin VB.CommandButton cmdMarkActive 
               Caption         =   "Mark Active"
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
               Left            =   3960
               TabIndex        =   136
               Top             =   5760
               Width           =   1095
            End
            Begin VB.TextBox txtEmpName 
               Appearance      =   0  'Flat
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
               Left            =   2400
               TabIndex        =   135
               Text            =   "EmpName"
               Top             =   5460
               Width           =   4275
            End
            Begin VB.TextBox txtEmpNo 
               Appearance      =   0  'Flat
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
               Left            =   420
               TabIndex        =   134
               Text            =   "EmpNo"
               Top             =   5460
               Width           =   1305
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   855
               Index           =   1
               Left            =   3720
               TabIndex        =   137
               Top             =   2760
               Visible         =   0   'False
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   1508
               Sorted          =   -1  'True
               LabelWrap       =   0   'False
               HideSelection   =   0   'False
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
               NumItems        =   9
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "No"
                  Object.Width           =   176
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Day"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Date"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Time In"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Time Out"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Absent"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Lates"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Undertime"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "OT"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSComctlLib.ListView lstDTRDisplay 
               Height          =   4050
               Left            =   120
               TabIndex        =   150
               Top             =   240
               Width           =   11640
               _ExtentX        =   20532
               _ExtentY        =   7144
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
               NumItems        =   14
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "No"
                  Object.Width           =   706
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Days"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "DTR Date"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "DTR ID"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Hours"
                  Object.Width           =   1147
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Time In"
                  Object.Width           =   1586
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Time Out"
                  Object.Width           =   1587
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Worked"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "Absent"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Text            =   "Lates"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   10
                  Text            =   "UDT"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   11
                  Text            =   "OT"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   12
                  Text            =   "NDiff"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   13
                  Text            =   "Remarks"
                  Object.Width           =   2205
               EndProperty
            End
            Begin MSComctlLib.ListView lstCon 
               Height          =   4050
               Left            =   120
               TabIndex        =   223
               Top             =   240
               Width           =   11640
               _ExtentX        =   20532
               _ExtentY        =   7144
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
               NumItems        =   30
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "No"
                  Object.Width           =   706
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Days"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "DTR Date"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Hours"
                  Object.Width           =   1147
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "1st Trip IN"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "1st Trip OUT"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "1st Trip HR"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "2nd Trip IN"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "2nd Trip OUT"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Text            =   "2nd Trip HR"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   10
                  Text            =   "3rd Trip IN"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   11
                  Text            =   "3rd Trip OUT"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   12
                  Text            =   "3rd Trip HR"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   13
                  Text            =   "4th Trip IN"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   14
                  Text            =   "4th Trip OUT"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   15
                  Text            =   "4th Trip HR"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   16
                  Text            =   "5th Trip IN"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   17
                  Text            =   "5th Trip OUT"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   18
                  Text            =   "5th Trip HR"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   19
                  Text            =   "6th Trip IN"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   20
                  Text            =   "6th Trip OUT"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   21
                  Text            =   "6th Trip HR"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   22
                  Text            =   "Worked Hr"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   23
                  Text            =   "Absent"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   24
                  Text            =   "Lates"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   25
                  Text            =   "UDT"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   26
                  Text            =   "OT"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   27
                  Text            =   "NDiff"
                  Object.Width           =   1235
               EndProperty
               BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   28
                  Text            =   "Remarks"
                  Object.Width           =   2205
               EndProperty
               BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   29
                  Object.Width           =   12
               EndProperty
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "RD:"
               Height          =   255
               Index           =   9
               Left            =   8760
               TabIndex        =   217
               Top             =   4365
               Width           =   855
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "VL:"
               Height          =   255
               Index           =   8
               Left            =   2760
               TabIndex        =   160
               Top             =   4365
               Width           =   615
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "SL:"
               Height          =   255
               Index           =   7
               Left            =   2040
               TabIndex        =   159
               Top             =   4365
               Width           =   615
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "NDiff:"
               Height          =   255
               Index           =   6
               Left            =   7680
               TabIndex        =   154
               Top             =   4365
               Width           =   855
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "OT:"
               Height          =   255
               Index           =   5
               Left            =   6720
               TabIndex        =   146
               Top             =   4365
               Width           =   855
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "UDT:"
               Height          =   255
               Index           =   4
               Left            =   5760
               TabIndex        =   145
               Top             =   4365
               Width           =   855
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "Lates:"
               Height          =   255
               Index           =   3
               Left            =   4680
               TabIndex        =   144
               Top             =   4365
               Width           =   975
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "Absent:"
               Height          =   255
               Index           =   2
               Left            =   3480
               TabIndex        =   143
               Top             =   4365
               Width           =   975
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "Pay Hrs:"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   142
               Top             =   4365
               Width           =   1095
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "Days:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   141
               Top             =   4365
               Width           =   735
            End
            Begin VB.Label lblclient 
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
               ForeColor       =   &H00800000&
               Height          =   165
               Index           =   15
               Left            =   1860
               TabIndex        =   138
               Top             =   5550
               Width           =   915
            End
         End
         Begin VB.Frame Frame7 
            ForeColor       =   &H00800000&
            Height          =   1335
            Left            =   2880
            TabIndex        =   177
            Top             =   960
            Visible         =   0   'False
            Width           =   8895
            Begin VB.CheckBox chkAllRec 
               Caption         =   "All Records"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   7320
               TabIndex        =   227
               Top             =   960
               Width           =   1335
            End
            Begin VB.CommandButton cmdTrips 
               Caption         =   "Next >>>"
               Height          =   405
               Index           =   1
               Left            =   2760
               Style           =   1  'Graphical
               TabIndex        =   225
               Top             =   840
               Width           =   1185
            End
            Begin VB.CommandButton cmdTrips 
               Caption         =   "<<<Previous"
               Height          =   405
               Index           =   0
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   224
               Top             =   840
               Width           =   1185
            End
            Begin VB.CheckBox chkAllNoDTR 
               Caption         =   "All blank DTR"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   7320
               TabIndex        =   190
               Top             =   720
               Width           =   1335
            End
            Begin VB.CommandButton cmdSaveDTR 
               Height          =   405
               Left            =   7680
               Picture         =   "frmTimeKeep.frx":1D2E
               Style           =   1  'Graphical
               TabIndex        =   189
               ToolTipText     =   "Save DTR changes"
               Top             =   240
               Width           =   465
            End
            Begin VB.TextBox txtDTR 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   55
               Left            =   480
               TabIndex        =   185
               Text            =   "00:00"
               Top             =   1320
               Visible         =   0   'False
               Width           =   885
            End
            Begin VB.TextBox txtDTR 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   56
               Left            =   1320
               TabIndex        =   183
               Text            =   "00:00"
               Top             =   1320
               Visible         =   0   'False
               Width           =   885
            End
            Begin VB.TextBox txtDTR 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   2
               Left            =   2880
               TabIndex        =   181
               Text            =   "0.00"
               Top             =   480
               Width           =   645
            End
            Begin VB.TextBox txtDTR 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   3
               Left            =   3480
               TabIndex        =   180
               Text            =   "0.00"
               Top             =   480
               Width           =   645
            End
            Begin VB.TextBox txtDTR 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   4
               Left            =   4080
               TabIndex        =   182
               Text            =   "0.00"
               Top             =   480
               Width           =   645
            End
            Begin VB.TextBox txtDTR 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   5
               Left            =   4680
               TabIndex        =   184
               Text            =   "0.00"
               Top             =   480
               Width           =   645
            End
            Begin VB.TextBox txtDTR 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   6
               Left            =   5280
               TabIndex        =   186
               Text            =   "0.00"
               Top             =   480
               Width           =   645
            End
            Begin VB.CommandButton cmdGetShift 
               Height          =   405
               Left            =   7200
               Picture         =   "frmTimeKeep.frx":2260
               Style           =   1  'Graphical
               TabIndex        =   179
               ToolTipText     =   "Get Shift Schedule"
               Top             =   240
               Width           =   465
            End
            Begin VB.TextBox txtDTR 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   7
               Left            =   6480
               TabIndex        =   188
               Text            =   "0.00"
               Top             =   480
               Visible         =   0   'False
               Width           =   645
            End
            Begin VB.CommandButton cmdRevert 
               Height          =   405
               Left            =   8160
               Picture         =   "frmTimeKeep.frx":26A2
               Style           =   1  'Graphical
               TabIndex        =   178
               ToolTipText     =   "Reset Time In/Out"
               Top             =   240
               Width           =   465
            End
            Begin VB.TextBox txtDTR 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   8
               Left            =   5880
               TabIndex        =   187
               Text            =   "0.00"
               Top             =   480
               Width           =   645
            End
            Begin MSComCtl2.DTPicker ChronoTime 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "H:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   4
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   240
               TabIndex        =   219
               ToolTipText     =   "Edit Time IN"
               Top             =   480
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "00:00"
               Format          =   143785986
               CurrentDate     =   39255
            End
            Begin MSComCtl2.DTPicker ChronoTime 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "H:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   4
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   1560
               TabIndex        =   220
               ToolTipText     =   "Edit Time OUT"
               Top             =   480
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "00:00"
               Format          =   143785986
               CurrentDate     =   39255
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Time OUT"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   70
               Left            =   1680
               TabIndex        =   226
               Top             =   870
               Width           =   1035
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "RD OT"
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
               Index           =   69
               Left            =   6600
               TabIndex        =   218
               Top             =   240
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "To:"
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
               Index           =   52
               Left            =   3360
               TabIndex        =   202
               Top             =   1560
               Width           =   315
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Time IN"
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
               Index           =   54
               Left            =   600
               TabIndex        =   201
               Top             =   225
               Width           =   675
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Second Cutoff from:"
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
               Index           =   55
               Left            =   240
               TabIndex        =   200
               Top             =   1560
               Width           =   1635
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Time OUT"
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
               Index           =   53
               Left            =   1920
               TabIndex        =   199
               Top             =   240
               Width           =   915
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Hours"
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
               Index           =   56
               Left            =   3000
               TabIndex        =   198
               Top             =   240
               Width           =   555
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Lates"
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
               Index           =   57
               Left            =   4200
               TabIndex        =   197
               Top             =   240
               Width           =   435
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Absent"
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
               Index           =   58
               Left            =   3600
               TabIndex        =   196
               Top             =   240
               Width           =   555
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "UDT"
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
               Index           =   59
               Left            =   4920
               TabIndex        =   195
               Top             =   240
               Width           =   435
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "OT"
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
               Index           =   60
               Left            =   5520
               TabIndex        =   194
               Top             =   240
               Width           =   315
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Day"
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
               Index           =   61
               Left            =   120
               TabIndex        =   193
               Top             =   225
               Width           =   315
            End
            Begin VB.Label lblclient 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "1"
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
               Index           =   62
               Left            =   0
               TabIndex        =   192
               Top             =   510
               Width           =   315
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "NDiff"
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
               Index           =   64
               Left            =   6000
               TabIndex        =   191
               Top             =   240
               Width           =   435
            End
         End
         Begin VB.Frame Frame9 
            ForeColor       =   &H00FF0000&
            Height          =   975
            Left            =   2880
            TabIndex        =   203
            Top             =   840
            Width           =   8895
            Begin VB.CommandButton cmdFirstHalf 
               Appearance      =   0  'Flat
               Caption         =   "First Half"
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
               Left            =   120
               TabIndex        =   208
               ToolTipText     =   "First Half"
               Top             =   600
               Width           =   4335
            End
            Begin VB.CommandButton cmdSecondHalf 
               Caption         =   "Second Half"
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
               Left            =   4440
               TabIndex        =   207
               ToolTipText     =   "Second Half"
               Top             =   600
               Width           =   4335
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
               ItemData        =   "frmTimeKeep.frx":27A4
               Left            =   1440
               List            =   "frmTimeKeep.frx":27A6
               Style           =   2  'Dropdown List
               TabIndex        =   206
               TabStop         =   0   'False
               Top             =   180
               Width           =   1365
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
               ItemData        =   "frmTimeKeep.frx":27A8
               Left            =   2865
               List            =   "frmTimeKeep.frx":27AA
               Sorted          =   -1  'True
               TabIndex        =   205
               TabStop         =   0   'False
               Text            =   "cmbYear"
               Top             =   180
               Width           =   930
            End
            Begin VB.CommandButton cmdDTRRefresh 
               Caption         =   "Compute DTR"
               Height          =   405
               Left            =   3840
               Style           =   1  'Graphical
               TabIndex        =   204
               ToolTipText     =   "Refresh"
               Top             =   150
               Width           =   1425
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Calendar Month:"
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
               TabIndex        =   211
               Top             =   240
               Width           =   1275
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Records:"
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
               Index           =   46
               Left            =   5400
               TabIndex        =   210
               Top             =   120
               Width           =   1875
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Records:"
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
               Index           =   63
               Left            =   5400
               TabIndex        =   209
               Top             =   360
               Width           =   3075
            End
         End
         Begin VB.Frame Frame3 
            Height          =   885
            Left            =   2880
            TabIndex        =   53
            Top             =   120
            Width           =   8895
            Begin VB.CommandButton cmdShiftSched 
               Height          =   405
               Left            =   8280
               Picture         =   "frmTimeKeep.frx":27AC
               Style           =   1  'Graphical
               TabIndex        =   229
               ToolTipText     =   "Save Shift Schedule"
               Top             =   240
               Width           =   465
            End
            Begin VB.ComboBox cmbShiftSchedule 
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
               Height          =   315
               ItemData        =   "frmTimeKeep.frx":2CDE
               Left            =   6120
               List            =   "frmTimeKeep.frx":2CE0
               TabIndex        =   228
               Text            =   "cmbShiftSchedule"
               Top             =   370
               Width           =   2040
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
               Index           =   10
               Left            =   4800
               MultiLine       =   -1  'True
               TabIndex        =   157
               Text            =   "frmTimeKeep.frx":2CE2
               Top             =   400
               Width           =   1275
            End
            Begin VB.CheckBox chkNoFilter 
               Caption         =   "All Employees"
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   1485
               TabIndex        =   149
               Top             =   150
               Visible         =   0   'False
               Width           =   1335
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
               Left            =   3720
               MultiLine       =   -1  'True
               TabIndex        =   148
               Text            =   "frmTimeKeep.frx":2CEE
               Top             =   400
               Width           =   1035
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
               Left            =   4800
               MultiLine       =   -1  'True
               TabIndex        =   139
               Text            =   "frmTimeKeep.frx":2CFA
               Top             =   400
               Width           =   675
            End
            Begin VB.CheckBox chkActiveClient 
               Caption         =   "Active"
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
               Left            =   4680
               TabIndex        =   5
               Top             =   400
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   825
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
               Left            =   180
               TabIndex        =   6
               Text            =   "Client Name"
               Top             =   400
               Width           =   2535
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
               Index           =   1
               Left            =   1530
               MultiLine       =   -1  'True
               TabIndex        =   4
               Text            =   "frmTimeKeep.frx":2D06
               Top             =   1320
               Width           =   2355
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
               Left            =   2820
               TabIndex        =   3
               Text            =   "Client Name"
               Top             =   400
               Width           =   855
            End
            Begin VB.ComboBox cmbTaxCode 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Height          =   315
               ItemData        =   "frmTimeKeep.frx":2D12
               Left            =   4080
               List            =   "frmTimeKeep.frx":2D1C
               TabIndex        =   7
               Text            =   "cmbTaxCode"
               Top             =   400
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Status:"
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
               Index           =   68
               Left            =   4800
               TabIndex        =   158
               Top             =   180
               Width           =   1215
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
               Left            =   6120
               TabIndex        =   140
               Top             =   180
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
               Left            =   3720
               TabIndex        =   57
               Top             =   180
               Width           =   1305
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Employee#:"
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
               Left            =   2820
               TabIndex        =   56
               Top             =   180
               Width           =   1455
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
               Index           =   0
               Left            =   180
               TabIndex        =   55
               Top             =   180
               Width           =   1815
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Employment Status:"
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
               Index           =   40
               Left            =   1530
               TabIndex        =   54
               Top             =   1095
               Width           =   2295
            End
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   8760
            X2              =   11400
            Y1              =   7800
            Y2              =   7800
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "By:"
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
            Index           =   24
            Left            =   5040
            TabIndex        =   103
            Top             =   5170
            Width           =   435
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
            Left            =   1890
            TabIndex        =   60
            Top             =   840
            Width           =   825
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Edit Date:"
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
            Left            =   6570
            TabIndex        =   59
            Top             =   5170
            Width           =   915
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Noted by:"
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
            Left            =   6000
            TabIndex        =   58
            Top             =   8040
            Width           =   735
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
            Left            =   180
            TabIndex        =   43
            Top             =   840
            Width           =   1365
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
            Left            =   180
            TabIndex        =   42
            Top             =   200
            Width           =   1365
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "If No (state reason):"
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
            Left            =   360
            TabIndex        =   41
            Top             =   7920
            Width           =   1545
         End
         Begin VB.Label lblclient 
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
            Index           =   12
            Left            =   4200
            TabIndex        =   40
            Top             =   8040
            Width           =   735
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7470
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTimeKeep.frx":2D30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CheckBox chkEditDTR 
         Height          =   375
         Left            =   2640
         Picture         =   "frmTimeKeep.frx":2D8E
         Style           =   1  'Graphical
         TabIndex        =   153
         ToolTipText     =   "Edit DTR Records"
         Top             =   30
         Width           =   495
      End
      Begin VB.CommandButton cmdEdit 
         Height          =   405
         Left            =   6240
         Picture         =   "frmTimeKeep.frx":31D0
         Style           =   1  'Graphical
         TabIndex        =   152
         ToolTipText     =   "Save"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CheckBox chkChangeCutoff 
         Height          =   375
         Left            =   2160
         Picture         =   "frmTimeKeep.frx":3702
         Style           =   1  'Graphical
         TabIndex        =   151
         ToolTipText     =   "Update Payroll Cut-off Dates"
         Top             =   30
         Width           =   495
      End
      Begin VB.CommandButton cmdDeductionTypes 
         Height          =   405
         Left            =   5400
         Picture         =   "frmTimeKeep.frx":3B44
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Deduction Types"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   1440
         Picture         =   "frmTimeKeep.frx":3F86
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   3360
         Picture         =   "frmTimeKeep.frx":4088
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmTimeKeep.frx":41D2
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "New"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   480
         Picture         =   "frmTimeKeep.frx":4704
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   960
         Picture         =   "frmTimeKeep.frx":4C36
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdClose 
         Height          =   405
         Left            =   3840
         Picture         =   "frmTimeKeep.frx":5168
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   800
         Y1              =   1500
         Y2              =   1200
      End
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "Agency Rate:"
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
      Index           =   23
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PURPOSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(1)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   210
   End
End
Attribute VB_Name = "frmTimeKeep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public i As Integer

Dim rsEmployeeBasicData As ADODB.Recordset
Dim rsInactiveEmployee As ADODB.Recordset
Dim rsClientData As ADODB.Recordset
Dim rstShift As ADODB.Recordset ' for Shift Schedule updated May 10, 2014
Dim mTripNo As Integer

'Dim rsDeductionTypes As ADODB.Recordset
'Dim rsEmpDeductions As ADODB.Recordset
'Dim rsGetDeductionTypes As ADODB.Recordset
Dim rsDTRRecords As ADODB.Recordset
Dim rsDTRPayroll As ADODB.Recordset

Dim StartFirstCutoff As Date
Dim EndFirstCutoff As Date
Dim StartSecondCutoff As Date
Dim EndSecondCutoff As Date
Dim FirstHalfFlag As Boolean    ' to determine whether first half or second half

Dim IniStartCutoff As Date      ' for initial Start Cutoff
Dim IniEndCutoff  As Date       ' for initial End Cutoff
Dim UseIniCutoff As Boolean     ' flag whether to use the initial cut-off
Dim StartCutoff As Date         ' New Start cut-off in date format
Dim EndCutoff As Date           ' New Ene cut-off in date format

Dim StrQuery As String
Dim mRecNo As Integer
Dim mCCode As Integer
Dim mCompany As String
Dim mEmpName As String
Dim ActiveCtr As Integer
Dim InActiveCtr As Integer
Dim RefreshSw As Boolean
Dim paymentbase As String
Dim CurrEmployee As String
Dim CurrAmount As String    ' holds the current value of salaries
Dim AmountSw As Boolean     ' variable for basic salary entries
Dim ActiveSw As Boolean
Dim FormLoadSw As Boolean

Dim workdaysperyr As Integer
Dim noofdayspermon As Integer

Dim AddFlag As Boolean
Dim EditFlag As Boolean
Dim AddDeductions As Boolean
Dim EditDeductions As Boolean
Dim mDeductionCode As Integer

Dim NoofPayments As Integer
Dim TotalPayments As Double
Dim NoEditFlag As Boolean
Dim mInterest As Double
Dim mTotalLoan As Double
Dim mPayments As Double
Dim ReProcessFlag As Boolean
Dim mDTRKEY As String

Dim DTRStartTime As String
Dim DTREndTime As String
Dim WorkHours As Integer

Private Sub chkChangeCutoff_Click()
If chkChangeCutoff.VALUE = 1 Then
    Frame6.Visible = True
    'Frame4.Enabled = False
Else
    Frame6.Visible = False
    'Frame4.Enabled = True
End If
End Sub

Private Sub chkEditDTR_Click()
If chkEditDTR.VALUE = 1 Then
    If CheckFinalPayroll = False Then     ' Allow only edit if payroll entries is not yet move.
        Frame7.Visible = True
        Frame10.Visible = False
        Frame9.Visible = False
        cmbCompany.Enabled = False
        ListView1(0).Enabled = False
        cmdReprocess.Enabled = False
        cmdLeaveApplication.Enabled = False
        cmdDTRManual.Enabled = False
        cmdDTRPayroll.Enabled = False
        cmdRefresh.Enabled = False
        cmdadd.Enabled = False
        cmddelete.Enabled = False
        cmdClose.Enabled = False
        If UCase(paymentbase) = "MONTHLY" Then
            txtDTR(2).Enabled = False
        Else
            txtDTR(2).Enabled = True
        End If
    Else
        MsgBox "Payroll Cut-off already Final." & Chr(13) & Chr(13) & "NOT ALLOW TO EDIT DTR RECORDS!", vbInformation, PayrollAlert
        chkEditDTR.VALUE = 0
    End If
Else
    Frame7.Visible = False
    Frame10.Visible = True
    Frame9.Visible = True
    cmbCompany.Enabled = True
    ListView1(0).Enabled = True
    cmdReprocess.Enabled = True
    cmdLeaveApplication.Enabled = True
    cmdDTRManual.Enabled = True
    cmdDTRPayroll.Enabled = True
    cmdRefresh.Enabled = True
    cmdadd.Enabled = True
    cmddelete.Enabled = True
    cmdClose.Enabled = True
End If
End Sub

Private Function CheckFinalPayroll() As Boolean
Dim mDateKey As Date
Dim rsCut As Recordset

CheckFinalPayroll = False
On Error Resume Next

If FirstHalfFlag = True Then
    mDateKey = Format(Left(cmbMonth.Text, 2) & "/01/" & Trim$(str(cmbYear.Text)), "mm/dd/yyyy")
Else
    mDateKey = Format(Left(cmbMonth.Text, 2) & "/16/" & Trim$(str(cmbYear.Text)), "mm/dd/yyyy")
End If
StrQuery = "SELECT * FROM cutofftable WHERE ClientCode = " & txtClientCode & " AND '" & mDateKey & "' BETWEEN StartCutoffDate AND EndCutoffDate AND BilledFlag = 1"
Set rsCut = ReturnRst(StrQuery, cnnstr, mstatic)
If rsCut.RecordCount > 0 Then
    CheckFinalPayroll = True
Else
    CheckFinalPayroll = False
End If
Set rsCut = Nothing

End Function

Private Sub ChronoTime_LostFocus(Index As Integer)
Dim mHr As Integer
Dim mMn As Integer
Dim mHr1 As Integer
Dim mMn1 As Integer
Dim mLate As Double

DTRStartTime = mID(txtClient(8).Text, InStr(1, txtClient(8).Text, "[", vbTextCompare) + 1, 5)
DTREndTime = mID(txtClient(8).Text, InStr(1, txtClient(8).Text, "-", vbTextCompare) + 2, 5)
mLate = 0#

' Compute for lates.
If Index = 0 Then
    If FMSFlag = False Then
    mHr = mID(DTRStartTime, 1, InStr(1, DTRStartTime, ":", vbTextCompare) - 1)
    mHr1 = mID(Format(ChronoTime(0), "HH:MM"), 1, InStr(1, Format(ChronoTime(0), "HH:MM"), ":", vbTextCompare) - 1)
    mMn = mID(DTRStartTime, InStr(1, DTRStartTime, ":", vbTextCompare) + 1, 2)
    mMn1 = mID(Format(ChronoTime(0), "HH:MM"), InStr(1, Format(ChronoTime(0), "HH:MM"), ":", vbTextCompare) + 1, 2)
    If mHr1 > mHr And mHr <> 0 Then
        mLate = mHr1 - mHr
        mLate = mLate + ((mMn1 - (mMn + 10)) / 60)
    End If
    If mHr1 = mHr And mMn1 > (mMn + 10) Then
        mLate = mLate + ((mMn1 - (mMn + 10)) / 60)
    End If
    End If
    txtDTR(4).Text = Format(mLate, "#0.00")
ElseIf Index = 1 Then ' compute for UDT
    If FMSFlag = False Then
    mHr = mID(DTREndTime, 1, InStr(1, DTREndTime, ":", vbTextCompare) - 1)
    mHr1 = mID(Format(ChronoTime(1), "HH:MM"), 1, InStr(1, Format(ChronoTime(1), "HH:MM"), ":", vbTextCompare) - 1)
    mMn = mID(DTREndTime, InStr(1, DTREndTime, ":", vbTextCompare) + 1, 2)
    If mMn = 0 Then
        mMn = 60 ' convert to 60 minutes
    End If
    mMn1 = mID(Format(ChronoTime(1), "HH:MM"), InStr(1, Format(ChronoTime(1), "HH:MM"), ":", vbTextCompare) + 1, 2)
    If mHr1 < mHr And mHr <> 0 Then
'        If mHr >= 23 Then
'            mHr = mHr + 1
'        End If
'        If mHr1 >= 23 Then
'            mHr1 = mHr1 + 1
'        End If
        mLate = mHr - (mHr1 + 1)
        If mLate = 0 Then ' compute for UDT in minutes
            mLate = ((mMn - (mMn1)) / 60)
        Else
            mLate = mLate + ((mMn - (mMn1)) / 60)
        End If
    End If
    If mHr1 = mHr And mMn1 < mMn And mMn <> 60 Then
        mLate = mLate + ((mMn - (mMn1)) / 60)
    End If
    End If
    txtDTR(5).Text = Format(mLate, "#0.00")
End If

''Compute for night differential.
mHr = mID(Format(ChronoTime(0), "HH:MM"), 1, InStr(1, Format(ChronoTime(0), "HH:MM"), ":", vbTextCompare) - 1)
mHr1 = mID(Format(ChronoTime(1), "HH:MM"), 1, InStr(1, Format(ChronoTime(1), "HH:MM"), ":", vbTextCompare) - 1)

mLate = 0

If (mHr >= 22 Or mHr = 0) And mHr1 <= 6 Then
    If mHr = 0 Then
        mLate = 0
    Else
        mLate = 24 - mHr
    End If
    If mHr1 > 6 Then
        mHr1 = 6
    End If
    mLate = mLate + mHr1
    mMn = mID(Format(ChronoTime(0), "HH:MM"), InStr(1, Format(ChronoTime(0), "HH:MM"), ":", vbTextCompare) + 1, 2)
    If mMn <> 0 Then
        mLate = mLate - (mMn / 60)
    End If
'    mMn = Mid(Format(ChronoTime(1), "HH:MM"), InStr(1, Format(ChronoTime(1), "HH:MM"), ":", vbTextCompare) + 1, 2)
'    If mHr1 < 6 And mMn <> 0 Then
'        mLate = mLate + (mMn / 60)
'    End If
ElseIf mHr > 18 And mHr1 <= 6 Then
    mLate = 2
    If mHr1 > 6 Then
        mHr1 = 6
    End If
    mLate = mLate + mHr1
    mMn = mID(Format(ChronoTime(0), "HH:MM"), InStr(1, Format(ChronoTime(0), "HH:MM"), ":", vbTextCompare) + 1, 2)
    If mMn <> 0 Then
        mLate = mLate - (mMn / 60)
    End If
'    mMn = Mid(Format(ChronoTime(1), "HH:MM"), InStr(1, Format(ChronoTime(1), "HH:MM"), ":", vbTextCompare) + 1, 2)
'    If mHr1 < 6 And mMn <> 0 Then
'        mLate = mLate + (mMn / 60)
'    End If

ElseIf (mHr >= 22 Or mHr = 0) And mHr1 > 6 Then
    If mHr = 0 Then
        mLate = 0
    Else
        mLate = 24 - mHr
    End If
    If mHr1 > 6 Then
        mHr1 = 6
    End If
    mLate = mLate + mHr1
    mMn = mID(Format(ChronoTime(0), "HH:MM"), InStr(1, Format(ChronoTime(0), "HH:MM"), ":", vbTextCompare) + 1, 2)
    If mMn <> 0 Then
        mLate = mLate - (mMn / 60)
    End If
'    mMn = Mid(Format(ChronoTime(1), "HH:MM"), InStr(1, Format(ChronoTime(1), "HH:MM"), ":", vbTextCompare) + 1, 2)
'    If mHr1 < 6 And mMn <> 0 Then
'        mLate = mLate + (mMn / 60)
'    End If

ElseIf mHr >= 14 And mHr1 >= 22 Then
    If mHr1 > 6 And mHr1 < 12 Then
        mLate = 2
        mHr1 = 6
    Else
        mLate = 0
        If mHr1 = 22 Then
            mHr1 = 0
        Else
            mHr1 = 24 - mHr1
        End If
    End If
    mLate = mLate + mHr1
    mMn = mID(Format(ChronoTime(0), "HH:MM"), InStr(1, Format(ChronoTime(0), "HH:MM"), ":", vbTextCompare) + 1, 2)
    If mMn <> 0 Then
        mLate = mLate - (mMn / 60)
    End If
    mMn = mID(Format(ChronoTime(1), "HH:MM"), InStr(1, Format(ChronoTime(1), "HH:MM"), ":", vbTextCompare) + 1, 2)
    If mHr1 < 6 And mMn <> 0 Then
        mLate = mLate + (mMn / 60)
    End If
ElseIf mHr >= 14 And mHr1 <= 6 Then
    If mHr1 <= 6 Then
        mLate = 2
    Else
        mLate = 0
        If mHr1 = 22 Then
            mHr1 = 0
        Else
            mHr1 = 24 - mHr1
        End If
    End If
    mLate = mLate + mHr1
    mMn = mID(Format(ChronoTime(0), "HH:MM"), InStr(1, Format(ChronoTime(0), "HH:MM"), ":", vbTextCompare) + 1, 2)
    If mMn <> 0 Then
        mLate = mLate - (mMn / 60)
    End If
    mMn = mID(Format(ChronoTime(1), "HH:MM"), InStr(1, Format(ChronoTime(1), "HH:MM"), ":", vbTextCompare) + 1, 2)
    If mHr1 < 6 And mMn <> 0 Then
        mLate = mLate + (mMn / 60)
    End If

End If

txtDTR(8).Text = Format(mLate, "#0")


''DECLARE @StartTime Varchar(10), @EndTime Varchar(10), @MidDiff Decimal(9,2)
''DECLARE @NightDiffHours Decimal(9,2)
''DECLARE @HR Decimal(9,2), @MIN Decimal(9,2)
''DECLARE @HR1 Decimal(9,2), @MIN1 Decimal(9,2)
'
'Dim mStartTime As String, mEndTime As String
'Dim mDTRTime As String, mDTREnd As String
'Dim mMidDiff As Double, mNightDiffHours As Double
'
'mMidDiff = (24 * 3600)
'mStartTime = "22:00"
'mEndTime = "06:00"
'mDTRTime = LTrim(RTrim(mHr))
''SET @DTREnd = LTRIM(RTRIM(@DTREnd))
''SET @NightDiffHours = 0.00
''
''IF LEN(@DTRTime) <> 5 OR LEN(@DTREnd) <> 5
''    GoTo Exit1
''
'-- Check first if schedule is Midnight.
'IF CONVERT(int, REPLACE(@DTRTime,  ':', '') )  >= CONVERT(int, REPLACE(@StartTime,  ':', '') )
'BEGIN
'    -- Compute for Start DTR Time.
'    SET @HR = CONVERT(DECIMAL(9,2), LEFT(@DTRTime,2))
'    SET @MIN = CONVERT(DECIMAL(9,2), RIGHT(@DTRTime,2)) * 60
'    SET @HR1 = CONVERT(DECIMAL(9,2), LEFT(@DTREnd,2))
'    IF @HR1 > 23
'        SET @HR1 = @HR1 - 24 -- updated 10/8/2011
'    SET @MIN1 = CONVERT(DECIMAL(9,2), RIGHT(@DTREnd,2)) * 60
'    IF @HR > 21
'        SET @NightDiffHours = @MidDiff - ( ( @HR * 3600 ) + @MIN )
'    -- Compute for End DTR Time.
'    IF @HR1 > 6
'    BEGIN
'        SET @HR1 = 6    -- set the end night diff to 6 if exceeded
'        SET @MIN1 = 0
'    End
'    SET @NightDiffHours =  @NightDiffHours + ( ( @HR1 * 3600 ) + @MIN1 )
'End
'ELSE IF CONVERT(int, REPLACE(@DTREnd,  ':', '') )  > CONVERT(int, REPLACE(@StartTime,  ':', '') )
'BEGIN
'    -- Compute for End DTR Time.
'    SET @HR = CONVERT(DECIMAL(9,2), LEFT(@DTREnd,2))
'    SET @MIN = CONVERT(DECIMAL(9,2), RIGHT(@DTREnd,2)) * 60
'    IF @HR > 22
'        SET @NightDiffHours = ( ( @HR * 3600 ) + @MIN ) - (22 * 3600)
'    Else
'    IF @HR = 6
'        SET @MIN = 0
'    IF @HR <= 6
'        SET @NightDiffHours =   ( @HR * 3600 ) + @MIN
'
'End
'ELSE IF CONVERT(int, REPLACE(@DTREnd,  ':', '') )  <= CONVERT(int, REPLACE(@EndTime,  ':', '') )
'BEGIN
'    -- Compute for End DTR Time.
'    SET @HR = CONVERT(DECIMAL(9,2), LEFT(@DTREnd,2))
'    SET @MIN = CONVERT(DECIMAL(9,2), RIGHT(@DTREnd,2)) * 60
'    IF @HR > 6
'    BEGIN
'        SET @HR = 6 -- set the end night diff to 6 if exceeded
'        SET @MIN = 0
'    End
'    SET @NightDiffHours =  @NightDiffHours + ( @HR * 3600 ) + @MIN
'
'    -- Check the Start Time.
'    SET @HR = CONVERT(DECIMAL(9,2), LEFT(@DTRTime,2))
'    SET @MIN = CONVERT(DECIMAL(9,2), RIGHT(@DTRTime,2)) * 60
'    IF @HR < 6
'        SET @NightDiffHours = @NightDiffHours -  ( ( @HR * 3600 ) + @MIN )
'
'End
'
'SET @NightDiffHours = CONVERT(int, (@NightDiffHours / 3600))

End Sub

'Private Sub cmbPaySystem_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

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

Private Sub cmbShiftSchedule_GotFocus()
cmbShiftSchedule.BackColor = &HFFFFC0
End Sub

Private Sub cmbShiftSchedule_LostFocus()
cmbShiftSchedule.BackColor = vbWhite
End Sub

Private Sub cmdCancelCutoff_Click()
chkChangeCutoff.VALUE = 0
Call chkChangeCutoff_Click
End Sub

Private Sub cmdClose_Click()
    Set rsEmployeeBasicData = Nothing
    Unload Me
End Sub

Private Sub cmdDTRManual_Click()
If lstDTRDisplay.ListItems.Count <> 0 Then
    Call GenerateDTRReport
Else
    MsgBox "No DTR records to print!", vbInformation, PayrollAlert
End If
End Sub

Private Sub GenerateDTRReport()
Dim rsDedRecords As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String, mRange2 As String
Dim sumString As String
Dim ctr As Long
Dim i As Integer, l As Integer
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

mHead = "DTR Records Report for " & ListView1(0).SelectedItem.Text
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy")
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "COMPANY NAME: " & cmbCompany.Text
Call PrintReportHeader(xlSheet, xlApp1, "A5", mHead, "", xlLeft, , False, False, False, , , 10, True)

Call PrintReportHeader(xlSheet, xlApp1, "A6:A7", "Record #", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "B6:B7", "DTR ID", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "C6:C7", "DTR Day", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "D6:D7", "DTR Date", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "E6:E7", "Worked", "Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "F6:F7", "Time In", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "G6:G7", "Time Out", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "H6:H7", "No. Hours", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "I6:I7", "Absent", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "J6:J7", "Late", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "K6:K7", "Undertime", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "L6:L7", "Night Diff", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "M6:M7", "Remarks", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 25)

For i = 1 To lstDTRDisplay.ListItems.Count
        xlSheet.Cells(ctr, 1).VALUE = i
        xlSheet.Cells(ctr, 2).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(3) ' DTR ID
        xlSheet.Cells(ctr, 2).NumberFormat = "00000000"
        xlSheet.Cells(ctr, 3).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(1) ' DTR Day
        xlSheet.Cells(ctr, 4).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(2) ' DTR Date
        xlSheet.Cells(ctr, 5).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(4) ' Work Hours
        xlSheet.Cells(ctr, 5).NumberFormat = "0.00"
        xlSheet.Cells(ctr, 6).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(5) ' Time In
        xlSheet.Cells(ctr, 7).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(6) ' Time Out
        xlSheet.Cells(ctr, 8).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(7) ' No Hours
        xlSheet.Cells(ctr, 9).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(8) ' Absent
        xlSheet.Cells(ctr, 10).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(9) ' Late
        xlSheet.Cells(ctr, 11).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(10) ' Undertime
        xlSheet.Cells(ctr, 12).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(12) ' Night Diff
        xlSheet.Cells(ctr, 13).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(13) ' remarks
        ctr = ctr + 1
Next i

For l = 69 To 76  ' E to R
    mRange = Chr(l) & 8 & ":" & Chr(l) & ctr
    xlSheet.Range(mRange).Select
    mRange = Chr(l) & ctr + 1
    xlSheet.Range(mRange).Activate
    If l = 70 Or l = 71 Then ' 70
    Else
        xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"
    End If
Next l

ctr = ctr + 2

'Box it.
mRange = "A6:M" & ctr
Call PlaceBox(xlSheet, xlApp1, mRange)

'ctr = ctr + 1
'mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
'mRange = "A" & ctr
'Call CustomizedHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 8, False, 25)

xlSheet.Range("A1").Select

ExitRtn2:

End Sub

Private Sub cmdDTRPayroll_Click()
Dim res

res = MsgBox("Include DTR of " & txtClient(2).Text & " for Payroll processing?", vbQuestion + vbYesNo, PayrollAlert)
If res = vbYes Then ' mark for payroll processing

    If FMSFlag = False Then
        ' Admin
        If FirstHalfFlag = True Then
            mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "01"
            StrQuery = "EXEC dbo.ComputeDTRCutOff " & cmbYear.Text & ", " & Val(Left(cmbMonth.Text, 2)) & ", 1, '" & txtClient(9).Text & "', 2, " & IIf(ReProcessFlag = True, 1, 0) & ", '" & StartCutoff & "', '" & EndCutoff & "'"
        Else
            mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "02"
            StrQuery = "EXEC dbo.ComputeDTRCutOff " & cmbYear.Text & ", " & Val(Left(cmbMonth.Text, 2)) & ", 2, '" & txtClient(9).Text & "', 2, " & IIf(ReProcessFlag = True, 1, 0) & ", '" & StartCutoff & "', '" & EndCutoff & "'"
        End If
        Call Action_Query(StrQuery, cnnstr)
        StrQuery = "UPDATE EmployeeBasicSalary SET DTRKEY = '" & mDTRKEY & "', DTRID = '" & txtClient(9).Text & "'" & _
            " WHERE EmployeeID = " & txtEmpNum 'txtClient(0)
        Call Action_Query(StrQuery, cnnstr)
    Else
        ' Driver / Conductor
        If FirstHalfFlag = True Then
            mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "01"
            StrQuery = "EXEC dbo.ComputeDTRCon " & cmbYear.Text & ", " & Val(Left(cmbMonth.Text, 2)) & ", 1, '" & txtClient(9).Text & "', 2, " & IIf(ReProcessFlag = True, 1, 0)
        Else
            mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "02"
            StrQuery = "EXEC dbo.ComputeDTRCon " & cmbYear.Text & ", " & Val(Left(cmbMonth.Text, 2)) & ", 2, '" & txtClient(9).Text & "', 2, " & IIf(ReProcessFlag = True, 1, 0)
        End If
        Call Action_Query(StrQuery, cnnstr)
        StrQuery = "UPDATE EmployeeBasicSalary SET DTRKEY = '" & mDTRKEY & "', DTRID = '" & txtClient(9).Text & "'" & _
            " WHERE EmployeeID = " & txtEmpNum 'txtClient(0)
        Call Action_Query(StrQuery, cnnstr)
    End If
    MsgBox txtClient(2) & " DTR was successfully imported for Payroll processing!", vbInformation, PayrollAlert
    lstDTRDisplay.ForeColor = vbBlue
    Call cmdDTRRefresh_Click
Else
    MsgBox "DTR importing was cancelled!", vbInformation, PayrollAlert
    lstDTRDisplay.ForeColor = vbBlack
End If

End Sub

Private Sub cmdDTRRefresh_Click()
Dim dtrYr As String, dtrMon As String

If rsDTRRecords.State = 1 Then
    rsDTRRecords.Close
End If

If rsDTRPayroll.State = 1 Then
    rsDTRPayroll.Close
End If

If FMSFlag = True Then
    Call GetFMSDTR
    Exit Sub
End If

'Check for New Cut-off. Updated 4/3/2016.
If cmbYear.Text = "" Then
    dtrYr = Trim$(str(Format(Date, "yyyy")))
Else
    dtrYr = Trim$(cmbYear.Text)
End If
'Get the Month for New Cut-Off.
dtrMon = Left(cmbMonth.Text, 2)
If FirstHalfFlag = True Then
    If Val(dtrMon) = 1 Then    'Need to check the month if January
        'dtrYr = Trim$(str(Val(dtrYr) - 1))
        'dtrMon = "12"
        If UseIniCutoff = False Then
            StartCutoff = Format("12/" & Trim$(str(Day(StartFirstCutoff))) & "/" & Trim$(str(Val(dtrYr) - 1)), "mm/dd/yyyy")
            EndCutoff = Format(dtrMon & "/" & Trim$(str(Day(EndFirstCutoff))) & "/" & dtrYr, "mm/dd/yyyy")
        Else
            StartCutoff = Format("12/" & Trim$(str(Day(IniStartCutoff))) & "/" & Trim$(str(Val(dtrYr) - 1)), "mm/dd/yyyy")
            EndCutoff = Format(dtrMon & "/" & Trim$(str(Day(IniEndCutoff))) & "/" & dtrYr, "mm/dd/yyyy")
        End If
    Else
        If UseIniCutoff = False Then
            StartCutoff = Format(Trim$(str(Val(dtrMon) - 1)) & "/" & Trim$(str(Day(StartFirstCutoff))) & "/" & dtrYr, "mm/dd/yyyy")
            EndCutoff = Format(dtrMon & "/" & Trim$(str(Day(EndFirstCutoff))) & "/" & dtrYr, "mm/dd/yyyy")
        Else
            StartCutoff = Format(dtrMon & "/" & Trim$(str(Day(IniStartCutoff))) & "/" & dtrYr, "mm/dd/yyyy")
            EndCutoff = Format(dtrMon & "/" & Trim$(str(Day(IniEndCutoff))) & "/" & dtrYr, "mm/dd/yyyy")
        End If
    End If
Else    ' For Second Cutoff
    If UseIniCutoff = False Then
        StartCutoff = Format(dtrMon & "/" & Trim$(str(Day(StartSecondCutoff))) & "/" & dtrYr, "mm/dd/yyyy")
        EndCutoff = Format(dtrMon & "/" & Trim$(str(Day(EndSecondCutoff))) & "/" & dtrYr, "mm/dd/yyyy")
    Else
        StartCutoff = Format(dtrMon & "/" & Trim$(str(Day(IniStartCutoff))) & "/" & dtrYr, "mm/dd/yyyy")
        EndCutoff = Format(dtrMon & "/" & Trim$(str(Day(IniEndCutoff))) & "/" & Trim$(str(dtrYr)), "mm/dd/yyyy")
    End If
End If

If chkNoFilter.VALUE = 1 Then ' ALl Employees
    If FirstHalfFlag = True Then
        ' Get DTR records for the current month First Cutoff.
        mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "01"

'        StrQuery = "SELECT * FROM vwNewDTRRecords WHERE DTRDay < 16 AND YEAR(DTR_Date) = " & _
'            IIf(cmbYear.Text = "", Format(Date, "yyyy"), cmbYear.Text) & _
'            " AND MONTH(DTR_Date) = " & Val(IIf(cmbMonth = "", Format(Date, "mm"), Left(cmbMonth.Text, 2))) & " ORDER BY DTRID "
    
        If UseIniCutoff = False Then
            StrQuery = "SELECT * FROM vwNewDTRRecords WHERE DTRDay < 16 AND YEAR(DTR_Date) = " & _
                IIf(cmbYear.Text = "", Format(Date, "yyyy"), cmbYear.Text) & _
                " AND MONTH(DTR_Date) = " & Val(IIf(cmbMonth = "", Format(Date, "mm"), Left(cmbMonth.Text, 2))) & " ORDER BY DTRID "
        Else    ' use initial dtr date range
            StrQuery = "SELECT * FROM vwNewDTRRecords WHERE DTRDay BETWEEN " & Format(DTIniStartCutoff.VALUE, """dd""") & " and " & Format(DTIniEndCutoff.VALUE, """dd""") & " AND YEAR(DTR_Date) = " & _
                IIf(cmbYear.Text = "", Format(Date, "yyyy"), cmbYear.Text) & _
                " AND MONTH(DTR_Date) = " & Val(IIf(cmbMonth = "", Format(Date, "mm"), Left(cmbMonth.Text, 2))) & " ORDER BY DTRID "
        End If
    
    Else
        ' Get DTR records for the current month Second Cutoff.
        mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "02"
'        StrQuery = "SELECT * FROM vwNewDTRRecords WHERE DTRDay > 15 AND YEAR(DTR_Date) = " & _
'            IIf(cmbYear.Text = "", Format(Date, "yyyy"), cmbYear.Text) & _
'            " AND MONTH(DTR_Date) = " & Val(IIf(cmbMonth = "", Format(Date, "mm"), Left(cmbMonth.Text, 2))) & " ORDER BY DTRID "
    
        If UseIniCutoff = False Then    ' Updated 3/27/2016 for new cutoff
            'StrQuery = "SELECT * FROM vwNewDTRRecords WHERE DTRDay > 15 AND YEAR(DTR_Date) = "
            StrQuery = "SELECT * FROM vwNewDTRRecords WHERE DTRDay > 15 AND YEAR(DTR_Date) = " & _
                IIf(cmbYear.Text = "", Format(Date, "yyyy"), cmbYear.Text) & _
                " AND MONTH(DTR_Date) = " & Val(IIf(cmbMonth = "", Format(Date, "mm"), Left(cmbMonth.Text, 2))) & " ORDER BY DTRID "
        Else    ' use initial dtr date range
            StrQuery = "SELECT * FROM vwNewDTRRecords WHERE DTRDay BETWEEN " & Format(DTIniStartCutoff.VALUE, """dd""") & " and " & Format(DTIniEndCutoff.VALUE, """dd""") & " AND YEAR(DTR_Date) = " & _
                IIf(cmbYear.Text = "", Format(Date, "yyyy"), cmbYear.Text) & _
                " AND MONTH(DTR_Date) = " & Val(IIf(cmbMonth = "", Format(Date, "mm"), Left(cmbMonth.Text, 2))) & " ORDER BY DTRID "
        End If
    End If
Else    ' for selected employee, updated for new cut-off
    If FirstHalfFlag = True Then
        ' Get DTR records for the current month First Cutoff.
        mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "01"
        StrQuery = "EXEC dbo.ComputeDTRCutOff " & cmbYear.Text & ", " & Val(Left(cmbMonth.Text, 2)) & ", 1, '" & txtClient(9).Text & "', 1, " & IIf(ReProcessFlag = True, 1, 0) & ", '" & StartCutoff & "', '" & EndCutoff & "'"
        Call Action_Query(StrQuery, cnnstr)
'        StrQuery = "SELECT * FROM DTRFinalRecords WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
'            " AND DTRKEY = '" & mDTRKEY & "' ORDER BY NoOfDay"
    
        If UseIniCutoff = False Then
            StrQuery = "SELECT * FROM DTRFinalRecords WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
                " AND DTRKEY = '" & mDTRKEY & "' ORDER BY NoOfDay"
        Else    ' Initial only
            StrQuery = "SELECT * FROM DTRFinalRecords WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
                " AND DTRDay BETWEEN " & Format(DTIniStartCutoff.VALUE, "dd") & " and " & Format(DTIniEndCutoff.VALUE, "dd") & _
                " AND DTRKEY = '" & mDTRKEY & "' ORDER BY NoOfDay"
        End If
    Else
        ' Get DTR records for the current month Second Cutoff
        mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "02"
        StrQuery = "EXEC dbo.ComputeDTRCutOff " & cmbYear.Text & ", " & Val(Left(cmbMonth.Text, 2)) & ", 2, '" & txtClient(9).Text & "', 1, " & IIf(ReProcessFlag = True, 1, 0) & ", '" & StartCutoff & "', '" & EndCutoff & "'"
        Call Action_Query(StrQuery, cnnstr)
'        StrQuery = "SELECT * FROM DTRFinalRecords WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
'            " AND DTRKEY = '" & mDTRKEY & "' ORDER BY NoOfDay"
    
        If UseIniCutoff = False Then
            StrQuery = "SELECT * FROM DTRFinalRecords WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
                " AND DTRKEY = '" & mDTRKEY & "' ORDER BY NoOfDay"
        Else    ' Initial only
            StrQuery = "SELECT * FROM DTRFinalRecords WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
                " AND DTRDay BETWEEN " & Format(DTIniStartCutoff.VALUE, "dd") & " and " & Format(DTIniEndCutoff.VALUE, "dd") & _
                " AND DTRKEY = '" & mDTRKEY & "' ORDER BY NoOfDay"
        End If
    End If
End If

Set rsDTRRecords = ReturnRst(StrQuery, cnnstr, mstatic)

'Update also the DTR from Payroll.
StrQuery = "SELECT * FROM DTRPayrollRecords WHERE DTRKEY = '" & mDTRKEY & "'"
Set rsDTRPayroll = ReturnRst(StrQuery, cnnstr, mstatic)

lblclient(46).Caption = "Records: " & rsDTRRecords.RecordCount
    
If ActiveCtr = 0 Then
    Call DisplayBasicData("", True)
Else
    Call DisplayBasicData(CurrEmployee, True)
End If

End Sub

Private Sub GetFMSDTR()

If chkNoFilter.VALUE = 1 Then ' All Employees
    If FirstHalfFlag = True Then
        ' Get DTR records for the current month First Cutoff.
        mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "01"
    
        If UseIniCutoff = False Then
            StrQuery = "SELECT * FROM vwDTRConductors WHERE DTRDay < 16 AND YEAR(DTR_Date) = " & _
                IIf(cmbYear.Text = "", Format(Date, "yyyy"), cmbYear.Text) & _
                " AND MONTH(DTR_Date) = " & Val(IIf(cmbMonth = "", Format(Date, "mm"), Left(cmbMonth.Text, 2))) & " ORDER BY DTRID, DTR_Date "
        Else    ' use initial dtr date range
            StrQuery = "SELECT * FROM vwDTRConductors WHERE DTRDay BETWEEN " & Format(DTIniStartCutoff.VALUE, """dd""") & " and " & Format(DTIniEndCutoff.VALUE, """dd""") & " AND YEAR(DTR_Date) = " & _
                IIf(cmbYear.Text = "", Format(Date, "yyyy"), cmbYear.Text) & _
                " AND MONTH(DTR_Date) = " & Val(IIf(cmbMonth = "", Format(Date, "mm"), Left(cmbMonth.Text, 2))) & " ORDER BY DTRID, DTR_Date "
        End If
    Else
        ' Get DTR records for the current month Second Cutoff.
        mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "02"
    
        If UseIniCutoff = False Then
            StrQuery = "SELECT * FROM vwDTRConductors WHERE DTRDay > 15 AND YEAR(DTR_Date) = " & _
                IIf(cmbYear.Text = "", Format(Date, "yyyy"), cmbYear.Text) & _
                " AND MONTH(DTR_Date) = " & Val(IIf(cmbMonth = "", Format(Date, "mm"), Left(cmbMonth.Text, 2))) & " ORDER BY DTRID, DTR_Date "
        Else    ' use initial dtr date range
            StrQuery = "SELECT * FROM vwDTRConductors WHERE DTRDay BETWEEN " & Format(DTIniStartCutoff.VALUE, """dd""") & " and " & Format(DTIniEndCutoff.VALUE, """dd""") & " AND YEAR(DTR_Date) = " & _
                IIf(cmbYear.Text = "", Format(Date, "yyyy"), cmbYear.Text) & _
                " AND MONTH(DTR_Date) = " & Val(IIf(cmbMonth = "", Format(Date, "mm"), Left(cmbMonth.Text, 2))) & " ORDER BY DTRID, DTR_Date "
        End If
    End If
Else
    If FirstHalfFlag = True Then
        ' Get DTR records for the current month First Cutoff.
        mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "01"
        StrQuery = "EXEC dbo.ComputeDTRCon " & cmbYear.Text & ", " & Val(Left(cmbMonth.Text, 2)) & ", 1, '" & txtClient(9).Text & "', 1, " & IIf(ReProcessFlag = True, 1, 0)
        Call Action_Query(StrQuery, cnnstr)
    
        If UseIniCutoff = False Then
            StrQuery = "SELECT * FROM DTRFinalConductors WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
                " AND DTRKEY = '" & mDTRKEY & "' ORDER BY NoOfDay"
        Else    ' Initial only
            StrQuery = "SELECT * FROM DTRFinalConductors WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
                " AND DTRDay BETWEEN " & Format(DTIniStartCutoff.VALUE, "dd") & " and " & Format(DTIniEndCutoff.VALUE, "dd") & _
                " AND DTRKEY = '" & mDTRKEY & "' ORDER BY NoOfDay"
        End If
    Else
        ' Get DTR records for the current month Second Cutoff
        mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "02"
        StrQuery = "EXEC dbo.ComputeDTRCon " & cmbYear.Text & ", " & Val(Left(cmbMonth.Text, 2)) & ", 2, '" & txtClient(9).Text & "', 1, " & IIf(ReProcessFlag = True, 1, 0)
        Call Action_Query(StrQuery, cnnstr)
    
        If UseIniCutoff = False Then
            StrQuery = "SELECT * FROM DTRFinalConductors WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
                " AND DTRKEY = '" & mDTRKEY & "' ORDER BY NoOfDay"
        Else    ' initial only
            StrQuery = "SELECT * FROM DTRFinalConductors WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
                " AND DTRDay BETWEEN " & Format(DTIniStartCutoff.VALUE, "dd") & " and " & Format(DTIniEndCutoff.VALUE, "dd") & _
                " AND DTRKEY = '" & mDTRKEY & "' ORDER BY NoOfDay"
        End If
    End If
End If

Set rsDTRRecords = ReturnRst(StrQuery, cnnstr, mstatic)
'Update also the DTR from Payroll.
StrQuery = "SELECT * FROM DTRPayrollConductors WHERE DTRKEY = '" & mDTRKEY & "'"
Set rsDTRPayroll = ReturnRst(StrQuery, cnnstr, mstatic)

lblclient(46).Caption = "Records: " & rsDTRRecords.RecordCount
    
If ActiveCtr = 0 Then
    Call DisplayBasicData("", True)
Else
    Call DisplayBasicData(CurrEmployee, True)
End If

End Sub

Private Sub cmdFirstHalf_Click()
If cmdFirstHalf.FontBold = False Then
If cmdSecondHalf.FontBold = True Then
    cmdSecondHalf.FontBold = False
    cmdFirstHalf.FontBold = True
    FirstHalfFlag = True
    mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "01"
Else
    cmdSecondHalf.FontBold = True
    cmdFirstHalf.FontBold = False
    FirstHalfFlag = False
    mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "02"
End If
End If

End Sub

Private Sub cmdGetShift_Click()
Dim mHr As Integer
Dim mMin As String

If txtClient(8).Text = "" Then
    MsgBox "No Shift Schedule defined. Please update Shift Schedule first!", vbInformation, PayrollAlert
Else
    DTRStartTime = mID(txtClient(8).Text, InStr(1, txtClient(8).Text, "[", vbTextCompare) + 1, 5)
    DTREndTime = mID(txtClient(8).Text, InStr(1, txtClient(8).Text, "-", vbTextCompare) + 2, 5)
    ChronoTime(0).VALUE = DTRStartTime
    ChronoTime(1).VALUE = DTREndTime
    
'    txtDTR(0).Text = DTRStartTime
'    txtDTR(1).Text = DTREndTime
    
    
' Updated 1/15/2011 - Bernard

'    If Val(Left(DTRStartTime, 2)) > 12 Then
'        mHr = Val(Left(DTRStartTime, 2)) - 12
'        mMin = Mid(DTRStartTime, InStr(1, DTRStartTime, ":", vbTextCompare), 3)
'        txtDTR(0).Text = Right("00" & mHr, 2) + mMin & " PM"
'        'txtDTR(0).Text = DTRStartTime & " PM" ' time in
'    Else
'        txtDTR(0).Text = DTRStartTime & " AM" ' time in
'    End If
'    If Val(Left(DTREndTime, 2)) > 12 Then
'        mHr = Val(Left(DTREndTime, 2)) - 12
'        mMin = Mid(DTREndTime, InStr(1, DTREndTime, ":", vbTextCompare), 3)
'        txtDTR(1).Text = Right("00" & mHr, 2) + mMin & " PM"
'        'txtDTR(1).Text = DTREndTime & " PM" ' time in
'    Else
'        txtDTR(1).Text = DTREndTime & " AM" ' time in
'    End If
    txtDTR(2).Text = "8"
    txtDTR(3).Text = "0"
    txtDTR(4).Text = "0"
    txtDTR(5).Text = "0"
End If

End Sub

Private Sub cmdLeaveApplication_Click()
'Call frmLeaveDayoff.Show
'frmLeaveDayoff.SetFocus
Call ExportToExcel
End Sub

Private Sub ExportToExcel()
Dim xlApp1 As Excel.Application
Dim xlSheet As Excel.Worksheet
Dim xlBook1 As Excel.Workbook
Dim rsDTR As ADODB.Recordset
Dim prevName As String, mRange As String, SumTotal As String
Dim CurrName As String, PageBreak As Integer
Dim ExcelTemplate As String
Dim LineCtr As Integer, LineCtr2 As Integer, i As Integer, HeadCtr As Integer, RecCtr As Integer

StrQuery = "SELECT * FROM vwDTRReport WHERE ClientCode = " & txtClientCode & " AND DTRKEY = '" & mDTRKEY & "' ORDER BY ClientName, Fullname, NoofDay"
Set rsDTR = ReturnRst(StrQuery, CnnstrHR, mstatic)

If rsDTR.RecordCount = 0 Then MsgBox "No records not found!", vbInformation: Exit Sub

Set xlApp1 = New Excel.Application

ExcelTemplate = App.Path & "\Templates\DTR Report Template.xls"
Set xlBook1 = xlApp1.Workbooks.Add(ExcelTemplate)
    
xlApp1.WindowState = xlMaximized
'xlApp1.Visible = True
xlApp1.DisplayAlerts = False

'Set xlSheet = book1.Worksheets.Add
'Set xlSheet = xlBook1.Worksheets(1)

xlApp1.Sheets("DTR Records").Select

LineCtr = 1
LineCtr2 = 6
HeadCtr = 1
PageBreak = 1

On Error GoTo ExitRtn

With rsDTR
    .MoveFirst

'PrevName = IIf(IsNull(!FullName), "No Name", LTrim(RTrim(!FullName)))
prevName = "XXX"
While Not (.EOF Or .BOF)
    CurrName = IIf(IsNull(!FullName), "No Name", LTrim(RTrim(!FullName)))
    ' If there is a change in Employee Name
    If prevName <> CurrName Then    ' Paste the template
        If prevName <> "XXX" Then
            'xlApp1.Sheets("DTR Template").Select
            'LineCtr2 = 1
            HeadCtr = HeadCtr + 1   '
    
            mRange = "D" & LineCtr2 & ":D" & HeadCtr
            xlApp1.Range(mRange).Select
            xlApp1.Selection.NumberFormat = "0.00"
            
            mRange = "G" & LineCtr2 & ":N" & HeadCtr
            xlApp1.Range(mRange).Select
            xlApp1.Selection.NumberFormat = "0.00"
            
            ' Print Totals.
            mRange = Chr(68) & HeadCtr
            xlApp1.Range(mRange).Activate
            xlApp1.ActiveSheet.Cells(HeadCtr, 4).VALUE = SumTotal
            xlApp1.ActiveSheet.Cells(HeadCtr, 4).Font.Bold = True
            SumTotal = Replace(SumTotal, Chr(68), Chr(71))
            For i = 71 To 78  ' G to N
                mRange = Chr(i) & HeadCtr
                xlApp1.Range(mRange).Activate
                xlApp1.ActiveSheet.Cells(HeadCtr, i - 64).VALUE = SumTotal
                xlApp1.ActiveSheet.Cells(HeadCtr, i - 64).Font.Bold = True
                SumTotal = Replace(SumTotal, Chr(i), Chr(i + 1))
            Next i
            mRange = "B" & HeadCtr & ":C" & HeadCtr
            xlApp1.ActiveSheet.Range(mRange).VALUE = "Totals:"
            xlApp1.ActiveSheet.Range(mRange).Font.Bold = True
            xlApp1.Range(mRange).Select
            xlApp1.Selection.MergeCells = True
            
            mRange = "B" & HeadCtr + 1 & ":F" & HeadCtr + 1
            xlApp1.ActiveSheet.Range(mRange).VALUE = "Checked by:"
            xlApp1.Range(mRange).Select
            xlApp1.Selection.MergeCells = True
    
            mRange = "G" & HeadCtr + 1 & ":N" & HeadCtr + 1
            xlApp1.ActiveSheet.Range(mRange).VALUE = "Verified by:"
            xlApp1.Range(mRange).Select
            xlApp1.Selection.MergeCells = True
            
            SumTotal = ""
            'mRange = "A" & HeadCtr - 22 & ":O" & HeadCtr
            'xlApp1.ActiveSheet.Range(mRange).Select
            'xlApp1.Selection.Copy
            If PageBreak = 2 Then
                HeadCtr = HeadCtr + 3
                PageBreak = 1
            Else
                HeadCtr = HeadCtr + 3  '2
                HeadCtr = HeadCtr + 3  '2
                PageBreak = PageBreak + 1
            End If
            
            mRange = "A" & HeadCtr + 3 & ":A" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "No.", "", xlCenter, xlCenter, True, False, True, 15, , , False, 2.75)
            mRange = "B" & HeadCtr + 3 & ":B" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "DTR Day", "", xlCenter, xlCenter, True, False, True, 15, , , False, 7)
            mRange = "C" & HeadCtr + 3 & ":C" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "DTR Date", "", xlCenter, xlCenter, True, False, True, 15, , , False, 8.57)
            mRange = "D" & HeadCtr + 3 & ":D" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Paid", "Hours", xlCenter, xlCenter, True, False, True, 15, , , False, 6.43)
            mRange = "E" & HeadCtr + 3 & ":E" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Time IN", "", xlCenter, xlCenter, True, False, True, 15, , , False, 7.57)
            mRange = "F" & HeadCtr + 3 & ":F" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Time OUT", "", xlCenter, xlCenter, True, False, True, 15, , , False, 7.57)
            mRange = "G" & HeadCtr + 3 & ":G" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Work", "Hours", xlCenter, xlCenter, True, False, True, 15, , , False, 6.71)
            mRange = "H" & HeadCtr + 3 & ":H" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Abs", "", xlCenter, xlCenter, True, False, True, 15, , , False, 3.63)
            mRange = "I" & HeadCtr + 3 & ":I" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Late", "", xlCenter, xlCenter, True, False, True, 15, , , False, 3.63)
            mRange = "J" & HeadCtr + 3 & ":J" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "UDT", "", xlCenter, xlCenter, True, False, True, 15, , , False, 3.63)
            mRange = "K" & HeadCtr + 3 & ":K" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Reg OT", "", xlCenter, xlCenter, True, False, True, 15, , , False, 4.13)
            mRange = "L" & HeadCtr + 3 & ":L" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Night", "Diff", xlCenter, xlCenter, True, False, True, 15, , , False, 5.5)
            mRange = "M" & HeadCtr + 3 & ":M" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Legal", "Hol", xlCenter, xlCenter, True, False, True, 15, , , False, 4.13)
            mRange = "N" & HeadCtr + 3 & ":N" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Spl", "Hol", xlCenter, xlCenter, True, False, True, 15, , , False, 4.13)
            mRange = "O" & HeadCtr + 3 & ":O" & HeadCtr + 4
            Call PrintReportHeader(xlSheet, xlApp1, mRange, "Remarks", "", xlCenter, xlCenter, True, False, True, 15, , , False, 10)
            HeadCtr = HeadCtr - 2
            'xlApp1.Sheets("DTR Records").Select
            'mRange = "A" & HeadCtr + 1 & ":O" & HeadCtr + 24  ' Orig
            mRange = "A" & HeadCtr + 2 & ":O" & HeadCtr + 24
            'xlApp1.ActiveSheet.Range(mRange).Select
            
            'xlApp1.Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            'xlApp1.Selection.Insert Shift:=xlDown
            'Rows("1:1").Select
            'Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            'mRange = "A" & HeadCtr + 2 & ":O" & HeadCtr + 24
            xlApp1.ActiveSheet.Range(mRange).Select
            xlApp1.ActiveSheet.Range(mRange).RowHeight = 13.5
            xlApp1.ActiveSheet.Range(mRange).Font.Name = "Tahoma"
            xlApp1.ActiveSheet.Range(mRange).Font.Size = 8
            Call PlaceBox(xlSheet, xlApp1, mRange)
            HeadCtr = HeadCtr + 2
            LineCtr2 = HeadCtr
        End If
        
        mRange = "A" & HeadCtr
        xlApp1.ActiveSheet.Range(mRange).VALUE = !ClientName
        xlApp1.ActiveSheet.Range(mRange).Font.Bold = True

        'Call PrintReportHeader(xlSheet, xlApp1, mRange, !ClientName, "", xlCenter, xlCenter, True, False, True, , , , True, 30)
        mRange = "G" & HeadCtr
        xlApp1.ActiveSheet.Range(mRange).VALUE = "Position: " & !empPosition
        xlApp1.ActiveSheet.Range(mRange).Font.Bold = True
        'Call PrintReportHeader(xlSheet, xlApp1, mRange, "Position: " & !empPosition, "", xlCenter, xlCenter, True, False, True, , , , True, 30)

        mRange = "L" & HeadCtr
        xlApp1.ActiveSheet.Range(mRange).VALUE = "Shift: " & !ShiftName
        xlApp1.ActiveSheet.Range(mRange).Font.Bold = True
        
        ' Get Dayoff.
        mRange = "L" & HeadCtr + 1
        xlApp1.ActiveSheet.Range(mRange).VALUE = "Dayoff: " & !Dayoff
        xlApp1.ActiveSheet.Range(mRange).Font.Bold = True

        mRange = "A" & HeadCtr + 1
        xlApp1.ActiveSheet.Range(mRange).VALUE = "DTR Records Report [" & !DTRKEY & "]"
        xlApp1.ActiveSheet.Range(mRange).Font.Bold = True
        'Call PrintReportHeader(xlSheet, xlApp1, mRange, "DTR Records Report", "", xlCenter, xlCenter, True, False, True, , , , True, 30)
        mRange = "G" & HeadCtr + 1
        xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(IsNull(!FullName), "No Name", !FullName)
        xlApp1.ActiveSheet.Range(mRange).Font.Bold = True
        'Call PrintReportHeader(xlSheet, xlApp1, mRange, IIf(IsNull(!FullName), "No Name", !FullName), "", xlCenter, xlCenter, True, False, True, , , , True, 30)
        
        mRange = "A" & HeadCtr + 2
        xlApp1.ActiveSheet.Range(mRange).VALUE = "Printed Date & Time: " & Format(Now, "mm/dd/yyyy hh:mm:ss")
        xlApp1.ActiveSheet.Range(mRange).Font.Bold = True
        'Call PrintReportHeader(xlSheet, xlApp1, mRange, "Printed Date & Time: " & Format(Now, "mm/dd/yyyy hh:mm:ss"), "", xlCenter, xlCenter, True, False, True, , , , True, 30)
        mRange = "G" & HeadCtr + 2
        xlApp1.ActiveSheet.Range(mRange).VALUE = "DTR ID: " & !DTRID
        xlApp1.ActiveSheet.Range(mRange).Font.Bold = True
        'Call PrintReportHeader(xlSheet, xlApp1, mRange, "DTR ID: " & !DTRID, "", xlCenter, xlCenter, True, False, True, , , , True, 30)
        mRange = "L" & HeadCtr + 2
        xlApp1.ActiveSheet.Range(mRange).VALUE = "EMP NO: " & !EmpNo
        xlApp1.ActiveSheet.Range(mRange).Font.Bold = True
        'Call PrintReportHeader(xlSheet, xlApp1, mRange, "EMP NO: " & !EmpNo, "", xlCenter, xlCenter, True, False, True, , , , True, 30)
        
        prevName = CurrName
        HeadCtr = HeadCtr + 5
        LineCtr = 1
    
    End If
        
    Call DisplayMsg("Processing " & IIf(IsNull(!FullName), "No Name", !FullName) & "...")
        
    mRange = "A" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = LineCtr
    mRange = "B" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(IsNull(!NameofDay), "Nameofday", !NameofDay)
    mRange = "C" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!ActualDate, "mm/dd/yyyy")
    mRange = "D" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!ReqWorkHours, "0.00")
    mRange = "E" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(IsNull(!TimeIN), "No IN", !TimeIN)
    mRange = "F" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(IsNull(!TimeOut), "No OUT", !TimeOut)
    mRange = "G" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!WorkHours, "0.00")
    mRange = "H" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!DTRAbsent, "0.00")
    mRange = "I" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!DTRLates, "0.00")
    mRange = "J" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!DTRUnderTime, "0.00")
    mRange = "K" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!DTROT, "0.00")
    mRange = "L" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = Format(!NDiff, "0.00")
    mRange = "M" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(IsNull(!LglOT), "0.00", Format(!LglOT, "0.00"))
    mRange = "N" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(IsNull(!SplOT), "0.00", Format(!SplOT, "0.00"))
    mRange = "O" & HeadCtr
    xlApp1.ActiveSheet.Range(mRange).VALUE = IIf(IsNull(!Remarks), "", !Remarks)

    If SumTotal = "" Then
        SumTotal = "=D" & HeadCtr
    Else
        SumTotal = SumTotal & "+D" & HeadCtr
    End If

    HeadCtr = HeadCtr + 1
    LineCtr = LineCtr + 1

    i = i + 1

NextRec:

    .MoveNext

Wend

End With

mRange = Chr(68) & HeadCtr
xlApp1.Range(mRange).Activate
xlApp1.ActiveSheet.Cells(HeadCtr, 4).VALUE = SumTotal
xlApp1.ActiveSheet.Cells(HeadCtr, 4).Font.Bold = True
SumTotal = Replace(SumTotal, Chr(68), Chr(71))
For i = 71 To 78  ' G to N
    mRange = Chr(i) & HeadCtr
    xlApp1.Range(mRange).Activate
    xlApp1.ActiveSheet.Cells(HeadCtr, i - 64).VALUE = SumTotal
    xlApp1.ActiveSheet.Cells(HeadCtr, i - 64).Font.Bold = True
    SumTotal = Replace(SumTotal, Chr(i), Chr(i + 1))
Next i
mRange = "B" & HeadCtr & ":C" & HeadCtr
xlApp1.ActiveSheet.Range(mRange).VALUE = "Totals:"
xlApp1.ActiveSheet.Range(mRange).Font.Bold = True
xlApp1.Range(mRange).Select
xlApp1.Selection.MergeCells = True

mRange = "B" & HeadCtr + 1 & ":F" & HeadCtr + 1
xlApp1.ActiveSheet.Range(mRange).VALUE = "Checked by:"
xlApp1.Range(mRange).Select
xlApp1.Selection.MergeCells = True

mRange = "G" & HeadCtr + 1 & ":N" & HeadCtr + 1
xlApp1.ActiveSheet.Range(mRange).VALUE = "Verified by:"
xlApp1.Range(mRange).Select
xlApp1.Selection.MergeCells = True

mRange = "A1"
xlApp1.ActiveSheet.Range(mRange).Select

xlApp1.Visible = True
Call DisplayMsg("Processing Done.", True)

MsgBox "Done.", vbInformation, PayrollAlert

Exit Sub

ExitRtn:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
    Resume Next

End Sub

Private Sub DisplayMsg(ByVal mText As String, Optional mFlag As Boolean = False)
    lblclient(65).Caption = mText
    If mFlag = False Then
        Frame8.Visible = True
    Else
        Frame8.Visible = False
    End If
End Sub

'Private Sub cmdMarkActive_Click()
'Dim res
'
'    res = MsgBox("This will mark " & txtEmpName & " active and will be included in Payroll! Do you want to proceed?", vbQuestion + vbYesNo, PayrollAlert)
'    If res = vbYes Then ' mark the employee active
'        StrQuery = "UPDATE employeebasicsalary SET employeestatus = 1 WHERE EmployeeID = " & txtEmpNum ' txtControlNo
'        Call Action_Query(StrQuery, cnnstr)
'        Call Form_Load
'        If ListView1(1).ListItems.Count = 0 Then
'            cmdMarkActive.Enabled = False
'            InActiveCtr = 0
'        Else
'            cmdMarkActive.Enabled = True
'        End If
'    End If
'
'End Sub

'Private Sub cmdNavigation_Click(Index As Integer)
'Call Navigation(rsEmployeeBasicData, Me.cmdNavigation(Index), _
'    Me.cmdNavigation(0), Me.cmdNavigation(1), Me.cmdNavigation(2), Me.cmdNavigation(3))
'
'If AddDeductions = True Or EditDeductions = True Then
'        Call EnableDisableDeductions(False)
'        Call EnableDisableDeductionButtons("Add", "Edit", True)
'        EditDeductions = False
'        AddDeductions = False
'        cmbTypes.Enabled = True
'End If
'
''On Error Resume Next
'If rsEmployeeBasicData.RecordCount <> 0 And SSTab1(1).Tab = 0 Then ' disable the add and edit buttons
'    Call DisplayBasicData(ListView1(0).SelectedItem.Text, True)
'End If
'Call ScrollRec
'
'End Sub

Private Sub cmdPreview_Click()
Call PreviewDTR
End Sub

'Private Function IsLeapYear2(Year As Long) As Boolean
'    IsLeapYear2 = (Int(Year / 200) - (Year / 200) = 0)
'End Function

Function IsLeapYear2(ByVal sYear As String) As Boolean
   If IsDate("02/29/" & sYear) Then
      IsLeapYear2 = True
   Else
      IsLeapYear2 = False
   End If
End Function

Private Sub PreviewDTR(Optional ReportName As String = "DTR Records.rpt")
Dim rsReport As New ADODB.Recordset
Dim cApp As CRAXDRT.Application
Dim cAppl As CRAXDRT.Application
Dim cReport As CRAXDRT.Report
Dim StartDate As String
Dim EndDate As String
Dim strPay As String

On Error GoTo errHandler
    
Me.MousePointer = vbHourglass
    
If FMSFlag = False Then
    'Get DTR records.
    If chkAllDTR.VALUE = 1 Then 'all
        strPay = "EXEC GetDTRRecords 1, " & txtClientCode & ", '" & mDTRKEY & "', '" & Trim$(txtClient(9).Text) & "'"
    Else
        strPay = "EXEC GetDTRRecords 0, " & txtClientCode & ", '" & mDTRKEY & "', '" & Trim$(txtClient(9).Text) & "'"
    End If
    'ConnectToSQL
Else
    'Get DTR records for FMS.
    If chkAllDTR.VALUE = 1 Then 'all
        strPay = "EXEC GetDTRRecordsCon 1, " & txtClientCode & ", '" & mDTRKEY & "', '" & Trim$(txtClient(9).Text) & "'"
    Else
        strPay = "EXEC GetDTRRecordsCon 0, " & txtClientCode & ", '" & mDTRKEY & "', '" & Trim$(txtClient(9).Text) & "'"
    End If
End If
Call Action_Query(strPay, cnnstr)

If FMSFlag = False Then
    strPay = "SELECT * FROM tmpDTR order by Fullname, ActualDate "
Else
    strPay = "SELECT * FROM tmpDTRCon order by Fullname, ActualDate "
    If MsgBox("Do you want to print in Legal Form?", vbQuestion + vbYesNo) = vbYes Then
        frmReportViewer.Caption = "Payroll Summary Report (Legal Form)"
        ReportName = "DTR Records Conductors for Legal.rpt"
    Else
        ReportName = "DTR Records Conductors.rpt"
    End If
End If
'Set rsReport = ReturnRst(strPay, cnnstr, mstatic)
Set rsReport = ReturnRst(strPay, cnnstr, mstatic)

If rsReport.RecordCount = 0 Then
    MsgBox "No record/s selected! Cannot proceed.", vbExclamation + vbOKOnly, PayrollAlert
    rsReport.Close
    Set rsReport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
End If
    
If Right(mDTRKEY, 4) = "0301" Then  ' For 2nd half use FebReport.rpt format 0202
    If IsLeapYear2(Left(mDTRKEY, 4)) Then
        ReportName = "FebReport2020.rpt"
    Else
        ReportName = "FebReport2019.rpt"
    End If
'    ReportName = "FebReport.rpt"
End If
    
    Set cAppl = New CRAXDRT.Application

    'Set cReport = cAppl.OpenReport(App.Path & "\Reports\Payroll Register.rpt")
    Set cReport = cAppl.OpenReport(App.Path & "\Reports\" & ReportName)

    cReport.DiscardSavedData
    cReport.Database.SetDataSource rsReport

    frmReportViewer.CRViewer91.ReportSource = cReport
    frmReportViewer.CRViewer91.ViewReport
    
'    frmReportViewer.CRViewer91.Height = Me.Height - 150
'    frmReportViewer.CRViewer91.Width = Me.Width - 150
    
    frmReportViewer.CRViewer91.Zoom 100

    Do While frmReportViewer.CRViewer91.IsBusy
        DoEvents
    Loop

    frmReportViewer.Show

Set cAppl = Nothing
Set cReport = Nothing
rsReport.Close
Set rsReport = Nothing
'CON.Close

Me.MousePointer = vbDefault

'Call frmTimeKeep.SetFocus

Exit Sub
errHandler:
    
    Me.MousePointer = vbDefault
    
    MsgBox "Error : " & Err.Number & ";" & Err.Description, vbExclamation + vbOKOnly, "Error"
    Err.Clear
    
Set cAppl = Nothing
Set cReport = Nothing
    
End Sub

Private Sub cmdReprocess_Click()

If MsgBox("This will read again DTR records of " & txtClient(2).Text & " from database." & Chr(13) & "Do you want to continue?", vbQuestion + vbYesNo) = vbYes Then
    ReProcessFlag = True
    Call cmdDTRRefresh_Click
    MsgBox txtClient(2).Text & " was successfully updated.", vbInformation, PayrollAlert
End If
ReProcessFlag = False

End Sub

Private Sub cmdRefresh_Click()

    mCompany = Trim(cmbCompany) ' Updated by Bernard Feb. 29, 2008
    mEmpName = Trim(txtClient(2))
    
    Call Form_Load
    If mCompany <> "" And mCompany <> cmbCompany Then
        'cmbCompany = mCompany
    End If
    Call DisplayBasicData(mEmpName, True)
    If SSTab1(1).Tab = 1 Then ' disable the add and edit buttons
        cmdadd.Enabled = False
        cmdedit.Enabled = False
    Else
        cmdadd.Enabled = True
        cmdedit.Enabled = True
    End If
    cmbCompany.ListIndex = 0
    'If RefreshSw = True Then
        MsgBox "Data was updated!", vbInformation, PayrollAlert
    'End If

End Sub

Private Sub cmdRevert_Click()
Dim mMsg As String

    DTRStartTime = ""
    DTREndTime = ""
    'txtDTR(0).Text = "-"
    'txtDTR(1).Text = "-"
    ChronoTime(0).VALUE = "00:00:00 AM"
    ChronoTime(1).VALUE = "00:00:00 AM"
    
    txtDTR(2).Text = "0"
    txtDTR(3).Text = "0"
    txtDTR(4).Text = "0"
    txtDTR(5).Text = "0"

mMsg = "This will reset DTR entries for " & mID(cmbMonth.Text, 4, 15) & " " & lblclient(62).Caption & ", " & cmbYear.Text & " to the ff:" & Chr(13) & Chr(13) & _
    "TimeIN: = -" & Chr(13) & _
    "TimeOut: = -" & Chr(13) & _
    "Work Hours: = 0.00" & Chr(13) & _
    "DTR Absent: = 0.00" & Chr(13) & _
    "DTR Lates: = 0.00" & Chr(13) & _
    "DTR UDT: = 0.00" & Chr(13) & _
    "DTR OT: = 0.00"

If MsgBox(mMsg, vbQuestion + vbYesNo) = vbYes Then
    If FMSFlag = False Then
        StrQuery = "UPDATE DTRFinalRecords SET ReqWorkHours = " & txtDTR(2).Text & _
        ", TimeIN = '" & Format(ChronoTime(0).VALUE, "HH:MM") & "', TimeOUT = '" & Format(ChronoTime(1).VALUE, "HH:MM") & "', WorkHours = " & WorkHours & _
        ", DTRAbsent = " & txtDTR(3).Text & ", DTRLates = " & txtDTR(4).Text & ", DTRUndertime = " & txtDTR(5).Text & _
        ", DTROT = " & txtDTR(6).Text & ", NDiff = " & txtDTR(8).Text & ", StartTime = '" & DTRStartTime & "', EndTime = '" & DTREndTime & "', Remarks = 'Reset DTR.'" & _
        ", DTR_Date = ActualDate " & _
        " WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
        " AND DTRKEY = '" & mDTRKEY & "' AND Noofday = " & Val(lblclient(62).Caption)
   Else
        StrQuery = "UPDATE DTRFinalConductors SET ReqWorkHours = " & txtDTR(2).Text & _
        ", TimeIN" & LTrim(mTripNo) & " = '" & Format(ChronoTime(0).VALUE, "HH:MM") & "', TimeOUT" & LTrim(mTripNo) & " = '" & Format(ChronoTime(1).VALUE, "HH:MM") & "', TotalHrs" & LTrim(mTripNo) & " = dbo._WorkHours('" & Format(ChronoTime(0).VALUE, "HH:MM") & "', '" & Format(ChronoTime(1).VALUE, "HH:MM") & "'), WorkHours = " & WorkHours & _
        ", DTRAbsent = " & txtDTR(3).Text & ", DTRLates = " & txtDTR(4).Text & ", DTRUndertime = " & txtDTR(5).Text & _
        ", DTROT = " & txtDTR(6).Text & ", NDiff = " & txtDTR(8).Text & ", StartTime = '" & DTRStartTime & "', EndTime = '" & DTREndTime & "', Remarks = 'Reset DTR.'" & _
        ", DTR_Date = ActualDate " & _
        " WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
        " AND DTRKEY = '" & mDTRKEY & "' AND Noofday = " & Val(lblclient(62).Caption)
    End If
    
    Call Action_Query(StrQuery, cnnstr)
    Call cmdDTRRefresh_Click
    MsgBox "DTR was reset!", vbInformation, PayrollAlert
Else
    MsgBox "Changes was cancelled!", vbInformation, PayrollAlert
End If

End Sub

Private Sub cmdSaveCutoff_Click()
Dim rsCutOff As ADODB.Recordset

StrQuery = "SELECT TOP 1 * FROM DTRCutOffTable"
Set rsCutOff = ReturnRst(StrQuery, cnnstr, mstatic)
If rsCutOff.RecordCount <> 0 Then
    rsCutOff.Close
    StrQuery = "UPDATE DTRCutOffTable SET StartFirstCutoff = '" & DTStartFirstCutoff.VALUE & "', EndFirstCutoff = '" & DTEndFirstCutoff.VALUE & _
        "', StartSecondCutoff = '" & DTStartSecondCutoff.VALUE & "', EndSecondCutoff = '" & DTEndSecondCutoff.VALUE & "'" & _
        ", InitialStartCutoff = '" & DTIniStartCutoff.VALUE & "', InitialEndCutoff = '" & DTIniEndCutoff.VALUE & "'" & _
        ", UseInitialCutoff = " & chkInitCutoff.VALUE
Else
    rsCutOff.Close
    StrQuery = "INSERT INTO DTRCutOffTable (StartFirstCutoff, EndFirstCutoff, StartSecondCutoff, EndSecondCutoff) " & _
        " VALUES ('" & DTStartFirstCutoff.VALUE & "', '" & DTEndFirstCutoff.VALUE & _
        "', '" & DTStartSecondCutoff.VALUE & "', '" & DTEndSecondCutoff.VALUE & "')"
End If
Action_Query StrQuery, cnnstr

StrQuery = "SELECT TOP 1 * FROM DTRCutoffTable"
Set rsCutOff = ReturnRst(StrQuery, cnnstr, mstatic)
If rsCutOff.RecordCount <> 0 Then
    StartFirstCutoff = rsCutOff!StartFirstCutoff
    EndFirstCutoff = rsCutOff!EndFirstCutoff
    StartSecondCutoff = rsCutOff!StartSecondCutoff
    EndSecondCutoff = rsCutOff!EndSecondCutoff
    IniStartCutoff = rsCutOff!InitialStartCutoff
    IniEndCutoff = rsCutOff!InitialEndCutoff
    UseIniCutoff = rsCutOff!UseInitialCutoff
    
    StartCutoff = StartFirstCutoff  ' New cutoff variables
    EndCutoff = EndFirstCutoff      ' New cutoff variables
    
    If UseIniCutoff = False Then
        cmdFirstHalf.Caption = "First Half [" & Format(StartFirstCutoff, "dd") & " - " & Format(EndFirstCutoff, "dd") & "]"
        cmdSecondHalf.Caption = "Second Half [" & Format(StartSecondCutoff, "dd") & " - " & Format(EndSecondCutoff, "dd") & "]"
    Else ' Initial Cutoff is applied
        'If Format(IniEndCutoff, "dd") < 16 Then ' 1st half
        If FirstHalfFlag = True Then
            cmdFirstHalf.Caption = "First Half [" & Format(IniStartCutoff, "dd") & " - " & Format(IniEndCutoff, "dd") & "]"
            cmdSecondHalf.Caption = "Second Half [" & Format(StartSecondCutoff, "dd") & " - " & Format(EndSecondCutoff, "dd") & "]"
        Else    ' initial is for 2nd shift
            cmdFirstHalf.Caption = "First Half [" & Format(StartFirstCutoff, "dd") & " - " & Format(EndFirstCutoff, "dd") & "]"
            cmdSecondHalf.Caption = "Second Half [" & Format(IniStartCutoff, "dd") & " - " & Format(IniEndCutoff, "dd") & "]"
        End If
    End If
Else
    cmdFirstHalf.Caption = "First Half [1 - 15]"
    cmdSecondHalf.Caption = "Second Half [16 - 31]"
End If
rsCutOff.Close

MsgBox "New Cutoff was successfully save!", vbInformation, PayrollAlert
chkChangeCutoff.VALUE = 0
Call chkChangeCutoff_Click

End Sub

Private Sub cmdSaveDTR_Click()
Dim mMsg As String
Dim mEditedBy As String, mEdit As String
Dim mRestOT As String, mCon As String

'mEditedBy = "Edit by [" & UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID))) + "]."
mEdit = "'Edit by [" & UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID))) + "].'"

'mEditedBy = " CASE WHEN Remarks = 'Absent' OR Remarks = 'No In/Out' OR Remarks = 'Sunday' THEN Remarks ELSE " & mEdit & " END "
'mEditedBy = " CASE WHEN Remarks = 'No In/Out' OR Remarks = 'Sunday' THEN Remarks ELSE " & mEdit & " END "

mEditedBy = mEdit

'If Val(txtDTR(7).Text) <> 0 Then
'    mRestOT = ", RestDayOT = " & txtDTR(2).Text & ", Remarks = 'RestPay', LeaveType = 'R', LeaveHours = " & txtDTR(2).Text
'Else
'    mRestOT = ", Remarks = '" & mEditedBy & "'"
'End If

If chkAllNoDTR.VALUE = 1 Then
    mMsg = "Generate DTR for " & lblclient(70).Caption & " to BLANK DTR entries, " & Chr(13) & _
        "according to Shift Schedule " & txtClient(8).Text & "."
    If FMSFlag = False Then
        If chkAllRec.VALUE = 1 Then
            mCon = " AND DTRKEY = '" & mDTRKEY & "' AND NameOfDay <> 'Sunday'"
        Else
            mCon = " AND DTRKEY = '" & mDTRKEY & "' AND TimeIN = '-' AND NameOfDay <> 'Sunday'"
        End If
        StrQuery = "UPDATE DTRFinalRecords SET ReqWorkHours = " & txtDTR(2).Text & _
        ", TimeIN = '" & Format(ChronoTime(0).VALUE, "HH:MM") & "', TimeOUT = '" & Format(ChronoTime(1).VALUE, "HH:MM") & "', WorkHours = " & WorkHours & _
        ", DTRAbsent = " & txtDTR(3).Text & ", DTRLates = " & txtDTR(4).Text & ", DTRUndertime = " & txtDTR(5).Text & _
        ", DTROT = " & txtDTR(6).Text & ", NDiff = " & txtDTR(8).Text & ", StartTime = '" & DTRStartTime & "', EndTime = '" & DTREndTime & "', Remarks = " & mEditedBy & "" & _
        ", DTR_Date = ActualDate " & _
        " WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & mCon ' _
        '" AND DTRKEY = '" & mDTRKEY & "' AND TimeIN = '-' AND NameOfDay <> 'Sunday'"
    Else
        If chkAllRec.VALUE = 1 Then
            mCon = " AND DTRKEY = '" & mDTRKEY & "' AND NameOfDay <> 'Sunday'"
        Else
            mCon = " AND DTRKEY = '" & mDTRKEY & "' AND (TimeIN" & LTrim(mTripNo) & " = '-' or TimeIN" & LTrim(mTripNo) & " = '00:00:00 AM') AND NameOfDay <> 'Sunday'"
        End If
        StrQuery = "UPDATE DTRFinalConductors SET ReqWorkHours = " & txtDTR(2).Text & _
        ", TimeIN" & LTrim(mTripNo) & " = '" & Format(ChronoTime(0).VALUE, "HH:MM") & "', TimeOUT" & LTrim(mTripNo) & " = '" & Format(ChronoTime(1).VALUE, "HH:MM") & "', TotalHrs" & LTrim(mTripNo) & " = dbo._WorkHours('" & Format(ChronoTime(0).VALUE, "HH:MM") & "', '" & Format(ChronoTime(1).VALUE, "HH:MM") & "'), WorkHours = " & WorkHours & _
        ", DTRAbsent = " & txtDTR(3).Text & ", DTRLates = " & txtDTR(4).Text & ", DTRUndertime = " & txtDTR(5).Text & _
        ", DTROT = " & txtDTR(6).Text & ", NDiff = " & txtDTR(8).Text & ", StartTime = '" & DTRStartTime & "', EndTime = '" & DTREndTime & "', Remarks = " & mEditedBy & "" & _
        ", DTR_Date = ActualDate " & _
        " WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & mCon ' _
        '" AND DTRKEY = '" & mDTRKEY & "' AND (TimeIN" & LTrim(mTripNo) & " = '-' or TimeIN" & LTrim(mTripNo) & " = '00:00:00 AM') AND NameOfDay <> 'Sunday'"
    End If
Else
    ' with updates in RD
    mMsg = "Save changes in DTR for " & mID(cmbMonth.Text, 4, 15) & " " & lblclient(62).Caption & ", " & cmbYear.Text & "?"
    If FMSFlag = False Then
        StrQuery = "UPDATE DTRFinalRecords SET ReqWorkHours = " & txtDTR(2).Text & _
        ", TimeIN = '" & Format(ChronoTime(0).VALUE, "HH:MM") & "', TimeOUT = '" & Format(ChronoTime(1).VALUE, "HH:MM") & "', WorkHours = " & WorkHours & _
        ", DTRAbsent = " & txtDTR(3).Text & ", DTRLates = " & txtDTR(4).Text & ", DTRUndertime = " & txtDTR(5).Text & _
        ", DTROT = " & txtDTR(6).Text & ", NDiff = " & txtDTR(8).Text & ", StartTime = '" & DTRStartTime & "', EndTime = '" & DTREndTime & "', Remarks = " & mEditedBy & "" & _
        ", DTR_Date = ActualDate " & _
        " WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
        " AND DTRKEY = '" & mDTRKEY & "' AND Noofday = " & Val(lblclient(62).Caption)
    Else
        StrQuery = "UPDATE DTRFinalConductors SET ReqWorkHours = " & txtDTR(2).Text & _
        ", TimeIN" & LTrim(mTripNo) & " = '" & Format(ChronoTime(0).VALUE, "HH:MM") & "', TimeOUT" & LTrim(mTripNo) & " = '" & Format(ChronoTime(1).VALUE, "HH:MM") & "', TotalHrs" & LTrim(mTripNo) & " = dbo._WorkHours('" & Format(ChronoTime(0).VALUE, "HH:MM") & "', '" & Format(ChronoTime(1).VALUE, "HH:MM") & "'), WorkHours = " & WorkHours & _
        ", DTRAbsent = " & txtDTR(3).Text & ", DTRLates = " & txtDTR(4).Text & ", DTRUndertime = " & txtDTR(5).Text & _
        ", DTROT = " & txtDTR(6).Text & ", NDiff = " & txtDTR(8).Text & ", StartTime = '" & DTRStartTime & "', EndTime = '" & DTREndTime & "', Remarks = " & mEditedBy & "" & _
        ", DTR_Date = ActualDate " & _
        " WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
        " AND DTRKEY = '" & mDTRKEY & "' AND Noofday = " & Val(lblclient(62).Caption)
    End If

End If
'IIf(mRestOT = "", "", mRestOT)

If MsgBox(mMsg, vbQuestion + vbYesNo) = vbYes Then
    Call Action_Query(StrQuery, cnnstr)
    Call cmdDTRRefresh_Click
    MsgBox "Changes was saved!", vbInformation, PayrollAlert
Else
    MsgBox "Changes was cancelled!", vbInformation, PayrollAlert
End If

End Sub

Private Sub cmdSecondHalf_Click()
If cmdSecondHalf.FontBold = False Then
If cmdSecondHalf.FontBold = True Then
    cmdSecondHalf.FontBold = False
    cmdFirstHalf.FontBold = True
    FirstHalfFlag = True
    mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "01"
Else
    cmdSecondHalf.FontBold = True
    cmdFirstHalf.FontBold = False
    FirstHalfFlag = False
    mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "02"
End If
End If
End Sub

Private Sub cmdShiftSched_Click()
        
If txtClient(8).Text = cmbShiftSchedule.Text Then
    MsgBox "No changes done in shift schedule, cannot update.", vbInformation, PayrollAlert
Else
    If MsgBox("Do you to save changes in Shift Schedule?", vbQuestion + vbYesNo) = vbYes Then
        mText1 = GetDescription(rstShift, "ShiftName", cmbShiftSchedule.Text, "ShiftCode", "", False)
        StrQuery = "UPDATE HREmployees SET ShiftCode = " & Val(mText1) & " where EmployeeID = " & txtEmpNum
        If Action_Query(StrQuery, cnnstr) = True Then
            ' Refresh the listview display and other info.
            StrQuery = "SELECT * FROM vwEmployeeBasicSalary ORDER BY Fullname"
            Set rsEmployeeBasicData = ReturnRst(StrQuery, cnnstr, mstatic)
            MsgBox "Shift Schedule was changed.", vbInformation, PayrollAlert
        Else
            MsgBox "Failed to change the shift schedule.", vbInformation, PayrollAlert
        End If
        txtClient(8) = IIf(IsNull(cmbShiftSchedule), "", cmbShiftSchedule) ' get the updated shift
    End If
End If
End Sub

Private Sub cmdTrips_Click(Index As Integer)
Dim mCaption As String
        
mTripNo = Val(Left(lblclient(70).Caption, 1))
If Index = 0 Then
    If mTripNo > 1 Then
        mTripNo = mTripNo - 1
    Else
        mTripNo = 1
    End If
ElseIf mTripNo < 6 Then
    mTripNo = mTripNo + 1
Else
    mTripNo = 6
End If

Select Case mTripNo
    Case 1
        lblclient(70).Caption = "1st TRIP"
    Case 2
        lblclient(70).Caption = "2nd TRIP"
    Case 3
        lblclient(70).Caption = "3rd TRIP"
    Case 4
        lblclient(70).Caption = "4th TRIP"
    Case 5
        lblclient(70).Caption = "5th TRIP"
    Case 6
        lblclient(70).Caption = "6th TRIP"
End Select
End Sub

Private Sub Form_Activate()
Dim rsCutOff As ADODB.Recordset

StrQuery = "SELECT TOP 1 * FROM DTRCutoffTable"
Set rsCutOff = ReturnRst(StrQuery, cnnstr, mstatic)
If rsCutOff.RecordCount <> 0 Then
    StartFirstCutoff = rsCutOff!StartFirstCutoff
    DTStartFirstCutoff.VALUE = StartFirstCutoff
    
    EndFirstCutoff = rsCutOff!EndFirstCutoff
    DTEndFirstCutoff.VALUE = EndFirstCutoff
    
    StartSecondCutoff = rsCutOff!StartSecondCutoff
    DTStartSecondCutoff.VALUE = StartSecondCutoff
    
    EndSecondCutoff = rsCutOff!EndSecondCutoff
    DTEndSecondCutoff.VALUE = EndSecondCutoff
    
    IniStartCutoff = IIf(IsNull(rsCutOff!InitialStartCutoff), rsCutOff!StartFirstCutoff, rsCutOff!InitialStartCutoff)
    DTIniStartCutoff.VALUE = IniStartCutoff
    
    IniEndCutoff = IIf(IsNull(rsCutOff!InitialEndCutoff), rsCutOff!EndFirstCutoff, rsCutOff!InitialEndCutoff)
    DTIniEndCutoff.VALUE = IniEndCutoff
    
    UseIniCutoff = IIf(IsNull(rsCutOff!UseInitialCutoff), 0, rsCutOff!UseInitialCutoff)
    chkInitCutoff.VALUE = IIf(UseIniCutoff = True, 1, 0)
    
    StartCutoff = StartFirstCutoff  ' New cutoff variables
    EndCutoff = EndFirstCutoff      ' New cutoff variables
    
    If UseIniCutoff = False Then
        cmdFirstHalf.Caption = "First Half [" & Format(StartFirstCutoff, "dd") & " - " & Format(EndFirstCutoff, "dd") & "]"
        cmdSecondHalf.Caption = "Second Half [" & Format(StartSecondCutoff, "dd") & " - " & Format(EndSecondCutoff, "dd") & "]"
    Else ' Initial Cutoff is applied
        'If Format(IniEndCutoff, "dd") < 16 Then ' 1st half
        If FirstHalfFlag = True Then
            cmdFirstHalf.Caption = "First Half [" & Format(DTIniStartCutoff, "dd") & " - " & Format(DTIniEndCutoff, "dd") & "]"
            cmdSecondHalf.Caption = "Second Half [" & Format(StartSecondCutoff, "dd") & " - " & Format(EndSecondCutoff, "dd") & "]"
        Else    ' initial is for 2nd shift
            cmdFirstHalf.Caption = "First Half [" & Format(StartFirstCutoff, "dd") & " - " & Format(EndFirstCutoff, "dd") & "]"
            cmdSecondHalf.Caption = "Second Half [" & Format(DTIniStartCutoff, "dd") & " - " & Format(DTIniEndCutoff, "dd") & "]"
        End If
    End If
Else
    cmdFirstHalf.Caption = "First Half [1 - 15]"
    cmdSecondHalf.Caption = "Second Half [16 - 31]"
End If
rsCutOff.Close

cmdDTRManual.Enabled = EnableDisableDTRPrinting()
cmdLeaveApplication.Enabled = EnableDisableDTRPrinting()

End Sub

Private Sub Form_Load()
Dim mComp As String
Dim mEmpN As String
Dim y As Integer

    TimeKeepFlag = True
    FirstHalfFlag = True
    ReProcessFlag = False
    If ActiveSw = False Then
        Me.Top = (mdiMain.Height - Me.Height) / 8
        Me.Left = (mdiMain.Width - Me.Width) / 2
    End If
    FormLoadSw = False
    txtCancelSw = "0"
    mTripNo = 1
    lblclient(70).Caption = "1st TRIP"
    FMSFlag = False
    ' Get Account Name.
    cmbCompany.Enabled = True
    mUserID = UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID)))
    'mUserID = LoginName
    ActiveSw = True
    RefreshSw = False
    Call GetEmployeeBasic
    If EditFlag = True Then
        rsEmployeeBasicData.Find " clientcode = " & mRecNo
    End If
    AddFlag = False
    EditFlag = False
    Call DisableTextBoxes(False)
    Call DisableCommands(False)
    'Call EnableDisableDeductions(False)
    AddDeductions = False
    EditDeductions = False
    txtRates(16).Enabled = False
    
    chkActiveClient.Enabled = False
    If ActiveCtr = 0 Then
        Call DisplayBasicData("", True)
    Else
        Call DisplayBasicData(CurrEmployee, True)
    End If
    
    ' Get the Shift Schedule. Updated May 10, 2014.
    StrQuery = "SELECT * FROM HRShiftSchedule WHERE ShiftStatus = 1 Order by ShiftShort"
    Set rstShift = ReturnRst(StrQuery, cnnstr, mstatic)
    cmbShiftSchedule.Clear
    cmbShiftSchedule.AddItem ""
    Do While Not rstShift.EOF
        cmbShiftSchedule.AddItem rstShift!ShiftName
        cmbShiftSchedule.ItemData(cmbShiftSchedule.NewIndex) = rstShift!ShiftCode
        rstShift.MoveNext
    Loop

'    StrQuery = "SELECT * FROM DeductionTypes WHERE DeductionCode <> 0 ORDER BY DeductionDescription"
'    Set rsDeductionTypes = ReturnRst(StrQuery, CnnstrHR, mstatic)
    ' Get all Active Employee Deductions from Query.
'    StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
'    'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ProtectedFlag = 0"
'    Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
    ' Get all deductions.
'    StrQuery = "SELECT * FROM DeductionTypes WHERE DeductionCode <> 0 ORDER BY DeductionDescription"
'    Set rsGetDeductionTypes = ReturnRst(StrQuery, CnnstrHR, mstatic)

    FormLoadSw = True
    
    If empBasicSw = True Then
        mComp = Trim(mCompanyName)
        mEmpN = Trim(frmPayrollEntry.lblEmpName)
        If mComp <> "" Then
            cmbCompany = mComp
            cmbCompany.Enabled = False
        End If
        Call DisplayBasicData(mEmpN, True)
        If SSTab1(1).Tab = 1 Then ' disable the add and edit buttons
            cmdadd.Enabled = False
            cmdedit.Enabled = False
        Else
            cmdadd.Enabled = True
            cmdedit.Enabled = True
        End If
        empBasicSw = False
        frmTimeKeep.Refresh
    End If

    ' Populate month and year.
    y = 0
    cmbYear.Clear
'    Do While y <= 10
'        DoEvents
'        cmbYear.AddItem 2003 + y
'        cmbYear.ItemData(cmbYear.NewIndex) = y
'        y = y + 1
'    Loop
    'Updated 7/8/2018
    Do While y <= 5
        DoEvents
        cmbYear.AddItem Year(Date) - y
        cmbYear.ItemData(cmbYear.NewIndex) = y
        y = y + 1
    Loop
    cmbYear.ListIndex = -1
    
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
    
'    cmbMonth.AddItem "January"
'    cmbMonth.AddItem "February"
'    cmbMonth.AddItem "March"
'    cmbMonth.AddItem "April"
'    cmbMonth.AddItem "May"
'    cmbMonth.AddItem "June"
'    cmbMonth.AddItem "July"
'    cmbMonth.AddItem "August"
'    cmbMonth.AddItem "September"
'    cmbMonth.AddItem "October"
'    cmbMonth.AddItem "November"
'    cmbMonth.AddItem "December"
    
    cmbYear.Text = Format(Now, "yyyy")
    cmbMonth.ListIndex = Format(Date, "m") - 1

    cmbCompany.ListIndex = 0
'' Open the Dialogbox
'CommonDialog1.ShowOpen
'' Set the selected path
'txtProgramDIR.Text = CommonDialog1.Filename

    If FirstHalfFlag = True Then
        mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "01"
    Else
        mDTRKEY = Trim$(str(cmbYear.Text)) & Left(cmbMonth.Text, 2) & "02"
    End If

cmdDTRManual.Enabled = EnableDisableDTRPrinting()
cmdLeaveApplication.Enabled = EnableDisableDTRPrinting()

''Get Cutoff dates.
'StrQuery = "SELECT TOP 1 * FROM DTRCutoffTable"
'Set rsCutOff = ReturnRst(StrQuery, CnnstrHR, mstatic)
'If rsCutOff.RecordCount <> 0 Then
'    StartFirstCutoff = rsCutOff!StartFirstCutoff
'    EndFirstCutoff = rsCutOff!EndFirstCutoff
'    StartSecondCutoff = rsCutOff!StartSecondCutoff
'    EndSecondCutoff = rsCutOff!EndSecondCutoff
'    cmdFirstHalf.Caption = "First Half [" & Format(StartFirstCutoff, "dd") & " - " & Format(EndFirstCutoff, "dd") & "]"
'    cmdSecondHalf.Caption = "Second Half [" & Format(StartSecondCutoff, "dd") & " - " & Format(EndSecondCutoff, "dd") & "]"
'Else
'    cmdFirstHalf.Caption = "First Half [01 - 15]"
'    cmdSecondHalf.Caption = "Second Half [16 - 31]"
'End If
    
' Get Cut off settings from table.

End Sub

Private Sub GetEmployeeBasic()
Dim rsTaxCode As ADODB.Recordset

'If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'    StrQuery = "SELECT * FROM clienttable WHERE CLIENTCODE2 = 'FMS' AND clientstatus = 1 and clientname is not null ORDER BY clientname"
'Else
'    StrQuery = "SELECT * FROM clienttable WHERE clientstatus = 1 and clientname is not null ORDER BY clientname"
'End If
' Get all active company.
StrQuery = GetClientRecord(0, True, " ORDER BY ClientName")
Set rsClientData = ReturnRst(StrQuery, cnnstr, mstatic)

If rsClientData.RecordCount <> 0 Then
    Call FillComboData(cmbCompany, rsClientData)
    rsClientData.MoveFirst
    mCCode = rsClientData!ClientCode
    txtClientCode = rsClientData!ClientCode
Else
    ' No client selected disable text boxes and command buttons.
    Call DisableCommands(True)
    Call DisableTextBoxes(False)
    Exit Sub
End If

'strQuery = "SELECT a.*, (trim(b.surname) + ', ' + trim(b.firstname)) AS fullname, " & _
'            "b.EmpNo AS EmpNo2, b.sss, b.philhealth, b.pagibig, b.acctno, b.TIN, c.employmentstatusdesc, d.clientname " & _
'            "FROM ((employeebasicsalary AS a LEFT JOIN HREmployees AS b ON a.EmployeeID = b.EmployeeID) " & _
'            "LEFT JOIN HREmploymentStatus AS c ON b.EmploymentStatus = c.EmploymentStatus) " & _
'            "LEFT JOIN ClientTable AS d ON a.clientcode = d.clientcode " & _
'            " ORDER BY b.surname, b.firstname"
'strQuery = "SELECT a.*, (trim(b.surname) + ', ' + trim(b.firstname)) AS fullname, "


'' Updated by Bernard, January 27, 2006.
'StrQuery = "SELECT a.*, " & _
'            "b.EmpNo AS EmpNo2, b.sss, b.philhealth, b.pagibig, b.acctno, b.TIN, c.employmentstatusdesc, d.clientname " & _
'            "FROM ((qryEmployeebasics AS a LEFT JOIN HREmployees AS b ON a.EmployeeID = b.EmployeeID) " & _
'            "LEFT JOIN HREmploymentStatus AS c ON b.EmploymentStatus = c.EmploymentStatus) " & _
'            "LEFT JOIN ClientTable AS d ON a.clientcode = d.clientcode " & _
'            "WHERE a.payrollstatus not in ('T','W','X','Y','Z')" & _
'            " ORDER BY b.surname, b.firstname"
'

'and b.employmentstatus not in ('T','W','X','Y','Z')
'SELECT FROM ((EmployeeBasicSalary INNER JOIN HREmployees ON EmployeeBasicSalary.EmployeeID = HREmployees.EmployeeID) INNER JOIN HREmploymentStatus ON HREmployees.EmploymentStatus = HREmploymentStatus.EmploymentStatus) INNER JOIN ClientTable ON EmployeeBasicSalary.ClientCode = ClientTable.ClientCode;

StrQuery = "SELECT * FROM vwEmployeeBasicSalary ORDER BY Fullname"
Set rsEmployeeBasicData = ReturnRst(StrQuery, cnnstr, mstatic)
If rsEmployeeBasicData.RecordCount <> 0 Then
    Call FilterEmployee(rsClientData!ClientName)
Else
    ' No client selected disable text boxes and command buttons
    Call DisableCommands(True)
    Call DisableTextBoxes(False)
End If

StrQuery = "SELECT id, Taxexemptioncode FROM taxexemptioncode ORDER BY id"
Set rsTaxCode = ReturnRst(StrQuery, cnnstr, mstatic)

If rsTaxCode.RecordCount <> 0 Then
    Call FillComboData(cmbTaxCode, rsTaxCode)
End If
Set rsTaxCode = Nothing

'If rsDTRRecords.State = 1 Then
'    rsDTRRecords.Close
'End If

' Get DTR records for the current month.
'StrQuery = "SELECT * FROM vwNewDTRRecords WHERE YEAR(DTR_Date) = " & _
'    IIf(cmbYear.Text = "", Format(Date, "yyyy"), cmbYear.Text) & _
'    " AND MONTH(DTR_Date) = " & Val(IIf(cmbMonth = "", Format(Date, "mm"), Left(cmbMonth.Text, 2))) & " ORDER BY DTRID "
'Set rsDTRRecords = ReturnRst(StrQuery, cnnstr, mstatic)
'
'lblclient(46).Caption = "Records: " & rsDTRRecords.RecordCount
    
StrQuery = "EXEC dbo.ComputeDTRCutOff " & Format(Now, "yyyy") & ", 1, 1, '" & IIf(txtClient(9).Text = "Client Name", "00000001", txtClient(9).Text) & "', 1, " & IIf(ReProcessFlag = True, 1, 0) & ", '" & StartFirstCutoff & "', '" & EndFirstCutoff & "'"
Call Action_Query(StrQuery, cnnstr)
'StrQuery = "SELECT * FROM DTRNoOfDays WHERE IncludedFlag = 1 ORDER BY NoOfDay"
StrQuery = "SELECT * FROM DTRFinalRecords WHERE IncludedFlag = 1 AND DTRID = '" & Trim$(txtClient(9).Text) & "'" & _
        " AND DTRKEY = '" & mDTRKEY & "' ORDER BY NoOfDay"

Set rsDTRRecords = ReturnRst(StrQuery, cnnstr, mstatic)
   
StrQuery = "SELECT * FROM DTRPayrollRecords WHERE DTRKEY = '" & mDTRKEY & "'"
Set rsDTRPayroll = ReturnRst(StrQuery, cnnstr, mstatic)
   
End Sub

Private Sub FilterEmployee(ByVal mCompany As String)
    rsClientData.MoveFirst
    rsClientData.Find "clientname = '" & Trim(mCompany) & "'"
    If rsClientData.EOF = True Then
        mCCode = 0
        Exit Sub
    Else
        FMSFlag = IIf(IsNull(rsClientData!ClientCode2) Or rsClientData!ClientCode2 <> "FMS", False, True) ' fmsflag
        If FMSFlag = False Then
            lstCon.Visible = False
            lstDTRDisplay.Visible = True
            cmdTrips(0).Visible = False
            cmdTrips(1).Visible = False
            lblclient(70).Visible = False
        Else
            lstDTRDisplay.Visible = False
            lstCon.Visible = True
            cmdTrips(0).Visible = True
            cmdTrips(1).Visible = True
            lblclient(70).Visible = True
        End If
        mCCode = rsClientData!ClientCode
        txtClientCode = rsClientData!ClientCode
        workdaysperyr = rsClientData!WorkingDaysperYear
        noofdayspermon = rsClientData!NoofDaysperMonth
    End If
    With rsEmployeeBasicData
        .Filter = "clientcode = " & mCCode
        If .RecordCount <> 0 Then
            'If Not .EOF And Not .BOF Then
            Call FillEmployeeList(ListView1(0), rsEmployeeBasicData)
            'rsEmployeeBasicData.MoveFirst
            lblclient(10).Caption = "Total:  " & ActiveCtr
             .Filter = "clientcode = " & mCCode
        Else
            Call ClearAllData
            Call clearList
        End If
    End With
End Sub

'Private Sub FilterDeductions(ByVal mEmployeeID As Integer)
'    With rsEmpDeductions
'        .Filter = "EmployeeID = " & mEmployeeID
'        Call FillComboData(cmbTypes, rsEmpDeductions)
'        If .RecordCount = 0 Then  'If .EOF
'            Call EnableDisableDeductions("", True)
'        Else
'            cmbTypes.ListIndex = 0
'        End If
'    End With
'End Sub

Private Sub DisableCommands(ByVal mSw As Boolean)

If AddFlag = True Or EditFlag = True Then
    cmdadd.Enabled = False
    cmdedit.Enabled = False
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
    cmddelete.Enabled = False
    cmdRefresh.Enabled = False
    cmdClose.Enabled = False
Else
    cmdadd.Enabled = True
    cmdedit.Enabled = True
    cmdsave.Enabled = False
    cmdcancel.Enabled = False
    cmddelete.Enabled = True
    cmdRefresh.Enabled = True
    cmdClose.Enabled = True
End If
If mSw = True Then ' no employee records selected, enable only add and exit button
    cmdedit.Enabled = False
    cmdRefresh.Enabled = False
    cmddelete.Enabled = False
End If

End Sub

Private Sub DisableTextBoxes(ByVal TFalse As Boolean)
    
    txtClient(10).Enabled = TFalse
    For i = 5 To 6
        txtClient(i).Enabled = TFalse
    Next i
    For i = 4 To 14 '0 to 14
        If i = 10 Then
            txtRates(i).Enabled = False
        Else
            txtRates(i).Enabled = TFalse
        End If
    Next i
    For i = 0 To 4
        txtRates(i).Enabled = False
    Next i
    For i = 19 To 20
        txtRates(i).Enabled = TFalse
    Next i
    For i = 0 To 10 '7
        chkRates(i).Enabled = TFalse
    Next i
    chkRates(1).Enabled = False
    If AddFlag = True Or EditFlag = True Then
        ListView1(0).Enabled = False
        cmbCompany.Enabled = False
        cmdNavigation(0).Enabled = False
        cmdNavigation(1).Enabled = False
        cmdNavigation(2).Enabled = False
        cmdNavigation(3).Enabled = False
        'chkActiveClient.Enabled = True
    Else
        ListView1(0).Enabled = True
        cmbCompany.Enabled = True
        cmdNavigation(0).Enabled = True
        cmdNavigation(1).Enabled = True
        cmdNavigation(2).Enabled = True
        cmdNavigation(3).Enabled = True
        'chkActiveClient.Enabled = False
    End If
    'cmbTaxCode.Enabled = TFalse
    cmbPaySystem.Enabled = TFalse
    
    cmbPayBasis.Enabled = False
    mskStartContract.Enabled = False
    mskEndContract.Enabled = False
    
    mskStartSSS.Enabled = TFalse
    mskEndSSS.Enabled = TFalse
    mskStartPbig.Enabled = TFalse
    mskEndPbig.Enabled = TFalse
    mskStartVale.Enabled = TFalse
    mskEndVale.Enabled = TFalse
    
End Sub

Private Sub FillEmployeeList(ByRef mLsEmp As ListView, _
                        ByRef mRs As Recordset)
Dim litem1 As ListItem
Dim litem2 As ListItem

    InActiveCtr = 0
    ActiveCtr = 0
    mLsEmp.View = lvwReport
    ListView1(1).View = lvwReport
    mLsEmp.ListItems.Clear
    ListView1(1).ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
        If mRs!EmployeeStatus = True Then ' if active employee
            Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName))
            litem1.SubItems(1) = mRs(0)
            ActiveCtr = ActiveCtr + 1
'        Else
'            Set litem2 = ListView1(1).ListItems.Add(, , IIf(IsNull(mRs!fullname), "Blank Name", mRs!fullname))
'            litem2.SubItems(1) = IIf(IsNull(mRs!ClientName), "Blank", mRs!ClientName)
'            litem2.SubItems(2) = IIf(IsNull(mRs!EmployeeID), "Blank", mRs!EmployeeID)
'            litem2.SubItems(3) = IIf(IsNull(mRs!StartContract), "Blank", Format(mRs!StartContract, "mm/dd/yyyy"))
'            litem2.SubItems(4) = IIf(IsNull(mRs!EndContract), "Blank", Format(mRs!EndContract, "mm/dd/yyyy"))
'            InActiveCtr = InActiveCtr + 1
        End If
        mRs.MoveNext
    Wend
    'Get all Inactive employees regardless of employers.
    mRs.MoveFirst
    'mRs.Filter = "Employeestatus = false"
    mRs.Filter = "Employeestatus = 0"
    While Not (mRs.EOF Or mRs.BOF)
        Set litem2 = ListView1(1).ListItems.Add(, , 1)
        'Set litem2 = ListView1(1).ListItems.Add(, , IIf(IsNull(mRs!FullName, "Blank", mRs!FullName)))
        litem2.SubItems(1) = IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName)
        litem2.SubItems(2) = IIf(IsNull(mRs!ClientName), "Blank", mRs!ClientName)
        'litem2.SubItems(3) = IIf(IsNull(mRs!EmployeeID), "Blank", mRs!EmployeeID)
        litem2.SubItems(3) = IIf(IsNull(mRs!EmpNo2), "Blank", mRs!EmpNo2)
        litem2.SubItems(4) = IIf(IsNull(mRs!StartContract) Or mRs!StartContract = #1/1/1900#, "__/__/____", Format(mRs!StartContract, "mm/dd/yyyy"))
        litem2.SubItems(5) = IIf(IsNull(mRs!endcontract) Or mRs!endcontract = #1/1/1900#, "__/__/____", Format(mRs!endcontract, "mm/dd/yyyy"))
        litem2.SubItems(6) = IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName)
        litem2.SubItems(7) = IIf(IsNull(mRs!EmployeeID), 1, mRs!EmployeeID)
        InActiveCtr = InActiveCtr + 1
        mRs.MoveNext
    Wend
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

Private Function EnableDisableDTRPrinting() As Boolean
Dim rsCom As ADODB.Recordset

    StrQuery = "SELECT TOP 1 Agencyedittaxflag from CompanyTable"
    Set rsCom = ReturnRst(StrQuery, cnnstr, mstatic)
    
    If rsCom.EOF = False Then
        EnableDisableDTRPrinting = rsCom!AgencyEditTaxFlag
    Else
        EnableDisableDTRPrinting = False
    End If
    
    Set rsCom = Nothing
   
End Function

Private Sub cmdAdd_Click()
Dim rsNewEmp As ADODB.Recordset

    ' Check whether there are new employees entered from HR table.
    StrQuery = "SELECT surname, firstname, EmployeeID, EmpNo FROM hremployees WHERE clientcode = " & _
        txtClientCode & " and EmployeeID not in (select EmployeeID from Employeebasicsalary) ORDER BY surname, firstname"
    Set rsNewEmp = ReturnRst(StrQuery, cnnstr, mstatic)
    If rsNewEmp.RecordCount = 0 Then
        Set rsNewEmp = Nothing
        MsgBox "No new employee selected for " & cmbCompany & " from HR table!", vbInformation, PayrollAlert
        Exit Sub
    End If
    Set rsNewEmp = Nothing
    AddFlag = True
    Call ClearAllData
    Call DisableTextBoxes(True)
    Call DisableCommands(False)
    chkActiveClient = 1
    txtCancelSw = "0"
    frmSearchTimeKeep.Show 1
    
    If txtCancelSw = "1" Then
        Call CancelNewEmployee
    End If
    txtCancelSw = "0"
End Sub

Private Sub save()
Dim mstr As String
Dim mDetails As String
Dim mRemarks As String

    If txtEmpNum = "txtclientcode" Then
        txtEmpNum = 0
    End If

    mCompany = Trim(cmbCompany)
    mEmpName = Trim(txtClient(2))
    For i = 0 To 7
        txtClient(i) = Replace(txtClient(i), "'", "''")
    Next i
    For i = 0 To 14
        If txtRates(i) = "" Then
            txtRates(i) = 0#
        End If
    Next i
    If Not IsDate(mskStartContract) Or Not IsDate(mskEndContract) Then
        mskStartContract = Format(#1/1/1900#, "mm/dd/yyyy")
        mskEndContract = Format(#1/1/1900#, "mm/dd/yyyy")
    End If
    If Not IsDate(mskStartSSS) Or Not IsDate(mskEndSSS) Then
        mskStartSSS = Format(#1/1/1900#, "mm/dd/yyyy")
        mskEndSSS = Format(#1/1/1900#, "mm/dd/yyyy")
    End If
    If Not IsDate(mskStartPbig) Or Not IsDate(mskEndPbig) Then
        mskStartPbig = Format(#1/1/1900#, "mm/dd/yyyy")
        mskEndPbig = Format(#1/1/1900#, "mm/dd/yyyy")
    End If
    If Not IsDate(mskStartVale) Or Not IsDate(mskEndVale) Then
        mskStartVale = Format(#1/1/1900#, "mm/dd/yyyy")
        mskEndVale = Format(#1/1/1900#, "mm/dd/yyyy")
    End If
    If AddFlag = True Then ' Add New Record
        StrQuery = "INSERT INTO EmployeeBasicSalary (AgencyCode, ClientCode, EmployeeID, EmpNo, EmployeeStatus, StartContract, EndContract, NoofMonthsHired, MonthlyRate,DailyRate,  HourlyRate, " & _
        " PaymentBasis, PaymentSystem, WithSSSDeduction, SSSRateperDay, WithPbigDeduction, PagibigRateperDay, WithPHealthDeduction, BasicSalaryFlag3, BasicSalaryFlag4, BasicSalaryFlag5, " & _
        " SSSLoanAmount, SSSLoanPayment, StartSSSPayment, LastSSSPayment, PbigLoanAmount, PbigLoanPayment, StartPBigPayment, LastPBigPayment, ValeLoanAmount, ValeLoanPayment, StartValePayment, LastValePayment," & _
        " PHealthRateperDay, WithTaxDeduction, WithEcolaPay, EcolaAmount, With13thMonthPay, Hourly13thMonthRate, AgencyBDOCardFee, AgencyTASAesla, ClientwServiceContract, " & _
        " ClientwCompPackage, IfnoCompensationPackage, Remarks, empTaxcode, NotedBy, Updatedby, DateUpdated ) " & _
        " VALUES ( 1," & txtClientCode & ", " & txtEmpNum & ", '" & txtClient(0) & "', 1, '" & IIf(mskStartContract = "__/__/____", #1/1/1900#, Format(mskStartContract, "mm/dd/yyyy")) & "', '" & IIf(mskEndContract = "__/__/____", #1/1/1900#, Format(mskEndContract, "mm/dd/yyyy")) & "', " & txtRates(3) & ", " & txtRates(0) & ", " & txtRates(1) & ", " & txtRates(2) & _
        ", '" & Trim(cmbPayBasis) & "', '" & Trim(cmbPaySystem) & "', " & chkRates(3) & ", " & txtRates(7) & ", " & chkRates(5) & ", " & txtRates(9) & ", " & chkRates(4) & ", " & chkRates(8) & ", " & chkRates(9) & ", " & chkRates(10) & _
        ", " & txtRates(11) & ", " & txtRates(12) & ", '" & IIf(mskStartSSS = "__/__/____", #1/1/1900#, Format(mskStartSSS, "mm/dd/yyyy")) & "', '" & IIf(mskEndSSS = "__/__/____", #1/1/1900#, Format(mskEndSSS, "mm/dd/yyyy")) & "', " & txtRates(14) & ", " & txtRates(13) & ", '" & IIf(mskStartPbig = "__/__/____", #1/1/1900#, Format(mskStartPbig, "mm/dd/yyyy")) & "', '" & IIf(mskEndPbig = "__/__/____", #1/1/1900#, Format(mskEndPbig, "mm/dd/yyyy")) & "', " & txtRates(20) & ", " & txtRates(19) & ", '" & IIf(mskStartVale = "__/__/____", #1/1/1900#, Format(mskStartVale, "mm/dd/yyyy")) & "', '" & IIf(mskEndVale = "__/__/____", #1/1/1900#, Format(mskEndVale, "mm/dd/yyyy")) & "'" & _
        ", " & txtRates(8) & ", " & chkRates(0) & ", " & chkRates(1) & ", " & txtRates(4) & ", " & chkRates(2) & ", " & txtRates(5) & ", " & txtRates(6) & ", " & txtRates(10) & ", " & chkRates(6) & _
        ", " & chkRates(7) & ", '" & txtClient(5) & "', '" & txtClient(6) & "', '" & Trim(cmbTaxCode) & "', '" & txtClient(10) & "', '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & "')"
        mstr = "New record was saved!"
        mDetails = "New record was saved in Employee Timekeep."
    ElseIf EditFlag = True Then ' Save edit
        StrQuery = "UPDATE EmployeeBasicSalary SET StartContract = '" & Format(mskStartContract, "mm/dd/yyyy") & "', EndContract = '" & Format(mskEndContract, "mm/dd/yyyy") & "', NoofMonthsHired = " & txtRates(3) & ", MonthlyRate = " & txtRates(0) & ", DailyRate = " & txtRates(1) & ", HourlyRate = " & txtRates(2) & ", " & _
        "PaymentBasis = '" & Trim(cmbPayBasis) & "', PaymentSystem = '" & Trim(cmbPaySystem) & "', WithSSSDeduction = " & chkRates(3) & ", SSSRateperDay = " & txtRates(7) & ", WithPbigDeduction = " & chkRates(5) & ", PagibigRateperDay = " & txtRates(9) & ", WithPHealthDeduction = " & chkRates(4) & ", " & _
        "BasicSalaryFlag3 = " & chkRates(8) & ", BasicSalaryFlag4 = " & chkRates(9) & ", BasicSalaryFlag5 = " & chkRates(10) & ", " & _
        "SSSLoanAmount = " & txtRates(11) & ", SSSLoanPayment = " & txtRates(12) & ", StartSSSPayment = '" & IIf(mskStartSSS = "__/__/____", #1/1/1900#, Format(mskStartSSS, "mm/dd/yyyy")) & "', LastSSSPayment = '" & IIf(mskEndSSS = "__/__/____", #1/1/1900#, Format(mskEndSSS, "mm/dd/yyyy")) & "', PBigLoanAmount = " & txtRates(14) & ", PbigLoanPayment = " & txtRates(13) & ", StartPBigPayment = '" & IIf(mskStartPbig = "__/__/____", #1/1/1900#, Format(mskStartPbig, "mm/dd/yyyy")) & "', LastPbigPayment = '" & IIf(mskEndPbig = Null, #1/1/1900#, Format(mskEndPbig, "mm/dd/yyyy")) & "', " & _
        "ValeLoanAmount = " & txtRates(20) & ", ValeLoanPayment = " & txtRates(19) & ", StartValePayment = '" & IIf(mskStartVale = "__/__/____", #1/1/1900#, Format(mskStartVale, "mm/dd/yyyy")) & "', LastValePayment = '" & IIf(mskEndVale = "__/__/____", #1/1/1900#, Format(mskEndVale, "mm/dd/yyyy")) & "', " & _
        "PHealthRateperDay = " & txtRates(8) & ", WithTaxDeduction = " & chkRates(0) & ", WithEcolaPay = " & chkRates(1) & ", EcolaAmount = " & txtRates(4) & ", With13thMonthPay = " & chkRates(2) & ", Hourly13thMonthRate = " & txtRates(5) & ", AgencyBDOCardFee = " & txtRates(6) & ", AgencyTASAesla = " & txtRates(10) & ", ClientwServiceContract = " & chkRates(6) & ", " & _
        "ClientwCompPackage = " & chkRates(7) & ", IfnoCompensationPackage = '" & txtClient(5) & "', Remarks = '" & txtClient(6) & "', EmpTaxCode = '" & Trim(cmbTaxCode) & "', NotedBy = '" & txtClient(10) & "', Updatedby = '" & mUserID & "', DateUpdated = '" & Format(Now, "mm/dd/yyyy") & "'" & _
        " WHERE EmployeeID = " & txtEmpNum 'txtClient(0)
        mstr = "Edit was saved!"
        mDetails = "Edit was saved in Employee Timekeep."
    End If
    If Action_Query(StrQuery, cnnstr) = True Then
        Call SaveUserTransactions(mDetails, mRemarks, txtEmpNum, txtClient(0))
    End If
    MsgBox mstr, vbInformation, PayrollAlert
    Call Form_Load
    cmbCompany = mCompany
    Call DisplayBasicData(mEmpName, True)

End Sub

Private Sub cmdCancel_Click()
    AddFlag = False
    EditFlag = False
    Call DisableTextBoxes(False)
    Call DisableCommands(False)
    SSTab1(1).TabVisible(0) = True
    'SSTab1(1).TabVisible(1) = True
End Sub

Private Sub CancelNewEmployee()
    AddFlag = False
    EditFlag = False
    Call DisableTextBoxes(False)
    Call DisableCommands(False)
    SSTab1(1).TabVisible(0) = True
    'SSTab1(1).TabVisible(1) = True
End Sub

Private Sub cmdDelete_Click()
    Dim strID As Integer
    Dim employeeMsg As String
    
    If DeleteRights = False Then
        MsgBox "You do not have rights to Delete Record!", vbInformation, PayrollAlert
        Exit Sub
    End If
    
'    If SSTab1(1).Tab = 1 Then
        StrQuery = "DELETE FROM EmployeeBasicSalary WHERE EmployeeID = " & txtEmpNum ' txtControlNo
        employeeMsg = "This will remove " & txtClient(2).Text & " from payroll entries. Do you want to proceed?"
'    Else
'        StrQuery = "UPDATE EmployeeBasicSalary SET EmployeeStatus = 0 WHERE EmployeeID = " & txtEmpNum ' txtClient(0)
'        employeeMsg = "This will mark " & Trim(txtClient(2)) & " inactive. Do you want to proceed?"
'    End If
    If MsgBox(employeeMsg, vbQuestion + vbYesNo) = vbYes Then
        Call Action_Query(StrQuery, cnnstr)
        mCompany = Trim(cmbCompany)
        mEmpName = Trim(txtClient(2))
        'Call Form_Load
        'cmbCompany = mCompany
    
'        strQuery = "SELECT a.*, (trim(b.surname) + ', ' + trim(b.firstname)) AS fullname, " & _
'            "b.EmpNo AS EmpNo2, b.sss, b.philhealth, b.pagibig, b.acctno, b.TIN, c.employmentstatusdesc, d.clientname " & _
'            "FROM ((employeebasicsalary AS a LEFT JOIN HREmployees AS b ON a.EmployeeID = b.EmployeeID) " & _
'            "LEFT JOIN HREmploymentStatus AS c ON b.EmploymentStatus = c.EmploymentStatus) " & _
'            "LEFT JOIN ClientTable AS d ON a.clientcode = d.clientcode " & _
'            " ORDER BY b.surname, b.firstname"

'' Updated by Bernard, January 27, 2006.
'        StrQuery = "SELECT a.*, " & _
'            "b.EmpNo AS EmpNo2, b.sss, b.philhealth, b.pagibig, b.acctno, b.TIN, c.employmentstatusdesc, d.clientname " & _
'            "FROM ((qryEmployeeBasics AS a LEFT JOIN HREmployees AS b ON a.EmployeeID = b.EmployeeID) " & _
'            "LEFT JOIN HREmploymentStatus AS c ON b.EmploymentStatus = c.EmploymentStatus) " & _
'            "LEFT JOIN ClientTable AS d ON a.clientcode = d.clientcode " & _
'            "WHERE a.payrollstatus not in ('T','W','X','Y','Z')" & _
'            " ORDER BY b.surname, b.firstname"
        
        StrQuery = "SELECT * FROM vwEmployeeBasicSalary ORDER BY Fullname"
        Set rsEmployeeBasicData = Nothing
        Set rsEmployeeBasicData = ReturnRst(StrQuery, cnnstr, mstatic)
        If rsEmployeeBasicData.RecordCount <> 0 Then
            Call FilterEmployee(mCompany)
            Call DisplayBasicData(mEmpName, True)
        Else
            ' No client selected disable text boxes and command buttons
            Call DisableCommands(True)
            Call DisableTextBoxes(False)
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    
MsgBox "For updates!", vbInformation, PayrollAlert
    
'    If EditRights = False Then
'        MsgBox "You do not have rights to Edit Employee!", vbInformation, PayrollAlert
'        Exit Sub
'    End If
'
'    EditFlag = True
'    AmountSw = False
'    'mRecNo = Text1
'    Call DisableTextBoxes(True)
'    Call DisableCommands(False)
'    'SSTab1(1).TabVisible(1) = False
'    'cmbPayBasis.SetFocus
'    'cmbTaxCode.SetFocus
'    'cmbPaySystem.SetFocus
End Sub

Private Sub cmdSave_Click()
    Call save
    SSTab1(1).TabVisible(0) = True
    'SSTab1(1).TabVisible(1) = True
End Sub

Private Sub lblExit_Click()
    SSTab1(2).Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
ActiveSw = False
FormLoadSw = False
End Sub

'Private Sub ListView1_Click(Index As Integer)
'Dim res
'
'    If ActiveCtr = 0 Then
'        Call DisplayBasicData("", False)
'    Else
'        If SSTab1(1).Tab = 1 Then ' in-active tab
'            If ListView1(1).ListItems.Count = 0 Then
'                Call DisplayBasicData("", False)
'            Else
'                Call DisplayBasicData(ListView1(1).SelectedItem.Text, False)
'                CurrEmployee = ListView1(1).SelectedItem.Text
'            End If
'        Else    ' Active Tab
'            If ListView1(0).ListItems.Count = 0 Then
'                Call DisplayBasicData("", False)
'            Else
'                Call DisplayBasicData(ListView1(0).SelectedItem.Text, False)
'                CurrEmployee = ListView1(0).SelectedItem.Text
'                If AddDeductions = True Or EditDeductions = True Then
'                    'res = MsgBox("Cancel changes to deductions?", vbQuestion + vbYesNo, PayrollAlert)
'                    'If res = vbYes Then ' cancel the changes
'                        Call EnableDisableDeductions(False)
'                        Call EnableDisableDeductionButtons("Add", "Edit", True)
'                        EditDeductions = False
'                        AddDeductions = False
'                        cmbTypes.Enabled = True
'                        MsgBox "Changes to deductions was cancelled!", vbInformation, PayrollAlert
'                    'End If
'                End If
'            End If
'        End If
'    End If
'End Sub

'Private Sub ListView1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    ListView1(Index).SortKey = ColumnHeader.Index - 1
'    ListView1(Index).Sorted = True
'End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    CurrEmployee = Item.Text
    Call DisplayBasicData(Item.Text, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub clearList()
ListView1(0).ListItems.Clear
End Sub

Private Sub ClearAllData()
    Dim X As Integer
    
    For X = 0 To 6
        txtClient(X).Text = ""
    Next X
    For X = 0 To 14
        txtRates(X).Text = 0#
    Next X
'    For X = 1 To 7
'        chkRates(X).Value = 0
'    Next X
'   Set the default value.
    txtRates(4) = 0
    txtRates(9) = 0#
    txtRates(5) = 0#
    txtRates(6) = 0#
    txtRates(7) = 0#
    txtRates(10) = 0#
    txtRates(8) = 0#
    txtRates(0) = 15
    txtRates(1) = 303
    txtRates(2) = 26
    txtRates(3) = 8
    
    'cmbTaxCode.Enabled = True
    cmbPaySystem.Text = "ATM"
    cmbPayBasis.Text = "Daily"
    mskStartContract = "__/__/____"
    mskEndContract = "__/__/____"
    mskStartSSS = "__/__/____"
    mskEndSSS = "__/__/____"
    mskStartPbig = "__/__/____"
    mskEndPbig = "__/__/____"
    mskStartVale = "__/__/____"
    mskEndVale = "__/__/____"

End Sub

Private Sub lstDTRDisplay_ItemClick(ByVal Item As MSComctlLib.ListItem)

lblclient(62).Caption = Item
'txtDTR(0).Text = Item.ListSubItems.Item(5).Text     ' time in
'txtDTR(1).Text = Item.ListSubItems.Item(6).Text     ' time out

If Val(Left(Item.ListSubItems.Item(5).Text, 2)) > 24 Then
    Item.ListSubItems.Item(5).Text = "23" & mID(Item.ListSubItems.Item(5).Text, 3, 5)
End If
If Val(Left(Item.ListSubItems.Item(6).Text, 2)) > 24 Then
    Item.ListSubItems.Item(6).Text = "23" & mID(Item.ListSubItems.Item(6).Text, 3, 5)
End If

ChronoTime(0).VALUE = IIf(Item.ListSubItems.Item(5).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(5).Text)    ' time in
ChronoTime(1).VALUE = IIf(Item.ListSubItems.Item(6).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(6).Text)     ' time out

txtDTR(2).Text = Item.ListSubItems.Item(4).Text     ' hours work
txtDTR(3).Text = Item.ListSubItems.Item(8).Text     ' absent
txtDTR(4).Text = Item.ListSubItems.Item(9).Text     ' lates
txtDTR(5).Text = Item.ListSubItems.Item(10).Text    ' UDT
txtDTR(6).Text = Item.ListSubItems.Item(11).Text    ' OT
txtDTR(8).Text = Item.ListSubItems.Item(12).Text    ' Ndiff
If Trim$(Item.ListSubItems.Item(13).Text) = "RestPay" Then
    txtDTR(7).Text = Item.ListSubItems.Item(4).Text     ' RD equal to hours work
Else
    txtDTR(7).Text = "0.00"
End If
WorkHours = "8.00" 'Format(Item.ListSubItems.Item(7).Text, "###9.99") ' Work Hours
lblclient(63).Caption = "Note: " & Item.ListSubItems.Item(13).Text
If WorkHours = 0 Then
    WorkHours = 9#
End If

DTRStartTime = mID(txtClient(8).Text, InStr(1, txtClient(8).Text, "[", vbTextCompare) + 1, 5)
DTREndTime = mID(txtClient(8).Text, InStr(1, txtClient(8).Text, "-", vbTextCompare) + 2, 5)
End Sub

Private Sub MaskEdBox5_GotFocus()
    Sendkeys "{HOME}+{END}"
    MaskEdBox5.BackColor = &HFFFFC0
End Sub

Private Sub MaskEdBox5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub MaskEdbox5_LostFocus()
    MaskEdBox5.BackColor = &HFFFFFF
End Sub

Private Sub MaskEdBox6_GotFocus()
    Sendkeys "{HOME}+{END}"
    MaskEdBox6.BackColor = &HFFFFC0
End Sub

Private Sub MaskEdBox6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub MaskEdbox6_LostFocus()
    MaskEdBox6.BackColor = &HFFFFFF
End Sub

Private Sub mskEndContract_LostFocus()
    mskEndContract.BackColor = &HFFFFFF
End Sub

Private Sub mskStartContract_LostFocus()
    mskStartContract.BackColor = &HFFFFFF
End Sub

Private Sub mskStartContract_GotFocus()
    Sendkeys "{HOME}+{END}"
    mskStartContract.BackColor = &HFFFFC0
End Sub

Private Sub mskStartContract_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub mskStartContract_Validate(Cancel As Boolean)
    If mskStartContract <> "__/__/____" And Not IsDate(mskStartContract) Then
        MsgBox "Please enter a valid date!!!", vbInformation, PayrollAlert
        Cancel = True
    End If
End Sub

Private Sub mskEndContract_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub mskEndContract_GotFocus()
    Sendkeys "{HOME}+{END}"
    mskEndContract.BackColor = &HFFFFC0
End Sub

Private Sub mskEndContract_Validate(Cancel As Boolean)
    If mskEndContract <> "__/__/____" And Not IsDate(mskEndContract) Then
        MsgBox "Please enter a valid date!!!", vbInformation, PayrollAlert
        Cancel = True
    End If
End Sub

Private Sub mskStartSSS_GotFocus()
    Sendkeys "{HOME}+{END}"
    mskStartSSS.BackColor = &HFFFFC0
End Sub

Private Sub mskStartSSS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub mskStartSSS_LostFocus()
    mskStartSSS.BackColor = &HFFFFFF
End Sub

Private Sub mskStartSSS_Validate(Cancel As Boolean)
    If mskStartSSS <> "__/__/____" And Not IsDate(mskStartSSS) Then
        MsgBox "Please enter a valid date!!!", vbInformation, PayrollAlert
        Cancel = True
    End If
End Sub

Private Sub mskEndSSS_GotFocus()
    Sendkeys "{HOME}+{END}"
    mskEndSSS.BackColor = &HFFFFC0
End Sub

Private Sub mskEndSSS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub mskEndSSS_LostFocus()
    mskEndSSS.BackColor = &HFFFFFF
End Sub

Private Sub mskEndSSS_Validate(Cancel As Boolean)
    If mskEndSSS <> "__/__/____" And Not IsDate(mskEndSSS) Then
        MsgBox "Please enter a valid date!!!", vbInformation, PayrollAlert
        Cancel = True
    End If
End Sub

Private Sub mskStartPbig_GotFocus()
    Sendkeys "{HOME}+{END}"
    mskStartPbig.BackColor = &HFFFFC0
End Sub

Private Sub mskStartPbig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub mskStartPbig_Validate(Cancel As Boolean)
    If mskStartPbig <> "__/__/____" And Not IsDate(mskStartPbig) Then
        MsgBox "Please enter a valid date!!!", vbInformation, PayrollAlert
        Cancel = True
    End If
End Sub

Private Sub mskStartPbig_LostFocus()
    mskStartPbig.BackColor = &HFFFFFF
End Sub

Private Sub mskEndPbig_GotFocus()
    Sendkeys "{HOME}+{END}"
    mskEndPbig.BackColor = &HFFFFC0
End Sub

Private Sub mskEndPbig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub mskEndPbig_Validate(Cancel As Boolean)
    If mskEndPbig <> "__/__/____" And Not IsDate(mskEndPbig) Then
        MsgBox "Please enter a valid date!!!", vbInformation, PayrollAlert
        Cancel = True
    End If
End Sub

Private Sub mskEndPbig_LostFocus()
    mskEndPbig.BackColor = &HFFFFFF
End Sub

Private Sub mskEndVale_GotFocus()
    Sendkeys "{HOME}+{END}"
    mskEndVale.BackColor = &HFFFFC0
End Sub

Private Sub mskEndVale_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub mskEndVale_Validate(Cancel As Boolean)
    If mskEndVale <> "__/__/____" And Not IsDate(mskEndVale) Then
        MsgBox "Please enter a valid date!!!", vbInformation, PayrollAlert
        Cancel = True
    End If
End Sub

Private Sub mskEndVale_LostFocus()
    mskEndVale.BackColor = &HFFFFFF
End Sub

Private Sub mskStartVale_GotFocus()
    Sendkeys "{HOME}+{END}"
    mskStartVale.BackColor = &HFFFFC0
End Sub

Private Sub mskStartVale_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub mskStartVale_Validate(Cancel As Boolean)
    If mskStartVale <> "__/__/____" And Not IsDate(mskStartVale) Then
        MsgBox "Please enter a valid date!!!", vbInformation, PayrollAlert
        Cancel = True
    End If
End Sub

Private Sub mskStartVale_LostFocus()
    mskStartVale.BackColor = &HFFFFFF
End Sub

Private Sub Option1_Click()
    If Option1.VALUE = True Then
        Frame5(1).Visible = True: Frame5(0).Visible = False: Frame5(2).Visible = False
    ElseIf Option2.VALUE = True Then
        Frame5(1).Visible = False: Frame5(0).Visible = True: Frame5(2).Visible = False
    Else
        Frame5(1).Visible = False: Frame5(0).Visible = False: Frame5(2).Visible = True
    End If
End Sub

Private Sub Option2_Click()
'    If Option1.VALUE = False Then Frame5(1).Visible = False: Frame5(0).Visible = True Else Frame5(1).Visible = True: Frame5(0).Visible = False
    If Option1.VALUE = True Then
        Frame5(1).Visible = True: Frame5(0).Visible = False: Frame5(2).Visible = False
    ElseIf Option2.VALUE = True Then
        Frame5(1).Visible = False: Frame5(0).Visible = True: Frame5(2).Visible = False
    Else
        Frame5(1).Visible = False: Frame5(0).Visible = False: Frame5(2).Visible = True
    End If
End Sub

Private Sub Option3_Click()
    If Option1.VALUE = True Then
        Frame5(1).Visible = True: Frame5(0).Visible = False: Frame5(2).Visible = False
    ElseIf Option2.VALUE = True Then
        Frame5(1).Visible = False: Frame5(0).Visible = True: Frame5(2).Visible = False
    Else
        Frame5(1).Visible = False: Frame5(0).Visible = False: Frame5(2).Visible = True
    End If
End Sub

Private Sub SSTab1_Click(Index As Integer, PreviousTab As Integer)
    If SSTab1(1).Tab = 1 Then ' disable the add and edit buttons
        cmdadd.Enabled = False
        cmdedit.Enabled = False
        If InActiveCtr = 0 Then
            cmdMarkActive.Enabled = False
        Else
            cmdMarkActive.Enabled = True
        End If
    Else
        cmdadd.Enabled = True
        cmdedit.Enabled = True
    End If
End Sub

Private Sub txtClient_GotFocus(Index As Integer)
    Sendkeys "{HOME}+{END}"
    txtClient(Index).BackColor = &HFFFFC0
End Sub

Private Sub txtClient_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub txtClient_LostFocus(Index As Integer)
    txtClient(Index).BackColor = &HFFFFFF
    If Index = 0 Then
        txtClient(21) = txtClient(0)
    End If
    If Index = 5 Then
        txtClient(5) = txtClient(5)
    End If
End Sub

Private Sub DisplayBasicData(ByVal SelectedEmployee As String, Optional recSw As Boolean)
Dim mDTRID As String
Dim litem2 As ListItem
Dim Dtrctr As Integer
Dim WithDTR As Integer
Dim Ttime As Double
Dim SLTime As Double, VLTime As Double, RDTime As Double, mLeaveType As String
Dim TDays As Double
Dim TAbsent As Double
Dim TLates As Double
Dim TUndertime As Double
Dim TOvertime As Double
Dim TNightDiff As Double
Dim IndexNo As Integer

On Error Resume Next

If rsEmployeeBasicData.RecordCount = 0 Then Exit Sub

IndexNo = 0

WithDTR = 0
Ttime = 0: TDays = 0: TAbsent = 0: TLates = 0: TUndertime = 0: TOvertime = 0
TNightDiff = 0: SLTime = 0: VLTime = 0: RDTime = 0

With rsEmployeeBasicData
    If recSw = False Then
        .MoveFirst
        .Find "Fullname = '" & SelectedEmployee & "'"
    End If
    If Not .EOF And Not .BOF Then
        chkActiveClient = IIf(!EmployeeStatus = True, 1, 0)
        txtEmpNum = IIf(IsNull(!EmployeeID), 0, !EmployeeID)
        txtClient(0) = IIf(IsNull(!EmpNo2), "", !EmpNo2)
        'txtClient(10) = IIf(IsNull(!EmploymentStatusDesc), "", !EmploymentStatusDesc)
        'Updated Feb. 18, 2011.
        txtClient(10) = IIf(IsNull(!PaymentBasis), "", LTrim(RTrim(!PaymentBasis)) & "/" & Left(!EmploymentStatusDesc, 3))
        paymentbase = !PaymentBasis
        
        txtClient(2) = IIf(IsNull(!FullName), "", !FullName)
        txtClient(3) = IIf(IsNull(!NotedBy), "", !NotedBy)
        txtClient(4) = IIf(IsNull(!DateUpdated), "", Format(!DateUpdated, "mm/dd/yyyy"))
        txtClient(5) = IIf(IsNull(!IfnoCompensationPackage), "", !IfnoCompensationPackage)
        txtClient(6) = IIf(IsNull(!Remarks), "", !Remarks)
        txtClient(7) = IIf(IsNull(!Updatedby), "", !Updatedby)
        txtClient(8) = IIf(IsNull(!ShiftName), "", !ShiftName) ' get the shift Name
        txtClient(9) = IIf(IsNull(!DTRID), "", !DTRID) ' get DTR ID
        
        'Shift Schedule. Updated May 10, 2014.
        mText1 = GetDescription(rstShift, "ShiftName", txtClient(8).Text, "ShiftName", "", False)
        If mText1 = "" Then
            cmbShiftSchedule.ListIndex = -1
            cmbShiftSchedule.Text = ""
        Else
            cmbShiftSchedule.Text = mText1
        End If
        
        'Get all DTR per employee.
        Dtrctr = 0
        mDTRID = LTrim(RTrim(!DTRID))
        rsDTRPayroll.Filter = ""
        If rsDTRRecords.State = 1 Then
            'rsDTRRecords.MoveFirst
            rsDTRRecords.Filter = ""
            rsDTRRecords.Filter = "DTRID = '" & mDTRID & "'"   ' look for dtr number
            ' Check if already mark for Payroll.
            rsDTRPayroll.Filter = "DTRID = '" & mDTRID & "'"   ' look for dtr number in Payroll DTR
            cmdReprocess.Caption = "Re-Process DTR " & mDTRID
            lstDTRDisplay.ListItems.Clear
            lstCon.ListItems.Clear
            ' Change the color of display to red to signifies DTR was in Payroll.
            If rsDTRPayroll.EOF = False Then
                lstDTRDisplay.ForeColor = vbBlue
                lstCon.ForeColor = vbBlue
            Else
                lstDTRDisplay.ForeColor = vbBlack
                lstCon.ForeColor = vbBlack
            End If
        
        'While Not (rsDTRRecords.EOF Or rsDTRRecords.BOF)
        While Not (rsDTRRecords.EOF)
            Dtrctr = Dtrctr + 1
            If FMSFlag = False Then
                lstCon.Visible = False
                lstDTRDisplay.Visible = True
                Set litem2 = lstDTRDisplay.ListItems.Add(, , rsDTRRecords!NoOfDay)
            Else
                lstDTRDisplay.Visible = False
                lstCon.Visible = True
                Set litem2 = lstCon.ListItems.Add(, , rsDTRRecords!NoOfDay)
            End If
            litem2.SubItems(1) = IIf(IsNull(rsDTRRecords!NameofDay), "NO Name of Day", rsDTRRecords!NameofDay)
            litem2.SubItems(2) = IIf(IsNull(rsDTRRecords!ActualDate), "No DTR Date", Format(rsDTRRecords!ActualDate, "mm/dd/yyyy"))
            
            If FMSFlag = False Then
                litem2.SubItems(3) = IIf(IsNull(rsDTRRecords!DTRID), "NO DTR ID", rsDTRRecords!DTRID)
                litem2.SubItems(4) = IIf(IsNull(rsDTRRecords!ReqWorkHours), "NO Req Hrs", rsDTRRecords!ReqWorkHours)
                litem2.SubItems(5) = IIf(IsNull(rsDTRRecords!TimeIN), IIf(IsNull(rsDTRRecords!StartTime), "-", rsDTRRecords!StartTime), rsDTRRecords!TimeIN)
                litem2.SubItems(6) = IIf(IsNull(rsDTRRecords!TimeOut), IIf(IsNull(rsDTRRecords!EndTime), "-", rsDTRRecords!EndTime), rsDTRRecords!TimeOut)
                
                If rsDTRRecords!TimeIN <> "-" Or rsDTRRecords!ReqWorkHours <> 0 Or rsDTRRecords!ReqWorkHours = Null Then
                    WithDTR = WithDTR + 1
                    TDays = TDays + 1
                    Ttime = Ttime + IIf(IsNull(rsDTRRecords!ReqWorkHours), 0#, rsDTRRecords!ReqWorkHours)
                ElseIf rsDTRRecords!Remarks = "Absent" Then
                    TAbsent = TAbsent + 1
                End If
                IndexNo = 6
            Else
                litem2.SubItems(3) = IIf(IsNull(rsDTRRecords!ReqWorkHours), "NO Req Hrs", rsDTRRecords!ReqWorkHours)
                litem2.SubItems(4) = IIf(IsNull(rsDTRRecords!TimeIN1), "-", rsDTRRecords!TimeIN1)
                litem2.SubItems(5) = IIf(IsNull(rsDTRRecords!TimeOut1), "-", rsDTRRecords!TimeOut1)
                litem2.SubItems(6) = IIf(IsNull(rsDTRRecords!TotalHrs1), "0.00", rsDTRRecords!TotalHrs1)
                litem2.SubItems(7) = IIf(IsNull(rsDTRRecords!TimeIN2), "-", rsDTRRecords!TimeIN2)
                litem2.SubItems(8) = IIf(IsNull(rsDTRRecords!TimeOut2), "-", rsDTRRecords!TimeOut2)
                litem2.SubItems(9) = IIf(IsNull(rsDTRRecords!TotalHrs2), "0.00", rsDTRRecords!TotalHrs2)
                
                litem2.SubItems(10) = IIf(IsNull(rsDTRRecords!TimeIN3), "-", rsDTRRecords!TimeIN3)
                litem2.SubItems(11) = IIf(IsNull(rsDTRRecords!TimeOut3), "-", rsDTRRecords!TimeOut3)
                litem2.SubItems(12) = IIf(IsNull(rsDTRRecords!TotalHrs3), "0.00", rsDTRRecords!TotalHrs3)
                litem2.SubItems(13) = IIf(IsNull(rsDTRRecords!TimeIN4), "-", rsDTRRecords!TimeIN4)
                litem2.SubItems(14) = IIf(IsNull(rsDTRRecords!TimeOut4), "-", rsDTRRecords!TimeOut4)
                litem2.SubItems(15) = IIf(IsNull(rsDTRRecords!TotalHrs4), "0.00", rsDTRRecords!TotalHrs4)
                
                litem2.SubItems(16) = IIf(IsNull(rsDTRRecords!TimeIN5), "-", rsDTRRecords!TimeIN5)
                litem2.SubItems(17) = IIf(IsNull(rsDTRRecords!TimeOut5), "-", rsDTRRecords!TimeOut5)
                litem2.SubItems(18) = IIf(IsNull(rsDTRRecords!TotalHrs5), "0.00", rsDTRRecords!TotalHrs5)
                litem2.SubItems(19) = IIf(IsNull(rsDTRRecords!TimeIN6), "-", rsDTRRecords!TimeIN6)
                litem2.SubItems(20) = IIf(IsNull(rsDTRRecords!TimeOut6), "-", rsDTRRecords!TimeOut6)
                litem2.SubItems(21) = IIf(IsNull(rsDTRRecords!TotalHrs6), "0.00", rsDTRRecords!TotalHrs6)
                
                If rsDTRRecords!ReqWorkHours <> 0 Or rsDTRRecords!ReqWorkHours <> Null Then
                    WithDTR = WithDTR + 1
                    TDays = TDays + 1
                    Ttime = Ttime + IIf(IsNull(rsDTRRecords!ReqWorkHours), 0#, rsDTRRecords!ReqWorkHours)
                ElseIf rsDTRRecords!Remarks = "Absent" Then
                    TAbsent = TAbsent + 1
                End If
                IndexNo = 21
            End If
            litem2.SubItems(IndexNo + 1) = IIf(IsNull(rsDTRRecords!WorkHours), "0.00", rsDTRRecords!WorkHours)
            litem2.SubItems(IndexNo + 2) = IIf(IsNull(rsDTRRecords!DTRAbsent), "0.00", rsDTRRecords!DTRAbsent)
            litem2.SubItems(IndexNo + 3) = IIf(IsNull(rsDTRRecords!DTRLates), "0.00", rsDTRRecords!DTRLates)
            
            If FMSFlag = False Then
                TLates = TLates + IIf(IsNull(rsDTRRecords!DTRLates), 0#, rsDTRRecords!DTRLates)
                litem2.SubItems(IndexNo + 4) = IIf(IsNull(rsDTRRecords!DTRUnderTime), "0.00", rsDTRRecords!DTRUnderTime)
                TUndertime = TUndertime + IIf(IsNull(rsDTRRecords!DTRUnderTime), 0#, rsDTRRecords!DTRUnderTime)
                litem2.SubItems(IndexNo + 5) = IIf(IsNull(rsDTRRecords!DTROT), "0.00", rsDTRRecords!DTROT)
                TOvertime = TOvertime + IIf(IsNull(rsDTRRecords!DTROT), 0#, rsDTRRecords!DTROT)
                litem2.SubItems(IndexNo + 6) = IIf(IsNull(rsDTRRecords!NDiff), "0.00", rsDTRRecords!NDiff)
                TNightDiff = TNightDiff + IIf(IsNull(rsDTRRecords!NDiff), 0#, rsDTRRecords!NDiff)
                litem2.SubItems(IndexNo + 7) = IIf(IsNull(rsDTRRecords!Remarks), "", rsDTRRecords!Remarks)
            Else
                TLates = TLates + IIf(IsNull(rsDTRRecords!DTRLates), 0#, rsDTRRecords!DTRLates)
                litem2.SubItems(IndexNo + 4) = IIf(IsNull(rsDTRRecords!DTRUnderTime), "0.00", rsDTRRecords!DTRUnderTime)
                TUndertime = TUndertime + IIf(IsNull(rsDTRRecords!DTRUnderTime), 0#, rsDTRRecords!DTRUnderTime)
                litem2.SubItems(IndexNo + 5) = IIf(IsNull(rsDTRRecords!DTROT), "0.00", rsDTRRecords!DTROT)
                TOvertime = TOvertime + IIf(IsNull(rsDTRRecords!DTROT), 0#, rsDTRRecords!DTROT)
                litem2.SubItems(IndexNo + 6) = IIf(IsNull(rsDTRRecords!NDiff), "0.00", rsDTRRecords!NDiff)
                TNightDiff = TNightDiff + IIf(IsNull(rsDTRRecords!NDiff), 0#, rsDTRRecords!NDiff)
                litem2.SubItems(IndexNo + 7) = IIf(IsNull(rsDTRRecords!Remarks), "", rsDTRRecords!Remarks)
            End If
            
'            Dtrctr = Dtrctr + 1
'            Set litem2 = lstDTRDisplay.ListItems.Add(, , rsDTRRecords!NoOfDay)
'            litem2.SubItems(1) = IIf(IsNull(rsDTRRecords!NameofDay), "NO Name of Day", rsDTRRecords!NameofDay)
'            litem2.SubItems(2) = IIf(IsNull(rsDTRRecords!ActualDate), "No DTR Date", Format(rsDTRRecords!ActualDate, "mm/dd/yyyy"))
'            litem2.SubItems(3) = IIf(IsNull(rsDTRRecords!DTRID), "NO DTR ID", rsDTRRecords!DTRID)
'            litem2.SubItems(4) = IIf(IsNull(rsDTRRecords!ReqWorkHours), "NO Req Hrs", rsDTRRecords!ReqWorkHours)
'            litem2.SubItems(5) = IIf(IsNull(rsDTRRecords!TimeIN), IIf(IsNull(rsDTRRecords!StartTime), "-", rsDTRRecords!StartTime), rsDTRRecords!TimeIN)
'            litem2.SubItems(6) = IIf(IsNull(rsDTRRecords!TimeOut), IIf(IsNull(rsDTRRecords!EndTime), "-", rsDTRRecords!EndTime), rsDTRRecords!TimeOut)
'            litem2.SubItems(7) = IIf(IsNull(rsDTRRecords!WorkHours), "0.00", rsDTRRecords!WorkHours)
'            litem2.SubItems(8) = IIf(IsNull(rsDTRRecords!DTRAbsent), "0.00", rsDTRRecords!DTRAbsent)
'            litem2.SubItems(9) = IIf(IsNull(rsDTRRecords!DTRLates), "0.00", rsDTRRecords!DTRLates)
'
'            TLates = TLates + IIf(IsNull(rsDTRRecords!DTRLates), 0#, rsDTRRecords!DTRLates)
'            litem2.SubItems(10) = IIf(IsNull(rsDTRRecords!DTRUnderTime), "0.00", rsDTRRecords!DTRUnderTime)
'            TUndertime = TUndertime + IIf(IsNull(rsDTRRecords!DTRUnderTime), 0#, rsDTRRecords!DTRUnderTime)
'            litem2.SubItems(11) = IIf(IsNull(rsDTRRecords!DTROT), "0.00", rsDTRRecords!DTROT)
'            TOvertime = TOvertime + IIf(IsNull(rsDTRRecords!DTROT), 0#, rsDTRRecords!DTROT)
'            litem2.SubItems(12) = IIf(IsNull(rsDTRRecords!NDiff), "0.00", rsDTRRecords!NDiff)
'            TNightDiff = TNightDiff + IIf(IsNull(rsDTRRecords!NDiff), 0#, rsDTRRecords!NDiff)
'            litem2.SubItems(13) = IIf(IsNull(rsDTRRecords!Remarks), "", rsDTRRecords!Remarks)
'
'            If rsDTRRecords!TimeIN <> "-" Or rsDTRRecords!ReqWorkHours <> 0 Then
'                WithDTR = WithDTR + 1
'                TDays = TDays + 1
'                Ttime = Ttime + IIf(IsNull(rsDTRRecords!ReqWorkHours), 0#, rsDTRRecords!ReqWorkHours)
'            'ElseIf rsDTRRecords!NameOfDay <> "Sunday" And rsDTRRecords!Remarks <> "Dayoff" And Left(rsDTRRecords!Remarks, 4) <> "Edit" Then
'            ElseIf rsDTRRecords!Remarks = "Absent" Then
'                TAbsent = TAbsent + 1
'            End If
            
            mLeaveType = "M, P, S, V, Y"
            If InStr(mLeaveType, LTrim(RTrim(rsDTRRecords!leavetype))) <> 0 Then
                'For Maternity and Sick Leave
                If InStr("M, S", LTrim(RTrim(rsDTRRecords!leavetype))) <> 0 Then
                    SLTime = SLTime + rsDTRRecords!LeaveHours
                Else
                    VLTime = VLTime + rsDTRRecords!LeaveHours
                End If
            ElseIf LTrim(RTrim(rsDTRRecords!leavetype)) = "R" Then  ' for Rest day
                    RDTime = RDTime + rsDTRRecords!LeaveHours
            End If
            rsDTRRecords.MoveNext
        Wend
        lblclient(46).Caption = "DTR Records: " & WithDTR
        rsDTRRecords.Filter = ""

        lblTimeKeep(0).Caption = "Days: " & TDays
        lblTimeKeep(1).Caption = "Pay Hrs: " & Ttime ' - SLTime - VLTime - RDTime)
        lblTimeKeep(2).Caption = "Absent: " & TAbsent
        lblTimeKeep(3).Caption = "Lates: " & TLates
        lblTimeKeep(4).Caption = "UDT: " & TUndertime
        lblTimeKeep(5).Caption = "OT: " & TOvertime
        lblTimeKeep(6).Caption = "NDiff: " & TNightDiff
        lblTimeKeep(7).Caption = "SL: " & IIf(SLTime <> 0, (SLTime / 8), 0)
        lblTimeKeep(8).Caption = "VL: " & IIf(VLTime <> 0, (VLTime / 8), 0)
        lblTimeKeep(9).Caption = "RD: " & IIf(RDTime <> 0, RDTime, 0)
        
        End If
        
'        'Requery all Active Employee Deductions.
'        StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
'        'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ProtectedFlag = 0"
'        Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
'        Call FilterDeductions(txtEmpNum)
    End If
End With

End Sub

Private Sub txtClient_Validate(Index As Integer, Cancel As Boolean)
    If txtClient(0) = "" Then MsgBox "Client name should not be blank!", vbInformation, PayrollAlert: Cancel = True
End Sub

Private Sub txtDTR_GotFocus(Index As Integer)
    Sendkeys "{HOME}+{END}"
    txtRates(Index).BackColor = &HFFFFC0
End Sub

Private Sub txtDTR_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strValid
    strValid = "0123456789." & Chr(13) & Chr(8)
    If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub txtRates_GotFocus(Index As Integer)
    Sendkeys "{HOME}+{END}"
    txtRates(Index).BackColor = &HFFFFC0
End Sub

Private Sub txtRates_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strValid
    strValid = "0123456789." & Chr(13) & Chr(8)
    If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub txtRates_LostFocus(Index As Integer)
    txtRates(Index).BackColor = &HFFFFFF
    If AmountSw = True Then
        CurrAmount = txtRates(Index).Text
        If Index <> 3 Then
            txtRates(Index).Text = Format(CheckAmount(CurrAmount), "##0.000")
        End If
        If Index = 0 Then 'Monthly
            'Compute for Daily and hourly automatically
            txtRates(1).Text = Format((CurrAmount * 12) / workdaysperyr, "##0.00")
            txtRates(2).Text = Format(txtRates(1).Text / 8, "##0.00")
        ElseIf Index = 1 Then 'Daily
            'Compute for Monthly and hourly automatically
            txtRates(0).Text = Format(CurrAmount * noofdayspermon, "##0.00")
            txtRates(2).Text = Format(CurrAmount / 8, "##0.00")
        ElseIf Index = 2 Then 'Hourly
            'Compute for Monthly and Daily automatically
            txtRates(1).Text = Format(CurrAmount * 8, "##0.000") ' Daily rate
            txtRates(0).Text = Format(txtRates(1).Text * noofdayspermon, "##0.000")  'Monthly Rate
        ElseIf Index = 17 Or Index = 15 Or Index = 18 Then 'copy to overall total text box
            txtRates(16) = Val(txtRates(17)) + Val(txtRates(15))
        End If
    Else
        AmountSw = True
    End If
End Sub

Private Function GetDescription(ByRef rsD As ADODB.Recordset, _
    ByVal mField As String, ByVal mValue As String, ByVal mField2 As String, _
    Optional mField3 As String = "", Optional mSw As Boolean = False, _
    Optional GetDetails As Boolean = False) As String

If rsD.RecordCount <> 0 Then
    rsD.MoveFirst
Else
    GetDescription = ""
    Exit Function
End If
If mSw = False Then
    StrQuery = mField & " = '" & mValue & "'"
Else
    StrQuery = mField & " = " & mValue
End If

rsD.Find StrQuery
If rsD.EOF = True Then
    GetDescription = ""
Else
    If mField3 <> "" Then
        GetDescription = LTrim(RTrim(mID(rsD.Fields(mField2), 1, 1) & "-" & rsD.Fields(mField3)))
    Else
        GetDescription = LTrim(RTrim(rsD.Fields(mField2)))
    End If
End If
If GetDetails = True And GetDescription = "" Then
    GetDescription = "1"
End If

End Function

Private Sub chkRates_GotFocus(Index As Integer)
'    SendKeys "{HOME}+{END}"
'    chkRates(Index).BackColor = &HFFFFC0
End Sub

Private Function CheckAmount(ByRef mAmt As String) As Double
Dim i As Integer
Dim DeciSw As Boolean
Dim strAmt As String

    DeciSw = False
    If IsNull(mAmt) Or mAmt = "" Or mAmt = "0" Then
        mAmt = "0.00"
        CheckAmount = 0#
    Else
        For i = 1 To Len(mAmt)
            If DeciSw = True And mID(mAmt, i, 1) = "." Then
            Else
                strAmt = strAmt & mID(mAmt, i, 1)
            End If
            If DeciSw = False And mID(mAmt, i, 1) = "." Then
                DeciSw = True
            End If
        Next
        mAmt = IIf(strAmt = ".", 0#, Format(strAmt, "0.000"))
        If mAmt > 99999.99 Then
            MsgBox "Entries should not exceed 99,999.99. Please correct it!", vbInformation, PayrollAlert
            CheckAmount = 0#
            Exit Function
        End If
    End If
    CheckAmount = mAmt
End Function

Private Sub ScrollRec()
Dim strFindMe As String
Dim itmFound As ListItem
   
    strFindMe = txtEmpNum 'txtClient(0)
    If SSTab1(1).Tab = 1 Then
        Set itmFound = ListView1(1).FindItem(strFindMe, 2, , lvwPartial)
    Else
        Set itmFound = ListView1(0).FindItem(strFindMe, 1, , lvwPartial)
    End If
    If itmFound Is Nothing Then
    Else
        itmFound.EnsureVisible
        itmFound.SELECTED = True
    End If
End Sub

Private Sub lstCon_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim mS As Integer

lblclient(62).Caption = Item

Select Case mTripNo
    Case 1
        ChronoTime(0).VALUE = IIf(Item.ListSubItems.Item(4).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(4).Text)          ' time in
        ChronoTime(1).VALUE = IIf(Item.ListSubItems.Item(5).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(5).Text)     ' time out
        txtDTR(2).Text = Item.ListSubItems.Item(6).Text     ' hours work
    Case 2
        ChronoTime(0).VALUE = IIf(Item.ListSubItems.Item(7).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(7).Text)          ' time in
        ChronoTime(1).VALUE = IIf(Item.ListSubItems.Item(8).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(8).Text)     ' time out
        txtDTR(2).Text = Item.ListSubItems.Item(9).Text     ' hours work
    Case 3
        ChronoTime(0).VALUE = IIf(Item.ListSubItems.Item(10).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(10).Text)          ' time in
        ChronoTime(1).VALUE = IIf(Item.ListSubItems.Item(11).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(11).Text)     ' time out
        txtDTR(2).Text = Item.ListSubItems.Item(12).Text     ' hours work
    Case 4
        ChronoTime(0).VALUE = IIf(Item.ListSubItems.Item(13).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(13).Text)          ' time in
        ChronoTime(1).VALUE = IIf(Item.ListSubItems.Item(14).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(14).Text)     ' time out
        txtDTR(2).Text = Item.ListSubItems.Item(15).Text     ' hours work
    Case 5
        ChronoTime(0).VALUE = IIf(Item.ListSubItems.Item(16).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(16).Text)          ' time in
        ChronoTime(1).VALUE = IIf(Item.ListSubItems.Item(17).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(17).Text)     ' time out
        txtDTR(2).Text = Item.ListSubItems.Item(18).Text     ' hours work
    Case 6
        ChronoTime(0).VALUE = IIf(Item.ListSubItems.Item(19).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(19).Text)          ' time in
        ChronoTime(1).VALUE = IIf(Item.ListSubItems.Item(20).Text = "-", "00:00:00 AM", Item.ListSubItems.Item(20).Text)     ' time out
        txtDTR(2).Text = Item.ListSubItems.Item(21).Text     ' hours work
End Select
mS = 21

txtDTR(3).Text = Item.ListSubItems.Item(mS + 2).Text    ' absent
txtDTR(4).Text = Item.ListSubItems.Item(mS + 3).Text    ' lates
txtDTR(5).Text = Item.ListSubItems.Item(mS + 4).Text    ' UDT
txtDTR(6).Text = Item.ListSubItems.Item(mS + 5).Text    ' OT
txtDTR(8).Text = Item.ListSubItems.Item(mS + 6).Text    ' Ndiff
'If Trim$(Item.ListSubItems.Item(mS + 13).Text) = "RestPay" Then
'    txtDTR(7).Text = Item.ListSubItems.Item(4).Text     ' RD equal to hours work
'Else
'    txtDTR(7).Text = "0.00"
'End If
'WorkHours = "8.00" 'Format(Item.ListSubItems.Item(7).Text, "###9.99") ' Work Hours
lblclient(63).Caption = "Note: " & Item.ListSubItems.Item(28).Text
'If WorkHours = 0 Then
'    WorkHours = 9#
'End If

DTRStartTime = mID(txtClient(8).Text, InStr(1, txtClient(8).Text, "[", vbTextCompare) + 1, 5)
DTREndTime = mID(txtClient(8).Text, InStr(1, txtClient(8).Text, "-", vbTextCompare) + 2, 5)

End Sub

