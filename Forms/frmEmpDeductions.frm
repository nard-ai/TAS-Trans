VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEmpDeductions 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Deduction Monitoring"
   ClientHeight    =   8865
   ClientLeft      =   2595
   ClientTop       =   2415
   ClientWidth     =   12390
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmEmpDeductions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   12390
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   480
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   14631
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
      TabCaption(0)   =   "Employee Deductions"
      TabPicture(0)   =   "frmEmpDeductions.frx":058A
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
         Left            =   8325
         Picture         =   "frmEmpDeductions.frx":05A6
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   8850
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   2
         Left            =   8040
         Picture         =   "frmEmpDeductions.frx":0AD8
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "First"
         Top             =   8850
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   1
         Left            =   7755
         Picture         =   "frmEmpDeductions.frx":0BD2
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "First"
         Top             =   8850
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   0
         Left            =   7470
         Picture         =   "frmEmpDeductions.frx":0CCC
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "First"
         Top             =   8850
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
         Height          =   7860
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   12015
         Begin VB.CommandButton cmdDelHistory 
            Caption         =   "Delete Blanks"
            Height          =   255
            Left            =   9840
            TabIndex        =   152
            ToolTipText     =   "Delete History of Deductions"
            Top             =   3285
            Width           =   1815
         End
         Begin VB.CommandButton cmdList 
            Caption         =   "Change Deduction Date"
            Height          =   255
            Index           =   1
            Left            =   5400
            TabIndex        =   147
            ToolTipText     =   "View previous schedule of Deductions"
            Top             =   3285
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton cmdList 
            Caption         =   "Apply Changes for Payroll"
            Height          =   255
            Index           =   0
            Left            =   3120
            TabIndex        =   146
            ToolTipText     =   "View previous schedule of Deductions"
            Top             =   3285
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CheckBox chkPayments 
            Caption         =   "Payments Only"
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
            Left            =   720
            TabIndex        =   145
            Top             =   7200
            Value           =   1  'Checked
            Width           =   1425
         End
         Begin VB.CheckBox chkAllDed 
            Caption         =   "All Deductions"
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
            Left            =   2640
            TabIndex        =   144
            Top             =   7440
            Width           =   1305
         End
         Begin VB.CommandButton cmdSchedule 
            Caption         =   "&Refresh Schedule"
            Height          =   255
            Left            =   7920
            TabIndex        =   141
            ToolTipText     =   "View previous schedule of Deductions"
            Top             =   3285
            Width           =   1815
         End
         Begin VB.CommandButton cmdCompute 
            Caption         =   "Compute Schedule"
            Height          =   375
            Left            =   4440
            TabIndex        =   140
            Top             =   7320
            Width           =   1575
         End
         Begin VB.CheckBox chkAllEmp 
            Caption         =   "All Employees"
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
            Left            =   2640
            TabIndex        =   139
            Top             =   7200
            Width           =   1305
         End
         Begin VB.CommandButton cmdExportToExcel 
            Caption         =   "Export to Excel"
            Height          =   255
            Left            =   480
            TabIndex        =   138
            Top             =   7440
            Width           =   1815
         End
         Begin VB.CommandButton cmdCancelComputation 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   7560
            TabIndex        =   137
            Top             =   7320
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdSaveComputation 
            Caption         =   "Save Schedule"
            Height          =   375
            Left            =   6000
            TabIndex        =   136
            Top             =   7320
            Visible         =   0   'False
            Width           =   1575
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
            Left            =   1320
            TabIndex        =   28
            Text            =   "Client Name"
            Top             =   8160
            Width           =   7095
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
            Left            =   6840
            TabIndex        =   105
            Top             =   8280
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
            Left            =   5340
            TabIndex        =   99
            Text            =   "Client Name"
            Top             =   8175
            Width           =   1125
         End
         Begin VB.OptionButton Option2 
            Caption         =   "HDMFl/SSS Loans"
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
            Left            =   5040
            TabIndex        =   98
            Top             =   8280
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Ecola/SSS Settings"
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
            Left            =   3240
            TabIndex        =   97
            Top             =   8280
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtEmpNum 
            Height          =   285
            Left            =   990
            TabIndex        =   72
            Text            =   "txtclientcode"
            Top             =   180
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtClientCode 
            Height          =   285
            Left            =   1800
            TabIndex        =   63
            Text            =   "txtclientcode"
            Top             =   180
            Visible         =   0   'False
            Width           =   1095
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
            Left            =   180
            TabIndex        =   25
            Top             =   8280
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2535
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
            Left            =   7470
            TabIndex        =   32
            Text            =   "Client Name"
            Top             =   8175
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
            Left            =   1320
            TabIndex        =   29
            Text            =   "Client Name"
            Top             =   8595
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Frame Frame3 
            Height          =   1365
            Left            =   3240
            TabIndex        =   49
            Top             =   0
            Width           =   8655
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
               ItemData        =   "frmEmpDeductions.frx":11FE
               Left            =   3120
               List            =   "frmEmpDeductions.frx":120E
               Style           =   2  'Dropdown List
               TabIndex        =   129
               Top             =   330
               Width           =   1965
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
               ItemData        =   "frmEmpDeductions.frx":1236
               Left            =   4440
               List            =   "frmEmpDeductions.frx":1240
               TabIndex        =   7
               Text            =   "cmbTaxCode"
               Top             =   945
               Width           =   1635
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
               Left            =   5220
               TabIndex        =   5
               Top             =   360
               Value           =   1  'Checked
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
               Top             =   945
               Width           =   4095
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
               Left            =   1290
               MultiLine       =   -1  'True
               TabIndex        =   4
               Text            =   "frmEmpDeductions.frx":1254
               Top             =   360
               Width           =   1755
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
               Left            =   180
               TabIndex        =   3
               Text            =   "Client Name"
               Top             =   360
               Width           =   1095
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
               Left            =   3120
               TabIndex        =   130
               Top             =   150
               Width           =   1335
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Tax Code:"
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
               Left            =   4440
               TabIndex        =   53
               Top             =   675
               Width           =   825
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
               Left            =   180
               TabIndex        =   52
               Top             =   135
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
               TabIndex        =   51
               Top             =   705
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
               Left            =   1290
               TabIndex        =   50
               Top             =   150
               Width           =   1455
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
            ItemData        =   "frmEmpDeductions.frx":1260
            Left            =   180
            List            =   "frmEmpDeductions.frx":126D
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   450
            Width           =   2925
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1935
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   1080
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
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
         Begin VB.Frame Frame5 
            Height          =   1455
            Index           =   1
            Left            =   3240
            TabIndex        =   77
            Top             =   7965
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
               TabIndex        =   103
               Top             =   1320
               Visible         =   0   'False
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
               TabIndex        =   82
               Text            =   "0.00"
               Top             =   1350
               Visible         =   0   'False
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
               TabIndex        =   102
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
               TabIndex        =   101
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
               TabIndex        =   91
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
               TabIndex        =   90
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
               TabIndex        =   83
               Text            =   "0.00"
               Top             =   1350
               Visible         =   0   'False
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
               TabIndex        =   94
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
               TabIndex        =   93
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
               TabIndex        =   92
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
               TabIndex        =   89
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
               TabIndex        =   88
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
               TabIndex        =   87
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
               TabIndex        =   86
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
               TabIndex        =   85
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
               TabIndex        =   84
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
               TabIndex        =   96
               Top             =   1395
               Visible         =   0   'False
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
               TabIndex        =   95
               Top             =   1395
               Visible         =   0   'False
               Width           =   855
            End
         End
         Begin VB.Frame Frame5 
            Height          =   1695
            Index           =   0
            Left            =   3240
            TabIndex        =   57
            Top             =   8085
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
               TabIndex        =   21
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
               TabIndex        =   22
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
               TabIndex        =   17
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
               TabIndex        =   18
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
               TabIndex        =   14
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
               TabIndex        =   13
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
               TabIndex        =   15
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
               TabIndex        =   16
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
               TabIndex        =   19
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
               TabIndex        =   20
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
               TabIndex        =   23
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
               TabIndex        =   24
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
               TabIndex        =   81
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
               TabIndex        =   80
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
               TabIndex        =   79
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
               TabIndex        =   78
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
               TabIndex        =   71
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
               TabIndex        =   70
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
               TabIndex        =   69
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
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
               TabIndex        =   64
               Top             =   165
               Width           =   825
            End
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
            Left            =   1890
            TabIndex        =   27
            Text            =   "Client Name"
            Top             =   8160
            Visible         =   0   'False
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
            Left            =   180
            TabIndex        =   26
            Top             =   8055
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2985
         End
         Begin VB.Frame Frame1 
            Height          =   1455
            Left            =   3240
            TabIndex        =   42
            Top             =   8025
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
               TabIndex        =   131
               Text            =   "15"
               Top             =   360
               Width           =   1005
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
               ItemData        =   "frmEmpDeductions.frx":1290
               Left            =   120
               List            =   "frmEmpDeductions.frx":12A3
               Style           =   2  'Dropdown List
               TabIndex        =   127
               Top             =   1020
               Width           =   1245
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
               TabIndex        =   12
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
               TabIndex        =   9
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
               TabIndex        =   8
               Text            =   "15"
               Top             =   405
               Width           =   1005
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
               TabIndex        =   10
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
               TabIndex        =   11
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
               Left            =   120
               TabIndex        =   128
               Top             =   840
               Width           =   1455
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
               TabIndex        =   48
               Top             =   765
               Width           =   1275
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
               TabIndex        =   47
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
               TabIndex        =   46
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
               TabIndex        =   45
               Top             =   180
               Width           =   1035
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
               TabIndex        =   44
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
               TabIndex        =   43
               Top             =   765
               Width           =   915
            End
         End
         Begin VB.Frame Frame5 
            Height          =   1695
            Index           =   2
            Left            =   3240
            TabIndex        =   106
            Top             =   1320
            Width           =   8655
            Begin VB.ComboBox cmbDedLabel 
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
               ItemData        =   "frmEmpDeductions.frx":12C4
               Left            =   6120
               List            =   "frmEmpDeductions.frx":12C6
               TabIndex        =   149
               Top             =   720
               Width           =   2445
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "Show Deduction History"
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
               Left            =   5880
               TabIndex        =   151
               Top             =   1030
               Width           =   2055
            End
            Begin VB.ComboBox cmbDedHeader 
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
               ItemData        =   "frmEmpDeductions.frx":12C8
               Left            =   5280
               List            =   "frmEmpDeductions.frx":12CA
               TabIndex        =   148
               Top             =   180
               Visible         =   0   'False
               Width           =   3285
            End
            Begin VB.TextBox txtRates 
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
               Index           =   23
               Left            =   3240
               TabIndex        =   120
               Top             =   1320
               Width           =   2355
            End
            Begin VB.TextBox txtDedID 
               Height          =   285
               Left            =   8040
               TabIndex        =   142
               Text            =   "txtDedID"
               Top             =   960
               Visible         =   0   'False
               Width           =   495
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
               Index           =   22
               Left            =   5040
               TabIndex        =   132
               Text            =   "0.00"
               Top             =   720
               Width           =   915
            End
            Begin VB.CommandButton cmdDelDed 
               Caption         =   "Del"
               Height          =   255
               Left            =   7560
               TabIndex        =   125
               Top             =   1320
               Width           =   855
            End
            Begin VB.CommandButton cmdEditDed 
               Caption         =   "Edit"
               Height          =   255
               Left            =   6720
               TabIndex        =   124
               Top             =   1320
               Width           =   855
            End
            Begin VB.CommandButton cmdAddDed 
               Caption         =   "New"
               Height          =   255
               Left            =   5880
               TabIndex        =   123
               Top             =   1320
               Width           =   855
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
               TabIndex        =   122
               Top             =   1680
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
               Left            =   3960
               TabIndex        =   117
               Text            =   "1"
               Top             =   720
               Width           =   915
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
               ItemData        =   "frmEmpDeductions.frx":12CC
               Left            =   1470
               List            =   "frmEmpDeductions.frx":12D9
               Style           =   2  'Dropdown List
               TabIndex        =   112
               Top             =   180
               Width           =   3285
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
               Left            =   990
               TabIndex        =   114
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
               TabIndex        =   113
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
               Left            =   2910
               TabIndex        =   116
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
               Index           =   15
               Left            =   1920
               TabIndex        =   115
               Text            =   "0.00"
               Top             =   720
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
               Left            =   90
               TabIndex        =   118
               Top             =   1305
               Width           =   1470
               _ExtentX        =   2593
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
               Left            =   1560
               TabIndex        =   119
               Top             =   1305
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               ForeColor       =   12582912
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
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Deduction Label"
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
               Left            =   6120
               TabIndex        =   155
               Top             =   490
               Visible         =   0   'False
               Width           =   1635
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
               Left            =   90
               TabIndex        =   154
               Top             =   1110
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
               Left            =   1560
               TabIndex        =   153
               Top             =   1110
               Width           =   735
            End
            Begin VB.Label lblclient 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   17
               Left            =   4800
               TabIndex        =   150
               Top             =   240
               Visible         =   0   'False
               Width           =   465
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
               Index           =   15
               Left            =   3240
               TabIndex        =   143
               Top             =   1110
               Width           =   975
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
               Index           =   37
               Left            =   5190
               TabIndex        =   133
               Top             =   525
               Width           =   795
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
               TabIndex        =   126
               Top             =   1800
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
               Left            =   4080
               TabIndex        =   121
               Top             =   525
               Width           =   795
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Deduction"
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
               Left            =   1080
               TabIndex        =   110
               Top             =   540
               Width           =   915
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Loan Amt"
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
               TabIndex        =   109
               Top             =   540
               Width           =   825
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
               Index           =   26
               Left            =   2940
               TabIndex        =   108
               Top             =   525
               Width           =   795
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
               Left            =   2130
               TabIndex        =   107
               Top             =   525
               Width           =   585
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
               TabIndex        =   111
               Top             =   240
               Width           =   1545
            End
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   3495
            Index           =   2
            Left            =   120
            TabIndex        =   134
            Top             =   3600
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   6165
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
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   3
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Deduction Type"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Count"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Loan Total"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Deduction"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Payments"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Deduction Date"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Remarks"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Width           =   3
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Object.Width           =   3
            EndProperty
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Schedules:"
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
            Left            =   240
            TabIndex        =   135
            Top             =   3360
            Width           =   1485
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
            TabIndex        =   100
            Top             =   8055
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
            Left            =   2250
            TabIndex        =   56
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
            TabIndex        =   55
            Top             =   8055
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
            Left            =   480
            TabIndex        =   54
            Top             =   8115
            Visible         =   0   'False
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
            TabIndex        =   41
            Top             =   840
            Width           =   1365
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Client Name:"
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
            TabIndex        =   40
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
            Left            =   240
            TabIndex        =   39
            Top             =   7920
            Visible         =   0   'False
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
            Left            =   480
            TabIndex        =   38
            Top             =   8280
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
            Picture         =   "frmEmpDeductions.frx":12FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   12390
      _ExtentX        =   21855
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdDeductionTypes 
         Height          =   405
         Left            =   3840
         Picture         =   "frmEmpDeductions.frx":135A
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Deduction Types"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         Enabled         =   0   'False
         Height          =   405
         Left            =   2745
         Picture         =   "frmEmpDeductions.frx":179C
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Delete"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   2175
         Picture         =   "frmEmpDeductions.frx":189E
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Refresh"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Enabled         =   0   'False
         Height          =   405
         Left            =   3360
         Picture         =   "frmEmpDeductions.frx":19E8
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "New"
         Top             =   0
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1080
         Picture         =   "frmEmpDeductions.frx":1F1A
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Save"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1620
         Picture         =   "frmEmpDeductions.frx":244C
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Cancel"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdedit 
         Enabled         =   0   'False
         Height          =   405
         Left            =   540
         Picture         =   "frmEmpDeductions.frx":297E
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Edit"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdClose 
         Height          =   405
         Left            =   0
         Picture         =   "frmEmpDeductions.frx":2AC8
         Style           =   1  'Graphical
         TabIndex        =   58
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
      TabIndex        =   36
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
      TabIndex        =   33
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
Attribute VB_Name = "frmEmpDeductions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public i As Integer
Dim rsEmployeeBasicData As ADODB.Recordset
Dim rsInactiveEmployee As ADODB.Recordset
Dim rsClientData As ADODB.Recordset
Dim rsDeductionTypes As ADODB.Recordset
Dim rsEmpDeductions As ADODB.Recordset
Dim rsGetDeductionTypes As ADODB.Recordset

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
Dim CurrAmount As String        ' holds the current value of salaries
Dim AmountSw As Boolean         ' variable for basic salary entries
Dim ActiveSw As Boolean
Dim FormLoadSw As Boolean
Dim ActiveDedFlag As Integer    ' signifies active or inactive deductions details

Dim workdaysperyr As Integer
Dim noofdayspermon As Integer
Dim mPrevIndex As Integer

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
Dim RecordSeq As Long

Private Sub chkRates_Click(Index As Integer)
If Index = 11 Then
    If chkRates(11).VALUE = 0 Then
        ActiveDedFlag = 1
        cmdAddDed.Enabled = True
        cmdEditDed.Enabled = True
        cmdDelDed.Enabled = True
        cmdDelHistory.Caption = "Delete Blanks"
    Else
        ActiveDedFlag = 0
        cmdAddDed.Enabled = False
        cmdEditDed.Enabled = False
        cmdDelDed.Enabled = False
        cmdSchedule.Enabled = True
        cmdDelHistory.Caption = "Delete History"
    End If
    Call FilterDeductions(txtEmpNum)
End If
End Sub

Private Sub cmbTaxCode_GotFocus()
    Sendkeys "{HOME}+{END}"
End Sub

Private Sub cmbTaxCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub cmbCompany_GotFocus()
    Sendkeys "{HOME}+{END}"
End Sub

Private Sub cmbPayBasis_GotFocus()
'SendKeys "{HOME}+{END}"
cmbPayBasis.Text = paymentbase
If cmbPayBasis.Text = "Commission" Then
    lblclient(43).Caption = "Commission Rate"
Else
    lblclient(43).Caption = "Hourly Rate"
End If
End Sub

Private Sub cmbPayBasis_LostFocus()
If cmbPayBasis.Text = "Daily" Then
    CurrAmount = txtRates(1).Text
    txtRates(0).Enabled = False
    txtRates(1).Enabled = True
    txtRates(2).Enabled = False
    txtRates(1).SetFocus
ElseIf cmbPayBasis.Text = "Hourly" Then
    CurrAmount = txtRates(2).Text
    txtRates(0).Enabled = False
    txtRates(1).Enabled = False
    txtRates(2).Enabled = True
    txtRates(2).SetFocus
Else ' Monthly
    CurrAmount = txtRates(0).Text
    txtRates(0).Enabled = True
    txtRates(1).Enabled = False
    txtRates(2).Enabled = False
    txtRates(0).SetFocus
End If
'SendKeys "{HOME}+{END}"
End Sub

Private Sub cmbPayBasis_KeyPress(KeyAscii As Integer)
'    If cmbPayBasis.Text = "Daily" Then
'        txtRates(0).Enabled = False
'        txtRates(1).Enabled = True
'    Else
'        txtRates(0).Enabled = True
'        txtRates(1).Enabled = False
'    End If
'    txtRates(2).Enabled = False
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub cmbPaySystem_GotFocus()
    'SendKeys "{HOME}+{END}"
End Sub

Private Sub cmbPaySystem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub cmbCompany_Click()
Call AutoSizeDropDownWidth(cmbCompany)
Call FilterEmployee(cmbCompany.Text)
If ActiveCtr = 0 Then
    Call DisplayBasicData("", False)
Else
    If ListView1(0).ListItems.Count > 0 Then
        Call DisplayBasicData(ListView1(0).SelectedItem.Text, False)
    End If
End If
End Sub

Private Sub EnableDisableDeductions(ByVal mTrue As Variant, Optional mBlankSw As Boolean = False)
If mBlankSw = False Then
    chkRates(12).Enabled = mTrue
    txtRates(17).Enabled = mTrue
    txtRates(18).Enabled = mTrue
    txtRates(15).Enabled = mTrue
    txtRates(16).Enabled = False
    txtRates(21).Enabled = mTrue
    txtRates(23).Enabled = mTrue
    MaskEdBox5.Enabled = mTrue
    MaskEdBox6.Enabled = mTrue
Else
    'chkRates(11).VALUE = 1
    chkRates(12).VALUE = 0
    txtRates(17).Text = mTrue
    txtRates(18).Text = mTrue
    txtRates(15).Text = mTrue
    'txtRates(16).Text = mTrue
    txtRates(21).Text = mTrue
    txtRates(23).Text = mTrue
    MaskEdBox5.Text = "__/__/____"
    MaskEdBox6.Text = "__/__/____"
End If
End Sub

Private Sub EnableDisableDeductionButtons(ByVal mAdd As String, ByVal mEdit As String, ByVal mDel As Boolean)
cmdAddDed.Caption = mAdd
cmdEditDed.Caption = mEdit
cmdDelDed.Enabled = mDel
End Sub

Private Sub cmbTypes_Click()
If AddDeductions = False And EditDeductions = False Then
    Call RefreshDeductions(cmbTypes.Text)
    If cmbTypes.Text = "Other Deductions" Then
        lblclient(15).Caption = "Remarks"
'        If AddDeductions = True Then
'            cmbDedHeader.Visible = True
'            lblclient(17).Visible = True
'            chkRates(11).Visible = False
'        End If
    Else
        lblclient(15).Caption = "Voucher No"
    End If
ElseIf cmbTypes.Text = "Other Deductions" Then
    cmbDedHeader.Visible = True
    lblclient(17).Visible = True
    chkRates(11).Visible = False
Else
    cmbDedHeader.Visible = False
    lblclient(17).Visible = False
    chkRates(11).Visible = True
End If
If FormLoadSw = True Then
    mDeductionCode = GetDeductionCode(cmbTypes.Text)    ' Bernard get deduction codes
End If
End Sub

Private Sub RefreshDeductions(ByVal mDesc As String)

If rsEmpDeductions.RecordCount <> 0 Then
    rsEmpDeductions.MoveFirst
    rsEmpDeductions.Find "DeductionDescription = '" & mDesc & "'"
    If rsEmpDeductions.EOF = False Then
        'chkRates(11).VALUE = IIf(rsEmpDeductions.Fields("DeductionStatus") = True, 1, 0)   ' 6
        txtRates(17).Text = Format(rsEmpDeductions.Fields("DeductionAmount"), "0.00")      ' 9
        txtRates(18).Text = Format(rsEmpDeductions.Fields("DeductionPayment"), "0.00")     ' 8
        MaskEdBox5.Text = Format(rsEmpDeductions.Fields("DedStartPayment"), "mm/dd/yyyy")   ' start deduction date ' 10
        MaskEdBox6.Text = Format(rsEmpDeductions.Fields("DedEndPayment"), "mm/dd/yyyy")     ' end deduction date   ' 11
        txtRates(15).Text = Format(rsEmpDeductions.Fields("TotalInterest"), "0.00")         ' interest 14
        txtRates(16).Text = Format(rsEmpDeductions.Fields("TotalDeduction"), "0.00")        ' total deductions 15
        txtRates(21).Text = Format(rsEmpDeductions.Fields("NoOfdeductions"), "0.00")        ' no of deductions 12
        txtRates(23).Text = IIf(IsNull(rsEmpDeductions.Fields("VoucherNo")), "", rsEmpDeductions.Fields("VoucherNo"))            ' for voucher number
        chkRates(12).VALUE = IIf(rsEmpDeductions.Fields("DeductionFlag") = True, 1, 0)      ' 17
        mDeductionCode = rsEmpDeductions.Fields("DeductionCode")                            ' 7
        txtDedID.Text = Val(rsEmpDeductions.Fields("DeductionID"))
'        If ActiveDedFlag = 0 Then ' old deductions
'            NoofPayments = rsEmpDeductions.Fields("DeductionCounter")                           ' payments counter  13
'            TotalPayments = rsEmpDeductions.Fields("PaymentsMade")                              ' payments made  16
'        Else
            NoofPayments = GetPaymentsMade(True, True)
            TotalPayments = GetPaymentsMade(True)                                                 ' payments made  16
'        End If
        txtRates(22).Text = Format(TotalPayments, "##0.00")
        NoEditFlag = 1 'IIf(rsEmpDeductions.Fields("NoEditFlag") = True, 1, 0)   ' flag whether the record is editable or not 20
        lblclient(36).Caption = "Payments: " & TotalPayments & " / " & NoofPayments
        mInterest = Val(txtRates(15).Text)
        mTotalLoan = Val(txtRates(16).Text)
    End If
End If

End Sub

Private Function OnPayrollProcess(ByVal mEmpno As Integer) As Boolean
'    strQuery = "SELECT EmployeeID FROM PayrollTrans WHERE clientcode = " & mClientCode & _
'        " and startpaydate = '" & StartDate & "' and endpaydate = '" & EndDate & "' and " & _
'        " EmployeeID = " & mEmpNo
'    Set rsPayrollTrans = ReturnRst(strQuery, cnnStr, mstatic)
End Function

Private Sub cmbTypes_KeyPress(KeyAscii As Integer)
    If AddDeductions = False And EditDeductions = False Then
        mPrevIndex = cmbTypes.ListIndex
        Call RefreshDeductions(cmbTypes.Text)
    End If
End Sub

Private Sub cmdAddDed_Click()
Dim EmpDeductionList As String
Dim mstr As String
Dim res
Dim ProtectFlag As Boolean
Dim mDetails As String
Dim mRemarks As String

    ProtectFlag = False
    On Error GoTo NextProc
    If cmdAddDed.Caption = "New" Then
        StrQuery = "SELECT * FROM DeductionTypes2 WHERE DeductionCode <> 0 ORDER BY DeductionDescription"
        Set rsGetDeductionTypes = ReturnRst(StrQuery, CnnstrHR, mstatic)
        
        ' Get all the types of Deductions not yet included for the current Employee.
        EmpDeductionList = ""
        If rsEmpDeductions.RecordCount <> 0 Then
            rsEmpDeductions.MoveFirst
            While Not rsEmpDeductions.EOF
                EmpDeductionList = EmpDeductionList & "[" & LTrim(RTrim(str(rsEmpDeductions.Fields(7)))) & "] "
                rsEmpDeductions.MoveNext
            Wend
        End If
        
        ' Fill the combo with New Deductions.
        Call FillDeductionList(cmbTypes, rsGetDeductionTypes, True, EmpDeductionList)
        If cmbTypes.ListCount = 0 Then
            MsgBox "No available deductions to add!", vbInformation, PayrollAlert
            GoTo Routine1
        End If
        Call EnableDisableDeductions(True)
        Call EnableDisableDeductions("", True)
        Call EnableDisableDeductionButtons("Save", "Cancel", False)
        Call GetOtherDed    ' update the list of other deductions
        
        ' mDeductionCode = GetDeductionCode(cmbTypes.Text)
        cmdSchedule.Enabled = False
        cmdCompute.Enabled = False
        AddDeductions = True
        chkRates(11).Enabled = False
        cmbCompany.Enabled = False
        ListView1(0).Enabled = False
        cmbTypes.SetFocus
    
    Else  ' Save changes to Deductions.
        
        ' Validate the entries first.
        If MaskEdBox5 = "__/__/____" Or MaskEdBox5 = "01/01/1900" Then
            MsgBox "Start Deduction Date should not be blank or equal to 01/01/1900!!!", vbInformation, PayrollAlert
            MaskEdBox5.SetFocus
            Exit Sub
        ElseIf MaskEdBox5 <> "__/__/____" And Not IsDate(MaskEdBox5) Then
            MsgBox "Please enter a valid Start Deduction date!!!", vbInformation, PayrollAlert
            MaskEdBox5.SetFocus
            Exit Sub
        End If
        If MaskEdBox6 = "__/__/____" Or MaskEdBox6 = "01/01/1900" Then
            MsgBox "End Deduction Date should not be blank or equal to 01/01/1900!!!", vbInformation, PayrollAlert
            MaskEdBox6.SetFocus
            Exit Sub
        ElseIf MaskEdBox6 <> "__/__/____" And Not IsDate(MaskEdBox6) Then
            MsgBox "Please enter a valid End Deduction date!!!", vbInformation, PayrollAlert
            MaskEdBox6.SetFocus
            Exit Sub
        End If
'        If MaskEdBox5 > MaskEdBox6 Then
'            MsgBox "Start Deduction date should be less than or equal to End Deduction date!!!", vbInformation, PayrollAlert
'            MaskEdBox5.SetFocus
'            Exit Sub
'        End If
        If Val(txtRates(21).Text) = 0# Then         ' And (MaskEdBox6 = "__/__/____" Or Format(Date, MaskEdBox6) = #1/1/1900#) Then
            MsgBox "Number of Payments is required, to monitor payments made!", vbInformation, PayrollAlert
            txtRates(21).SetFocus
            Exit Sub
        End If
        
        rsEmpDeductions.Requery
'        If rsEmpDeductions.Fields("NoEditFlag") = True Then
'            MsgBox "Currently processing payroll for this client! Cannot save changes.", vbInformation, PayrollAlert
'            GoTo Routine1
'        End If
    
        If EditDeductions = True Then
            If cmbTypes.Text = "Other Deductions" Then
                Call SaveOtherDed
                GoTo Routine1
            End If
            If TotalPayments <> 0 And TotalPayments = Val(txtRates(16)) Then ' deductions is fully paid
                ProtectFlag = True
            ElseIf TotalPayments <> 0 Then ' ask whether they wanted to update the details of payment
                res = MsgBox("Total Loan not equal to Total Payments made." & Chr(13) & _
                    "Do you want to update the Deduction Payment for " & cmbTypes.Text & "?" & Chr(13) & Chr(13) & _
                    "This will save the old Payment Details for reference and insert updated entries!", vbQuestion + vbYesNo, PayrollAlert)
                If res = vbYes Then ' save old details for reference
                    ProtectFlag = True
                Else
                    MsgBox "You need to save old Payment Details for reference if payment is not yet completed!", vbInformation, PayrollAlert
                    txtRates(17).SetFocus
                    Exit Sub
                End If
            End If
        End If
        
        'Save/Edit the entries.
        If AddDeductions = True Then         ' Insert New Deductions
            If cmbTypes.Text = "Other Deductions" Then
                Call SaveOtherDed
                GoTo Routine1
            End If
            StrQuery = "INSERT INTO EmployeeDeductions (EmployeeID, EmpNo, AgencyCode, ClientCode, DeductionStatus, DeductionCode, DeductionPayment, DeductionAmount, VoucherNo, DedStartPayment, DedEndPayment, TotalInterest, TotalDeduction, NoOfDeductions, DeductionFlag, EntryBy, EntryDate) " & _
                " VALUES ( " & txtEmpNum & ", '" & txtClient(0) & "', " & CompanyCode & ", " & mCCode & ", " & ActiveDedFlag & ", " & mDeductionCode & ", " & txtRates(18) & ", " & txtRates(17) & ", '" & txtRates(23) & "'" & _
                ", '" & IIf(MaskEdBox5 = "__/__/____", #1/1/1900#, Format(MaskEdBox5, "mm/dd/yyyy")) & "', '" & IIf(MaskEdBox6 = "__/__/____", #1/1/1900#, Format(MaskEdBox6, "mm/dd/yyyy")) & _
                "', " & txtRates(15) & ", " & txtRates(16) & ", " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", " & chkRates(12) & ", '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & "')"
            mstr = "Deduction was saved!"
            mDetails = "Deduction was added with total of " & txtRates(16) & " for " & txtClient(2)
            mRemarks = "New Deduction."
            Call Action_Query(StrQuery, CnnstrHR)
        ElseIf EditDeductions = True Then       ' Save Edited Deductions
            If ProtectFlag = True Then          ' Save old details for reference
                StrQuery = "UPDATE EmployeeDeductions SET DeductionFlag = 1, DeductionStatus = 0 " & _
                    " Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode & " AND DeductionID = " & Val(txtDedID.Text)
                    
                    mstr = "Deduction Details was Updated!"
                    mDetails = "Deduction Total was updated to " & txtRates(16) & " for " & txtClient(2)
                    mRemarks = "Update Deduction."
                    
                    If Action_Query(StrQuery, CnnstrHR) = True Then
                        ' Need to save old Payment details.
                        StrQuery = "UPDATE EmployeeDedSchedules SET DeductionStatus = 0, DeductionFlag = 1, Remarks = LTRIM(RTRIM(Remarks)) + ' Updated.'" & _
                            " WHERE EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode & _
                            " AND DeductionFlag = 0 AND DeductionID = " & Val(txtDedID.Text)
                        Call Action_Query(StrQuery, CnnstrHR)
                        'Insert the new deduction entries.
                        StrQuery = "INSERT INTO EmployeeDeductions (EmployeeID, EmpNo, AgencyCode, ClientCode, DeductionStatus, DeductionCode, DeductionPayment, DeductionAmount, VoucherNo, DedStartPayment, DedEndPayment, TotalInterest, TotalDeduction, NoOfDeductions, DeductionFlag, EntryBy, EntryDate) " & _
                            " VALUES ( " & txtEmpNum & ", '" & txtClient(0) & "', " & CompanyCode & ", " & mCCode & ", 1, " & mDeductionCode & ", " & txtRates(18) & ", " & txtRates(17) & ", '" & txtRates(23) & _
                            "', '" & IIf(MaskEdBox5 = "__/__/____", #1/1/1900#, Format(MaskEdBox5, "mm/dd/yyyy")) & "', '" & IIf(MaskEdBox6 = "__/__/____", #1/1/1900#, Format(MaskEdBox6, "mm/dd/yyyy")) & _
                            "', " & txtRates(15) & ", " & txtRates(16) & ", " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", " & chkRates(12) & ", '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & "')"
                        Call Action_Query(StrQuery, CnnstrHR)
                        'Reset counters to 0.
                        TotalPayments = 0
                        txtRates(22).Text = "0.00"
                        NoofPayments = 0
                    End If
            Else
                StrQuery = "UPDATE EmployeeDeductions SET EmployeeID = " & txtEmpNum & ", EmpNo = '" & txtClient(0) & "', AgencyCode = " & CompanyCode & ", ClientCode = " & mCCode & _
                    ", DeductionStatus = " & ActiveDedFlag & ", DeductionCode = " & mDeductionCode & ", DeductionPayment = " & txtRates(18) & ", DeductionAmount = " & txtRates(17) & ", VoucherNo = '" & txtRates(23) & "'" & _
                    ", DedStartPayment = '" & IIf(MaskEdBox5 = "__/__/____", #1/1/1900#, Format(MaskEdBox5, "mm/dd/yyyy")) & "', DedEndPayment = '" & IIf(MaskEdBox6 = "__/__/____", #1/1/1900#, Format(MaskEdBox6, "mm/dd/yyyy")) & _
                    "', TotalInterest = " & txtRates(15) & ", TotalDeduction = " & txtRates(16) & ", NoOfDeductions = " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", DeductionFlag = 0, EntryBy = '" & mUserID & "', EntryDate = '" & Format(Now, "mm/dd/yyyy") & _
                    "' Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode & " AND DeductionID = " & Val(txtDedID.Text)
                    mstr = "Deduction was edited!"
                    mDetails = "Deduction Total was edited to " & txtRates(16) & " for " & txtClient(2)
                    mRemarks = "Edit Deduction."
                    Call Action_Query(StrQuery, CnnstrHR)
            End If
        End If
        
        Call SaveUserTransactions(mDetails, mRemarks, txtEmpNum, txtClient(0))
        ListView1(2).ListItems.Clear
        MsgBox mstr, vbInformation, PayrollAlert
        
Routine1:

        Call ReQueryDeductions
        Call EnableDisableDeductions(False)
        Call EnableDisableDeductionButtons("New", "Edit", True)
        
        'cmdSchedule.Enabled = True
        cmdCompute.Enabled = True
        cmbCompany.Enabled = True
        ListView1(0).Enabled = True

        EditDeductions = False
        AddDeductions = False
        chkRates(11).Enabled = True
        
        cmbTypes.Enabled = True
        
        cmbDedHeader.Visible = False
        lblclient(17).Visible = False
        chkRates(11).Visible = True     'selection history
        
    End If
    
    Exit Sub
    
NextProc:

    MsgBox Err.Description
    Err.Clear
    
End Sub

Private Sub GetOtherDed()
Dim rsDed As Recordset

On Error GoTo ErrRtn11

Rtn22:

'If cmbTypes.Text = "Other Deductions" Then
    StrQuery = "SELECT DedID, OtherDedName FROM HROtherDed ORDER BY OtherDedName"
    Set rsDed = ReturnRst(StrQuery, CnnstrHR, mstatic)
    If rsDed.RecordCount <> 0 Then
        Call FillComboData(cmbDedHeader, rsDed)
        cmbDedHeader.ListIndex = 0
    rsDed.Close
    End If
'End If
Set rsDed = Nothing

Exit Sub

ErrRtn11:

    'MsgBox Err.Description & " " & str(Err.Number)
    If Err.Number = 91 Then
        StrQuery = "CREATE TABLE [dbo].[HROtherDed] (" & _
            "[DedID] [int] IDENTITY (1, 1) NOT NULL , " & _
            "[DeductionCode] [int] Default (0) NOT NULL , " & _
            "[OtherDedName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL " & _
            ") ON [PRIMARY]"
        Call Action_Query(StrQuery, CnnstrHR)
        'MsgBox "Error was corrected.", vbInformation, PayrollAlert
        GoTo Rtn22
    End If
End Sub

Private Sub SaveOtherDed()  ' for Other Deductions only
Dim rsOldID As Recordset

On Error GoTo ExitRtn33

If EditDeductions = True Then
'    StrQuery = "UPDATE EmployeeDeductions SET EmployeeID = " & txtEmpNum & ", EmpNo = '" & txtClient(0) & "', AgencyCode = " & CompanyCode & ", ClientCode = " & mCCode & _
'        ", DeductionStatus = " & ActiveDedFlag & ", DeductionCode = " & mDeductionCode & ", DeductionPayment = " & txtRates(18) & ", DeductionAmount = " & txtRates(17) & ", VoucherNo = '" & txtRates(23) & "'" & _
'        ", DedStartPayment = '" & IIf(MaskEdBox5 = "__/__/____", #1/1/1900#, Format(MaskEdBox5, "mm/dd/yyyy")) & "', DedEndPayment = '" & IIf(MaskEdBox6 = "__/__/____", #1/1/1900#, Format(MaskEdBox6, "mm/dd/yyyy")) & _
'        "', TotalInterest = " & txtRates(15) & ", TotalDeduction = " & txtRates(16) & ", NoOfDeductions = " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", DeductionFlag = 0, EntryBy = '" & mUserID & "', EntryDate = '" & Format(Now, "mm/dd/yyyy") & _
'        "' Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode & " AND DeductionID = " & Val(txtDedID.Text)
'        Call Action_Query(StrQuery, CnnstrHR)

    StrQuery = "DELETE EmployeeDedSchedules WHERE EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode & _
        " AND DeductionID = " & Val(txtDedID.Text) & " AND RecordSeq = " & RecordSeq
    Call Action_Query(StrQuery, CnnstrHR)
    Call SaveDeductionEntries(MaskEdBox5, 1, txtRates(18), txtRates(17), txtRates(23), True)
    GetPaymentsMade (False)
    cmdSchedule.Enabled = True
Else
    ' Save first previous DeductionID.
    StrQuery = "SELECT DeductionID FROM EmployeeDeductions " & _
        " Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode & _
        " AND DeductionStatus = 1 "
    Set rsOldID = ReturnRst(StrQuery, CnnstrHR, mstatic)
    If rsOldID.RecordCount <> 0 Then
        txtDedID.Text = Val(rsOldID.Fields("DeductionID"))
        ' Update only.
        StrQuery = "UPDATE EmployeeDeductions SET EmployeeID = " & txtEmpNum & ", EmpNo = '" & txtClient(0) & "', AgencyCode = " & CompanyCode & ", ClientCode = " & mCCode & _
            ", DeductionStatus = " & ActiveDedFlag & ", DeductionCode = " & mDeductionCode & ", DeductionPayment = " & txtRates(18) & ", DeductionAmount = " & txtRates(17) & ", VoucherNo = '" & txtRates(23) & "'" & _
            ", DedStartPayment = '" & IIf(MaskEdBox5 = "__/__/____", #1/1/1900#, Format(MaskEdBox5, "mm/dd/yyyy")) & "', DedEndPayment = '" & IIf(MaskEdBox6 = "__/__/____", #1/1/1900#, Format(MaskEdBox6, "mm/dd/yyyy")) & _
            "', TotalInterest = " & txtRates(15) & ", TotalDeduction = " & txtRates(16) & ", NoOfDeductions = " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", DeductionFlag = 0, EntryBy = '" & mUserID & "', EntryDate = '" & Format(Now, "mm/dd/yyyy") & _
            "' Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode & " AND DeductionID = " & Val(txtDedID.Text)
        Call Action_Query(StrQuery, CnnstrHR)
    
'            ", DeductionStatus = " & ActiveDedFlag & ", DeductionCode = " & mDeductionCode & ", DeductionPayment = " & txtRates(18) & ", DeductionAmount = " & txtRates(17) & ", VoucherNo = '" & txtRates(23) & "'"
    Else
        ' Delete first if existing.
        StrQuery = "DELETE FROM EmployeeDeductions " & _
            " Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode & _
            " AND DeductionID = " & Val(txtDedID.Text) & " AND DeductionStatus = 1 "
        Call Action_Query(StrQuery, CnnstrHR)
        StrQuery = "INSERT INTO EmployeeDeductions (EmployeeID, EmpNo, AgencyCode, ClientCode, DeductionStatus, DeductionCode, DeductionPayment, DeductionAmount, VoucherNo, DedStartPayment, DedEndPayment, TotalInterest, TotalDeduction, NoOfDeductions, DeductionFlag, EntryBy, EntryDate) " & _
            " VALUES ( " & txtEmpNum & ", '" & txtClient(0) & "', " & CompanyCode & ", " & mCCode & ", " & ActiveDedFlag & ", " & mDeductionCode & ", " & txtRates(18) & ", " & txtRates(17) & ", '" & txtRates(23) & "'" & _
            ", '" & IIf(MaskEdBox5 = "__/__/____", #1/1/1900#, Format(MaskEdBox5, "mm/dd/yyyy")) & "', '" & IIf(MaskEdBox6 = "__/__/____", #1/1/1900#, Format(MaskEdBox6, "mm/dd/yyyy")) & _
            "', " & txtRates(15) & ", " & txtRates(16) & ", " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", " & chkRates(12) & ", '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & "')"
        Call Action_Query(StrQuery, CnnstrHR)
    End If
    Set rsOldID = Nothing
End If

' Save it for payroll deductions.
' Call SaveDeductionEntries(MaskEdBox5, 1, txtRates(18), txtRates(17), txtRates(23))

' Update Other Deduction table.
Call UpdateOtherDeductions

Exit Sub

ExitRtn33:

    MsgBox "Failed to save: " & Err.Description, vbInformation, PayrollAlert
    
End Sub

Private Sub UpdateOtherDeductions()

StrQuery = "INSERT INTO HROtherDed (DeductionCode, OtherDedName) " & _
    " VALUES ( " & mCCode & ", '" & cmbDedHeader & "')"
Call Action_Query(StrQuery, CnnstrHR)
'Call GetOtherDed

End Sub

Private Sub cmdCancelComputation_Click()
cmdCancelComputation.Visible = False
cmdSaveComputation.Visible = False
cmdSchedule.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Set rsEmployeeBasicData = Nothing
    Unload Me
End Sub

Private Function GetDeductionCode(ByVal mDesc As String) As Integer

GetDeductionCode = 0
If mDesc <> "" Or mDesc <> Null Then
    If InStr(1, mDesc, "(", vbTextCompare) > 0 Then
        mDesc = mID(mDesc, 1, InStr(1, mDesc, "(", vbTextCompare) - 1)
    End If
    rsGetDeductionTypes.MoveFirst
    rsGetDeductionTypes.Find "DeductionDescription = '" & mDesc & "'"
    If rsGetDeductionTypes.EOF = False Then
        GetDeductionCode = rsGetDeductionTypes.Fields(0)
    End If
End If
    
End Function

Private Sub cmdCompute_Click()
    If cmbTypes.Text = "Other Deductions" Then
        Call ComputeDeductions(True)
    ElseIf MsgBox("This will recompute payment schedules for " & cmbTypes.Text & "! Do you want to continue?", vbQuestion + vbYesNo) = vbYes Then
        Call ComputeDeductions
    End If
End Sub

Private Sub ComputeDeductions(Optional mSave As Boolean = False)
Dim mLsEmp As ListView
Dim litem1 As ListItem
Dim Dedctr As Integer
Dim mDate As Date
Dim mDed As Double, mPayment As Double, mRunBalance As Double

On Error GoTo err1

If AddDeductions = False And EditDeductions = False Then
    Call RefreshDeductions(cmbTypes.Text)
End If

InActiveCtr = 0
ActiveCtr = 0
Set mLsEmp = ListView1(2)
mLsEmp.View = lvwReport
mLsEmp.ListItems.Clear
Dedctr = 0
mPayment = 0#
mRunBalance = 0#
mDate = MaskEdBox5  ' Get the start date
mDed = 0#
'While Dedctr <= txtRates(21).Text
While mDed < Val(txtRates(16).Text)
    If Dedctr > 0 Then
        mDate = GetScheduleDate(mDate)
    End If
    Dedctr = Dedctr + 1
    If Dedctr > 80 Then
        MsgBox "To many deduction schedules made! The system will not proceed.", vbInformation, PayrollAlert
        GoTo ExitDed
    End If
    mPayment = IIf(IsNull(txtRates(18).Text), "0.00", txtRates(18).Text)        ' deduction per payroll
    If (Val(txtRates(16).Text) - mDed) > Val(txtRates(18).Text) Then
        mPayment = IIf(IsNull(txtRates(18).Text), "0.00", txtRates(18).Text)    ' deduction per payroll
    Else
        mPayment = Val(txtRates(16).Text) - mDed
    End If
    Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(cmbTypes), "Blank Ded", cmbTypes)) ' Deduction Types
        litem1.SubItems(1) = IIf(IsNull(cmbTypes), "Blank Ded", cmbTypes)
        litem1.SubItems(2) = Dedctr ' Counter
        litem1.SubItems(3) = Format(IIf(IsNull(txtRates(16).Text), "0.00", txtRates(16).Text) - mDed, "##0.00") ' Total balance
        mRunBalance = Format(IIf(IsNull(txtRates(16).Text), "0.00", txtRates(16).Text) - mDed, "##0.00") ' Total balance
        litem1.SubItems(4) = Format(mPayment, "##0.00")
        litem1.SubItems(5) = "0.00" ' Get payments made
        litem1.SubItems(6) = mDate
    If mSave = True Then
        If cmbTypes.Text = "Other Deductions" Then
            Call SaveDeductionEntries(mDate, Dedctr, mPayment, mRunBalance, txtRates(23), True)
            litem1.SubItems(7) = txtRates(23).Text & " (" & cmbDedHeader & ")" ' Remarks
        Else
            Call SaveDeductionEntries(mDate, Dedctr, mPayment, mRunBalance)
            litem1.SubItems(7) = "For Deduction."
        End If
    Else
        litem1.SubItems(7) = "For Schedule."
    End If
    mDed = mDed + Val(IIf(IsNull(txtRates(18).Text), "0.00", txtRates(18).Text))
Wend
    
ExitDed:

MaskEdBox6 = Format(mDate, "mm/dd/yyyy")
txtRates(21).Text = Dedctr

If mSave = False Then
    cmdSaveComputation.Visible = True
    cmdCancelComputation.Visible = True
    MsgBox "Done with the computation of Deduction Schedules!", vbInformation, PayrollAlert
    cmdSchedule.Enabled = False
Else
    cmdSaveComputation.Visible = False
    cmdCancelComputation.Visible = False
    MsgBox "Reminder: " & Chr(13) & " Deduction Computation was successfully save for Payroll Deductions!", vbInformation, PayrollAlert
    cmdSchedule.Enabled = True
End If

Exit Sub

err1:

'MsgBox Err.Description
Err.Clear
'Resume Next
End Sub

Private Function GetScheduleDate(ByVal asDate As Date) As Date
Dim newDate As Date

If Format(asDate, "dd") <= 11 Then  ' 16
'    GetScheduleDate = Format(str(Format(asDate, "mm")) & "/16/" & str(Format(asDate, "yyyy")), "mm/dd/yyyy")
    GetScheduleDate = Format(str(Format(asDate, "mm")) & "/26/" & str(Format(asDate, "yyyy")), "mm/dd/yyyy")
Else
    If Format(asDate, "mm") = 12 Then
'        GetScheduleDate = Format("01/01/" & str(Format(asDate, "yyyy") + 1), "mm/dd/yyyy")
        GetScheduleDate = Format("01/11/" & str(Format(asDate, "yyyy") + 1), "mm/dd/yyyy")
    Else
'        GetScheduleDate = Format(str(Format(asDate, "mm") + 1) & "/01/" & str(Format(asDate, "yyyy")), "mm/dd/yyyy")
        GetScheduleDate = Format(str(Format(asDate, "mm") + 1) & "/11/" & str(Format(asDate, "yyyy")), "mm/dd/yyyy")
    End If
End If

End Function

Private Sub SaveDeductionEntries(ByVal mDate As Date, ByVal Dedno As Integer, _
    ByVal mPayment As Double, ByVal mRunBal As Double, Optional mRemarks As String = "For Deduction.", _
    Optional OtherDedFlag As Boolean = False)

Dim mOtherDed As String

If cmbDedHeader.Text = "" Then
    mOtherDed = "Other Deduction"
Else
    mOtherDed = cmbDedHeader.Text
End If
'Delete first those old schedules, when no payments took place.
If Dedno = 1 And TotalPayments = 0 And mRemarks = "For Deduction." Then
    StrQuery = "DELETE EmployeeDedSchedules WHERE EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode & _
        " AND DeductionID = " & Val(txtDedID.Text)
    Call Action_Query(StrQuery, CnnstrHR)
End If

'Insert the new schedules.
StrQuery = "INSERT INTO EmployeeDedSchedules (Remarks, EmployeeName, EmployeeID, EmpNo, DeductionDescription, AgencyCode, ClientCode, DeductionStatus, DeductionCode, DeductionPayment, RunningBalance, DeductionAmount, VoucherNo, DedStartPayment, DedEndPayment, TotalInterest, TotalDeduction, NoOfDeductions, DeductionCounter, DeductionFlag, EntryBy, EntryDate, DeductionID) " & _
    " VALUES ( '" & mRemarks & "', '" & txtClient(2).Text & "', " & txtEmpNum & ", '" & txtClient(0) & "', '" & cmbTypes & "', " & CompanyCode & ", " & mCCode & ", " & ActiveDedFlag & ", " & mDeductionCode & ", " & mPayment & ", " & mRunBal & ", " & txtRates(17) & ", '" & IIf(OtherDedFlag = False, txtRates(23), mOtherDed) & _
    "', '" & mDate & "', '" & mDate & _
    "', " & txtRates(15) & ", " & txtRates(16) & ", " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", " & Dedno & ", " & chkRates(12) & ", '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & "', " & Val(txtDedID.Text) & ")"
Call Action_Query(StrQuery, CnnstrHR)

'MsgBox "Deduction Schedule for " & cmbTypes & " was updated!", vbInformation, PayrollAlert
'    ", '" & IIf(mDate = "__/__/____", #1/1/1900#, Format(mDate, "mm/dd/yyyy")) & "', '" & IIf(mDate = "__/__/____", #1/1/1900#, Format(mDate, "mm/dd/yyyy"))

End Sub

Private Sub cmdDeductionTypes_Click()

    Call frmDeductionTypes.Show
    
    ' Update the deduction types recordset with new entries.
    StrQuery = "SELECT * FROM DeductionTypes2 WHERE DeductionCode <> 0 ORDER BY DeductionDescription"
    Set rsDeductionTypes = ReturnRst(StrQuery, CnnstrHR, mstatic)
    
    Call ReQueryDeductions
    
'    ' Update the recordset for Employee Deductions.
'    StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
'    'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ProtectedFlag = 0"
'    Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
'
'    ' Re-filter deductions.
'    If txtEmpNum = "txtclientcode" Then
'        txtEmpNum = 0
'    End If
'
'    Call FilterDeductions(txtEmpNum)

End Sub

Private Sub cmdDelDed_Click()
Dim res
Dim mRem As String

    mRem = mUserID
    If cmbTypes.Text = "" Then MsgBox "No Active Deductions, cannot be deleted!": Exit Sub
    mDeductionCode = GetDeductionCode(cmbTypes.Text)
    If mDeductionCode = 8 Then
        res = MsgBox("Do you want to delete this selected other deduction?", vbQuestion + vbYesNo, PayrollAlert)
        If res = vbYes Then ' delete deductions
            StrQuery = "DELETE EmployeeDedSchedules WHERE EmployeeID = " & txtEmpNum & _
                " AND DeductionCode = " & mDeductionCode & _
                " AND DeductionID = " & Val(txtDedID.Text) & " AND DeductionStatus = 1 AND RecordSeq = " & RecordSeq
            Call Action_Query(StrQuery, CnnstrHR)
            Call cmdSchedule_Click
            GoTo Exit1
        End If
    End If
    If TotalPayments <> 0 Then
'        res = MsgBox("There were already " & NoofPayments & " payments made amounting to " & TotalPayments & ". Do you want to mark it as fully paid instead?", vbQuestion + vbYesNo, PayrollAlert)
'        If res = vbYes Then ' delete deductions
'            mRem = "Marked as fully paid by " & mUserID & ", last " & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM")
'            StrQuery = "UPDATE EmployeeDeductions SET EmployeeID = " & txtEmpNum & ", EmpNo = '" & txtClient(0) & "', AgencyCode = " & CompanyCode & ", ClientCode = " & mCCode & _
'                    ", DeductionStatus = " & chkRates(11) & ", DeductionCode = " & mDeductionCode & ", DeductionPayment = " & IIf(txtRates(18) = "", 0, txtRates(18)) & ", DeductionAmount = " & IIf(txtRates(17) = "", 0, txtRates(17)) & _
'                    ", DedStartPayment = '" & IIf(MaskEdBox5 = "__/__/____", #1/1/1900#, Format(MaskEdBox5, "mm/dd/yyyy")) & "', DedEndPayment = '" & IIf(MaskEdBox6 = "__/__/____", #1/1/1900#, Format(MaskEdBox6, "mm/dd/yyyy")) & _
'                    "', TotalInterest = " & IIf(txtRates(15) = "", 0, txtRates(15)) & ", TotalDeduction = " & txtRates(16) & ", NoOfDeductions = " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", DeductionFlag = " & chkRates(12) & ", EntryBy = '" & mRem & "', EntryDate = '" & Format(Now, "mm/dd/yyyy") & "'" & _
'                    ", Protectedflag = 1" & _
'                    " Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode
'                    '" Where AgencyCode = " & CompanyCode & " and clientcode = " & mCCode & " and EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode ' Updated by Bernard
'            Call Action_Query(StrQuery, CnnstrHR)
'            MsgBox "Deduction was marked as Fully Paid!"
'        Else
            'MsgBox "You need to update the Deduction Information or set Total Deduction equal to " & TotalPayments & ", if you do not want to deduc it from Payroll.", vbInformation, PayrollAlert
            MsgBox "There were already " & NoofPayments & " payments made amounting to " & TotalPayments & "." & Chr(13) & _
                "Please proceed with Edit, cannot delete entries with active payments!", vbInformation, PayrollAlert
'        End If
    Else
        res = MsgBox("Delete " & cmbTypes.Text & " deductions? No deduction will reflect to Payroll if you delete.", vbQuestion + vbYesNo, PayrollAlert)
        If res = vbYes Then ' Delete Active Deductions only
            StrQuery = "DELETE FROM EmployeeDeductions " & _
                " Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode & _
                " AND DeductionID = " & Val(txtDedID.Text) & " AND DeductionStatus = 1 "
            ' Delete also records in deduction schedules.
            If Action_Query(StrQuery, CnnstrHR) = True Then
                StrQuery = "DELETE EmployeeDedSchedules WHERE EmployeeID = " & txtEmpNum & _
                    " AND DeductionCode = " & mDeductionCode & _
                    " AND DeductionID = " & Val(txtDedID.Text) & " AND DeductionStatus = 1"
                Call Action_Query(StrQuery, CnnstrHR)
            End If
            
'            Call EnableDisableDeductions(False)
'            Call EnableDisableDeductionButtons("Add", "Edit", True)
'            cmdEditDed.Caption = "Edit"
'            cmbTypes.Enabled = True
'            strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode
'            Set rsEmpDeductions = ReturnRst(strQuery, CnnstrHR, mstatic)
'            ' Re-filter deductions.
'            Call FilterDeductions(txtEmpNum)
        End If
    End If
    
Exit1:
    
    Call EnableDisableDeductions(False)
    Call EnableDisableDeductionButtons("New", "Edit", True)
    cmdEditDed.Caption = "Edit"
    cmbTypes.Enabled = True
    
    Call ReQueryDeductions
'    StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
'    'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ProtectedFlag = 0"
'    Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
'    ' Re-filter deductions.
'    Call FilterDeductions(txtEmpNum)

End Sub

Private Sub ReQueryDeductions()
Dim mTypes As String

On Error GoTo err11

mTypes = cmbTypes.Text
    
StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)

' Re-filter deductions.
If txtEmpNum = "txtclientcode" Then
    txtEmpNum = 0
End If

' Re-filter Deductions.
Call FilterDeductions(txtEmpNum)

If cmbTypes.ListCount <> 0 And mTypes <> "" Then
    cmbTypes.Text = mTypes
End If

Exit Sub

err11:

    Err.Clear

End Sub

Private Sub cmdDelHistory_Click()
Static mTag As Boolean

If cmbTypes.Text = "" Then
    MsgBox "Deduction types is blank, kindly select deduction first!"
    Exit Sub
End If
    If cmdDelHistory.Caption = "Delete Blanks" Then ' Delete Blanks
        MsgBox "Kindly input the exact Loan Amount in Payment box above."
        cmdDelHistory.Caption = "Start Deletion"
        mTag = True
        txtRates(22).SetFocus
    ElseIf cmdDelHistory.Caption = "Start Deletion" Then  ' Start Deletion
        If Val(txtRates(22)) = 0# Or txtRates(22) = "" Then
            MsgBox "No input value in Payment Box, cannot proceed!"
        Else
            If mTag = True Then ' Delete Blanks
                Call DeleteDedRecords(0)
                cmdDelHistory.Caption = "Delete Blanks"
            Else
                Call DeleteDedRecords(1)
                cmdDelHistory.Caption = "Delete History"
            End If
        End If
    Else   ' Delete History
        MsgBox "Kindly input the Deduction ID enclosed with () in Payment box above."
        cmdDelHistory.Caption = "Start Deletion"
        mTag = False
        txtRates(22).SetFocus
    End If
End Sub

Private Sub DeleteDedRecords(ByVal mCap As Boolean)
Dim res

If mCap = 0 Then ' set Blank Desc deduction status to 0
    res = MsgBox("Start Deleting blanks for " & cmbTypes & " deductions?", vbQuestion + vbYesNo, PayrollAlert)
    If res = vbYes Then ' delete blanks
        StrQuery = "UPDATE EmployeeDedSchedules SET DeductionStatus = 0, DeductionDescription = '" & cmbTypes.Text & "'" & _
        " WHERE EmployeeID = " & txtEmpNum & " and ClientCode = " & mCCode & _
        " AND DeductionCode = " & mDeductionCode & " AND DeductionStatus = 1" & _
        " AND DeductionAmount = " & Val(txtRates(22)) & " AND DeductionDescription = '' "
        If Action_Query(StrQuery, CnnstrHR) = True Then
            MsgBox "Successfully deleted Deduction Schedules for blanks amounting to " & txtRates(22) & "!"
        Else
            MsgBox "Failed to delete Deduction Schedules for blanks amounting to " & txtRates(22) & "!"
        End If
    End If
Else ' Delete History
    res = MsgBox("Start Deleting History for " & cmbTypes & " deductions?", vbQuestion + vbYesNo, PayrollAlert)
    If res = vbYes Then ' delete blanks
        StrQuery = "DELETE EmployeeDedSchedules " & _
            " WHERE EmployeeID = " & txtEmpNum & " and ClientCode = " & mCCode & _
            " AND DeductionCode = " & mDeductionCode & " AND DeductionID = " & Val(txtRates(22).Text) & _
            " AND DeductionStatus = 0 "
        If Action_Query(StrQuery, CnnstrHR) = True Then
            MsgBox "Successfully deleted Deduction History records!"
        Else
            MsgBox "Failed to delete Deduction History records!"
        End If
    End If
End If
End Sub

Private Sub cmdEditDed_Click()
Dim res
    
    If cmdEditDed.Caption = "Edit" Then
        If cmbTypes.Text = "" Then MsgBox "No Active Deductions, add new deduction first!": GoTo NextProc
        StrQuery = "SELECT * FROM DeductionTypes2 WHERE DeductionCode <> 0 ORDER BY DeductionDescription"
        Set rsGetDeductionTypes = ReturnRst(StrQuery, CnnstrHR, mstatic)
        
        cmbTypes.Enabled = False        ' do not let the user change it
        Call EnableDisableDeductions(True)
        Call EnableDisableDeductionButtons("Save", "Cancel", False)
        mDeductionCode = GetDeductionCode(cmbTypes.Text)
        
'        If TotalPayments <> 0 Then  ' protect the start date from editing if there were payments made already
'            MaskEdBox5.Enabled = False
'        Else
'            MaskEdBox5.Enabled = True
'        End If
        cmdSchedule.Enabled = False
        cmdCompute.Enabled = False
        cmbCompany.Enabled = False
        ListView1(0).Enabled = False
        
        EditDeductions = True
        chkRates(11).Enabled = False
        txtRates(17).SetFocus
    Else
        'res = MsgBox("Cancel changes to Deductions?", vbQuestion + vbYesNo, PayrollAlert)
        'If res = vbYes Then ' cancel the changes
        Call EnableDisableDeductions(False)
        Call EnableDisableDeductionButtons("New", "Edit", True)
        cmdEditDed.Caption = "Edit"
        EditDeductions = False
        AddDeductions = False
        chkRates(11).Enabled = True
        cmbCompany.Enabled = True
        ListView1(0).Enabled = True
        
        cmbTypes.Enabled = True
        
        cmdSchedule.Enabled = True
        cmdCompute.Enabled = True
        
        MsgBox "Changes to deductions was cancelled!", vbInformation, PayrollAlert
        'End If
    End If

        cmbDedHeader.Visible = False
        lblclient(17).Visible = False
        chkRates(11).Visible = True     'selection history

NextProc:

End Sub

Private Sub cmdExportToExcel_Click()
Dim rsDed As ADODB.Recordset
Dim mCon As String

mCon = ""
If chkPayments.VALUE = 1 Then ' for payments only
    mCon = " AND PaymentsMade > 0.00 "
End If

If chkAllEmp.VALUE = 1 Then ' all employees
    'StrQuery = "SELECT * FROM EmployeeDedSchedules WHERE ClientCode = " & mCCode & " ORDER BY EmployeeName, DeductionCode, DedStartPayment"
    '    " WHERE ClientCode = " & mCCode & " AND DeductionCode = " & mDeductionCode
    If chkAllDed.VALUE = 1 Then ' All deductions
        StrQuery = "SELECT * FROM EmployeeDedSchedules " & _
            " WHERE ClientCode = " & mCCode & mCon & _
            " AND DeductionStatus = " & ActiveDedFlag & _
            " ORDER BY EmployeeName, DeductionCode, DedStartPayment"
    Else ' selected deduction only
        StrQuery = "SELECT * FROM EmployeeDedSchedules " & _
            " WHERE ClientCode = " & mCCode & " AND DeductionCode = " & mDeductionCode & mCon & _
            " AND DeductionStatus = " & ActiveDedFlag & _
            " ORDER BY EmployeeName, DeductionCode, DedStartPayment"
    End If
Else
    If chkAllDed.VALUE = 1 Then ' all Deductions
        StrQuery = "SELECT * FROM EmployeeDedSchedules " & _
            " WHERE ClientCode = " & mCCode & " AND EmployeeID = " & txtEmpNum & mCon & _
            " AND DeductionStatus = " & ActiveDedFlag & _
            " ORDER BY EmployeeName, DeductionCode, DedStartPayment"
    Else ' selected deduction only
        StrQuery = "SELECT * FROM EmployeeDedSchedules " & _
            " WHERE ClientCode = " & mCCode & " AND DeductionCode = " & mDeductionCode & " AND EmployeeID = " & txtEmpNum & mCon & _
            " AND DeductionStatus = " & ActiveDedFlag & _
            " ORDER BY EmployeeName, DeductionCode, DedStartPayment"
    End If
End If
    
Set rsDed = ReturnRst(StrQuery, CnnstrHR, mstatic)
If rsDed.RecordCount = 0 Then
    ListView1(2).ListItems.Clear
    MsgBox "No deduction schedule found for!", vbInformation, PayrollAlert
    Set rsDed = Nothing
    Exit Sub
Else
    Call RetrieveDeductions(rsDed)
End If
Set rsDed = Nothing

If ListView1(2).ListItems.Count > 0 Then
    Call ExportDedToExcel(ListView1(2))
Else
    MsgBox "No records to print!", vbInformation, PayrollAlert
End If

End Sub

Private Sub ExportDedToExcel(ByRef lstDTRDisplay As ListView)
Dim rsDedRecords As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String, mRange2 As String
Dim sumString As String, prevName As String, prevDed As String
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

If chkAllEmp.VALUE = 0 Then
    mHead = "Payroll Deduction Entries Report for " & ListView1(0).SelectedItem.Text
Else
    mHead = "Payroll Deduction Entries Report for All Employees"
End If

Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy")
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)
mHead = "COMPANY NAME: " & cmbCompany.Text
Call PrintReportHeader(xlSheet, xlApp1, "A5", mHead, "", xlLeft, , False, False, False, , , 10, True)

Call PrintReportHeader(xlSheet, xlApp1, "A6:A7", "Record #", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "B6:B7", "Employee Name", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 16.5)
Call PrintReportHeader(xlSheet, xlApp1, "C6:C7", "Deduction", "Type", xlCenter, xlCenter, True, True, True, 15, , 8, False, 13)
Call PrintReportHeader(xlSheet, xlApp1, "D6:D7", "Payment", "Counter", xlCenter, xlCenter, True, True, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "E6:E7", "Running", "Balance", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "F6:F7", "Deduction", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "G6:G7", "Payroll", "Deduction", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "H6:H7", "Deduction", "Date", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
Call PrintReportHeader(xlSheet, xlApp1, "I6:I7", "Remarks", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 23)
        
prevName = lstDTRDisplay.ListItems.Item(1).SubItems(8)
prevDed = lstDTRDisplay.ListItems.Item(1).SubItems(1)
For i = 1 To lstDTRDisplay.ListItems.Count
    If prevName <> lstDTRDisplay.ListItems.Item(i).SubItems(8) Or _
        prevDed <> lstDTRDisplay.ListItems.Item(i).SubItems(1) Then
        ctr = ctr + 2
        prevName = lstDTRDisplay.ListItems.Item(i).SubItems(8)
        prevDed = lstDTRDisplay.ListItems.Item(i).SubItems(1)
    End If
        xlSheet.Cells(ctr, 1).VALUE = i
        xlSheet.Cells(ctr, 2).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(8) ' Name
        xlSheet.Cells(ctr, 3).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(1) ' Description
        xlSheet.Cells(ctr, 4).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(2) ' Counter
        xlSheet.Cells(ctr, 4).NumberFormat = "00"
        xlSheet.Cells(ctr, 5).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(3) ' Total Loan
        xlSheet.Cells(ctr, 5).NumberFormat = "0.00"
        xlSheet.Cells(ctr, 6).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(4) ' Deduction
        xlSheet.Cells(ctr, 6).NumberFormat = "0.00"
        xlSheet.Cells(ctr, 7).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(5) ' Payments
        xlSheet.Cells(ctr, 7).NumberFormat = "0.00"
        xlSheet.Cells(ctr, 8).VALUE = Format(lstDTRDisplay.ListItems.Item(i).SubItems(6), "mm/dd/yyyy") ' Date
        xlSheet.Cells(ctr, 9).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(7) ' Remarks
        ctr = ctr + 1
Next i

For l = 70 To 71  ' F to G
    mRange = Chr(l) & 8 & ":" & Chr(l) & ctr
    xlSheet.Range(mRange).Select
    mRange = Chr(l) & ctr + 1
    xlSheet.Range(mRange).Activate
    xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"
    xlSheet.Range(mRange).Font.Bold = True
Next l

ctr = ctr + 2

'Box it.
mRange = "A6:I" & ctr
Call PlaceBox(xlSheet, xlApp1, mRange)
With xlApp1.Range(mRange)
    .Font.Name = "Arial"   'Font.Name
    .Font.Size = 8         'Font.size
End With

xlSheet.Range("A1").Select

ExitRtn2:

End Sub

'Private Sub cmdList_Click(Index As Integer)
'If Index = 0 Then
'    Call EditRecord(Index)
'Else
'    Call EditRecord(Index)
'End If
'End Sub
'
'Private Sub EditRecord(Index As Integer)
'Dim res
'Dim mRem As String
'
'If Index = 0 Then ' Means edit
'    If cmdList(0).Caption = "Edit Selected Record" Then
'        If cmbTypes.Text = "" Then MsgBox "No Active Deductions, add new deduction first!": GoTo NextProc
'        cmbTypes.Enabled = False        ' do not let the user change it
'        Call EnableDisableDeductions(True)
'        cmdList(0).Caption = "Save Selected Record"
'        cmdList(1).Caption = "Cancel Edits"
'        mDeductionCode = GetDeductionCode(cmbTypes.Text)
'
'        cmdAddDed.Enabled = False
'        cmdEditDed.Enabled = False
'        cmdDelDed.Enabled = False
'        txtRates(17).SetFocus
'    Else    ' Save Edited record
'        Call SaveEditedRecord
'        cmdList(0).Caption = "Edit Selected Record"
'        cmdList(1).Caption = "Delete Selected Record"
'    End If
'Else ' Means Delete
'    If cmdList(1).Caption = "Delete Selected Record" Then
'        If cmbTypes.Text = "" Then MsgBox "No Active Deductions, cannot be deleted!": Exit Sub
'        mDeductionCode = GetDeductionCode(cmbTypes.Text)
'        If TotalPayments <> 0 Then
'            MsgBox "There were payments made amounting to " & TotalPayments & "." & Chr(13) & _
'                "Please proceed with Edit, cannot delete entries with active payments!", vbInformation, PayrollAlert
'        Else
'            res = MsgBox("Delete " & txtRates(18).Text & " deductions?", vbQuestion + vbYesNo, PayrollAlert)
'            If res = vbYes Then ' Delete current record
'                ' Delete records in deduction schedules.
'                If Action_Query(StrQuery, CnnstrHR) = True Then
'                    StrQuery = "DELETE EmployeeDedSchedules WHERE EmployeeID = " & txtEmpNum & _
'                        " AND DeductionCode = " & mDeductionCode & _
'                        " AND DeductionID = " & Val(txtDedID.Text) & " AND DeductionStatus = 1"
'                    Call Action_Query(StrQuery, CnnstrHR)
'                End If
'
'            End If
'        End If
'        Call ReQueryDeductions
'    Else ' Cancel changes to record
'        Call EnableDisableDeductions(False)
'        cmdList(0).Caption = "Edit Selected Record"
'        cmdList(1).Caption = "Delete Selected Record"
'
'        cmdAddDed.Enabled = True
'        cmdEditDed.Enabled = True
'        cmdDelDed.Enabled = True
'        MsgBox "Changes to deductions was cancelled!", vbInformation, PayrollAlert
'    End If
'
'End If
'
'NextProc:
'
'End Sub

'Private Sub SaveEditedRecord()
'Dim res, mPayment As Double, mDate1 As Date, mDate2 As Date
'
'' Validate the entries first.
'If MaskEdBox5 = "__/__/____" Or MaskEdBox5 = "01/01/1900" Then
'    MsgBox "Start Deduction Date should not be blank or equal to 01/01/1900!!!", vbInformation, PayrollAlert
'    MaskEdBox5.SetFocus
'    Exit Sub
'ElseIf MaskEdBox5 <> "__/__/____" And Not IsDate(MaskEdBox5) Then
'    MsgBox "Please enter a valid Start Deduction date!!!", vbInformation, PayrollAlert
'    MaskEdBox5.SetFocus
'    Exit Sub
'End If
'If MaskEdBox6 = "__/__/____" Or MaskEdBox6 = "01/01/1900" Then
'    MsgBox "End Deduction Date should not be blank or equal to 01/01/1900!!!", vbInformation, PayrollAlert
'    MaskEdBox6.SetFocus
'    Exit Sub
'ElseIf MaskEdBox6 <> "__/__/____" And Not IsDate(MaskEdBox6) Then
'    MsgBox "Please enter a valid End Deduction date!!!", vbInformation, PayrollAlert
'    MaskEdBox6.SetFocus
'    Exit Sub
'End If
'If Val(txtRates(21).Text) = 0# Then         ' And (MaskEdBox6 = "__/__/____" Or Format(Date, MaskEdBox6) = #1/1/1900#) Then
'    MsgBox "Number of Payments is required, to monitor payments made!", vbInformation, PayrollAlert
'    txtRates(21).SetFocus
'    Exit Sub
'End If
'
'rsEmpDeductions.Requery
'
'res = MsgBox("Do you want to save changes?", vbQuestion + vbYesNo, PayrollAlert)
'If res = vbNo Then
'    Exit Sub
'End If
'
'mPayment = IIf(IsNull(txtRates(18).Text), "0.00", txtRates(18).Text)        ' deduction per payroll
'mDate1 = MaskEdBox5  ' Get the start date
'mDate2 = MaskEdBox6  ' Get the end date
'
'StrQuery = "UPDATE EmployeeDedSchedules SET VoucherNo = '" & txtRates(23) & "', DeductionPayment = " & mPayment & ", DeductionAmount = " & txtRates(17) & _
'    ", DedStartPayment = '" & mDate1 & "', DedEndPayment = '" & mDate2 & "', TotalInterest = " & txtRates(15) & ", TotalDeduction = " & txtRates(16) & ", NoOfDeductions = " & IIf(txtRates(21) = "", 0#, txtRates(21)) & _
'    ", DeductionFlag = " & chkRates(12) & ", EntryBy = '" & mUserID & "', EntryDate = '" & Format(Now, "mm/dd/yyyy") & "', DeductionID = " & Val(txtDedID.Text) & _
'    " WHERE EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode & _
'    " AND DeductionFlag = 0 AND DeductionID = " & Val(txtDedID.Text)
'Call Action_Query(StrQuery, CnnstrHR)
'
'mDetails = "Deductions was edited!"
'Call SaveUserTransactions(mDetails, mRemarks)
'ListView1(2).ListItems.Clear
'MsgBox mstr, vbInformation, PayrollAlert
'
'Call ReQueryDeductions
'Call EnableDisableDeductions(False)
'cmdList(0).Caption = "Edit Selected Record"
'cmdList(1).Caption = "Delete Selected Record"
'
'cmdAddDed.Enabled = True
'cmdEditDed.Enabled = True
'cmdDelDed.Enabled = True
'cmbTypes.Enabled = True
'
'Exit Sub
'
'NextProc:
'
'    MsgBox Err.Description
'    Err.Clear
'
'End Sub

Private Sub cmdList_Click(Index As Integer) ' Updated 3/27/2016 for New Cutoff 26-10 and 11-25
Dim mDate As Date, mSched As Date
Dim i As Integer

With cmdList
If Index = 1 Then
    If MsgBox("Change Deduction Dates to New Cut-off?", vbQuestion + vbYesNo) = vbYes Then
        With ListView1(2)
        For i = 1 To .ListItems.Count
            .ListItems.Item(i).SELECTED = True
            mDate = Format(.SelectedItem.ListSubItems.Item(6).Text, "mm/dd/yyyy")    ' get the start date
            mSched = mDate
            If Day(mDate) = 16 Then     ' check if day 16, then change to day 25
                mSched = mDate + 9
            ElseIf Day(mDate) = 25 Then
                mSched = mDate - 9
            End If
            .SelectedItem.SubItems(6) = mSched
            StrQuery = "UPDATE EmployeeDedSchedules SET DedStartPayment = '" & mSched & "', DedEndPayment = '" & mSched & "'" & _
                " Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode & _
                " AND DeductionID = " & Val(txtDedID.Text) & " AND Remarks = 'For Deduction.'" & _
                " AND DedStartPayment = '" & mDate & "'"
            If mSched <> mDate Then
                If Action_Query(StrQuery, CnnstrHR) = False Then
                    MsgBox "Deduction date " & mDate & " failed to save. Try it again."
                End If
            End If
            .ListItems.Item(i).SELECTED = True
        Next i
        End With
    End If
End If
End With

End Sub

Private Sub cmdNavigation_Click(Index As Integer)
Call Navigation(rsEmployeeBasicData, Me.cmdNavigation(Index), _
    Me.cmdNavigation(0), Me.cmdNavigation(1), Me.cmdNavigation(2), Me.cmdNavigation(3))

If AddDeductions = True Or EditDeductions = True Then
        Call EnableDisableDeductions(False)
        Call EnableDisableDeductionButtons("New", "Edit", True)
        EditDeductions = False
        AddDeductions = False
        cmbTypes.Enabled = True
End If

'On Error Resume Next
If rsEmployeeBasicData.RecordCount <> 0 And SSTab1(1).Tab = 0 Then ' disable the add and edit buttons
    Call DisplayBasicData(ListView1(0).SelectedItem.Text, True)
End If
'Call ScrollRec

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

Private Sub cmdSaveComputation_Click()
If MsgBox("Do you want to save the computation for Payroll Deductions?", vbQuestion + vbYesNo) = vbYes Then
    Call ComputeDeductions(True)
    'Need to update also the entries in schedules.
    StrQuery = "UPDATE EmployeeDeductions SET DedEndPayment = '" & IIf(MaskEdBox6 = "__/__/____", #1/1/1900#, Format(MaskEdBox6, "mm/dd/yyyy")) & _
        "', NoOfDeductions = " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", EntryBy = '" & mUserID & "', EntryDate = '" & Format(Now, "mm/dd/yyyy") & _
        "' Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode
    Action_Query StrQuery, CnnstrHR
    
    Call ReQueryDeductions

'    ' Requery all Active Employee Deductions.
'    StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
'    'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ProtectedFlag = 0"
'    Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
'    Call FilterDeductions(txtEmpNum)
Else
    MsgBox "Computation of deductions not save for Payroll Schedules.", vbInformation, PayrollAlert
End If
End Sub

Private Sub cmdSaveEdit_Click()
'save the changes to table.

End Sub

Private Sub cmdSchedule_Click()
GetPaymentsMade (False)
End Sub

Private Function GetPaymentsMade(Optional PayFlag As Boolean = False, Optional wCount As Boolean = False) As Double
Dim rsDed As ADODB.Recordset

GetPaymentsMade = 0

'Get Deduction Schedules.
If PayFlag = False Then
    If chkRates(11).VALUE = 1 Then
        txtDedID.Text = Replace(mID(cmbTypes.Text, InStr(1, cmbTypes.Text, "(", vbTextCompare) + 1, 6), ")", "")
    End If
    ' Filter for Active or Inactive Schedules only.
    StrQuery = "SELECT * FROM EmployeeDedSchedules " & _
        " WHERE EmployeeID = " & txtEmpNum & " AND DeductionCode = " & mDeductionCode & _
        " AND DeductionID = " & Val(txtDedID.Text) & " AND DeductionStatus = " & ActiveDedFlag & _
        " ORDER BY EmployeeID, DeductionCode, DedStartPayment"
    
    Set rsDed = ReturnRst(StrQuery, CnnstrHR, mstatic)
    If rsDed.RecordCount = 0 Then
        ListView1(2).ListItems.Clear
        cmdSaveComputation.Visible = False
        cmdCancelComputation.Visible = False
        MsgBox "No deduction schedule found" & IIf(cmbTypes.Text = "", "!", " for " & cmbTypes.Text & "."), vbInformation, PayrollAlert
    Else
        Call RetrieveDeductions(rsDed)
        'MsgBox "Schedule was retrieved!", vbInformation, PayrollAlert
    End If

Else    ' Get all the payments made, check the deductionFlag for 0 value only.
    StrQuery = "SELECT SUM(PaymentsMade) As Payments, COUNT(PaymentsMade) As PayCount FROM EmployeeDedSchedules " & _
        " WHERE EmployeeID = " & txtEmpNum & " AND DeductionCode = " & mDeductionCode & " AND PaymentsMade <> 0 " & _
        " AND DeductionID = " & Val(txtDedID.Text) & " AND DeductionStatus = " & ActiveDedFlag
    Set rsDed = ReturnRst(StrQuery, CnnstrHR, mstatic)
    If rsDed.RecordCount <> 0 Then
        If wCount = True Then   ' return the payment count
            GetPaymentsMade = IIf(IsNull(rsDed!PayCount), 0, rsDed!PayCount)
        Else
            GetPaymentsMade = IIf(IsNull(rsDed!payments), 0, rsDed!payments)
        End If
    Else
        GetPaymentsMade = 0
    End If

End If

Set rsDed = Nothing

End Function

Private Sub RetrieveDeductions(ByRef rsDed As Recordset)
Dim mLsEmp As ListView
Dim litem1 As ListItem
Dim Dedctr As Integer
Dim mDate As Date
Dim mDed As Double, mPayment As Double
Dim mDesc As String

On Error GoTo err1

InActiveCtr = 0
ActiveCtr = 0
Set mLsEmp = ListView1(2)
mLsEmp.View = lvwReport
mLsEmp.ListItems.Clear

With rsDed
    Dedctr = 0
    mPayment = 0#
    mDed = 0#
    mDesc = !DeductionDescription
    While Not (.EOF Or .BOF)
        Dedctr = !DeductionCounter
        mPayment = IIf(IsNull(!DeductionPayment), "0.00", !DeductionPayment)  ' deduction per payroll
        mDate = Format(!DedStartpayment, "mm/dd/yyyy") ' get the start date
        Set litem1 = mLsEmp.ListItems.Add(, , Dedctr) ' Deduction Types
            litem1.SubItems(1) = IIf(IsNull(!DeductionDescription), "Blank Ded", !DeductionDescription)
            litem1.SubItems(2) = IIf(IsNull(!DeductionCounter), 0, !DeductionCounter) ' ded Counter
            litem1.SubItems(3) = Format(IIf(IsNull(!RunningBalance), "0.00", !RunningBalance), "##0.00") ' Running Balance
            litem1.SubItems(4) = Format(mPayment, "##0.00")
            litem1.SubItems(5) = Format(IIf(IsNull(!PaymentsMade), "0.00", !PaymentsMade), "##0.00") ' Get payments made
            litem1.SubItems(6) = mDate
            If cmbTypes.Text = "Other Deductions" Then
                litem1.SubItems(7) = IIf(IsNull(!Remarks), "Other Deductions.", !Remarks & " (" & Trim(!VoucherNo) & ")")
            Else
                litem1.SubItems(7) = IIf(IsNull(!Remarks), "Scheduled.", !Remarks)
            End If
            litem1.SubItems(8) = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
            litem1.SubItems(9) = !RecordSeq
        .MoveNext
    Wend
End With

Exit Sub

err1:

    MsgBox Err.Description

End Sub

'Private Sub RetrieveDeductionsManipulated(ByRef rsDed As Recordset)
'Dim mLsEmp As ListView
'Dim litem1 As ListItem
'Dim Dedctr As Integer
'Dim mDate As Date
'Dim mDed As Double, mPayment As Double
'Dim RemainingBal As Double, AccuPayments As Double, PerPayrollDed As Double
'Dim mDesc As String, ZeroPay As Boolean, mDeduc As Double
'
'On Error GoTo Err1
'
'InActiveCtr = 0
'ActiveCtr = 0
'Set mLsEmp = ListView1(2)
'mLsEmp.View = lvwReport
'mLsEmp.ListItems.Clear
'
'With rsDed
'    Dedctr = 0
'    mPayment = 0#
'    mDed = 0#
'    mDesc = !DeductionDescription
'    RemainingBal = !TotalDeduction
'    AccuPayments = 0#
'    ZeroPay = False
'
'    While Not (.EOF Or .BOF)
'        Dedctr = !DeductionCounter
'        mPayment = IIf(IsNull(!DeductionPayment), "0.00", !DeductionPayment)  ' deduction per payroll
'        mDeduc = IIf(IsNull(!PaymentsMade), "0.00", !PaymentsMade)
'
'        If cmbTypes.Text = "Other Deductions" Then
'            RemainingBal = Format(IIf(IsNull(!RunningBalance), "0.00", !RunningBalance), "##0.00")
'        Else
'            If Left(!Remarks, 6) = "Ded. (" Then
'                mDeduc = 0#
'            End If
'            If ZeroPay = True And mDeduc <> 0# Then
'                mPayment = mPayment + (PerPayrollDed - AccuPayments)
'                ZeroPay = False
'            ElseIf ZeroPay = True And mDeduc = 0# Then
'                RemainingBal = !TotalDeduction - (AccuPayments + mPayment)
'                mPayment = (PerPayrollDed - AccuPayments)
'            End If
'            If Dedctr = rsDed.RecordCount Then
'                mPayment = RemainingBal
'            End If
'        End If
'        mDate = Format(!DedStartpayment, "mm/dd/yyyy") ' get the start date
'        Set litem1 = mLsEmp.ListItems.Add(, , Dedctr) ' Deduction Types
'            litem1.SubItems(1) = IIf(IsNull(!DeductionDescription), "Blank Ded", !DeductionDescription)
'            litem1.SubItems(2) = IIf(IsNull(!DeductionCounter), 0, !DeductionCounter) ' ded Counter
'            'litem1.SubItems(3) = Format(IIf(IsNull(!RunningBalance), "0.00", !RunningBalance), "##0.00") ' Running Balance
'            litem1.SubItems(3) = Format(RemainingBal, "##0.00") ' Running Balance
'            litem1.SubItems(4) = Format(mPayment, "##0.00")
'            litem1.SubItems(5) = Format(mDeduc, "##0.00") ' Get payments made
'            litem1.SubItems(6) = mDate
'            If cmbTypes.Text = "Other Deductions" Then
'                litem1.SubItems(7) = IIf(IsNull(!Remarks), "Other Deductions.", !Remarks & " (" & Trim(!VoucherNo) & ")")
'            Else
'                litem1.SubItems(7) = IIf(IsNull(!Remarks), "Scheduled.", !Remarks)
'            End If
'            litem1.SubItems(8) = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
'            litem1.SubItems(9) = !RecordSeq
'
'        PerPayrollDed = PerPayrollDed + IIf(IsNull(!DeductionPayment), "0.00", !DeductionPayment)
'        If ZeroPay = False Then
'            AccuPayments = AccuPayments + mDeduc
'        Else
'            AccuPayments = AccuPayments + IIf(IsNull(!DeductionPayment), "0.00", !DeductionPayment)
'        End If
'        If IsNull(!TotalDeduction) Or !TotalDeduction = 0# Then
'            RemainingBal = 0#
'        Else
'            RemainingBal = !TotalDeduction - AccuPayments
'        End If
'        If mDeduc = 0# Then
'            ZeroPay = True
'        Else
'            ZeroPay = False
'        End If
'
'        .MoveNext
'
'    Wend
'End With
'
'Exit Sub
'
'Err1:
'
'    MsgBox Err.Description
'
'End Sub

'Private Sub RetrieveDeductionsOrigError(ByRef rsDed As Recordset)
'Dim mLsEmp As ListView
'Dim litem1 As ListItem
'Dim Dedctr As Integer
'Dim mDate As Date
'Dim mDed As Double, mPayment As Double
'Dim mDesc As String
'
'On Error GoTo err1
'
'InActiveCtr = 0
'ActiveCtr = 0
'Set mLsEmp = ListView1(2)
'mLsEmp.View = lvwReport
'mLsEmp.ListItems.Clear
'
'With rsDed
'    Dedctr = 0
'    mPayment = 0#
'    mDed = 0#
'    mDesc = !DeductionDescription
'    While Not (.EOF Or .BOF)
'        If chkRates(13).VALUE = 1 Then  ' Paid only.
'            Dedctr = !DeductionCounter
'            mPayment = IIf(IsNull(!DeductionPayment), "0.00", !DeductionPayment)  ' deduction per payroll
'            mDate = Format(!DedStartPayment, "mm/dd/yyyy") ' get the start date
'            Set litem1 = mLsEmp.ListItems.Add(, , Dedctr) ' Deduction Types
'                litem1.SubItems(1) = IIf(IsNull(!DeductionDescription), "Blank Ded", !DeductionDescription)
'                litem1.SubItems(2) = IIf(IsNull(!DeductionCounter), 0, !DeductionCounter) ' ded Counter
'                litem1.SubItems(3) = Format(IIf(IsNull(!TotalDeduction), "0.00", !TotalDeduction) - mDed, "##0.00") ' Total balance
'                litem1.SubItems(4) = Format(mPayment, "##0.00")
'                litem1.SubItems(5) = Format(IIf(IsNull(!PaymentsMade), "0.00", !PaymentsMade), "##0.00") ' Get payments made
'                litem1.SubItems(6) = mDate
'                litem1.SubItems(7) = IIf(IsNull(!Remarks), "Scheduled.", !Remarks)
'                litem1.SubItems(8) = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
'        Else
'            If mDesc <> !DeductionDescription Then
'                mDed = 0
'                mDesc = !DeductionDescription
'            End If
'            mDate = Format(!DedStartPayment, "mm/dd/yyyy") ' get the start date
'            Dedctr = !DeductionCounter
'            If (!TotalDeduction - mDed) > !DeductionPayment Then
'                mPayment = IIf(IsNull(!DeductionPayment), "0.00", !DeductionPayment)  ' deduction per payroll
'            Else
'                mPayment = !TotalDeduction - mDed
'            End If
'            Set litem1 = mLsEmp.ListItems.Add(, , Dedctr) ' Deduction Types
'                litem1.SubItems(1) = IIf(IsNull(!DeductionDescription), "Blank Ded", !DeductionDescription)
'                litem1.SubItems(2) = Dedctr ' Counter
'                litem1.SubItems(3) = Format(IIf(IsNull(!TotalDeduction), "0.00", !TotalDeduction) - mDed, "##0.00") ' Total balance
'                litem1.SubItems(4) = Format(mPayment, "##0.00")
'                litem1.SubItems(5) = Format(IIf(IsNull(!PaymentsMade), "0.00", !PaymentsMade), "##0.00") ' Get payments made
'                litem1.SubItems(6) = mDate
'                litem1.SubItems(7) = IIf(IsNull(!Remarks), "Scheduled.", !Remarks)
'                litem1.SubItems(8) = IIf(IsNull(!EmployeeName), "No Name", !EmployeeName)
'
'                mDed = mDed + Val(IIf(IsNull(!DeductionPayment), "0.00", !DeductionPayment))
'            End If
'        .MoveNext
'    Wend
'End With
'
'Exit Sub
'
'err1:
'
'    MsgBox Err.Description
'
'End Sub

Private Sub Form_Load()
Dim mComp As String
Dim mEmpN As String

    TimeKeepFlag = False
    If ActiveSw = False Then
        Me.Top = (mdiMain.Height - Me.Height) / 8
        Me.Left = (mdiMain.Width - Me.Width) / 2
    End If
    FormLoadSw = False
    ActiveDedFlag = 1
    
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
    Call EnableDisableDeductions(False)
    AddDeductions = False
    EditDeductions = False
    txtRates(16).Enabled = False
    
    chkActiveClient.Enabled = False
    If ActiveCtr = 0 Then
        Call DisplayBasicData("", True)
    Else
        Call DisplayBasicData(CurrEmployee, True)
    End If
    
    ' Get all Deduction Types.
    StrQuery = "SELECT * FROM DeductionTypes2 WHERE DeductionCode <> 0 ORDER BY DeductionDescription"
    Set rsDeductionTypes = ReturnRst(StrQuery, CnnstrHR, mstatic)
    ' Get all Active Employee Deductions from Query.
    StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
    'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ProtectedFlag = 0"
    Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
    ' Get all deductions.
    StrQuery = "SELECT * FROM DeductionTypes2 WHERE DeductionCode <> 0 ORDER BY DeductionDescription"
    Set rsGetDeductionTypes = ReturnRst(StrQuery, CnnstrHR, mstatic)
    
    ' Updated June 25, 2013.
    Call GetOtherDed
    
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
        frmBasicSalary.Refresh
    End If
    cmdadd.Enabled = False
    cmdedit.Enabled = False
    cmddelete.Enabled = False

End Sub

Private Sub GetEmployeeBasic()
Dim rsTaxCode As ADODB.Recordset

'If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'    StrQuery = "SELECT * FROM clienttable WHERE CLIENTCODE2 = 'FMS' AND clientstatus = 1 and clientname is not null ORDER BY clientname"
'Else
'    StrQuery = "SELECT * FROM clienttable WHERE clientstatus = 1 and clientname is not null ORDER BY clientname"
'End If
' Get all active company.
StrQuery = GetClientRecord(0, True, " ORDER BY CLientName")
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

End Sub

Private Sub FilterEmployee(ByVal mCompany As String)

On Error GoTo ExitSub1

rsClientData.MoveFirst
rsClientData.Find "clientname = '" & Trim(mCompany) & "'"
If rsClientData.EOF = True Then
    mCCode = 0
    Exit Sub
Else
    mCCode = rsClientData!ClientCode
    txtClientCode = rsClientData!ClientCode
    workdaysperyr = rsClientData!WorkingDaysperYear
    noofdayspermon = rsClientData!NoofDaysperMonth
End If
With rsEmployeeBasicData
    .Filter = ""
    .MoveFirst
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

ExitSub1:

End Sub

Private Sub FilterDeductions(ByVal mEmployeeID As Integer)
On Error GoTo err1

    With rsEmpDeductions
       ' .Filter = "EmployeeID = " & mEmployeeID
        .Filter = "EmployeeID = " & mEmployeeID & " AND DeductionStatus = " & ActiveDedFlag
        Call FillDeductionList(cmbTypes, rsEmpDeductions)  ' do not call it again
        If .RecordCount = 0 Then  'If .EOF
            Call EnableDisableDeductions("", True)
        'Else
        '    cmbTypes.ListIndex = mPrevIndex
        End If
    End With

Exit Sub

err1:
    
    cmbTypes.ListIndex = 0
    Err.Clear
    
End Sub

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
    
    txtClient(3).Enabled = TFalse
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
    ListView1(0).View = lvwReport
    mLsEmp.ListItems.Clear
    ListView1(0).ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
        If mRs!EmployeeStatus = True Then ' if active employee
            Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName))
            'litem1.SubItems(0) = mRs(0)
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

Private Function FillDeductionList(ByRef mCmb As ComboBox, _
                        mRs As Recordset, Optional FilterFlag As Boolean, _
                        Optional ComboList As String = "")
Dim mVar As String

mCmb.Clear
While Not (mRs.EOF Or mRs.BOF)
    'If chkRates(11).VALUE = 1 Then ' if Active only
    If ActiveDedFlag = 1 Then ' if Active only
        mVar = Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1)), "'", "''"))
    Else    ' Show old deduction information
        mVar = Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1)), "'", "''"))
        'mVar = Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1) & " # " & mRs!deductionid), "'", "''"))
        mVar = mVar & "(" & mRs!deductionid & ")"
    End If
    
    If FilterFlag = False Then
        mCmb.AddItem mVar
        'mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1)), "'", "''"))
        mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
    Else    ' need to filter the lists
        If InStr(1, ComboList, "[" & LTrim(RTrim(str(mRs.Fields(0)))) & "]", 1) = 0 Or mRs.Fields(0) = 8 Then ' always included 8 for Other Ded
            mCmb.AddItem mVar
            'mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1)), "'", "''"))
            mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
        End If
    End If
    mRs.MoveNext
Wend

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
frmEmployeeSearch.Show 1
    
End Sub

'Private Sub Save()
'Dim mstr As String
'Dim mDetails As String
'Dim mRemarks As String
'
'If txtEmpNum = "txtclientcode" Then
'    txtEmpNum = 0
'End If
'
'mCompany = Trim(cmbCompany)
'mEmpName = Trim(txtClient(2))
'For i = 0 To 7
'    txtClient(i) = Replace(txtClient(i), "'", "''")
'Next i
'For i = 0 To 14
'    If txtRates(i) = "" Then
'        txtRates(i) = 0#
'    End If
'Next i
'If Not IsDate(mskStartContract) Or Not IsDate(mskEndContract) Then
'    mskStartContract = Format(#1/1/1900#, "mm/dd/yyyy")
'    mskEndContract = Format(#1/1/1900#, "mm/dd/yyyy")
'End If
'If Not IsDate(mskStartSSS) Or Not IsDate(mskEndSSS) Then
'    mskStartSSS = Format(#1/1/1900#, "mm/dd/yyyy")
'    mskEndSSS = Format(#1/1/1900#, "mm/dd/yyyy")
'End If
'If Not IsDate(mskStartPbig) Or Not IsDate(mskEndPbig) Then
'    mskStartPbig = Format(#1/1/1900#, "mm/dd/yyyy")
'    mskEndPbig = Format(#1/1/1900#, "mm/dd/yyyy")
'End If
'If Not IsDate(mskStartVale) Or Not IsDate(mskEndVale) Then
'    mskStartVale = Format(#1/1/1900#, "mm/dd/yyyy")
'    mskEndVale = Format(#1/1/1900#, "mm/dd/yyyy")
'End If
'If AddFlag = True Then ' Add New Record
'    StrQuery = "INSERT INTO EmployeeBasicSalary (AgencyCode, ClientCode, EmployeeID, EmpNo, EmployeeStatus, StartContract, EndContract, NoofMonthsHired, MonthlyRate,DailyRate,  HourlyRate, " & _
'    " PaymentBasis, PaymentSystem, WithSSSDeduction, SSSRateperDay, WithPbigDeduction, PagibigRateperDay, WithPHealthDeduction, BasicSalaryFlag3, BasicSalaryFlag4, BasicSalaryFlag5, " & _
'    " SSSLoanAmount, SSSLoanPayment, StartSSSPayment, LastSSSPayment, PbigLoanAmount, PbigLoanPayment, StartPBigPayment, LastPBigPayment, ValeLoanAmount, ValeLoanPayment, StartValePayment, LastValePayment," & _
'    " PHealthRateperDay, WithTaxDeduction, WithEcolaPay, EcolaAmount, With13thMonthPay, Hourly13thMonthRate, AgencyBDOCardFee, AgencyTASAesla, ClientwServiceContract, " & _
'    " ClientwCompPackage, IfnoCompensationPackage, Remarks, empTaxcode, NotedBy, Updatedby, DateUpdated ) " & _
'    " VALUES ( 1," & txtClientCode & ", " & txtEmpNum & ", '" & txtClient(0) & "', 1, '" & IIf(mskStartContract = "__/__/____", #1/1/1900#, Format(mskStartContract, "mm/dd/yyyy")) & "', '" & IIf(mskEndContract = "__/__/____", #1/1/1900#, Format(mskEndContract, "mm/dd/yyyy")) & "', " & txtRates(3) & ", " & txtRates(0) & ", " & txtRates(1) & ", " & txtRates(2) & _
'    ", '" & Trim(cmbPayBasis) & "', '" & Trim(cmbPaySystem) & "', " & chkRates(3) & ", " & txtRates(7) & ", " & chkRates(5) & ", " & txtRates(9) & ", " & chkRates(4) & ", " & chkRates(8) & ", " & chkRates(9) & ", " & chkRates(10) & _
'    ", " & txtRates(11) & ", " & txtRates(12) & ", '" & IIf(mskStartSSS = "__/__/____", #1/1/1900#, Format(mskStartSSS, "mm/dd/yyyy")) & "', '" & IIf(mskEndSSS = "__/__/____", #1/1/1900#, Format(mskEndSSS, "mm/dd/yyyy")) & "', " & txtRates(14) & ", " & txtRates(13) & ", '" & IIf(mskStartPbig = "__/__/____", #1/1/1900#, Format(mskStartPbig, "mm/dd/yyyy")) & "', '" & IIf(mskEndPbig = "__/__/____", #1/1/1900#, Format(mskEndPbig, "mm/dd/yyyy")) & "', " & txtRates(20) & ", " & txtRates(19) & ", '" & IIf(mskStartVale = "__/__/____", #1/1/1900#, Format(mskStartVale, "mm/dd/yyyy")) & "', '" & IIf(mskEndVale = "__/__/____", #1/1/1900#, Format(mskEndVale, "mm/dd/yyyy")) & "'" & _
'    ", " & txtRates(8) & ", " & chkRates(0) & ", " & chkRates(1) & ", " & txtRates(4) & ", " & chkRates(2) & ", " & txtRates(5) & ", " & txtRates(6) & ", " & txtRates(10) & ", " & chkRates(6) & _
'    ", " & chkRates(7) & ", '" & txtClient(5) & "', '" & txtClient(6) & "', '" & Trim(cmbTaxCode) & "', '" & txtClient(3) & "', '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & "')"
'    mstr = "New record was saved!"
'ElseIf EditFlag = True Then ' Save edit
'    StrQuery = "UPDATE EmployeeBasicSalary SET StartContract = '" & Format(mskStartContract, "mm/dd/yyyy") & "', EndContract = '" & Format(mskEndContract, "mm/dd/yyyy") & "', NoofMonthsHired = " & txtRates(3) & ", MonthlyRate = " & txtRates(0) & ", DailyRate = " & txtRates(1) & ", HourlyRate = " & txtRates(2) & ", " & _
'    "PaymentBasis = '" & Trim(cmbPayBasis) & "', PaymentSystem = '" & Trim(cmbPaySystem) & "', WithSSSDeduction = " & chkRates(3) & ", SSSRateperDay = " & txtRates(7) & ", WithPbigDeduction = " & chkRates(5) & ", PagibigRateperDay = " & txtRates(9) & ", WithPHealthDeduction = " & chkRates(4) & ", " & _
'    "BasicSalaryFlag3 = " & chkRates(8) & ", BasicSalaryFlag4 = " & chkRates(9) & ", BasicSalaryFlag5 = " & chkRates(10) & ", " & _
'    "SSSLoanAmount = " & txtRates(11) & ", SSSLoanPayment = " & txtRates(12) & ", StartSSSPayment = '" & IIf(mskStartSSS = "__/__/____", #1/1/1900#, Format(mskStartSSS, "mm/dd/yyyy")) & "', LastSSSPayment = '" & IIf(mskEndSSS = "__/__/____", #1/1/1900#, Format(mskEndSSS, "mm/dd/yyyy")) & "', PBigLoanAmount = " & txtRates(14) & ", PbigLoanPayment = " & txtRates(13) & ", StartPBigPayment = '" & IIf(mskStartPbig = "__/__/____", #1/1/1900#, Format(mskStartPbig, "mm/dd/yyyy")) & "', LastPbigPayment = '" & IIf(mskEndPbig = Null, #1/1/1900#, Format(mskEndPbig, "mm/dd/yyyy")) & "', " & _
'    "ValeLoanAmount = " & txtRates(20) & ", ValeLoanPayment = " & txtRates(19) & ", StartValePayment = '" & IIf(mskStartVale = "__/__/____", #1/1/1900#, Format(mskStartVale, "mm/dd/yyyy")) & "', LastValePayment = '" & IIf(mskEndVale = "__/__/____", #1/1/1900#, Format(mskEndVale, "mm/dd/yyyy")) & "', " & _
'    "PHealthRateperDay = " & txtRates(8) & ", WithTaxDeduction = " & chkRates(0) & ", WithEcolaPay = " & chkRates(1) & ", EcolaAmount = " & txtRates(4) & ", With13thMonthPay = " & chkRates(2) & ", Hourly13thMonthRate = " & txtRates(5) & ", AgencyBDOCardFee = " & txtRates(6) & ", AgencyTASAesla = " & txtRates(10) & ", ClientwServiceContract = " & chkRates(6) & ", " & _
'    "ClientwCompPackage = " & chkRates(7) & ", IfnoCompensationPackage = '" & txtClient(5) & "', Remarks = '" & txtClient(6) & "', EmpTaxCode = '" & Trim(cmbTaxCode) & "', NotedBy = '" & txtClient(3) & "', Updatedby = '" & mUserID & "', DateUpdated = '" & Format(Now, "mm/dd/yyyy") & "'" & _
'    " WHERE EmployeeID = " & txtEmpNum 'txtClient(0)
'    mstr = "Edit was saved!"
'End If
'If Action_Query(StrQuery, cnnstr) = True Then
'    Call SaveUserTransactions(mDetails, mRemarks)
'End If
'MsgBox mstr, vbInformation, PayrollAlert
'Call Form_Load
'cmbCompany = mCompany
'Call DisplayBasicData(mEmpName, True)
'
'End Sub

Private Sub cmdCancel_Click()
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
        employeeMsg = "This will remove employee from Payroll list. Do you want to proceed?"
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
        'Set rsEmployeeBasicData = Nothing
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
    If EditRights = False Then
        MsgBox "You do not have rights to Edit Employee!", vbInformation, PayrollAlert
        Exit Sub
    End If
    
    EditFlag = True
    AmountSw = False
    'mRecNo = Text1
    Call DisableTextBoxes(True)
    Call DisableCommands(False)
    'SSTab1(1).TabVisible(1) = False
    'cmbPayBasis.SetFocus
    'cmbTaxCode.SetFocus
    cmbPaySystem.SetFocus
End Sub

'Private Sub cmdSave_Click()
'    Call Save
'    SSTab1(1).TabVisible(0) = True
'    SSTab1(1).TabVisible(1) = True
'End Sub

Private Sub lblExit_Click()
    SSTab1(2).Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
ActiveSw = False
FormLoadSw = False
End Sub

Private Sub ListView1_Click(Index As Integer)
Dim res
    
    If ActiveCtr = 0 Then
        Call DisplayBasicData("", False)
    Else
        If SSTab1(1).Tab = 1 Then ' in-active tab
            If ListView1(1).ListItems.Count = 0 Then
                Call DisplayBasicData("", False)
            Else
                Call DisplayBasicData(ListView1(1).SelectedItem.Text, False)
                CurrEmployee = ListView1(1).SelectedItem.Text
            End If
        Else    ' Active Tab
            If ListView1(0).ListItems.Count = 0 Then
                Call DisplayBasicData("", False)
            Else
                Call DisplayBasicData(ListView1(0).SelectedItem.Text, False)
                CurrEmployee = ListView1(0).SelectedItem.Text
                If AddDeductions = True Or EditDeductions = True Then
                    'res = MsgBox("Cancel changes to deductions?", vbQuestion + vbYesNo, PayrollAlert)
                    'If res = vbYes Then ' cancel the changes
                    Call EnableDisableDeductions(False)
                    Call EnableDisableDeductionButtons("New", "Edit", True)
                    EditDeductions = False
                    AddDeductions = False
                    cmbTypes.Enabled = True
                    MsgBox "Changes to deductions was cancelled!", vbInformation, PayrollAlert
                    'End If
                End If
            End If
        End If
    End If
    
'    If AddDeductions = False And EditDeductions = False And cmbTypes.ListCount <> 0 Then
'        cmbTypes.ListIndex = 0
'    End If
    
End Sub

'Private Sub ListView1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    ListView1(Index).SortKey = ColumnHeader.Index - 1
'    ListView1(Index).Sorted = True
'End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
Dim SaveDed As String

On Error GoTo ErrorExit1

    Call DisplayBasicData(Item.Text, False)
    CurrEmployee = Item.Text
    
    If cmbTypes.Text = "Other Deductions" Then
        txtRates(17).Text = Item.ListSubItems(3).Text   ' total
        txtRates(18).Text = Item.ListSubItems(4).Text   ' deduction
        MaskEdBox5.Text = Format(Item.ListSubItems(6).Text, "mm/dd/yyyy")     ' date
        txtRates(23).Text = Item.ListSubItems(7).Text   ' remarks
        RecordSeq = Item.ListSubItems(9).Text   ' remarks
    End If
    
    'Call cmdSchedule_Click

'    'call DisplayBasicData(Item.ListSubItems(1).Text, False)
'    If SSTab1(1).Tab = 1 Then
'        txtEmpNum = Item.ListSubItems(7).Text
' '       TxtControlNo = Item.ListSubItems(3).Text
'        txtEmpNo = Item.ListSubItems(3).Text
'        txtEmpName = Item.ListSubItems(1).Text
'        mCompany = Item.ListSubItems(2).Text
'    End If
ErrorExit1:

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

'
'Private Sub Option1_Click()
'    If Option1.VALUE = True Then
'        Frame5(1).Visible = True: Frame5(0).Visible = False: Frame5(2).Visible = False
'    ElseIf Option2.VALUE = True Then
'        Frame5(1).Visible = False: Frame5(0).Visible = True: Frame5(2).Visible = False
'    Else
'        Frame5(1).Visible = False: Frame5(0).Visible = False: Frame5(2).Visible = True
'    End If
'End Sub
'
'Private Sub Option2_Click()
''    If Option1.VALUE = False Then Frame5(1).Visible = False: Frame5(0).Visible = True Else Frame5(1).Visible = True: Frame5(0).Visible = False
'    If Option1.VALUE = True Then
'        Frame5(1).Visible = True: Frame5(0).Visible = False: Frame5(2).Visible = False
'    ElseIf Option2.VALUE = True Then
'        Frame5(1).Visible = False: Frame5(0).Visible = True: Frame5(2).Visible = False
'    Else
'        Frame5(1).Visible = False: Frame5(0).Visible = False: Frame5(2).Visible = True
'    End If
'End Sub
'
'Private Sub Option3_Click()
'    If Option1.VALUE = True Then
'        Frame5(1).Visible = True: Frame5(0).Visible = False: Frame5(2).Visible = False
'    ElseIf Option2.VALUE = True Then
'        Frame5(1).Visible = False: Frame5(0).Visible = True: Frame5(2).Visible = False
'    Else
'        Frame5(1).Visible = False: Frame5(0).Visible = False: Frame5(2).Visible = True
'    End If
'End Sub

'Private Sub SSTab1_Click(Index As Integer, PreviousTab As Integer)
'    If SSTab1(1).Tab = 1 Then ' disable the add and edit buttons
'        cmdadd.Enabled = False
'        cmdedit.Enabled = False
'        If InActiveCtr = 0 Then
'            cmdMarkActive.Enabled = False
'        Else
'            cmdMarkActive.Enabled = True
'        End If
'    Else
'        cmdadd.Enabled = True
'        cmdedit.Enabled = True
'    End If
'End Sub

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

On Error GoTo ExitSub1

'If rsEmployeeBasicData.RecordCount = 0 Then Exit Sub

With rsEmployeeBasicData
    If recSw = False Then
        .MoveFirst
        .Find "Fullname = '" & SelectedEmployee & "'"
    End If
    If Not .EOF And Not .BOF Then
        chkActiveClient = IIf(!EmployeeStatus = True, 1, 0)
        txtEmpNum = IIf(IsNull(!EmployeeID), 0, !EmployeeID)
        txtClient(0) = IIf(IsNull(!EmpNo2), "", !EmpNo2)
        txtClient(1) = IIf(IsNull(!EmploymentStatusDesc), "", !EmploymentStatusDesc)
        txtClient(2) = IIf(IsNull(!FullName), "", !FullName)
        txtClient(3) = IIf(IsNull(!NotedBy), "", !NotedBy)
        'txtClient(4) = IIf(IsNull(!EntryDate), "", Format(!EntryDate, "mm/dd/yyyy")) ' should be date updated
        txtClient(4) = IIf(IsNull(!DateUpdated), "", Format(!DateUpdated, "mm/dd/yyyy"))
        txtClient(5) = IIf(IsNull(!IfnoCompensationPackage), "", !IfnoCompensationPackage)
        txtClient(6) = IIf(IsNull(!Remarks), "", !Remarks)
        txtClient(7) = IIf(IsNull(!Updatedby), "", !Updatedby)

        If cmbPayBasis.Text = "Commission" Then
            lblclient(41).Caption = "Commission Rate"
        Else
            lblclient(41).Caption = "Monthly Rate"
        End If
        txtRates(0) = IIf(!MonthlyRate = "" Or IsNull(!MonthlyRate), 0#, Format(!MonthlyRate, "###0.00"))
        txtRates(1) = IIf(!DailyRate = "" Or IsNull(!DailyRate), 0#, Format(!DailyRate, "###0.00"))
        txtRates(2) = IIf(!HourlyRate = "" Or IsNull(!HourlyRate), 0#, Format(!HourlyRate, "###0.000"))
        txtRates(3) = IIf(!NoofMonthsHired = "" Or IsNull(!NoofMonthsHired), 0#, !NoofMonthsHired)
        ' Ecola Amount
        txtRates(4) = IIf(!EcolaAmount = "" Or IsNull(!EcolaAmount), 0#, Format(!EcolaAmount, "###0.00"))
        txtRates(5) = IIf(!Hourly13thMonthRate = "" Or IsNull(!Hourly13thMonthRate), 0#, Format(!Hourly13thMonthRate, "###0.00"))
        txtRates(6) = IIf(!AgencyBDOCardFee = "" Or IsNull(!AgencyBDOCardFee), 0#, Format(!AgencyBDOCardFee, "###0.00"))
        txtRates(7) = IIf(!SSSRateperDay = "" Or IsNull(!SSSRateperDay), 0#, Format(!SSSRateperDay, "###0.00"))
        txtRates(8) = IIf(!PHealthRateperDay = "" Or IsNull(!PHealthRateperDay), 0#, Format(!PHealthRateperDay, "###0.00"))
        txtRates(9) = IIf(!PagibigRateperDay = "" Or IsNull(!PagibigRateperDay), 0#, Format(!PagibigRateperDay, "###0.00"))
        txtRates(10) = IIf(!AgencyTASAesla = "" Or IsNull(!AgencyTASAesla), 0#, Format(!AgencyTASAesla, "###0.00"))

        txtRates(11) = IIf(!SSSLoanAmount = "" Or IsNull(!SSSLoanAmount), 0#, Format(!SSSLoanAmount, "###0.00"))
        txtRates(12) = IIf(!SSSLoanPayment = "" Or IsNull(!SSSLoanPayment), 0#, Format(!SSSLoanPayment, "###0.00"))
        txtRates(14) = IIf(!PbigLoanAmount = "" Or IsNull(!PbigLoanAmount), 0#, Format(!PbigLoanAmount, "###0.00"))
        txtRates(13) = IIf(!PbigLoanPayment = "" Or IsNull(!PbigLoanPayment), 0#, Format(!PbigLoanPayment, "###0.00"))
        txtRates(20) = IIf(!ValeLoanAmount = "" Or IsNull(!ValeLoanAmount), 0#, Format(!ValeLoanAmount, "###0.00"))
        txtRates(19) = IIf(!ValeLoanPayment = "" Or IsNull(!ValeLoanPayment), 0#, Format(!ValeLoanPayment, "###0.00"))

        cmbTaxCode = IIf(IsNull(!EmpTaxCode), "HF", !EmpTaxCode)
        'cmbTaxCode = IIf(IsNull(!EmpCivilStatus), "HF", !EmpCivilStatus)
        cmbPayBasis = IIf(IsNull(!PaymentBasis), "Daily", !PaymentBasis)
        paymentbase = IIf(IsNull(!PaymentBasis), "Daily", !PaymentBasis)
        cmbPaySystem = IIf(IsNull(!PaymentSystem), "ATM", !PaymentSystem)

        mskStartContract = IIf(IsNull(!StartContract) Or !StartContract = #1/1/1900#, "__/__/____", Format(!StartContract, "mm/dd/yyyy"))
        mskEndContract = IIf(IsNull(!endcontract) Or !endcontract = #1/1/1900#, "__/__/____", Format(!endcontract, "mm/dd/yyyy"))
        ' SSS Loan Dates
        mskStartSSS = IIf(IsNull(!StartSSSPayment) Or !StartSSSPayment = #1/1/1900#, "__/__/____", Format(!StartSSSPayment, "mm/dd/yyyy"))
        mskEndSSS = IIf(IsNull(!LastSSSPayment) Or !LastSSSPayment = #1/1/1900#, "__/__/____", Format(!LastSSSPayment, "mm/dd/yyyy"))
        mskStartPbig = IIf(IsNull(!StartPbigPayment) Or !StartPbigPayment = #1/1/1900#, "__/__/____", Format(!StartPbigPayment, "mm/dd/yyyy"))
        mskEndPbig = IIf(IsNull(!LastPbigPayment) Or !LastPbigPayment = #1/1/1900#, "__/__/____", Format(!LastPbigPayment, "mm/dd/yyyy"))
        mskStartVale = IIf(IsNull(!StartValePayment) Or !StartValePayment = #1/1/1900#, "__/__/____", Format(!StartValePayment, "mm/dd/yyyy"))
        mskEndVale = IIf(IsNull(!LastValePayment) Or !LastValePayment = #1/1/1900#, "__/__/____", Format(!LastValePayment, "mm/dd/yyyy"))

        chkRates(3) = IIf(!WithSSSDeduction = True, 1, 0)
        chkRates(5) = IIf(!WithPbigDeduction = True, 1, 0)
        chkRates(8) = IIf(!BasicSalaryFlag3 = True, 1, 0)
        chkRates(9) = IIf(!BasicSalaryFlag4 = True, 1, 0)
        chkRates(10) = IIf(!BasicSalaryFlag5 = True, 1, 0)
        chkRates(4) = IIf(!WithPHealthDeduction = True, 1, 0)
        chkRates(0) = IIf(!WithTaxDeduction = True, 1, 0)
        chkRates(1) = IIf(!WithEcolaPay = True, 1, 0)
        chkRates(2) = IIf(!With13thMonthPay = True, 1, 0)
        chkRates(6) = IIf(!ClientwServiceContract = True, 1, 0)
        chkRates(7) = IIf(!ClientwCompPackage = True, 1, 0)

        Call ReQueryDeductions

'        ' Requery all Active Employee Deductions.
'        StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
'        'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ProtectedFlag = 0"
'        Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
'        Call FilterDeductions(txtEmpNum)
    End If
End With

ExitSub1:

End Sub

Private Sub txtClient_Validate(Index As Integer, Cancel As Boolean)
    If txtClient(0) = "" Then MsgBox "Client name should not be blank!", vbInformation, PayrollAlert: Cancel = True
End Sub

Private Sub txtRates_GotFocus(Index As Integer)
    Sendkeys "{HOME}+{END}"
    txtRates(Index).BackColor = &HFFFFC0
End Sub

Private Sub txtRates_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strValid

If Index <> 23 Then
    strValid = "0123456789." & Chr(13) & Chr(8)
    If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
End If
If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub txtRates_LostFocus(Index As Integer)
    txtRates(Index).BackColor = &HFFFFFF
    If AmountSw = True And Index <> 23 Then
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
'        If mAmt > 99999.99 Then
'            MsgBox "Entries should not exceed 99,999.99. Please correct it!", vbInformation, PayrollAlert
'            CheckAmount = 0#
'            Exit Function
'        End If
    End If
    CheckAmount = mAmt
End Function

'Private Sub ScrollRec()
'Dim strFindMe As String
'Dim itmFound As ListItem
'
'    strFindMe = txtEmpNum 'txtClient(0)
'    If SSTab1(1).Tab = 1 Then
'        Set itmFound = ListView1(1).FindItem(strFindMe, 2, , lvwPartial)
'    Else
'        Set itmFound = ListView1(0).FindItem(strFindMe, 1, , lvwPartial)
'    End If
'    If itmFound Is Nothing Then
'    Else
'        itmFound.EnsureVisible
'        itmFound.SELECTED = True
'    End If
'End Sub

