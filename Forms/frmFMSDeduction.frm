VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmFMSDeduction 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FMS Contributions"
   ClientHeight    =   7665
   ClientLeft      =   2595
   ClientTop       =   2415
   ClientWidth     =   11325
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmFMSDeduction.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmFMSDeduction.frx":058A
   ScaleHeight     =   7665
   ScaleWidth      =   11325
   Begin TabDlg.SSTab SSTab1 
      Height          =   7125
      Index           =   1
      Left            =   0
      TabIndex        =   36
      Top             =   480
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   12568
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
      TabCaption(0)   =   "Driver and Conductor SSS Contributions"
      TabPicture(0)   =   "frmFMSDeduction.frx":0894
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
      Tab(0).Control(5)=   "txtDepartment"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDBName"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtBundleFlag"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDriver"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox txtDriver 
         Height          =   285
         Left            =   600
         TabIndex        =   201
         Text            =   "txtclientcode"
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBundleFlag 
         Height          =   285
         Left            =   2400
         TabIndex        =   200
         Text            =   "0"
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDBName 
         Height          =   285
         Left            =   1800
         TabIndex        =   199
         Text            =   "txtclientcode"
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDepartment 
         Height          =   285
         Left            =   1200
         TabIndex        =   198
         Text            =   "txtclientcode"
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   3
         Left            =   2925
         Picture         =   "frmFMSDeduction.frx":08B0
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   7230
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   2
         Left            =   2640
         Picture         =   "frmFMSDeduction.frx":0DE2
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "First"
         Top             =   7230
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   1
         Left            =   2355
         Picture         =   "frmFMSDeduction.frx":0EDC
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "First"
         Top             =   7230
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   0
         Left            =   2070
         Picture         =   "frmFMSDeduction.frx":0FD6
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "First"
         Top             =   7230
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
         Height          =   6585
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   11055
         Begin VB.CheckBox chkAll 
            Caption         =   "All Employees"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   840
            TabIndex        =   226
            Top             =   5760
            Width           =   1335
         End
         Begin VB.TextBox txtFund 
            Height          =   285
            Index           =   4
            Left            =   2280
            TabIndex        =   224
            Text            =   "10000"
            Top             =   5400
            Width           =   735
         End
         Begin VB.TextBox txtFund 
            Height          =   285
            Index           =   3
            Left            =   840
            TabIndex        =   212
            Text            =   "12000"
            Top             =   5400
            Width           =   735
         End
         Begin VB.CommandButton cmdExport 
            Caption         =   "Export to Excel"
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
            TabIndex        =   202
            Top             =   6120
            Width           =   2775
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
            Left            =   1680
            TabIndex        =   187
            Top             =   7200
            Width           =   2895
         End
         Begin VB.TextBox txtCancelSw 
            Height          =   285
            Left            =   2520
            TabIndex        =   148
            Text            =   "0"
            Top             =   120
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.CommandButton cmdDTRManual 
            Caption         =   "Export DTR to Excel Report"
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
            Left            =   -240
            TabIndex        =   147
            ToolTipText     =   "Manual DTR Entry or Updates"
            Top             =   6720
            Width           =   2895
         End
         Begin VB.CommandButton cmdLeaveApplication 
            BackColor       =   &H00FF0000&
            Caption         =   "Print Processed DTR for Payroll"
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
            Left            =   -240
            MaskColor       =   &H00FF0000&
            TabIndex        =   139
            ToolTipText     =   "Leave/Dayoff Application Form"
            Top             =   6960
            Width           =   2895
         End
         Begin VB.Frame Frame3 
            Height          =   765
            Left            =   3240
            TabIndex        =   53
            Top             =   120
            Width           =   7695
            Begin VB.TextBox txtBegBal 
               Height          =   285
               Index           =   4
               Left            =   5760
               TabIndex        =   223
               Text            =   "0.00"
               Top             =   400
               Width           =   975
            End
            Begin VB.CommandButton cmdEditBalance 
               Caption         =   "Edit"
               Height          =   315
               Left            =   6840
               Style           =   1  'Graphical
               TabIndex        =   222
               Top             =   360
               Width           =   705
            End
            Begin VB.CheckBox chkNoFilter 
               Caption         =   "All Employees"
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   1485
               TabIndex        =   150
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
               Left            =   4680
               MultiLine       =   -1  'True
               TabIndex        =   149
               Text            =   "frmFMSDeduction.frx":1508
               Top             =   400
               Width           =   975
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
               Left            =   5040
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
               Width           =   3375
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
               Text            =   "frmFMSDeduction.frx":1514
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
               Left            =   3660
               TabIndex        =   3
               Text            =   "Client Name"
               Top             =   400
               Width           =   975
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
               ItemData        =   "frmFMSDeduction.frx":1520
               Left            =   4080
               List            =   "frmFMSDeduction.frx":152A
               TabIndex        =   7
               Text            =   "cmbTaxCode"
               Top             =   400
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Beg. Balance:"
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
               Left            =   5760
               TabIndex        =   140
               Top             =   180
               Width           =   1455
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Civil Status:"
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
               Left            =   4680
               TabIndex        =   57
               Top             =   180
               Width           =   1425
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Employee ID:"
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
               Left            =   3660
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
         Begin VB.Frame Frame4 
            Height          =   5520
            Left            =   3240
            TabIndex        =   133
            Top             =   960
            Width           =   7695
            Begin VB.CommandButton cmdGenerate 
               Caption         =   "Read TRIPS Records"
               Height          =   315
               Left            =   5400
               Style           =   1  'Graphical
               TabIndex        =   217
               Top             =   2640
               Width           =   2145
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
               ItemData        =   "frmFMSDeduction.frx":153E
               Left            =   840
               List            =   "frmFMSDeduction.frx":1540
               Style           =   2  'Dropdown List
               TabIndex        =   216
               TabStop         =   0   'False
               Top             =   2655
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
               ItemData        =   "frmFMSDeduction.frx":1542
               Left            =   2760
               List            =   "frmFMSDeduction.frx":1544
               Sorted          =   -1  'True
               TabIndex        =   215
               TabStop         =   0   'False
               Text            =   "cmbYear"
               Top             =   2655
               Width           =   1365
            End
            Begin VB.TextBox txtFund 
               Height          =   285
               Index           =   2
               Left            =   5040
               TabIndex        =   211
               Text            =   "txtclientcode"
               Top             =   5160
               Width           =   975
            End
            Begin VB.TextBox txtFund 
               Height          =   285
               Index           =   1
               Left            =   3000
               TabIndex        =   210
               Text            =   "txtclientcode"
               Top             =   5160
               Width           =   975
            End
            Begin VB.TextBox txtFund 
               Height          =   285
               Index           =   0
               Left            =   600
               TabIndex        =   209
               Text            =   "200.00"
               Top             =   5160
               Width           =   855
            End
            Begin VB.CommandButton cmdCompute 
               Caption         =   "Compute SSS"
               Height          =   315
               Left            =   6120
               Style           =   1  'Graphical
               TabIndex        =   207
               Top             =   5160
               Width           =   1425
            End
            Begin VB.Frame Frame8 
               BorderStyle     =   0  'None
               Height          =   735
               Left            =   720
               TabIndex        =   180
               Top             =   6360
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
                  TabIndex        =   181
                  Top             =   240
                  Width           =   5025
               End
            End
            Begin VB.Frame Frame7 
               ForeColor       =   &H00800000&
               Height          =   735
               Left            =   0
               TabIndex        =   165
               Top             =   0
               Width           =   7695
               Begin VB.ComboBox cmbYear2 
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
                  ItemData        =   "frmFMSDeduction.frx":1546
                  Left            =   3000
                  List            =   "frmFMSDeduction.frx":1548
                  Sorted          =   -1  'True
                  TabIndex        =   219
                  TabStop         =   0   'False
                  Text            =   "cmbYear2"
                  Top             =   240
                  Width           =   1485
               End
               Begin VB.ComboBox cmbMonth2 
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
                  ItemData        =   "frmFMSDeduction.frx":154A
                  Left            =   840
                  List            =   "frmFMSDeduction.frx":154C
                  Style           =   2  'Dropdown List
                  TabIndex        =   218
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1485
               End
               Begin VB.CommandButton cmdDisplay 
                  Caption         =   "Display Payments"
                  Height          =   315
                  Left            =   5520
                  Style           =   1  'Graphical
                  TabIndex        =   197
                  Top             =   240
                  Width           =   2025
               End
               Begin VB.TextBox txtDTR 
                  Alignment       =   2  'Center
                  Height          =   285
                  Index           =   8
                  Left            =   6360
                  TabIndex        =   178
                  Text            =   "0.00"
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.TextBox txtDTR 
                  Alignment       =   2  'Center
                  Height          =   285
                  Index           =   7
                  Left            =   5760
                  TabIndex        =   176
                  Text            =   "0.00"
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   405
               End
               Begin VB.TextBox txtDTR 
                  Alignment       =   2  'Center
                  Height          =   285
                  Index           =   6
                  Left            =   5760
                  TabIndex        =   170
                  Text            =   "0.00"
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.TextBox txtDTR 
                  Alignment       =   2  'Center
                  Height          =   285
                  Index           =   5
                  Left            =   5160
                  TabIndex        =   169
                  Text            =   "0.00"
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.TextBox txtDTR 
                  Alignment       =   2  'Center
                  Height          =   285
                  Index           =   4
                  Left            =   4800
                  TabIndex        =   168
                  Text            =   "0.00"
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.CheckBox chkAllNoDTR 
                  Caption         =   "All blank DTR"
                  ForeColor       =   &H00800000&
                  Height          =   255
                  Left            =   6120
                  TabIndex        =   175
                  Top             =   1005
                  Width           =   1335
               End
               Begin VB.Label lblclient 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Month:"
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
                  TabIndex        =   221
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.Label lblclient 
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
                  ForeColor       =   &H00800000&
                  Height          =   255
                  Index           =   46
                  Left            =   2400
                  TabIndex        =   220
                  Top             =   225
                  Width           =   795
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
                  Left            =   6480
                  TabIndex        =   179
                  Top             =   960
                  Visible         =   0   'False
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
                  Left            =   6000
                  TabIndex        =   173
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   315
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
                  Left            =   5400
                  TabIndex        =   172
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   435
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
                  Left            =   4920
                  TabIndex        =   171
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   435
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
                  TabIndex        =   167
                  Top             =   840
                  Width           =   1635
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
                  TabIndex        =   166
                  Top             =   840
                  Width           =   315
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "Change Cutoff Date"
               ForeColor       =   &H00800000&
               Height          =   1815
               Left            =   0
               TabIndex        =   154
               Top             =   6360
               Visible         =   0   'False
               Width           =   7455
               Begin VB.CheckBox chkInitCutoff 
                  Caption         =   "Use initial Cutoff"
                  ForeColor       =   &H00800000&
                  Height          =   255
                  Left            =   5280
                  TabIndex        =   186
                  Top             =   1320
                  Width           =   1935
               End
               Begin VB.CommandButton cmdCancelCutoff 
                  Height          =   405
                  Left            =   6360
                  Picture         =   "frmFMSDeduction.frx":154E
                  Style           =   1  'Graphical
                  TabIndex        =   164
                  ToolTipText     =   "Cancel"
                  Top             =   480
                  Width           =   465
               End
               Begin VB.CommandButton cmdSaveCutoff 
                  Height          =   405
                  Left            =   5520
                  Picture         =   "frmFMSDeduction.frx":1A80
                  Style           =   1  'Graphical
                  TabIndex        =   155
                  ToolTipText     =   "Save Cutoff"
                  Top             =   480
                  Width           =   465
               End
               Begin MSComCtl2.DTPicker DTStartFirstCutoff 
                  Height          =   375
                  Left            =   1920
                  TabIndex        =   156
                  Top             =   315
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   661
                  _Version        =   393216
                  Format          =   61997057
                  CurrentDate     =   40262
               End
               Begin MSComCtl2.DTPicker DTEndFirstCutoff 
                  Height          =   375
                  Left            =   3720
                  TabIndex        =   157
                  Top             =   315
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   661
                  _Version        =   393216
                  Format          =   61997057
                  CurrentDate     =   40262
               End
               Begin MSComCtl2.DTPicker DTStartSecondCutoff 
                  Height          =   375
                  Left            =   1920
                  TabIndex        =   158
                  Top             =   720
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   661
                  _Version        =   393216
                  Format          =   61997057
                  CurrentDate     =   40262
               End
               Begin MSComCtl2.DTPicker DTEndSecondCutoff 
                  Height          =   375
                  Left            =   3720
                  TabIndex        =   159
                  Top             =   720
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   661
                  _Version        =   393216
                  Format          =   61997057
                  CurrentDate     =   40262
               End
               Begin MSComCtl2.DTPicker DTIniStartCutoff 
                  Height          =   375
                  Left            =   1920
                  TabIndex        =   182
                  Top             =   1200
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   661
                  _Version        =   393216
                  Format          =   61997057
                  CurrentDate     =   40262
               End
               Begin MSComCtl2.DTPicker DTIniEndCutoff 
                  Height          =   375
                  Left            =   3720
                  TabIndex        =   183
                  Top             =   1200
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   661
                  _Version        =   393216
                  Format          =   61997057
                  CurrentDate     =   40262
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
                  TabIndex        =   185
                  Top             =   1320
                  Width           =   1635
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
                  TabIndex        =   184
                  Top             =   1320
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
                  TabIndex        =   163
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
                  Index           =   50
                  Left            =   3360
                  TabIndex        =   162
                  Top             =   480
                  Width           =   315
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
                  TabIndex        =   161
                  Top             =   465
                  Width           =   1275
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
                  TabIndex        =   160
                  Top             =   840
                  Width           =   1635
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
               Top             =   5700
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
               Top             =   5700
               Width           =   1305
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   1095
               Index           =   1
               Left            =   480
               TabIndex        =   137
               Top             =   6480
               Visible         =   0   'False
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   1931
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
            Begin MSComctlLib.ListView lstPayments 
               Height          =   2010
               Index           =   0
               Left            =   120
               TabIndex        =   151
               Top             =   3000
               Width           =   7440
               _ExtentX        =   13123
               _ExtentY        =   3545
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
                  Text            =   "Trans Date"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Bundle No"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Commission"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Surety Fund"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Bus No"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Route"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Object.Width           =   4
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Object.Width           =   9
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Object.Width           =   4
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Object.Width           =   4
               EndProperty
               BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   10
                  Object.Width           =   4
               EndProperty
               BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   11
                  Object.Width           =   4
               EndProperty
               BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   12
                  Object.Width           =   4
               EndProperty
               BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   13
                  Object.Width           =   5
               EndProperty
            End
            Begin MSComctlLib.ListView lstPayments 
               Height          =   1770
               Index           =   1
               Left            =   120
               TabIndex        =   208
               Top             =   840
               Width           =   7440
               _ExtentX        =   13123
               _ExtentY        =   3122
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
               NumItems        =   15
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Payment Date"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Trans Date"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Beg Bal."
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Commission"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Surety Fund"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "SSS"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "PhilHealth"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "HDMF"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "Tax"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Text            =   "Advance"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   10
                  Text            =   "SSS Loan"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   11
                  Text            =   "HDMF Loan"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   12
                  Text            =   "Particulars"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   13
                  Text            =   "Others"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   14
                  Object.Width           =   5
               EndProperty
            End
            Begin VB.Label lblclient 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   54
               Left            =   2280
               TabIndex        =   214
               Top             =   2640
               Width           =   795
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Month:"
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
               Left            =   240
               TabIndex        =   213
               Top             =   2655
               Width           =   1155
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Fund:"
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
               Left            =   4080
               TabIndex        =   206
               Top             =   5160
               Width           =   1035
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Commission:"
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
               Left            =   1560
               TabIndex        =   205
               Top             =   5160
               Width           =   1395
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Tax:"
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
               Left            =   120
               TabIndex        =   204
               Top             =   5160
               Width           =   555
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "NDiff:"
               Height          =   255
               Index           =   6
               Left            =   6600
               TabIndex        =   177
               Top             =   5805
               Width           =   855
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "OT:"
               Height          =   255
               Index           =   5
               Left            =   5640
               TabIndex        =   146
               Top             =   5805
               Width           =   855
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "UDT:"
               Height          =   255
               Index           =   4
               Left            =   4680
               TabIndex        =   145
               Top             =   5805
               Width           =   855
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "Lates:"
               Height          =   255
               Index           =   3
               Left            =   3600
               TabIndex        =   144
               Top             =   5805
               Width           =   975
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "Absent:"
               Height          =   255
               Index           =   2
               Left            =   2400
               TabIndex        =   143
               Top             =   5805
               Width           =   975
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "Pay Hrs:"
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   142
               Top             =   5805
               Width           =   1215
            End
            Begin VB.Label lblTimeKeep 
               Caption         =   "Days:"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   141
               Top             =   5805
               Width           =   855
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
            Left            =   480
            TabIndex        =   108
            Top             =   6960
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
            TabIndex        =   102
            Text            =   "Client Name"
            Top             =   5170
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
            Left            =   480
            TabIndex        =   101
            Top             =   6720
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
            Left            =   480
            TabIndex        =   100
            Top             =   6840
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
            Left            =   2010
            TabIndex        =   29
            Text            =   "Client Name"
            Top             =   6720
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
            Left            =   300
            TabIndex        =   28
            Top             =   6975
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
            Left            =   300
            TabIndex        =   27
            Top             =   6600
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
            Left            =   1020
            TabIndex        =   30
            Text            =   "Client Name"
            Top             =   7035
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
            Left            =   7470
            TabIndex        =   34
            Text            =   "Client Name"
            Top             =   5170
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
            Left            =   3720
            TabIndex        =   31
            Text            =   "Client Name"
            Top             =   7035
            Width           =   1365
         End
         Begin VB.Frame Frame1 
            Height          =   1455
            Left            =   11400
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
               ItemData        =   "frmFMSDeduction.frx":1FB2
               Left            =   180
               List            =   "frmFMSDeduction.frx":1FC2
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
               ItemData        =   "frmFMSDeduction.frx":1FEA
               Left            =   180
               List            =   "frmFMSDeduction.frx":1FFD
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
            ItemData        =   "frmFMSDeduction.frx":201E
            Left            =   180
            List            =   "frmFMSDeduction.frx":202B
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   450
            Width           =   3045
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4215
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   1080
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   7435
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
         Begin VB.Frame Frame5 
            Height          =   1695
            Index           =   1
            Left            =   11520
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
            Left            =   11400
            TabIndex        =   109
            Top             =   120
            Visible         =   0   'False
            Width           =   5415
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
               ItemData        =   "frmFMSDeduction.frx":204E
               Left            =   1470
               List            =   "frmFMSDeduction.frx":205B
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
            Left            =   11760
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
         Begin VB.TextBox txtBundle 
            BackColor       =   &H00FFFFC0&
            Height          =   320
            Index           =   2
            Left            =   1080
            TabIndex        =   195
            Top             =   2880
            Width           =   1935
         End
         Begin VB.TextBox txtBundle 
            BackColor       =   &H00FFFFC0&
            Height          =   320
            Index           =   0
            Left            =   1080
            TabIndex        =   190
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox txtBundle 
            BackColor       =   &H00FFFFC0&
            Height          =   320
            Index           =   1
            Left            =   1080
            TabIndex        =   191
            Top             =   2520
            Width           =   1935
         End
         Begin VB.TextBox txtBundle 
            BackColor       =   &H00FFFFC0&
            Height          =   320
            Index           =   3
            Left            =   1080
            TabIndex        =   196
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdBundleNo 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   2640
            TabIndex        =   188
            Top             =   2160
            Width           =   300
         End
         Begin VB.Label lblclient 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bond:"
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
            Left            =   1440
            TabIndex        =   225
            Top             =   5400
            Width           =   675
         End
         Begin VB.Label lblclient 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ceiling:"
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
            Left            =   120
            TabIndex        =   203
            Top             =   5400
            Width           =   675
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Route.:"
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
            Index           =   71
            Left            =   120
            TabIndex        =   194
            Top             =   3300
            Width           =   855
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Bus No.:"
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
            Index           =   70
            Left            =   120
            TabIndex        =   193
            Top             =   2940
            Width           =   855
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Dept.:"
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
            Left            =   120
            TabIndex        =   192
            Top             =   2580
            Width           =   855
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Bundle No.:"
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
            Left            =   120
            TabIndex        =   189
            Top             =   2220
            Width           =   855
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   1080
            X2              =   3975
            Y1              =   6600
            Y2              =   6615
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
            Left            =   1650
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
            Left            =   3000
            TabIndex        =   58
            Top             =   7095
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
            Left            =   480
            TabIndex        =   41
            Top             =   6765
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
            Left            =   300
            TabIndex        =   40
            Top             =   7095
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
            Picture         =   "frmFMSDeduction.frx":207E
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
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CheckBox chkEditDTR 
         Height          =   375
         Left            =   1200
         Picture         =   "frmFMSDeduction.frx":20DC
         Style           =   1  'Graphical
         TabIndex        =   174
         ToolTipText     =   "Edit DTR Records"
         Top             =   30
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdEdit 
         Height          =   405
         Left            =   6240
         Picture         =   "frmFMSDeduction.frx":251E
         Style           =   1  'Graphical
         TabIndex        =   153
         ToolTipText     =   "Save"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CheckBox chkChangeCutoff 
         Height          =   375
         Left            =   10320
         Picture         =   "frmFMSDeduction.frx":2A50
         Style           =   1  'Graphical
         TabIndex        =   152
         ToolTipText     =   "Update Payroll Cut-off Dates"
         Top             =   30
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdDeductionTypes 
         Height          =   405
         Left            =   5400
         Picture         =   "frmFMSDeduction.frx":2E92
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Deduction Types"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   9600
         Picture         =   "frmFMSDeduction.frx":32D4
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Delete"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   120
         Picture         =   "frmFMSDeduction.frx":33D6
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   8160
         Picture         =   "frmFMSDeduction.frx":3520
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "New"
         Top             =   0
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   8640
         Picture         =   "frmFMSDeduction.frx":3A52
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Save"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   9120
         Picture         =   "frmFMSDeduction.frx":3F84
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Cancel"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdClose 
         Height          =   405
         Left            =   600
         Picture         =   "frmFMSDeduction.frx":44B6
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
Attribute VB_Name = "frmFMSDeduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer

Dim rsEmployeeBasicData As ADODB.Recordset
Dim rsInactiveEmployee As ADODB.Recordset
Dim rsClientData As ADODB.Recordset

Dim StrQuery As String
Dim mRecNo As Integer
Dim mCCode As Integer
Dim mCompany As String
Dim mEmpName As String
Dim ActiveCtr As Integer
Dim InActiveCtr As Integer
Dim RefreshSw As Boolean
Dim CurrEmployee As String
Dim CurrAmount As String    ' holds the current value of salaries
Dim AmountSw As Boolean     ' variable for basic salary entries
Dim ActiveSw As Boolean
Dim FormLoadSw As Boolean

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

Private Sub cmdBegBal_Click()

StrQuery = "EXEC dbo.GetFMSCommissions '" & txtDriver.Text & "', " & cmbYear & ", '" & cmbMonth.Text & _
        "', '" & txtDBName & "', '" & txtDepartment.Text & "', '" & txtClient(0).Text & "'"

If Action_Query(StrQuery, cnnstr) = False Then
    MsgBox "Failed to get records!", vbInformation, PayrollAlert
End If

End Sub

Private Sub cmbMonth2_Click()
cmdExport.Caption = "Export " & cmbMonth2.Text & " to Excel"
End Sub

Private Sub cmdBundleNo_Click()
txtBundleFlag = 0
'Call frmBundleNo.Show(1)
End Sub

Private Sub cmdClose_Click()
Set rsEmployeeBasicData = Nothing
Unload Me
End Sub

'Private Sub GenerateDTRReport()
'Dim rsDedRecords As ADODB.Recordset
'Dim xlApp1 As Excel.Application
'Dim xlBook1 As Excel.Workbook
'Dim xlSheet As Excel.Worksheet
'Dim mTotalPay As Double
'Dim mTotalDeduction As Double
'Dim mRange As String, mRange2 As String
'Dim sumString As String
'Dim ctr As Long
'Dim i As Integer, l As Integer
'Dim cPos As Integer
'Dim LineCtr As Integer
'Dim mHead As String
'Dim mDeduc As Double
'Dim StartDeduction As Variant
'Dim EndDeduction As Variant
'Dim DeductionYear As Integer
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
'ctr = 8
'i = 1
'
''Updated by Bernard May 13, 2008.
'mHead = CompanyName & " (" & CompanyShort & ")"
'Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 10, True)
'
'mHead = "DTR Records Report for " & ListView1(0).SelectedItem.Text
'Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
'mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy")
'Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)
'mHead = "COMPANY NAME: " & cmbCompany.Text
'Call PrintReportHeader(xlSheet, xlApp1, "A5", mHead, "", xlLeft, , False, False, False, , , 10, True)
'
'Call PrintReportHeader(xlSheet, xlApp1, "A6:A7", "Record #", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "B6:B7", "DTR ID", "", xlCenter, xlCenter, True, False, True, 15, , 9, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "C6:C7", "DTR Day", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "D6:D7", "DTR Date", "", xlCenter, xlCenter, True, True, True, 15, , 9, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "E6:E7", "Worked", "Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "F6:F7", "Time In", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "G6:G7", "Time Out", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "H6:H7", "No. Hours", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "I6:I7", "Absent", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "J6:J7", "Late", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "K6:K7", "Undertime", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "L6:L7", "Night Diff", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 10)
'Call PrintReportHeader(xlSheet, xlApp1, "M6:M7", "Remarks", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 25)
'
'For i = 1 To .ListItems.Count
'        xlSheet.Cells(ctr, 1).VALUE = i
'        xlSheet.Cells(ctr, 2).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(3) ' DTR ID
'        xlSheet.Cells(ctr, 2).NumberFormat = "00000000"
'        xlSheet.Cells(ctr, 3).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(1) ' DTR Day
'        xlSheet.Cells(ctr, 4).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(2) ' DTR Date
'        xlSheet.Cells(ctr, 5).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(4) ' Work Hours
'        xlSheet.Cells(ctr, 5).NumberFormat = "0.00"
'        xlSheet.Cells(ctr, 6).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(5) ' Time In
'        xlSheet.Cells(ctr, 7).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(6) ' Time Out
'        xlSheet.Cells(ctr, 8).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(7) ' No Hours
'        xlSheet.Cells(ctr, 9).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(8) ' Absent
'        xlSheet.Cells(ctr, 10).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(9) ' Late
'        xlSheet.Cells(ctr, 11).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(10) ' Undertime
'        xlSheet.Cells(ctr, 12).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(12) ' Night Diff
'        xlSheet.Cells(ctr, 13).VALUE = lstDTRDisplay.ListItems.Item(i).SubItems(13) ' remarks
'        ctr = ctr + 1
'Next i
'
'For l = 69 To 76  ' E to R
'    mRange = Chr(l) & 8 & ":" & Chr(l) & ctr
'    xlSheet.Range(mRange).Select
'    mRange = Chr(l) & ctr + 1
'    xlSheet.Range(mRange).Activate
'    If l = 70 Or l = 71 Then ' 70
'    Else
'        xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"
'    End If
'Next l
'
'ctr = ctr + 2
'
''Box it.
'mRange = "A6:M" & ctr
'Call PlaceBox(xlSheet, xlApp1, mRange)
'
''ctr = ctr + 1
''mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
''mRange = "A" & ctr
''Call CustomizedHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 8, False, 25)
'
'xlSheet.Range("A1").Select
'
'ExitRtn2:
'
'End Sub

Private Sub cmdDTRPayroll_Click()
Dim res

res = MsgBox("Include DTR of " & txtClient(2).Text & " for Payroll processing?", vbQuestion + vbYesNo, PayrollAlert)
If res = vbYes Then ' mark for payroll processing

Else
    MsgBox "DTR importing was cancelled!", vbInformation, PayrollAlert
End If

End Sub

Private Sub cmdCompute_Click()
Dim rsTrips As ADODB.Recordset
Dim TransMonth As String
Dim res

'CREATE PROCEDURE dbo.DriverConductorSSSComputation
'@CeilingFund  Float,        -- Fund Ceiling
'@TotCommission Float,       -- Total Commission
'@TotFund Float,     -- Total Fund
'@mMonthName Varchar(20),
'@mYear int,
'@EmpName Varchar(50),
'@EmpID Varchar(20),
'@TaxCode Varchar(20),
'@BegBalance Float,
'@DBName Varchar(30),
'@Department Varchar(30)

'StrQuery = "EXEC dbo.DriverConductorSSSComputation " & Val(txtFund(3).Text) & ", " & Val(txtFund(1).Text) & ", " & Val(txtFund(2).Text) & _
'        ", '" & cmbMonth.Text & "', " & cmbYear & ", '" & txtClient(2).Text & "', '" & txtClient(0).Text & "', '" & txtClient(9).Text & "', " & _
'        txtBegBal(4).Text & ", '" & txtDBName & "', '" & txtDepartment.Text & "'"

'TransMonth = Mid(cmbMonth.Text, 4, 3) & " " & str(cmbYear)

StrQuery = "SELECT PaymentDate, AppliedMonth, AppliedYear FROM DriverConductorPayments " & _
    "WHERE EmployeeID = '" & Trim$(txtClient(0)) & "' AND AppliedMonth = " & Val(Left(cmbMonth.Text, 2)) & " AND AppliedYear = " & cmbYear.Text
Set rsTrips = ReturnRst(StrQuery, cnnstr, mstatic)
If rsTrips.RecordCount > 0 Then
    res = MsgBox(txtClient(2).Text & " already processed! Do you want to DELETE and Reprocess?", vbQuestion + vbYesNo, PayrollAlert)
    If res = vbNo Then ' do not proceed
        Call RefreshPayments(Val(Mid(cmbMonth.Text, 1, 2)), cmbYear.Text)
        Set rsTrips = Nothing
        Exit Sub
    Else    ' Delete first the record.
        StrQuery = "DELETE DriverConductorPayments " & _
            "WHERE EmployeeID = '" & Trim$(txtClient(0)) & "' AND AppliedMonth = " & Val(Left(cmbMonth.Text, 2)) & " AND AppliedYear = " & cmbYear.Text
        If Action_Query(StrQuery, cnnstr) = False Then
            MsgBox "Failed to delete record!", vbInformation, PayrollAlert
            Exit Sub
        End If
    End If
End If
rsTrips.Close

TransMonth = Mid(cmbMonth.Text, 4, 15)
StrQuery = "INSERT INTO DriverConductorPayments (ClientCode, EmployeeID, FullName, Positions, Department, TaxCode, " & _
        "PaymentDate, TransMonth, TransYear, BegBalance, Commission, SuretyFund, AppliedMonth, AppliedYear) VALUES (" & _
        mCCode & ", '" & txtClient(0) & "', '" & txtClient(2) & "', '" & Trim$(txtDriver) & "', '" & txtDepartment.Text & "', '" & txtClient(9) & "', '" & _
        Format(Now, "mm/dd/yyyy") & "', '" & TransMonth & "', " & cmbYear.Text & ", " & Replace(txtBegBal(4), ",", "") & ", " & Replace(txtFund(1), ",", "") & ", " & Replace(txtFund(2), ",", "") & ", " & Val(Left(cmbMonth.Text, 2)) & ", " & cmbYear.Text & ")"

If Action_Query(StrQuery, cnnstr) = False Then
    MsgBox "Failed to insert new record!", vbInformation, PayrollAlert
Else ' Compute for SSS, HDMF, Philhealth, Tax
    StrQuery = "EXEC dbo.DriverConductorSSSComputation " & Val(txtFund(3).Text) & ", " & Val(txtFund(1).Text) & ", " & Val(txtFund(2).Text) & _
        ", '" & cmbMonth.Text & "', " & cmbYear & ", '" & txtClient(2).Text & "', '" & txtClient(0).Text & "', '" & txtClient(9).Text & "', " & _
        txtFund(4).Text & ", '" & txtDBName & "', '" & txtDepartment.Text & "', " & Val(txtFund(0).Text) & ", '" & mUserID & "', " & Val(txtFund(0).Text)
    If Action_Query(StrQuery, cnnstr) = False Then
        MsgBox "Failed to update record!", vbInformation, PayrollAlert
    Else
        Call RefreshPayments(Left(cmbMonth2.Text, 1), cmbYear.Text)
        
'        StrQuery = "SELECT CONVERT(Char(3), TransDate) + ' ' + CONVERT(Char(4), YEAR(TransDate)), Commission, SuretyFund, SSSPayment, PHealth, Pagibig, CashAdvance, SSSLoan, PagibigLoan, Particulars, Others " & _
'            " AppliedMonth, AppliedYear FROM DriverConductorPayments WHERE EmployeeID = '" & txtClient(0) & "' ORDER by TransDate"
'        Set rsTrips = returnrst(StrQuery, cnnstr, mstatic)
'        If rsTrips.RecordCount > 0 Then
'            Call FillListFund(lstPayments(1), rsTrips, 7)
'        Else
'            lstPayments(1).ListItems.Clear
'            MsgBox "No records selected!", vbInformation, PayrollAlert
'        End If
    
    End If
End If

'If action_query(StrQuery, cnnstr) = False Then
'    MsgBox "Failed to get records!", vbInformation, PayrollAlert
'Else
'    StrQuery = "SELECT TransDate, BundleNo, Commission, Funds, BusNo, Route, EmployeeID FROM tmpFMSCommissions ORDER by BUNDLENo"
'    Set rsTrips = returnrst(StrQuery, cnnstr, mstatic)
'    If rsTrips.RecordCount > 0 Then
'        Call FillListFund(lstPayments(0), rsTrips, 5)
'    Else
'        Call DisplayTotalFunds(0, 0, 0)
'        lstPayments(0).ListItems.Clear
'        MsgBox "No records selected!", vbInformation, PayrollAlert
'    End If
'End If

End Sub

Private Sub RefreshPayments(ByVal mMon As Integer, ByVal mYr As Integer)
Dim mQtr As String
Dim rsTrips As ADODB.Recordset

'Select Case Val(Left(cmbMonth2.Text, 2))

Select Case mMon
    Case 1:
        mQtr = "(1, 2, 3)"
    Case 2:
        mQtr = "(4, 5, 6)"
    Case 3:
        mQtr = "(7, 8, 9)"
    Case Else
        mQtr = "(10, 11, 12)"
End Select

StrQuery = "SELECT PaymentDate, SUBSTRING(TransMonth, 1, 3) + ' ' + CONVERT(char(4), Transyear) AS TransDate, BegBalance, Commission, SuretyFund, EESSS, EEPHealth, EEHDMF, EETax, CashAdvance, SSSLoan, HDMFLoan, Particulars, Others, " & _
    " AppliedMonth, AppliedYear FROM DriverConductorPayments WHERE EmployeeID = '" & txtClient(0) & "' AND AppliedMonth IN " & mQtr & " AND AppliedYear = " & mYr & " ORDER by TransDate"
Set rsTrips = ReturnRst(StrQuery, cnnstr, mstatic)
If rsTrips.RecordCount > 0 Then
    Call FillListFund(lstPayments(1), rsTrips, 10, False)
Else
    lstPayments(1).ListItems.Clear
    MsgBox "No records selected!", vbInformation, PayrollAlert
End If

End Sub

Private Sub cmdDisplay_Click()
Call RefreshPayments(Left(cmbMonth2.Text, 1), cmbYear2.Text)
End Sub

Private Sub cmdEditBalance_Click()
Dim rsBal As ADODB.Recordset
Dim res

If cmdEditBalance.Caption = "Edit" Then
    txtFund(4).Enabled = True
    txtFund(4).SetFocus
    cmdEditBalance.Caption = "Save"
Else    ' Save changes to HR database.
    res = MsgBox("Save changes to beginning balance of " & txtClient(2).Text & "?", vbQuestion + vbYesNo, PayrollAlert)
    If res = vbYes Then ' mark for payroll processing
        StrQuery = "UPDATE " & Trim$(txtDBName) & "..EMP201 SET Accumulate_Fund = " & Replace(txtFund(4).Text, ",", "") & " WHERE name = '" & Trim$(txtClient(2).Text) & "' and ID = '" & Trim$(txtClient(0).Text) & "'"
        If Action_Query(StrQuery, cnnstr) = True Then
            MsgBox "Beginning balance successfully updated!", vbInformation, PayrollAlert
            rsEmployeeBasicData.Requery
        Else
            MsgBox "Failed to update beginning balance!", vbInformation, PayrollAlert
        End If
    End If
    txtFund(4).Enabled = False
    cmdEditBalance.Caption = "Edit"
End If

End Sub

Private Sub cmdExport_Click()
Call GenerateDriverReportInExcel(Left(cmbMonth2.Text, 1), cmbYear2.Text)
End Sub

Private Sub GenerateDriverReportInExcel(ByVal mMon As Integer, ByVal mYear As Integer)
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
Dim rsContri As Recordset
Dim mQtr As String

On Error GoTo ErrorRtn2

Select Case mMon
    Case 1:
        mQtr = "(1, 2, 3)"
    Case 2:
        mQtr = "(4, 5, 6)"
    Case 3:
        mQtr = "(7, 8, 9)"
    Case 4:
        mQtr = "(10, 11, 12)"
End Select

If chkAll.VALUE = True Then ' all
    StrQuery = "SELECT * FROM DriverconductorPayments " & _
        " WHERE AppliedMonth IN " & mQtr & " AND AppliedYear = " & mYear & _
        " AND Positions = '" & Trim$(txtDriver) & "' " & _
        " Order By Fullname, AppliedMonth, AppliedYear "
Else
    StrQuery = "SELECT * FROM DriverconductorPayments " & _
        " WHERE EmployeeID = '" & Trim$(txtClient(0)) & "'" & _
        " AND AppliedMonth IN " & mQtr & " AND AppliedYear = " & mYear & _
        " AND Positions = '" & Trim$(txtDriver) & "' " & _
        " Order By Fullname, AppliedMonth, AppliedYear "
End If

Set rsContri = ReturnRst(StrQuery, cnnstr, mstatic)
If rsContri.RecordCount = 0 Then
    MsgBox "No records selected!", vbInformation, PayrollAlert
    Exit Sub
End If

With rsContri

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)

ctr = 9
i = 1

'Updated by Bernard May 13, 2008.
'mHead = CompanyName & " (" & CompanyShort & ")"
mHead = cmbCompany.Text
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 12, True)
mHead = Trim$(txtDriver) & " SSS, PHealth and Other Compensation Report"
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, False)
mHead = "For " & cmbMonth2.Text & " " & cmbYear2.Text & Space(20) & Format(Now, "mm/dd/yyyy hh:nn:ss")
Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, False)

Call PrintReportHeader(xlSheet, xlApp1, "A7:A8", "Number", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "B7:B8", "Employee Name", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 21.4)
Call PrintReportHeader(xlSheet, xlApp1, "C7:C8", "Trans Date", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "D7:D8", "Beginning", "Balance", xlCenter, xlBottom, True, False, True, 15, , 9, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "E7:E8", "Surety", "Fund", xlCenter, xlBottom, True, True, True, 15, , 9, False, 8.7)

Call PrintReportHeader(xlSheet, xlApp1, "F7:F8", "EE SSS", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "G7:G8", "ER SSS", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "H7:H8", "ECC SSS", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "I7:I8", "EE PHealth", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "J7:J8", "ER PHealth", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "K7:K8", "EE HDMF", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "L7:L8", "ER HDMF", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "M7:M8", "TAX", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 8.7)

.MoveFirst
Do While Not .EOF
    xlSheet.Cells(ctr, 1).VALUE = i
    xlSheet.Cells(ctr, 2).VALUE = !FullName
    xlSheet.Cells(ctr, 3).VALUE = Left(!TransMonth, 3) & " " & !AppliedYear
    xlSheet.Cells(ctr, 4).VALUE = Format(!BegBalance, "##0.00")
    xlSheet.Cells(ctr, 5).VALUE = Format(!SuretyFund, "##0.00")
    xlSheet.Cells(ctr, 6).VALUE = Format(!EESSS, "##0.00")
    xlSheet.Cells(ctr, 7).VALUE = Format(!ERSSS, "##0.00")
    xlSheet.Cells(ctr, 8).VALUE = Format(!ECC, "##0.00")
    xlSheet.Cells(ctr, 9).VALUE = Format(!EEPHealth, "##0.00")
    xlSheet.Cells(ctr, 10).VALUE = Format(!ERPHealth, "##0.00")
    xlSheet.Cells(ctr, 11).VALUE = Format(!EEHDMF, "##0.00")
    xlSheet.Cells(ctr, 12).VALUE = Format(!ERHDMF, "##0.00")
    xlSheet.Cells(ctr, 13).VALUE = Format(!EETax, "##0.00")
    
    'xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
    
    'xlSheet.Cells(ctr, 4).FormulaR1C1 = "=RC[-2]-RC[-1]" '"=RC[-2]+RC[-1]"
    ctr = ctr + 1
    i = i + 1
    .MoveNext
Loop

End With

'ctr = ctr + 3
'xlSheet.Cells(ctr, 2).VALUE = "TOTAL HEADCOUNT"
'
'mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
'ctr = ctr + 3
'xlSheet.Cells(ctr, 2).VALUE = "TOTAL HEADCOUNT"
'Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)
'
'
'mRange = "C" & 7 & ":C" & ctr
'xlSheet.Range(mRange).Select
'mRange2 = "C" & ctr
'xlSheet.Range(mRange2).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-3]C)"
'
'mRange = "D" & 7 & ":D" & ctr
'xlSheet.Range(mRange).Select
'mRange2 = "D" & ctr
'xlSheet.Range(mRange2).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-3]C)"
'
'mRange = "E" & 7 & ":E" & ctr
'xlSheet.Range(mRange).Select
'mRange2 = "E" & ctr
'xlSheet.Range(mRange2).Activate
'xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-3]C)"

'Box it.
mRange = "A7:M" & ctr + 3
Call PutBox(xlSheet, xlApp1, mRange)

'ctr = ctr + 7
'xlSheet.Cells(ctr, 2).VALUE = "Prepared by: ______________"
'xlSheet.Cells(ctr + 3, 2).VALUE = "Note by: ______________"
xlSheet.Range("A1").Select

Exit Sub


ErrorRtn2:

    MsgBox Err.Description, vbInformation, PayrollAlert
    Err.Clear
    
End Sub

Private Sub PutBox(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal mRange As String)

xlSheet.Range(mRange).Select
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
With xlSheet.Range(mRange)
    .Font.Name = "Arial"   'Font.Name
    .Font.Size = 8         'Font.size
End With

End Sub

Private Sub cmdGenerate_Click()
Dim rsTrips As ADODB.Recordset

' Exec GetFMSCommissions 'DRIVER', 2009, 'November', 'TGC', 'TAS TRANS', '26389'

StrQuery = "EXEC dbo.GetFMSCommissions '" & txtDriver.Text & "', " & cmbYear & ", '" & cmbMonth.Text & _
        "', '" & txtDBName & "', '" & txtDepartment.Text & "', '" & txtClient(0).Text & "'"

If Action_Query(StrQuery, cnnstr) = False Then
    MsgBox "Failed to get records!", vbInformation, PayrollAlert
Else
    StrQuery = "SELECT TransDate, BundleNo, Commission, Funds, BusNo, Route, EmployeeID FROM tmpFMSCommissions ORDER by BUNDLENo"
    Set rsTrips = ReturnRst(StrQuery, cnnstr, mstatic)
    If rsTrips.RecordCount > 0 Then
        Call FillListFund(lstPayments(0), rsTrips, 5)
    Else
        Call DisplayTotalFunds(0, 0, 0)
        lstPayments(0).ListItems.Clear
        MsgBox "No records selected!", vbInformation, PayrollAlert
    End If
End If

'    StrQuery = "SELECT LEFT(DATENAME(month, b.TrxDate), 3) + ' ' + CONVERT(char(4), YEAR(b.TrxDate)) AS TransDate, b.BusNo AS BusNo, " & _
'        "b.Route AS Route, b.BundleNo AS BundleNo, b.ID_CON AS EmployeeID, " & _
'        "b.C_Comission AS Commission, b.c_Fund AS Funds, " & _
'        "FROM " & txtDBName.Text & "..Trips b " & _
'        "Where Year(b.TRXDATE) = " & cmbYear.Text & " AND DATENAME(month, b.TRXDate) = '" & cmbMonth.Text & "' " & _
'        " AND b.ID_CON = '" & Trim$(txtClient(0).Text) & "' AND Conductor = '" & Trim$(txtClient(2).Text) & "' " & _
'        " AND b.Dapart = '" & txtDepartment.Text & "'"

End Sub

Public Sub FillListFund(ByRef mLsviv As ListView, _
                        mRs As Recordset, NoOfCol As Integer, Optional DateFormat As Boolean = True)
Dim litem As ListItem
Dim i As Integer, y As Integer, X As Integer
Dim totCom As Double, totFund As Double, totBundle As Integer

totCom = 0#: totFund = 0#: totBundle = 0
y = 0: X = 1
mLsviv.View = lvwReport
mLsviv.ListItems.Clear

While Not (mRs.EOF Or mRs.BOF)
    Set litem = mLsviv.ListItems.Add(, , IIf(IsNull(mRs(0)), "    ", mRs(0)))
        For i = 1 To NoOfCol
            If IsDate(mRs(i)) = True Then
                If DateFormat = True Then
                    litem.SubItems(i) = Format(mRs(i), "dd/mm/yyyy")
                Else
                    litem.SubItems(i) = Trim$(IIf(IsNull(mRs(i)), "    ", mRs(i)))
                End If
            ElseIf IsNumeric(mRs(i)) = True Then
                If mRs(i) = 0 Then
                    litem.SubItems(i) = "0.00"
                Else
                    litem.SubItems(i) = Val(mRs(i))
                End If
            Else
                litem.SubItems(i) = Trim$(IIf(IsNull(mRs(i)), "    ", mRs(i)))
            End If
            
            If mRs.Fields(i).Name = "Commission" Then
                totCom = totCom + Val(mRs(i))
            ElseIf mRs.Fields(i).Name = "Funds" Then
                totFund = totFund + Val(mRs(i))
            End If
        Next
        If y = 0 Then
            y = 1
        Else
            y = 0
        End If
    X = X + 1
    totBundle = totBundle + 1
    mRs.MoveNext
Wend
Call DisplayTotalFunds(totCom, totFund, totBundle)
End Sub

Private Sub DisplayTotalFunds(mCom As Double, mFund As Double, mBundle As Integer)

'txtFund(0).Text = mBundle
txtFund(1).Text = mCom
txtFund(2).Text = mFund

End Sub

Private Sub Form_Activate()
'Call DisplayGrossTable
End Sub

Private Sub Form_Load()
Dim mComp As String
Dim mEmpN As String
Dim y As Integer

    If ActiveSw = False Then
        Me.Top = (mdiMain.Height - Me.Height) / 8
        Me.Left = (mdiMain.Width - Me.Width) / 2
    End If
    FormLoadSw = False
    txtCancelSw = "0"
    txtFund(0).Text = "200.00"
    
    ' Get Account Name.
    cmbCompany.Enabled = True
    mUserID = UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID)))
    'mUserID = LoginName
    ActiveSw = True
    RefreshSw = False
    Call GetEmployeeBasic
    If EditFlag = True Then
        rsEmployeeBasicData.Find " ClientCode = " & mRecNo
    End If
    AddFlag = False
    EditFlag = False
    AddDeductions = False
    EditDeductions = False
    txtRates(16).Enabled = False
    
    chkActiveClient.Enabled = False
    If ActiveCtr = 0 Then
        Call DisplayBasicData("", True)
    Else
        Call DisplayBasicData(CurrEmployee, True)
    End If
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
    y = 1
    cmbYear.Clear
    Do While y <= 10
        DoEvents
        cmbYear.AddItem 2003 + y
        cmbYear.ItemData(cmbYear.NewIndex) = y
        y = y + 1
    Loop
    
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
    cmbYear.Text = Format(Now, "yyyy")
    'cmbMonth.ListIndex = 0
    cmbMonth.Text = Trim$(Format(Date, "mm")) & "-" & (Format(Date, "mmmm"))

    y = 1
    cmbYear2.Clear
    Do While y <= 10
        DoEvents
        cmbYear2.AddItem 2003 + y
        cmbYear2.ItemData(cmbYear2.NewIndex) = y
        y = y + 1
    Loop
    
    'Populate Month.
    cmbMonth2.AddItem "1st Quarter"
    cmbMonth2.AddItem "2nd Quarter"
    cmbMonth2.AddItem "3rd Quarter"
    cmbMonth2.AddItem "4th Quarter"
    cmbYear2.Text = Format(Now, "yyyy")
    cmbMonth2.Text = "1st Quarter"
    cmbCompany.ListIndex = 0
    
End Sub

Private Sub GetEmployeeBasic()
Dim rsTaxCode As ADODB.Recordset

'Get FMS Clients only.
StrQuery = "SELECT * FROM clienttable WHERE ClientCode2 = 'FMS' AND clientstatus = 1 and clientname is not null ORDER BY clientname"
Set rsClientData = ReturnRst(StrQuery, cnnstr, mstatic)

If rsClientData.RecordCount <> 0 Then
    Call FillComboData(cmbCompany, rsClientData)
    rsClientData.MoveFirst
    mCCode = rsClientData!ClientCode
    txtClientCode = rsClientData!ClientCode
    'Get the Department and Database name.
    txtDepartment = rsClientData!Department
    txtDBName = rsClientData!DatabaseName
    txtDriver = rsClientData!DriverConductor
Else
    Exit Sub
End If
    
StrQuery = "SELECT [ID] AS EmployeeID, [Name] AS FullName, Positions, Company, TaxCode, Accumulate_Fund FROM " & Trim$(txtDBName) & "..EMP201 WHERE MODE = 'Active' ORDER BY [Name]"
Set rsEmployeeBasicData = ReturnRst(StrQuery, cnnstr, mstatic)
If rsEmployeeBasicData.RecordCount <> 0 Then
    Call FilterEmployee(rsClientData!ClientName)
End If

StrQuery = "SELECT id, Taxexemptioncode FROM TaxExemptioncode ORDER BY id"
Set rsTaxCode = ReturnRst(StrQuery, cnnstr, mstatic)

If rsTaxCode.RecordCount <> 0 Then
    Call FillComboData(cmbTaxCode, rsTaxCode)
End If
Set rsTaxCode = Nothing

End Sub

Private Sub FilterEmployee(ByVal mCompany As String)
    rsClientData.MoveFirst
    rsClientData.Find "clientname = '" & Trim(mCompany) & "'"
    If rsClientData.EOF = True Then
        mCCode = 0
        Exit Sub
    Else
        mCCode = rsClientData!ClientCode
        txtClientCode = rsClientData!ClientCode
        'Get the Department and Database name.
        txtDepartment = rsClientData!Department
        txtDBName = rsClientData!DatabaseName
        txtDriver = rsClientData!DriverConductor
    End If
    With rsEmployeeBasicData
        '.Filter = "clientcode = " & mCCode
        .Filter = "Company = '" & Trim$(txtDepartment) & "' AND Positions = '" & Trim$(txtDriver) & "'"
        
        If .RecordCount <> 0 Then
            'If Not .EOF And Not .BOF Then
            Call FillEmployeeList(ListView1(0), rsEmployeeBasicData)
            'rsEmployeeBasicData.MoveFirst
            lblclient(10).Caption = "Total:  " & ActiveCtr
             '.Filter = "clientcode = " & mCCode
            .Filter = "Company = '" & Trim$(txtDepartment) & "' AND Positions = '" & Trim$(txtDriver) & "'"
        Else
            Call ClearAllData
            Call clearList
        End If
    End With
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

Private Sub FillEmployeeList(ByRef mLsEmp As ListView, _
                        ByRef mRs As Recordset)
Dim litem1 As ListItem
Dim litem2 As ListItem

    InActiveCtr = 0
    ActiveCtr = 0
    mLsEmp.View = lvwReport
    mLsEmp.ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
        Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName))
        litem1.SubItems(1) = mRs(0)
        ActiveCtr = ActiveCtr + 1
        mRs.MoveNext
    Wend

End Sub

'Private Sub FillEmployeeListOrig(ByRef mLsEmp As ListView, _
'                        ByRef mRs As Recordset)
'Dim litem1 As ListItem
'Dim litem2 As ListItem
'
'    InActiveCtr = 0
'    ActiveCtr = 0
'    mLsEmp.View = lvwReport
'    ListView1(1).View = lvwReport
'    mLsEmp.ListItems.Clear
'    ListView1(1).ListItems.Clear
'    While Not (mRs.EOF Or mRs.BOF)
'        If mRs!EmployeeStatus = True Then ' if active employee
'            Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName))
'            litem1.SubItems(1) = mRs(0)
'            ActiveCtr = ActiveCtr + 1
'        End If
'        mRs.MoveNext
'    Wend
'    'Get all Inactive employees regardless of employers.
'    mRs.MoveFirst
'    'mRs.Filter = "Employeestatus = false"
'    mRs.Filter = "Employeestatus = 0"
'    While Not (mRs.EOF Or mRs.BOF)
'        Set litem2 = ListView1(1).ListItems.Add(, , 1)
'        'Set litem2 = ListView1(1).ListItems.Add(, , IIf(IsNull(mRs!FullName, "Blank", mRs!FullName)))
'        litem2.SubItems(1) = IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName)
'        litem2.SubItems(2) = IIf(IsNull(mRs!ClientName), "Blank", mRs!ClientName)
'        'litem2.SubItems(3) = IIf(IsNull(mRs!EmployeeID), "Blank", mRs!EmployeeID)
'        litem2.SubItems(3) = IIf(IsNull(mRs!EmpNo2), "Blank", mRs!EmpNo2)
'        litem2.SubItems(4) = IIf(IsNull(mRs!StartContract) Or mRs!StartContract = #1/1/1900#, "__/__/____", Format(mRs!StartContract, "mm/dd/yyyy"))
'        litem2.SubItems(5) = IIf(IsNull(mRs!endcontract) Or mRs!endcontract = #1/1/1900#, "__/__/____", Format(mRs!endcontract, "mm/dd/yyyy"))
'        litem2.SubItems(6) = IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName)
'        litem2.SubItems(7) = IIf(IsNull(mRs!EmployeeID), 1, mRs!EmployeeID)
'        InActiveCtr = InActiveCtr + 1
'        mRs.MoveNext
'    Wend
'End Sub

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

'Private Sub lstDTRDisplay_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
'lblclient(62).Caption = Item
'txtDTR(0).Text = Item.ListSubItems.Item(1).Text ' min range
'txtDTR(1).Text = Item.ListSubItems.Item(2).Text ' max range
'txtDTR(2).Text = Item.ListSubItems.Item(3).Text ' amount
'txtDTR(3).Text = Item.ListSubItems.Item(4).Text ' percentage
'
'End Sub

Private Sub txtClient_GotFocus(Index As Integer)
    SendKeys "{HOME}+{END}"
    txtClient(Index).BackColor = &HFFFFC0
End Sub

Private Sub txtClient_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
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
Dim TDays As Double
Dim TAbsent As Double
Dim TLates As Double
Dim TUndertime As Double
Dim TOvertime As Double
Dim TNightDiff As Double

If rsEmployeeBasicData.RecordCount = 0 Then Exit Sub
    
WithDTR = 0
Ttime = 0: TDays = 0: TAbsent = 0: TLates = 0: TUndertime = 0: TOvertime = 0
TNightDiff = 0

With rsEmployeeBasicData
    If recSw = False Then
        .MoveFirst
        .Find "Fullname = '" & SelectedEmployee & "'"
    End If
    If Not .EOF And Not .BOF Then
        'chkActiveClient = IIf(!EmployeeStatus = True, 1, 0)
        txtClient(0) = IIf(IsNull(!EmployeeID), "", !EmployeeID)
        'txtClient(1) = IIf(IsNull(!EmploymentStatusDesc), "", !EmploymentStatusDesc)
        txtClient(2) = IIf(IsNull(!FullName), "", !FullName)
        'txtClient(3) = IIf(IsNull(!NotedBy), "", !NotedBy)
        'txtClient(4) = IIf(IsNull(!DateUpdated), "", Format(!DateUpdated, "mm/dd/yyyy"))
        'txtClient(5) = IIf(IsNull(!IfnoCompensationPackage), "", !IfnoCompensationPackage)
        'txtClient(6) = IIf(IsNull(!Remarks), "", !Remarks)
        'txtClient(7) = IIf(IsNull(!Updatedby), "", !Updatedby)
        txtBegBal(4) = IIf(IsNull(!Accumulate_Fund), 0#, !Accumulate_Fund)   ' accumulated fund
        txtClient(9) = IIf(IsNull(!TaxCode), "", !TaxCode) ' get Tax Code
      
    End If
End With

End Sub

'Private Sub DisplayGrossTable()
'Dim litem2 As ListItem
'Dim Dtrctr As Integer
'Dim rsGross As ADODB.Recordset
'
'StrQuery = "SELECT * FROM tblFMSGrossRange ORDER BY GrossMin"
'Set rsGross = returnrst(StrQuery, cnnstr, mstatic)
'
'Dtrctr = 0
'lstDTRDisplay.ListItems.Clear
'With rsGross
'If .RecordCount <> 0 Then
'    .MoveFirst
'    While Not (.EOF)
'        Dtrctr = Dtrctr + 1
'        Set litem2 = lstDTRDisplay.ListItems.Add(, , !BonusID)
'            litem2.SubItems(1) = IIf(IsNull(!GrossMin), 0, !GrossMin)
'            litem2.SubItems(2) = IIf(IsNull(!GrossMax), 0, !GrossMax)
'            litem2.SubItems(3) = IIf(IsNull(!BonusAmount), 0, !BonusAmount)
'            litem2.SubItems(4) = IIf(IsNull(!BonusPercentage), 0, !BonusPercentage)
'        .MoveNext
'    Wend
'End If
'End With
'Set rsGross = Nothing
'
'End Sub

Private Sub txtDTR_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strValid
    strValid = "0123456789." & Chr(13) & Chr(8)
    If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFund_GotFocus(Index As Integer)
    'SendKeys "{HOME}+{END}"
    'txtFund(Index).BackColor = &HFFFFC0
End Sub

Private Sub txtFund_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFund_LostFocus(Index As Integer)
    'txtFund(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtRates_GotFocus(Index As Integer)
    SendKeys "{HOME}+{END}"
    txtRates(Index).BackColor = &HFFFFC0
End Sub

