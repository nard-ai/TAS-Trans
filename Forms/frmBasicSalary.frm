VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBasicSalary 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Payroll Master"
   ClientHeight    =   6795
   ClientLeft      =   2595
   ClientTop       =   2415
   ClientWidth     =   9060
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmBasicSalary.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9060
   Begin TabDlg.SSTab SSTab1 
      Height          =   6165
      Index           =   1
      Left            =   90
      TabIndex        =   34
      Top             =   540
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   10874
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Employee Payroll Master"
      TabPicture(0)   =   "frmBasicSalary.frx":058A
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
      TabCaption(1)   =   "In-Active Employees"
      TabPicture(1)   =   "frmBasicSalary.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   3
         Left            =   8325
         Picture         =   "frmBasicSalary.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   5790
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   2
         Left            =   8040
         Picture         =   "frmBasicSalary.frx":0AF4
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "First"
         Top             =   5790
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   1
         Left            =   7755
         Picture         =   "frmBasicSalary.frx":0BEE
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "First"
         Top             =   5790
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   0
         Left            =   7470
         Picture         =   "frmBasicSalary.frx":0CE8
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "First"
         Top             =   5790
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
         Height          =   5580
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   8655
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
            Top             =   4800
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
            TabIndex        =   112
            Top             =   4560
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
            TabIndex        =   106
            Text            =   "Client Name"
            Top             =   5170
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
            TabIndex        =   105
            Top             =   4560
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
            TabIndex        =   104
            Top             =   4560
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtEmpNum 
            Height          =   285
            Left            =   990
            TabIndex        =   79
            Text            =   "txtclientcode"
            Top             =   180
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtClientCode 
            Height          =   285
            Left            =   1800
            TabIndex        =   69
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
            Top             =   4320
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
            Left            =   1320
            TabIndex        =   29
            Text            =   "Client Name"
            Top             =   5110
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Frame Frame3 
            Height          =   1365
            Left            =   3240
            TabIndex        =   49
            Top             =   0
            Width           =   5415
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
               ItemData        =   "frmBasicSalary.frx":121A
               Left            =   3120
               List            =   "frmBasicSalary.frx":122A
               Style           =   2  'Dropdown List
               TabIndex        =   138
               Top             =   330
               Width           =   1365
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
               ItemData        =   "frmBasicSalary.frx":1252
               Left            =   4320
               List            =   "frmBasicSalary.frx":125C
               TabIndex        =   7
               Text            =   "cmbTaxCode"
               Top             =   945
               Width           =   1035
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
               Left            =   4500
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
               Text            =   "frmBasicSalary.frx":1270
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
               TabIndex        =   139
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
               Left            =   4320
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
            ItemData        =   "frmBasicSalary.frx":127C
            Left            =   180
            List            =   "frmBasicSalary.frx":1289
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   450
            Width           =   2925
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   3150
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   1080
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   5556
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
            TabIndex        =   84
            Top             =   1365
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
               TabIndex        =   110
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
               TabIndex        =   89
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
               TabIndex        =   109
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
               TabIndex        =   108
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
               TabIndex        =   98
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
               TabIndex        =   97
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
               TabIndex        =   90
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
               TabIndex        =   101
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
               TabIndex        =   100
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
               TabIndex        =   99
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
               TabIndex        =   96
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
               TabIndex        =   95
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
               TabIndex        =   94
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
               TabIndex        =   93
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
               TabIndex        =   92
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
               TabIndex        =   91
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
               TabIndex        =   103
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
               TabIndex        =   102
               Top             =   1395
               Visible         =   0   'False
               Width           =   855
            End
         End
         Begin VB.Frame Frame5 
            Height          =   1695
            Index           =   0
            Left            =   3240
            TabIndex        =   61
            Top             =   2805
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
               Left            =   -600
               TabIndex        =   21
               Text            =   "0.00"
               Top             =   1905
               Visible         =   0   'False
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
               Caption         =   "Allowance per Cut-off:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   375
               Index           =   35
               Left            =   210
               TabIndex        =   88
               Top             =   1215
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
               TabIndex        =   87
               Top             =   1125
               Visible         =   0   'False
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
               TabIndex        =   86
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
               TabIndex        =   85
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
               TabIndex        =   78
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
               TabIndex        =   77
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
               TabIndex        =   76
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
               TabIndex        =   75
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
               TabIndex        =   74
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
               TabIndex        =   73
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
               TabIndex        =   72
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
               TabIndex        =   71
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
            Top             =   4800
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
            Top             =   4575
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2985
         End
         Begin VB.Frame Frame1 
            Height          =   1455
            Left            =   3240
            TabIndex        =   42
            Top             =   1305
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
               TabIndex        =   140
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
               ItemData        =   "frmBasicSalary.frx":12AC
               Left            =   120
               List            =   "frmBasicSalary.frx":12BF
               Style           =   2  'Dropdown List
               TabIndex        =   136
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
               TabIndex        =   137
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
            Left            =   4200
            TabIndex        =   113
            Top             =   2565
            Width           =   5415
            Begin VB.CommandButton cmdDelDed 
               Caption         =   "Del"
               Height          =   255
               Left            =   4680
               TabIndex        =   134
               Top             =   1320
               Width           =   615
            End
            Begin VB.CommandButton cmdEditDed 
               Caption         =   "Edit"
               Height          =   255
               Left            =   4080
               TabIndex        =   133
               Top             =   1320
               Width           =   615
            End
            Begin VB.CommandButton cmdAddDed 
               Caption         =   "Add"
               Height          =   255
               Left            =   3360
               TabIndex        =   132
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
               TabIndex        =   131
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
               TabIndex        =   130
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
               TabIndex        =   122
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
               ItemData        =   "frmBasicSalary.frx":12E0
               Left            =   1470
               List            =   "frmBasicSalary.frx":12ED
               Style           =   2  'Dropdown List
               TabIndex        =   121
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
               TabIndex        =   125
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
               TabIndex        =   123
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
               TabIndex        =   129
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
               TabIndex        =   128
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
               TabIndex        =   126
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
               TabIndex        =   127
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
               TabIndex        =   135
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
               TabIndex        =   124
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
               TabIndex        =   119
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
               TabIndex        =   118
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
               TabIndex        =   117
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
               TabIndex        =   116
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
               TabIndex        =   115
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
               TabIndex        =   114
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
               TabIndex        =   120
               Top             =   240
               Width           =   1545
            End
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
            TabIndex        =   107
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
            Left            =   2250
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
            TabIndex        =   55
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
            Left            =   480
            TabIndex        =   54
            Top             =   5110
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
            Top             =   5280
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
            Top             =   4800
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Height          =   5640
         Left            =   -74910
         TabIndex        =   56
         Top             =   360
         Width           =   8700
         Begin MSComctlLib.ListView ListView1 
            Height          =   4695
            Index           =   1
            Left            =   180
            TabIndex        =   59
            Top             =   810
            Width           =   8310
            _ExtentX        =   14658
            _ExtentY        =   8281
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
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Employee Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Company Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Emp Number"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Start Contract"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "End Contract"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "active"
               Object.Width           =   2
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Width           =   2
            EndProperty
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
            Left            =   7290
            TabIndex        =   70
            Top             =   180
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
            Left            =   2880
            TabIndex        =   68
            Text            =   "EmpName"
            Top             =   180
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
            Left            =   900
            TabIndex        =   67
            Text            =   "EmpNo"
            Top             =   180
            Width           =   1305
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000000&
            BorderWidth     =   4
            Index           =   0
            X1              =   195
            X2              =   8385
            Y1              =   630
            Y2              =   630
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
            Left            =   2340
            TabIndex        =   58
            Top             =   270
            Width           =   915
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Emp #:"
            BeginProperty Font 
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
            Left            =   270
            TabIndex        =   57
            Top             =   270
            Width           =   915
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
            Picture         =   "frmBasicSalary.frx":1310
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
      Width           =   9060
      _ExtentX        =   15981
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
         Picture         =   "frmBasicSalary.frx":136E
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Deduction Types"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   2745
         Picture         =   "frmBasicSalary.frx":17B0
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   2175
         Picture         =   "frmBasicSalary.frx":18B2
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmBasicSalary.frx":19FC
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "New"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1080
         Picture         =   "frmBasicSalary.frx":1F2E
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1620
         Picture         =   "frmBasicSalary.frx":2460
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   540
         Picture         =   "frmBasicSalary.frx":2992
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdClose 
         Height          =   405
         Left            =   3285
         Picture         =   "frmBasicSalary.frx":2ADC
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
Attribute VB_Name = "frmBasicSalary"
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

Private Sub cmbTaxCode_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub cmbTaxCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbCompany_GotFocus()
'    SendKeys "{HOME}+{END}"
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
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbPaySystem_GotFocus()
    'SendKeys "{HOME}+{END}"
End Sub

Private Sub cmbPaySystem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbCompany_Click()
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
        chkRates(11).Enabled = mTrue
        chkRates(12).Enabled = mTrue
        txtRates(17).Enabled = mTrue
        txtRates(18).Enabled = mTrue
        txtRates(15).Enabled = mTrue
        txtRates(16).Enabled = False
        txtRates(21).Enabled = mTrue
        MaskEdBox5.Enabled = mTrue
        MaskEdBox6.Enabled = mTrue
    Else
        chkRates(11).VALUE = 1
        chkRates(12).VALUE = 0
        txtRates(17).Text = mTrue
        txtRates(18).Text = mTrue
        txtRates(15).Text = mTrue
        'txtRates(16).Text = mTrue
        txtRates(21).Text = mTrue
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
            chkRates(11).VALUE = IIf(rsEmpDeductions.Fields("DeductionStatus") = True, 1, 0)    ' 6
            txtRates(17).Text = Format(rsEmpDeductions.Fields("DeductionAmount"), "00.00")   ' 9
            txtRates(18).Text = Format(rsEmpDeductions.Fields("DeductionPayment"), "00.00")  ' 8
            MaskEdBox5.Text = Format(rsEmpDeductions.Fields("DedStartPayment"), "mm/dd/yyyy")  ' start deduction date ' 10
            MaskEdBox6.Text = Format(rsEmpDeductions.Fields("DedEndPayment"), "mm/dd/yyyy")  ' end deduction date   ' 11
            txtRates(15).Text = Format(rsEmpDeductions.Fields("TotalInterest"), "0.00")  ' interest 14
            txtRates(16).Text = Format(rsEmpDeductions.Fields("TotalDeduction"), "0.00")  ' total deductions 15
            txtRates(21).Text = Format(rsEmpDeductions.Fields("NoOfdeductions"), "0.00")  ' no of deductions 12
            chkRates(12).VALUE = IIf(rsEmpDeductions.Fields("DeductionFlag") = True, 1, 0)  ' 17
            mDeductionCode = rsEmpDeductions.Fields("DeductionCode")    ' 7
            NoofPayments = rsEmpDeductions.Fields("DeductionCounter")       ' payments counter  13
            TotalPayments = rsEmpDeductions.Fields("PaymentsMade")      ' payments made  16
            NoEditFlag = IIf(rsEmpDeductions.Fields("NoEditFlag") = True, 1, 0)   ' flag whether the record is editable or not 20
            lblclient(36).Caption = "Payments: " & TotalPayments & " / " & NoofPayments
            
            mInterest = txtRates(15).Text
            mTotalLoan = txtRates(16).Text
            
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
    If cmdAddDed.Caption = "Add" Then
        StrQuery = "SELECT * FROM DeductionTypes2 WHERE DeductionCode <> 0 ORDER BY DeductionDescription"
        Set rsGetDeductionTypes = ReturnRst(StrQuery, CnnstrHR, mstatic)
        
        ' Get all the types of Deductions not yet included for the current Employee.
        EmpDeductionList = ""
        If rsEmpDeductions.RecordCount <> 0 Then
            rsEmpDeductions.MoveFirst
            While Not rsEmpDeductions.EOF
                EmpDeductionList = EmpDeductionList & LTrim(RTrim(str(rsEmpDeductions.Fields(7)))) & " "
                rsEmpDeductions.MoveNext
            Wend
        End If
        
        ' Fill the combo with New Deductions.
        Call FillComboData(cmbTypes, rsGetDeductionTypes, True, EmpDeductionList)
        If cmbTypes.ListCount = 0 Then
            MsgBox "No available deductions to add!", vbInformation, PayrollAlert
            GoTo Routine1
        End If
        Call EnableDisableDeductions(True)
        Call EnableDisableDeductions("", True)
        Call EnableDisableDeductionButtons("Save", "Cancel", False)
        ' mDeductionCode = GetDeductionCode(cmbTypes.Text)
        AddDeductions = True
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
            If TotalPayments <> 0 And TotalPayments = Val(txtRates(16)) Then ' already paid
                res = MsgBox("Total Loan equal to Total Payments made. Do you want to mark it as fully paid?", vbQuestion + vbYesNo, PayrollAlert)
                If res = vbYes Then ' mark as fully paid
                    ProtectFlag = True
                Else
                    MsgBox "Total Deduction Amount should be greater than Payments Made if not yet mark as fully paid.", vbInformation, PayrollAlert
                    txtRates(17).SetFocus
                    Exit Sub
                End If
            ElseIf TotalPayments <> 0 Then ' compute for total entries if tally if there were payments already made
                If Val(txtRates(16)) <> ((txtRates(18) * (txtRates(21) - NoofPayments)) + TotalPayments) Then
                    MsgBox "No of payments (" & txtRates(21) & " less " & NoofPayments & " no of payments) times deduction (" & txtRates(18) & ") + Payments made (" & TotalPayments & _
                        ") equal to " & ((txtRates(18) * (txtRates(21) - NoofPayments)) + TotalPayments) & ". Not equal to Total Deduction (" & txtRates(16) & "). Please check!", vbInformation, PayrollAlert
                        'txtRates(17).SetFocus
                        'Exit Sub
                End If
            Else
                If Val(txtRates(16)) <> (txtRates(18) * txtRates(21)) Then
                    MsgBox "No of payments (" & txtRates(21) & ") times deduction (" & txtRates(18) & ")" & _
                        " equal to " & (txtRates(18) * txtRates(21)) & ". Not equal to Total Deduction (" & txtRates(16) & "). Please check!", vbInformation, PayrollAlert
                        'txtRates(17).SetFocus
                        'Exit Sub
                End If
            End If
        End If
        If AddDeductions = True Then                ' Insert New Deductions
            StrQuery = "INSERT INTO EmployeeDeductions (EmployeeID, EmpNo, AgencyCode, ClientCode, DeductionStatus, DeductionCode, DeductionPayment, DeductionAmount, DedStartPayment, DedEndPayment, TotalInterest, TotalDeduction, NoOfDeductions, DeductionFlag, EntryBy, EntryDate) " & _
                " VALUES ( " & txtEmpNum & ", '" & txtClient(0) & "', " & CompanyCode & ", " & mCCode & ", " & chkRates(11) & ", " & mDeductionCode & ", " & txtRates(18) & ", " & txtRates(17) & _
                ", '" & IIf(MaskEdBox5 = "__/__/____", #1/1/1900#, Format(MaskEdBox5, "mm/dd/yyyy")) & "', '" & IIf(MaskEdBox6 = "__/__/____", #1/1/1900#, Format(MaskEdBox6, "mm/dd/yyyy")) & _
                "', " & txtRates(15) & ", " & txtRates(16) & ", " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", " & chkRates(12) & ", '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & "')"
            mstr = "Deduction was saved!"
            mDetails = "Deduction was added with total of " & txtRates(16) & " for " & txtClient(0)
            mRemarks = "New deduction."
        ElseIf EditDeductions = True Then           ' Save Edited Deductions
            If ProtectFlag = True Then ' mark it as fully paid
                StrQuery = "UPDATE EmployeeDeductions SET EmployeeID = " & txtEmpNum & ", EmpNo = '" & txtClient(0) & "', AgencyCode = " & CompanyCode & ", ClientCode = " & mCCode & _
                    ", DeductionStatus = " & chkRates(11) & ", DeductionCode = " & mDeductionCode & ", DeductionPayment = " & txtRates(18) & ", DeductionAmount = " & txtRates(17) & _
                    ", DedStartPayment = '" & IIf(MaskEdBox5 = "__/__/____", #1/1/1900#, Format(MaskEdBox5, "mm/dd/yyyy")) & "', DedEndPayment = '" & IIf(MaskEdBox6 = "__/__/____", #1/1/1900#, Format(MaskEdBox6, "mm/dd/yyyy")) & _
                    "', TotalInterest = " & txtRates(15) & ", TotalDeduction = " & txtRates(16) & ", NoOfDeductions = " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", DeductionFlag = " & chkRates(12) & ", EntryBy = '" & mUserID & "', EntryDate = '" & Format(Now, "mm/dd/yyyy") & "'" & _
                    ", Protectedflag = 1" & _
                    " Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode
                    '" Where AgencyCode = " & CompanyCode & " and clientcode = " & mCCode & " and EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode ' Updated by Bernard May 30, 2009
                    mstr = "Deduction was marked as Fully Paid!"
                    mDetails = "Deduction was edited to " & txtRates(16) & " for " & txtClient(0)
                    mRemarks = "Edit deduction."
            Else
                StrQuery = "UPDATE EmployeeDeductions SET EmployeeID = " & txtEmpNum & ", EmpNo = '" & txtClient(0) & "', AgencyCode = " & CompanyCode & ", ClientCode = " & mCCode & _
                    ", DeductionStatus = " & chkRates(11) & ", DeductionCode = " & mDeductionCode & ", DeductionPayment = " & txtRates(18) & ", DeductionAmount = " & txtRates(17) & _
                    ", DedStartPayment = '" & IIf(MaskEdBox5 = "__/__/____", #1/1/1900#, Format(MaskEdBox5, "mm/dd/yyyy")) & "', DedEndPayment = '" & IIf(MaskEdBox6 = "__/__/____", #1/1/1900#, Format(MaskEdBox6, "mm/dd/yyyy")) & _
                    "', TotalInterest = " & txtRates(15) & ", TotalDeduction = " & txtRates(16) & ", NoOfDeductions = " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", DeductionFlag = " & chkRates(12) & ", EntryBy = '" & mUserID & "', EntryDate = '" & Format(Now, "mm/dd/yyyy") & _
                    "' Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode
                    '"' Where AgencyCode = " & CompanyCode & " and clientcode = " & mCCode & " and EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode ' updated by Bernard
                    mstr = "Deduction was edited!"
                    mDetails = "Deduction was edited to " & txtRates(16) & " for " & txtClient(0)
                    mRemarks = "Edit deduction."
            End If
        End If
        
        If Action_Query(StrQuery, CnnstrHR) = True Then
            Call SaveUserTransactions(mDetails, mRemarks, txtEmpNum, txtClient(0))
        End If
        MsgBox mstr, vbInformation, PayrollAlert
        
Routine1:

        'Requery all Active Employee Deductions.
        StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
        'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ProtectedFlag = 0"
        Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
        Call FilterDeductions(txtEmpNum)
        
        Call EnableDisableDeductions(False)
        Call EnableDisableDeductionButtons("Add", "Edit", True)
        EditDeductions = False
        AddDeductions = False
        cmbTypes.Enabled = True
        
    End If
    
    Exit Sub
    
NextProc:

    MsgBox Err.Description
    Err.Clear
    Resume Next
    
End Sub

Private Sub cmdClose_Click()
    Set rsEmployeeBasicData = Nothing
    Unload Me
End Sub

Private Function GetDeductionCode(ByVal mDesc As String) As Integer

    GetDeductionCode = 0
    If mDesc <> "" Or mDesc <> Null Then
        rsGetDeductionTypes.MoveFirst
        rsGetDeductionTypes.Find "DeductionDescription = '" & mDesc & "'"
        If rsGetDeductionTypes.EOF = False Then
            GetDeductionCode = rsGetDeductionTypes.Fields(0)
        End If
    End If
    
End Function

Private Sub cmdDeductionTypes_Click()

    Call frmDeductionTypes.Show
    
    ' Update the deduction types recordset with new entries.
    StrQuery = "SELECT * FROM DeductionTypes2 WHERE DeductionCode <> 0 ORDER BY DeductionDescription"
    Set rsDeductionTypes = ReturnRst(StrQuery, CnnstrHR, mstatic)
    
    ' Update the recordset for Employee Deductions.
    StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
    'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ProtectedFlag = 0"
    Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
    
    ' Re-filter deductions.
    If txtEmpNum = "txtclientcode" Then
        txtEmpNum = 0
    End If
    
    Call FilterDeductions(txtEmpNum)

End Sub

Private Sub cmdDelDed_Click()
Dim res
Dim mRem As String

    mRem = mUserID
    If cmbTypes.Text = "" Then MsgBox "No Active Deductions, cannot be deleted!": Exit Sub
    mDeductionCode = GetDeductionCode(cmbTypes.Text)
    If TotalPayments <> 0 Then
        res = MsgBox("There were already " & NoofPayments & " payments made amounting to " & TotalPayments & ". Do you want to mark it as fully paid instead?", vbQuestion + vbYesNo, PayrollAlert)
        If res = vbYes Then ' delete deductions
            mRem = "Marked as fully paid by " & mUserID & ", last " & Format(Now, "mm/dd/yyyy hh:mm:ss AMPM")
            StrQuery = "UPDATE EmployeeDeductions SET EmployeeID = " & txtEmpNum & ", EmpNo = '" & txtClient(0) & "', AgencyCode = " & CompanyCode & ", ClientCode = " & mCCode & _
                    ", DeductionStatus = " & chkRates(11) & ", DeductionCode = " & mDeductionCode & ", DeductionPayment = " & IIf(txtRates(18) = "", 0, txtRates(18)) & ", DeductionAmount = " & IIf(txtRates(17) = "", 0, txtRates(17)) & _
                    ", DedStartPayment = '" & IIf(MaskEdBox5 = "__/__/____", #1/1/1900#, Format(MaskEdBox5, "mm/dd/yyyy")) & "', DedEndPayment = '" & IIf(MaskEdBox6 = "__/__/____", #1/1/1900#, Format(MaskEdBox6, "mm/dd/yyyy")) & _
                    "', TotalInterest = " & IIf(txtRates(15) = "", 0, txtRates(15)) & ", TotalDeduction = " & txtRates(16) & ", NoOfDeductions = " & IIf(txtRates(21) = "", 0#, txtRates(21)) & ", DeductionFlag = " & chkRates(12) & ", EntryBy = '" & mRem & "', EntryDate = '" & Format(Now, "mm/dd/yyyy") & "'" & _
                    ", Protectedflag = 1" & _
                    " Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode
                    '" Where AgencyCode = " & CompanyCode & " and clientcode = " & mCCode & " and EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode ' Updated by Bernard
            Call Action_Query(StrQuery, CnnstrHR)
            MsgBox "Deduction was marked as Fully Paid!"
        Else
            MsgBox "You need to update the Deduction Information or set Total Deduction equal to " & TotalPayments & ", if you do not want to deduc it from Payroll.", vbInformation, PayrollAlert
        End If
    Else
        res = MsgBox("Delete " & cmbTypes.Text & " deduction? No deduction will reflect to Payroll if you delete.", vbQuestion + vbYesNo, PayrollAlert)
        If res = vbYes Then ' delete deductions
            StrQuery = "DELETE FROM EmployeeDeductions " & _
                " Where EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode
                '" Where AgencyCode = " & CompanyCode & " and clientcode = " & mCCode & " and EmployeeID = " & txtEmpNum & " and DeductionCode = " & mDeductionCode ' updated by bernard
            Call Action_Query(StrQuery, CnnstrHR)
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
    
    Call EnableDisableDeductions(False)
    Call EnableDisableDeductionButtons("Add", "Edit", True)
    cmdEditDed.Caption = "Edit"
    cmbTypes.Enabled = True
    StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
    'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ProtectedFlag = 0"
    Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
    ' Re-filter deductions.
    Call FilterDeductions(txtEmpNum)
    
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
        EditDeductions = True
        txtRates(17).SetFocus
    Else
        'res = MsgBox("Cancel changes to Deductions?", vbQuestion + vbYesNo, PayrollAlert)
        'If res = vbYes Then ' cancel the changes
            Call EnableDisableDeductions(False)
            Call EnableDisableDeductionButtons("Add", "Edit", True)
            cmdEditDed.Caption = "Edit"
            EditDeductions = False
            AddDeductions = False
            cmbTypes.Enabled = True
            MsgBox "Changes to deductions was cancelled!", vbInformation, PayrollAlert
        'End If
    End If

NextProc:

End Sub

Private Sub cmdMarkActive_Click()
Dim res

    res = MsgBox("This will mark " & txtEmpName & " active and will be included in Payroll! Do you want to proceed?", vbQuestion + vbYesNo, PayrollAlert)
    If res = vbYes Then ' mark the employee active
        StrQuery = "UPDATE employeebasicsalary SET employeestatus = 1 WHERE EmployeeID = " & txtEmpNum ' txtControlNo
        Call Action_Query(StrQuery, cnnstr)
        Call Form_Load
        If ListView1(1).ListItems.Count = 0 Then
            cmdMarkActive.Enabled = False
            InActiveCtr = 0
        Else
            cmdMarkActive.Enabled = True
        End If
    End If
    
End Sub

Private Sub cmdNavigation_Click(Index As Integer)
Call Navigation(rsEmployeeBasicData, Me.cmdNavigation(Index), _
    Me.cmdNavigation(0), Me.cmdNavigation(1), Me.cmdNavigation(2), Me.cmdNavigation(3))

If AddDeductions = True Or EditDeductions = True Then
        Call EnableDisableDeductions(False)
        Call EnableDisableDeductionButtons("Add", "Edit", True)
        EditDeductions = False
        AddDeductions = False
        cmbTypes.Enabled = True
End If

'On Error Resume Next
If rsEmployeeBasicData.RecordCount <> 0 And SSTab1(1).Tab = 0 Then ' disable the add and edit buttons
    Call DisplayBasicData(ListView1(0).SelectedItem.Text, True)
End If
Call ScrollRec

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

Private Sub Form_Load()
Dim mComp As String
Dim mEmpN As String

    TimeKeepFlag = False
    If ActiveSw = False Then
        Me.Top = (mdiMain.Height - Me.Height) / 8
        Me.Left = (mdiMain.Width - Me.Width) / 2
    End If
    FormLoadSw = False
    
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

End Sub

Private Sub GetEmployeeBasic()
Dim rsTaxCode As ADODB.Recordset

'If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'    StrQuery = "SELECT * FROM clienttable WHERE CLIENTCODE2 = 'FMS' AND clientstatus = 1 and clientname is not null ORDER BY clientname"
'Else
'    StrQuery = "SELECT * FROM clienttable WHERE clientstatus = 1 and clientname is not null ORDER BY clientname"
'End If
'Get all active company.
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
    cmbCompany.ListIndex = 0
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
cmbCompany.ListIndex = 0

End Sub

Private Sub FilterEmployee(ByVal mCompany As String)

On Error GoTo Exit52
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

Exit52:

End Sub

Private Sub FilterDeductions(ByVal mEmployeeID As Integer)
    With rsEmpDeductions
        .Filter = "EmployeeID = " & mEmployeeID
        Call FillComboData(cmbTypes, rsEmpDeductions)
        If .RecordCount = 0 Then  'If .EOF
            Call EnableDisableDeductions("", True)
        Else
            cmbTypes.ListIndex = 0
        End If
    End With
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
    txtRates(4).Enabled = TFalse
    For i = 19 To 20
        txtRates(i).Enabled = TFalse
    Next i
    For i = 0 To 10 '7
        chkRates(i).Enabled = TFalse
    Next i
    'chkRates(1).Enabled = False ' enabled by bernard july 15, 2011
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
        ", " & chkRates(7) & ", '" & txtClient(5) & "', '" & txtClient(6) & "', '" & Trim(cmbTaxCode) & "', '" & txtClient(3) & "', '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & "')"
        mstr = "New record was saved!"
        mDetails = "New record was inserted in Employee basic salary."
    ElseIf EditFlag = True Then ' Save edit
        StrQuery = "UPDATE EmployeeBasicSalary SET StartContract = '" & Format(mskStartContract, "mm/dd/yyyy") & "', EndContract = '" & Format(mskEndContract, "mm/dd/yyyy") & "', NoofMonthsHired = " & txtRates(3) & ", MonthlyRate = " & txtRates(0) & ", DailyRate = " & txtRates(1) & ", HourlyRate = " & txtRates(2) & ", " & _
        "PaymentBasis = '" & Trim(cmbPayBasis) & "', PaymentSystem = '" & Trim(cmbPaySystem) & "', WithSSSDeduction = " & chkRates(3) & ", SSSRateperDay = " & txtRates(7) & ", WithPbigDeduction = " & chkRates(5) & ", PagibigRateperDay = " & txtRates(9) & ", WithPHealthDeduction = " & chkRates(4) & ", " & _
        "BasicSalaryFlag3 = " & chkRates(8) & ", BasicSalaryFlag4 = " & chkRates(9) & ", BasicSalaryFlag5 = " & chkRates(10) & ", " & _
        "SSSLoanAmount = " & txtRates(11) & ", SSSLoanPayment = " & txtRates(12) & ", StartSSSPayment = '" & IIf(mskStartSSS = "__/__/____", #1/1/1900#, Format(mskStartSSS, "mm/dd/yyyy")) & "', LastSSSPayment = '" & IIf(mskEndSSS = "__/__/____", #1/1/1900#, Format(mskEndSSS, "mm/dd/yyyy")) & "', PBigLoanAmount = " & txtRates(14) & ", PbigLoanPayment = " & txtRates(13) & ", StartPBigPayment = '" & IIf(mskStartPbig = "__/__/____", #1/1/1900#, Format(mskStartPbig, "mm/dd/yyyy")) & "', LastPbigPayment = '" & IIf(mskEndPbig = Null, #1/1/1900#, Format(mskEndPbig, "mm/dd/yyyy")) & "', " & _
        "ValeLoanAmount = " & txtRates(20) & ", ValeLoanPayment = " & txtRates(19) & ", StartValePayment = '" & IIf(mskStartVale = "__/__/____", #1/1/1900#, Format(mskStartVale, "mm/dd/yyyy")) & "', LastValePayment = '" & IIf(mskEndVale = "__/__/____", #1/1/1900#, Format(mskEndVale, "mm/dd/yyyy")) & "', " & _
        "PHealthRateperDay = " & txtRates(8) & ", WithTaxDeduction = " & chkRates(0) & ", WithEcolaPay = " & chkRates(1) & ", EcolaAmount = " & txtRates(4) & ", With13thMonthPay = " & chkRates(2) & ", Hourly13thMonthRate = " & txtRates(5) & ", AgencyBDOCardFee = " & txtRates(6) & ", AgencyTASAesla = " & txtRates(10) & ", ClientwServiceContract = " & chkRates(6) & ", " & _
        "ClientwCompPackage = " & chkRates(7) & ", IfnoCompensationPackage = '" & txtClient(5) & "', Remarks = '" & txtClient(6) & "', EmpTaxCode = '" & Trim(cmbTaxCode) & "', NotedBy = '" & txtClient(3) & "', Updatedby = '" & mUserID & "', DateUpdated = '" & Format(Now, "mm/dd/yyyy") & "'" & _
        " WHERE EmployeeID = " & txtEmpNum 'txtClient(0)
        mstr = "Edit was saved!"
        mDetails = "Edit was done in Employee basic salary."
    End If
    If Action_Query(StrQuery, cnnstr) = True Then
        MsgBox mstr, vbInformation, PayrollAlert
'        Call SaveUserTransactions(mDetails, mRemarks, txtEmpNum, txtClient(0))
    End If
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
    SSTab1(1).TabVisible(1) = True
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
    
    If EditRights = False Then
        MsgBox "You do not have rights to Edit Employee!", vbInformation, PayrollAlert
        Exit Sub
    End If
    
    EditFlag = True
    AmountSw = False
    'mRecNo = Text1
    Call DisableTextBoxes(True)
    Call DisableCommands(False)
    SSTab1(1).TabVisible(1) = False
    'cmbPayBasis.SetFocus
    'cmbTaxCode.SetFocus
    cmbPaySystem.SetFocus
End Sub

Private Sub cmdSave_Click()
    Call save
    SSTab1(1).TabVisible(0) = True
    SSTab1(1).TabVisible(1) = True
End Sub

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
                        Call EnableDisableDeductionButtons("Add", "Edit", True)
                        EditDeductions = False
                        AddDeductions = False
                        cmbTypes.Enabled = True
                        MsgBox "Changes to deductions was cancelled!", vbInformation, PayrollAlert
                    'End If
                End If
            End If
        End If
    End If
End Sub

Private Sub ListView1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1(Index).SortKey = ColumnHeader.Index - 1
    ListView1(Index).Sorted = True
End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Call DisplayBasicData(Item.Text, False)
    CurrEmployee = Item.Text
    'call DisplayBasicData(Item.ListSubItems(1).Text, False)
    If SSTab1(1).Tab = 1 Then
        txtEmpNum = Item.ListSubItems(7).Text
 '       TxtControlNo = Item.ListSubItems(3).Text
        txtEmpNo = Item.ListSubItems(3).Text
        txtEmpName = Item.ListSubItems(1).Text
        mCompany = Item.ListSubItems(2).Text
    End If
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
    SendKeys "{HOME}+{END}"
    MaskEdBox5.BackColor = &HFFFFC0
End Sub

Private Sub MaskEdBox5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub MaskEdbox5_LostFocus()
    MaskEdBox5.BackColor = &HFFFFFF
End Sub

Private Sub MaskEdBox6_GotFocus()
    SendKeys "{HOME}+{END}"
    MaskEdBox6.BackColor = &HFFFFC0
End Sub

Private Sub MaskEdBox6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
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
    SendKeys "{HOME}+{END}"
    mskStartContract.BackColor = &HFFFFC0
End Sub

Private Sub mskStartContract_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub mskStartContract_Validate(Cancel As Boolean)
    If mskStartContract <> "__/__/____" And Not IsDate(mskStartContract) Then
        MsgBox "Please enter a valid date!!!", vbInformation, PayrollAlert
        Cancel = True
    End If
End Sub

Private Sub mskEndContract_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub mskEndContract_GotFocus()
    SendKeys "{HOME}+{END}"
    mskEndContract.BackColor = &HFFFFC0
End Sub

Private Sub mskEndContract_Validate(Cancel As Boolean)
    If mskEndContract <> "__/__/____" And Not IsDate(mskEndContract) Then
        MsgBox "Please enter a valid date!!!", vbInformation, PayrollAlert
        Cancel = True
    End If
End Sub

Private Sub mskStartSSS_GotFocus()
    SendKeys "{HOME}+{END}"
    mskStartSSS.BackColor = &HFFFFC0
End Sub

Private Sub mskStartSSS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
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
    SendKeys "{HOME}+{END}"
    mskEndSSS.BackColor = &HFFFFC0
End Sub

Private Sub mskEndSSS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
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
    SendKeys "{HOME}+{END}"
    mskStartPbig.BackColor = &HFFFFC0
End Sub

Private Sub mskStartPbig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
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
    SendKeys "{HOME}+{END}"
    mskEndPbig.BackColor = &HFFFFC0
End Sub

Private Sub mskEndPbig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
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
    SendKeys "{HOME}+{END}"
    mskEndVale.BackColor = &HFFFFC0
End Sub

Private Sub mskEndVale_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
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
    SendKeys "{HOME}+{END}"
    mskStartVale.BackColor = &HFFFFC0
End Sub

Private Sub mskStartVale_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
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

On Error GoTo ExitSub1
If rsEmployeeBasicData.RecordCount = 0 Then Exit Sub
    
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
        
        ' Requery all Active Employee Deductions.
        StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
        'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ProtectedFlag = 0"
        Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
        Call FilterDeductions(txtEmpNum)
    End If
End With

ExitSub1:

End Sub

Private Sub txtClient_Validate(Index As Integer, Cancel As Boolean)
    If txtClient(0) = "" Then MsgBox "Client name should not be blank!", vbInformation, PayrollAlert: Cancel = True
End Sub

Private Sub txtRates_GotFocus(Index As Integer)
    SendKeys "{HOME}+{END}"
    txtRates(Index).BackColor = &HFFFFC0
End Sub

Private Sub txtRates_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strValid
    strValid = "0123456789." & Chr(13) & Chr(8)
    If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys "{TAB}"
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

Public Sub CommandPass(ByVal strCommand As String)

Select Case strCommand
    Case "Close"
        Unload Me
End Select

End Sub

