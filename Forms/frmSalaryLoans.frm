VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBasicSalary 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Basic Salary"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmBasicSalary.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9045
   Begin TabDlg.SSTab SSTab1 
      Height          =   6360
      Index           =   1
      Left            =   90
      TabIndex        =   46
      Top             =   540
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   11218
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
      TabCaption(0)   =   "Employee Basic Salary"
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
         TabIndex        =   99
         Top             =   6030
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   2
         Left            =   8040
         Picture         =   "frmBasicSalary.frx":0AF4
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "First"
         Top             =   6030
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   1
         Left            =   7755
         Picture         =   "frmBasicSalary.frx":0BEE
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "First"
         Top             =   6030
         Width           =   300
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   270
         Index           =   0
         Left            =   7470
         Picture         =   "frmBasicSalary.frx":0CE8
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "First"
         Top             =   6030
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
         Height          =   5820
         Index           =   4
         Left            =   90
         TabIndex        =   49
         Top             =   360
         Width           =   8655
         Begin VB.TextBox txtEmpNum 
            Height          =   285
            Left            =   990
            TabIndex        =   95
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
            Left            =   1890
            TabIndex        =   39
            Text            =   "Client Name"
            Top             =   5040
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
            TabIndex        =   38
            Top             =   4815
            Value           =   1  'Checked
            Width           =   2985
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
            Left            =   7740
            TabIndex        =   28
            Text            =   "50.00"
            Top             =   3825
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
            Left            =   5040
            TabIndex        =   27
            Text            =   "5.5"
            Top             =   3825
            Width           =   735
         End
         Begin VB.TextBox txtClientCode 
            Height          =   285
            Left            =   1800
            TabIndex        =   83
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
            Left            =   5940
            TabIndex        =   37
            Top             =   4815
            Value           =   1  'Checked
            Width           =   2535
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
            Left            =   5940
            TabIndex        =   25
            Top             =   3555
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
            Left            =   5940
            TabIndex        =   23
            Top             =   3195
            Value           =   1  'Checked
            Width           =   1725
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
            Left            =   5940
            TabIndex        =   21
            Top             =   2835
            Value           =   1  'Checked
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
            Height          =   315
            Index           =   9
            Left            =   7740
            TabIndex        =   26
            Text            =   "0.00"
            Top             =   3465
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
            Left            =   7740
            TabIndex        =   24
            Text            =   "0.00"
            Top             =   3150
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
            Index           =   7
            Left            =   7740
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   2835
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
            Index           =   5
            Left            =   5040
            TabIndex        =   19
            Text            =   "20.83"
            Top             =   3195
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
            Index           =   4
            Left            =   5040
            TabIndex        =   17
            Text            =   "50.00"
            Top             =   2835
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
            Index           =   0
            Left            =   4995
            TabIndex        =   9
            Text            =   "15"
            Top             =   1710
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
            Index           =   6
            Left            =   900
            TabIndex        =   40
            Text            =   "Client Name"
            Top             =   5445
            Width           =   1935
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
            Index           =   4
            Left            =   6030
            TabIndex        =   44
            Text            =   "Client Name"
            Top             =   5445
            Width           =   1125
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
            Left            =   3600
            TabIndex        =   41
            Text            =   "Client Name"
            Top             =   5445
            Width           =   1485
         End
         Begin VB.Frame Frame3 
            Height          =   1365
            Left            =   3240
            TabIndex        =   63
            Top             =   0
            Width           =   5415
            Begin VB.ComboBox cmbTaxCode 
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
               Left            =   4320
               List            =   "frmBasicSalary.frx":1224
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
               Left            =   1890
               MultiLine       =   -1  'True
               TabIndex        =   4
               Text            =   "frmBasicSalary.frx":1238
               Top             =   360
               Width           =   2475
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
               Width           =   1575
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
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
               Left            =   1890
               TabIndex        =   64
               Top             =   150
               Width           =   1815
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1455
            Left            =   3240
            TabIndex        =   54
            Top             =   1305
            Width           =   5415
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
               TabIndex        =   15
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
               TabIndex        =   11
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
               Left            =   2970
               TabIndex        =   10
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
               ItemData        =   "frmBasicSalary.frx":1244
               Left            =   180
               List            =   "frmBasicSalary.frx":1251
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
               ItemData        =   "frmBasicSalary.frx":126D
               Left            =   180
               List            =   "frmBasicSalary.frx":127D
               Style           =   2  'Dropdown List
               TabIndex        =   12
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
               TabIndex        =   14
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
               TabIndex        =   62
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
               TabIndex        =   61
               Top             =   180
               Width           =   1455
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Rate:"
               BeginProperty Font 
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
               Left            =   1800
               TabIndex        =   60
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Daily Rate:"
               BeginProperty Font 
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
               Left            =   2970
               TabIndex        =   59
               Top             =   180
               Width           =   915
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "Hourly Rate:"
               BeginProperty Font 
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
               TabIndex        =   58
               Top             =   180
               Width           =   915
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
               TabIndex        =   57
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
               TabIndex        =   56
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
               TabIndex        =   55
               Top             =   765
               Width           =   915
            End
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
            Left            =   3330
            TabIndex        =   20
            Top             =   3555
            Value           =   1  'Checked
            Width           =   1815
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
            ItemData        =   "frmBasicSalary.frx":1298
            Left            =   180
            List            =   "frmBasicSalary.frx":12A5
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   450
            Width           =   2925
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   3030
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   1080
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   5345
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
            Left            =   3330
            TabIndex        =   16
            Top             =   2835
            Value           =   1  'Checked
            Width           =   1725
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
            Left            =   3330
            TabIndex        =   18
            Top             =   3195
            Width           =   1725
         End
         Begin VB.Frame Frame5 
            Height          =   735
            Left            =   180
            TabIndex        =   75
            Top             =   4050
            Width           =   8295
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
               Left            =   4230
               TabIndex        =   33
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
               Index           =   13
               Left            =   5130
               TabIndex        =   34
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
               Index           =   12
               Left            =   990
               TabIndex        =   30
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
               Left            =   90
               TabIndex        =   29
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
               Left            =   1890
               TabIndex        =   31
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
               Left            =   2970
               TabIndex        =   32
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
               Left            =   6030
               TabIndex        =   35
               Top             =   360
               Width           =   1105
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
               Left            =   7110
               TabIndex        =   36
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
               Left            =   4320
               TabIndex        =   94
               Top             =   180
               Width           =   825
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "PBig Paymnt"
               BeginProperty Font 
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
               Left            =   5130
               TabIndex        =   93
               Top             =   180
               Width           =   915
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
               Left            =   6210
               TabIndex        =   92
               Top             =   180
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
               Left            =   7290
               TabIndex        =   91
               Top             =   180
               Width           =   915
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
               Left            =   3150
               TabIndex        =   90
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
               Left            =   2070
               TabIndex        =   89
               Top             =   180
               Width           =   1005
            End
            Begin VB.Label lblclient 
               BackStyle       =   0  'Transparent
               Caption         =   "SSS Paymnt:"
               BeginProperty Font 
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
               Left            =   990
               TabIndex        =   88
               Top             =   165
               Width           =   915
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
               Left            =   180
               TabIndex        =   85
               Top             =   165
               Width           =   825
            End
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Bond:"
            BeginProperty Font 
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
            Left            =   6750
            TabIndex        =   87
            Top             =   3870
            Width           =   1095
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Card Trans. Fee:"
            BeginProperty Font 
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
            Left            =   3690
            TabIndex        =   86
            Top             =   3870
            Width           =   1275
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
            TabIndex        =   74
            Top             =   840
            Width           =   825
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry Date:"
            BeginProperty Font 
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
            Left            =   5130
            TabIndex        =   69
            Top             =   5490
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
            Left            =   2880
            TabIndex        =   68
            Top             =   5490
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
            Top             =   5085
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
            Left            =   180
            TabIndex        =   50
            Top             =   5445
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Height          =   6000
         Left            =   -74910
         TabIndex        =   70
         Top             =   360
         Width           =   8700
         Begin MSComctlLib.ListView ListView1 
            Height          =   5055
            Index           =   1
            Left            =   180
            TabIndex        =   73
            Top             =   810
            Width           =   8310
            _ExtentX        =   14658
            _ExtentY        =   8916
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
            TabIndex        =   84
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
            TabIndex        =   82
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
            TabIndex        =   81
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
            TabIndex        =   72
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
            TabIndex        =   71
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
            Picture         =   "frmBasicSalary.frx":12C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   2745
         Picture         =   "frmBasicSalary.frx":1326
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   2175
         Picture         =   "frmBasicSalary.frx":1428
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmBasicSalary.frx":1572
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "New"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1080
         Picture         =   "frmBasicSalary.frx":1AA4
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1620
         Picture         =   "frmBasicSalary.frx":1FD6
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   540
         Picture         =   "frmBasicSalary.frx":2508
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdClose 
         Height          =   405
         Left            =   3285
         Picture         =   "frmBasicSalary.frx":2652
         Style           =   1  'Graphical
         TabIndex        =   76
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
      TabIndex        =   48
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
      TabIndex        =   45
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
Dim rsClientData As ADODB.Recordset ' for later deletion, it is already declare in main
Dim strQuery As String
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
        
Dim workdaysperyr As Integer
Dim noofdayspermon As Integer

Dim AddFlag As Boolean
Dim EditFlag As Boolean

Private Sub cmbCompany_Change()
'   Call FilterEmployee(cmbCompany.Text)
'    If ActiveCtr = 0 Then
'        Call DisplayBasicData("", False)
'    Else
'        Call DisplayBasicData(ListView1(0).SelectedItem.Text, False)
'    End If
'    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

'Private Sub cmbCutoff_Click()
'    If cmbCutoff.Text = "Weekly" Then
'        txtRates(0) = 7
'    ElseIf cmbCutoff.Text = "Semi-Monthly" Then
'        txtRates(0) = 15
'    Else
'        txtRates(0) = 30
'    End If
'End Sub

Private Sub cmbTaxCode_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub cmbTaxCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbCompany_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub cmbPayBasis_GotFocus()
    'SendKeys "{HOME}+{END}"
    cmbPayBasis.Text = paymentbase
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
'    cmbCompany.Refresh
End Sub

Private Sub cmbCompany_KeyPress(KeyAscii As Integer)
'    Call FilterEmployee(cmbCompany.Text)
'    If ActiveCtr = 0 Then
'        Call DisplayBasicData("", False)
'    Else
'        Call DisplayBasicData(ListView1(0).SelectedItem.Text, False)
'    End If
'    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdClose_Click()
Set rsEmployeeBasicData = Nothing
Unload Me
End Sub

Private Sub cmdMarkActive_Click()
Dim res

    res = MsgBox("This will mark " & txtEmpName & " active and will be included in Payroll! Do you want to proceed?", vbQuestion + vbYesNo, PayrollAlert)
    If res = vbYes Then ' mark the employee active
        strQuery = "UPDATE employeebasicsalary SET employeestatus = true WHERE employeenum = " & txtEmpNum ' txtempno
        Call Action_Query(strQuery, Cnnstr)
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
If SSTab1(1).Tab = 0 Then ' disable the add and edit buttons
    Call DisplayBasicData(ListView1(0).SelectedItem.Text, True)
End If
Call ScrollRec
End Sub

Private Sub cmdRefresh_Click()
    mCompany = Trim(cmbCompany)
    mEmpName = Trim(txtClient(2))
    Call Form_Load
    If mCompany <> "" Then
        cmbCompany = mCompany
    End If
    Call DisplayBasicData(mEmpName, True)
    If SSTab1(1).Tab = 1 Then ' disable the add and edit buttons
        cmdadd.Enabled = False
        cmdedit.Enabled = False
    Else
        cmdadd.Enabled = True
        cmdedit.Enabled = True
    End If
    'If RefreshSw = True Then
        MsgBox "Data was updated!", vbInformation, PayrollAlert
    'End If
End Sub

Private Sub Form_Load()

    If ActiveSw = False Then
        Me.Top = (mdiMain.Height - Me.Height) / 8
        Me.Left = (mdiMain.Width - Me.Width) / 2
    End If
    
    mUserID = UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID)))
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
    chkActiveClient.Enabled = False
    If ActiveCtr = 0 Then
        Call DisplayBasicData("", True)
    Else
        'Call DisplayBasicData(ListView1(0).SelectedItem.Text, True)
        Call DisplayBasicData(CurrEmployee, True)
    End If
End Sub

Private Sub GetEmployeeBasic()
Dim rsTaxCode As ADODB.Recordset

' Get all active company
strQuery = "SELECT * FROM clienttable WHERE clientstatus = true and clientname is not null ORDER BY clientname"
Set rsClientData = ReturnRst(strQuery, Cnnstr, mstatic)

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

strQuery = "SELECT a.*, (trim(b.surname) + ', ' + trim(b.firstname)) AS fullname, " & _
            "b.controlno AS controlno2, b.sss, b.philhealth, b.pagibig, b.acctno, b.taxid, c.employmentstatusdesc, d.clientname " & _
            "FROM ((employeebasicsalary AS a LEFT JOIN HREmployees AS b ON a.employeenum = b.employeenum) " & _
            "LEFT JOIN HREmploymentStatus AS c ON b.EmploymentStatus = c.EmploymentStatus) " & _
            "LEFT JOIN ClientTable AS d ON a.clientcode = d.clientcode " & _
            " ORDER BY b.surname, b.firstname"
'SELECT FROM ((EmployeeBasicSalary INNER JOIN HREmployees ON EmployeeBasicSalary.EmployeeNum = HREmployees.EmployeeNum) INNER JOIN HREmploymentStatus ON HREmployees.EmploymentStatus = HREmploymentStatus.EmploymentStatus) INNER JOIN ClientTable ON EmployeeBasicSalary.ClientCode = ClientTable.ClientCode;

Set rsEmployeeBasicData = ReturnRst(strQuery, Cnnstr, mstatic)
If rsEmployeeBasicData.RecordCount <> 0 Then
    Call FilterEmployee(rsClientData!ClientName)
Else
    ' No client selected disable text boxes and command buttons
    Call DisableCommands(True)
    Call DisableTextBoxes(False)
End If

strQuery = "SELECT id, taxexemptioncode FROM taxexemptioncode ORDER BY id"
Set rsTaxCode = ReturnRst(strQuery, Cnnstr, mstatic)

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

Private Sub DisableCommands(ByVal mSw As Boolean)

If AddFlag = True Or EditFlag = True Then
    cmdadd.Enabled = False
    cmdedit.Enabled = False
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
    cmddelete.Enabled = False
    cmdrefresh.Enabled = False
    cmdClose.Enabled = False
Else
    cmdadd.Enabled = True
    cmdedit.Enabled = True
    cmdsave.Enabled = False
    cmdcancel.Enabled = False
    cmddelete.Enabled = True
    cmdrefresh.Enabled = True
    cmdClose.Enabled = True
End If
If mSw = True Then ' no employee records selected, enable only add and exit button
    cmdedit.Enabled = False
    cmdrefresh.Enabled = False
    cmddelete.Enabled = False
End If

End Sub

Private Sub DisableTextBoxes(ByVal TFalse As Boolean)
    
    txtClient(3).Enabled = TFalse
    For i = 5 To 6
        txtClient(i).Enabled = TFalse
    Next i
    For i = 0 To 14
        txtRates(i).Enabled = TFalse
    Next i
'    For i = 0 To 7
'        chkRates(i).Enabled = TFalse
'    Next i
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
    cmbTaxCode.Enabled = TFalse
    cmbPaySystem.Enabled = TFalse
    cmbPayBasis.Enabled = TFalse
    mskStartContract.Enabled = TFalse
    mskEndContract.Enabled = TFalse
    mskStartSSS.Enabled = TFalse
    mskEndSSS.Enabled = TFalse
    mskStartPbig.Enabled = TFalse
    mskEndPbig.Enabled = TFalse
    
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
'            litem2.SubItems(2) = IIf(IsNull(mRs!EmployeeNum), "Blank", mRs!EmployeeNum)
'            litem2.SubItems(3) = IIf(IsNull(mRs!StartContract), "Blank", Format(mRs!StartContract, "mm/dd/yyyy"))
'            litem2.SubItems(4) = IIf(IsNull(mRs!EndContract), "Blank", Format(mRs!EndContract, "mm/dd/yyyy"))
'            InActiveCtr = InActiveCtr + 1
        End If
        mRs.MoveNext
    Wend
    'Get all Inactive employees regardless of employers.
    mRs.MoveFirst
    mRs.Filter = "employeestatus = false"
    While Not (mRs.EOF Or mRs.BOF)
        Set litem2 = ListView1(1).ListItems.Add(, , 1)
        'Set litem2 = ListView1(1).ListItems.Add(, , IIf(IsNull(mRs!FullName, "Blank", mRs!FullName)))
        litem2.SubItems(1) = IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName)
        litem2.SubItems(2) = IIf(IsNull(mRs!ClientName), "Blank", mRs!ClientName)
        'litem2.SubItems(3) = IIf(IsNull(mRs!employeenum), "Blank", mRs!employeenum)
        litem2.SubItems(3) = IIf(IsNull(mRs!ControlNo2), "Blank", mRs!ControlNo2)
        litem2.SubItems(4) = IIf(IsNull(mRs!StartContract) Or mRs!StartContract = #1/1/1900#, "__/__/____", Format(mRs!StartContract, "mm/dd/yyyy"))
        litem2.SubItems(5) = IIf(IsNull(mRs!EndContract) Or mRs!EndContract = #1/1/1900#, "__/__/____", Format(mRs!EndContract, "mm/dd/yyyy"))
        litem2.SubItems(6) = IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName)
        litem2.SubItems(7) = IIf(IsNull(mRs!employeenum), 1, mRs!employeenum)
        InActiveCtr = InActiveCtr + 1
        mRs.MoveNext
    Wend
End Sub

Private Function FillComboData(ByRef mCmb As ComboBox, _
                        mRs As Recordset)
    mCmb.Clear
    While Not (mRs.EOF Or mRs.BOF)
        mCmb.AddItem Trim(Replace(mRs.Fields(1), "'", "''"))
        mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
        mRs.MoveNext
    Wend
End Function

Private Sub cmdAdd_Click()
Dim rsNewEmp As ADODB.Recordset

    ' Check whether there are new employees entered from HR table.
    strQuery = "SELECT surname, firstname, employeenum, controlno FROM hremployees WHERE clientcode = " & _
        txtClientCode & " and employeenum not in (select employeenum from employeebasicsalary) ORDER BY surname, firstname"
    Set rsNewEmp = ReturnRst(strQuery, Cnnstr, mstatic)
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

    mCompany = Trim(cmbCompany)
    mEmpName = Trim(txtClient(2))
    For i = 0 To 6
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
    If AddFlag = True Then ' Add New Record
'        strQuery = "INSERT INTO EmployeeBasicSalary ( ClientName, ClientAddress, ClientTelNo, ClientTeleFax, ClientOwner, ClientContactPerson, " & _
'        "ClientContactPosition, ClientAssistant, ClientAssistantPosition, ClientBankName, ClientBankAccountNo, ClientPhilHealth, ClientSSSNo, " & _
'        "ClientTaxNumber, ClientPagibig, ClientCutoff, WorkHrsPerDay, WorkingDaysPerYear, NoofDaysperMonth, ClientCutoffDates, OverTimeRate, SundayRate, " & _
'        "TardinessRate, NDiffRate, RestDayRate, OTwoPremiumRate, SpecialHolidayRate, LegalHolidayRate, LateChargesinMin, " & _
'        "ClientAgencyFlag, ClientAgencyRate, ClientChargeperDay, ClientEcolaFlag, ECOLARate, With13MonthPay, Month13RatePerDay, ClientTaxFlag, ClientTaxPercentFlag, " & _
'        "ClientTaxPercentage, ClientSSSFlag, ClientSSSRate, ClientPhilHealthFlag, ClientPHRateperDay, ClientPagibigflag, ClientPagibigRate, WithVatCharges, VatPercentRate, " & _
'        "ClientEcolaTaxFlag, EcolawLateDeduction, EcolaDoubleonLegal, EcolaDoubleonSpl, PickupofDTR, SalaryRelease, BillingSchedule, Updatedby, ClientStatus ) " & _
'        " VALUES ( '" & txtClient(0) & "', '" & txtClient(1) & "', '" & txtClient(2) & "', '" & txtClient(3) & "', '" & txtClient(4) & "', '" & txtClient(5) & "', '" & _
'        txtClient(6) & "', '" & txtClient(7) & "', '" & txtClient(8) & "', '" & txtClient(9) & "', '" & txtClient(10) & "', '" & txtClient(11) & "', '" & txtClient(12) & "', '" & _
'        txtClient(13) & "', '" & txtClient(14) & "', '" & cmbCutoff & "', " & txtRates(0) & ", " & txtRates(1) & ", " & txtRates(2) & ", " & txtRates(3) & ", " & txtRates(5) & ", " & txtRates(6) & ", " & _
'        txtRates(7) & ", " & txtRates(10) & ", " & txtRates(11) & ", " & txtRates(12) & ", " & txtRates(8) & ", " & txtRates(13) & ", " & txtRates(16) & ", " & _
'        IIf(optCharge1 = True, 1, 0) & ", " & txtRates(4) & ", " & txtRates(9) & ", " & chkRates(1) & ", " & txtRates(14) & ", " & chkRates(4) & ", " & txtRates(20) & ", " & chkRates(7) & ", " & chkRates(3) & ", " & _
'        txtRates(15) & ", " & chkRates(5) & ", " & txtRates(18) & ", " & chkRates(6) & ", " & txtRates(19) & ", " & chkRates(2) & ", " & txtRates(21) & ", " & chkRates(13) & ", " & txtRates(17) & ", " & _
'        chkRates(9) & ", " & chkRates(10) & ", " & chkRates(12) & ", " & chkRates(11) & ", '" & txtClient(16) & "', '" & txtClient(17) & "', '" & txtClient(18) & "', '" & txtClient(19) & "', " & chkActiveClient & ")"

        strQuery = "INSERT INTO EmployeeBasicSalary (AgencyCode, ClientCode, Employeenum, ControlNo, EmployeeStatus, StartContract, EndContract, NoofMonthsHired, MonthlyRate,DailyRate,  HourlyRate, " & _
        " PaymentBasis, PaymentSystem, WithSSSDeduction, SSSRateperDay, WithPbigDeduction, PagibigRateperDay, WithPHealthDeduction, " & _
        " SSSLoanAmount, SSSLoanPayment, StartSSSPayment, LastSSSPayment, PbigLoanAmount, PbigLoanPayment, StartPBigPayment, LastPBigPayment," & _
        " PHealthRateperDay, WithTaxDeduction, WithEcolaPay, EcolaAmount, With13thMonthPay, Hourly13thMonthRate, AgencyBDOCardFee, AgencyCashBond, ClientwServiceContract, " & _
        " ClientwCompPackage, IfnoCompensationPackage , Remarks, EmpTaxCode, NotedBy, Updatedby, DateUpdated ) " & _
        " VALUES ( 1," & txtClientCode & ", " & txtEmpNum & ", '" & txtClient(0) & "', 1, '" & IIf(mskStartContract = "__/__/____", #1/1/1900#, Format(mskStartContract, "mm/dd/yyyy")) & "', '" & IIf(mskEndContract = "__/__/____", #1/1/1900#, Format(mskEndContract, "mm/dd/yyyy")) & "', " & txtRates(3) & ", " & txtRates(0) & ", " & txtRates(1) & ", " & txtRates(2) & _
        ", '" & Trim(cmbPayBasis) & "', '" & Trim(cmbPaySystem) & "', " & chkRates(3) & ", " & txtRates(7) & ", " & chkRates(5) & ", " & txtRates(9) & ", " & chkRates(4) & _
        ", " & txtRates(11) & ", " & txtRates(12) & ", '" & IIf(mskStartSSS = "__/__/____", #1/1/1900#, Format(mskStartSSS, "mm/dd/yyyy")) & "', '" & IIf(mskEndSSS = "__/__/____", #1/1/1900#, Format(mskEndSSS, "mm/dd/yyyy")) & "', " & txtRates(14) & ", " & txtRates(13) & ", '" & IIf(mskStartPbig = "__/__/____", #1/1/1900#, Format(mskStartPbig, "mm/dd/yyyy")) & "', '" & IIf(mskEndPbig = "__/__/____", #1/1/1900#, Format(mskEndPbig, "mm/dd/yyyy")) & "'" & _
        ", " & txtRates(8) & ", " & chkRates(0) & ", " & chkRates(1) & ", " & txtRates(4) & ", " & chkRates(2) & ", " & txtRates(5) & ", " & txtRates(6) & ", " & txtRates(10) & ", " & chkRates(6) & _
        ", " & chkRates(7) & ", '" & txtClient(5) & "', '" & txtClient(6) & "', '" & Trim(cmbTaxCode) & "', '" & txtClient(3) & "', '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & "')"
        
        mstr = "New record was saved!"
    ElseIf EditFlag = True Then ' save edit
        strQuery = "UPDATE EmployeeBasicSalary SET StartContract = '" & Format(mskStartContract, "mm/dd/yyyy") & "', EndContract = '" & Format(mskEndContract, "mm/dd/yyyy") & "', NoofMonthsHired = " & txtRates(3) & ", MonthlyRate = " & txtRates(0) & ", DailyRate = " & txtRates(1) & ", HourlyRate = " & txtRates(2) & ", " & _
        "PaymentBasis = '" & Trim(cmbPayBasis) & "', PaymentSystem = '" & Trim(cmbPaySystem) & "', WithSSSDeduction = " & chkRates(3) & ", SSSRateperDay = " & txtRates(7) & ", WithPbigDeduction = " & chkRates(5) & ", PagibigRateperDay = " & txtRates(9) & ", WithPHealthDeduction = " & chkRates(4) & ", " & _
        "SSSLoanAmount = " & txtRates(11) & ", SSSLoanPayment = " & txtRates(12) & ", StartSSSPayment = '" & IIf(mskStartSSS = "__/__/____", #1/1/1900#, Format(mskStartSSS, "mm/dd/yyyy")) & "', LastSSSPayment = '" & IIf(mskEndSSS = "__/__/____", #1/1/1900#, Format(mskEndSSS, "mm/dd/yyyy")) & "', PBigLoanAmount = " & txtRates(14) & ", PbigLoanPayment = " & txtRates(13) & ", StartPBigPayment = '" & IIf(mskStartPbig = "__/__/____", #1/1/1900#, Format(mskStartPbig, "mm/dd/yyyy")) & "', LastPbigPayment = '" & IIf(mskEndPbig = Null, #1/1/1900#, Format(mskEndPbig, "mm/dd/yyyy")) & "', " & _
        "PHealthRateperDay = " & txtRates(8) & ", WithTaxDeduction = " & chkRates(0) & ", WithEcolaPay = " & chkRates(1) & ", EcolaAmount = " & txtRates(4) & ", With13thMonthPay = " & chkRates(2) & ", Hourly13thMonthRate = " & txtRates(5) & ", AgencyBDOCardFee = " & txtRates(6) & ", AgencyCashBond = " & txtRates(10) & ", ClientwServiceContract = " & chkRates(6) & ", " & _
        "ClientwCompPackage = " & chkRates(7) & ", IfnoCompensationPackage = '" & txtClient(5) & "', Remarks = '" & txtClient(6) & "', EmpTaxCode = '" & Trim(cmbTaxCode) & "', NotedBy = '" & txtClient(3) & "', Updatedby = '" & mUserID & "', DateUpdated = '" & Format(Now, "mm/dd/yyyy") & "'" & _
        " WHERE employeenum = " & txtEmpNum 'txtClient(0)
        mstr = "Edit was saved!"
    End If
    Call Action_Query(strQuery, Cnnstr)
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
    SSTab1(1).TabVisible(1) = True
End Sub

Private Sub cmdDelete_Click()
    Dim strID As Integer
    Dim employeeMsg As String
    
    If SSTab1(1).Tab = 1 Then
        strQuery = "DELETE FROM EmployeeBasicSalary WHERE employeenum = " & txtEmpNum ' txtempno
        employeeMsg = "This will permanently delete " & txtEmpName & ". Do you want to proceed?"
    Else
        strQuery = "UPDATE EmployeeBasicSalary SET employeestatus = false WHERE employeenum = " & txtEmpNum ' txtClient(0)
        employeeMsg = "This will mark " & Trim(txtClient(2)) & " inactive. Do you want to proceed?"
    End If
    If MsgBox(employeeMsg, vbQuestion + vbYesNo) = vbYes Then
        Call Action_Query(strQuery, Cnnstr)
        mCompany = Trim(cmbCompany)
        mEmpName = Trim(txtClient(2))
        'Call Form_Load
        'cmbCompany = mCompany
    
        strQuery = "SELECT a.*, (trim(b.surname) + ', ' + trim(b.firstname)) AS fullname, " & _
            "b.controlno AS controlno2, b.sss, b.philhealth, b.pagibig, b.acctno, b.taxid, c.employmentstatusdesc, d.clientname " & _
            "FROM ((employeebasicsalary AS a LEFT JOIN HREmployees AS b ON a.employeenum = b.employeenum) " & _
            "LEFT JOIN HREmploymentStatus AS c ON b.EmploymentStatus = c.EmploymentStatus) " & _
            "LEFT JOIN ClientTable AS d ON a.clientcode = d.clientcode " & _
            " ORDER BY b.surname, b.firstname"
        
        Set rsEmployeeBasicData = Nothing
        Set rsEmployeeBasicData = ReturnRst(strQuery, Cnnstr, mstatic)
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
    EditFlag = True
    AmountSw = False
    'mRecNo = Text1
    Call DisableTextBoxes(True)
    Call DisableCommands(False)
    SSTab1(1).TabVisible(1) = False
    'cmbPayBasis.SetFocus
    cmbTaxCode.SetFocus
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
End Sub

Private Sub ListView1_Click(Index As Integer)
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
        Else
            If ListView1(0).ListItems.Count = 0 Then
                Call DisplayBasicData("", False)
            Else
                Call DisplayBasicData(ListView1(0).SelectedItem.Text, False)
                CurrEmployee = ListView1(0).SelectedItem.Text
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
    txtRates(4) = 6.5
    txtRates(9) = 10#
    txtRates(5) = 1.25
    txtRates(6) = 1.3
    txtRates(7) = 1
    txtRates(10) = 1.5
    txtRates(8) = 1.3
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

If rsEmployeeBasicData.RecordCount = 0 Then Exit Sub
    
With rsEmployeeBasicData
    If recSw = False Then
        .MoveFirst
        .Find "fullname = '" & SelectedEmployee & "'"
    End If
    If Not .EOF And Not .BOF Then
        chkActiveClient = IIf(!EmployeeStatus = True, 1, 0)
        txtEmpNum = IIf(IsNull(!employeenum), 0, !employeenum)
        txtClient(0) = IIf(IsNull(!ControlNo2), "", !ControlNo2)
        txtClient(1) = IIf(IsNull(!EmploymentStatusdesc), "", !EmploymentStatusdesc)
        txtClient(2) = IIf(IsNull(!FullName), "", !FullName)
        txtClient(3) = IIf(IsNull(!NotedBy), "", !NotedBy)
        txtClient(4) = IIf(IsNull(!EntryDate), "", Format(!EntryDate, "mm/dd/yyyy"))
        txtClient(5) = IIf(IsNull(!IfnoCompensationPackage), "", !IfnoCompensationPackage)
        txtClient(6) = IIf(IsNull(!Remarks), "", !Remarks)
        
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
        txtRates(10) = IIf(!AgencyCashBond = "" Or IsNull(!AgencyCashBond), 0#, Format(!AgencyCashBond, "###0.00"))
        
        txtRates(11) = IIf(!SSSLoanAmount = "" Or IsNull(!SSSLoanAmount), 0#, Format(!SSSLoanAmount, "###0.00"))
        txtRates(12) = IIf(!SSSLoanPayment = "" Or IsNull(!SSSLoanPayment), 0#, Format(!SSSLoanPayment, "###0.00"))
        txtRates(14) = IIf(!PbigLoanAmount = "" Or IsNull(!PbigLoanAmount), 0#, Format(!PbigLoanAmount, "###0.00"))
        txtRates(13) = IIf(!PbigLoanPayment = "" Or IsNull(!PbigLoanPayment), 0#, Format(!PbigLoanPayment, "###0.00"))
        
        cmbTaxCode = IIf(IsNull(!EmpTaxCode), "HF", !EmpTaxCode)
        cmbPayBasis = IIf(IsNull(!PaymentBasis), "Daily", !PaymentBasis)
        paymentbase = IIf(IsNull(!PaymentBasis), "Daily", !PaymentBasis)
        cmbPaySystem = IIf(IsNull(!PaymentSystem), "ATM", !PaymentSystem)
        
        mskStartContract = IIf(IsNull(!StartContract) Or !StartContract = #1/1/1900#, "__/__/____", Format(!StartContract, "mm/dd/yyyy"))
        mskEndContract = IIf(IsNull(!EndContract) Or !EndContract = #1/1/1900#, "__/__/____", Format(!EndContract, "mm/dd/yyyy"))
        ' SSS Loan Dates
        mskStartSSS = IIf(IsNull(!StartSSSPayment) Or !StartSSSPayment = #1/1/1900#, "__/__/____", Format(!StartSSSPayment, "mm/dd/yyyy"))
        mskEndSSS = IIf(IsNull(!LastSSSPayment) Or !LastSSSPayment = #1/1/1900#, "__/__/____", Format(!LastSSSPayment, "mm/dd/yyyy"))
        mskStartPbig = IIf(IsNull(!StartPbigPayment) Or !StartPbigPayment = #1/1/1900#, "__/__/____", Format(!StartPbigPayment, "mm/dd/yyyy"))
        mskEndPbig = IIf(IsNull(!LastPbigPayment) Or !LastPbigPayment = #1/1/1900#, "__/__/____", Format(!LastPbigPayment, "mm/dd/yyyy"))
        
        chkRates(3) = IIf(!WithSSSDeduction = True, 1, 0)
        chkRates(5) = IIf(!WithPbigDeduction = True, 1, 0)
        chkRates(4) = IIf(!WithPHealthDeduction = True, 1, 0)
        chkRates(0) = IIf(!WithTaxDeduction = True, 1, 0)
        chkRates(1) = IIf(!WithEcolaPay = True, 1, 0)
        chkRates(2) = IIf(!With13thMonthPay = True, 1, 0)
        chkRates(6) = IIf(!ClientwServiceContract = True, 1, 0)
        chkRates(7) = IIf(!ClientwCompPackage = True, 1, 0)
    End If
End With

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
            If DeciSw = True And Mid(mAmt, i, 1) = "." Then
            Else
                strAmt = strAmt & Mid(mAmt, i, 1)
            End If
            If DeciSw = False And Mid(mAmt, i, 1) = "." Then
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

