VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmFixedRates 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAX TABLE"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   10980
   Begin TabDlg.SSTab SSTab1 
      Height          =   5145
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   630
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   9075
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483639
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
      TabCaption(0)   =   "Client Name"
      TabPicture(0)   =   "frmFixedRates.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblclient(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblclient(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblclient(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblclient(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblclient(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblclient(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblclient(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblclient(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblclient(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblclient(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblclient(11)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblclient(12)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblclient(13)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblclient(14)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblclient(10)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ListView1(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtClient(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtClient(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtClient(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtClient(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtClient(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtClient(5)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtClient(6)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtClient(7)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtClient(8)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkActiveClient"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtClient(9)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtClient(11)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtClient(12)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtClient(13)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtClient(14)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtClient(10)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "Configure Fixed Rates"
      TabPicture(1)   =   "frmFixedRates.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblclient(17)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblclient(15)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblclient(16)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblclient(18)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblclient(19)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblclient(20)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Line1(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblclient(21)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtClient(15)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmbCutoff"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame2(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtRates(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtRates(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtRates(2)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtRates(3)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmbClient"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      Begin VB.ComboBox cmbClient 
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "frmFixedRates.frx":0038
         Left            =   810
         List            =   "frmFixedRates.frx":0045
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   450
         Width           =   4005
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   7830
         TabIndex        =   57
         Text            =   "8"
         Top             =   990
         Width           =   465
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   6390
         TabIndex        =   56
         Text            =   "26"
         Top             =   990
         Width           =   465
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   4860
         TabIndex        =   55
         Text            =   "303"
         Top             =   990
         Width           =   465
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   3510
         TabIndex        =   54
         Text            =   "15"
         Top             =   990
         Width           =   465
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fixed Rates"
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
         Height          =   3615
         Index           =   0
         Left            =   90
         TabIndex        =   52
         Top             =   1350
         Width           =   8295
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   17
            Left            =   3960
            TabIndex        =   102
            Text            =   "Client Name"
            Top             =   2790
            Width           =   4185
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   18
            Left            =   3960
            TabIndex        =   100
            Text            =   "Client Name"
            Top             =   3150
            Width           =   4185
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   16
            Left            =   3960
            TabIndex        =   97
            Text            =   "Client Name"
            Top             =   2430
            Width           =   4185
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   16
            Left            =   7650
            TabIndex        =   95
            Text            =   "3.38"
            Top             =   900
            Width           =   465
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
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
            Height          =   3705
            Index           =   2
            Left            =   2700
            TabIndex        =   93
            Top             =   0
            Width           =   15
         End
         Begin VB.CheckBox chkRates 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Editable Tax Deductions"
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   8
            Left            =   180
            TabIndex        =   89
            Top             =   3150
            Width           =   2355
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ECOLA Conditions"
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
            Height          =   1005
            Index           =   1
            Left            =   2790
            TabIndex        =   88
            Top             =   1350
            Width           =   5415
            Begin VB.CheckBox chkRates 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Double ECOLA on Legal Holiday"
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
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   12
               Left            =   2610
               TabIndex        =   94
               Top             =   540
               Width           =   2625
            End
            Begin VB.CheckBox chkRates 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Double ECOLA on Spl. Holiday"
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
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   11
               Left            =   2610
               TabIndex        =   92
               Top             =   270
               Width           =   2625
            End
            Begin VB.CheckBox chkRates 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ECOLA with Late Deductions"
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
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   10
               Left            =   90
               TabIndex        =   91
               Top             =   540
               Width           =   2625
            End
            Begin VB.CheckBox chkRates 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Taxable ECOLA"
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
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   9
               Left            =   90
               TabIndex        =   90
               Top             =   270
               Width           =   2355
            End
         End
         Begin VB.CheckBox chkRates 
            BackColor       =   &H00C0C0C0&
            Caption         =   "W/ Tax Deductions"
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   180
            TabIndex        =   86
            Top             =   2790
            Width           =   1815
         End
         Begin VB.CheckBox chkRates 
            BackColor       =   &H00C0C0C0&
            Caption         =   "W/ PHealth Deductions"
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   6
            Left            =   180
            TabIndex        =   85
            Top             =   2430
            Width           =   2085
         End
         Begin VB.CheckBox chkRates 
            BackColor       =   &H00C0C0C0&
            Caption         =   "W/ SSS Deductions"
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   180
            TabIndex        =   84
            Top             =   2070
            Width           =   1815
         End
         Begin VB.CheckBox chkRates 
            BackColor       =   &H00C0C0C0&
            Caption         =   "W/ 13th Month Pay"
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   180
            TabIndex        =   83
            Top             =   1710
            Width           =   1815
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   5
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   15
            Left            =   1980
            TabIndex        =   81
            Text            =   "3.5%"
            Top             =   1260
            Width           =   645
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   5
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   14
            Left            =   1980
            TabIndex        =   79
            Text            =   "50.00"
            Top             =   900
            Width           =   645
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   13
            Left            =   5850
            TabIndex        =   75
            Text            =   "1.3"
            Top             =   900
            Width           =   465
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   12
            Left            =   7650
            TabIndex        =   73
            Text            =   "1"
            Top             =   540
            Width           =   465
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   11
            Left            =   5850
            TabIndex        =   71
            Text            =   "1.3"
            Top             =   540
            Width           =   465
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   10
            Left            =   4140
            TabIndex        =   69
            Text            =   "1.5"
            Top             =   540
            Width           =   465
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   5
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   9
            Left            =   1980
            TabIndex        =   68
            Text            =   "50.00"
            Top             =   540
            Width           =   645
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   8
            Left            =   4140
            TabIndex        =   66
            Text            =   "1.3"
            Top             =   900
            Width           =   465
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   7
            Left            =   7650
            TabIndex        =   64
            Text            =   "1"
            Top             =   180
            Width           =   465
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   6
            Left            =   5850
            TabIndex        =   62
            Text            =   "1.3"
            Top             =   180
            Width           =   465
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   5
            Left            =   4140
            TabIndex        =   59
            Text            =   "1.25"
            Top             =   180
            Width           =   465
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   5
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   4
            Left            =   1980
            TabIndex        =   58
            Text            =   "6.5%"
            Top             =   180
            Width           =   645
         End
         Begin VB.CheckBox chkRates 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Agency Rate"
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   77
            Top             =   270
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkRates 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ECOLA Payment"
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   78
            Top             =   630
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkRates 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pag-ibig Deductions"
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   80
            Top             =   990
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkRates 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fixed Tax Rate"
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   82
            Top             =   1350
            Width           =   1815
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Billing Schedule:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   2790
            TabIndex        =   101
            Top             =   3240
            Width           =   1185
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Release:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   2790
            TabIndex        =   99
            Top             =   2880
            Width           =   1185
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Pickup of DTR:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   2790
            TabIndex        =   98
            Top             =   2520
            Width           =   1185
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Sun/Lgl OT Rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   6390
            TabIndex        =   96
            Top             =   990
            Width           =   1365
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Legal Hol. Rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   4680
            TabIndex        =   76
            Top             =   990
            Width           =   1545
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "OT w/o Premium:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   6390
            TabIndex        =   74
            Top             =   630
            Width           =   1365
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Rest Day Rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   4680
            TabIndex        =   72
            Top             =   630
            Width           =   1185
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Night Differential:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   2790
            TabIndex        =   70
            Top             =   630
            Width           =   1545
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Special Hol. Rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   2790
            TabIndex        =   67
            Top             =   990
            Width           =   1365
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Tardiness Rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   6480
            TabIndex        =   65
            Top             =   270
            Width           =   1365
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Sunday Rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   4770
            TabIndex        =   63
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Regular OT:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3150
            TabIndex        =   61
            Top             =   270
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmbCutoff 
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "frmFixedRates.frx":0068
         Left            =   810
         List            =   "frmFixedRates.frx":0075
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   990
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   10
         Left            =   -69870
         TabIndex        =   45
         Text            =   "Client Name"
         Top             =   3780
         Width           =   3195
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   14
         Left            =   -74100
         TabIndex        =   42
         Text            =   "Client Name"
         Top             =   4410
         Width           =   3195
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   13
         Left            =   -69870
         TabIndex        =   41
         Text            =   "Client Name"
         Top             =   4410
         Width           =   3195
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   12
         Left            =   -74100
         TabIndex        =   38
         Text            =   "Client Name"
         Top             =   4095
         Width           =   3195
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   11
         Left            =   -69870
         TabIndex        =   37
         Text            =   "Client Name"
         Top             =   4095
         Width           =   3195
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   9
         Left            =   -74100
         TabIndex        =   35
         Text            =   "Client Name"
         Top             =   3780
         Width           =   3195
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   15
         Left            =   5670
         TabIndex        =   33
         Text            =   "Client Name"
         Top             =   450
         Width           =   2655
      End
      Begin VB.CheckBox chkActiveClient 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Active Client"
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -67980
         TabIndex        =   31
         Top             =   4770
         Value           =   1  'Checked
         Width           =   1365
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   8
         Left            =   -70860
         TabIndex        =   29
         Text            =   "Client Name"
         Top             =   3330
         Width           =   4185
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   7
         Left            =   -70860
         TabIndex        =   27
         Text            =   "Client Name"
         Top             =   3010
         Width           =   4185
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   -70860
         TabIndex        =   25
         Text            =   "Client Name"
         Top             =   2700
         Width           =   4185
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   -70860
         TabIndex        =   23
         Text            =   "Client Name"
         Top             =   2380
         Width           =   4185
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   4
         Left            =   -70860
         TabIndex        =   21
         Text            =   "Client Name"
         Top             =   2070
         Width           =   4185
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   -70860
         TabIndex        =   19
         Text            =   "Client Name"
         Top             =   1750
         Width           =   4185
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   -70860
         TabIndex        =   17
         Text            =   "Client Name"
         Top             =   1440
         Width           =   4185
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
         ForeColor       =   &H00FF0000&
         Height          =   675
         Index           =   1
         Left            =   -70860
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "frmFixedRates.frx":0098
         Top             =   760
         Width           =   4185
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   -70860
         TabIndex        =   13
         Text            =   "Client Name"
         Top             =   450
         Width           =   4185
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3120
         Index           =   0
         Left            =   -74820
         TabIndex        =   11
         Top             =   540
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   5503
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
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
         NumItems        =   0
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Day:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   6930
         TabIndex        =   53
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   4
         Index           =   0
         X1              =   0
         X2              =   8550
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Days/Month:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   5400
         TabIndex        =   51
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Days/Yr:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   4140
         TabIndex        =   50
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Days/CutOff:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   2520
         TabIndex        =   49
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Cut-Off:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   180
         TabIndex        =   48
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Acct No.:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -70770
         TabIndex        =   46
         Top             =   3870
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "HDMF:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -74820
         TabIndex        =   44
         Top             =   4470
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "TIN No.:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -70770
         TabIndex        =   43
         Top             =   4470
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "SSS No.:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -74820
         TabIndex        =   40
         Top             =   4155
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "P-Health:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -70770
         TabIndex        =   39
         Top             =   4155
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -74820
         TabIndex        =   36
         Top             =   3870
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   4950
         TabIndex        =   34
         Top             =   540
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   180
         TabIndex        =   32
         Top             =   540
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -71850
         TabIndex        =   30
         Top             =   3330
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Assistant:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -71850
         TabIndex        =   28
         Top             =   3010
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -71850
         TabIndex        =   26
         Top             =   2700
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -71850
         TabIndex        =   24
         Top             =   2380
         Width           =   915
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -71850
         TabIndex        =   22
         Top             =   2070
         Width           =   555
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "TeleFax:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -71850
         TabIndex        =   20
         Top             =   1750
         Width           =   825
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. #:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -71850
         TabIndex        =   18
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lblclient 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Left            =   -71850
         TabIndex        =   16
         Top             =   760
         Width           =   915
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
         Index           =   0
         Left            =   -71850
         TabIndex        =   14
         Top             =   450
         Width           =   915
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3705
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   6120
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   6535
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "HF"
      TabPicture(0)   =   "frmFixedRates.frx":00A4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "HF1"
      TabPicture(1)   =   "frmFixedRates.frx":00C0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin MSComctlLib.ListView ListView1 
         Height          =   3030
         Index           =   1
         Left            =   270
         TabIndex        =   12
         Top             =   450
         Width           =   7230
         _ExtentX        =   12753
         _ExtentY        =   5345
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
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
         NumItems        =   0
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
            Picture         =   "frmFixedRates.frx":00DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   794
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdedit 
         Height          =   495
         Left            =   630
         Picture         =   "frmFixedRates.frx":013A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   495
         Left            =   1800
         Picture         =   "frmFixedRates.frx":0284
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdsave 
         Height          =   495
         Left            =   1215
         Picture         =   "frmFixedRates.frx":07B6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdadd 
         Height          =   495
         Left            =   0
         Picture         =   "frmFixedRates.frx":0CE8
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "New"
         Top             =   0
         Width           =   585
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   495
         Left            =   2400
         Picture         =   "frmFixedRates.frx":121A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmddelete 
         Height          =   495
         Left            =   3015
         Picture         =   "frmFixedRates.frx":1364
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   600
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
      TabIndex        =   60
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
      TabIndex        =   2
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
Attribute VB_Name = "FrmFixedRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public i As Integer
'
Dim blSave As Boolean

Private Sub cmdrefresh_Click()
   Call loadlv(SSTab1(1).Caption & "taxtable", SSTab1(0).Caption)
End Sub

Private Sub Form_Load()
    Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Call SSTab1_Click(1, 1)
End Sub

Private Sub cmdadd_Click()
    blSave = True
    SSTab1(2).Visible = True
    Text(0).Enabled = True
    Text(1).Enabled = True
    
    Call clearText
    Text(1).Text = "99999.99"
End Sub

Private Sub sAVE()
    If blSave = True Then
          If CheckText = True Then
        If SSTab1(1).Tab <> 4 Then
            SQLstring = "Insert into " & SSTab1(1).Caption & "taxtable (TAXCODE,salrate2,basetax,perover) values ('" & SSTab1(0).Caption & "','" & Text(0).Text & "','" & Text(1).Text & "','" & Text(2).Text & "','" & Text(3).Text & "')"
        Else
            SQLstring = "Insert into " & SSTab1(1).Caption & "taxtable (salrate1,salrate2,basetax,perover) values ('" & Text(0).Text & "','" & Text(1).Text & "','" & Text(2).Text & "','" & Text(3).Text & "')"
        End If
        Else
        Exit Sub
        
         End If
    Else
        SQLstring = "Update " & SSTab1(1).Caption & "taxtable  set salrate2='" & Text(1).Text & "',basetax='" & Text(2).Text & "',perover='" & Text(3).Text & "' where id =" & ListView1.SelectedItem.Text & ""
    End If
    Call Action_Query(SQLstring, Cnnstr)
    SSTab1(2).Visible = False
    Call loadlv(SSTab1(1).Caption & "taxtable", SSTab1(0).Caption)
End Sub

Function CheckText() As Boolean
If Val(Text(0)) > Val(Text(1)) Then
MsgBox "Salary rate 1 must be less than salary rate 2"
CheckText = False
Else
CheckText = True

End If


End Function
Private Sub cmdCancel_Click()
    SSTab1(2).Visible = False
End Sub

Private Sub cmddelete_Click()
    Dim strID As Integer
    If MsgBox("Are you sure want to remove this item?", vbQuestion + vbYesNo) = vbYes Then
        strID = ListView1.ListItems(ListView1.SelectedItem.Index).Text
        Call Action_Query("Delete from " & SSTab1(1).Caption & "taxtable where id=" & strID & "", Cnnstr)
        Call loadlv(SSTab1(1).Caption & "taxtable", SSTab1(0).Caption)
    End If
End Sub

Private Sub cmdedit_Click()
    blSave = False
    'Text(0).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text
    Text(1).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text
    Text(2).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text
    Text(3).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(3).Text
    SSTab1(2).Visible = True
    Text(0).Enabled = False
    Text(1).Enabled = False
End Sub

Private Sub cmdsave_Click()
    Call sAVE
End Sub
Private Sub lblExit_Click()
    SSTab1(2).Visible = False
End Sub

Private Sub SSTab1_Click(Index As Integer, PreviousTab As Integer)
    Dim StrTab As String
    StrTab = SSTab1(0).Caption
    If SSTab1(1).Tab = 0 Then
       Call loadlv("dailytaxtable", StrTab)
    ElseIf SSTab1(1).Tab = 1 Then
       Call loadlv("weeklytaxtable", StrTab)
    ElseIf SSTab1(1).Tab = 2 Then
       Call loadlv("semimonthlytaxtable", StrTab)
    ElseIf SSTab1(1).Tab = 3 Then
       Call loadlv("monthlytaxtable", StrTab)
    ElseIf SSTab1(1).Tab = 4 Then
       Call loadlv("yearlytaxtable")
    End If
End Sub
Private Sub loadlv(ByVal TableName As String, Optional ByVal StrTab As String)
    Dim rsList As New ADODB.Recordset
    ListView1.ListItems.Clear
        If SSTab1(1).Tab = 4 Then
            'Set rsList = ReturnRst("Select salrate1 as [RATE FROM],salrate2 AS [RATE TO],basetax AS [BASE TAX], perover AS [PER OVER]  from " & tableName & " order by perover ", Cnnstr, mstatic)
            Set rsList = ReturnRst("Select ID, salrate2 AS [RATE],basetax AS [BASE TAX], perover AS [PER OVER] from " & TableName & " order by ID ", Cnnstr, mstatic)
            SSTab1(0).Visible = False
        Else
            'Set rsList = ReturnRst("Select id,basesalary,percentagetax,exemptionamt,bracket from " & tableName & "   where taxstatus='" & StrTab & "' order by id", Cnnstr, mstatic)
            Set rsList = ReturnRst("Select ID,salrate2 AS [RATE],basetax AS [BASE TAX], perover AS [PER OVER] from " & TableName & "   where taxcode='" & StrTab & "' order by ID ", Cnnstr, mstatic)
            SSTab1(0).Visible = True
        End If
        
        If rsList.EOF = False Then
            PopulateList ListView1, rsList
        End If
        ListView1.ColumnHeaders(1).Width = 0
        ListView1.ColumnHeaders(2).Width = 2000
        ListView1.ColumnHeaders(3).Width = 2000
       ' ListView1.ColumnHeaders(4).Width = 1600
        If SSTab1(1).Tab <> 3 Then
            ListView1.ColumnHeaders(4).Width = 2000
            ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
        End If
        ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
        ListView1.ColumnHeaders(3).Alignment = lvwColumnRight
       ' ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub clearText()
    Dim X As Integer
    For X = 0 To 3
    Text(X).Text = ""
    Next X
End Sub

