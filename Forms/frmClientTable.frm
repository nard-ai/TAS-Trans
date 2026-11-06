VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmClientTable 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Maintenance Table"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmClientTable.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdNavigation 
      Height          =   315
      Index           =   0
      Left            =   7440
      Picture         =   "frmClientTable.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   117
      ToolTipText     =   "First"
      Top             =   6390
      Width           =   300
   End
   Begin VB.CommandButton cmdNavigation 
      Height          =   315
      Index           =   1
      Left            =   7725
      Picture         =   "frmClientTable.frx":1374
      Style           =   1  'Graphical
      TabIndex        =   116
      ToolTipText     =   "First"
      Top             =   6390
      Width           =   300
   End
   Begin VB.CommandButton cmdNavigation 
      Height          =   315
      Index           =   2
      Left            =   8010
      Picture         =   "frmClientTable.frx":146E
      Style           =   1  'Graphical
      TabIndex        =   115
      ToolTipText     =   "First"
      Top             =   6390
      Width           =   300
   End
   Begin VB.CommandButton cmdNavigation 
      Height          =   315
      Index           =   3
      Left            =   8295
      Picture         =   "frmClientTable.frx":1568
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   6390
      Width           =   300
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6525
      Index           =   1
      Left            =   90
      TabIndex        =   59
      Top             =   540
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   11509
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
      TabPicture(0)   =   "frmClientTable.frx":1A9A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2(4)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Client Fixed Rates"
      TabPicture(1)   =   "frmClientTable.frx":1AB6
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
      Tab(1).Control(15)=   "txtClient(21)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
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
         Height          =   315
         Index           =   21
         Left            =   810
         TabIndex        =   112
         Text            =   "Client Name"
         Top             =   450
         Width           =   4095
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
         Height          =   5625
         Index           =   4
         Left            =   -74910
         TabIndex        =   96
         Top             =   360
         Width           =   8565
         Begin VB.ComboBox cmbDriver 
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
            ItemData        =   "frmClientTable.frx":1AD2
            Left            =   1560
            List            =   "frmClientTable.frx":1ADC
            TabIndex        =   137
            Text            =   "cmbDriver"
            Top             =   4920
            Width           =   2655
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
            Index           =   25
            Left            =   4320
            TabIndex        =   2
            Text            =   "Client Name"
            Top             =   270
            Width           =   4185
         End
         Begin VB.TextBox txtClient 
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
            Height          =   195
            Index           =   24
            Left            =   3090
            TabIndex        =   123
            Text            =   "Client Name"
            Top             =   5280
            Width           =   1395
         End
         Begin VB.TextBox txtClient 
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
            Height          =   195
            Index           =   23
            Left            =   5880
            TabIndex        =   122
            Text            =   "Client Name"
            Top             =   5280
            Width           =   1395
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   118
            Text            =   "Text1"
            Top             =   5490
            Visible         =   0   'False
            Width           =   495
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
            Index           =   0
            Left            =   4320
            TabIndex        =   3
            Text            =   "Client Name"
            Top             =   600
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
            ForeColor       =   &H00C00000&
            Height          =   675
            Index           =   1
            Left            =   4320
            MultiLine       =   -1  'True
            TabIndex        =   4
            Text            =   "frmClientTable.frx":1AF3
            Top             =   945
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   2
            Left            =   4320
            TabIndex        =   5
            Text            =   "Client Name"
            Top             =   1620
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   3
            Left            =   4320
            TabIndex        =   6
            Text            =   "Client Name"
            Top             =   1935
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   4
            Left            =   4320
            TabIndex        =   7
            Text            =   "Client Name"
            Top             =   2250
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   5
            Left            =   4320
            TabIndex        =   8
            Text            =   "Client Name"
            Top             =   2565
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   6
            Left            =   4320
            TabIndex        =   9
            Text            =   "Client Name"
            Top             =   2880
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   7
            Left            =   4320
            TabIndex        =   10
            Text            =   "Client Name"
            Top             =   3195
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   8
            Left            =   4320
            TabIndex        =   11
            Text            =   "Client Name"
            Top             =   3510
            Width           =   4185
         End
         Begin VB.CheckBox chkActiveClient 
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
            Left            =   5160
            TabIndex        =   18
            Top             =   4920
            Value           =   1  'Checked
            Width           =   1275
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
            Height          =   315
            Index           =   9
            Left            =   810
            TabIndex        =   12
            Text            =   "Client Name"
            Top             =   3960
            Width           =   3375
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
            Height          =   315
            Index           =   11
            Left            =   5130
            TabIndex        =   16
            Text            =   "Client Name"
            Top             =   4275
            Width           =   3375
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
            Height          =   315
            Index           =   12
            Left            =   810
            TabIndex        =   13
            Text            =   "Client Name"
            Top             =   4275
            Width           =   3375
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
            Height          =   315
            Index           =   13
            Left            =   5130
            TabIndex        =   17
            Text            =   "Client Name"
            Top             =   4590
            Width           =   3375
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
            Height          =   315
            Index           =   14
            Left            =   810
            TabIndex        =   14
            Text            =   "Client Name"
            Top             =   4590
            Width           =   3375
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
            Height          =   315
            Index           =   10
            Left            =   5130
            TabIndex        =   15
            Text            =   "Client Name"
            Top             =   3960
            Width           =   3375
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   3480
            Index           =   0
            Left            =   90
            TabIndex        =   0
            Top             =   270
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   6138
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Client Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Client ID"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Driver/Conductor:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   120
            TabIndex        =   138
            Top             =   5040
            Width           =   1395
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Client Group:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3330
            TabIndex        =   136
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Updated by:"
            BeginProperty Font 
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
            Index           =   40
            Left            =   1920
            TabIndex        =   125
            Top             =   5280
            Width           =   1395
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Update:"
            BeginProperty Font 
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
            Index           =   39
            Left            =   4620
            TabIndex        =   124
            Top             =   5280
            Width           =   1395
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
            Left            =   3330
            TabIndex        =   111
            Top             =   270
            Width           =   915
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
            Left            =   3330
            TabIndex        =   110
            Top             =   945
            Width           =   915
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
            Left            =   3330
            TabIndex        =   109
            Top             =   1620
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
            Left            =   3330
            TabIndex        =   108
            Top             =   1935
            Width           =   825
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
            Left            =   3330
            TabIndex        =   107
            Top             =   2250
            Width           =   555
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
            Left            =   3330
            TabIndex        =   106
            Top             =   2565
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
            Left            =   3330
            TabIndex        =   105
            Top             =   2880
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
            Left            =   3330
            TabIndex        =   104
            Top             =   3195
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
            Left            =   3330
            TabIndex        =   103
            Top             =   3510
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
            Index           =   9
            Left            =   90
            TabIndex        =   102
            Top             =   4050
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
            Left            =   4320
            TabIndex        =   101
            Top             =   4335
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
            Left            =   90
            TabIndex        =   100
            Top             =   4335
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
            Left            =   4320
            TabIndex        =   99
            Top             =   4650
            Width           =   915
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Zip Code:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   90
            TabIndex        =   98
            Top             =   4650
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
            Left            =   4320
            TabIndex        =   97
            Top             =   4050
            Width           =   915
         End
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
         Left            =   8010
         TabIndex        =   23
         Text            =   "8"
         Top             =   990
         Width           =   555
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
         Index           =   2
         Left            =   6480
         TabIndex        =   22
         Text            =   "26"
         Top             =   990
         Width           =   555
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
         Index           =   1
         Left            =   4860
         TabIndex        =   21
         Text            =   "303"
         Top             =   990
         Width           =   555
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
         Index           =   0
         Left            =   3510
         TabIndex        =   20
         Text            =   "15"
         Top             =   990
         Width           =   555
      End
      Begin VB.Frame Frame2 
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
         Height          =   4995
         Index           =   0
         Left            =   120
         TabIndex        =   75
         Top             =   1320
         Width           =   8565
         Begin VB.CheckBox chkRates 
            Caption         =   "Use Tax Table 2018"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Index           =   18
            Left            =   3480
            TabIndex        =   141
            Top             =   4200
            Width           =   1755
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
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
            Index           =   23
            Left            =   5160
            TabIndex        =   135
            Text            =   "0.00"
            Top             =   4605
            Width           =   645
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
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
            Index           =   22
            Left            =   2250
            TabIndex        =   134
            Text            =   "0.00"
            Top             =   4605
            Width           =   645
         End
         Begin VB.CheckBox chkRates 
            Caption         =   "Add Hazard Pay:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Index           =   16
            Left            =   180
            TabIndex        =   132
            Top             =   4515
            Width           =   1515
         End
         Begin VB.CheckBox chkRates 
            Caption         =   "Half SSS"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   15
            Left            =   2040
            TabIndex        =   131
            Top             =   2520
            Width           =   855
         End
         Begin VB.CheckBox chkRates 
            Caption         =   "OT is NON-Taxable"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Index           =   14
            Left            =   180
            TabIndex        =   129
            Top             =   3720
            Width           =   2715
         End
         Begin VB.CheckBox chkRates 
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
            TabIndex        =   127
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkRates 
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
            TabIndex        =   126
            Top             =   1870
            Value           =   1  'Checked
            Width           =   1815
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
            Index           =   22
            Left            =   4230
            TabIndex        =   53
            Text            =   "Client Name"
            Top             =   2520
            Width           =   4185
         End
         Begin VB.TextBox txtClient 
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
            Index           =   20
            Left            =   7020
            TabIndex        =   58
            Text            =   "Client Name"
            Top             =   3810
            Width           =   1395
         End
         Begin VB.TextBox txtClient 
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
            Index           =   19
            Left            =   4230
            TabIndex        =   57
            Text            =   "Client Name"
            Top             =   3810
            Width           =   1395
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
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
            Index           =   21
            Left            =   2250
            TabIndex        =   37
            Text            =   "15.54"
            Top             =   3240
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
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
            Index           =   20
            Left            =   2250
            TabIndex        =   30
            Text            =   "20.83"
            Top             =   1530
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.OptionButton optCharge2 
            Caption         =   "Client Charge per Day"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   180
            TabIndex        =   25
            Top             =   540
            Value           =   -1  'True
            Width           =   1995
         End
         Begin VB.OptionButton optCharge1 
            Caption         =   "Client Charge per Cutoff"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   270
            Width           =   1995
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
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
            Index           =   19
            Left            =   2250
            TabIndex        =   35
            Text            =   "15.54"
            Top             =   2880
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
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
            Index           =   18
            Left            =   2250
            TabIndex        =   33
            Text            =   "15.54"
            Top             =   2520
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
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
            Index           =   17
            Left            =   2250
            TabIndex        =   39
            Text            =   "0.00"
            Top             =   4245
            Width           =   645
         End
         Begin VB.CheckBox chkRates 
            Caption         =   "Exclude SSS, Pagibig...from Tax"
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
            Height          =   405
            Index           =   13
            Left            =   180
            TabIndex        =   38
            Top             =   3405
            Width           =   2715
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
            Index           =   17
            Left            =   4230
            TabIndex        =   55
            Text            =   "Client Name"
            Top             =   3170
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   18
            Left            =   4230
            TabIndex        =   56
            Text            =   "Client Name"
            Top             =   3480
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   16
            Left            =   4230
            TabIndex        =   54
            Text            =   "Client Name"
            Top             =   2850
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   16
            Left            =   8000
            TabIndex        =   48
            Text            =   "2.60"
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
            Height          =   4035
            Index           =   2
            Left            =   2970
            TabIndex        =   87
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame2 
            Caption         =   "ECOLA Options"
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
            Height          =   1185
            Index           =   1
            Left            =   3060
            TabIndex        =   86
            Top             =   1260
            Width           =   5415
            Begin VB.CheckBox chkRates 
               Caption         =   "ECOLA Payment on Holidays"
               BeginProperty Font 
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
               Left            =   2370
               TabIndex        =   120
               Top             =   810
               Width           =   2865
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "Other Income is Taxable"
               BeginProperty Font 
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
               TabIndex        =   119
               Top             =   810
               Width           =   2625
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "ECOLA Payment on Legal Holiday"
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
               Left            =   2370
               TabIndex        =   52
               Top             =   540
               Width           =   2865
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "ECOLA Payment on Spl. Holiday"
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
               Left            =   2370
               TabIndex        =   51
               Top             =   270
               Width           =   2865
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "Allowance is Taxable"
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
               TabIndex        =   50
               Top             =   540
               Width           =   2625
            End
            Begin VB.CheckBox chkRates 
               Caption         =   "ECOLA is Taxable"
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
               TabIndex        =   64
               Top             =   270
               Value           =   1  'Checked
               Width           =   2355
            End
         End
         Begin VB.CheckBox chkRates 
            Caption         =   "W/ PHealth Dedtns"
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
            TabIndex        =   34
            Top             =   2835
            Value           =   1  'Checked
            Width           =   2085
         End
         Begin VB.CheckBox chkRates 
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
            TabIndex        =   32
            Top             =   2520
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
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
            Index           =   15
            Left            =   2250
            TabIndex        =   31
            Text            =   "3.5"
            Top             =   2160
            Width           =   645
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
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
            Index           =   14
            Left            =   2250
            TabIndex        =   29
            Text            =   "50.00"
            Top             =   1170
            Visible         =   0   'False
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   13
            Left            =   6075
            TabIndex        =   47
            Text            =   "1.6"
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   12
            Left            =   8000
            TabIndex        =   45
            Text            =   "1.50"
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   11
            Left            =   6075
            TabIndex        =   44
            Text            =   "0.3"
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   10
            Left            =   4410
            TabIndex        =   43
            Text            =   "0.1"
            Top             =   540
            Width           =   465
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
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
            Index           =   9
            Left            =   2250
            TabIndex        =   27
            Text            =   "10.00"
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   8
            Left            =   4410
            TabIndex        =   46
            Text            =   "0.5"
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   7
            Left            =   8000
            TabIndex        =   42
            Text            =   "0.30"
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   6
            Left            =   6075
            TabIndex        =   41
            Text            =   "0.3"
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   5
            Left            =   4410
            TabIndex        =   40
            Text            =   "1.25"
            Top             =   180
            Width           =   465
         End
         Begin VB.TextBox txtRates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
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
            Index           =   4
            Left            =   2250
            TabIndex        =   26
            Text            =   "6.5"
            Top             =   180
            Width           =   645
         End
         Begin VB.CheckBox chkRates 
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
            TabIndex        =   28
            Top             =   1170
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkRates 
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
            TabIndex        =   36
            Top             =   3165
            Value           =   1  'Checked
            Width           =   1815
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
            Height          =   105
            Index           =   3
            Left            =   90
            TabIndex        =   92
            Top             =   810
            Width           =   2715
         End
         Begin VB.CheckBox chkRates 
            Caption         =   "No TAX if Basic Rate less"
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
            TabIndex        =   128
            Top             =   2190
            Width           =   2115
         End
         Begin VB.CheckBox chkRates 
            Caption         =   "Deduc Union Dues:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Index           =   17
            Left            =   3480
            TabIndex        =   133
            Top             =   4515
            Width           =   1755
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
            Height          =   105
            Index           =   6
            Left            =   120
            TabIndex        =   140
            Top             =   4080
            Width           =   8235
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "New PhilHealth Rate:"
            BeginProperty Font 
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
            Index           =   43
            Left            =   240
            TabIndex        =   139
            Top             =   4275
            Width           =   1755
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Cut-Off Period:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3060
            TabIndex        =   121
            Top             =   2565
            Width           =   1185
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Rate/ Day"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   2160
            TabIndex        =   113
            Top             =   950
            Width           =   735
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Update:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   5760
            TabIndex        =   94
            Top             =   3810
            Width           =   1185
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Updated by:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3060
            TabIndex        =   93
            Top             =   3810
            Width           =   1185
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
            Left            =   3060
            TabIndex        =   91
            Top             =   3500
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
            Left            =   3060
            TabIndex        =   90
            Top             =   3210
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
            Left            =   3060
            TabIndex        =   89
            Top             =   2890
            Width           =   1185
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "RD on LGL Rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   6705
            TabIndex        =   88
            Top             =   990
            Width           =   1245
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Legal Holiday:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   4950
            TabIndex        =   85
            Top             =   990
            Width           =   1545
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "RD on SPL Rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   6705
            TabIndex        =   84
            Top             =   630
            Width           =   1245
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Rest Day OT:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   4950
            TabIndex        =   83
            Top             =   630
            Width           =   1785
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
            Left            =   3060
            TabIndex        =   82
            Top             =   630
            Width           =   1545
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Special Holiday:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3060
            TabIndex        =   81
            Top             =   990
            Width           =   1365
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "RD Spl/Lgl Rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   6705
            TabIndex        =   80
            Top             =   270
            Width           =   1245
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
            Left            =   4950
            TabIndex        =   79
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Regular OT Rate:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3060
            TabIndex        =   78
            Top             =   270
            Width           =   1335
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
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "frmClientTable.frx":1AFF
         Left            =   810
         List            =   "frmClientTable.frx":1B0C
         Style           =   2  'Dropdown List
         TabIndex        =   19
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   15
         Left            =   5670
         TabIndex        =   69
         Text            =   "Client Name"
         Top             =   450
         Width           =   2835
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
         Left            =   7110
         TabIndex        =   76
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         Index           =   0
         X1              =   270
         X2              =   8460
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
         Left            =   5490
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
         Top             =   1080
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
         TabIndex        =   70
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
         TabIndex        =   68
         Top             =   540
         Width           =   915
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
            Picture         =   "frmClientTable.frx":1B2F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdReport 
         Height          =   405
         Left            =   4200
         Picture         =   "frmClientTable.frx":1B8D
         Style           =   1  'Graphical
         TabIndex        =   130
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdClose 
         Height          =   405
         Left            =   3600
         Picture         =   "frmClientTable.frx":1CD7
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   630
         Picture         =   "frmClientTable.frx":1DD1
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1800
         Picture         =   "frmClientTable.frx":1F1B
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1215
         Picture         =   "frmClientTable.frx":244D
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmClientTable.frx":297F
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "New"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   2400
         Picture         =   "frmClientTable.frx":2EB1
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   3015
         Picture         =   "frmClientTable.frx":2FFB
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
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
      TabIndex        =   77
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
      TabIndex        =   49
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
Attribute VB_Name = "frmClientTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public i As Integer
Dim rsClientData As ADODB.Recordset
Dim StrQuery As String
Dim mRecNo As Integer

Dim AddFlag As Boolean
Dim EditFlag As Boolean

Private Sub cmbCutoff_Click()
    If cmbCutoff.Text = "Weekly" Then
        txtRates(0) = 7
    ElseIf cmbCutoff.Text = "Semi-Monthly" Then
        txtRates(0) = 15
    Else
        txtRates(0) = 30
    End If
End Sub

Private Sub cmbCutoff_GotFocus()
    SendKeys "{HOME}+{END}"
    cmbCutoff.BackColor = &HFFFFC0
End Sub

Private Sub cmbCutoff_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbCutoff_LostFocus()
    cmbCutoff.BackColor = &HFFFFFF
End Sub

Private Sub cmdClose_Click()
Set rsClientData = Nothing
Unload Me
End Sub

Private Sub cmdNavigation_Click(Index As Integer)
Call Navigation(rsClientData, Me.cmdNavigation(Index), _
Me.cmdNavigation(0), Me.cmdNavigation(1), Me.cmdNavigation(2), Me.cmdNavigation(3))
Call DisplayClientData(ListView1(0).SelectedItem.Text, True)
Call ScrollRec
End Sub

Private Sub cmdRefresh_Click()
Call Form_Load
MsgBox "Data was updated!", vbInformation, PayrollAlert
End Sub

Private Sub cmdReport_Click()
Dim rsItems As New ADODB.Recordset
    
'    If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'        StrQuery = "Select * from clientTable where  CLIENTCODE2 = 'FMS' AND Clientcode = " & Text1
'    Else
'        StrQuery = "Select * from clientTable where  Clientcode = " & Text1
'    End If
    
    StrQuery = GetClientRecord(Text1)
    Set rsItems = ReturnRst(StrQuery, cnnstr, mstatic)

    'Set rsItems = ReturnRst("Select * from clientTable where Clientcode = " & Text1, cnnstr, mstatic)
        If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
        Dim tmpimg As RptImage
        Set tmpimg = rptReportClient.Sections(2).Controls("ImgLogo")
        On Error GoTo b:
        If LoadPictureFlag = True Then
            Set tmpimg.Picture = LoadPicture(logoPath)
        End If
b:
        With rptReportClient
        Set .DataSource = rsItems
             .DataMember = rsItems.DataMember
            .Sections("Section2").Controls("lblAddress").Caption = getAgencyAddress(CompanyCode) ' "Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
             .Show
        End With
    
    Set rsItems = Nothing
    
End Sub

Private Sub Form_Load()
    
If EditFlag = False Then
    Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
End If
    mUserID = UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID)))
    Call GetClientData
    If EditFlag = True Then
        rsClientData.Find " clientcode = " & mRecNo
    End If
    AddFlag = False
    EditFlag = False
    Call DisableTextBoxes(False)
    Call DisableCommands(False)
    Call DisplayClientData(ListView1(0).SelectedItem.Text, True)
End Sub

Private Sub GetClientData()
    
'If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'    StrQuery = "SELECT * from ClientTable WHERE CLIENTCODE2 = 'FMS' AND clientname is not null ORDER BY clientname"
'Else
'    StrQuery = "SELECT * from ClientTable WHERE clientname is not null ORDER BY clientname"
'End If

StrQuery = GetClientRecord(0, False, " ORDER BY ClientName")

Set rsClientData = ReturnRst(StrQuery, cnnstr, mstatic)

If rsClientData.RecordCount <> 0 Then
    Call FillClientList(ListView1(0), rsClientData)
    rsClientData.MoveFirst
Else
    ' No client selected disable text boxes and command buttons
    Call DisableCommands(True)
    Call DisableTextBoxes(False)
End If

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
If mSw = True Then ' no client records selected, enable only add and exit button
    cmdedit.Enabled = False
    cmdrefresh.Enabled = False
    cmddelete.Enabled = False
End If

End Sub

Private Sub DisableTextBoxes(ByVal TFalse As Boolean)
    
    For i = 0 To 22
        txtClient(i).Enabled = TFalse
    Next i
    txtClient(25).Enabled = TFalse
    cmbDriver.Enabled = TFalse
    For i = 0 To 23  ' 21
        txtRates(i).Enabled = TFalse
    Next i
    For i = 1 To 7
        chkRates(i).Enabled = True
    Next i
    For i = 9 To 13
        chkRates(i).Enabled = True
    Next i
    chkRates(15).Enabled = True
    If AddFlag = True Or EditFlag = True Then
        ListView1(0).Enabled = False
        cmdNavigation(0).Enabled = False
        cmdNavigation(1).Enabled = False
        cmdNavigation(2).Enabled = False
        cmdNavigation(3).Enabled = False
    Else
        ListView1(0).Enabled = True
        cmdNavigation(0).Enabled = True
        cmdNavigation(1).Enabled = True
        cmdNavigation(2).Enabled = True
        cmdNavigation(3).Enabled = True
    End If
    cmbCutoff.Enabled = TFalse
    txtClient(19).Enabled = False
    txtClient(20).Enabled = False
    txtClient(23).Enabled = False
    txtClient(24).Enabled = False
    txtClient(15).Enabled = False
    txtClient(21).Enabled = False
    chkActiveClient.Enabled = True
    optCharge1.Enabled = True
    optCharge2.Enabled = True
End Sub

'Private Sub FillClientList(ByRef mLsEmp As ListView, _
'                        mRs As Recordset)
'Dim litem1 As ListItem
'
'    mLsEmp.View = lvwReport
'    mLsEmp.ListItems.Clear
'    While Not (mRs.EOF Or mRs.BOF)
'        Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!ClientName), "Unknown Client", Replace(mRs!ClientName, "'", "''")))
'        litem1.SubItems(1) = mRs(0)
'        mRs.MoveNext
'    Wend
'End Sub

Private Sub cmdAdd_Click()
    AddFlag = True
    Call ClearTextBox
    Call DisableTextBoxes(True)
    Call DisableCommands(False)
    chkActiveClient = 1
    SSTab1(1).TabVisible(1) = True
    txtClient(25).SetFocus
End Sub

Private Sub save()
Dim mstr As String
Dim mDetails As String
Dim mRemarks As String
Dim mDrvFlag As Integer
Dim mFMSFlag As String

    For i = 0 To 22
        txtClient(i) = Replace(txtClient(i), "'", "''")
    Next i
    For i = 0 To 21
        If txtRates(i) = "" Then
            txtRates(i) = 0#
        End If
    Next i
    txtClient(19) = mUserID
    txtClient(24) = mUserID
    
    mDrvFlag = 0: mFMSFlag = "FMS"
    If Trim$(cmbDriver.Text) = "DRIVER" Then
        mDrvFlag = 1
    ElseIf Trim$(cmbDriver.Text) = "" Then
        mFMSFlag = ""
    End If
    ' Changes SundayRate changed to RestDayRate April 23, 2009.
    If AddFlag = True Then ' Add New Record
        StrQuery = "INSERT INTO ClientTable ( DriverFlag, DriverConductor, ClientCode2, ClientName, ClientReportName, ClientAddress, ClientTelNo, ClientTeleFax, ClientOwner, ClientContactPerson, " & _
        "ClientContactPosition, ClientAssistant, ClientAssistantPosition, ClientBankName, ClientBankAccountNo, ClientPhilHealth, ClientSSSNo, " & _
        "ClientTaxNumber, ClientPagibig, ClientCutoff, ClientCutoffDates, WorkingDaysPerYear, NoofDaysperMonth, WorkHrsPerDay , OverTimeRate, RestDayRate, " & _
        "TardinessRate, NDiffRate, SundayRate, OTwoPremiumRate, SpecialHolidayRate, LegalHolidayRate, LateChargesinMin, HazardPayFlag, HazardPayRate, UnionDuesFlag, EcolaPayonRestDay, UnionDuesRate, " & _
        "ClientAgencyFlag, ClientAgencyRate, ClientChargeperDay, ClientEcolaFlag, ECOLARate, With13MonthPay, Month13RatePerDay, ClientTaxFlag, ClientTaxPercentFlag, " & _
        "ClientTaxPercentage, ClientSSSFlag, ClientSSSRate, ClientPhilHealthFlag, ClientPHRateperDay, ClientPagibigflag, ClientPagibigRate, WithVatCharges, ClientFlag3, ClientFlag4, VatPercentRate, " & _
        "ClientEcolaTaxFlag, EcolawLateDeduction, EcolaDoubleonLegal, EcolaDoubleonSpl, EcolaPayonHolidays, CutoffPeriod, PickupofDTR, SalaryRelease, BillingSchedule, Updatedby, ClientStatus ) " & _
        " VALUES ( " & mDrvFlag & ", '" & Trim$(cmbDriver.Text) & "', '" & mFMSFlag & "', '" & txtClient(0) & "', '" & txtClient(25) & "', '" & txtClient(1) & "', '" & txtClient(2) & "', '" & txtClient(3) & "', '" & txtClient(4) & "', '" & txtClient(5) & "', '" & _
        txtClient(6) & "', '" & txtClient(7) & "', '" & txtClient(8) & "', '" & txtClient(9) & "', '" & txtClient(10) & "', '" & txtClient(11) & "', '" & txtClient(12) & "', '" & _
        txtClient(13) & "', '" & txtClient(14) & "', '" & cmbCutoff.Text & "', " & txtRates(0) & ", " & txtRates(1) & ", " & txtRates(2) & ", " & txtRates(3) & ", " & txtRates(5) & ", " & txtRates(6) & ", " & _
        txtRates(7) & ", " & txtRates(10) & ", " & txtRates(11) & ", " & txtRates(12) & ", " & txtRates(8) & ", " & txtRates(13) & ", " & txtRates(16) & ", " & chkRates(16) & ", " & txtRates(22) & ", " & chkRates(17) & ", " & chkRates(18) & ", " & txtRates(23) & ", " & _
        IIf(optCharge1 = True, 1, 0) & ", " & txtRates(4) & ", " & txtRates(9) & ", " & chkRates(1) & ", " & txtRates(14) & ", " & chkRates(4) & ", " & txtRates(20) & ", " & chkRates(7) & ", " & chkRates(3) & ", " & _
        txtRates(15) & ", " & chkRates(5) & ", " & txtRates(18) & ", " & chkRates(6) & ", " & txtRates(19) & ", " & chkRates(2) & ", " & txtRates(21) & ", " & chkRates(13) & ", " & chkRates(14) & ", " & chkRates(15) & ", " & txtRates(17) & ", " & _
        chkRates(9) & ", " & chkRates(10) & ", " & chkRates(12) & ", " & chkRates(11) & ", " & chkRates(8) & ", '" & txtClient(22) & "', '" & txtClient(16) & "', '" & txtClient(17) & "', '" & txtClient(18) & "', '" & txtClient(19) & "', " & chkActiveClient & ")"
        mstr = "New record was saved!"
        mDetails = "New client " & txtClient(25) & " was inserted."
        mRemarks = "New Client."
    ElseIf EditFlag = True Then ' save edit
        'SQLstring = "Update " & SSTab1(1).Caption & "taxtable  set salrate2='" & Text(1).Text & "',basetax='" & Text(2).Text & "',perover='" & Text(3).Text & "' where id =" & ListView1.SelectedItem.Text & ""
        StrQuery = "UPDATE ClientTable SET DriverFlag = " & mDrvFlag & ", DriverConductor = '" & Trim$(cmbDriver.Text) & "', ClientCode2 = '" & mFMSFlag & "', ClientName = '" & txtClient(0) & "', ClientReportName = '" & Trim(txtClient(25)) & "', ClientAddress = '" & Trim(txtClient(1)) & "', ClientTelNo = '" & txtClient(2) & "', ClientTeleFax = '" & txtClient(3) & "', ClientOwner = '" & txtClient(4) & "', ClientContactPerson = '" & txtClient(5) & "', " & _
        "ClientContactPosition = '" & txtClient(6) & "', ClientAssistant = '" & txtClient(7) & "', ClientAssistantPosition = '" & txtClient(8) & " ', ClientBankName = '" & txtClient(9) & "', ClientBankAccountNo = '" & txtClient(10) & "', ClientPhilHealth = '" & txtClient(11) & "', ClientSSSNo = '" & txtClient(12) & "', " & _
        "ClientTaxNumber = '" & txtClient(13) & "', ClientPagibig = '" & txtClient(14) & "', ClientCutoff = '" & cmbCutoff.Text & "', ClientCutoffDates = " & txtRates(0) & ", WorkingDaysPerYear = " & txtRates(1) & ", NoofDaysperMonth = " & txtRates(2) & ", WorkHrsPerDay = " & txtRates(3) & ", OverTimeRate = " & txtRates(5) & ", SundayRate = " & txtRates(6) & ", " & _
        "TardinessRate = " & txtRates(7) & ", NDiffRate = " & txtRates(10) & ", RestDayRate = " & txtRates(11) & ", OTwoPremiumRate = " & txtRates(12) & ", SpecialHolidayRate = " & txtRates(8) & ", LegalHolidayRate = " & txtRates(13) & ", LateChargesinMin = " & txtRates(16) & ", " & _
        "ClientAgencyFlag = " & IIf(optCharge1 = True, 1, 0) & " , ClientAgencyRate = " & txtRates(4) & ", ClientChargeperDay = " & txtRates(9) & ", ClientEcolaFlag = " & chkRates(1) & " , ECOLARate = " & txtRates(14) & ", With13MonthPay = " & chkRates(4) & ", Month13RatePerDay = " & txtRates(20) & ", ClientTaxFlag = " & chkRates(7) & ", ClientTaxPercentFlag = " & chkRates(3) & ", " & _
        "ClientTaxPercentage = " & txtRates(15) & ", ClientSSSFlag = " & chkRates(5) & ", ClientSSSRate = " & txtRates(18) & ", ClientPhilHealthFlag = " & chkRates(6) & ", ClientPHRateperDay = " & txtRates(19) & ", ClientPagibigflag = " & chkRates(2) & ", ClientPagibigRate = " & txtRates(21) & ", WithVatCharges = " & chkRates(13) & ", ClientFlag3 = " & chkRates(14) & ", ClientFlag4 = " & chkRates(15) & ", VatPercentRate = " & txtRates(17) & ", " & _
        "HazardPayFlag = " & chkRates(16) & ", HazardPayRate = " & txtRates(22) & ", UnionDuesFlag = " & chkRates(17) & ", EcolaPayonRestDay = " & chkRates(18) & ", UnionDuesRate = " & txtRates(23) & ", " & _
        "ClientEcolaTaxFlag = " & chkRates(9) & ", EcolawLateDeduction = " & chkRates(10) & " , EcolaDoubleonLegal = " & chkRates(12) & ", EcolaDoubleonSpl = " & chkRates(11) & ", EcolaPayonHolidays = " & chkRates(8) & ", CutoffPeriod = '" & txtClient(22) & "', PickupofDTR = '" & txtClient(16) & "', SalaryRelease = '" & txtClient(17) & "', BillingSchedule = '" & txtClient(18) & "', Updatedby = '" & txtClient(19) & "', DateUpdated = '" & Format(Now, "mm/dd/yyyy") & "', ClientStatus = " & chkActiveClient & " WHERE clientcode = " & Text1
        mstr = "Edit was saved!"
        mDetails = "Client " & txtClient(25) & " was updated."
        mRemarks = "Update Client."
    End If
    'Commented by Bernard 5/19/15, employee info was invoked also
    If Action_Query(StrQuery, cnnstr) = True Then
        Call SaveUserTransactions(mDetails, mRemarks, 999, "000", txtClient(21))
    End If
'    If Action_Query(strQuery, cnnStr) = True Then
'        Call Action_Query(strQuery, CnnstrHR)
'    Else
'        MsgBox "Client table was not successfully updated. Please repeat process.", vbInformation, PayrollAlert
'    End If
    MsgBox mstr, vbInformation, PayrollAlert
    Call Form_Load
    Call ScrollRec
End Sub

'Function CheckText() As Boolean
'If Val(Text(0)) > Val(Text(1)) Then
'MsgBox "Salary rate 1 must be less than salary rate 2"
'CheckText = False
'Else
'CheckText = True
'
'End If
'
'
'End Function

Private Sub cmdCancel_Click()
    'SSTab1(2).TabEnabled = True
    AddFlag = False
    EditFlag = False
    Call DisableTextBoxes(False)
    Call DisableCommands(False)
End Sub

Private Sub cmdDelete_Click()
    Dim rsemp As ADODB.Recordset
    
    Set rsemp = New ADODB.Recordset
    If DeleteRights = False Then
        MsgBox "You do not have rights to Delete Client Record!", vbInformation, PayrollAlert
        Exit Sub
    End If
    
    If MsgBox("Delete client " & txtClient(25) & "?", vbQuestion + vbYesNo) = vbYes Then
        'Check first whether there are employees under this client.
        StrQuery = "Select ClientCode from EmployeeBasicSalary where clientcode = " & Text1
        Set rsemp = ReturnRst(StrQuery, cnnstr)
        If rsemp.EOF = False Then
            MsgBox "There were employees connected to this client. You need to delete first all employees connected to this client before deletion.", vbInformation, PayrollAlert
        Else
            Call Action_Query("DELETE FROM clienttable WHERE clientcode = " & Text1, cnnstr)
            Call Form_Load
        End If
        Set rsemp = Nothing
    End If
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    mRecNo = Text1
    Call DisableTextBoxes(True)
    Call DisableCommands(False)
    'txtClient(0).SetFocus
    txtClient(25).SetFocus
    
'    Text(1).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text
'    Text(2).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text
'    Text(3).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(3).Text
'    SSTab1(2).Visible = True
End Sub

Private Sub cmdSave_Click()
    Call save
End Sub

Private Sub lblExit_Click()
    SSTab1(2).Visible = False
End Sub

Private Sub ListView1_Click(Index As Integer)
    Call DisplayClientData(ListView1(0).SelectedItem.Text, False)
End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Call DisplayClientData(Item.Text, False)
End Sub

'Private Sub SSTab1_Click(Index As Integer, PreviousTab As Integer)
'    Dim StrTab As String
'    StrTab = SSTab1(0).Caption
'    If SSTab1(1).Tab = 0 Then
'       Call loadlv("dailytaxtable", StrTab)
'    ElseIf SSTab1(1).Tab = 1 Then
'       Call loadlv("weeklytaxtable", StrTab)
'    ElseIf SSTab1(1).Tab = 2 Then
'       Call loadlv("semimonthlytaxtable", StrTab)
'    ElseIf SSTab1(1).Tab = 3 Then
'       Call loadlv("monthlytaxtable", StrTab)
'    ElseIf SSTab1(1).Tab = 4 Then
'       Call loadlv("yearlytaxtable")
'    End If
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub ClearTextBox()
    Dim X As Integer
    
    For X = 0 To 21
        txtClient(X).Text = ""
    Next X
    txtClient(25).Text = ""
    For X = 0 To 23
        txtRates(X).Text = 0
    Next X
    For X = 1 To 7
        chkRates(X).VALUE = 0
    Next X
    For X = 9 To 13
        chkRates(X).Enabled = True
    Next X
    chkRates(4).VALUE = 1
    ' Set the default value.
    txtRates(4) = 6.5
    txtRates(9) = 10#
    txtRates(14) = 50
    txtRates(5) = 1.25
    txtRates(6) = 1.3
    txtRates(7) = 0.3
    txtRates(10) = 1.5
    txtRates(11) = 1.3
    txtRates(12) = 1.5
    txtRates(8) = 1.3
    txtRates(13) = 1.3
    txtRates(16) = 2.6
    cmbCutoff.Text = "Semi-Monthly"
    txtRates(0) = 15
    txtRates(1) = 303
    txtRates(2) = 26
    txtRates(3) = 8
    chkActiveClient.Enabled = True
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
        txtClient(15) = txtClient(5)
    End If
    txtClient(15) = Replace(txtClient(15), "'", "''")
End Sub

Private Sub DisplayClientData(ByVal SelectedClient As String, Optional recSw As Boolean)

With rsClientData
    If recSw = False Then
        .MoveFirst
        .Find "Clientname = '" & SelectedClient & "'"
    End If
    If Not .EOF And Not .BOF Then
        Text1 = !ClientCode
        txtClient(0) = IIf(IsNull(!ClientName), "", !ClientName)
        cmbDriver.Text = IIf(IsNull(!DriverConductor), "", !DriverConductor)
        txtClient(25) = IIf(IsNull(!ClientReportName), "", !ClientReportName)
        txtClient(1) = IIf(IsNull(!ClientAddress), "", !ClientAddress)
        txtClient(2) = IIf(IsNull(!ClientTelNo), "", !ClientTelNo)
        txtClient(3) = IIf(IsNull(!ClientTeleFax), "", !ClientTeleFax)
        txtClient(4) = IIf(IsNull(!ClientOwner), "", !ClientOwner)
        txtClient(5) = IIf(IsNull(!ClientContactPerson), "", !ClientContactPerson)
        txtClient(6) = IIf(IsNull(!ClientContactPosition), "", !ClientContactPosition)
        txtClient(7) = IIf(IsNull(!ClientAssistant), "", !ClientAssistant)
        txtClient(8) = IIf(IsNull(!ClientAssistantPosition), "", !ClientAssistantPosition)
        txtClient(9) = IIf(IsNull(!ClientBankName), "", !ClientBankName)
        txtClient(10) = IIf(IsNull(!ClientBankAccountNo), "", !ClientBankAccountNo)
        txtClient(11) = IIf(IsNull(!ClientPhilHealth), "", !ClientPhilHealth)
        txtClient(12) = IIf(IsNull(!ClientSSSNo), "", !ClientSSSNo)
        txtClient(13) = IIf(IsNull(!ClientTaxNumber), "", !ClientTaxNumber)
        txtClient(14) = IIf(IsNull(!ClientPagibig), "", !ClientPagibig)
        ' Tab 2
        txtClient(21) = IIf(IsNull(!ClientName), "", !ClientName)
        txtClient(15) = IIf(IsNull(!ClientContactPerson), "", !ClientContactPerson)
        cmbCutoff.Text = IIf(IsNull(!ClientCutoff), "", !ClientCutoff)
        txtRates(0) = IIf(!ClientCutoffDates = "", 0#, !ClientCutoffDates)
        txtRates(1) = IIf(!WorkingDaysperYear = "", 0#, !WorkingDaysperYear)
        txtRates(2) = IIf(!NoofDaysperMonth = "", 0#, !NoofDaysperMonth)
        txtRates(3) = IIf(!WorkHrsPerDay = "", 0#, !WorkHrsPerDay)
        ' Overtime Charges
        txtRates(5) = IIf(!OverTimeRate = "", 0#, Format(!OverTimeRate, "0.00"))
        txtRates(6) = IIf(!SundayRate = "", 0#, Format(!SundayRate, "0.00"))
        txtRates(7) = IIf(!TardinessRate = "", 0#, Format(!TardinessRate, "0.00"))
        txtRates(10) = IIf(!NDiffRate = "", 0#, Format(!NDiffRate, "0.00"))
        txtRates(11) = IIf(!RestDayRate = "", 0#, Format(!RestDayRate, "0.00"))
        txtRates(12) = IIf(!OTwoPremiumRate = "", 0#, Format(!OTwoPremiumRate, "0.00"))
        txtRates(8) = IIf(!SpecialHolidayRate = "", 0#, Format(!SpecialHolidayRate, "0.00"))
        txtRates(13) = IIf(!LegalHolidayRate = "", 0#, Format(!LegalHolidayRate, "0.00"))
        txtRates(16) = IIf(!LateChargesinMin = "", 0#, Format(!LateChargesinMin, "0.00"))
        ' Agency Charges
        If !ClientAgencyFlag = True Then    ' if charges is per cutoff
            optCharge1 = True
        Else
            optCharge2 = True
        End If
        txtRates(4) = IIf(!ClientAgencyRate = "", 0#, Format(!ClientAgencyRate, "0.00"))
        txtRates(9) = IIf(!clientchargeperday = "", 0#, Format(!clientchargeperday, "0.00"))
        chkRates(1) = IIf(!ClientEcolaFlag = True, 1, 0)
        txtRates(14) = IIf(!ecolarate = "", 0#, Format(!ecolarate, "0.00"))
        chkRates(4) = IIf(!With13MonthPay = True, 1, 0)
        txtRates(20) = IIf(!Month13RateperDay = "", 0#, Format(!Month13RateperDay, "0.00"))
        ' With Tax
        chkRates(7) = IIf(!ClientTaxFlag = True, 1, 0)
        chkRates(3) = IIf(!ClientTaxPercentFlag = True, 1, 0)
        txtRates(15) = IIf(!ClientTaxPercentage = "", 0#, Format(!ClientTaxPercentage, "0.00"))
        chkRates(5) = IIf(!ClientSSSFlag = True, 1, 0)
        txtRates(18) = IIf(!ClientSSSRate = "", 0#, Format(!ClientSSSRate, "0.00"))
        chkRates(6) = IIf(!ClientPhilHealthFlag = True, 1, 0)
        txtRates(19) = IIf(!ClientPHRateperDay = "", 0#, Format(!ClientPHRateperDay, "0.00"))
        chkRates(2) = IIf(!ClientPagibigFlag = True, 1, 0)
        txtRates(21) = IIf(!ClientPagibigRate = "", 0#, Format(!ClientPagibigRate, "0.00"))
        chkRates(13) = IIf(!withvatcharges = True, 1, 0)    ' use in excluding sss, pagibig, phealth from tax computation
        chkRates(14) = IIf(!ClientFlag3 = True, 1, 0)    ' use as flag for OT non-taxable, updated Jan 17, 2015
        chkRates(15) = IIf(!ClientFlag4 = True, 1, 0)    ' use for half sss,philhealth and pagibig computation
        txtRates(17) = IIf(!VatPercentRate = "", 0#, Format(!VatPercentRate, "0.00"))
        ' Ecola Options
        chkRates(9) = IIf(!ClientEcolaTaxFlag = True, 1, 0)
        chkRates(10) = IIf(!EcolawLateDeduction = True, 1, 0)
        chkRates(12) = IIf(!EcolaDoubleonLegal = True, 1, 0)
        chkRates(11) = IIf(!EcolaDoubleonSpl = True, 1, 0)
        'chkRates(0) = IIf(!EcolaPayonRestDay = True, 1, 0)
        chkRates(8) = IIf(!EcolaPayOnHolidays = True, 1, 0)
        'SSS/PH Add on
        chkRates(16) = IIf(!HazardPayFlag = True, 1, 0)
        chkRates(17) = IIf(!UnionDuesFlag = True, 1, 0)
        'Updated Jan. 6, 2018 for Tax table 2018.
        chkRates(18) = IIf(!EcolaPayonRestDay = True, 1, 0)
        
        txtRates(22) = IIf(!HazardPayRate = "" Or IsNull(!HazardPayRate), 0#, Format(!HazardPayRate, "0.00"))
        txtRates(23) = IIf(!UnionDuesRate = "" Or IsNull(!UnionDuesRate), 0#, Format(!UnionDuesRate, "0.00"))
        
        txtClient(22) = IIf(IsNull(!CutOffPeriod), "", !CutOffPeriod)
        txtClient(16) = IIf(IsNull(!PickupofDTR), "", !PickupofDTR)
        txtClient(17) = IIf(IsNull(!SalaryRelease), "", !SalaryRelease)
        txtClient(18) = IIf(IsNull(!BillingSchedule), "", !BillingSchedule)
        
        txtClient(19) = IIf(IsNull(!Updatedby), "", !Updatedby)
        txtClient(20) = IIf(IsNull(!DateUpdated), "", Format(!DateUpdated, "mmm dd, yyyy"))
        txtClient(24) = IIf(IsNull(!Updatedby), "", !Updatedby)
        txtClient(23) = IIf(IsNull(!DateUpdated), "", Format(!DateUpdated, "mmm dd, yyyy"))
        
        chkActiveClient = IIf(!ClientStatus = True, 1, 0)
    End If
End With

End Sub

Private Sub txtClient_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        If txtClient(0) = "" Then MsgBox "Client Group Name should not be blank!", vbInformation, PayrollAlert: Cancel = True
    ElseIf Index = 25 Then
        If txtClient(25) = "" Then MsgBox "Client Name should not be blank!", vbInformation, PayrollAlert: Cancel = True
    End If
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
    txtRates(Index).Text = Format(txtRates(Index).Text, "##,#00.00")
End Sub

Private Sub chkRates_GotFocus(Index As Integer)
'    SendKeys "{HOME}+{END}"
'    chkRates(Index).BackColor = &HFFFFC0
End Sub

Private Sub ScrollRec()
Dim strFindMe As String
Dim itmFound As ListItem
   
    strFindMe = txtClient(0)
    If SSTab1(1).Tab = 0 Then
'        Set itmFound = ListView1(1).FindItem(strFindMe, 2, , lvwPartial)
 '   Else
        Set itmFound = ListView1(0).FindItem(strFindMe, 0, , lvwPartial)
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

