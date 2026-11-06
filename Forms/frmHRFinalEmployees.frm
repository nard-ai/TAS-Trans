VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmHRFinalEmployees 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Information"
   ClientHeight    =   6840
   ClientLeft      =   4455
   ClientTop       =   3885
   ClientWidth     =   10665
   Icon            =   "frmHRFinalEmployees.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   10665
   Begin TabDlg.SSTab SSTab1 
      Height          =   4125
      Left            =   135
      TabIndex        =   6
      Top             =   2160
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   7276
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483638
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Personal Data"
      TabPicture(0)   =   "frmHRFinalEmployees.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Contract"
      TabPicture(1)   =   "frmHRFinalEmployees.frx":05A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame16"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame15"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame12"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame14"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Family"
      TabPicture(2)   =   "frmHRFinalEmployees.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(2)=   "Frame1"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Educational"
      TabPicture(3)   =   "frmHRFinalEmployees.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame10"
      Tab(3).Control(1)=   "Frame8"
      Tab(3).Control(2)=   "Frame5(1)"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "References"
      TabPicture(4)   =   "frmHRFinalEmployees.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame9"
      Tab(4).Control(1)=   "Frame11"
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame14 
         Height          =   870
         Left            =   135
         TabIndex        =   285
         Top             =   315
         Width           =   5055
         Begin VB.TextBox txtreasonresign 
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
            Left            =   1890
            MaxLength       =   40
            TabIndex        =   39
            Top             =   495
            Width           =   2640
         End
         Begin VB.ComboBox cmbEmpStatus 
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
            Height          =   315
            ItemData        =   "frmHRFinalEmployees.frx":0616
            Left            =   1920
            List            =   "frmHRFinalEmployees.frx":0618
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   135
            Width           =   2640
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   64
            Left            =   180
            TabIndex        =   306
            Top             =   540
            Width           =   630
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Employee Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   60
            Left            =   180
            TabIndex        =   286
            Top             =   180
            Width           =   1200
         End
      End
      Begin VB.Frame Frame12 
         Height          =   1995
         Left            =   135
         TabIndex        =   255
         Top             =   1125
         Width           =   5055
         Begin VB.ComboBox cmbBranch 
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
            Height          =   315
            ItemData        =   "frmHRFinalEmployees.frx":061A
            Left            =   1890
            List            =   "frmHRFinalEmployees.frx":061C
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   945
            Width           =   2685
         End
         Begin VB.ComboBox cmbPosition 
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
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   225
            Width           =   2685
         End
         Begin VB.CommandButton cmdAddPosition 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   285
            Left            =   4635
            TabIndex        =   316
            Top             =   240
            Width           =   330
         End
         Begin VB.CommandButton cmdAddDept 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   285
            Left            =   4635
            TabIndex        =   274
            Top             =   1305
            Width           =   330
         End
         Begin VB.CommandButton cmdAddBranch 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   285
            Left            =   4635
            TabIndex        =   273
            Top             =   945
            Width           =   330
         End
         Begin VB.ComboBox cmbDept 
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
            Height          =   315
            ItemData        =   "frmHRFinalEmployees.frx":061E
            Left            =   1890
            List            =   "frmHRFinalEmployees.frx":0620
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1305
            Width           =   2685
         End
         Begin VB.TextBox txtsupervisor 
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
            Left            =   1890
            MaxLength       =   40
            TabIndex        =   45
            Top             =   1665
            Width           =   2685
         End
         Begin VB.ComboBox cmbCompany 
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
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   585
            Width           =   2685
         End
         Begin VB.TextBox txtfirstchoice 
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
            Left            =   1890
            MaxLength       =   25
            TabIndex        =   41
            Top             =   225
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Department"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   92
            Left            =   180
            TabIndex        =   271
            Top             =   1395
            Width           =   825
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Work Station"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   72
            Left            =   180
            TabIndex        =   270
            Top             =   1035
            Width           =   915
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Supervisor (Authority)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   63
            Left            =   180
            TabIndex        =   258
            Top             =   1710
            Width           =   1620
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Company Connected"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   96
            Left            =   180
            TabIndex        =   257
            Top             =   675
            Width           =   1500
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   19
            Left            =   180
            TabIndex        =   256
            Top             =   315
            Width           =   1125
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1590
         Left            =   -66990
         TabIndex        =   290
         Top             =   2400
         Width           =   2115
         Begin VB.TextBox txtmedical 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1305
            TabIndex        =   103
            Top             =   1035
            Width           =   660
         End
         Begin VB.TextBox txtinterview 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1305
            TabIndex        =   102
            Top             =   630
            Width           =   660
         End
         Begin VB.TextBox txtexam 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1305
            TabIndex        =   101
            Top             =   270
            Width           =   660
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Medical Grade"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   91
            Left            =   120
            TabIndex        =   293
            Top             =   1125
            Width           =   1035
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Interview Grade"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   90
            Left            =   120
            TabIndex        =   292
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Exam. Grade"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   89
            Left            =   135
            TabIndex        =   291
            Top             =   315
            Width           =   930
         End
      End
      Begin VB.Frame Frame15 
         Height          =   690
         Left            =   5310
         TabIndex        =   287
         Top             =   315
         Width           =   4875
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
            Left            =   1290
            TabIndex        =   50
            Top             =   270
            Width           =   1065
            _ExtentX        =   1879
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
            Left            =   3630
            TabIndex        =   51
            Top             =   270
            Width           =   1065
            _ExtentX        =   1879
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
            Left            =   2475
            TabIndex        =   289
            Top             =   315
            Width           =   1275
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
            Left            =   90
            TabIndex        =   288
            Top             =   315
            Width           =   1275
         End
      End
      Begin VB.Frame Frame4 
         Height          =   915
         Left            =   135
         TabIndex        =   245
         Top             =   3060
         Width           =   5070
         Begin MSMask.MaskEdBox mskdateres 
            Height          =   285
            Left            =   1215
            TabIndex        =   48
            Top             =   540
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskdateprob 
            Height          =   285
            Left            =   3825
            TabIndex        =   49
            Top             =   540
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskdatereg 
            Height          =   285
            Left            =   3825
            TabIndex        =   47
            Top             =   180
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskendof 
            Height          =   285
            Left            =   7830
            TabIndex        =   77
            ToolTipText     =   "225"
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskDateHired 
            Height          =   285
            Left            =   1215
            TabIndex        =   46
            Top             =   195
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Date Hired"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   16
            Left            =   135
            TabIndex        =   250
            Top             =   225
            Width           =   750
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Date Resign"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   47
            Left            =   135
            TabIndex        =   249
            Top             =   585
            Width           =   870
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Date Regular"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   48
            Left            =   2715
            TabIndex        =   248
            Top             =   210
            Width           =   930
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Date Probation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Index           =   52
            Left            =   2715
            TabIndex        =   247
            Top             =   570
            Width           =   1050
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "End Contract"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   65
            Left            =   6660
            TabIndex        =   246
            Top             =   270
            Width           =   930
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "In case of Emergency"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1815
         Index           =   1
         Left            =   -67935
         TabIndex        =   251
         Top             =   495
         Width           =   3105
         Begin VB.TextBox txtertel 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   990
            MaxLength       =   15
            TabIndex        =   118
            Top             =   1215
            Width           =   2010
         End
         Begin VB.TextBox txtemergency 
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
            Left            =   990
            MaxLength       =   40
            TabIndex        =   116
            Top             =   270
            Width           =   2010
         End
         Begin VB.TextBox txteraddress 
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
            Height          =   510
            Left            =   990
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   117
            Top             =   630
            Width           =   2010
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Telephone#"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   15
            Left            =   90
            TabIndex        =   254
            Top             =   1350
            Width           =   870
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   14
            Left            =   90
            TabIndex        =   253
            Top             =   720
            Width           =   585
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Contact Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   390
            Index           =   13
            Left            =   90
            TabIndex        =   252
            Top             =   270
            Width           =   660
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Dependents"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   -74865
         TabIndex        =   236
         Top             =   2400
         Width           =   7800
         Begin VB.TextBox txtdep5 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   4365
            MaxLength       =   50
            TabIndex        =   98
            Top             =   810
            Width           =   2055
         End
         Begin VB.TextBox txtdep1 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   495
            MaxLength       =   50
            TabIndex        =   90
            Top             =   465
            Width           =   2010
         End
         Begin VB.TextBox txtdep2 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   480
            MaxLength       =   50
            TabIndex        =   92
            Top             =   825
            Width           =   2010
         End
         Begin VB.TextBox txtdep3 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   465
            MaxLength       =   50
            TabIndex        =   94
            Top             =   1185
            Width           =   2040
         End
         Begin VB.TextBox txtdep4 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   4365
            MaxLength       =   50
            TabIndex        =   96
            Top             =   450
            Width           =   2085
         End
         Begin VB.TextBox txtnumdep 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   6660
            MaxLength       =   50
            TabIndex        =   100
            Top             =   1170
            Width           =   435
         End
         Begin MSMask.MaskEdBox mskdep2 
            Height          =   285
            Left            =   2610
            TabIndex        =   93
            Top             =   810
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskdep1 
            Height          =   285
            Left            =   2595
            TabIndex        =   91
            Top             =   450
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskdep3 
            Height          =   285
            Left            =   2640
            TabIndex        =   95
            Top             =   1200
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskdep4 
            Height          =   285
            Left            =   6660
            TabIndex        =   97
            Top             =   450
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskdep5 
            Height          =   285
            Left            =   6660
            TabIndex        =   99
            Top             =   810
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Date of Birth"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   103
            Left            =   6705
            TabIndex        =   244
            Top             =   225
            Width           =   915
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Dep5"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   35
            Left            =   3915
            TabIndex        =   243
            Top             =   900
            Width           =   375
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Dep4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   42
            Left            =   3915
            TabIndex        =   242
            Top             =   585
            Width           =   375
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Dep3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   41
            Left            =   90
            TabIndex        =   241
            Top             =   1215
            Width           =   375
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Dep2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   40
            Left            =   90
            TabIndex        =   240
            Top             =   900
            Width           =   375
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Dep1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   39
            Left            =   90
            TabIndex        =   239
            Top             =   540
            Width           =   375
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Total No. of Dependents"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   27
            Left            =   4635
            TabIndex        =   238
            Top             =   1260
            Width           =   1770
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Date of Birth"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   33
            Left            =   2700
            TabIndex        =   237
            Top             =   225
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Family Background"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   -74865
         TabIndex        =   230
         Top             =   540
         Width           =   10050
         Begin VB.TextBox txtContactPhone 
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
            Left            =   8040
            TabIndex        =   89
            Top             =   1440
            Width           =   1860
         End
         Begin VB.TextBox txtContactAddress 
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
            Left            =   3360
            MaxLength       =   50
            TabIndex        =   88
            Top             =   1440
            Width           =   4140
         End
         Begin VB.TextBox txtContactPerson 
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
            MaxLength       =   50
            TabIndex        =   87
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox txtsjob 
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
            Left            =   7575
            TabIndex        =   86
            Top             =   1080
            Width           =   2340
         End
         Begin VB.TextBox txtmjob 
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
            Left            =   7575
            TabIndex        =   83
            Top             =   735
            Width           =   2340
         End
         Begin VB.TextBox txtfjob 
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
            Left            =   7560
            TabIndex        =   80
            Top             =   375
            Width           =   2340
         End
         Begin VB.TextBox txtspousename 
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
            MaxLength       =   50
            TabIndex        =   84
            Top             =   1095
            Width           =   2055
         End
         Begin VB.TextBox txtmothername 
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
            MaxLength       =   50
            TabIndex        =   81
            Top             =   735
            Width           =   2055
         End
         Begin VB.TextBox txtfathername 
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
            MaxLength       =   50
            TabIndex        =   78
            Top             =   375
            Width           =   2055
         End
         Begin VB.TextBox txtfaddress 
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
            Left            =   3345
            MaxLength       =   50
            TabIndex        =   79
            Top             =   375
            Width           =   4140
         End
         Begin VB.TextBox txtmaddress 
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
            Left            =   3345
            MaxLength       =   50
            TabIndex        =   82
            Top             =   735
            Width           =   4140
         End
         Begin VB.TextBox txtspouseaddress 
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
            Left            =   3360
            TabIndex        =   85
            Top             =   1095
            Width           =   4140
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Tel #:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   23
            Left            =   7560
            TabIndex        =   326
            Top             =   1440
            Width           =   435
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Contact Info"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   22
            Left            =   240
            TabIndex        =   325
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Occupation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   29
            Left            =   8220
            TabIndex        =   272
            Top             =   180
            Width           =   810
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   26
            Left            =   2100
            TabIndex        =   235
            Top             =   135
            Width           =   405
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Spouse Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   38
            Left            =   180
            TabIndex        =   234
            Top             =   1095
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Mother Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   37
            Left            =   180
            TabIndex        =   233
            Top             =   735
            Width           =   960
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Father Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   36
            Left            =   180
            TabIndex        =   232
            Top             =   375
            Width           =   930
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   32
            Left            =   5160
            TabIndex        =   231
            Top             =   135
            Width           =   585
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Educ'l Background"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74865
         TabIndex        =   221
         Top             =   495
         Width           =   6765
         Begin VB.TextBox txtcollegemajor 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   4770
            MaxLength       =   20
            TabIndex        =   115
            Top             =   1455
            Width           =   1920
         End
         Begin VB.TextBox txtcollegecourse 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   4770
            MaxLength       =   20
            TabIndex        =   110
            Top             =   735
            Width           =   1920
         End
         Begin VB.TextBox txtpostcourse 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   4770
            MaxLength       =   20
            TabIndex        =   106
            Top             =   385
            Width           =   1920
         End
         Begin VB.TextBox txtprimarysch 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   114
            Top             =   1455
            Width           =   2820
         End
         Begin VB.TextBox txtsecsch 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   112
            Top             =   1095
            Width           =   2820
         End
         Begin VB.TextBox txtcollegesch 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   109
            Top             =   735
            Width           =   2820
         End
         Begin VB.TextBox txtpostsch 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   105
            Top             =   385
            Width           =   2820
         End
         Begin VB.TextBox txtdatepost 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   930
            TabIndex        =   104
            Top             =   385
            Width           =   915
         End
         Begin VB.TextBox txtdatecollege 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   930
            TabIndex        =   107
            Top             =   735
            Width           =   915
         End
         Begin VB.TextBox txtdatesec 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   930
            TabIndex        =   111
            Top             =   1095
            Width           =   915
         End
         Begin VB.TextBox txtdateprimary 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   930
            TabIndex        =   113
            Top             =   1455
            Width           =   915
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Major In - "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   49
            Left            =   4770
            TabIndex        =   229
            Top             =   1215
            Width           =   750
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Course/Degree"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   50
            Left            =   5325
            TabIndex        =   228
            Top             =   135
            Width           =   1095
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Name of School"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   51
            Left            =   2700
            TabIndex        =   227
            Top             =   180
            Width           =   1110
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Graduated"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   53
            Left            =   855
            TabIndex        =   226
            Top             =   180
            Width           =   1155
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Primary"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   54
            Left            =   135
            TabIndex        =   225
            Top             =   1575
            Width           =   540
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Secondary"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   55
            Left            =   135
            TabIndex        =   224
            Top             =   1215
            Width           =   765
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "College"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   56
            Left            =   135
            TabIndex        =   223
            Top             =   855
            Width           =   525
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Post Graduate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   390
            Index           =   57
            Left            =   135
            TabIndex        =   222
            Top             =   360
            Width           =   675
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Employment History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   -74865
         TabIndex        =   214
         Top             =   2385
         Width           =   10050
         Begin VB.TextBox txtsuperior3 
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
            Left            =   7980
            MaxLength       =   50
            TabIndex        =   136
            Top             =   1140
            Width           =   1935
         End
         Begin VB.TextBox txtsuperior2 
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
            Left            =   7980
            MaxLength       =   50
            TabIndex        =   130
            Top             =   780
            Width           =   1935
         End
         Begin VB.TextBox txtsuperior1 
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
            Left            =   7980
            MaxLength       =   50
            TabIndex        =   124
            Top             =   420
            Width           =   1935
         End
         Begin VB.TextBox txtreason3 
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
            Left            =   6345
            MaxLength       =   50
            TabIndex        =   135
            Top             =   1140
            Width           =   1575
         End
         Begin VB.TextBox txtreason2 
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
            Left            =   6345
            MaxLength       =   50
            TabIndex        =   129
            Top             =   780
            Width           =   1575
         End
         Begin VB.TextBox txtreason1 
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
            Left            =   6345
            MaxLength       =   50
            TabIndex        =   123
            Top             =   420
            Width           =   1575
         End
         Begin VB.TextBox txtpos3 
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
            Left            =   4710
            MaxLength       =   50
            TabIndex        =   134
            Top             =   1140
            Width           =   1575
         End
         Begin VB.TextBox txtpos2 
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
            Left            =   4710
            MaxLength       =   50
            TabIndex        =   128
            Top             =   780
            Width           =   1575
         End
         Begin VB.TextBox txtpos1 
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
            Left            =   4710
            MaxLength       =   50
            TabIndex        =   122
            Top             =   420
            Width           =   1575
         End
         Begin VB.TextBox txtto3 
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
            Left            =   3615
            MaxLength       =   40
            TabIndex        =   133
            Top             =   1140
            Width           =   1035
         End
         Begin VB.TextBox txtto2 
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
            Left            =   3615
            MaxLength       =   40
            TabIndex        =   127
            Top             =   780
            Width           =   1035
         End
         Begin VB.TextBox txtto1 
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
            Left            =   3615
            MaxLength       =   40
            TabIndex        =   121
            Top             =   420
            Width           =   1035
         End
         Begin VB.TextBox txtfrom3 
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
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   132
            Top             =   1140
            Width           =   1035
         End
         Begin VB.TextBox txtfrom2 
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
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   126
            Top             =   780
            Width           =   1035
         End
         Begin VB.TextBox txtfrom1 
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
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   120
            Top             =   420
            Width           =   1035
         End
         Begin VB.TextBox txtnameofcompany3 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   131
            Top             =   1140
            Width           =   2295
         End
         Begin VB.TextBox txtnameofcompany2 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   125
            Top             =   780
            Width           =   2295
         End
         Begin VB.TextBox txtnameofcompany1 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   119
            Top             =   420
            Width           =   2295
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Immediate Superior"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   84
            Left            =   8235
            TabIndex        =   220
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Reason for Leaving"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   81
            Left            =   6435
            TabIndex        =   219
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Position"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   80
            Left            =   5205
            TabIndex        =   218
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   78
            Left            =   4035
            TabIndex        =   217
            Top             =   180
            Width           =   180
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "From"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   77
            Left            =   3000
            TabIndex        =   216
            Top             =   180
            Width           =   360
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name of Company"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   73
            Left            =   480
            TabIndex        =   215
            Top             =   180
            Width           =   1320
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "References"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   -74820
         TabIndex        =   207
         Top             =   2385
         Width           =   10050
         Begin VB.TextBox txtrefname1 
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
            Left            =   135
            MaxLength       =   50
            TabIndex        =   146
            Top             =   390
            Width           =   2295
         End
         Begin VB.TextBox txtrefjob1 
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
            Left            =   2535
            MaxLength       =   30
            TabIndex        =   147
            Top             =   390
            Width           =   1440
         End
         Begin VB.TextBox txtrefaddress1 
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
            Left            =   3990
            MaxLength       =   50
            TabIndex        =   148
            Top             =   390
            Width           =   3225
         End
         Begin VB.TextBox txtrefphone1 
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
            Left            =   7275
            MaxLength       =   35
            TabIndex        =   149
            Top             =   390
            Width           =   1065
         End
         Begin VB.TextBox txtrefrel1 
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
            Left            =   8370
            MaxLength       =   30
            TabIndex        =   150
            Top             =   390
            Width           =   1530
         End
         Begin VB.TextBox txtrefname2 
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
            Left            =   135
            MaxLength       =   50
            TabIndex        =   151
            Top             =   750
            Width           =   2295
         End
         Begin VB.TextBox txtrefjob2 
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
            Left            =   2535
            MaxLength       =   30
            TabIndex        =   152
            Top             =   750
            Width           =   1440
         End
         Begin VB.TextBox txtrefaddress2 
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
            Left            =   3990
            MaxLength       =   50
            TabIndex        =   153
            Top             =   750
            Width           =   3225
         End
         Begin VB.TextBox txtrefphone2 
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
            Left            =   7275
            MaxLength       =   35
            TabIndex        =   154
            Top             =   750
            Width           =   1065
         End
         Begin VB.TextBox txtrefrel2 
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
            Left            =   8370
            MaxLength       =   30
            TabIndex        =   155
            Top             =   750
            Width           =   1530
         End
         Begin VB.TextBox txtrefname3 
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
            Left            =   135
            MaxLength       =   50
            TabIndex        =   156
            Top             =   1110
            Width           =   2295
         End
         Begin VB.TextBox txtrefjob3 
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
            Left            =   2535
            MaxLength       =   30
            TabIndex        =   157
            Top             =   1110
            Width           =   1440
         End
         Begin VB.TextBox txtrefaddress3 
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
            Left            =   3990
            MaxLength       =   50
            TabIndex        =   158
            Top             =   1110
            Width           =   3225
         End
         Begin VB.TextBox txtrefphone3 
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
            Left            =   7275
            MaxLength       =   35
            TabIndex        =   159
            Top             =   1110
            Width           =   1065
         End
         Begin VB.TextBox txtrefrel3 
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
            Left            =   8370
            MaxLength       =   30
            TabIndex        =   160
            Top             =   1110
            Width           =   1530
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "References"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   74
            Left            =   315
            TabIndex        =   213
            Top             =   450
            Width           =   825
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   75
            Left            =   975
            TabIndex        =   212
            Top             =   150
            Width           =   405
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Occupation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   76
            Left            =   2835
            TabIndex        =   211
            Top             =   150
            Width           =   810
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Business Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   79
            Left            =   5040
            TabIndex        =   210
            Top             =   150
            Width           =   1245
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Business Phone"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   82
            Left            =   7245
            TabIndex        =   209
            Top             =   135
            Width           =   1110
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Relationships"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   83
            Left            =   8685
            TabIndex        =   208
            Top             =   150
            Width           =   945
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1680
         Left            =   -74820
         TabIndex        =   197
         Top             =   540
         Width           =   9915
         Begin VB.TextBox txtdateseminar3 
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
            Height          =   285
            Left            =   7845
            MaxLength       =   40
            TabIndex        =   145
            Top             =   1125
            Width           =   1890
         End
         Begin VB.TextBox txtdateseminar2 
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
            Height          =   285
            Left            =   7845
            MaxLength       =   40
            TabIndex        =   142
            Top             =   735
            Width           =   1890
         End
         Begin VB.TextBox txtdateseminar1 
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
            Height          =   285
            Left            =   7845
            MaxLength       =   40
            TabIndex        =   139
            Top             =   375
            Width           =   1890
         End
         Begin VB.TextBox txtdialect1 
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
            Height          =   285
            Left            =   405
            MaxLength       =   25
            TabIndex        =   137
            Top             =   375
            Width           =   1845
         End
         Begin VB.TextBox txtdialect2 
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
            Height          =   285
            Left            =   420
            MaxLength       =   25
            TabIndex        =   140
            Top             =   735
            Width           =   1845
         End
         Begin VB.TextBox txtdialect3 
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
            Height          =   285
            Left            =   420
            MaxLength       =   25
            TabIndex        =   143
            Top             =   1095
            Width           =   1845
         End
         Begin VB.TextBox txtseminar3 
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
            Height          =   285
            Left            =   2550
            MaxLength       =   40
            TabIndex        =   144
            Top             =   1125
            Width           =   5205
         End
         Begin VB.TextBox txtseminar2 
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
            Height          =   285
            Left            =   2550
            MaxLength       =   40
            TabIndex        =   141
            Top             =   735
            Width           =   5205
         End
         Begin VB.TextBox txtseminar1 
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
            Height          =   285
            Left            =   2550
            MaxLength       =   40
            TabIndex        =   138
            Top             =   375
            Width           =   5205
         End
         Begin VB.Label Label22 
            Caption         =   "(3)"
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   71
            Left            =   180
            TabIndex        =   206
            Top             =   1155
            Width           =   255
         End
         Begin VB.Label Label22 
            Caption         =   "(2)"
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   70
            Left            =   180
            TabIndex        =   205
            Top             =   795
            Width           =   255
         End
         Begin VB.Label Label22 
            Caption         =   "(1)"
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   69
            Left            =   180
            TabIndex        =   204
            Top             =   435
            Width           =   255
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Languages/Dialects"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   68
            Left            =   585
            TabIndex        =   203
            Top             =   135
            Width           =   1395
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dates"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   67
            Left            =   8475
            TabIndex        =   202
            Top             =   135
            Width           =   1230
         End
         Begin VB.Label Label22 
            Caption         =   "(3)"
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   62
            Left            =   2310
            TabIndex        =   201
            Top             =   1155
            Width           =   255
         End
         Begin VB.Label Label22 
            Caption         =   "(2)"
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   61
            Left            =   2310
            TabIndex        =   200
            Top             =   795
            Width           =   255
         End
         Begin VB.Label Label22 
            Caption         =   "(1)"
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   59
            Left            =   2310
            TabIndex        =   199
            Top             =   435
            Width           =   255
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trainings/Seminars Attended"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   58
            Left            =   3465
            TabIndex        =   198
            Top             =   135
            Width           =   3510
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   -74865
         TabIndex        =   165
         Top             =   540
         Width           =   10050
         Begin VB.TextBox txtpagibig 
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
            Left            =   8280
            TabIndex        =   29
            Top             =   1980
            Width           =   1695
         End
         Begin VB.ComboBox cmbcivilstatus 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   8295
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1575
            Width           =   1695
         End
         Begin VB.OptionButton optmale 
            Appearance      =   0  'Flat
            Caption         =   "Male"
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
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   5190
            TabIndex        =   18
            Top             =   1275
            Width           =   735
         End
         Begin VB.OptionButton optfemale 
            Appearance      =   0  'Flat
            Caption         =   "Female"
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
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   5190
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1620
            Width           =   840
         End
         Begin VB.TextBox txtissueat 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5985
            TabIndex        =   36
            Top             =   2715
            Width           =   1365
         End
         Begin VB.TextBox txtrescer 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3840
            TabIndex        =   35
            Top             =   2745
            Width           =   1215
         End
         Begin VB.TextBox txtnbi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1080
            TabIndex        =   34
            Top             =   2715
            Width           =   1695
         End
         Begin VB.TextBox txtreligion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1080
            TabIndex        =   26
            Top             =   1995
            Width           =   1695
         End
         Begin VB.TextBox txteyecolor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3840
            TabIndex        =   31
            Top             =   2355
            Width           =   1215
         End
         Begin VB.TextBox txthaircolor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1080
            TabIndex        =   30
            Top             =   2355
            Width           =   1695
         End
         Begin VB.TextBox txtacctno 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   8280
            TabIndex        =   33
            Top             =   2355
            Width           =   1695
         End
         Begin VB.TextBox txthouseno 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   945
            TabIndex        =   7
            Top             =   180
            Width           =   705
         End
         Begin VB.TextBox txtheight 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6675
            MaxLength       =   8
            TabIndex        =   19
            Top             =   1215
            Width           =   540
         End
         Begin VB.TextBox txtweight 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6690
            MaxLength       =   8
            TabIndex        =   21
            Top             =   1620
            Width           =   540
         End
         Begin VB.TextBox txttel2 
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
            Left            =   8265
            MaxLength       =   15
            TabIndex        =   22
            Top             =   1260
            Width           =   1695
         End
         Begin VB.TextBox txttel1 
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
            Left            =   8265
            MaxLength       =   15
            TabIndex        =   15
            Top             =   900
            Width           =   1695
         End
         Begin VB.TextBox txtprovinceaddress 
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
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   13
            Top             =   900
            Width           =   3255
         End
         Begin VB.TextBox txtaddress1 
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
            Left            =   4905
            TabIndex        =   12
            Top             =   540
            Width           =   5055
         End
         Begin VB.TextBox txtbirthplace 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   23
            Top             =   1620
            Width           =   1680
         End
         Begin VB.TextBox txtstreet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2160
            TabIndex        =   8
            Top             =   180
            Width           =   1800
         End
         Begin VB.TextBox txtbarangay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4455
            TabIndex        =   9
            Top             =   180
            Width           =   2370
         End
         Begin VB.TextBox txtcity 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   450
            TabIndex        =   11
            Top             =   540
            Width           =   3480
         End
         Begin VB.TextBox txtprovince 
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
            Left            =   5640
            TabIndex        =   14
            Top             =   900
            Width           =   2070
         End
         Begin VB.TextBox txtdistrict 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7425
            TabIndex        =   10
            Top             =   180
            Width           =   2535
         End
         Begin VB.TextBox txtphilhealth 
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
            Left            =   5970
            TabIndex        =   32
            Top             =   2355
            Width           =   1365
         End
         Begin VB.TextBox txtage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3825
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   17
            Top             =   1260
            Width           =   495
         End
         Begin VB.TextBox txtcitizen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3840
            MaxLength       =   20
            TabIndex        =   24
            Top             =   1635
            Width           =   1200
         End
         Begin MSMask.MaskEdBox mskbday 
            Height          =   285
            Left            =   1080
            TabIndex        =   16
            Top             =   1245
            Width           =   1200
            _ExtentX        =   2117
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
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msktax 
            Height          =   285
            Left            =   5985
            TabIndex        =   28
            Top             =   1980
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            Enabled         =   0   'False
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###-###-###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msksss 
            Height          =   285
            Left            =   3840
            TabIndex        =   27
            Top             =   1980
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   12582912
            Enabled         =   0   'False
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##-#######-#"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskissue 
            Height          =   285
            Left            =   8280
            TabIndex        =   37
            Top             =   2745
            Width           =   1695
            _ExtentX        =   2990
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
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Issue On"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   46
            Left            =   7440
            TabIndex        =   195
            Top             =   2835
            Width           =   645
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Issue At"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   45
            Left            =   5160
            TabIndex        =   194
            Top             =   2835
            Width           =   600
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Res. Cer."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   44
            Left            =   3000
            TabIndex        =   193
            Top             =   2835
            Width           =   690
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "NBI"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   43
            Left            =   120
            TabIndex        =   192
            Top             =   2835
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Religion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   120
            TabIndex        =   191
            Top             =   2130
            Width           =   555
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Eye Color"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   95
            Left            =   3000
            TabIndex        =   190
            Top             =   2475
            Width           =   690
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Hair Color"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   94
            Left            =   120
            TabIndex        =   189
            Top             =   2475
            Width           =   705
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Acct. No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   93
            Left            =   7440
            TabIndex        =   188
            Top             =   2475
            Width           =   675
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Civil Status"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   88
            Left            =   7440
            TabIndex        =   187
            Top             =   1680
            Width           =   795
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "# of House"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   186
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Weight"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   12
            Left            =   6045
            TabIndex        =   185
            Top             =   1695
            Width           =   510
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Height"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   11
            Left            =   6045
            TabIndex        =   184
            Top             =   1335
            Width           =   465
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "PAGIBIG#"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   6
            Left            =   7440
            TabIndex        =   183
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "BirthDate"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   182
            Top             =   1380
            Width           =   675
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Cell No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   8
            Left            =   7440
            TabIndex        =   181
            Top             =   1320
            Width           =   555
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Tel. #"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   7
            Left            =   7800
            TabIndex        =   180
            Top             =   960
            Width           =   435
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Provincial Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   179
            Top             =   1020
            Width           =   1305
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Address 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   0
            Left            =   4020
            TabIndex        =   178
            Top             =   630
            Width           =   720
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Age"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   0
            Left            =   3000
            TabIndex        =   177
            Top             =   1380
            Width           =   285
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Citizenship"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   5
            Left            =   3000
            TabIndex        =   176
            Top             =   1740
            Width           =   765
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Birth Place"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   175
            Top             =   1740
            Width           =   750
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "SSS#"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   2
            Left            =   3000
            TabIndex        =   174
            Top             =   2100
            Width           =   390
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Tax ID#"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   3
            Left            =   5160
            TabIndex        =   173
            Top             =   2100
            Width           =   600
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Street"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   98
            Left            =   1695
            TabIndex        =   172
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Brgy."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   99
            Left            =   4020
            TabIndex        =   171
            Top             =   255
            Width           =   390
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "City"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   100
            Left            =   120
            TabIndex        =   170
            Top             =   615
            Width           =   285
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Province"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   101
            Left            =   4920
            TabIndex        =   169
            Top             =   1020
            Width           =   615
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "District"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   102
            Left            =   6885
            TabIndex        =   168
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "PhilHealth"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   104
            Left            =   5160
            TabIndex        =   167
            Top             =   2445
            Width           =   705
         End
      End
      Begin VB.Frame Frame13 
         Height          =   345
         Left            =   5280
         TabIndex        =   278
         Top             =   4320
         Visible         =   0   'False
         Width           =   4875
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
            Left            =   3360
            TabIndex        =   279
            Text            =   "15"
            Top             =   600
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   13
            Left            =   540
            TabIndex        =   309
            Top             =   405
            Width           =   1095
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   12
            Left            =   3735
            TabIndex        =   308
            Top             =   405
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   11
            Left            =   2115
            TabIndex        =   307
            Top             =   405
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3030
         Index           =   0
         Left            =   5310
         TabIndex        =   280
         Top             =   960
         Width           =   4920
         Begin VB.ComboBox cmbAllowance 
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
            ItemData        =   "frmHRFinalEmployees.frx":0622
            Left            =   120
            List            =   "frmHRFinalEmployees.frx":0635
            Style           =   2  'Dropdown List
            TabIndex        =   334
            Top             =   1440
            Width           =   1485
         End
         Begin VB.CheckBox chkFund 
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   4560
            TabIndex        =   75
            ToolTipText     =   "Active if checked"
            Top             =   4440
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.CheckBox chkMisc 
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   4560
            TabIndex        =   74
            ToolTipText     =   "Active if checked"
            Top             =   3900
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.ComboBox cmbPayBasis 
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
            ItemData        =   "frmHRFinalEmployees.frx":0668
            Left            =   120
            List            =   "frmHRFinalEmployees.frx":0678
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   345
            Width           =   1485
         End
         Begin VB.TextBox txtRate 
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
            Left            =   1680
            TabIndex        =   53
            Text            =   "15"
            Top             =   345
            Width           =   1005
         End
         Begin VB.TextBox txtMonthsHired 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   2850
            TabIndex        =   54
            Text            =   "15"
            Top             =   345
            Width           =   645
         End
         Begin VB.TextBox txtMonthlyRate 
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
            Left            =   3705
            TabIndex        =   55
            Text            =   "1000"
            Top             =   345
            Width           =   1005
         End
         Begin VB.TextBox TxtMFPayment 
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
            Left            =   1140
            TabIndex        =   62
            Text            =   "0.00"
            Top             =   3900
            Width           =   915
         End
         Begin VB.TextBox txtPFPayment 
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
            Left            =   1680
            TabIndex        =   57
            Text            =   "0.00"
            Top             =   1440
            Width           =   1005
         End
         Begin VB.TextBox TxtMFAmount 
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
            Left            =   135
            TabIndex        =   61
            Text            =   "0.00"
            Top             =   3900
            Width           =   915
         End
         Begin VB.TextBox txtPFAmount 
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
            Left            =   1680
            TabIndex        =   56
            Text            =   "0.00"
            Top             =   2040
            Width           =   1005
         End
         Begin VB.TextBox txtCBpayment 
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
            Left            =   1140
            TabIndex        =   66
            Text            =   "0.00"
            Top             =   4440
            Width           =   915
         End
         Begin VB.TextBox txtCBtotal 
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
            Left            =   135
            TabIndex        =   65
            Text            =   "0.00"
            Top             =   4440
            Width           =   915
         End
         Begin MSMask.MaskEdBox mskStartCB 
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
            Left            =   2310
            TabIndex        =   67
            Top             =   4440
            Width           =   1035
            _ExtentX        =   1826
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
         Begin MSMask.MaskEdBox mskEndCB 
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
            Left            =   3450
            TabIndex        =   68
            Top             =   4440
            Width           =   1035
            _ExtentX        =   1826
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
         Begin MSMask.MaskEdBox MskMFStartPayment 
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
            Left            =   2310
            TabIndex        =   63
            Top             =   3900
            Width           =   1035
            _ExtentX        =   1826
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
         Begin MSMask.MaskEdBox MskMFEndPayment 
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
            Left            =   3450
            TabIndex        =   64
            Top             =   3900
            Width           =   1035
            _ExtentX        =   1826
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
         Begin MSMask.MaskEdBox MskPFStartPayment 
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
            Left            =   2310
            TabIndex        =   59
            Top             =   3600
            Width           =   1035
            _ExtentX        =   1826
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
         Begin MSMask.MaskEdBox MskPFEndPayment 
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
            Left            =   3450
            TabIndex        =   60
            Top             =   3600
            Width           =   1035
            _ExtentX        =   1826
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
         Begin VB.TextBox TxtEcolaAmount 
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
            Left            =   1680
            TabIndex        =   69
            Text            =   "0.00"
            Top             =   720
            Width           =   1005
         End
         Begin VB.TextBox txtBillingRate 
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
            Left            =   1140
            TabIndex        =   76
            Text            =   "0.00"
            Top             =   3600
            Width           =   915
         End
         Begin VB.TextBox txtSuretyPercent 
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
            Left            =   2160
            TabIndex        =   58
            Text            =   "0.00"
            Top             =   3360
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CheckBox ChkEcola 
            Caption         =   "Show Deductions"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1680
            TabIndex        =   297
            Top             =   3480
            Visible         =   0   'False
            Width           =   1890
         End
         Begin MSMask.MaskEdBox mskStartCash 
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
            Left            =   2310
            TabIndex        =   70
            Top             =   4800
            Width           =   1035
            _ExtentX        =   1826
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
         Begin MSMask.MaskEdBox mskEndCash 
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
            Left            =   3450
            TabIndex        =   73
            Top             =   4800
            Width           =   1035
            _ExtentX        =   1826
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
         Begin MSMask.MaskEdBox mskStartUniform 
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
            Left            =   2310
            TabIndex        =   71
            Top             =   3600
            Width           =   1035
            _ExtentX        =   1826
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
         Begin MSMask.MaskEdBox mskEndUniform 
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
            Left            =   3450
            TabIndex        =   72
            Top             =   3600
            Width           =   1035
            _ExtentX        =   1826
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
            Caption         =   "BSC Allowance 2:"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   240
            TabIndex        =   336
            Top             =   2040
            Width           =   1515
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   120
            X2              =   4680
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Allowance Cutoff"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   333
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "(%)"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   2040
            TabIndex        =   323
            Top             =   3480
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Basis"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   120
            TabIndex        =   322
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label lblRAte 
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
            Left            =   1680
            TabIndex        =   321
            Top             =   120
            Width           =   1095
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
            Left            =   2850
            TabIndex        =   320
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label3 
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
            Left            =   3705
            TabIndex        =   319
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Allowance Rate"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   1665
            TabIndex        =   304
            Top             =   1200
            Width           =   1155
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "(Payments)"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   1185
            TabIndex        =   303
            Top             =   3675
            Width           =   1035
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Start Payment"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   2310
            TabIndex        =   302
            Top             =   3480
            Width           =   1065
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "End Payment"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3450
            TabIndex        =   301
            Top             =   3495
            Width           =   975
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Start Payment"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   2310
            TabIndex        =   300
            Top             =   3675
            Width           =   1065
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "End Payment"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3465
            TabIndex        =   299
            Top             =   3675
            Width           =   1095
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "LMC Ded."
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   180
            TabIndex        =   295
            Top             =   3675
            Width           =   1035
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Ecola Rate"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   780
            TabIndex        =   294
            Top             =   735
            Width           =   1035
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "End Payment"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   3465
            TabIndex        =   284
            Top             =   4215
            Width           =   975
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Start Payment"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   2310
            TabIndex        =   283
            Top             =   4215
            Width           =   1065
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "(Payments)"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   1185
            TabIndex        =   282
            Top             =   4215
            Width           =   1035
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Planters"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   165
            TabIndex        =   281
            Top             =   4215
            Width           =   825
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Cash/Adv."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   165
            TabIndex        =   315
            Top             =   3720
            Width           =   855
         End
      End
      Begin VB.Frame Frame16 
         Height          =   3030
         Left            =   5310
         TabIndex        =   324
         Top             =   960
         Visible         =   0   'False
         Width           =   4920
      End
   End
   Begin VB.Frame Frame7 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   5760
      TabIndex        =   276
      Top             =   45
      Width           =   4770
      Begin VB.Label lblCOMPANY 
         Alignment       =   1  'Right Justify
         Caption         =   "COMPANY NAME"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   120
         TabIndex        =   296
         Top             =   45
         Width           =   4380
      End
   End
   Begin VB.CommandButton cmdlast 
      Height          =   255
      Left            =   9975
      Picture         =   "frmHRFinalEmployees.frx":06A0
      Style           =   1  'Graphical
      TabIndex        =   268
      ToolTipText     =   "Last record"
      Top             =   6345
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Height          =   255
      Left            =   9735
      Picture         =   "frmHRFinalEmployees.frx":07EA
      Style           =   1  'Graphical
      TabIndex        =   267
      ToolTipText     =   "Next record"
      Top             =   6345
      Width           =   255
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   255
      Left            =   9495
      Picture         =   "frmHRFinalEmployees.frx":0934
      Style           =   1  'Graphical
      TabIndex        =   266
      ToolTipText     =   "Previous record"
      Top             =   6345
      Width           =   255
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   255
      Left            =   9255
      Picture         =   "frmHRFinalEmployees.frx":0A7E
      Style           =   1  'Graphical
      TabIndex        =   265
      ToolTipText     =   "Top record"
      Top             =   6345
      Width           =   255
   End
   Begin VB.TextBox txtHide 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Height          =   240
      Left            =   9720
      MouseIcon       =   "frmHRFinalEmployees.frx":0BC8
      MousePointer    =   99  'Custom
      TabIndex        =   196
      Text            =   "Hide"
      Top             =   5625
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   108
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdReportSolo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4080
         Picture         =   "frmHRFinalEmployees.frx":0ED2
         Style           =   1  'Graphical
         TabIndex        =   314
         ToolTipText     =   "Search"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdRecontract 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         Picture         =   "frmHRFinalEmployees.frx":1314
         Style           =   1  'Graphical
         TabIndex        =   312
         ToolTipText     =   "Transfer Employee"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton btnEmployeeSearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4680
         Picture         =   "frmHRFinalEmployees.frx":1756
         Style           =   1  'Graphical
         TabIndex        =   277
         ToolTipText     =   "Search"
         Top             =   0
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton btnEmpRefresh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3525
         Picture         =   "frmHRFinalEmployees.frx":1858
         Style           =   1  'Graphical
         TabIndex        =   275
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdundo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2940
         Picture         =   "frmHRFinalEmployees.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   166
         ToolTipText     =   "Undo"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdupdate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         Picture         =   "frmHRFinalEmployees.frx":200C
         Style           =   1  'Graphical
         TabIndex        =   163
         ToolTipText     =   "Update"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2355
         Picture         =   "frmHRFinalEmployees.frx":253E
         Style           =   1  'Graphical
         TabIndex        =   164
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdedit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   625
         Picture         =   "frmHRFinalEmployees.frx":2640
         Style           =   1  'Graphical
         TabIndex        =   162
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdNew 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   45
         Picture         =   "frmHRFinalEmployees.frx":278A
         Style           =   1  'Graphical
         TabIndex        =   161
         ToolTipText     =   "New"
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
   Begin TabDlg.SSTab SSTab2 
      Height          =   6120
      Left            =   120
      TabIndex        =   259
      Top             =   600
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   10795
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmHRFinalEmployees.frx":2CBC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label22(17)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label22(18)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label22(20)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label22(21)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label22(25)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label22(9)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label22(24)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label22(28)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtfirstname"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtmiddlename"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtnickname"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtsurname"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtControlNo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtID"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdCopy"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtLastDate"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmbChangeCompany"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdChangeCompany"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "EmpPicture"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtGenInfo"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtShiftSched"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "CommonDialog1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtDTRID"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmbShiftSchedule"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtClientCode"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      Begin VB.TextBox txtClientCode 
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6120
         MaxLength       =   15
         TabIndex        =   335
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cmbShiftSchedule 
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
         Height          =   315
         ItemData        =   "frmHRFinalEmployees.frx":2CD8
         Left            =   6360
         List            =   "frmHRFinalEmployees.frx":2CDA
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "cmbShiftSchedule"
         Top             =   810
         Width           =   2160
      End
      Begin VB.TextBox txtDTRID 
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3765
         MaxLength       =   15
         TabIndex        =   331
         ToolTipText     =   "Daily time record ID #"
         Top             =   180
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7680
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtShiftSched 
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
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   6360
         TabIndex        =   330
         Top             =   810
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtGenInfo 
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
         Left            =   120
         TabIndex        =   328
         Top             =   1200
         Width           =   8415
      End
      Begin VB.PictureBox EmpPicture 
         Height          =   1335
         Left            =   8760
         MouseIcon       =   "frmHRFinalEmployees.frx":2CDC
         MousePointer    =   4  'Icon
         ScaleHeight     =   1275
         ScaleWidth      =   1395
         TabIndex        =   327
         ToolTipText     =   "Double click to change picture on Add or Edit"
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdChangeCompany 
         Caption         =   "Select"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   318
         ToolTipText     =   "Select Company"
         Top             =   1710
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.ComboBox cmbChangeCompany 
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
         Height          =   315
         ItemData        =   "frmHRFinalEmployees.frx":2FE6
         Left            =   360
         List            =   "frmHRFinalEmployees.frx":2FE8
         Style           =   2  'Dropdown List
         TabIndex        =   317
         Top             =   1710
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.TextBox txtLastDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   7080
         TabIndex        =   310
         Text            =   "Date"
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   ">>"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   305
         Top             =   1740
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox TxtID 
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3285
         MaxLength       =   20
         TabIndex        =   298
         Top             =   1740
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox TxtControlNo 
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1260
         MaxLength       =   15
         TabIndex        =   0
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox txtsurname 
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
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   810
         Width           =   1575
      End
      Begin VB.TextBox txtnickname 
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
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   4
         Top             =   810
         Width           =   1575
      End
      Begin VB.TextBox txtmiddlename 
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
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   3
         Top             =   810
         Width           =   1575
      End
      Begin VB.TextBox txtfirstname 
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
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   2
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   28
         Left            =   3000
         TabIndex        =   332
         Top             =   225
         Width           =   570
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Schedule"
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
         Index           =   24
         Left            =   6360
         TabIndex        =   329
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
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
         Index           =   9
         Left            =   6120
         TabIndex        =   311
         Top             =   5760
         Width           =   930
      End
      Begin VB.Label Label1 
         Height          =   240
         Left            =   4545
         TabIndex        =   269
         Top             =   45
         Width           =   1950
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nick Name"
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
         Index           =   25
         Left            =   4800
         TabIndex        =   264
         Top             =   615
         Width           =   735
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middlename"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   3240
         TabIndex        =   263
         Top             =   615
         Width           =   840
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Firstname"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   1680
         TabIndex        =   262
         Top             =   615
         Width           =   705
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Surname"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   135
         TabIndex        =   261
         Top             =   615
         Width           =   630
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   135
         TabIndex        =   260
         Top             =   225
         Width           =   990
      End
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2280
      TabIndex        =   313
      Top             =   6760
      Width           =   6375
   End
End
Attribute VB_Name = "frmHRFinalEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form Name:    HR Main Form
' Last Update:  January 17, 2010
' Updated by:   Bernard O. Sahagun
' Client:       Tas Trans

'Option Explicit

Public rs As ADODB.Recordset
Public prevControlNo As String
Public PrevCompany As String
Public PrevStartContract As String
Public PrevEndContract As String
Public Cid As Integer
Public PaymentBasis As String
Public MonthlyRate
Public DailyRate
Public HourlyRate
Public CommissionRate
Public WithEcolaPay As Integer
Public WithMiscPay As Integer
Public WithFundPay As Integer

Public EcolaAmount
Public BrCode  As Integer
Public PosCode  As Integer
Public DeptCode  As Integer
Public mShiftCode As Integer
 
Public rsHREmp As ADODB.Recordset   ' HR employees recordset
Public rsHREmpEdit As ADODB.Recordset

Dim PictureFile As String

Dim rsDetails As ADODB.Recordset
Dim rsRef As ADODB.Recordset
Dim rsDialect As ADODB.Recordset
Dim rsSem As ADODB.Recordset
Dim rsCivil As ADODB.Recordset
Dim rsEmpStatus As ADODB.Recordset
Dim rsEmpRec As ADODB.Recordset
Dim rsEm As ADODB.Recordset
Dim rsDelete As ADODB.Recordset
Dim rsemp As ADODB.Recordset
Dim rstPos As ADODB.Recordset
Dim rstShift As ADODB.Recordset
Dim rsIsExist As ADODB.Recordset

Dim dep1bdate, dep2bdate, dep3bdate, dep4bdate, dep5bdate, Status, sqlString, sqlString2
Public TmpSurname As String
Public TmpFirstname As String
Public TmpMiddlename As String

Dim rstclient As ADODB.Recordset
Dim rstBranch As ADODB.Recordset
Dim rstDept As ADODB.Recordset

Dim HRDeptSw As Boolean
Dim HRPosSw As Boolean
Dim HRBranchSw As Boolean

Dim OrigAcctNo As String
Dim origCivilStatus As String
Dim OrigEmpStatus As String
Dim OrigCompany As String
Dim OrigDateHired As Variant
Dim OrigStartContract As Variant
Dim OrigEndContract As Variant
Dim OrigPayBasis As String
Dim OrigRate As String
Dim OrigMonthly As String
Dim OrigTASAesla As String
Dim OrigSSSNo As String
Dim OrigPHealthNo As String
Dim OrigPagibig As String
    
Private Sub cmbCompany_LostFocus()
Cid = cmbCompany.ItemData(cmbCompany.ListIndex)
End Sub

Private Sub cmbChangeCompany_LostFocus()
' Get the specific client code for filtering of data.
CurrentClientCode = cmbChangeCompany.ItemData(cmbChangeCompany.ListIndex)
End Sub

Private Sub cmbDept_Click()
 If HRDeptSw = True Then
    Call RefreshDept
    HRDeptSw = False
End If
End Sub

Private Sub cmbBranch_Click()
 If HRBranchSw = True Then
    Call RefreshBranch
    HRBranchSw = False
End If
End Sub

Private Sub cmbPosition_Click()
If HRPosSw = True Then
    Call RefreshPosition
    HRPosSw = False
End If
End Sub

'Private Sub cmbUniform_Click()
'  On Error Resume Next
'
'        TxtEcolaAmount.Visible = False
'        mskStartUniform.Visible = False
'        mskEndUniform.Visible = False
'        txtBillingRate.Visible = False
'        mskStartCash.Visible = False
'        mskEndCash.Visible = False
'
'    If cmbUniform.ListIndex = 0 Then ' Cash Advance
'        TxtEcolaAmount.Visible = True
'        mskStartCash.Visible = True
'        mskEndCash.Visible = True
'    Else        ' Uniform
'        txtBillingRate.Visible = True
'        mskStartUniform.Visible = True
'        mskEndUniform.Visible = True
'    End If
'
'End Sub

Private Sub cmdAddPosition_Click()
    HRPosSw = True
    frmPositions.Show
End Sub

Private Sub cmdRecontract_Click()
     
    reContract = False
    If MsgBox("Do you want to transfer this employee to other Company?", vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
    End If

    editflagship = True
    reContract = True
    SSTab1.Tab = 1
    
    cmbCompany.Enabled = True
    addflagship = False
    lblMessage.Caption = "You are about to TRANSFER this employee to other Company."
    Call DisAble_ActionButtons
    cmbCompany.SetFocus
    
End Sub

Private Sub cmdReportSolo_Click()
Dim strHistory As String
Dim rsItems As ADODB.Recordset

' "IIf(AgencyCode = 1,'" & getAgencyNamebyID(1) & "', IIf(AgencyCode = 2,'" & getAgencyNamebyID(2) & "', IIf(AgencyCode = 3,'" & getAgencyNamebyID(3) & "', IIf(AgencyCode = 4,'Not Defined', )))) AS AgencyName, " & _

'original code
'strHistory = "SELECT EmployeeID, EmpNo, AgencyCode, ClientCode, '" & PublicFunctions.getClientNamebyEmployeeID(txtID.Text) & "' as ClientName, FirstName, Middlename, Surname, NickName, EmpPosition, Age, ApplicationDate, EmpStatus, BirthDate, Address1, HouseNo, Street, Barangay, District, City, Address2, Town, Province1, Province, Contact, Cphone, CAddress, Gender, Weight, Height, ExamGrade, InterviewGrade, TelNum1, TelNum2, TaxCode, DateHired, DateResign, SSS, Philhealth, PHCode, Pagibig, EmploymentStatus, AcctNo, Citizen, Birthplace, HairColor, EyeColor, Religion, DateRegular, DateProb, TIN, MedicalGrade, [Primary], DatePrimary, Secondary, DateSecondary, College, DateCollege, Degree, Major, PostGraduate, DatePostGraduate, PostGraduateCourse, FatherName, Faddress, FWork, MotherName, Maddress, MWork, NBI, ResCer, IssueAt, IssueOn, Supervisor, ReasonRes, Dept_No, Confi, " & _
'            "'" & PublicFunctions.getAgencyName(CompanyShort) & "' AS AgencyName, " & _
'            " PicturePath, EndOfContract, EmploymentType, DeptCode, BrCode, JobCode, JobClass, Paymentbasis, MonthlyRate, DailyRate, HourlyRate, StartContract, EndContract, NoOfMonthsHired, WithEcolaPay,    , CBamount, CBpayment," & _
'            " CBstartpayment , CBendpayment, MiscFee, MFAmount, MFpayment, MFStartPayment, MFEndPayment, CashAdvance, PFAmount, PFPayment, PFStartPayment, PFEndPayment, Remarks, USERID, LastUpdate" & _
'            " From HREmployees where EmployeeID=" & txtID.Text & _
'            " Union" & _
'            " SELECT EmployeeID, EmpNo, AgencyCode, ClientCode, ClientName, FirstName, Middlename, Surname, NickName, EmpPosition, Age, ApplicationDate, EmpStatus, BirthDate, Address1, HouseNo, Street, Barangay, District, City, Address2, Town, Province1, Province, Contact, Cphone, CAddress, Gender, Weight, Height, ExamGrade, InterviewGrade, TelNum1, TelNum2, TaxCode, DateHired, DateResign, SSS, Philhealth, PHCode, Pagibig, EmploymentStatus, AcctNo, Citizen, Birthplace, HairColor, EyeColor, Religion, DateRegular, DateProb, TIN, MedicalGrade, [Primary], DatePrimary, Secondary, DateSecondary, College, DateCollege, Degree, Major, PostGraduate, DatePostGraduate, PostGraduateCourse, FatherName, Faddress, FWork, MotherName, Maddress, MWork, NBI, ResCer, IssueAt, IssueOn, Supervisor, ReasonRes, Dept_No, Confi, " & _
'            " AgencyName, " & _
'            " PicturePath, EndOfContract, EmploymentType, DeptCode, BrCode, JobCode, JobClass, Paymentbasis, MonthlyRate, DailyRate, HourlyRate, StartContract, EndContract, NoOfMonthsHired, WithEcolaPay, EcolaAmount, CBamount, CBpayment, " & _
'            " CBstartpayment , CBendpayment, MiscFee, MFAmount, MFpayment, MFStartPayment, MFEndPayment, CashAdvance, PFAmount, PFPayment, PFStartPayment, PFEndPayment, Remarks, USERID, LastUpdate " & _
'            " FROM HREmployees_History where EmployeeID=" & txtID.Text

strHistory = "SELECT EmployeeID, EmpNo, AgencyCode, ClientCode, '" & PublicFunctions.getClientNamebyEmployeeID(txtID.Text) & "' as ClientName, FirstName, Middlename, Surname, NickName, EmpPosition, Age, ApplicationDate, EmpStatus, BirthDate, Address1, HouseNo, Street, Barangay, District, City, Address2, Town, Province1, Province, Contact, Cphone, CAddress, Gender, Weight, Height, ExamGrade, InterviewGrade, TelNum1, TelNum2, TaxCode, DateHired, DateResign, SSS, Philhealth, PHCode, Pagibig, EmploymentStatus, AcctNo, Citizen, Birthplace, HairColor, EyeColor, Religion, DateRegular, DateProb, TIN, MedicalGrade, [Primary], DatePrimary, Secondary, DateSecondary, College, DateCollege, Degree, Major, PostGraduate, DatePostGraduate, PostGraduateCourse, FatherName, Faddress, FWork, MotherName, Maddress, MWork, NBI, ResCer, IssueAt, IssueOn, Supervisor, ReasonRes, Dept_No, Confi, " & _
            "'" & PublicFunctions.getAgencyName(CompanyShort) & "' AS AgencyName, " & _
            " PicturePath, EndOfContract, EmploymentType, DeptCode, BrCode, PositionCode, JobCode, JobClass, Paymentbasis, CommissionRate, MonthlyRate, DailyRate, HourlyRate, StartContract, EndContract, NoOfMonthsHired, WithEcolaPay, WithMiscPay, WithFundPay, EcolaAmount, WithMiscPay, WithFundPay, CBamount, CBpayment," & _
            " CBstartpayment , CBendpayment, MiscFee, MFAmount, MFpayment, MFStartPayment, MFEndPayment, CashAdvance, PFAmount, PFPayment, PFStartPayment, PFEndPayment, Remarks, USERID, LastUpdate" & _
            " From HREmployees where EmployeeID = " & txtID.Text & _
            " Union" & _
            " SELECT EmployeeID, EmpNo, AgencyCode, ClientCode, ClientName, FirstName, Middlename, Surname, NickName, EmpPosition, Age, ApplicationDate, EmpStatus, BirthDate, Address1, HouseNo, Street, Barangay, District, City, Address2, Town, Province1, Province, Contact, Cphone, CAddress, Gender, Weight, Height, ExamGrade, InterviewGrade, TelNum1, TelNum2, TaxCode, DateHired, DateResign, SSS, Philhealth, PHCode, Pagibig, EmploymentStatus, AcctNo, Citizen, Birthplace, HairColor, EyeColor, Religion, DateRegular, DateProb, TIN, MedicalGrade, [Primary], DatePrimary, Secondary, DateSecondary, College, DateCollege, Degree, Major, PostGraduate, DatePostGraduate, PostGraduateCourse, FatherName, Faddress, FWork, MotherName, Maddress, MWork, NBI, ResCer, IssueAt, IssueOn, Supervisor, ReasonRes, Dept_No, Confi, " & _
            " AgencyName, " & _
            " PicturePath, EndOfContract, EmploymentType, DeptCode, BrCode, PositionCode, JobCode, JobClass, Paymentbasis, CommissionRate, MonthlyRate, DailyRate, HourlyRate, StartContract, EndContract, NoOfMonthsHired, WithEcolaPay, WithMiscPay, WithFundPay, EcolaAmount, WithMiscPay, WithFundPay, CBamount, CBpayment, " & _
            " CBstartpayment , CBendpayment, MiscFee, MFAmount, MFpayment, MFStartPayment, MFEndPayment, CashAdvance, PFAmount, PFPayment, PFStartPayment, PFEndPayment, Remarks, USERID, LastUpdate " & _
            " FROM HREmployees_History where EmployeeID = " & txtID.Text

        'Set rsItems = ReturnRst("Select * from qryHRhiredREPORT where EmpNo='" & txtControlNo.Text & "'", Cnnstr, mstatic)
        Set rsItems = ReturnRst(strHistory, cnnstr, mstatic)
        If rsItems.EOF = True Then MsgBox "Record not found! Please save first.", vbInformation: Exit Sub
         
        Dim tmpimg As RptImage
        
        Set tmpimg = rptReportSolo.Sections(2).Controls("ImgLogo")
        
        On Error GoTo b:
        'Set tmpimg.Picture = LoadPicture(App.Path & "..\Logos\" & CompanyCode & " logo.gif")
        If LoadPictureFlag = True Then
            Set tmpimg.Picture = LoadPicture(logoPath)
        End If
b:
        With rptReportSolo
        Set .DataSource = rsItems
             .DataMember = rsItems.DataMember
            .Sections("Section2").Controls("lblAddress").Caption = getAgencyAddress(CompanyCode) ' "Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
            .Sections("Section2").Controls("lblEmployeeID").Caption = rsItems!EmpNo
            .Sections("Section2").Controls("lblCompleteName").Caption = RTrim(Me.txtsurname.Text) & ", " & RTrim(Me.txtfirstname.Text) & " " & RTrim(Me.txtmiddlename.Text)
            '.Sections("Section2").Controls("lblClientName").Caption = Me.cmbCompany.Text
            '.Sections("Section2").Controls("lblDesignation").Caption = Me.txtpos1.Text
            '.Sections("Section2").Controls("lblHiringDate").Caption = IIf(IsNull(rsItems!datehired) = True, "", rsItems!datehired)
            .Sections("Section2").Controls("lblBirthplace").Caption = IIf(IsNull(rsItems!Birthplace) = True, "", RTrim(rsItems!Birthplace))
            .Sections("Section2").Controls("lblLandline").Caption = IIf(IsNull(rsItems!TelNum1) = True, "", RTrim(rsItems!TelNum1))
            .Sections("Section2").Controls("lblCellPhone").Caption = IIf(IsNull(rsItems!Cphone) = True, "", RTrim(rsItems!Cphone))
            
            '.Sections("Section2").Controls("lblNameOfSpouse").Caption = rsItems!EmpNo
            .Sections("Section2").Controls("lblMailingAddress").Caption = IIf(IsNull(rsItems!HouseNo) = True, "", RTrim(rsItems!HouseNo)) & " " & IIf(IsNull(rsItems!Street) = True, "", RTrim(rsItems!Street)) & " " & IIf(IsNull(rsItems!Barangay) = True, "", RTrim(rsItems!Barangay)) & " " & IIf(IsNull(rsItems!District) = True, "", RTrim(rsItems!District)) & " " & IIf(IsNull(rsItems!city) = True, "", RTrim(rsItems!city))
            .Sections("Section2").Controls("lblPhilhealth").Caption = IIf(IsNull(rsItems!PhilHealth) = True, "", rsItems!PhilHealth)
            .Sections("Section2").Controls("lblTIN").Caption = IIf(IsNull(rsItems!TIN) = True, "", rsItems!TIN)
            .Sections("Section2").Controls("lblSSS").Caption = IIf(IsNull(rsItems!sss) = True, "", rsItems!sss)
            .Sections("Section2").Controls("lblPagibig").Caption = IIf(IsNull(rsItems!Pagibig) = True, "", rsItems!Pagibig)
            
            Dim age As String
             If mskbday.Text = "__/__/____" Or IsNull(rsItems!BIRTHDATE) = True Then
                 age = ""
            Else
                age = DateDiff("yyyy", Format(mskbday, "mm/dd/yyyy"), Date)
            End If
            
       
            .Sections("Section2").Controls("lblAge").Caption = age
           
            .Sections("Section2").Controls("lblHeight").Caption = IIf(IsNull(rsItems!Height) = True, "", rsItems!Height)
            .Sections("Section2").Controls("lblReligion").Caption = IIf(IsNull(rsItems!Religion) = True, "", rsItems!Religion)
            .Sections("Section2").Controls("lblCivilStatus").Caption = IIf(IsNull(rsItems!TaxCode) = True, "", rsItems!TaxCode)
            
            .Show
           
        End With
End Sub

'Private Sub EmpPicture_Click()
'
'MsgBox "Get New Picture", vbInformation, PayrollAlert
'End Sub

Private Sub EmpPicture_DblClick()
If addflagship = True Or editflagship = True Then
    CommonDialog1.Filter = "Image Files|*.jpg|"
    CommonDialog1.ShowOpen
    PictureFile = CommonDialog1.Filename
    EmpPicture.Picture = LoadPicture(PictureFile)
End If
End Sub

'Private Sub EmpPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.MousePointer = vbHourglass
'End Sub

Private Sub Form_Activate()
    'If AccountType = "Accounting" Or AccountType = "Administrator" Or AccountType = "FMS Payroll" Or AccountType = "FMS Admin" Then
    If ForPayroll(AccountType) = True Then
        Me.Frame5(0).Visible = True
        Frame16.Visible = False
    Else
        Me.Frame5(0).Visible = False
        Frame16.Visible = True
    End If
End Sub

Private Sub Form_Load()

    PictureFile = App.Path & "\pics\noPhoto.jpg"
    
    'If AccountType = "Accounting" Or AccountType = "Administrator" Or AccountType = "FMS Payroll" Then
    If ForPayroll(AccountType) = True Then
        Me.Frame5(0).Visible = True
        Frame16.Visible = False
    Else
        Me.Frame5(0).Visible = False
        Frame16.Visible = True
    End If

    Call TextEnabled(False)
    
    On Error Resume Next
    SSTab1.Tab = 1
    FeMP = True

    Me.Left = 0
    Me.Top = 0

    Call OpenReferenceTable ' open reference table
    
'    StrQuery = "SELECT * from TaxExemptioncode"
'    Set rsCivil = ReturnRst(StrQuery, cnnstr, mstatic)
'
'    'Set rscivil = cnnHR.Execute("Select * from TaxExemptioncode")
'    Do While Not rsCivil.EOF
'        cmbcivilstatus.AddItem rsCivil!TaxExemptioncode
'        rsCivil.MoveNext
'    Loop
'    rsCivil.Close
'
'    StrQuery = "Select * from HRemploymentStatus"
'    Set rsEmpStatus = ReturnRst(StrQuery, cnnstr, mstatic)
'
'    'Set rsempstatus = cnnHR.Execute("Select * from HRemploymentstatus")
'    Do While Not rsEmpStatus.EOF
'        cmbEmpStatus.AddItem UCase(Mid(rsEmpStatus!employmentstatus, 1, 1) & "-" & rsEmpStatus!EmploymentStatusdesc)
'        rsEmpStatus.MoveNext
'    Loop
'    rsEmpStatus.Close
'
'    Call LoadCompany(CompanyCode)
'
'    StrQuery = "SELECT * FROM HRWorkStation ORDER BY brname"
'    Set rstBranch = ReturnRst(StrQuery, cnnstr, mstatic)
'
'    'Set rstBranch = ReturnRst("SELECT * FROM HRBRANCH WHERE BrStatus = True ORDER BY brname", CnnstrHR)
'    'Set rstBranch = cnnHR.Execute("SELECT * from HRBRANCH order by brname")
'
'    Do While Not rstBranch.EOF
'        cmbBranch.AddItem UCase(IIf(IsNull(rstBranch!BrName), "", rstBranch!BrName))
'        cmbBranch.ItemData(cmbBranch.NewIndex) = rstBranch!BrCode
'        rstBranch.MoveNext
'    Loop
'    rstBranch.Close
'
'    'Set rstDept = cnnHR.Execute("select * from HRDEPT order by deptname")
'    StrQuery = "select * from HRDEPT WHERE DeptStatus = 1 order by deptname"
'    Set rstDept = ReturnRst(StrQuery, cnnstr, mstatic)
'    'Set rstDept = ReturnRst("select * from HRDEPT order by deptname", cnnstr) ' updated by Bernard 10/17/08
'
'    Do While Not rstDept.EOF
'        cmbDept.AddItem UCase(rstDept!deptname)
'        cmbDept.ItemData(cmbDept.NewIndex) = rstDept!DeptCode
'        rstDept.MoveNext
'    Loop
'        rstDept.Close
'
'    StrQuery = "Select * from HRPositions order by positionname"
'    Set rstDept = ReturnRst(StrQuery, cnnstr, mstatic)
'
'    Do While Not rstDept.EOF
'        cmbPosition.AddItem UCase(rstDept!PositionName)
'        cmbPosition.ItemData(cmbPosition.NewIndex) = rstDept!PositionCode
'        rstDept.MoveNext
'    Loop
'        rstDept.Close
'
'    cmbShiftSched.Clear
'    StrQuery = "Select * From HRShiftSchedule WHERE ShiftStatus = 1 Order by ShiftID"
'    Set rstDept = ReturnRst(StrQuery, cnnstr, mstatic)
'    Do While Not rstDept.EOF
'        cmbShiftSched.AddItem rstDept!ShiftCode
'        rstDept.MoveNext
'    Loop
'    rstDept.Close
'    cmbShiftSched.ListIndex = 0
'
'    Set rsHREmp = New ADODB.Recordset
'    Set rsDetails = New ADODB.Recordset
'    Set rsRef = New ADODB.Recordset
'    Set rsDialect = New ADODB.Recordset
'    Set rsEmpRec = New ADODB.Recordset
'    Set rsem = New ADODB.Recordset
'
'    If CurrentClientCode = 0 Then
'        StrQuery = "SELECT * FROM HREmployees WHERE agencycode = " & CompanyCode & "  order by  surname, firstname"
'    Else
'        StrQuery = "SELECT * FROM HREmployees WHERE agencycode = " & CompanyCode & " AND ClientCode = " & CurrentClientCode & "  order by  surname, firstname"
'    End If
'
'    'StrQuery = "Select * from HREmployees where agencycode = " & CompanyCode & "  order by  surname"
'    Set rsHREmp = ReturnRst(StrQuery, cnnstr, mstatic)
'    'rsHREmp.Open "select * from HREmployees where agencycode=" & CompanyCode & "  order by  surname", cnnHR, adOpenKeyset, adLockOptimistic
'    StrQuery = "Select * from HRDetails WHERE DetailsFlag = 1"
'    Set rsDetails = ReturnRst(StrQuery, cnnstr, mstatic)
'    'rsdetails.Open "Select * from HRDetails", cnnHR, adOpenKeyset, adLockOptimistic
'    StrQuery = "Select * from HRCharReference WHERE ReferenceFlag = 1"
'    Set rsRef = ReturnRst(StrQuery, cnnstr, mstatic)
'    'rsref.Open "Select * from HRCharReference", cnnHR, adOpenKeyset, adLockOptimistic
'    StrQuery = "Select * from HRDialects WHERE DialectFlag = 1"
'    Set rsDialect = ReturnRst(StrQuery, cnnstr, mstatic)
'    'rsdialect.Open "Select * from HRDialects", cnnHR, adOpenKeyset, adLockOptimistic
'    StrQuery = "Select * from HREmploymentRecord WHERE EmploymentFlag = 1"
'    Set rsEmpRec = ReturnRst(StrQuery, cnnstr, mstatic)
'    'rsemprec.Open "Select * from HREmploymentRecord", cnnHR, adOpenKeyset, adLockOptimistic
'    StrQuery = "Select * from HRSeminar WHERE SeminarFlag = 1"
'    Set rsSem = ReturnRst(StrQuery, cnnstr, mstatic)
'    'rssem.Open "select * from HRSeminar", cnnHR, adOpenKeyset, adLockOptimistic
    
    If rsHREmp.RecordCount > 0 Then
        cmdundo.Enabled = False
        cmdupdate.Enabled = False
    Else
       Call ClearTextBox
       cmdRecontract.Enabled = False
       cmdfirst.Enabled = False
       cmdPrev.Enabled = False
       cmdNext.Enabled = False
       cmdlast.Enabled = False
       cmdedit.Enabled = False
       cmddelete.Enabled = False
       cmdundo.Enabled = False
       cmdupdate.Enabled = False
    End If
    EmpNo = TxtControlNo.Text
    'cmbUniform.Text = "Cash/Adv"
    'cmbAllowance.Text = ""
    
    'Dim strQuery As String
       ' If AddFlag = True Then
       '     Call cmdnew_Click
       ' End If
      '  If EditFlag = True Then
       '      rsHREmp.Find "EmployeeID =" & frmHRMainEmployee.TxtEmployeeID.Text
        '     Call cmdEdit_Click
         '       Showdata rsHREmp
          '  End If
 
'    Set rstclient = Nothing
'    Set rstBranch = Nothing
'    Set rstDept = Nothing
'    Set rsDetails = Nothing
'    Set rsRef = Nothing
'    Set rsDialect = Nothing
'    Set rsEmpRec = Nothing
'    Set rsCivil = Nothing
'    Set rsSem = Nothing
'    Set rsEmpStatus = Nothing

End Sub

Private Sub OpenReferenceTable()

    Set rstclient = New ADODB.Recordset
    Set rsDetails = New ADODB.Recordset
    Set rsRef = New ADODB.Recordset
    Set rsDialect = New ADODB.Recordset
    Set rsEmpRec = New ADODB.Recordset
    Set rsCivil = New ADODB.Recordset
    Set rsSem = New ADODB.Recordset
    Set rstPos = New ADODB.Recordset
    Set rsEmpStatus = New ADODB.Recordset
    Set rstShift = New ADODB.Recordset

    StrQuery = "SELECT * from TaxExemptioncode"
    Set rsCivil = ReturnRst(StrQuery, cnnstr, mstatic)
    cmbcivilstatus.Clear
    Do While Not rsCivil.EOF
        cmbcivilstatus.AddItem rsCivil!TaxExemptionCode
        rsCivil.MoveNext
    Loop
    
    StrQuery = "Select * from HRemploymentStatus"
    Set rsEmpStatus = ReturnRst(StrQuery, cnnstr, mstatic)
    cmbEmpStatus.Clear
    Do While Not rsEmpStatus.EOF
        cmbEmpStatus.AddItem UCase(mID(rsEmpStatus!employmentstatus, 1, 1) & "-" & rsEmpStatus!EmploymentStatusDesc)
        rsEmpStatus.MoveNext
    Loop

    'Call LoadCompany(CompanyCode)
    cmbCompany.Clear
    cmbChangeCompany.Clear
    
'    If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'        StrQuery = "SELECT * FROM clienttable where CLIENTCODE2 = 'FMS' AND clientstatus = 1 order by clientname "
'    Else
'        StrQuery = "SELECT * FROM clienttable where clientstatus = 1 order by clientname "
'    End If
    StrQuery = GetClientRecord(0, True, " ORDER BY ClientName")
    'StrQuery = "SELECT * FROM clienttable where clientstatus = 1 order by clientname "
    Set rstclient = ReturnRst(StrQuery, cnnstr, mstatic)
    
    CurrentClientCode = 0
    Do While Not rstclient.EOF
        cmbCompany.AddItem rstclient!ClientName
        cmbCompany.ItemData(cmbCompany.NewIndex) = rstclient!ClientCode
        cmbChangeCompany.AddItem rstclient!ClientName
        cmbChangeCompany.ItemData(cmbChangeCompany.NewIndex) = rstclient!ClientCode
        rstclient.MoveNext
    Loop
    cmbChangeCompany.ListIndex = -1
    
    StrQuery = "SELECT * FROM HRWorkStation ORDER BY brName"
    Set rstBranch = ReturnRst(StrQuery, cnnstr, mstatic)
    cmbBranch.Clear
    Do While Not rstBranch.EOF
        cmbBranch.AddItem UCase(IIf(IsNull(rstBranch!BrName), "", rstBranch!BrName))
        cmbBranch.ItemData(cmbBranch.NewIndex) = rstBranch!BrCode
        rstBranch.MoveNext
    Loop
        
    StrQuery = "SELECT * FROM HRDEPT WHERE DeptStatus = 1 order by deptname"
    Set rstDept = ReturnRst(StrQuery, cnnstr, mstatic)
    cmbDept.Clear
    Do While Not rstDept.EOF
        cmbDept.AddItem UCase(rstDept!deptname)
        cmbDept.ItemData(cmbDept.NewIndex) = rstDept!DeptCode
        rstDept.MoveNext
    Loop
        
    StrQuery = "SELECT * FROM HRPositions order by PositionName"
    Set rstPos = ReturnRst(StrQuery, cnnstr, mstatic)
    cmbPosition.Clear
    Do While Not rstPos.EOF
        cmbPosition.AddItem UCase(rstPos!Positionname)
        cmbPosition.ItemData(cmbPosition.NewIndex) = rstPos!Positioncode
        rstPos.MoveNext
    Loop
        
    StrQuery = "SELECT * FROM HRShiftSchedule WHERE ShiftStatus = 1 Order by ShiftShort"
    Set rstShift = ReturnRst(StrQuery, cnnstr, mstatic)
    cmbShiftSchedule.Clear
    cmbShiftSchedule.AddItem ""
    Do While Not rstShift.EOF
        cmbShiftSchedule.AddItem rstShift!ShiftName
        cmbShiftSchedule.ItemData(cmbShiftSchedule.NewIndex) = rstShift!ShiftCode
        rstShift.MoveNext
    Loop
        
    If CurrentClientCode = 0 Then
        StrQuery = "SELECT * FROM HREmployees WHERE agencycode = " & CompanyCode & "  order by  surname, firstname"
    Else
        StrQuery = "SELECT * FROM HREmployees WHERE agencycode = " & CompanyCode & " AND ClientCode = " & CurrentClientCode & "  order by  surname, firstname"
    End If
    Set rsHREmp = ReturnRst(StrQuery, cnnstr, mstatic)
    
    StrQuery = "Select * from HRDetails WHERE DetailsFlag = 1"
    Set rsDetails = ReturnRst(StrQuery, cnnstr, mstatic)
    
    StrQuery = "Select * from HRCharReference WHERE ReferenceFlag = 1"
    Set rsRef = ReturnRst(StrQuery, cnnstr, mstatic)
    
    StrQuery = "Select * from HRDialects WHERE DialectFlag = 1"
    Set rsDialect = ReturnRst(StrQuery, cnnstr, mstatic)
    
    StrQuery = "Select * from HREmploymentRecord WHERE EmploymentFlag = 1"
    Set rsEmpRec = ReturnRst(StrQuery, cnnstr, mstatic)
    
    StrQuery = "Select * from HRSeminar WHERE SeminarFlag = 1"
    Set rsSem = ReturnRst(StrQuery, cnnstr, mstatic)
   
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

Public Sub LoadCompany(ByVal CompanyCode As Integer)

Me.cmbCompany.Clear
Me.cmbChangeCompany.Clear

'If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'    StrQuery = "SELECT * FROM clienttable where CLIENTCODE2 = 'FMS' AND clientstatus = 1 order by clientname "
'Else
'    StrQuery = "SELECT * FROM clienttable where clientstatus = 1 order by clientname "
'End If
StrQuery = GetClientRecord(0, True, " ORDER BY ClientName")
Set rstclient = ReturnRst(StrQuery, cnnstr, mstatic)
CurrentClientCode = 0
  
Do While Not rstclient.EOF
    cmbCompany.AddItem rstclient!ClientName
    cmbCompany.ItemData(cmbCompany.NewIndex) = rstclient!ClientCode
    cmbChangeCompany.AddItem rstclient!ClientName
    cmbChangeCompany.ItemData(cmbChangeCompany.NewIndex) = rstclient!ClientCode
    rstclient.MoveNext
Loop
rstclient.Close
cmbChangeCompany.ListIndex = -1
    
End Sub

Public Sub Showdata(ByRef rec As ADODB.Recordset)
Dim fs As FileSystemObject
Dim mText1 As String

    On Error Resume Next
    
    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    
    If rec.EOF = True Then
        Call ClearTextBox
        cmdedit.Enabled = False
        cmddelete.Enabled = False
    Else
        'cmbAllowance.Text = ""
        txtClientCode.Text = rec!ClientCode
        txtID.Text = rec!EmployeeID
        TxtControlNo.Text = IIf(IsNull(rec!EmpNo), "<control no>", rec!EmpNo)
        txtDTRID.Text = IIf(IsNull(rec!DTRID), "00000000", rec!DTRID)
        txtfirstchoice.Text = IIf(IsNull(rec!empPosition), "", rec!empPosition)
        txtnickname.Text = IIf(IsNull(rec!nickname), "", rec!nickname)
        txtsurname.Text = IIf(IsNull(rec!Surname), "", rec!Surname)
        txtfirstname.Text = IIf(IsNull(rec!FirstName), "", rec!FirstName)
     
        txtLastDate.Text = IIf(IsNull(rec!LastUpdate), "__/__/____", Format(rec!LastUpdate, "mm/dd/yyyy hh:mm:ss AMPM"))
        
        txtmiddlename.Text = IIf(IsNull(rec!MiddleName), "", rec!MiddleName)
        TmpMiddlename = txtmiddlename.Text
        TmpSurname = Me.txtsurname.Text
        TmpFirstname = Me.txtfirstname.Text
    
        txtaddress1.Text = IIf(IsNull(rec!Address1), "", rec!Address1)
        txthouseno.Text = IIf(IsNull(rec!HouseNo), "", rec!HouseNo)
        txtstreet.Text = IIf(IsNull(rec!Street), "", rec!Street)
        txtbarangay.Text = IIf(IsNull(rec!Barangay), "", rec!Barangay)
        txtdistrict.Text = IIf(IsNull(rec!District), "", rec!District)
        txtprovince.Text = IIf(IsNull(rec!Province1), "", rec!Province1)
        txtcity.Text = IIf(IsNull(rec!city), "", rec!city)
        txtprovinceaddress.Text = IIf(IsNull(rec!Province), "", rec!Province)
       
        mskbday.Text = IIf(IsNull(rec!BIRTHDATE) = True Or rec!BIRTHDATE = "", "__/__/____", Format(rec!BIRTHDATE, "mm/dd/yyyy"))
        If mskbday.Text = "__/__/____" Or IsNull(rec!BIRTHDATE) = True Then
            txtage.Text = ""
        Else
            txtage.Text = DateDiff("yyyy", Format(mskbday, "mm/dd/yyyy"), Date)
            ' Update the age.
            If rec!age <> txtage.Text Then
                sqlString = "Update HREmployees set age=" & txtage.Text & "" & _
                          " WHERE  EmployeeID = " & txtID.Text & ""
                Call Action_Query(sqlString, cnnstr)
            End If
        End If
        
        txtemergency.Text = IIf(IsNull(rec!Contact), "", rec!Contact)
        txteraddress.Text = IIf(IsNull(rec!CAddress), "", rec!CAddress)
        txtertel.Text = IIf(IsNull(rec!Cphone), "", rec!Cphone)
        Me.Caption = "EMPLOYEES INFORMATION [Updated by: " + UCase(GETuSERNAME(IIf(IsNull(rec!USERID), "0", rec!USERID))) + "]"
        
        If LTrim(RTrim(rec!Gender)) = "M" Then
            optmale.VALUE = 1
            optfemale.VALUE = 0
        ElseIf LTrim(RTrim(rec!Gender)) = "F" Then
            optmale.VALUE = 0
            optfemale.VALUE = 1
        Else
            optmale.VALUE = 0 'False
            optfemale.VALUE = 0 'False
        End If
        txtheight.Text = IIf(IsNull(rec!Height), "", Replace(rec!Height, """", "'"))
        txtweight.Text = IIf(IsNull(rec!Weight), "", rec!Weight)
        txttel1.Text = IIf(IsNull(rec!TelNum1), "", rec!TelNum1)
        txttel2.Text = IIf(IsNull(rec!TelNum2), "", rec!TelNum2)
        txtexam.Text = IIf(IsNull(rec!examgrade), "", rec!examgrade)
        txtinterview.Text = IIf(IsNull(rec!interviewgrade), "", rec!interviewgrade)
        txtmedical.Text = IIf(IsNull(rec!medicalgrade), "", rec!medicalgrade)
        msksss.Text = IIf(IsNull(rec!sss), "__-_______-_", rec!sss)
        txtphilhealth = IIf(IsNull(rec!PhilHealth), "", rec!PhilHealth)
        txtpagibig.Text = IIf(IsNull(rec!Pagibig), "", rec!Pagibig)
        txtacctno.Text = IIf(IsNull(rec!AcctNo), "", rec!AcctNo)
        
        txtbirthplace.Text = IIf(IsNull(rec!Birthplace), "", rec!Birthplace)
        txthaircolor.Text = IIf(IsNull(rec!haircolor), "", rec!haircolor)
        txteyecolor.Text = IIf(IsNull(rec!eyecolor), "", rec!eyecolor)
        txtreligion.Text = IIf(IsNull(rec!Religion), "", rec!Religion)
             
        msktax.Text = IIf(IsNull(rec!TIN), "___-___-___", rec!TIN)
        txtprimarysch.Text = IIf(IsNull(rec![Primary]), "", rec![Primary])
        txtdateprimary.Text = IIf(IsNull(rec!dateprimary), "", rec!dateprimary)
        txtsecsch.Text = IIf(IsNull(rec!Secondary), "", rec!Secondary)
        txtdatesec.Text = IIf(IsNull(rec!DateSecondary), "", rec!DateSecondary)
        txtcollegesch.Text = IIf(IsNull(rec!college), "", rec!college)
        txtdatecollege.Text = IIf(IsNull(rec!datecollege), "", rec!datecollege)
        txtcitizen.Text = IIf(IsNull(rec!citizen), "", rec!citizen)
        txtcollegecourse.Text = IIf(IsNull(rec!Degree), "", rec!Degree)
        txtcollegemajor.Text = IIf(IsNull(rec!Major), "", rec!Major)
        txtpostsch.Text = IIf(IsNull(rec!postgraduate), "", rec!postgraduate)
        txtdatepost.Text = IIf(IsNull(rec!datepostgraduate), "", rec!datepostgraduate)
        txtpostcourse.Text = IIf(IsNull(rec!postgraduatecourse), "", rec!postgraduatecourse)
        txtfathername.Text = IIf(IsNull(rec!fathername), "", rec!fathername)
        txtfaddress.Text = IIf(IsNull(rec!faddress), "", rec!faddress)
        txtfjob.Text = IIf(IsNull(rec!fwork), "", rec!fwork)
        txtmothername.Text = IIf(IsNull(rec!mothername), "", rec!mothername)
        txtmaddress.Text = IIf(IsNull(rec!maddress), "", rec!maddress)
        txtmjob.Text = IIf(IsNull(rec!mwork), "", rec!mwork)
        txtnbi.Text = IIf(IsNull(rec!NBI), "", rec!NBI)
        
        txtrescer.Text = IIf(IsNull(rec!ResCer), "", rec!ResCer)
        txtissueat.Text = IIf(IsNull(rec!issueat), "", rec!issueat)
        mskissue.Text = IIf(IsNull(rec!IssueOn), "__/__/____", Format(rec!IssueOn, "mm/dd/yyyy"))
        
        mskdateprob.Text = IIf(IsNull(rec!dateprob), "__/__/____", Format(rec!dateprob, "mm/dd/yyyy"))
        MskDateHired.Text = IIf(IsNull(rec!datehired), "__/__/____", Format(rec!datehired, "mm/dd/yyyy"))
        mskdatereg.Text = IIf(IsNull(rec!dateregular), "__/__/____", Format(rec!dateregular, "mm/dd/yyyy"))
        mskdateres.Text = IIf(IsNull(rec!dateresign), "__/__/____", Format(rec!dateresign, "mm/dd/yyyy"))
        txtsupervisor.Text = IIf(IsNull(rec!Supervisor), "", rec!Supervisor)
        txtreasonresign.Text = IIf(IsNull(rec!reasonres), "", rec!reasonres)
        Me.mskendof.Text = IIf(IsNull(rec!EndOfContract) = True Or rec!EndOfContract = "", "__/__/____", Format(rec!EndOfContract, "mm/dd/yyyy"))
    
        txtMonthlyRate = Format(rec!MonthlyRate, "##0.000")
        If rec!PaymentBasis = "Hourly" Then
            cmbPayBasis.ListIndex = 0
            txtRate = IIf(IsNull(rec!HourlyRate), "", Format(rec!HourlyRate, "##0.000"))
        ElseIf rec!PaymentBasis = "Daily" Then
            cmbPayBasis.ListIndex = 1
            txtRate = IIf(IsNull(rec!DailyRate), "", Format(rec!DailyRate, "##0.00"))
        ElseIf rec!PaymentBasis = "Monthly" Then
            cmbPayBasis.ListIndex = 2
            txtRate = IIf(IsNull(rec!MonthlyRate), "", Format(rec!MonthlyRate, "##0.00"))
        ElseIf rec!PaymentBasis = "Commission" Then
            cmbPayBasis.ListIndex = 3
            txtRate = IIf(IsNull(rec!CommissionRate), "", Format(rec!CommissionRate, "##0.00"))
        Else
            cmbPayBasis.ListIndex = -1
            txtRate = ""
       End If
       
       mskStartContract.Text = IIf(IsNull(rec!StartContract) = True Or rec!StartContract = "__/__/____", "__/__/____", Format(rec!StartContract, "mm/dd/yyyy"))
       mskEndContract.Text = IIf(IsNull(rec!endcontract) = True Or rec!endcontract = "__/__/____", "__/__/____", Format(rec!endcontract, "mm/dd/yyyy"))
       txtMonthsHired.Text = IIf(IsNull(rec!NoofMonthsHired) = True, "", rec!NoofMonthsHired)
       
       txtCBtotal.Text = IIf(IsNull(rec!Cbamount) = True, "", Format(rec!Cbamount, "##0.00"))
       txtCBpayment.Text = IIf(IsNull(rec!CBPayment) = True, "", Format(rec!CBPayment, "##0.00"))
       mskStartCB.Text = IIf(IsNull(rec!CBStartPayment) = True Or rec!CBStartPayment = "__/__/____", "__/__/____", Format(rec!CBStartPayment, "mm/dd/yyyy"))
       mskEndCB.Text = IIf(IsNull(rec!CBEndPayment) = True Or rec!CBEndPayment = "__/__/____", "__/__/____", Format(rec!CBEndPayment, "mm/dd/yyyy"))
       
       mskStartCash.Text = IIf(IsNull(rec!CaStartPayment) = True Or rec!CaStartPayment = "__/__/____", "__/__/____", Format(rec!CaStartPayment, "mm/dd/yyyy"))
       mskEndCash.Text = IIf(IsNull(rec!CaEndPayment) = True Or rec!CaEndPayment = "__/__/____", "__/__/____", Format(rec!CaEndPayment, "mm/dd/yyyy"))
       mskStartUniform.Text = IIf(IsNull(rec!UNIStartPayment) = True Or rec!UNIStartPayment = "__/__/____", "__/__/____", Format(rec!UNIStartPayment, "mm/dd/yyyy"))
       mskEndUniform.Text = IIf(IsNull(rec!UNIEndPayment) = True Or rec!UNIEndPayment = "__/__/____", "__/__/____", Format(rec!UNIEndPayment, "mm/dd/yyyy"))

       'txtMiscFee.Text = IIf(IsNull(rec!MiscFee), "", Format(rec!MiscFee, "##0.00"))
       'txtPlacementFee.Text = IIf(IsNull(rec!CashAdvance), "", Format(rec!CashAdvance, "##0.00"))
       
       txtBillingRate = IIf(IsNull(rec!MiscFee), "0.00", Format(rec!MiscFee, "##0.00")) ' Uniform
       TxtEcolaAmount.Text = IIf(IsNull(rec!EcolaAmount), "0.00", Format(rec!EcolaAmount, "##0.00")) ' dating Cash Advance
        
       TxtMFAmount = IIf(IsNull(rec!MFAmount), "0.00", Format(rec!MFAmount, "##0.00"))  ' mfAmount
       TxtMFPayment = IIf(IsNull(rec!MFpayment), "0.00", Format(rec!MFpayment, "##0.00"))
       MskMFStartPayment.Text = IIf(IsNull(rec!MFstartpayment) = True Or rec!MFstartpayment = "__/__/____", "__/__/____", Format(rec!MFstartpayment, "mm/dd/yyyy"))
       MskMFEndPayment.Text = IIf(IsNull(rec!MFendpayment) = True Or rec!MFendpayment = "__/__/____", "__/__/____", Format(rec!MFendpayment, "mm/dd/yyyy"))
       
       txtPFAmount.Text = IIf(IsNull(rec!PFAmount), "0.00", Format(rec!PFAmount, "##0.00")) ' pfamount
       txtPFPayment = IIf(IsNull(rec!PFpayment), "0.00", Format(rec!PFpayment, "##0.00"))
       MskPFStartPayment.Text = IIf(IsNull(rec!PFstartpayment) = True Or rec!PFstartpayment = "__/__/____", "__/__/____", Format(rec!PFstartpayment, "mm/dd/yyyy"))
       MskPFEndPayment.Text = IIf(IsNull(rec!PFendpayment) = True Or rec!PFendpayment = "__/__/____", "__/__/____", Format(rec!PFendpayment, "mm/dd/yyyy"))
       
       txtSuretyPercent = IIf(IsNull(rec!FundPercentage), "0.00", Format(rec!FundPercentage, "##0.00"))
        
        If IsNull(rec!Town) Then
            cmbAllowance.ListIndex = 0
            cmbAllowance.Text = ""
        ElseIf rec!Town = "1st Cutoff" Then
            cmbAllowance.ListIndex = 1
            cmbAllowance.Text = "1st Cutoff"
        ElseIf rec!Town = "2nd Cutoff" Then
            cmbAllowance.ListIndex = 2
            cmbAllowance.Text = "2nd Cutoff"
        ElseIf rec!Town = "All Cutoff" Then
            cmbAllowance.ListIndex = 3
            cmbAllowance.Text = "All Cutoff"
        ElseIf rec!Town = "Per Day" Then
            cmbAllowance.ListIndex = 4
            cmbAllowance.Text = "Per Day"
        End If
        'cmbAllowance.Text = IIf(rec!Town = "", "", rec!Town) ' updated July 16, 2016
        
'       If rec!WithEcolaPay = 1 Or rec!WithEcolaPay = True Then
'            ChkEcola.VALUE = 1
'       Else
'            ChkEcola.VALUE = 0
'       End If
       
       If rec!WithMiscPay = 1 Then
            chkMisc.VALUE = 1
       Else
            chkMisc.VALUE = 0
       End If
       
       If rec!WithFundPay = 1 Then
            chkFund.VALUE = 1
       Else
            chkFund.VALUE = 0
       End If
       
       'Coop Dues.
       'TxtEcolaAmount.Text = IIf(IsNull(rec!EcolaAmount) = True, "", Format(rec!EcolaAmount, "##0.00"))
       
       'txtBillingRate.Text = IIf(IsNull(rec!BillingRate), "0.00", Format(rec!BillingRate, "##0.00"))
       'txtShiftSched.Text = IIf(IsNull(rec!ShiftName), "", rec!ShiftName)
        
        DoEvents
    
    End If
     
    'Shift Schedule
    mText1 = GetDescription(rstShift, "ShiftCode", rec!ShiftCode, "ShiftName", "", True)
'    txtShiftSched.Text = GetDescription(rstShift, "ShiftCode", rec!ShiftCode, "ShiftName")
'    cmbShiftSchedule.Text = txtShiftSched.Text
    If mText1 = "" Then
        cmbShiftSchedule.ListIndex = -1
        cmbShiftSchedule.Text = ""
    Else
        cmbShiftSchedule.Text = mText1
    End If

'    strQuery = "Select ShiftCode From HRShiftSchedule WHERE ShiftCode = '" & rec!ShiftCode & "'"
'    Set rs = ReturnRst(strQuery, cnnstr, mstatic)
'    If Not rs.EOF Then
'        cmbShiftSched.Text = rs.Fields(0)
'        txtShiftSched.Text = rs.Fields(0)
'    Else
'        cmbShiftSched.ListIndex = -1
'        'txtShiftSched.Text = ""
'    End If
'    rs.Close
    
    'TAX EXEMPTION
    cmbcivilstatus.Text = GetDescription(rsCivil, "TaxExemptionCode", IIf(IsNull(rec!TaxCode), "", rec!TaxCode), "TaxExemptionCode")
    If cmbcivilstatus.Text = "" Then
        cmbcivilstatus.ListIndex = -1
    End If
    
'    strQuery = "Select taxexemptioncode From taxexemptioncode Where taxexemptioncode ='" & IIf(IsNull(rec!TaxCode), "", rec!TaxCode) & "'"
'    Set rs = ReturnRst(strQuery, cnnstr, mstatic)
'    'Set rs = cnnHR.Execute("Select taxexemptioncode From taxexemptioncode Where taxexemptioncode ='" & IIf(IsNull(rec!TaxCode), "", rec!TaxCode) & "'")
'    If Not rs.EOF Then
'        cmbcivilstatus.Text = rs.Fields(0)
'    Else
'        cmbcivilstatus.ListIndex = -1
'    End If
'    rs.Close
     
     'CLIENT
'     strQuery = "Select * From clienttable Where clientcode = " & IIf(IsNull(rec!ClientCode), 0, rec!ClientCode)
'     Set rs = ReturnRst(strQuery, cnnstr, mstatic)
     'Set rs = ReturnRst("Select * From clienttable Where clientcode =" & IIf(IsNull(rec!ClientCode), 0, rec!ClientCode) & "", cnnstr)
     cmbCompany.Text = GetDescription(rstclient, "ClientCode", IIf(IsNull(rec!ClientCode), 1, rec!ClientCode), "ClientName")
            'If Not rs.EOF Then
            If cmbCompany.Text <> "" Then
                'cmbCompany.Text = rs.Fields("clientname")
                lblCompany.Caption = UCase(cmbCompany.Text)
                If SuperUser = False And rstclient.Fields("ClientFlag2") = True Then
                If txtRate.Visible = True Then txtRate.Visible = False
                    If txtMonthsHired.Visible = True Then txtMonthsHired.Visible = False
                    If cmbPayBasis.Visible = True Then cmbPayBasis.Visible = False
                Else
                    If txtRate.Visible = False Then txtRate.Visible = True
                    If txtMonthsHired.Visible = False Then txtMonthsHired.Visible = True
                    If Me.cmbPayBasis.Visible = False Then Me.cmbPayBasis.Visible = True
                End If
            Else
                cmbCompany.ListIndex = -1
            End If
'            rs.Close

     'HR BRANCH
    'Set rs = cnnHR.Execute("Select * From HRBRANCH Where brcode =" & IIf(IsNull(rec!BrCode), 0, rec!BrCode))
'    strQuery = "SELECT * FROM HRWorkStation WHERE brcode = " & IIf(IsNull(rec!BrCode), 0, rec!BrCode)
'    Set rs = ReturnRst(strQuery, cnnstr, mstatic)
    'Set rs = ReturnRst("SELECT * FROM HRBRANCH WHERE brcode = " & IIf(IsNull(rec!BrCode), 0, rec!BrCode), CnnstrHR)
    cmbBranch.Text = GetDescription(rstBranch, "BrCode", IIf(IsNull(rec!BrCode), 0, rec!BrCode), "BrName")
            'If Not rs.EOF Then
            If cmbBranch.Text = "" Then
'                cmbBranch.Text = rs.Fields(2)
'            Else
                cmbBranch.ListIndex = -1
            End If
'            rs.Close
            
      'HR DEPARTMENT
'      strQuery = "Select * From HRDept Where deptcode =" & IIf(IsNull(rec!DeptCode), 0, rec!DeptCode)
'      Set rs = ReturnRst(strQuery, cnnstr, mstatic)
      cmbDept.Text = GetDescription(rstDept, "DeptCode", IIf(IsNull(rec!DeptCode), 0, rec!DeptCode), "DeptName")
      'Set rs = cnnHR.Execute("Select * From HRdept Where deptcode =" & IIf(IsNull(rec!DeptCode), 0, rec!DeptCode))
           'If Not rs.EOF Then
           If cmbDept.Text = "" Then
'                cmbDept.Text = rs.Fields(1)
'            Else
                cmbDept.ListIndex = -1
            End If
'      rs.Close
      
'      strQuery = "Select * From HRPositions Where positioncode =" & IIf(IsNull(rec!PositionCode), 0, rec!PositionCode)
        cmbPosition.Text = GetDescription(rstPos, "PositionCode", IIf(IsNull(rec!Positioncode), 0, rec!Positioncode), "PositionName")
        ' Updated 1/21/2011 Bernard
        txtfirstchoice.Text = cmbPosition.Text

'      Set rs = ReturnRst(strQuery, cnnstr, mstatic)
           'If Not rs.EOF Then
           If cmbPosition.Text = "" Then
'                cmbPosition.Text = rs.Fields(1)
'            Else
                cmbPosition.ListIndex = -1
            End If
'      rs.Close
      
      'HR EMPLOYMENT STATUS
'      strQuery = "Select * From HREmploymentStatus Where Employmentstatus = '" & Trim(IIf(IsNull(rec!employmentstatus), "", rec!employmentstatus)) & "'"
'      'strQuery = "Select * From HREmploymentstatus Where employmentstatus = '" & Trim(IIf(IsNull(rec!employmentstatus), "", rec!employmentstatus)) & "'"
'      Set rs1 = ReturnRst(strQuery, cnnstr, mstatic)
      cmbEmpStatus.Text = GetDescription(rsEmpStatus, "Employmentstatus", Trim(IIf(IsNull(rec!employmentstatus), "", rec!employmentstatus)), "EmploymentstatusDesc", "EmploymentstatusDesc")
            'If Not rs1.EOF Then
            If cmbEmpStatus.Text = "" Then
'                cmbEmpStatus.Text = Mid(rs1.Fields(0), 1, 1) & "-" & rs1.Fields(1)
'            Else
                cmbEmpStatus.ListIndex = -1
            End If
'            rs1.Close
            
      'HR DIALECTS
'      strQuery = "Select * from HRDialects where EmployeeID = " & TxtID.Text
'      Set rsDialect = ReturnRst(strQuery, cnnstr, mstatic)
'      'Set rsdialect = cnnHR.Execute("Select * from HRDialects where EmployeeID='" & TxtControlNo.Text & "'")
      StrQuery = GetDescription(rsDialect, "EmployeeID", txtID.Text, "Dialect1")
         'If rsDialect.EOF = False Then
         If StrQuery <> "" Then
             Showdialect rsDialect
'             rsDialect.MoveNext
         Else
             txtdialect1 = ""
             txtdialect2 = ""
             txtdialect3 = ""
         End If
'        rsDialect.Close
        
      'HR DETAILS
'      strQuery = "Select * from HRDetails where EmployeeID = " & TxtID.Text
'      Set rsDetails = ReturnRst(strQuery, cnnstr, mstatic)
'      'Set rsdetails = cnnHR.Execute("Select * from HRDetails where EmployeeID='" & TxtControlNo.Text & "'")
        
      StrQuery = GetDescription(rsDetails, "EmployeeID", txtID.Text, "SpouseName", , , True)
        'If rsDetails.EOF = False Then
        If StrQuery <> "" Then
            Showdetails rsDetails
'            rsDetails.MoveNext
        Else
             txtspousename.Text = ""
             txtspouseaddress.Text = ""
             txtsjob.Text = ""
             txtnumdep.Text = ""
             TxtContactPerson.Text = ""
             txtContactAddress.Text = ""
             txtContactPhone.Text = ""
             
             txtdep1.Text = ""
             txtdep2.Text = ""
             txtdep3.Text = ""
             txtdep4.Text = ""
             mskdep1.Text = "__/__/____"
             mskdep2.Text = "__/__/____"
             mskdep3.Text = "__/__/____"
             mskdep4.Text = "__/__/____"
             mskdep5.Text = "__/__/____"
        End If
'        rsDetails.Close
    
      'HR CHARACTER REFERENCE.
'      strQuery = "Select * from HRCharReference where EmployeeID = " & TxtID.Text
'      Set rsRef = ReturnRst(strQuery, cnnstr, mstatic)
      'Set rsref = cnnHR.Execute("Select * from HRCharReference where EmployeeID='" & TxtControlNo.Text & "'")
         
      StrQuery = GetDescription(rsRef, "EmployeeID", txtID.Text, "Name")
         'If rsRef.EOF = False Then
         If StrQuery <> "" Then
            ShowRef rsRef
'            rsRef.MoveNext
        Else
            txtrefname1.Text = ""
            txtrefjob1.Text = ""
            txtrefaddress1.Text = ""
            txtrefphone1.Text = ""
            txtrefrel1.Text = ""
            txtrefname2.Text = ""
            txtrefjob2.Text = ""
            txtrefaddress2.Text = ""
            txtrefphone2.Text = ""
            txtrefrel2.Text = ""
            txtrefname3.Text = ""
            txtrefjob3.Text = ""
            txtrefaddress3.Text = ""
            txtrefphone3.Text = ""
            txtrefrel3.Text = ""
            End If
'        rsRef.Close

'        strQuery = "Select * from HREmploymentRecord where EmployeeID = " & TxtID.Text
'        Set rsEmpRec = ReturnRst(strQuery, cnnstr, mstatic)
'        'Set rsemprec = cnnHR.Execute("Select * from HREmploymentRecord where EmployeeID='" & TxtControlNo.Text & "'")
      
      StrQuery = GetDescription(rsEmpRec, "EmployeeID", txtID.Text, "Company")
        'If rsEmpRec.EOF = False Then
        If StrQuery <> "" Then
            ShowEmpRec rsEmpRec
'            rsEmpRec.MoveNext
        Else
            txtnameofcompany1.Text = ""
            txtfrom1.Text = ""
            txtto1.Text = ""
            txtpos1.Text = ""
            txtreason1.Text = ""
            txtsuperior1.Text = ""
            txtnameofcompany2.Text = ""
            txtfrom2.Text = ""
            txtto2.Text = ""
            txtpos2.Text = ""
            txtreason2.Text = ""
            txtsuperior2.Text = ""
            txtnameofcompany3.Text = ""
            txtfrom3.Text = ""
            txtto3.Text = ""
            txtpos3.Text = ""
            txtreason3.Text = ""
            txtsuperior3.Text = ""
            End If
'        rsEmpRec.Close
        
        'HR SEMINAR
'        strQuery = "Select * from HRSeminar where EmployeeID = " & TxtID.Text
'        Set rsSem = ReturnRst(strQuery, cnnstr, mstatic)
'        'Set rssem = cnnHR.Execute("Select * from HRSeminar where EmployeeID='" & TxtControlNo.Text & "'")
      
      StrQuery = GetDescription(rsSem, "EmployeeID", txtID.Text, "Seminar1")
        'If rsSem.EOF = False Then
        If StrQuery <> "" Then
            ShowSem rsSem
        Else
            txtseminar1.Text = ""
            txtdateseminar1.Text = ""
            txtseminar2.Text = ""
            txtdateseminar2.Text = ""
            txtseminar3.Text = ""
            txtdateseminar3.Text = ""
        End If
'
'        rsSem.Close

        ' Load the picture.
        If LoadPictureFlag = True Then
            If fs.FileExists(rec!picturepath) = True Then
                EmpPicture.Picture = LoadPicture(rec!picturepath)
            ElseIf fs.FileExists(App.Path & "\pics\noPhoto.jpg") = True Then
                EmpPicture.Picture = LoadPicture(App.Path & "\pics\noPhoto.jpg")
            End If
        End If
        
        txtGenInfo.Text = IIf(LTrim(RTrim(mID(cmbEmpStatus, 3))) = "", "No Status", LTrim(RTrim(mID(cmbEmpStatus, 3)))) & _
            "/ " & IIf(cmbPosition = "", "No Position", cmbPosition) & "/ " & IIf(cmbDept = "", "No dept", cmbDept) & _
            "/ " & IIf(txtaddress1 = "", "No Address", txtaddress1) & _
            "/ " & IIf(IsNull(rec!datehired), "No Date Hired", Format(rec!datehired, "mm/dd/yyyy"))
        
End Sub

Private Sub DisAble_ActionButtons()
    cmdNew.Enabled = False
    cmdedit.Enabled = False
    cmdupdate.Enabled = True
    cmdundo.Enabled = True
    cmddelete.Enabled = False
    cmdfirst.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdlast.Enabled = False
    cmdundo.Enabled = True
    btnEmpRefresh.Enabled = False
    cmdRecontract.Enabled = False
    btnEmployeeSearch.Enabled = False
End Sub

Private Sub cmbcivil_GotFocus()
    cmbcivilstatus.BackColor = &HFFFFC0
    Set oCommonFunction.ValidAlpha = cmbcivilstatus
End Sub

Private Sub cmbcivil_LostFocus()
    cmbcivilstatus.BackColor = vbWhite
End Sub

Public Sub btnEmpRefresh_Click()
On Error Resume Next

    If rsHREmp.State = 1 Then rsHREmp.Close
    'StrQuery = "Select * from HREmployees where agencycode = " & CompanyCode & "  order by surname, firstname"
    If CurrentClientCode = 0 Then
        StrQuery = "SELECT * FROM HREmployees WHERE agencycode = " & CompanyCode & "  order by  surname, firstname"
    Else
        StrQuery = "SELECT * FROM HREmployees WHERE agencycode = " & CompanyCode & " AND ClientCode = " & CurrentClientCode & "  order by  surname, firstname"
    End If
    Set rsHREmp = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsHREmp.Open "Select * from HREmployees where agencycode=" & CompanyCode & "  order by surname, firstname", cnnHR, adOpenKeyset, adLockOptimistic
    If rsHREmp.EOF = False Then
        Showdata rsHREmp
        Call OpenReferenceTable ' open reference table
        cmdNext.Enabled = True
        cmdlast.Enabled = True
        MsgBox "HR records was updated!", vbInformation, PayrollAlert
    Else
        MsgBox "All records deleted, the form will be closed.", vbInformation
        Unload Me
        Exit Sub
    End If
    lblMessage.Caption = ""
    btnEmpRefresh.Enabled = True
    btnEmployeeSearch.Enabled = True
End Sub

Private Sub cmbBranch_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
 If HRBranchSw = True Then
    Call RefreshBranch
    HRBranchSw = False
End If
End Sub

Private Sub cmbCivilStatus_GotFocus()
cmbcivilstatus.BackColor = &HFFFFC0
End Sub

Private Sub cmbCivilStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbCivilStatus_LostFocus()
cmbcivilstatus.BackColor = vbWhite
End Sub

Private Sub cmbCompany_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbdept_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
 If HRDeptSw = True Then
    Call RefreshDept
    HRDeptSw = False
End If
End Sub

Private Sub cmbPosition_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
 If HRPosSw = True Then
    Call RefreshPosition
    HRPosSw = False
End If
End Sub

Private Sub cmbShiftSchedule_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SSTab1.Tab = 0
    txthouseno.SetFocus
End If
End Sub

Private Sub cmbShiftSchedule_GotFocus()
cmbShiftSchedule.BackColor = &HFFFFC0
End Sub

Private Sub cmbShiftSchedule_LostFocus()
cmbShiftSchedule.BackColor = vbWhite
End Sub

Private Sub cmbempstatus_GotFocus()
    cmbEmpStatus.BackColor = &HFFFFC0
End Sub

Private Sub cmbempstatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbempstatus_LostFocus()
    cmbEmpStatus.BackColor = vbWhite
End Sub

Private Sub cmbPayBasis_Click()
  On Error Resume Next
    If cmbPayBasis.ListIndex = 0 Then lblRAte.Caption = "Hourly Rate"
    If cmbPayBasis.ListIndex = 1 Then lblRAte.Caption = "Daily Rate"
    If cmbPayBasis.ListIndex = 2 Then lblRAte.Caption = "Monthly Rate"
    If cmbPayBasis.ListIndex = 3 Then lblRAte.Caption = "Commission %"
    
    If addflagship <> True Then
        If cmbPayBasis.ListIndex = 0 Then txtRate.Text = Format(rsHREmp("hourlyrate"), "#00.000")
        If cmbPayBasis.ListIndex = 1 Then txtRate.Text = Format(rsHREmp("dailyrate"), "#00.00")
        If cmbPayBasis.ListIndex = 2 Then txtRate.Text = Format(rsHREmp("monthlyrate"), "#00.00")
        If cmbPayBasis.ListIndex = 3 Then txtRate.Text = Format(rsHREmp("CommissionRate"), "#00.00")
   End If
   If txtRate.Enabled = True Then txtRate.SetFocus
End Sub

Private Sub cmdAddBranch_Click()
    frmWorkStation.Show 1
End Sub

Private Sub cmdAddDept_Click()
    HRDeptSw = True
    frmDept.Show
End Sub

Private Sub cmdDelete_Click()
Dim oldControlNo As String
Dim oldEmpNo As String
Dim res


On Error Resume Next

    If DeleteRights = False Then
        MsgBox "You do not have rights to Delete Employee!", vbInformation, PayrollAlert
        Exit Sub
    End If

oldEmpNo = TxtControlNo.Text
    
    
Set rsDelete = New ADODB.Recordset
StrQuery = "Select * from EmployeeBasicSalary where EmployeeID = " & Me.txtID.Text & "  and EmployeeStatus = 1"
Set rsDelete = ReturnRst(StrQuery, cnnstr, mstatic)
'rsDelete.Open "Select * from employeebasicsalary where EmployeeID=" & Me.txtID.Text & "  and employeestatus=true", cnn, adOpenKeyset, adLockOptimistic

If rsDelete.EOF = False Then
    MsgBox "This employee is included in Payroll List, cannot be deleted. Remove it in Employee Payroll Master table first.", vbInformation, PayrollAlert
    Exit Sub
Else
    Call DeleteEmployee
    If mHRSEARCH = True Then
          Call frmHRSearch.lvwdisplay.ListItems.Clear
          Call frmHRSearch.chkStatus_Click(1)
    End If
    
    If oldControlNo = TxtControlNo.Text Then
        Call btnEmpRefresh_Click
    End If
  
End If
rsDelete.Close
Set rsDelete = Nothing
End Sub

Private Function GetEmployeeID() As Integer
    On Error GoTo err2
    
    StrQuery = "Select EmployeeID from hremployees where EmpNo='" & Me.TxtControlNo.Text & "'  and agencycode = " & CompanyCode
    Set rsemp = ReturnRst(StrQuery, cnnstr, mstatic)
    'Set rsemp = ReturnRst("Select EmployeeID from hremployees where EmpNo='" & Me.TxtControlNo.Text & "'  and agencycode=" & CompanyCode, CnnstrHR, mKeySet)
    If rsemp.EOF = False Then
        GetEmployeeID = rsemp.Fields(0)
    Else
        GetEmployeeID = 0
    End If
    rsemp.Close
    Set rsemp = Nothing
    Exit Function

err2:
    Err.Clear
    GetEmployeeID = 0

End Function

Private Sub DeleteEmployee()
Dim sqlString As String
Dim res, id

On Error Resume Next
 
  res = MsgBox("You are about to delete 1 record." & vbCrLf & " " & vbCrLf & "If you click Yes, You won't able to undo this delete operation. " & vbCrLf & "Are you sure you want to delete the Employee No: " & TxtControlNo & "?", vbQuestion + vbYesNo, CompanyName)
  If res = vbYes Then
      id = Me.txtID.Text
     sqlString = "Delete from HREmployees where EmployeeID=" & id & " and agencycode=" & CompanyCode
     Call Action_Query(sqlString, cnnstr)
     
    If Trim(TxtControlNo.Text) <> "" Then
    sqlString = "Delete from HRDetails where EmployeeID = " & id:        Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HRCharReference where EmployeeID = " & id:          Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HRDialects where EmployeeID=" & id:           Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HREmploymentRecord where EmployeeID=" & id:       Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HRSeminar where EmployeeID=" & id:          Call Action_Query(sqlString, cnnstr)
    End If
    
    sqlString = "Delete from HREmployeesAssignment where EmployeeID=" & id:        Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HREmployeesLeaves where EmployeeID=" & id:          Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HREmployeesOffenses where EmployeeID=" & id:       Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HREmployeesTraining where EmployeeID=" & id:       Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HREmployeeBenefits where EmployeeID=" & id:        Call Action_Query(sqlString, cnnstr)
    sqlString = "Delete from HREmployeesAppraisal where EmployeeID=" & id:       Call Action_Query(sqlString, cnnstr)
    'frmStatus.Show 1
    MsgBox "The record of [" & txtsurname.Text & ", " & txtfirstname.Text & "] was deleted.", vbInformation
    Call btnEmpRefresh_Click
End If
End Sub

Public Sub cmdEdit_Click()
    If EditRights = False Then
        MsgBox "You do not have rights to Edit Employee!", vbInformation, PayrollAlert
        Exit Sub
    End If
    
    DoEvents
    'cmbshiftsched.Visible = True
    OrigAcctNo = txtacctno.Text ' save account no
    origCivilStatus = cmbcivilstatus.Text
    OrigEmpStatus = cmbEmpStatus.Text
    OrigCompany = cmbCompany.Text
    OrigDateHired = MskDateHired.Text
    OrigStartContract = mskStartContract.Text
    OrigEndContract = mskEndContract.Text
    OrigPayBasis = cmbPayBasis.Text
    OrigRate = txtRate.Text
    OrigMonthly = txtMonthlyRate.Text
    OrigTASAesla = txtCBtotal.Text
    
    reContract = False
    lblMessage.Caption = "You are about to EDIT this employee."
    TmpSurname = Me.txtsurname.Text
    TmpFirstname = Me.txtfirstname.Text
    TmpMiddlename = Me.txtmiddlename.Text
    
    prevControlNo = Me.TxtControlNo.Text
    PrevCompany = Me.cmbCompany.Text
    PrevStartContract = Me.mskStartContract.Text
    PrevEndContract = Me.mskEndContract.Text
    
    editflagship = True
    EditFlag = True
    addflagship = False
    AddFlag = False
    addflagship = False
    cmdCopy.Enabled = False
    cmdRecontract.Enabled = False
    Call DisAble_ActionButtons
    Call TextEnabled(True)
    'Bernard 282008 disabled the EmpNo ,not editable
    TxtControlNo.Enabled = True
    Call CheckUser1(operatorID)
    txtsurname.SetFocus
End Sub

Private Sub cmdfirst_Click()
    rsHREmp.MoveFirst
    cmdfirst.Enabled = False
    cmdPrev.Enabled = False
    cmdlast.Enabled = True
    cmdNext.Enabled = True
    Showdata rsHREmp
End Sub

Private Sub cmdlast_Click()
    rsHREmp.MoveLast
    cmdlast.Enabled = False
    cmdNext.Enabled = False
    cmdPrev.Enabled = True
    cmdfirst.Enabled = True
    Showdata rsHREmp
End Sub

Private Sub cmdnew_Click()

    If AddRights = False Then
        MsgBox "You do not have rights to add New Employee!", vbInformation, PayrollAlert
        Exit Sub
    End If
    addflagship = True
    editflagship = False
    'LblUpdateBy.Caption = UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID)))
    Me.Caption = "EMPLOYEES INFORMATION   [Updated by: " + UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID))) + "]"
    'Me.Caption = "EMPLOYEES INFORMATION [Last updated by: " + LoginName + "]"
    lblCompany.Caption = "For Selection of Company"
    Call Add
    Call ClearTextBox
    lblMessage.Caption = "You are about to ADD new employee..."
    'txtControlNo.Text = Format(Date, "yy") & "-" & Format(Date, "mm") & "-" & Format(NewID + 1, "00000")
    TxtControlNo.Text = GetNewID
    txtDTRID.Text = Right("00000000" & GetNewDTRID, 8) ' get new dtr id
    TempNewID = TxtControlNo
    'frmHRFinalEmployees.TxtControlNo.Text = "<auto>"
    frmHRFinalEmployees.TxtControlNo.Text = TxtControlNo
    
    Call TextEnabled(True)
    'Bernard 282008 disabled the EmpNo ,not editable
    TxtControlNo.Enabled = False
    txtsurname.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsHREmp.MoveNext
    cmdPrev.Enabled = True
    cmdfirst.Enabled = True
    If rsHREmp.EOF Then
        rsHREmp.MoveLast
        cmdNext.Enabled = False
        cmdlast.Enabled = False
        MsgBox "End of file encountered!", vbInformation, CompanyName
        Exit Sub
    End If
    Showdata rsHREmp
End Sub

Private Sub cmdPrev_Click()
    rsHREmp.MovePrevious
    cmdNext.Enabled = True
    cmdlast.Enabled = True
    If rsHREmp.BOF Then
        rsHREmp.MoveFirst
        cmdPrev.Enabled = False
        cmdfirst.Enabled = False
        MsgBox "Beginning of file encountered!", vbInformation, CompanyName
    End If
    'Employeenum = rsHREmp!EmployeeID  ' comment by bernard
    Showdata rsHREmp
End Sub

Private Sub cmdsearch_Click()
    frmHRSearch.Show vbModal
End Sub

Private Sub cmdundo_Click()
   
    If AddFlag = True Then
        Unload Me
        Exit Sub
    End If
    lblMessage.Caption = ""
    addflagship = False
    editflagship = False
    'cmbshiftsched.Visible = False
    Call TextEnabled(False)
    TxtControlNo.Enabled = False
    
    'Call LoadCompany(CompanyCode)
    Showdata rsHREmp
    btnEmpRefresh.Enabled = True
    btnEmployeeSearch.Enabled = True
    cmdundo.Enabled = False
    cmdupdate.Enabled = False
    cmdNew.Enabled = True
    cmddelete.Enabled = True
    cmdedit.Enabled = True
    cmdfirst.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdlast.Enabled = True
    cmdRecontract.Enabled = True
    cmdAddBranch.Enabled = False
    cmdAddDept.Enabled = False
    cmdAddPosition.Enabled = False
    cmbCompany.Enabled = False
End Sub

Public Sub Cancel()
    Call TextEnabled(False)
End Sub

Private Sub TextEnabled(ByVal vlfalse As Boolean)
    cmdAddBranch.Enabled = vlfalse
    cmdAddDept.Enabled = vlfalse
    cmdAddPosition.Enabled = vlfalse
    
    cmbAllowance.Enabled = vlfalse
    
    TxtControlNo.Enabled = vlfalse
    txtprimarysch.Enabled = vlfalse
    txtdateprimary.Enabled = vlfalse
    txtnickname.Enabled = vlfalse
    txtsurname.Enabled = vlfalse
    txtfirstname.Enabled = vlfalse
    txtmiddlename.Enabled = vlfalse
    txtfirstchoice.Enabled = vlfalse
    mskbday.Enabled = vlfalse
    txtage.Enabled = vlfalse
    txtaddress1.Enabled = vlfalse
    txthouseno.Enabled = vlfalse
    txtstreet.Enabled = vlfalse
    txtbarangay.Enabled = vlfalse
    txtdistrict.Enabled = vlfalse
    txtcity.Enabled = vlfalse
    txtprovince.Enabled = vlfalse
    txtprovinceaddress.Enabled = vlfalse
    txtemergency.Enabled = vlfalse
    txteraddress.Enabled = vlfalse
    txtertel.Enabled = vlfalse
    If AddFlag = True Or addflagship = True Then  ' enable selection for add
        cmbCompany.Enabled = vlfalse
    End If
    cmbDept.Enabled = vlfalse
    cmbPosition.Enabled = vlfalse
    cmbBranch.Enabled = vlfalse
    cmbShiftSchedule.Enabled = vlfalse
    
    optmale.Enabled = vlfalse
    optfemale.Enabled = vlfalse
    txtheight.Enabled = vlfalse
    txtweight.Enabled = vlfalse
    txttel1.Enabled = vlfalse
    txttel2.Enabled = vlfalse
    cmbcivilstatus.Enabled = vlfalse
    txtexam.Enabled = vlfalse
    txtinterview.Enabled = vlfalse
    txtmedical.Enabled = vlfalse
    msksss.Enabled = vlfalse
    txtphilhealth.Enabled = vlfalse
    txtpagibig.Enabled = vlfalse
    txtacctno.Enabled = vlfalse
    mskissue.Enabled = vlfalse
    txtcitizen.Enabled = vlfalse
    txtbirthplace.Enabled = vlfalse
    txthaircolor.Enabled = vlfalse
    txteyecolor.Enabled = vlfalse
    txtreligion.Enabled = vlfalse
    msktax.Enabled = vlfalse
    txtpostsch.Enabled = vlfalse
    txtdatepost.Enabled = vlfalse
    txtsecsch.Enabled = vlfalse
    txtdatesec.Enabled = vlfalse
    txtcollegesch.Enabled = vlfalse
    txtdatecollege.Enabled = vlfalse
    txtcollegecourse.Enabled = vlfalse
    txtcollegemajor.Enabled = vlfalse
    txtdatepost.Enabled = vlfalse
    txtpostcourse.Enabled = vlfalse
    txtfathername.Enabled = vlfalse
    txtfaddress.Enabled = vlfalse
    txtfjob.Enabled = vlfalse
    txtmothername.Enabled = vlfalse
    txtmaddress.Enabled = vlfalse
    txtmjob.Enabled = vlfalse
    txtspousename.Enabled = vlfalse
    
    TxtContactPerson.Enabled = vlfalse
    txtContactAddress.Enabled = vlfalse
    txtContactPhone.Enabled = vlfalse
    
    txtspouseaddress.Enabled = vlfalse
    txtsjob.Enabled = vlfalse
    txtnbi.Enabled = vlfalse
    txtrescer.Enabled = vlfalse
    txtissueat.Enabled = vlfalse
    txtdep1.Enabled = vlfalse
    txtdep2.Enabled = vlfalse
    txtdep3.Enabled = vlfalse
    txtdep4.Enabled = vlfalse
    txtdep5.Enabled = vlfalse
    mskdep1.Enabled = vlfalse
    mskdep2.Enabled = vlfalse
    mskdep3.Enabled = vlfalse
    mskdep4.Enabled = vlfalse
    mskdep5.Enabled = vlfalse
    txtnumdep.Enabled = vlfalse
    txtdialect1.Enabled = vlfalse
    txtdialect2.Enabled = vlfalse
    txtdialect3.Enabled = vlfalse
    txtseminar1.Enabled = vlfalse
    txtseminar2.Enabled = vlfalse
    txtseminar3.Enabled = vlfalse
    txtdateseminar1.Enabled = vlfalse
    txtdateseminar2.Enabled = vlfalse
    txtdateseminar3.Enabled = vlfalse
    txtnameofcompany1.Enabled = vlfalse
    txtnameofcompany2.Enabled = vlfalse
    txtnameofcompany3.Enabled = vlfalse
    txtfrom1.Enabled = vlfalse
    txtfrom2.Enabled = vlfalse
    txtfrom3.Enabled = vlfalse
    txtto1.Enabled = vlfalse
    txtto2.Enabled = vlfalse
    txtto3.Enabled = vlfalse
    txtpos1.Enabled = vlfalse
    txtpos2.Enabled = vlfalse
    txtpos3.Enabled = vlfalse
    txtreason1.Enabled = vlfalse
    txtreason2.Enabled = vlfalse
    txtreason3.Enabled = vlfalse
    txtsuperior1.Enabled = vlfalse
    txtsuperior2.Enabled = vlfalse
    txtsuperior3.Enabled = vlfalse
    txtrefname1.Enabled = vlfalse
    txtrefname2.Enabled = vlfalse
    txtrefname3.Enabled = vlfalse
    txtrefjob1.Enabled = vlfalse
    txtrefjob2.Enabled = vlfalse
    txtrefjob3.Enabled = vlfalse
    txtrefaddress1.Enabled = vlfalse
    txtrefaddress2.Enabled = vlfalse
    txtrefaddress3.Enabled = vlfalse
    txtrefphone1.Enabled = vlfalse
    txtrefphone2.Enabled = vlfalse
    txtrefphone3.Enabled = vlfalse
    txtrefrel1.Enabled = vlfalse
    txtrefrel2.Enabled = vlfalse
    txtrefrel3.Enabled = vlfalse
    MskDateHired.Enabled = vlfalse
    mskdateprob.Enabled = vlfalse
    mskdatereg.Enabled = vlfalse
    mskdateres.Enabled = vlfalse
    cmbEmpStatus.Enabled = vlfalse
    txtsupervisor.Enabled = vlfalse
    txtreasonresign.Enabled = vlfalse
    mskendof.Enabled = vlfalse
    cmbPayBasis.Enabled = vlfalse
    txtRate.Enabled = vlfalse
    txtMonthlyRate.Enabled = vlfalse
    mskStartContract.Enabled = vlfalse
    mskEndContract.Enabled = vlfalse
    txtMonthsHired.Enabled = vlfalse
    TxtEcolaAmount.Enabled = vlfalse
    txtBillingRate.Enabled = vlfalse
    txtCBtotal.Enabled = vlfalse
    txtCBpayment.Enabled = vlfalse
    mskStartCB.Enabled = vlfalse
    mskEndCB.Enabled = vlfalse
    
    mskStartCash.Enabled = vlfalse
    mskEndCash.Enabled = vlfalse
    mskStartUniform.Enabled = vlfalse
    mskEndUniform.Enabled = vlfalse

    TxtMFAmount.Enabled = vlfalse
    txtPFAmount.Enabled = vlfalse
    
    TxtMFPayment.Enabled = vlfalse
    txtPFPayment.Enabled = vlfalse
    txtSuretyPercent.Enabled = vlfalse
    
    MskMFStartPayment.Enabled = vlfalse
    MskMFEndPayment.Enabled = vlfalse
    MskPFStartPayment.Enabled = vlfalse
    MskPFEndPayment.Enabled = vlfalse
    Frame11.Enabled = vlfalse
    Frame9.Enabled = vlfalse
    'ChkEcola.Enabled = vlfalse
    chkMisc.Enabled = vlfalse
    chkFund.Enabled = vlfalse
    
End Sub

Public Function ValidateTextbox() As Boolean

On Error Resume Next

    If TxtControlNo.Text = "" Then MsgBox "Employee number is required.", vbInformation: TxtControlNo.SetFocus: GoTo last
    If txtsurname.Text = "" Then MsgBox "Surname is required.", vbInformation: txtsurname.SetFocus: GoTo last
    If txtfirstname.Text = "" Then MsgBox "Firstname is required.", vbInformation: txtfirstname.SetFocus: GoTo last
     
    If Me.cmbEmpStatus.ListIndex = -1 Then MsgBox "Please select employee status.", vbInformation: SSTab1.Tab = 1: Me.cmbEmpStatus.SetFocus: GoTo last
    
    'If txtmiddlename.Text = "" Then MsgBox "Middle name/Middle initial is required", vbInformation: txtmiddlename.SetFocus: GoTo last
    If cmbCompany.Text = "" Then MsgBox "Please select company for this employee.", vbInformation: SSTab1.Tab = 1: cmbCompany.SetFocus: GoTo last
    
    If addflagship = True And cmbPayBasis.Text = "" Then
        cmbPayBasis.Text = "Daily"
    End If
    
    If Me.cmbPayBasis.ListIndex = -1 Then MsgBox "Please select Payment Basis for this employee.", vbInformation: SSTab1.Tab = 1: Me.cmbPayBasis.SetFocus: GoTo last
    If Me.txtRate.Text = 0 Or Me.txtRate.Text = "" Then
        txtRate.Text = 5#
    End If
    If Me.txtRate.Enabled = True And (Me.txtRate.Text = 0 Or Me.txtRate.Text = "") Then MsgBox "Please enter Salary Rate of this employee.", vbInformation: SSTab1.Tab = 1: Me.txtRate.Enabled = True: Me.txtRate.SetFocus: GoTo last
    
    ' Updated January 2, 2009.
    If Me.txtMonthlyRate.Text = "" Or Me.txtMonthlyRate.Text = 0 Then
        Me.txtMonthlyRate.Text = 10
    End If
    If (Me.txtMonthlyRate.Text = 0 Or Me.txtMonthlyRate.Text = "") And cmbPayBasis.Text <> "Commission" Then MsgBox "Please enter Monthly salary rate of this employee.", vbInformation: SSTab1.Tab = 1: Me.txtMonthlyRate.Enabled = True: Me.txtMonthlyRate.SetFocus: GoTo last
     
    If Me.cmbcivilstatus.ListIndex = -1 Then MsgBox "Please select Civil Status of this employee.", vbInformation: SSTab1.Tab = 0: Me.cmbcivilstatus.SetFocus: GoTo last
  
    If Left(cmbEmpStatus.Text, 1) = "C" Then ' ask for contract date
    If mskStartContract.Text = "__/__/____" Or mskStartContract.Text = "__/__/____" Then
        MsgBox "Please enter Start and End contract of employee.", vbInformation
        SSTab1.Tab = 1
        mskStartContract.SetFocus
        GoTo last
    Else
        If Format(mskEndContract.Text, "yyyymmdd") < Format(mskStartContract.Text, "yyyymmdd") Then
            MsgBox "End Contract should be greater than Start Contract!!", vbInformation
            SSTab1.Tab = 1
            mskEndContract.SetFocus
            GoTo last
        End If
    End If
    End If
 
    Dim date1 As Date
    Dim date2 As Date
     
'    If MskPFStartPayment <> "__/__/____" And IsDate(MskPFStartPayment) = True And MskPFEndPayment <> "__/__/____" And IsDate(MskPFEndPayment) = True Then
'     date1 = MskPFStartPayment.Text
'     date2 = Me.MskPFEndPayment.Text
'
'        If date1 > date2 Then
'             MsgBox "PF end date must not be earlier than start date.", vbInformation
'             MskPFEndPayment.SetFocus
'             GoTo last
'        End If
'     End If
'
'        If MskMFStartPayment <> "__/__/____" And IsDate(MskMFStartPayment) = True And MskMFEndPayment <> "__/__/____" And IsDate(MskMFEndPayment) = True Then
'     date1 = MskMFStartPayment.Text
'     date2 = MskMFEndPayment.Text
'
'        If date1 > date2 Then
'             MsgBox "MF end date must not be earlier than start date.", vbInformation
'             MskMFEndPayment.SetFocus
'             GoTo last
'        End If
'     End If
'
'  If mskStartCB <> "__/__/____" And IsDate(mskStartCB) = True And mskEndCB <> "__/__/____" And IsDate(mskEndCB) = True Then
'     date1 = mskStartCB.Text
'     date2 = mskEndCB.Text
'
'        If date1 > date2 Then  ' should be less than
'             MsgBox "CB end date must not be earlier than start date.", vbInformation
'             mskEndCB.SetFocus
'             GoTo last
'        End If
'     End If
   
    If RTrim(txtRate.Text) = "" Then txtRate = "0.000"
    ValidateTextbox = True
    Exit Function

last:
    
    ValidateTextbox = False

End Function

Private Sub cmdUpdate_Click()
     If ValidateTextbox = False Then Exit Sub
     If AddFlag = True Or addflagship = True Then 'SAVE FOR NEW EMPLOYEES
         TempNewID = TxtControlNo.Text ' getnewid
         If SaveNewHREmployeeRecord = False Then Exit Sub
         Call SetButtonsave
         Call TextEnabled(False)
         cmbCompany.Enabled = False
         GetNewID (True) ' update the employee sequence number table
         AddFlag = False
         addflagship = False
         EditFlag = False
     End If
    
     If EditFlag = True Or editflagship = True Then 'UPDATE EMPLOYEES
         If UpdateHREmployeeRecord = False Then Exit Sub
            If reContract = True Then
                 MsgBox "[" & txtsurname.Text & ", " & txtfirstname.Text & "] was successfully transfer to " & cmbCompany.Text & "!", vbInformation
                 Call btnEmpRefresh_Click
             Else
                 MsgBox "[" & txtsurname.Text & ", " & txtfirstname.Text & "] was successfully updated!", vbInformation
             End If
         Call DisAble_ActionButtons
         Call TextEnabled(False)
         cmbCompany.Enabled = False
     End If
     TxtControlNo.Enabled = False
     
     'TxtControlNo.Text = TempNewID
     Call SetButtonsave
     lblMessage.Caption = ""
     addflagship = False
     editflagship = False
     Call cmbempstatus_Click
     
     If mHRSEARCH = True Then 'REFRESH HR SEARCH FORM
     If TmpMiddlename <> txtmiddlename.Text Or TmpSurname <> Me.txtsurname.Text Or TmpFirstname <> Me.txtfirstname.Text Then Call frmHRSearch.lvwdisplay.ListItems.Clear
         Call frmHRSearch.chkStatus_Click(1)
     End If
End Sub

Public Sub SaveToHistory()
    Dim rsemp As New ADODB.Recordset

'    StrQuery = "Select * from HREmployees_HISTORY where EmployeeID = '" & TxtID.Text & _
'        "' and agencycode = " & CompanyCode & " and convert(varchar(10), startcontract, 112) = " & _
'        Format(Me.mskStartContract.Text, "yyyymmdd") & " and convert(varchar(10), endcontract, 112) = '" & _
'        Format(Me.mskEndContract.Text, "yyyymmdd") & "'"
'    'StrQuery = "Select * from HREmployees_HISTORY where EmployeeID = '" & txtID.Text & "' and agencycode = " & CompanyCode & " and startcontract = #" & Me.mskStartContract.Text & "# and endcontract=#" & Me.mskEndContract.Text & "#"
'
'    Set rsemp = ReturnRst(StrQuery, cnnstr, mstatic)
'    If rsemp.EOF = True Then
        
    StrQuery = "INSERT INTO HREmployees_HISTORY SELECT * FROM HREMPLOYEES WHERE EmployeeID = " & txtID.Text & " and agencycode = " & CompanyCode
    Call Action_Query(StrQuery, cnnstr)

'    End If
End Sub

Private Function UpdateHREmployeeRecord() As Boolean

    Dim tempCompanyCode As Integer
    Dim tempConnStr As String
    Dim SuccessFlag As Boolean
    
    TempNewID = TxtControlNo.Text
    tempCompanyCode = CompanyCode
    tempConnStr = cnnstr
    
    If tempConnStr = "" Then Exit Function
    
    'SAVE TO HISTORY FIRST.
    If reContract = True Then
        Call SaveToHistory
    End If
    
    Call ComputeValue

    '2192008
    Dim strWhere As String
    Dim rsCheck As New ADODB.Recordset
    
    SuccessFlag = False
    StrQuery = "SELECT EmployeeID FROM HREmployees WHERE EmployeeID = " & txtID.Text 'EmpNo =  '" & txtControlNo.Text & "'"
    Set rsCheck = ReturnRst(StrQuery, tempConnStr, mstatic)
    If rsCheck.EOF = True Then ' Employee number does not exist, insert it
        sqlString = "INSERT INTO HREmployees (EmployeeID, agencycode, EmpNo, DTRID, FirstName, Empposition, age, middlename, Surname, nickname, BIRTHDATE, Address1, houseno, street, barangay, district, city, province1, province, contact, caddress,cphone,gender,Height,Weight,Telnum1,Telnum2,TaxCode,examgrade,interviewgrade,medicalgrade,sss,philhealth, pagibig,AcctNo, citizen, birthplace,haircolor, eyecolor, Religion, TIN, [Primary],dateprimary, secondary, datesecondary,college, datecollege, degree, Major, postgraduate," & _
         " datepostgraduate, postgraduatecourse, fathername, faddress, fwork, mothername, maddress, mwork, nbi, rescer, issueat, issueon, datehired, dateprob, dateregular, dateresign, employmentstatus, supervisor, reasonres, endofcontract, ClientCode, brcode, positioncode, " & _
         " PaymentBasis,HourlyRate,DailyRate, MonthlyRate, CommissionRate, StartContract, EndContract, NoofMonthsHired, cbamount, cbpayment,cbstartpayment, cbendpayment, caStartPayment, caEndPayment, uniStartPayment, uniEndPayment, MFAmount, MFPayment,MFStartPayment,MFEndPayment, PFAmount, PFpayment, FundPercentage, PFstartpayment, PFendpayment, WithEcolaPay, WithMiscPay, WithFundPay, CashAdvance, MiscFee, [USERID], DeptCode, PicturePath, ShiftCode, remarks, Town, LastUpdate) " & _
         " VALUES (" & txtID.Text & "," & tempCompanyCode & ", '" & TempNewID & "', '" & txtDTRID.Text & "', " & ChkVal(txtfirstname) & "," & ChkVal(txtfirstchoice) & "," & IIf(txtage.Text = "", 0, txtage.Text) & "," & ChkVal(txtmiddlename) & "," & ChkVal(txtsurname) & "," & ChkVal(txtnickname) & "," & ChkMsk(mskbday) & "," & ChkVal(txtaddress1) & "," & ChkVal(txthouseno) & "," & ChkVal(txtstreet) & "," & ChkVal(txtbarangay) & "," & ChkVal(txtdistrict) & "," & ChkVal(txtcity) & "," & ChkVal(txtprovince) & "" & _
         "," & ChkVal(txtprovinceaddress) & "," & ChkVal(txtemergency) & "," & ChkVal(txteraddress) & "," & ChkVal(txtertel) & ",'" & IIf(optmale.VALUE, "M", "F") & "','" & Replace(txtheight, "'", """") & "'," & ChkVal(txtweight) & "," & ChkVal(txttel1) & "," & ChkVal(txttel2) & ",'" & IIf(cmbcivilstatus.Text = "", "S", cmbcivilstatus.Text) & "'," & ChkVal(txtexam) & "," & ChkVal(txtinterview) & "," & ChkVal(txtmedical) & "," & ChkMsk(msksss) & "," & ChkVal(txtphilhealth) & "," & ChkVal(txtpagibig) & "," & ChkVal(txtacctno) & "" & _
         "," & ChkVal(txtcitizen) & "," & ChkVal(txtbirthplace) & "," & ChkVal(txthaircolor) & "," & ChkVal(txteyecolor) & "," & ChkVal(txtreligion) & "," & ChkMsk(msktax) & "," & ChkVal(txtprimarysch) & "," & ChkVal(txtdateprimary) & "," & ChkVal(txtsecsch) & "," & ChkVal(txtdatesec) & "," & ChkVal(txtcollegesch) & "," & ChkVal(txtdatecollege) & "," & ChkVal(txtcollegecourse) & "," & ChkVal(txtcollegemajor) & "," & ChkVal(txtpostsch) & "," & ChkVal(txtdatepost) & "" & _
         "," & ChkVal(txtpostcourse) & "," & ChkVal(txtfathername) & "," & ChkVal(txtfjob) & "," & ChkVal(txtfaddress) & "," & ChkVal(txtmothername) & "," & ChkVal(txtmaddress) & "," & ChkVal(txtmjob) & "," & ChkVal(txtnbi) & "," & ChkVal(txtrescer) & "," & ChkVal(txtissueat) & "," & ChkMsk(mskissue) & "," & ChkMsk(MskDateHired) & "," & ChkMsk(mskdateprob) & "," & ChkMsk(mskdatereg) & "" & _
         "," & ChkMsk(mskdateres) & ",'" & Left(cmbEmpStatus.Text, 1) & "'," & ChkVal(txtsupervisor) & "," & ChkVal(txtreasonresign) & "," & ChkMsk(mskendof) & "," & Cid & "," & BrCode & ", " & PosCode & ", '" & PaymentBasis & "'," & HourlyRate & "," & DailyRate & "," & MonthlyRate & "," & CommissionRate & "," & ChkMsk(mskStartContract) & "," & ChkMsk(mskEndContract) & "," & IIf(txtMonthsHired.Text = "", 5, txtMonthsHired.Text) & "" & _
         "," & IIf(txtCBtotal = "", 0, Format(txtCBtotal, "##0.00")) & "," & IIf(txtCBpayment = "", 0, Format(txtCBpayment, "##0.00")) & "," & ChkMsk(mskStartCB) & "," & ChkMsk(mskEndCB) & "," & ChkMsk(mskStartCash) & "," & ChkMsk(mskEndCash) & "," & ChkMsk(mskStartUniform) & "," & ChkMsk(mskEndUniform) & "" & _
         "," & IIf(TxtMFAmount = "", 0, Format(TxtMFAmount, "##0.00")) & "," & IIf(TxtMFPayment = "", 0, Format(TxtMFPayment, "##0.00")) & "," & ChkMsk(MskMFStartPayment) & "," & ChkMsk(MskMFEndPayment) & "" & _
         "," & IIf(txtPFAmount = "", 0, Format(txtPFAmount, "##0.00")) & "," & IIf(txtPFPayment = "", 0, Format(txtPFPayment, "##0.00")) & "," & IIf(txtSuretyPercent = "", 0, Format(txtSuretyPercent, "##0.00")) & "," & ChkMsk(MskPFStartPayment) & "," & ChkMsk(MskPFEndPayment) & "" & _
         ",'" & WithEcolaPay & ",'" & WithMiscPay & ",'" & WithFundPay & "'," & IIf(TxtEcolaAmount = "", 0, Format(TxtEcolaAmount, "##0.00")) & ",0 ," & IIf(txtBillingRate = "", 0, Format(txtBillingRate, "##0.00")) & ", '" & operatorID & "'," & DeptCode & ", '" & PictureFile & "', " & mShiftCode & ", '" & IIf(Me.txtreasonresign.Text = "", " ", txtreasonresign.Text) & "', '" & cmbAllowance.Text & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
        If Action_Query(sqlString, tempConnStr) = True Then
            SuccessFlag = True
        End If
    Else
        sqlString = "Update HREmployees set EmpNo='" & RTrim(TempNewID) & "'"
        sqlString2 = ", FirstName=" & ChkVal(txtfirstname) & ", Empposition =" & ChkVal(txtfirstchoice) & ", age=" & IIf(txtage.Text = "", 0, txtage.Text) & ", middlename=" & ChkVal(txtmiddlename) & ", Surname=" & ChkVal(txtsurname) & ", nickname=" & ChkVal(txtnickname) & ", BIRTHDATE=" & ChkMsk(mskbday) & ", Address1=" & ChkVal(txtaddress1) & ", houseno=" & ChkVal(txthouseno) & ", street=" & ChkVal(txtstreet) & ", barangay=" & ChkVal(txtbarangay) & ",district=" & ChkVal(txtdistrict) & ", city=" & ChkVal(txtcity) & ", province1=" & ChkVal(txtprovince) & ", province=" & ChkVal(txtprovinceaddress) & ", contact=" & ChkVal(txtemergency) & ", caddress=" & ChkVal(txteraddress) & ",cphone=" & ChkVal(txtertel) & ",gender='" & IIf(optmale.VALUE, "M", "F") & "',Height='" & Replace(txtheight, "'", """") & "',Weight=" & ChkVal(txtweight) & "" & _
       " ,Telnum1=" & ChkVal(txttel1) & ",Telnum2=" & ChkVal(txttel2) & ",TaxCode='" & IIf(cmbcivilstatus.Text = "", "S", cmbcivilstatus.Text) & "',examgrade=" & ChkVal(txtexam) & ",interviewgrade=" & ChkVal(txtinterview) & ",medicalgrade=" & ChkVal(txtmedical) & ",sss=" & ChkMsk(msksss) & ",philhealth=" & ChkVal(txtphilhealth) & ", pagibig=" & ChkVal(txtpagibig) & ",AcctNo=" & ChkVal(txtacctno) & "" & _
       ", citizen=" & ChkVal(txtcitizen) & ", birthplace=" & ChkVal(txtbirthplace) & ",haircolor=" & ChkVal(txthaircolor) & ", eyecolor=" & ChkVal(txteyecolor) & ", Religion=" & ChkVal(txtreligion) & ", TIN=" & ChkMsk(msktax) & ", [Primary]=" & ChkVal(txtprimarysch) & ",dateprimary=" & ChkVal(txtdateprimary) & ", secondary=" & ChkVal(txtsecsch) & ", datesecondary=" & ChkVal(txtdatesec) & ",college=" & ChkVal(txtcollegesch) & ", datecollege=" & ChkVal(txtdatecollege) & ", degree=" & ChkVal(txtcollegecourse) & ", Major=" & ChkVal(txtcollegemajor) & ", postgraduate=" & ChkVal(txtpostsch) & "," & _
       " datepostgraduate= " & ChkVal(txtdatepost) & ", postgraduatecourse=" & ChkVal(txtpostcourse) & " , fathername=" & ChkVal(txtfathername) & ", faddress=" & ChkVal(txtfaddress) & ", fwork=" & ChkVal(txtfjob) & ", mothername=" & ChkVal(txtmothername) & ", maddress=" & ChkVal(txtmaddress) & ", mwork=" & ChkVal(txtmjob) & ", nbi=" & ChkVal(txtnbi) & ", rescer=" & ChkVal(txtrescer) & ", issueat=" & ChkVal(txtissueat) & ", issueon = " & ChkMsk(mskissue) & ", datehired=" & ChkMsk(MskDateHired) & ", dateprob=" & ChkMsk(mskdateprob) & ", dateregular=" & ChkMsk(mskdatereg) & ", dateresign=" & ChkMsk(mskdateres) & ", employmentstatus='" & Left(cmbEmpStatus.Text, 1) & "'" & _
       ", supervisor=" & ChkVal(txtsupervisor) & ", reasonres=" & ChkVal(txtreasonresign) & ", endofcontract=" & ChkMsk(mskendof) & ", ClientCode=" & Cid & ", brcode=" & BrCode & ", PositionCode = " & PosCode & "" & _
       ", PaymentBasis='" & PaymentBasis & "',HourlyRate=" & HourlyRate & ",DailyRate=" & DailyRate & ", MonthlyRate=" & MonthlyRate & ", CommissionRate=" & CommissionRate & ", StartContract=" & ChkMsk(mskStartContract) & ", EndContract=" & ChkMsk(mskEndContract) & ", NoofMonthsHired=" & IIf(txtMonthsHired.Text = "", 5, txtMonthsHired.Text) & ", cbamount=" & IIf(txtCBtotal = "", 0, Format(txtCBtotal, "##0.00")) & ", cbpayment=" & IIf(txtCBpayment = "", 0, Format(txtCBpayment, "##0.00")) & ",cbstartpayment=" & ChkMsk(mskStartCB) & ", cbendpayment=" & ChkMsk(mskEndCB) & ", caStartPayment=" & ChkMsk(mskStartCash) & ", caEndPayment=" & ChkMsk(mskEndCash) & ",unistartpayment=" & ChkMsk(mskStartUniform) & ", uniEndpayment=" & ChkMsk(mskEndUniform) & "" & _
       ", MFAmount = " & IIf(TxtMFAmount.Text = "", 0, Format(TxtMFAmount, "##0.00")) & ", MFPayment=" & IIf(TxtMFPayment = "", 0, Format(TxtMFPayment, "##0.00")) & " , MFstartpayment=" & ChkMsk(MskMFStartPayment) & ", MFendpayment=" & ChkMsk(MskMFEndPayment) & "" & _
       ", PFAmount = " & IIf(txtPFAmount = "", 0, Format(txtPFAmount, "##0.00")) & ", PFPayment=" & IIf(txtPFPayment = "", 0, Format(txtPFPayment, "##0.00")) & ", FundPercentage = " & IIf(txtSuretyPercent = "", 0, Format(txtSuretyPercent, "##0.00")) & ", PFSTARTPAYMENT=" & ChkMsk(MskPFStartPayment) & ", PFendpayment=" & ChkMsk(MskPFEndPayment) & "" & _
       ", WithEcolaPay='" & WithEcolaPay & "', WithMiscPay='" & WithMiscPay & "', WithFundPay='" & WithFundPay & "', EcolaAmount=" & IIf(TxtEcolaAmount = "", 0#, Format(TxtEcolaAmount, "###0.00")) & ", MiscFee = " & IIf(txtBillingRate = "", 0, Format(txtBillingRate, "##0.00")) & ", [USERID]='" & operatorID & "', Deptcode=" & DeptCode & ", PicturePath = '" & PictureFile & "', ShiftCode = " & mShiftCode & ", Town = '" & cmbAllowance.Text & "', Remarks = '" & IIf(Me.txtreasonresign.Text = "", " ", txtreasonresign.Text) & "', LastUpdate = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'"
        strWhere = " where  EmployeeID = " & txtID.Text 'EmpNo= '" & TempNewID & "'"  CASHADvance = txtecolaamount.text
        
        If reContract = True And AdminClient = Cid Then
            sqlString = sqlString & ", HRClientCode = " & txtClientCode & " " & sqlString2 & strWhere
        Else
            sqlString = sqlString & sqlString2 & strWhere
        End If
        
        If Action_Query(sqlString, tempConnStr) = True Then
            SuccessFlag = True
'====================================================================================================
            If OrigAcctNo <> txtacctno.Text Then
                Call SaveUserTransactions("Change of Acct# from " & OrigAcctNo & " to " & txtacctno.Text & ".", "Accnt# change.", 0, "", cmbCompany.Text)
            End If
            If origCivilStatus <> cmbcivilstatus.Text Then
                Call SaveUserTransactions("Change of Civil Status from " & origCivilStatus & " to " & cmbcivilstatus.Text & ".", "Civil Status change.", 0, "", cmbCompany.Text)
            End If
            If OrigEmpStatus <> cmbEmpStatus.Text Then
                Call SaveUserTransactions("Change in Employment Status from " & OrigEmpStatus & " to " & cmbEmpStatus.Text & ".", "Employment Status change.", 0, "", cmbCompany.Text)
            End If
            If OrigCompany <> cmbCompany.Text Then
                Call SaveUserTransactions("Change in client from " & OrigCompany & " to " & cmbCompany.Text & ".", "Client change.", 0, "", cmbCompany.Text)
            End If
            If OrigDateHired <> MskDateHired.Text Then
                Call SaveUserTransactions("Change in date hired from " & OrigDateHired & " to " & MskDateHired.Text & ".", "Date Hired change.", 0, "", cmbCompany.Text)
            End If
            If OrigStartContract <> mskStartContract.Text Then
                Call SaveUserTransactions("Change in Start Contract Date from " & OrigStartContract & " to " & mskStartContract.Text & ".", "Start Contract change.", 0, "", cmbCompany.Text)
            End If
            If OrigEndContract <> mskEndContract.Text Then
                Call SaveUserTransactions("Change in End Contract Date from " & OrigEndContract & " to " & mskEndContract.Text & ".", "End Contract change.", 0, "", cmbCompany.Text)
            End If
            If OrigPayBasis <> cmbPayBasis.Text Then
                Call SaveUserTransactions("Change in Payment basis from " & OrigPayBasis & " to " & cmbPayBasis.Text & ".", "Payment Basis change.", 0, "", cmbCompany.Text)
            End If
            If cmbPayBasis.Text = "Daily" Then
                If Format(OrigRate, "0.##") <> Format(txtRate.Text, "0.##") Then
                    Call SaveUserTransactions("Change in Daily Rate from " & OrigRate & " to " & txtRate.Text & ".", "Daily Rate change.", 0, "", cmbCompany.Text)
                End If
            Else
                If Format(OrigMonthly, "0.##") <> Format(txtMonthlyRate.Text, "0.##") Then
                    Call SaveUserTransactions("Change in Monthly Rate from " & OrigMonthly & " to " & txtMonthlyRate.Text & ".", "Monthly Rate change.", 0, "", cmbCompany.Text)
                End If
            End If
            If Format(OrigTASAesla, "0.##") <> Format(txtCBtotal.Text, "0.##") Then
                Call SaveUserTransactions("Change in Cash Bond rate from " & OrigTASAesla & " to " & txtCBtotal.Text & ".", "Cash Bond Rate change.", 0, "", cmbCompany.Text)
            End If
'====================================================================================================
        End If
     End If
     
    'Update EmployeeBasicSalary Table to transfer also to new client if included in payroll list.
    Dim rsemp As ADODB.Recordset
    Dim rsTmp2 As ADODB.Recordset
    
    StrQuery = "Select EmployeeID FROM EmployeeBasicSalary where EmployeeID = " & txtID.Text & " and agencycode = " & tempCompanyCode
    Set rsemp = ReturnRst(StrQuery, tempConnStr, mstatic)
    If rsemp.EOF = False Then
        'Update Basic Salary table with previous entries.
        StrQuery = "UPDATE EmployeeBasicSalary SET clientcode = " & Cid & ", EmpNo = '" & RTrim(TempNewID) & "', " & _
                " Updatedby = '" & LoginName & "', DateUpdated = '" & Format(Now, "mm/dd/yyyy") & "'" & _
                " WHERE EmployeeID = " & txtID.Text & " and agencycode = " & tempCompanyCode
        Call Action_Query(StrQuery, tempConnStr)
    End If
    rsemp.Close
    
    Call EditDetails    ' hrdetails
    Call EditEmpRec     ' hremploymentrecord
    Call EditDialect    ' HRDialects
    Call EditChar       ' HRCharReference
    Call EditSem        ' HRSeminar
    
    ' Check whether employee was consider in-active. Then do not include in payroll list.
    If InStr(1, "ABCPR", Left(cmbEmpStatus.Text, 1)) = 0 Then
        MarkAsInActive
        'MsgBox "[" & txtsurname.Text & ", " & txtfirstname.Text & "] was excluded from payroll list!", vbInformation, PayrollAlert
    End If
    UpdateHREmployeeRecord = True
    rsHREmp.Requery
    rsHREmp.Find " EmployeeID = " & txtID.Text & ""
    ' Showdata rsHREmp

Repeat1:

    Exit Function

endSave:
    
    MsgBox "Encountered Exception Error in updating. The data you want to save might be invalid. Please review your entry.", vbInformation
    
    UpdateHREmployeeRecord = False

End Function

Private Function IsContractSave() As Boolean
Dim CompanyName2 As String
Dim intretval

'        StrDbpath = GetSetting("DataSearch", CompanyShort, "Path2Version2")
'        CompanyName2 = CompanyName '"SAPI Manpower & Mgt. Services"

    If CompanyCode = 2 Then
        StrDbpath = GetSetting("DataSearch", "SAPI", "Path2Version2")
        CompanyName2 = "SAPI Manpower & Mgt. Services"
    ElseIf CompanyCode = 1 Then
        CompanyName2 = "Jobs and Careers Inc."
        StrDbpath = GetSetting("DataSearch", "JCI", "Path2Version2")
    End If
 
    If StrDbpath <> "" And mskStartContract.Text <> "__/__/____" And mskEndContract.Text <> "__/__/____" Then
        CnnstrCopyDestination = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrDbpath
        '   CnnstrCopyDestination = "Driver={Microsoft Access Driver (*.mdb)};" & _
                        "Dbq=" & StrDbpath & ";" & _
                        "Uid=;Pwd=alpha;"
        intretval = OpenDatabaseAdo(cnn, cnnstr)
        If intretval <> 0 Then
            'MsgBox "Unable To Establish the Connection. ", vbCritical
            'Exit Sub
        Else
            If mID(cmbEmpStatus.Text, 1, 1) <> "Z" And mID(cmbEmpStatus.Text, 1, 1) <> "X" And mID(cmbEmpStatus.Text, 1, 1) <> "W" And mID(cmbEmpStatus.Text, 1, 1) <> "T" Then
                StrQuery = "SELECT * FROM HREmployees WHERE EmpNo = '" & TxtControlNo & "' and startcontract=#" & Me.mskStartContract.Text & "# and endcontract=#" & Me.mskEndContract.Text & "# and EmploymentStatus ='" & mID(cmbEmpStatus.Text, 1, 1) & "'"
                Dim rsemp As New ADODB.Recordset
                Set rsemp = ReturnRst(StrQuery, CnnstrCopyDestination, mstatic)
                If rsemp.EOF = False Then
                    MsgBox "CANNOT SAVE." & vbCrLf & " " & vbCrLf & " Same Employee No. and employee status with same contract found  in " & CompanyName2 & " database. Please update the following: Employees status, contract and employee no. ", vbInformation
                  IsContractSave = True
                  Exit Function
                End If
            End If
        End If
    End If
        
    IsContractSave = False
End Function

Private Sub EditSem()
Dim rsSem As ADODB.Recordset

' For Update 2010.
Set rsSem = New ADODB.Recordset

'strQuery = "Select * from HRSeminar where EmployeeID = '" & TempNewID & "'"
'Set rssem = ReturnRst(strQuery, cnnstr, mstatic)
rsSem.Open "Select * from HRSeminar where EmployeeID = " & txtID.Text, cnnstr, adOpenKeyset, adLockOptimistic
'
With rsSem
If .EOF = True Then
    .AddNew
    !EmployeeID = GetEmployeeID()  'TempNewID
End If
       !seminar1 = txtseminar1.Text
       !seminar1date = txtdateseminar1.Text
       !seminar2 = txtseminar2.Text
       !seminar2date = txtdateseminar2.Text
       !seminar3 = txtseminar3.Text
       !seminar3date = txtdateseminar3.Text
     .Update
End With
rsSem.Close
Set rsSem = Nothing

End Sub

Private Sub EditEmpRec()
'On Error Resume Next
Dim rsEmpRec As ADODB.Recordset

' For Update 2010, Bernard.
Set rsEmpRec = New ADODB.Recordset
'strQuery = "Select * from HREmploymentRecord where EmployeeID = '" & TempNewID & "'"
'Set rsemprec = ReturnRst(strQuery, cnnstr, mstatic)
rsEmpRec.Open "Select * from HREmploymentRecord where EmployeeID = " & txtID.Text, cnnstr, adOpenKeyset, adLockOptimistic
'
With rsEmpRec
If .EOF = True Then
    .AddNew
    !EmployeeID = GetEmployeeID() 'TempNewID
End If
       !Company = txtnameofcompany1.Text
       !From = txtfrom1.Text
       !To = txtto1.Text
       !Position = txtpos1.Text
       !reason = txtreason1.Text
       !superior = txtsuperior1.Text
       !company1 = txtnameofcompany2.Text
       !From1 = txtfrom2.Text
       !to1 = txtto2.Text
       !Position1 = txtpos2.Text
       !reason1 = txtreason2.Text
       !superior1 = txtsuperior2.Text
       !company2 = txtnameofcompany3.Text
       !From2 = txtfrom3.Text
       !to2 = txtto3.Text
       !Position2 = txtpos3.Text
       !reason2 = txtreason3.Text
       !superior2 = txtsuperior3.Text
    .Update
End With
rsEmpRec.Close
Set rsEmpRec = Nothing

End Sub

Private Sub EditDialect()
Dim rsDialect As ADODB.Recordset

' For Update 2010, bernard.
Set rsDialect = New ADODB.Recordset
'strQuery = "Select * from HRDialects where EmployeeID = '" & TempNewID & "'"
'Set rsdialect = ReturnRst(strQuery, cnnstr, mstatic)
rsDialect.Open "Select * from HRDialects where EmployeeID = " & txtID.Text, cnnstr, adOpenKeyset, adLockOptimistic
'
With rsDialect
If .EOF = True Then
    .AddNew
    !EmployeeID = GetEmployeeID()  'TempNewID
End If
       !dialect1 = txtdialect1.Text
       !dialect2 = txtdialect2.Text
       !dialect3 = txtdialect3.Text
    .Update
End With
rsDialect.Close
Set rsDialect = Nothing

End Sub

Private Sub EditDetails()
'Dim rsDetails As ADODB.Recordset

'' For Update 2010, Bernard.
Set rsDetails = New ADODB.Recordset
'StrQuery = "Select * FROM HRDetails WHERE EmployeeID = " & TxtID.Text
'Set rsdetails = ReturnRst(StrQuery, cnnstr, mstatic)
rsDetails.Open "Select * from HRDetails where EmployeeID = '" & txtID.Text & "'", cnnstr, adOpenKeyset, adLockOptimistic

''Dim dept1Bdate, dept2Bdate, dept3Bdate, dept4Bdate, dept5Bdate
'    dep1bdate = IIf(mskdep1.Text = "__/__/____", ChkMsk(mskdep1), "#" & Format(mskdep1.Text, "mm/dd/yyyy") & "#")
'    dep2bdate = IIf(mskdep2.Text = "__/__/____", ChkMsk(mskdep2), "#" & Format(mskdep2.Text, "mm/dd/yyyy") & "#")
'    dep3bdate = IIf(mskdep3.Text = "__/__/____", ChkMsk(mskdep3), "#" & Format(mskdep3.Text, "mm/dd/yyyy") & "#")
'    dep4bdate = IIf(mskdep4.Text = "__/__/____", ChkMsk(mskdep4), "#" & Format(mskdep4.Text, "mm/dd/yyyy") & "#")
'    dep5bdate = IIf(mskdep5.Text = "__/__/____", ChkMsk(mskdep5), "#" & Format(mskdep5.Text, "mm/dd/yyyy") & "#")
'
'StrQuery = "UPDATE HRDetails SET " & _
'      " Spousename = " & ChkVal(txtspousename) & ", SpouseAddress = " & ChkVal(txtspouseaddress) & _
'      ", SpouseWork = " & ChkVal(txtsjob) & ", Dependents = " & ChkVal(txtnumdep) & ", dep1name = " & _
'      ChkVal(txtdep1) & ", dep2name = " & ChkVal(txtdep2) & ", dep3name = " & ChkVal(txtdep3) & _
'      ", dep4name = " & ChkVal(txtdep4) & ", dep5name = " & ChkVal(txtdep5) & ", dep1bdate = " & _
'      txtdep1 & ", dep2bdate = " & txtdep2 & ", dep3bdate = " & txtdep3 & ", dep4bdate = " & _
'      txtdep4 & ", dep5bdate = " & txtdep5 & " WHERE EmployeeID = " & txtID.txt
'If Action_Query(StrQuery, cnnstr) = False Then
'    MsgBox "HR details was not successfully saved!", vbInformation, payrollert
'End If

With rsDetails
If .EOF = True Then
    .AddNew
    !EmployeeID = txtID.Text
    !EmpNo = TempNewID
End If
       !SpouseName = IIf(txtspousename.Text <> "", txtspousename.Text, "")
       !spouseaddress = txtspouseaddress.Text
       !spousework = txtsjob.Text
       
       !ContactName = IIf(TxtContactPerson.Text <> "", TxtContactPerson.Text, "")
       !Contactaddress = IIf(txtContactAddress.Text <> "", txtContactAddress.Text, "")
       !Contactphone = IIf(txtContactPhone.Text <> "", txtContactPhone.Text, "")
       
       !Dependents = txtnumdep.Text
       !dep1name = txtdep1.Text
       !dep2name = txtdep2.Text
       !dep3name = txtdep3.Text
       !dep4name = txtdep4.Text
       !dep5name = txtdep5.Text
       !dep1bdate = IIf(mskdep1.Text = "__/__/____", Null, Format(mskdep1.Text, "mm/dd/yyyy"))
       !dep2bdate = IIf(mskdep2.Text = "__/__/____", Null, Format(mskdep2.Text, "mm/dd/yyyy"))
       !dep3bdate = IIf(mskdep3.Text = "__/__/____", Null, Format(mskdep3.Text, "mm/dd/yyyy"))
       !dep4bdate = IIf(mskdep4.Text = "__/__/____", Null, Format(mskdep4.Text, "mm/dd/yyyy"))
       !dep5bdate = IIf(mskdep5.Text = "__/__/____", Null, Format(mskdep5.Text, "mm/dd/yyyy"))
    .Update
End With
rsDetails.Close
    
'Updated by Bernard, May 26, 2012.
StrQuery = "Select * from HRDetails WHERE DetailsFlag = 1"
Set rsDetails = ReturnRst(StrQuery, cnnstr, mstatic)

End Sub

Private Sub EditChar()
Dim rsRef As ADODB.Recordset

' For Update 2010, bernard
Set rsRef = New ADODB.Recordset
'strQuery = "Select * from HRCharReference where EmployeeID='" & TempNewID & "'"
'Set rsref = ReturnRst(strQuery, cnnstr, mstatic)
rsRef.Open "Select * from HRCharReference where EmployeeID = " & txtID.Text, cnnstr, adOpenKeyset, adLockOptimistic
'
With rsRef
If .EOF = True Then
    .AddNew
    !EmployeeID = txtID.Text 'TempNewID
End If
       !Name = txtrefname1.Text
       !occupation = txtrefjob1.Text
       !Address = txtrefaddress1.Text
       !telnum = txtrefphone1.Text
       !Relation = txtrefrel1.Text
       !Name1 = txtrefname2.Text
       !occupation1 = txtrefjob2.Text
       !Address1 = txtrefaddress2.Text
       !TelNum1 = txtrefphone2.Text
       !Relation1 = txtrefrel2.Text
       !Name2 = txtrefname3.Text
       !occupation2 = txtrefjob3.Text
       !Address2 = txtrefaddress3.Text
       !TelNum2 = txtrefphone3.Text
       !Relation2 = txtrefrel3.Text
    .Update
End With
rsRef.Close
Set rsRef = Nothing
End Sub

Private Function ChkVal(ByRef mTBox As TextBox) As String
    Dim mAmt As String
    If Trim(mTBox.Text) = "" Then
        mAmt = Trim(Replace(mTBox, "'", "''"))
        mAmt = "NULL"
    Else
    'Modified by Bernard, replace single quote with double quote for Height field.
        mAmt = "'" & Trim(Replace(mTBox.Text, "'", "''")) & "'"
    End If
    ChkVal = mAmt
End Function

Private Function ChkCmb(ByRef mTBox As ComboBox) As String
    Dim mAmt As String
    If Trim(mTBox.Text) = "" Then
        mAmt = Trim(Replace(mTBox, "'", ""))
        mAmt = "null"
    Else
        mAmt = "'" & mTBox.Text & "'"
    End If
    ChkCmb = mAmt
End Function

Private Function ChkMsk(ByRef mTBox As MaskEdBox) As String
    Dim mAmt As String
    If mTBox.Text = "__/__/____" Then
        mAmt = "null"
    Else
        mAmt = "'" & mTBox.Text & "'"
    End If
    ChkMsk = mAmt
End Function

Private Function ComputeValue()
On Error Resume Next
   Cid = cmbCompany.ItemData(cmbCompany.ListIndex)
    
'    MonthlyRate = txtRate.Text * getField("NoofDaysperMonth", Cid)
'    CommissionRate = 0#
'    PaymentBasis = IIf(cmbPayBasis.Text = "", " ", cmbPayBasis.Text)
'    If cmbPayBasis.ListIndex = 0 Then 'hourly
'            HourlyRate = Format(Val(txtRate.Text), "##0.000")
'            DailyRate = 8 * txtRate.Text
'            MonthlyRate = (8 * txtRate.Text) * getField("NoofDaysperMonth", Cid)
'    ElseIf cmbPayBasis.ListIndex = 1 Then 'daily
'            HourlyRate = txtRate.Text / 8
'            DailyRate = Format(Val(txtRate.Text), "##0.00")
'            MonthlyRate = txtRate.Text * getField("NoofDaysperMonth", Cid)
'    ElseIf cmbPayBasis.ListIndex = 2 Then 'monthly
'            MonthlyRate = Format(Val(txtRate.Text), "##0.00")
'            DailyRate = txtRate.Text * 12 / getField("Workingdaysperyear", Cid)
'            HourlyRate = (txtRate.Text) * 12 / getField("Workingdaysperyear", Cid) / 8
'    ElseIf cmbPayBasis.ListIndex = 3 Then ' Commission Based
'            MonthlyRate = Format(Val(txtRate.Text), "##0.00")
'            CommissionRate = Format(Val(txtRate.Text), "##0.00")
'            DailyRate = 0#
'            HourlyRate = 0#
'    End If
    
    MonthlyRate = (Val(txtRate.Text) * getField("Workingdaysperyear", Cid)) / 12 ' new formula
    CommissionRate = 0#
    PaymentBasis = IIf(cmbPayBasis.Text = "", " ", cmbPayBasis.Text)
    If cmbPayBasis.ListIndex = 0 Then 'hourly
            HourlyRate = Val(txtRate.Text)
            DailyRate = (8 * txtRate.Text)
            MonthlyRate = ((8 * Val(txtRate.Text)) * getField("Workingdaysperyear", Cid)) / 12
    ElseIf cmbPayBasis.ListIndex = 1 Then 'Daily
            HourlyRate = (txtRate.Text / 8)
            DailyRate = Val(txtRate.Text)
            'MonthlyRate = txtRate.Text * getField("NoofDaysperMonth", Cid)
            MonthlyRate = (Val(txtRate.Text) * getField("Workingdaysperyear", Cid)) / 12
    ElseIf cmbPayBasis.ListIndex = 2 Then 'monthly
            MonthlyRate = Val(txtRate.Text)
            DailyRate = (Val(txtRate.Text) * 12) / getField("Workingdaysperyear", Cid)
            HourlyRate = ((Val(txtRate.Text) * 12) / getField("Workingdaysperyear", Cid)) / 8
    ElseIf cmbPayBasis.ListIndex = 3 Then ' Commission Based
            MonthlyRate = Val(txtRate.Text)
            CommissionRate = Val(txtRate.Text)
            DailyRate = 0#
            HourlyRate = 0#
    End If
    
    If MonthlyRate = "" Then MonthlyRate = 0
    If DailyRate = "" Then DailyRate = 0
    If HourlyRate = "" Then HourlyRate = 0
    If CommissionRate = "" Then CommissionRate = 0
        
'    If ChkEcola.VALUE = 1 Then
'        WithEcolaPay = 1
'        EcolaAmount = IIf(TxtEcolaAmount.Text = "", "0.00", Format(TxtEcolaAmount.Text, "##0.00"))
'    Else
'        WithEcolaPay = 0
'        EcolaAmount = 0
'    End If
        
    'EcolaAmount = IIf(TxtEcolaAmount.Text = "", "0.00", Format(TxtEcolaAmount.Text, "##0.00"))
    If chkMisc.VALUE = 1 Then   ' deduc misc
        WithMiscPay = 1
    Else
        WithMiscPay = 0
    End If
    If chkFund.VALUE = 1 Then   ' deduc fund
        WithFundPay = 1
    Else
        WithFundPay = 0
    End If
        
    If cmbBranch.Text = "" Then
       BrCode = 0
    Else
       BrCode = Me.cmbBranch.ItemData(Me.cmbBranch.ListIndex)
    End If
    If cmbDept.Text = "" Then
        DeptCode = 0
    Else
        DeptCode = Me.cmbDept.ItemData(Me.cmbDept.ListIndex)
    End If
    If cmbPosition.Text = "" Then
        PosCode = 0
    Else
        PosCode = Me.cmbPosition.ItemData(Me.cmbPosition.ListIndex)
    End If
    If cmbShiftSchedule.Text = "" Then
        mShiftCode = 0
    Else
        mShiftCode = Me.cmbShiftSchedule.ItemData(Me.cmbShiftSchedule.ListIndex)
    End If
End Function

Public Function MarkAsInActive()
    On Error Resume Next
    sqlString = "UPDATE EmployeeBasicSalary SET Employeestatus = 0 where EmployeeID = " & Me.txtID.Text
    Call Action_Query(sqlString, cnnstr)
End Function

Public Function validate_insert() As Boolean
   Dim rsValidation As ADODB.Recordset
   
   StrQuery = "SELECT * FROM HREmployees WHERE surname='" & RTrim(txtsurname.Text) & "' and firstname='" & RTrim(txtfirstname.Text) & "' and middlename='" & RTrim(txtmiddlename.Text) & "'"
   Set rsValidation = ReturnRst(StrQuery, cnnstr, mstatic)
   'Set rsValidation = ReturnRst("Select * from HREmployees where surname='" & RTrim(txtsurname.Text) & "' and firstname='" & RTrim(txtfirstname.Text) & "'", CnnstrHR, mKeySet)
    
   With rsValidation
        If .EOF = False Then
        MsgBox "CANNOT SAVE because employee no. is already used by [" & rsValidation!Surname & ", " & rsValidation!FirstName & "]" & vbCrLf & "" & vbCrLf & "", vbInformation
        TxtControlNo.SetFocus
        validate_insert = False
        Exit Function
        End If
   End With
   validate_insert = True
End Function

Public Function SaveNewHREmployeeRecord() As Boolean
    
    If validate_insert = False Then Exit Function
    
    Dim rsValidation As ADODB.Recordset
    
    StrQuery = "SELECT * FROM HREmployees WHERE EmpNo = '" & TempNewID & "'"
    Set rsValidation = ReturnRst(StrQuery, cnnstr, mstatic)
    
    With rsValidation
        If .EOF = False Then
            MsgBox "Cannot save because employee no. " + TempNewID + " is already used by [" & rsValidation!Surname & ", " & rsValidation!FirstName & "]" & vbCrLf & "" & vbCrLf & "Please cancel then press again the new button.", vbInformation
            Call GetNewID(True)
        Exit Function
        End If
    End With
    Set rsValidation = Nothing
    
    Call ComputeValue
    sqlString = "INSERT INTO HREmployees (EmpNo, DTRID, AgencyCode, FirstName, Empposition, age, middlename, Surname, nickname, BIRTHDATE, Address1, houseno, street, barangay,district, city, province1, province, contact, caddress,cphone,gender,Height,Weight,Telnum1,Telnum2,TaxCode,examgrade,interviewgrade,medicalgrade,sss,philhealth, pagibig,AcctNo, citizen, birthplace,haircolor, eyecolor, Religion, TIN, [Primary],dateprimary, secondary, datesecondary,college, datecollege, degree, Major, postgraduate," & _
           " datepostgraduate, postgraduatecourse, fathername, faddress, fwork, mothername, maddress, mwork, nbi, rescer, issueat, issueon, datehired, dateprob, dateregular, dateresign, employmentstatus, supervisor, reasonres, endofcontract, ClientCode, brcode, PositionCode, " & _
           " PaymentBasis,HourlyRate, DailyRate, MonthlyRate, CommissionRate, StartContract, EndContract, NoofMonthsHired, cbamount, cbpayment,cbstartpayment, cbendpayment, caStartPayment, caEndPayment, uniStartPayment, uniEndPayment, MFAmount, MFPayment, MFStartPayment, MFEndPayment, PFAmount, PFpayment, FundPercentage, PFstartpayment,PFendpayment, WithEcolaPay, WithMiscPay, WithFundPay, CashAdvance, EcolaAmount, MiscFee, [USERID], DeptCode, PicturePath, ShiftCode, remarks, Town, LastUpdate) " & _
           " values ( '" & TempNewID & "', '" & txtDTRID.Text & "', " & CompanyCode & "," & ChkVal(txtfirstname) & "," & ChkVal(txtfirstchoice) & "," & IIf(txtage.Text = "", 0, txtage.Text) & "," & ChkVal(txtmiddlename) & "," & ChkVal(txtsurname) & "," & ChkVal(txtnickname) & "," & ChkMsk(mskbday) & "," & ChkVal(txtaddress1) & "," & ChkVal(txthouseno) & "," & ChkVal(txtstreet) & "," & ChkVal(txtbarangay) & "," & ChkVal(txtdistrict) & "," & ChkVal(txtcity) & "," & ChkVal(txtprovince) & "" & _
           "," & ChkVal(txtprovinceaddress) & "," & ChkVal(txtemergency) & "," & ChkVal(txteraddress) & "," & ChkVal(txtertel) & ",'" & IIf(optmale.VALUE, "M", "F") & "','" & Replace(txtheight, "'", """") & "'," & ChkVal(txtweight) & "," & ChkVal(txttel1) & "," & ChkVal(txttel2) & ",'" & IIf(cmbcivilstatus.Text = "", "S", cmbcivilstatus.Text) & "'," & ChkVal(txtexam) & "," & ChkVal(txtinterview) & "," & ChkVal(txtmedical) & "," & ChkMsk(msksss) & "," & ChkVal(txtphilhealth) & "," & ChkVal(txtpagibig) & "," & ChkVal(txtacctno) & "" & _
           "," & ChkVal(txtcitizen) & "," & ChkVal(txtbirthplace) & "," & ChkVal(txthaircolor) & "," & ChkVal(txteyecolor) & "," & ChkVal(txtreligion) & "," & ChkMsk(msktax) & "," & ChkVal(txtprimarysch) & "," & ChkVal(txtdateprimary) & "," & ChkVal(txtsecsch) & "," & ChkVal(txtdatesec) & "," & ChkVal(txtcollegesch) & "," & ChkVal(txtdatecollege) & "," & ChkVal(txtcollegecourse) & "," & ChkVal(txtcollegemajor) & "," & ChkVal(txtpostsch) & "," & ChkVal(txtdatepost) & "" & _
           "," & ChkVal(txtpostcourse) & "," & ChkVal(txtfathername) & "," & ChkVal(txtfaddress) & "," & ChkVal(txtfjob) & "," & ChkVal(txtmothername) & "," & ChkVal(txtmaddress) & "," & ChkVal(txtmjob) & "," & ChkVal(txtnbi) & "," & ChkVal(txtrescer) & "," & ChkVal(txtissueat) & "," & ChkMsk(mskissue) & "," & ChkMsk(MskDateHired) & "," & ChkMsk(mskdateprob) & "," & ChkMsk(mskdatereg) & "" & _
           "," & ChkMsk(mskdateres) & ",'" & Left(cmbEmpStatus.Text, 1) & "'," & ChkVal(txtsupervisor) & "," & ChkVal(txtreasonresign) & "," & ChkMsk(mskendof) & "," & Cid & "," & BrCode & ", " & PosCode & ",'" & PaymentBasis & "'," & HourlyRate & "," & DailyRate & "," & MonthlyRate & "," & CommissionRate & "," & ChkMsk(mskStartContract) & "," & ChkMsk(mskEndContract) & "," & IIf(txtMonthsHired.Text = "", 5, txtMonthsHired.Text) & "" & _
           "," & IIf(txtCBtotal = "", 0, Format(txtCBtotal, "##0.00")) & "," & IIf(txtCBpayment = "", 0, Format(txtCBpayment, "##0.00")) & "," & ChkMsk(mskStartCB) & "," & ChkMsk(mskEndCB) & "," & ChkMsk(mskStartCash) & "," & ChkMsk(mskEndCash) & "," & ChkMsk(mskStartUniform) & "," & ChkMsk(mskEndUniform) & "" & _
           "," & IIf(TxtMFAmount = "", 0, Format(TxtMFAmount, "##0.00")) & "," & IIf(TxtMFPayment = "", 0, Format(TxtMFPayment, "##0.00")) & "," & ChkMsk(MskMFStartPayment) & "," & ChkMsk(MskMFEndPayment) & "" & _
            "," & IIf(txtPFAmount = "", 0, Format(txtPFAmount, "##0.00")) & "," & IIf(txtPFPayment = "", 0, Format(txtPFPayment, "##0.00")) & "," & IIf(txtSuretyPercent = "", 0, Format(txtSuretyPercent, "##0.00")) & "," & ChkMsk(MskPFStartPayment) & "," & ChkMsk(MskPFEndPayment) & "" & _
           ",'" & WithEcolaPay & "','" & WithMiscPay & "','" & WithFundPay & "'," & IIf(TxtEcolaAmount = "", 0, Format(TxtEcolaAmount, "##0.00")) & ",0 , " & IIf(txtBillingRate = "", 0, Format(txtBillingRate, "##0.00")) & ", '" & operatorID & "'," & DeptCode & ", '" & PictureFile & "', " & mShiftCode & ", '" & IIf(Me.txtreasonresign.Text = "", "", txtreasonresign.Text) & "', '" & cmbAllowance.Text & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
    Call Action_Query(sqlString, cnnstr)
    
    'TxtControlNo.Text = TempNewID
    txtID.Text = GetEmployeeID()
        'Original Code.
'        sqlString = "Insert into hremployees (EmployeeID, EmpNo , AGENCYCODE, FirstName, Empposition, age, middlename, Surname, nickname, BIRTHDATE, Address1, houseno, street, barangay,district, city,province1, province, contact, caddress,cphone,gender,Height,Weight,Telnum1,Telnum2,TaxCode,examgrade,interviewgrade,medicalgrade,sss,philhealth, pagibig,AcctNo, citizen, birthplace,haircolor, eyecolor, Religion, TIN, [Primary],dateprimary, secondary, datesecondary,college, datecollege, degree, Major, postgraduate," & _
'         " datepostgraduate, postgraduatecourse, fathername, faddress, fwork, mothername, maddress, mwork, nbi, rescer, issueat, issueon, datehired, dateprob, dateregular, dateresign, employmentstatus, supervisor, reasonres, endofcontract, ClientCode, brcode, " & _
'         " PaymentBasis,HourlyRate,DailyRate, MonthlyRate, StartContract, EndContract, NoofMonthsHired, cbamount, cbpayment,cbstartpayment, cbendpayment,  MiscFee,MFPayment,MFStartPayment,MFEndPayment, CashAdvance,PFpayment,PFstartpayment,PFendpayment, WithEcolaPay,  ecolaamount, BillingRate, [USERID], deptcode,remarks, LastUpdate) " & _
'         " values (" & txtID.Text & ",'" & TempNewID & "'," & CompanyCode & "," & ChkVal(txtfirstname) & "," & ChkVal(txtfirstchoice) & "," & IIf(txtage.Text = "", 0, txtage.Text) & "," & ChkVal(txtmiddlename) & "," & ChkVal(txtsurname) & "," & ChkVal(txtnickname) & "," & ChkMsk(mskbday) & "," & ChkVal(txtaddress1) & "," & ChkVal(txthouseno) & "," & ChkVal(txtstreet) & "," & ChkVal(txtbarangay) & "," & ChkVal(txtdistrict) & "," & ChkVal(txtcity) & "," & ChkVal(txtprovince) & "" & _
'         "," & ChkVal(txtprovinceaddress) & "," & ChkVal(txtemergency) & "," & ChkVal(txteraddress) & "," & ChkVal(txtertel) & ",'" & IIf(optmale.VALUE, "M", "F") & "','" & Replace(txtheight, "'", """") & "'," & ChkVal(txtweight) & "," & ChkVal(txttel1) & "," & ChkVal(txttel2) & ",'" & IIf(cmbcivilstatus.Text = "", "S", cmbcivilstatus.Text) & "'," & ChkVal(txtexam) & "," & ChkVal(txtinterview) & "," & ChkVal(txtmedical) & "," & ChkMsk(msksss) & "," & ChkVal(txtphilhealth) & "," & ChkVal(txtpagibig) & "," & ChkVal(txtacctno) & "" & _
'         "," & ChkVal(txtcitizen) & "," & ChkVal(txtbirthplace) & "," & ChkVal(txthaircolor) & "," & ChkVal(txteyecolor) & "," & ChkVal(txtreligion) & "," & ChkMsk(msktax) & "," & ChkVal(txtprimarysch) & "," & ChkVal(txtdateprimary) & "," & ChkVal(txtsecsch) & "," & ChkVal(txtdatesec) & "," & ChkVal(txtcollegesch) & "," & ChkVal(txtdatecollege) & "," & ChkVal(txtcollegecourse) & "," & ChkVal(txtcollegemajor) & "," & ChkVal(txtpostsch) & "," & ChkVal(txtdatepost) & "" & _
'         "," & ChkVal(txtpostcourse) & "," & ChkVal(txtfathername) & "," & ChkVal(txtfjob) & "," & ChkVal(txtfaddress) & "," & ChkVal(txtmothername) & "," & ChkVal(txtmaddress) & "," & ChkVal(txtmjob) & "," & ChkVal(txtnbi) & "," & ChkVal(txtrescer) & "," & ChkVal(txtissueat) & "," & ChkMsk(mskissue) & "," & ChkMsk(MskDateHired) & "," & ChkMsk(mskdateprob) & "," & ChkMsk(mskdatereg) & "" & _
'         "," & ChkMsk(mskdateres) & ",'" & Left(cmbempstatus.Text, 1) & "'," & ChkVal(txtsupervisor) & "," & ChkVal(txtreasonresign) & "," & ChkMsk(mskendof) & "," & Cid & "," & BrCode & ",'" & PaymentBasis & "'," & HourlyRate & "," & DailyRate & "," & MonthlyRate & "," & ChkMsk(mskStartContract) & "," & ChkMsk(mskEndContract) & "," & IIf(txtMonthsHired.Text = "", 5, txtMonthsHired.Text) & "" & _
'         "," & IIf(txtCBtotal = "", 0, Format(txtCBtotal, "##0.00")) & "," & IIf(txtCBpayment = "", 0, Format(txtCBpayment, "##0.00")) & "," & ChkMsk(mskStartCB) & "," & ChkMsk(mskEndCB) & "" & _
'         "," & IIf(TxtMFAmount = "", 0, Format(TxtMFAmount, "##0.00")) & "," & IIf(TxtMFPayment = "", 0, Format(TxtMFPayment, "##0.00")) & "," & ChkMsk(MskMFStartPayment) & "," & ChkMsk(MskMFEndPayment) & "" & _
'          "," & IIf(txtPFAmount = "", 0, Format(txtPFAmount, "##0.00")) & "," & IIf(txtPFPayment = "", 0, Format(txtPFPayment, "##0.00")) & "," & ChkMsk(MskPFStartPayment) & "," & ChkMsk(MskPFEndPayment) & "" & _
'         ",'" & WithEcolaPay & "'," & EcolaAmount & ", " & IIf(txtBillingRate = "", 0, Format(txtBillingRate, "##0.00")) & ", '" & operatorID & "'," & DeptCode & ",'" & IIf(Me.txtreasonresign.Text = "", " ", txtreasonresign.Text) & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
    
'        sqlString = "Insert into hremployees (EmpNo , AGENCYCODE, FirstName, Empposition, age, middlename, Surname, nickname, BIRTHDATE, Address1, houseno, street, barangay,district, city,province1, province, contact, caddress,cphone,gender,Height,Weight,Telnum1,Telnum2,TaxCode,examgrade,interviewgrade,medicalgrade,sss,philhealth, pagibig,AcctNo, citizen, birthplace,haircolor, eyecolor, Religion, TIN, [Primary],dateprimary, secondary, datesecondary,college, datecollege, degree, Major, postgraduate," & _
'         " datepostgraduate, postgraduatecourse, fathername, faddress, fwork, mothername, maddress, mwork, nbi, rescer, issueat, issueon, datehired, dateprob, dateregular, dateresign, employmentstatus, supervisor, reasonres, endofcontract, ClientCode, brcode, " & _
'         " PaymentBasis,HourlyRate,DailyRate, MonthlyRate, StartContract, EndContract, NoofMonthsHired, cbamount, cbpayment,cbstartpayment, cbendpayment,  MiscFee,MFPayment,MFStartPayment,MFEndPayment, CashAdvance,PFpayment,PFstartpayment,PFendpayment, WithEcolaPay,  ecolaamount, BillingRate, [USERID], deptcode,remarks, LastUpdate) " & _
'         " VALUES ('" & TempNewID & "'," & CompanyCode & "," & ChkVal(txtfirstname) & "," & ChkVal(txtfirstchoice) & "," & IIf(txtage.Text = "", 0, txtage.Text) & "," & ChkVal(txtmiddlename) & "," & ChkVal(txtsurname) & "," & ChkVal(txtnickname) & "," & ChkMsk(mskbday) & "," & ChkVal(txtaddress1) & "," & ChkVal(txthouseno) & "," & ChkVal(txtstreet) & "," & ChkVal(txtbarangay) & "," & ChkVal(txtdistrict) & "," & ChkVal(txtcity) & "," & ChkVal(txtprovince) & "" & _
'         "," & ChkVal(txtprovinceaddress) & "," & ChkVal(txtemergency) & "," & ChkVal(txteraddress) & "," & ChkVal(txtertel) & ",'" & IIf(optmale.VALUE, "M", "F") & "','" & Replace(txtheight, "'", """") & "'," & ChkVal(txtweight) & "," & ChkVal(txttel1) & "," & ChkVal(txttel2) & ",'" & IIf(cmbcivilstatus.Text = "", "S", cmbcivilstatus.Text) & "'," & ChkVal(txtexam) & "," & ChkVal(txtinterview) & "," & ChkVal(txtmedical) & "," & ChkMsk(msksss) & "," & ChkVal(txtphilhealth) & "," & ChkVal(txtpagibig) & "," & ChkVal(txtacctno) & "" & _
'         "," & ChkVal(txtcitizen) & "," & ChkVal(txtbirthplace) & "," & ChkVal(txthaircolor) & "," & ChkVal(txteyecolor) & "," & ChkVal(txtreligion) & "," & ChkMsk(msktax) & "," & ChkVal(txtprimarysch) & "," & ChkVal(txtdateprimary) & "," & ChkVal(txtsecsch) & "," & ChkVal(txtdatesec) & "," & ChkVal(txtcollegesch) & "," & ChkVal(txtdatecollege) & "," & ChkVal(txtcollegecourse) & "," & ChkVal(txtcollegemajor) & "," & ChkVal(txtpostsch) & "," & ChkVal(txtdatepost) & "" & _
'         "," & ChkVal(txtpostcourse) & "," & ChkVal(txtfathername) & "," & ChkVal(txtfjob) & "," & ChkVal(txtfaddress) & "," & ChkVal(txtmothername) & "," & ChkVal(txtmaddress) & "," & ChkVal(txtmjob) & "," & ChkVal(txtnbi) & "," & ChkVal(txtrescer) & "," & ChkVal(txtissueat) & "," & ChkMsk(mskissue) & "," & ChkMsk(MskDateHired) & "," & ChkMsk(mskdateprob) & "," & ChkMsk(mskdatereg) & "" & _
'         "," & ChkMsk(mskdateres) & ",'" & Left(cmbempstatus.Text, 1) & "'," & ChkVal(txtsupervisor) & "," & ChkVal(txtreasonresign) & "," & ChkMsk(mskendof) & "," & Cid & "," & BrCode & ",'" & PaymentBasis & "'," & HourlyRate & "," & DailyRate & "," & MonthlyRate & "," & ChkMsk(mskStartContract) & "," & ChkMsk(mskEndContract) & "," & IIf(txtMonthsHired.Text = "", 5, txtMonthsHired.Text) & "" & _
'         "," & IIf(txtCBtotal = "", 0, Format(txtCBtotal, "##0.00")) & "," & IIf(txtCBpayment = "", 0, Format(txtCBpayment, "##0.00")) & "," & ChkMsk(mskStartCB) & "," & ChkMsk(mskEndCB) & "" & _
'         "," & IIf(TxtMFAmount = "", 0, Format(TxtMFAmount, "##0.00")) & "," & IIf(TxtMFPayment = "", 0, Format(TxtMFPayment, "##0.00")) & "," & ChkMsk(MskMFStartPayment) & "," & ChkMsk(MskMFEndPayment) & "" & _
'          "," & IIf(txtPFAmount = "", 0, Format(txtPFAmount, "##0.00")) & "," & IIf(txtPFPayment = "", 0, Format(txtPFPayment, "##0.00")) & "," & ChkMsk(MskPFStartPayment) & "," & ChkMsk(MskPFEndPayment) & "" & _
'         ",'" & WithEcolaPay & "'," & EcolaAmount & ", " & IIf(txtBillingRate = "", 0, Format(txtBillingRate, "##0.00")) & ", '" & operatorID & "'," & DeptCode & ",'" & IIf(Me.txtreasonresign.Text = "", " ", txtreasonresign.Text) & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
'
'    Call Action_Query(sqlString, cnnstr)
    
    MsgBox "[" & txtsurname.Text & ", " & txtfirstname.Text & " " & txtmiddlename.Text & "] was successfully saved", vbInformation
  
    On Error Resume Next

    'Dependents
    Dim dept1Bdate, dept2Bdate, dept3Bdate, dept4Bdate, dept5Bdate
'    dep1bdate = IIf(mskdep1.Text = "__/__/____", ChkMsk(mskdep1), "#" & Format(mskdep1.Text, "mm/dd/yyyy") & "#")
'    dep2bdate = IIf(mskdep2.Text = "__/__/____", ChkMsk(mskdep2), "#" & Format(mskdep2.Text, "mm/dd/yyyy") & "#")
'    dep3bdate = IIf(mskdep3.Text = "__/__/____", ChkMsk(mskdep3), "#" & Format(mskdep3.Text, "mm/dd/yyyy") & "#")
'    dep4bdate = IIf(mskdep4.Text = "__/__/____", ChkMsk(mskdep4), "#" & Format(mskdep4.Text, "mm/dd/yyyy") & "#")
'    dep5bdate = IIf(mskdep5.Text = "__/__/____", ChkMsk(mskdep5), "#" & Format(mskdep5.Text, "mm/dd/yyyy") & "#")
    
    dep1bdate = IIf(mskdep1.Text = "__/__/____", ChkMsk(mskdep1), Format(mskdep1.Text, "mm/dd/yyyy"))
    dep2bdate = IIf(mskdep2.Text = "__/__/____", ChkMsk(mskdep2), Format(mskdep2.Text, "mm/dd/yyyy"))
    dep3bdate = IIf(mskdep3.Text = "__/__/____", ChkMsk(mskdep3), Format(mskdep3.Text, "mm/dd/yyyy"))
    dep4bdate = IIf(mskdep4.Text = "__/__/____", ChkMsk(mskdep4), Format(mskdep4.Text, "mm/dd/yyyy"))
    dep5bdate = IIf(mskdep5.Text = "__/__/____", ChkMsk(mskdep5), Format(mskdep5.Text, "mm/dd/yyyy"))

    'HR Details
    sqlString = "INSERT INTO hrdetails " & _
      "(EmployeeID,EmpNo,spousename,spouseaddress,spousework,Dependents ,dep1name,dep2name,dep3name,dep4name,dep5name,dep1bdate,dep2bdate,dep3bdate,dep4bdate,dep5bdate, contactname, contactaddress, contactphone )" & _
      " values (" & txtID.Text & ", '" & TempNewID & "'," & ChkVal(txtspousename) & "," & ChkVal(txtspouseaddress) & "," & ChkVal(txtsjob) & "," & ChkVal(txtnumdep) & "," & ChkVal(txtdep1) & "," & ChkVal(txtdep2) & "," & ChkVal(txtdep3) & "," & ChkVal(txtdep4) & "," & ChkVal(txtdep5) & "," & dep1bdate & "," & dep2bdate & "," & dep3bdate & "," & dep4bdate & "," & dep5bdate & _
      ", '" & TxtContactPerson.Text & "', '" & txtContactAddress.Text & "', '" & txtContactPhone.Text & "')"
    Call Action_Query(sqlString, cnnstr)
  
    'HR Reference
    sqlString = "INSERT INTO hrCharReference " & _
      "(EmployeeID, EmpNo, Name, occupation, address, telnum, relation, name1, occupation1, address1, telnum1 " & _
      ", relation1, name2,occupation2,address2,telnum2,relation2 )" & _
      " VALUES (" & txtID.Text & ", '" & TempNewID & "'," & ChkVal(txtrefname1) & "," & ChkVal(txtrefjob1) & "," & ChkVal(txtrefaddress1) & "," & ChkVal(txtrefphone1) & "," & ChkVal(txtrefrel1) & "," & ChkVal(txtrefname2) & "," & ChkVal(txtrefjob2) & "," & ChkVal(txtrefaddress2) & "," & ChkVal(txtrefphone2) & "," & ChkVal(txtrefrel2) & "," & ChkVal(txtrefname3) & "," & ChkVal(txtrefjob3) & "," & ChkVal(txtrefjob3) & "," & ChkVal(txtrefphone3) & "," & ChkVal(txtrefrel3) & ")"
    Call Action_Query(sqlString, cnnstr)

    'HR Employment record
    sqlString = "INSERT INTO hrEmploymentRecord " & _
      "(EmployeeID, EmpNo,company,[from],[to],[position],reason,superior,company1,from1,to1," & _
      " position1,reason1,superior1,company2,from2,to2,position2,reason2,superior2)" & _
      " values (" & txtID.Text & ", '" & TempNewID & "'," & ChkVal(txtnameofcompany1) & "," & ChkVal(txtfrom1) & "," & ChkVal(txtto1) & "," & ChkVal(txtpos1) & "," & ChkVal(txtreason1) & "," & ChkVal(txtsuperior1) & "," & _
                                        " " & ChkVal(txtnameofcompany2) & "," & ChkVal(txtfrom2) & "," & ChkVal(txtto2) & "," & ChkVal(txtpos2) & "," & ChkVal(txtreason2) & "," & ChkVal(txtsuperior2) & "," & _
                                        " " & ChkVal(txtnameofcompany3) & "," & ChkVal(txtfrom3) & "," & ChkVal(txtto3) & "," & ChkVal(txtpos3) & "," & ChkVal(txtreason3) & "," & ChkVal(txtsuperior3) & ")"
    Call Action_Query(sqlString, cnnstr)
           
    'HR Seminar records
    sqlString = "INSERT INTO hrSeminar " & _
      "(EmployeeID, EmpNo,seminar1,seminar1date,seminar2,seminar2date,seminar3,seminar3date)" & _
      " values (" & txtID.Text & ", '" & TempNewID & "'," & ChkVal(txtseminar1) & "," & ChkVal(txtdateseminar1) & "," & _
                                        " " & ChkVal(txtseminar2) & "," & ChkVal(txtdateseminar2) & "," & _
                                        " " & ChkVal(txtseminar3) & "," & ChkVal(txtdateseminar3) & ")"
   Call Action_Query(sqlString, cnnstr)
               
    'HR Dialects records.
    sqlString = "INSERT INTO HRDialects " & _
      "(EmployeeID, EmpNo,dialect1,dialect2,dialect3) " & _
      " values (" & txtID.Text & ", '" & TempNewID & "'," & ChkVal(txtdialect1) & "," & ChkVal(txtdialect2) & "," & ChkVal(txtdialect3) & ")"
        Call Action_Query(sqlString, cnnstr)
           
'    If rsHREmp.State = 1 Then rsHREmp.Close
    Set rsHREmp = New ADODB.Recordset
    'StrQuery = "Select * from HREmployees where agencycode=" & CompanyCode & " order by  surname"
    If CurrentClientCode = 0 Then
        StrQuery = "SELECT * FROM HREmployees WHERE agencycode = " & CompanyCode & "  order by  surname, firstname"
    Else
        StrQuery = "SELECT * FROM HREmployees WHERE agencycode = " & CompanyCode & " AND ClientCode = " & CurrentClientCode & "  order by  surname, firstname"
    End If
    
    Set rsHREmp = ReturnRst(StrQuery, cnnstr, mstatic)
    'rsHREmp.Open "Select * from HREmployees where agencycode=" & CompanyCode & " order by  surname", cnnHR, adOpenKeyset, adLockOptimistic
    
    SaveNewHREmployeeRecord = True
    Exit Function
    
endSave:
    
    MsgBox "Encountered Exception in saving. The data you want to save might be invalid. Please review your entry.", vbInformation
    SaveNewHREmployeeRecord = False
    
End Function

Public Function getField(ByVal fldname As String, ByVal clientID As Integer) As Integer
    Dim rs As ADODB.Recordset
   
'    If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'        StrQuery = "Select " & fldname & " From clienttable where CLIENTCODE2 = 'FMS' AND clientCODE = " & clientID & ""
'    Else
'        StrQuery = "Select " & fldname & " From clienttable where clientCODE = " & clientID & ""
'    End If
    StrQuery = GetClientRecord(clientID, , , fldname)
    Set rs = ReturnRst(StrQuery, cnnstr, mstatic)
    If Not rs.EOF Then
       getField = rs(0)
    End If
    
    Set rs = Nothing
    
End Function

Private Sub SetButtonsave()
    cmdNew.Enabled = True
    cmdedit.Enabled = True
    cmdupdate.Enabled = False
    cmddelete.Enabled = True
    cmdfirst.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdlast.Enabled = True
    cmdundo.Enabled = False
    btnEmpRefresh.Enabled = True
    btnEmployeeSearch.Enabled = True
    cmdRecontract.Enabled = True
End Sub

Public Sub RefreshBranch()

'Set rstBranch = cnnHR.Execute("select * from HRBRANCH")
StrQuery = "SELECT * FROM HRWorkStation ORDER BY BRName"
Set rstBranch = ReturnRst(StrQuery, cnnstr, mstatic)
'Set rstBranch = ReturnRst("SELECT * FROM HRBRANCH WHERE BrStatus = True order by BRName", cnnstr) ' updated by Bernard 01/02/09

cmbBranch.Clear
 Do While Not rstBranch.EOF
    cmbBranch.AddItem UCase(IIf(IsNull(rstBranch!BrName), "", rstBranch!BrName))
    cmbBranch.ItemData(cmbBranch.NewIndex) = rstBranch!BrCode
    rstBranch.MoveNext
Loop
    rstBranch.Close
End Sub

Public Sub RefreshDept()
cmbDept.Clear

'Set rstDept = cnnHR.Execute("select * from HRDEPT Order by DeptName")
StrQuery = "SELECT * FROM HRDEPT order by deptname"
Set rstDept = ReturnRst(StrQuery, cnnstr, mstatic)
'Set rstDept = ReturnRst("select * from HRDEPT order by deptname", cnnstr) ' updated by Bernard 10/17/08

 Do While Not rstDept.EOF
    cmbDept.AddItem UCase(rstDept!deptname)
    cmbDept.ItemData(cmbDept.NewIndex) = rstDept!DeptCode
    rstDept.MoveNext
Loop
rstDept.Close
End Sub

Public Sub RefreshPosition()
cmbPosition.Clear

StrQuery = "Select * from HRPositions order by positionname"
Set rstDept = ReturnRst(StrQuery, cnnstr, mstatic)

 Do While Not rstDept.EOF
    cmbPosition.AddItem UCase(rstDept!Positionname)
    cmbPosition.ItemData(cmbPosition.NewIndex) = rstDept!Positioncode
    rstDept.MoveNext
Loop
rstDept.Close
End Sub

Public Function CheckUser1(ByVal operatorID As Integer)
Dim RsUsers As ADODB.Recordset

Set RsUsers = New ADODB.Recordset
StrQuery = "Select * from userslist Where operatorid = " & operatorID & ""
Set RsUsers = ReturnRst(StrQuery, cnnstr, mstatic)

' RsUsers.Open "select * from userslist where operatorid=" & operatorID & "", cnnstr, adOpenDynamic, adLockOptimistic
 
With RsUsers
    If .Fields("maintenancerights") = True Then
         cmdAddBranch.Enabled = True
         cmdAddDept.Enabled = True
         cmdAddPosition.Enabled = True
    End If
End With

End Function

Public Sub Showdialect(ByRef rec1 As ADODB.Recordset)
    Dim rsDialect As ADODB.Recordset
    txtdialect1 = IIf(IsNull(rec1!dialect1), "", rec1!dialect1)
    txtdialect2 = IIf(IsNull(rec1!dialect2), "", rec1!dialect2)
    txtdialect3 = IIf(IsNull(rec1!dialect3), "", rec1!dialect3)
End Sub

Public Sub Showdetails(Rec2 As ADODB.Recordset)
Dim rsDetails As ADODB.Recordset
   txtspousename.Text = IIf(IsNull(Rec2!SpouseName), "", Rec2!SpouseName)
   txtspouseaddress.Text = IIf(IsNull(Rec2!spouseaddress), "", Rec2!spouseaddress)
   txtsjob.Text = IIf(IsNull(Rec2!spousework), "", Rec2!spousework)
   TxtContactPerson.Text = IIf(IsNull(Rec2!ContactName), "", Rec2!ContactName)
   txtContactAddress.Text = IIf(IsNull(Rec2!Contactaddress), "", Rec2!Contactaddress)
   txtContactPhone.Text = IIf(IsNull(Rec2!Contactphone), "", Rec2!Contactphone)
   
   txtnumdep.Text = IIf(IsNull(Rec2!Dependents), 0, Rec2!Dependents)
   txtdep1.Text = IIf(IsNull(Rec2!dep1name), "", Rec2!dep1name)
   txtdep2.Text = IIf(IsNull(Rec2!dep2name), "", Rec2!dep2name)
   txtdep3.Text = IIf(IsNull(Rec2!dep3name), "", Rec2!dep3name)
   txtdep4.Text = IIf(IsNull(Rec2!dep4name), "", Rec2!dep4name)
  txtdep5.Text = IIf(IsNull(Rec2!dep5name), "", Rec2!dep5name)
   
   mskdep1.Text = IIf(IsNull(Rec2!dep1bdate), "__/__/____", Format(Rec2!dep1bdate, "mm/dd/yyyy"))
   mskdep2.Text = IIf(IsNull(Rec2!dep2bdate), "__/__/____", Format(Rec2!dep2bdate, "mm/dd/yyyy"))
   mskdep3.Text = IIf(IsNull(Rec2!dep3bdate), "__/__/____", Format(Rec2!dep3bdate, "mm/dd/yyyy"))
   mskdep4.Text = IIf(IsNull(Rec2!dep4bdate), "__/__/____", Format(Rec2!dep4bdate, "mm/dd/yyyy"))
   mskdep5.Text = IIf(IsNull(Rec2!dep5bdate), "__/__/____", Format(Rec2!dep5bdate, "mm/dd/yyyy"))
End Sub
Public Sub ShowRef(rec3 As ADODB.Recordset)
Dim rsRef As ADODB.Recordset
    txtrefname1.Text = IIf(IsNull(rec3!Name), "", rec3!Name)
    txtrefjob1.Text = IIf(IsNull(rec3!occupation), "", rec3!occupation)
    txtrefaddress1.Text = IIf(IsNull(rec3!Address), "", rec3!Address)
    txtrefphone1.Text = IIf(IsNull(rec3!telnum), "", rec3!telnum)
    txtrefrel1.Text = IIf(IsNull(rec3!Relation), "", rec3!Relation)

    txtrefname2.Text = IIf(IsNull(rec3!Name1), "", rec3!Name1)
    txtrefjob2.Text = IIf(IsNull(rec3!occupation1), "", rec3!occupation1)
    txtrefaddress2.Text = IIf(IsNull(rec3!Address1), "", rec3!Address1)
    txtrefphone2.Text = IIf(IsNull(rec3!TelNum1), "", rec3!TelNum1)
    txtrefrel2.Text = IIf(IsNull(rec3!Relation1), "", rec3!Relation1)

    txtrefname3.Text = IIf(IsNull(rec3!Name2), "", rec3!Name2)
    txtrefjob3.Text = IIf(IsNull(rec3!occupation2), "", rec3!occupation2)
    txtrefaddress3.Text = IIf(IsNull(rec3!Address2), "", rec3!Address2)
    txtrefphone3.Text = IIf(IsNull(rec3!TelNum2), "", rec3!TelNum2)
    txtrefrel3.Text = IIf(IsNull(rec3!Relation2), "", rec3!Relation2)

End Sub
Public Sub ShowEmpRec(rec4 As ADODB.Recordset)
Dim rsEmpRec As ADODB.Recordset
    txtnameofcompany1.Text = IIf(IsNull(rec4!Company), "", rec4!Company)
    txtfrom1.Text = IIf(IsNull(rec4!From), "", rec4!From)
    txtto1.Text = IIf(IsNull(rec4!To), "", rec4!To)
    txtpos1.Text = IIf(IsNull(rec4!Position), "", rec4!Position)
    txtreason1.Text = IIf(IsNull(rec4!reason), "", rec4!reason)
    txtsuperior1.Text = IIf(IsNull(rec4!superior), "", rec4!superior)

    txtnameofcompany2.Text = IIf(IsNull(rec4!company1), "", rec4!company1)
    txtfrom2.Text = IIf(IsNull(rec4!From1), "", rec4!From1)
    txtto2.Text = IIf(IsNull(rec4!to1), "", rec4!to1)
    txtpos2.Text = IIf(IsNull(rec4!Position1), "", rec4!Position1)
    txtreason2.Text = IIf(IsNull(rec4!reason1), "", rec4!reason1)
    txtsuperior2.Text = IIf(IsNull(rec4!superior1), "", rec4!superior1)

    txtnameofcompany3.Text = IIf(IsNull(rec4!company2), "", rec4!company2)
    txtfrom3.Text = IIf(IsNull(rec4!From2), "", rec4!From2)
    txtto3.Text = IIf(IsNull(rec4!to2), "", rec4!to2)
    txtpos3.Text = IIf(IsNull(rec4!Position2), "", rec4!Position2)
    txtreason3.Text = IIf(IsNull(rec4!reason2), "", rec4!reason2)
    txtsuperior3.Text = IIf(IsNull(rec4!superior2), "", rec4!superior2)

End Sub
Public Sub ShowSem(rec5 As ADODB.Recordset)

    txtseminar1.Text = IIf(IsNull(rec5!seminar1), "", rec5!seminar1)
    txtdateseminar1.Text = IIf(IsNull(rec5!seminar1date), "", rec5!seminar1date)

    txtseminar2.Text = IIf(IsNull(rec5!seminar2), "", rec5!seminar2)
    txtdateseminar2.Text = IIf(IsNull(rec5!seminar2date), "", rec5!seminar2date)

    txtseminar3.Text = IIf(IsNull(rec5!seminar3), "", rec5!seminar3)
    txtdateseminar3.Text = IIf(IsNull(rec5!seminar3date), "", rec5!seminar3date)

End Sub

Private Sub ClearTextBox(Optional ByVal reContract As Boolean = False)
    
    If reContract = False Then
        txtsurname.Text = ""
        txtfirstname.Text = ""
        txtmiddlename.Text = ""
        TxtControlNo.Text = ""
          
        txtMonthsHired.Text = "5"
        TxtEcolaAmount.Text = "0.00"
        txtBillingRate.Text = "0.00"
        txtCBtotal.Text = "0.00"
        txtCBpayment.Text = "0.00"
        mskStartCB.Text = "__/__/____"
        mskEndCB.Text = "__/__/____"
        
        mskStartCash.Text = "__/__/____"
        mskEndCash.Text = "__/__/____"
        mskStartUniform.Text = "__/__/____"
        mskEndUniform.Text = "__/__/____"

        MskMFStartPayment.Text = "__/__/____"
        MskMFEndPayment.Text = "__/__/____"
        MskPFStartPayment.Text = "__/__/____"
        MskPFEndPayment.Text = "__/__/____"
        TxtMFAmount.Text = "0.00"
        TxtMFPayment.Text = "0.00"
        txtPFPayment.Text = "0.00"
        txtSuretyPercent.Text = "0.00"
        txtPFAmount.Text = "0.00"
        
        txtprimarysch.Text = ""
        txtdateprimary.Text = ""
        txtnickname.Text = ""
        txtfirstchoice.Text = ""
        mskbday.Text = "__/__/____"
        txtage.Text = ""
        txtaddress1.Text = ""
        txthouseno.Text = ""
        txtstreet.Text = ""
        txtbarangay.Text = ""
        txtdistrict.Text = ""
        txtcity.Text = ""
        txtprovince.Text = ""
        txtprovinceaddress.Text = ""
        txtemergency.Text = ""
        txteraddress.Text = ""
        txtertel.Text = ""
        optmale.VALUE = False
        optfemale.VALUE = False
        txtheight.Text = ""
        txtweight.Text = ""
        txttel1.Text = ""
        txttel2.Text = ""
        txtexam.Text = ""
        txtinterview.Text = ""
        txtmedical.Text = ""
        msksss.Text = "__-_______-_"
        txtphilhealth.Text = ""
        txtpagibig.Text = ""
        txtacctno.Text = ""
        txtcitizen.Text = ""
        txtbirthplace.Text = ""
        txthaircolor.Text = ""
        txteyecolor.Text = ""
        txtreligion.Text = ""
        msktax.Text = "___-___-___"
        mskissue.Text = "__/__/____"
        txtpostsch.Text = ""
        txtdatepost.Text = ""
        txtsecsch.Text = ""
        txtdatesec.Text = ""
        txtcollegesch.Text = ""
        txtdatecollege.Text = ""
        txtcollegecourse.Text = ""
        txtcollegemajor.Text = ""
        txtdatepost.Text = ""
        txtpostcourse.Text = ""
        txtfathername.Text = ""
        txtfaddress.Text = ""
        txtfjob.Text = ""
        txtmothername.Text = ""
        txtmaddress.Text = ""
        txtmjob.Text = ""
        txtspousename.Text = ""
        txtspouseaddress.Text = ""
        txtsjob.Text = ""
        
        TxtContactPerson.Text = ""
        txtContactAddress.Text = ""
        txtContactPhone.Text = ""
        
        txtnbi.Text = ""
        txtrescer.Text = ""
        txtissueat.Text = ""
        txtdep1.Text = ""
        txtdep2.Text = ""
        txtdep3.Text = ""
        txtdep4.Text = ""
        txtdep5.Text = ""
        mskdep1.Text = "__/__/____"
        mskdep2.Text = "__/__/____"
        mskdep3.Text = "__/__/____"
        mskdep4.Text = "__/__/____"
        mskdep5.Text = "__/__/____"
        txtnumdep.Text = 0
        txtdialect1.Text = ""
        txtdialect2.Text = ""
        txtdialect3.Text = ""
        txtseminar1.Text = ""
        txtseminar2.Text = ""
        txtseminar3.Text = ""
        txtdateseminar1.Text = ""
        txtdateseminar2.Text = ""
        txtdateseminar3.Text = ""
        txtnameofcompany1.Text = ""
        txtnameofcompany2.Text = ""
        txtnameofcompany3.Text = ""
        txtfrom1.Text = ""
        txtfrom2.Text = ""
        txtfrom3.Text = ""
        txtto1.Text = ""
        txtto2.Text = ""
        txtto3.Text = ""
        txtpos1.Text = ""
        txtpos2.Text = ""
        txtpos3.Text = ""
        txtreason1.Text = ""
        txtreason2.Text = ""
        txtreason3.Text = ""
        txtsuperior1.Text = ""
        txtsuperior2.Text = ""
        txtsuperior3.Text = ""
        txtrefname1.Text = ""
        txtrefname2.Text = ""
        txtrefname3.Text = ""
        txtrefjob1.Text = ""
        txtrefjob2.Text = ""
        txtrefjob3.Text = ""
        txtrefaddress1.Text = ""
        txtrefaddress2.Text = ""
        txtrefaddress3.Text = ""
        txtrefphone1.Text = ""
        txtrefphone2.Text = ""
        txtrefphone3.Text = ""
        txtrefrel1.Text = ""
        txtrefrel2.Text = ""
        txtrefrel3.Text = ""
        mskdateprob.Text = "__/__/____"
        MskDateHired.Text = "__/__/____"
        mskdatereg.Text = "__/__/____"
        mskdateres.Text = "__/__/____"
        
    End If

    txtRate.Text = "0.00"
    txtMonthlyRate.Text = "0.00"
   
    mskStartContract.Text = "__/__/____"
    mskEndContract.Text = "__/__/____"
        
    
    cmbcivilstatus.ListIndex = -1
    cmbEmpStatus.ListIndex = -1

    cmbEmpStatus.ListIndex = -1
    txtsupervisor.Text = ""
    txtreasonresign.Text = ""
    mskendof.Text = "__/__/____"
    
    cmbPayBasis.ListIndex = -1
  
   
'    If Not Trim(projectlong) = "" Then
'        cmbCompany.Text = projectlong
'    Else
'        cmbCompany.ListIndex = -1
'    End If
    cmbBranch.ListIndex = -1
    cmbDept.ListIndex = -1
    cmbPosition.ListIndex = -1
    cmbShiftSchedule.ListIndex = -1    ' new code
End Sub

Private Sub mskbday_GotFocus()
mskbday.BackColor = &HFFFFC0
End Sub

Private Sub mskbday_LostFocus()
mskbday.BackColor = vbWhite
If mskbday.Text <> "__/__/____" Then
    txtage.Text = DateDiff("yyyy", Format(mskbday, "mm/dd/yyyy"), Date)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If TxtControlNo.Enabled = True Then
        If MsgBox("Are you sure you want to close the form without saving?", vbYesNo + vbQuestion) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    
    FeMP = False
  
    Set rsHREmp = Nothing
    Set rsDetails = Nothing
    Set rsRef = Nothing
    Set rsDialect = Nothing
    Set rsEmpRec = Nothing
    Set rsCivil = Nothing
    Set rsSem = Nothing
    Set rsEmpStatus = Nothing
    Set rsEmpRec = Nothing
    Set rsDelete = Nothing
    Set rsemp = Nothing
End Sub

'Private Sub mskbday1_GotFocus()
'mskbday1.BackColor = &HFFFFC0
'Call mskHLigths(mskbday1)
'End Sub

Private Sub mskbday_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

'Private Sub mskbday1_LostFocus()
'mskbday1.BackColor = vbWhite
'If mskbday1.Text <> "__/__/____" Then
'txtage.Text = DateDiff("yyyy", Format(mskbday1, "mm/dd/yyyy"), Date)
'End If
'End Sub

'Private Sub mskc1birth_GotFocus()
'mskc1birth.BackColor = &HFFFFC0
'
'End Sub

Private Sub mskc1birth_LostFocus()
mskc1birth.BackColor = vbWhite
End Sub

Private Sub mskc2birth_GotFocus()
mskc2birth.BackColor = &HFFFFC0
End Sub

Private Sub mskc2birth_LostFocus()
mskc2birth.BackColor = vbWhite
End Sub

Private Sub mskfbirth_GotFocus()
mskfbirth.BackColor = &HFFFFC0
End Sub

Private Sub mskfbirth_LostFocus()
mskfbirth.BackColor = vbWhite
End Sub

Private Sub mskmbirth_GotFocus()
mskmbirth.BackColor = &HFFFFC0
End Sub

Private Sub mskmbirth_LostFocus()
mskmbirth.BackColor = vbWhite
End Sub

Private Sub msks1birth_GotFocus()
msks1birth.BackColor = &HFFFFC0
End Sub

Private Sub msks1birth_LostFocus()
msks1birth.BackColor = vbWhite
End Sub

Private Sub msks2birth_GotFocus()
msks2birth.BackColor = &HFFFFC0
End Sub

Private Sub msks2birth_LostFocus()
msks2birth.BackColor = vbWhite
End Sub

Private Sub msksbirth_GotFocus()
msksbirth.BackColor = &HFFFFC0
End Sub

Private Sub msksbirth_LostFocus()
msksbirth.BackColor = vbWhite
End Sub

Private Sub mskbday_Validate(Cancel As Boolean)
If mskbday <> "__/__/____" And IsDate(mskbday) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskbday.SetFocus
End Sub

Private Sub MskDateHired_GotFocus()
MskDateHired.BackColor = &HFFFFC0
End Sub

Private Sub MskDateHired_LostFocus()
MskDateHired.BackColor = vbWhite
End Sub

Private Sub MskDateHired_Validate(Cancel As Boolean)
If MskDateHired <> "__/__/____" And IsDate(MskDateHired) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: MskDateHired.SetFocus
End Sub

Private Sub mskdateprob_GotFocus()
mskdateprob.BackColor = &HFFFFC0
End Sub

Private Sub mskdateprob_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
Private Sub mskdatehired_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskdateprob_LostFocus()
mskdateprob.BackColor = vbWhite
End Sub

Private Sub mskdateprob_Validate(Cancel As Boolean)
If mskdateprob <> "__/__/____" And IsDate(mskdateprob) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskdateprob.SetFocus

End Sub

Private Sub mskdatereg_GotFocus()
mskdatereg.BackColor = &HFFFFC0
End Sub

Private Sub mskdatereg_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskdatereg_LostFocus()
mskdatereg.BackColor = vbWhite
End Sub

Private Sub mskdatereg_Validate(Cancel As Boolean)
If mskdatereg <> "__/__/____" And IsDate(mskdatereg) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskdatereg.SetFocus

End Sub

Private Sub mskdateres_GotFocus()
mskdateres.BackColor = &HFFFFC0
End Sub

Private Sub mskdateres_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskdateres_LostFocus()
mskdateres.BackColor = vbWhite
End Sub

Private Sub mskdateres_Validate(Cancel As Boolean)
If mskdateres <> "__/__/____" And IsDate(mskdateres) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskdateres.SetFocus

End Sub

Private Sub mskdep1_GotFocus()
mskdep1.BackColor = &HFFFFC0
End Sub

Private Sub mskdep1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskdep1_LostFocus()
mskdep1.BackColor = vbWhite
End Sub

Private Sub mskdep1_Validate(Cancel As Boolean)
If mskdep1 <> "__/__/____" And IsDate(mskdep1) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskdep1.SetFocus
End Sub

Private Sub mskdep2_GotFocus()
mskdep2.BackColor = &HFFFFC0
End Sub

Private Sub mskdep2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskdep2_LostFocus()
mskdep2.BackColor = vbWhite
End Sub

Private Sub mskdep2_Validate(Cancel As Boolean)
If mskdep2 <> "__/__/____" And IsDate(mskdep2) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskdep2.SetFocus
End Sub

Private Sub mskdep3_GotFocus()
mskdep3.BackColor = &HFFFFC0
End Sub

Private Sub mskdep3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskdep3_LostFocus()
mskdep3.BackColor = vbWhite
End Sub

Private Sub mskdep3_Validate(Cancel As Boolean)
If mskdep3 <> "__/__/____" And IsDate(mskdep3) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskdep3.SetFocus
End Sub

Private Sub mskdep4_GotFocus()
mskdep4.BackColor = &HFFFFC0
End Sub

Private Sub mskdep4_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskdep4_LostFocus()
mskdep4.BackColor = vbWhite
End Sub

Private Sub mskdep4_Validate(Cancel As Boolean)
If mskdep4 <> "__/__/____" And IsDate(mskdep4) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskdep4.SetFocus

End Sub

Private Sub mskdep5_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskdep5_Validate(Cancel As Boolean)
If mskdep5 <> "__/__/____" And IsDate(mskdep5) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskdep5.SetFocus
End Sub

Private Sub mskStartCash_GotFocus()
mskStartCash.BackColor = &HFFFFC0
Call mskHLigths(mskStartCash)
End Sub

Private Sub mskStartCash_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskStartCash_LostFocus()
mskStartCash.BackColor = vbWhite
End Sub

Private Sub mskStartCash_Validate(Cancel As Boolean)
If mskStartCash <> "__/__/____" And IsDate(mskStartCash) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskStartCash.SetFocus
End Sub

Private Sub mskStartuniform_GotFocus()
mskStartUniform.BackColor = &HFFFFC0
Call mskHLigths(mskStartUniform)
End Sub

Private Sub mskStartuniform_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskStartUniform_LostFocus()
mskStartUniform.BackColor = vbWhite
End Sub

Private Sub mskStartUniform_Validate(Cancel As Boolean)
If mskStartUniform <> "__/__/____" And IsDate(mskStartUniform) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskStartUniform.SetFocus
End Sub

Private Sub mskEndCB_GotFocus()
mskEndCB.BackColor = &HFFFFC0
Call mskHLigths(mskEndCB)
End Sub

Private Sub mskEndCB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskEndCB_LostFocus()
mskEndCB.BackColor = vbWhite
End Sub

Private Sub mskEndCB_Validate(Cancel As Boolean)
If mskEndCB <> "__/__/____" And IsDate(mskEndCB) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskEndCB.SetFocus
End Sub

Private Sub mskEndContract_GotFocus()
mskEndContract.BackColor = &HFFFFC0
Call mskHLigths(mskEndContract)
End Sub

Private Sub mskEndContract_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskEndContract_LostFocus()
mskEndContract.BackColor = vbWhite
End Sub

Private Sub mskEndContract_Validate(Cancel As Boolean)
If mskEndContract <> "__/__/____" And IsDate(mskEndContract) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskEndContract.SetFocus
End Sub

Private Sub mskendof_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskissue_GotFocus()
mskissue.BackColor = &HFFFFC0
Call mskHLigths(mskissue)
End Sub

Private Sub mskissue_LostFocus()
mskissue.BackColor = vbWhite
End Sub

Private Sub mskissue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SSTab1.Tab = 1
    cmbEmpStatus.SetFocus
End If
End Sub

Private Sub mskissue_Validate(Cancel As Boolean)
If mskissue <> "__/__/____" And IsDate(mskissue) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskissue.SetFocus
End Sub

Private Sub MskMFEndPayment_GotFocus()
MskMFEndPayment.BackColor = &HFFFFC0
Call mskHLigths(MskMFEndPayment)
End Sub

Private Sub MskMFEndPayment_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub MskMFEndPayment_LostFocus()
MskMFEndPayment.BackColor = vbWhite
End Sub

Private Sub MskMFEndPayment_Validate(Cancel As Boolean)
If MskMFEndPayment <> "__/__/____" And IsDate(MskMFEndPayment) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: MskMFEndPayment.SetFocus
End Sub

Private Sub MskMFStartPayment_GotFocus()
MskMFStartPayment.BackColor = &HFFFFC0
Call mskHLigths(MskMFStartPayment)
End Sub

Private Sub MskMFStartPayment_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub MskMFStartPayment_LostFocus()
MskMFStartPayment.BackColor = vbWhite
End Sub

Private Sub MskMFStartPayment_Validate(Cancel As Boolean)
If MskMFStartPayment <> "__/__/____" And IsDate(MskMFStartPayment) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: MskMFStartPayment.SetFocus
End Sub

Private Sub MskPFEndPayment_GotFocus()
MskPFEndPayment.BackColor = &HFFFFC0
Call mskHLigths(MskPFEndPayment)

End Sub

Private Sub MskPFEndPayment_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub MskPFEndPayment_LostFocus()
MskPFEndPayment.BackColor = vbWhite
End Sub

Private Sub MskPFEndPayment_Validate(Cancel As Boolean)
If MskPFEndPayment <> "__/__/____" And IsDate(MskPFEndPayment) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: MskPFEndPayment.SetFocus
End Sub

Private Sub MskPFStartPayment_GotFocus()
MskPFStartPayment.BackColor = &HFFFFC0
Call mskHLigths(MskPFStartPayment)

End Sub

Private Sub MskPFStartPayment_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub MskPFStartPayment_LostFocus()
MskPFStartPayment.BackColor = vbWhite
End Sub

Private Sub MskPFStartPayment_Validate(Cancel As Boolean)
If MskPFStartPayment <> "__/__/____" And IsDate(MskPFStartPayment) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: MskPFStartPayment.SetFocus
End Sub

Private Sub msksss_GotFocus()
msksss.BackColor = &HFFFFC0
Call mskHLigths(msksss)
End Sub

Private Sub msksss_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub msksss_LostFocus()
msksss.BackColor = vbWhite
End Sub

Private Sub mskStartCB_GotFocus()
mskStartCB.BackColor = &HFFFFC0
Call mskHLigths(mskStartCB)
End Sub

Private Sub mskStartCB_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskStartCB_LostFocus()
mskStartCB.BackColor = vbWhite
End Sub

Private Sub mskStartCB_Validate(Cancel As Boolean)
If mskStartCB <> "__/__/____" And IsDate(mskStartCB) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskStartCB.SetFocus
End Sub

Private Sub mskStartContract_GotFocus()
mskStartContract.BackColor = &HFFFFC0
Call mskHLigths(mskStartContract)
End Sub

Private Sub mskStartContract_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskStartContract_LostFocus()
mskStartContract.BackColor = vbWhite
End Sub

Private Sub mskStartContract_Validate(Cancel As Boolean)
If mskStartContract <> "__/__/____" And IsDate(mskStartContract) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskStartContract.SetFocus

End Sub

Private Sub msktax_GotFocus()
msktax.BackColor = &HFFFFC0
Call mskHLigths(msktax)
End Sub

Private Sub msktax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub msktax_LostFocus()
msktax.BackColor = vbWhite
End Sub

Private Sub optfemale_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then txtheight.SetFocus
End Sub

Private Sub optmale_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then optfemale.SetFocus

End Sub


Private Sub SetNavigationButtons(ByVal VALUE As Boolean)
    cmdNew.Enabled = False
    cmdedit.Enabled = False
    cmdupdate.Enabled = True
    cmdundo.Enabled = True
    'cmdsearch.Enabled = False
    'cmdhired.Enabled = False
    cmddelete.Enabled = False
    cmdfirst.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdlast.Enabled = False
    cmdRecontract.Enabled = False
End Sub

Private Sub Add()
cmdNew.Enabled = False
cmdedit.Enabled = False
cmdupdate.Enabled = True
cmdundo.Enabled = True
cmddelete.Enabled = False
cmdfirst.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdlast.Enabled = False
cmdRecontract.Enabled = False
End Sub

Private Function IsExist() As Boolean

If editflagship = False Then
Set rsIsExist = New ADODB.Recordset
StrQuery = "SELECT * FROM HREmployees WHERE firstname = '" & RTrim(Me.txtfirstname.Text) & "' and surname='" & RTrim(Me.txtsurname.Text) & "' and Middlename='" & RTrim(Me.txtmiddlename.Text) & "' and Agencycode = " & CompanyCode
Set rsIsExist = ReturnRst(StrQuery, cnnstr, mstatic)
'rsIsExist.Open "Select * from HREmployees where firstname='" & RTrim(Me.txtfirstname.Text) & "' and surname='" & RTrim(Me.txtsurname.Text) & "'  and agencycode=" & CompanyCode, cnnHR, adOpenKeyset, adLockOptimistic

With rsIsExist
    If .EOF = False Then
        If MsgBox("Employee name [ " & txtsurname & "," & txtfirstname.Text & "] already in the database. Do you want to edit this employee?", vbYesNo + vbInformation, PayrollAlert) = vbYes Then
        
        Dim rsHREmptemp As New ADODB.Recordset
        
        StrQuery = "SELECT * FROM HREmployees WHERE firstname='" & RTrim(Me.txtfirstname.Text) & "' and surname='" & RTrim(Me.txtsurname.Text) & "'  and agencycode=" & CompanyCode
        Set rsHREmptemp = ReturnRst(StrQuery, cnnstr, mstatic)
        'rsHREmptemp.Open "select * from HREmployees where firstname='" & RTrim(Me.txtfirstname.Text) & "' and surname='" & RTrim(Me.txtsurname.Text) & "'  and agencycode=" & CompanyCode, cnnHR, adOpenKeyset, adLockOptimistic
             If rsHREmptemp.EOF = False Then
                editflagship = True
                addflagship = False
                Showdata rsHREmptemp
             End If
        End If
       IsExist = True
'       rsHREmptemp.Close
       Set rsHREmptemp = Nothing
    Exit Function
    End If
End With
End If
IsExist = False
End Function

'Private Sub ChkEcola_Click()
'    If cmdUpdate.Enabled = True Then
'        If ChkEcola.VALUE = 1 Then
'            TxtEcolaAmount.Enabled = True
'            'TxtEcolaAmount.SetFocus ' 2010
'        Else
'            TxtEcolaAmount.Enabled = False
'            TxtEcolaAmount.Text = "0.00"
'        End If
'    End If
'End Sub

Private Sub cmbempstatus_Click()
    If cmdedit.Enabled = True Then
        Status = UCase(mID(cmbEmpStatus.Text, 1, 1))
    If Status = "T" Or Status = "W" Or Status = "X" Or Status = "Y" Or Status = "Z" Then
        cmdCopy.Enabled = True
        Me.cmdRecontract.Enabled = True
    Else
        cmdCopy.Enabled = False
        'Me.cmdRecontract.Enabled = False
    End If
    End If
End Sub

Private Sub cmdCopy_Click()
    If cmdupdate.Enabled = True Then
         MsgBox "Please update or save first your task before you transfer.", vbInformation + vbInformation
    Exit Sub
    End If
    frmHRSelectDestination.Show 1
End Sub

Private Sub cmbPayBasis_GotFocus()
cmbPayBasis.BackColor = &HFFFFC0
End Sub

Private Sub cmbPayBasis_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbPayBasis_LostFocus()
    cmbPayBasis.BackColor = vbWhite
End Sub

Private Sub txtacctno_GotFocus()
txtacctno.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtacctno
End Sub

Private Sub txtacctno_LostFocus()
txtacctno.BackColor = vbWhite
End Sub

Private Sub txtActype_GotFocus()
txtActype.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtActype
End Sub

Private Sub txtActype_LostFocus()
txtActype.BackColor = vbWhite
End Sub

Private Sub txtaddress1_GotFocus()
txtaddress1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtaddress1
End Sub
Private Sub txtaddress1_LostFocus()
    txtaddress1.BackColor = vbWhite
End Sub

Private Sub txtCBpayment_GotFocus()
txtCBpayment.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txtCBpayment
Call txtHLigths(txtCBpayment)
End Sub

Private Sub txtCBpayment_LostFocus()
txtCBpayment.BackColor = vbWhite
Call ToZeroText(txtCBpayment)
End Sub

Private Sub txtCBtotal_GotFocus()
txtCBtotal.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txtCBtotal
Call txtHLigths(txtCBtotal)
End Sub

Private Sub txtCBtotal_LostFocus()
txtCBtotal.BackColor = vbWhite
Call ToZeroText(txtCBtotal)
End Sub

Private Sub txtControlNo_GotFocus()
Set oCommonFunction.ValidAlpha = TxtControlNo
Call txtHLigths(TxtControlNo)
End Sub

Private Sub TxtEcolaAmount_GotFocus()
    TxtEcolaAmount.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = TxtEcolaAmount
    Call txtHLigths(TxtEcolaAmount)
End Sub

Private Sub TxtEcolaAmount_LostFocus()
TxtEcolaAmount.BackColor = vbWhite
Call ToZeroText(TxtEcolaAmount)
End Sub

Private Sub txtexam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtheight_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtbirthplace.SetFocus
End Sub

Private Sub txtHide_Click()
Me.Height = 2400
End Sub

Private Sub txthouseno_GotFocus()
txthouseno.BackColor = &HFFFFC0
End Sub

Private Sub txthouseno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtmedical_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SSTab1.Tab = 3
    txtdatepost.SetFocus
End If
End Sub

Private Sub txtMFamount_GotFocus()
TxtMFAmount.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = TxtMFAmount
Call txtHLigths(TxtMFAmount)
End Sub

Private Sub TxtMFAmount_LostFocus()
TxtMFAmount.BackColor = vbWhite
Call ToZeroText(TxtMFAmount)
End Sub

Private Sub TxtMFPayment_GotFocus()
TxtMFPayment.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = TxtMFPayment
Call txtHLigths(TxtMFPayment)
End Sub

Private Sub TxtMFPayment_LostFocus()
TxtMFPayment.BackColor = vbWhite
Call ToZeroText(TxtMFPayment)
End Sub

Private Sub txtMonthsHired_GotFocus()
txtMonthsHired.BackColor = &HFFFFC0
Set oCommonFunction.ValidInteger = txtMonthsHired
Call txtHLigths(txtMonthsHired)
End Sub

Private Sub txtMonthsHired_LostFocus()
txtMonthsHired.BackColor = vbWhite
End Sub

Private Sub txtMonthsHired_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtBillingRate_GotFocus()
txtBillingRate.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txtBillingRate
Call txtHLigths(txtBillingRate)
End Sub

Private Sub txtBillingRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
'    SSTab1.Tab = 2
'    '    txtfathername.SetFocus
'    'mskStartUniform.SetFocus
'End If
End Sub

Private Sub txtBillingRate_LostFocus()
txtBillingRate.BackColor = vbWhite
Call ToZeroText(txtBillingRate)
End Sub

Private Sub txtnickname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtnumdep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtPFAmount_GotFocus()
txtPFAmount.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txtPFAmount
Call txtHLigths(txtPFAmount)
End Sub

Private Sub ToZeroText(ByVal txtbox As TextBox)
If Val(txtbox.Text) > 0 Then
    txtbox.Text = Format(txtbox.Text, "##0.000")
Else
    txtbox.Text = "0.000"
End If
End Sub

Private Sub txtPFAmount_LostFocus()
txtPFAmount.BackColor = vbWhite
Call ToZeroText(txtPFAmount)
End Sub

Private Sub txtPFPayment_GotFocus()
txtPFPayment.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txtPFPayment
Call txtHLigths(txtPFPayment)
End Sub

Private Sub txtPFPayment_LostFocus()
txtPFPayment.BackColor = vbWhite
Call ToZeroText(txtPFPayment)
End Sub

Private Sub txtSuretyPercent_GotFocus()
txtSuretyPercent.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txtSuretyPercent
Call txtHLigths(txtSuretyPercent)
End Sub

Private Sub txtSuretyPercent_LostFocus()
txtSuretyPercent.BackColor = vbWhite
Call ToZeroText(txtSuretyPercent)
End Sub

Private Sub txtRate_GotFocus()
txtRate.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txtRate
Call txtHLigths(txtRate)
End Sub

Private Sub txtRate_LostFocus()

On Error GoTo chkClient

txtRate.BackColor = vbWhite
Call ToZeroText(txtRate)

Cid = cmbCompany.ItemData(cmbCompany.ListIndex)  ' previously remarked

'If cmbPayBasis.ListIndex = 0 Then 'hourly
'        MonthlyRate = (8 * txtRate.Text) * getField("NoofDaysperMonth", Cid)
'ElseIf cmbPayBasis.ListIndex = 1 Then 'daily
'        MonthlyRate = txtRate.Text * getField("NoofDaysperMonth", Cid)
'ElseIf cmbPayBasis.ListIndex = 2 Then 'monthly
'        MonthlyRate = Format(Val(txtRate.Text), "##0.00")
'ElseIf cmbPayBasis.ListIndex = 3 Then 'Commission Based
'        MonthlyRate = Format(Val(txtRate.Text), "##0.00")
'End If
'txtMonthlyRate = Format(MonthlyRate, "##0.000")

If cmbPayBasis.ListIndex = 0 Then 'Hourly
        MonthlyRate = ((8 * txtRate.Text) * getField("WorkingDaysPerYear", Cid)) / 12
ElseIf cmbPayBasis.ListIndex = 1 Then 'Daily
        MonthlyRate = (txtRate.Text * getField("WorkingDaysPerYear", Cid)) / 12
ElseIf cmbPayBasis.ListIndex = 2 Then 'Monthly
        MonthlyRate = Val(txtRate.Text)
ElseIf cmbPayBasis.ListIndex = 3 Then 'Commission Based
        MonthlyRate = Val(txtRate.Text)
End If
txtMonthlyRate = MonthlyRate


chkClient:

    If Cid = 0 Or Cid = -1 Then
        MsgBox "Please select Client for this employee first to compute for Monthly Salary", vbInformation, PayrollAlert
    End If
    
End Sub

Private Sub txtMonthlyRate_GotFocus()
txtMonthlyRate.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txtMonthlyRate
Call txtHLigths(txtMonthlyRate)
End Sub

Private Sub txtMonthlyRate_LostFocus()
txtMonthlyRate.BackColor = vbWhite
Call ToZeroText(txtMonthlyRate)
End Sub

Private Sub txtreasonresign_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtrefrel3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SSTab1.Tab = 0
    txthouseno.SetFocus
End If
End Sub

Private Sub txtstreet_GotFocus()
txtstreet.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtstreet
End Sub
Private Sub txtcity_GotFocus()
txtcity.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtcity
End Sub
Private Sub txtcity_LostFocus()
txtcity.BackColor = vbWhite
End Sub

Private Sub txtstreet_LostFocus()
txtstreet.BackColor = vbWhite
End Sub
Private Sub txtbarangay_GotFocus()
txtbarangay.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtbarangay
End Sub
Private Sub txtdistrict_GotFocus()
txtdistrict.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdistrict
End Sub
Private Sub txtprovince_GotFocus()
txtprovince.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtprovince
End Sub
Private Sub txtprovince_LostFocus()
txtprovince.BackColor = vbWhite
End Sub

Private Sub txtdistrict_LostFocus()
txtdistrict.BackColor = vbWhite
End Sub

Private Sub txtbarangay_LostFocus()
txtbarangay.BackColor = vbWhite
End Sub


Private Sub txthouseno_LostFocus()
txthouseno.BackColor = vbWhite
End Sub

Private Sub txtage_GotFocus()
txtage.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtage
End Sub


Private Sub txtage_LostFocus()
txtage.BackColor = vbWhite
End Sub

Private Sub txtappstatus_LostFocus()
txtappstatus.BackColor = vbWhite
End Sub

Private Sub txtbirthplace_GotFocus()
txtbirthplace.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtbirthplace
End Sub


Private Sub txtbirthplace_LostFocus()
txtbirthplace.BackColor = vbWhite
End Sub

Private Sub txtbro1name_GotFocus()
txtbro1name.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtbro1name
End Sub

Private Sub txtbro1name_LostFocus()
txtbro1name.BackColor = vbWhite
End Sub

Private Sub txtbro2name_GotFocus()
txtbro2name.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtbro2name
End Sub

Private Sub txtbro2name_LostFocus()
txtbro2name.BackColor = vbWhite
End Sub

Private Sub txtchild1name_GotFocus()
txtchild1name.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtchild1name
End Sub

Private Sub txtchild1name_LostFocus()
txtchild1name.BackColor = vbWhite
End Sub

Private Sub txtchild2name_GotFocus()
txtchild2name.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtchild2name
End Sub

Private Sub txtchild2name_LostFocus()
txtchild2name.BackColor = vbWhite
End Sub

Private Sub txtcitizen_GotFocus()
txtcitizen.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtcitizen
End Sub

Private Sub txtcitizen_LostFocus()
txtcitizen.BackColor = vbWhite
End Sub

Private Sub txtcityaddress_GotFocus()
txtcityaddress.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtcityaddress
End Sub

Private Sub txtcityaddress_LostFocus()
txtcityaddress.BackColor = vbWhite
End Sub

Private Sub txtcollegecourse_GotFocus()
txtcollegecourse.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtcollegecourse
End Sub


Private Sub txtcollegecourse_LostFocus()
txtcollegecourse.BackColor = vbWhite
End Sub

Private Sub txtcollegemajor_GotFocus()
txtcollegemajor.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtcollegemajor
End Sub



Private Sub txtcollegemajor_LostFocus()
txtcollegemajor.BackColor = vbWhite
End Sub

Private Sub txtcollegesch_GotFocus()
txtcollegesch.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtcollegesch
End Sub



Private Sub txtcollegesch_LostFocus()
txtcollegesch.BackColor = vbWhite
End Sub

Private Sub txtdateapplied_GotFocus()
txtdateapplied.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdateapplied
End Sub



Private Sub txtdateapplied_LostFocus()
txtdateapplied.BackColor = vbWhite
End Sub

Private Sub txtdatecollege_GotFocus()
txtdatecollege.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdatecollege
End Sub


Private Sub txtdatecollege_LostFocus()
txtdatecollege.BackColor = vbWhite
End Sub

Private Sub txtdatehired_GotFocus()
txtdatehired.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdatehired
End Sub


Private Sub txtdatehired_LostFocus()
txtdatehired.BackColor = vbWhite
End Sub

Private Sub txtdatepost_GotFocus()
txtdatepost.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdatepost
End Sub


Private Sub txtdatepost_LostFocus()
txtdatepost.BackColor = vbWhite
End Sub

Private Sub txtdateprimary_GotFocus()
txtdateprimary.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdateprimary
End Sub


Private Sub txtdateprimary_LostFocus()
txtdateprimary.BackColor = vbWhite
End Sub

Private Sub txtdatesec_GotFocus()
txtdatesec.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdatesec
End Sub


Private Sub txtdatesec_LostFocus()
txtdatesec.BackColor = vbWhite
End Sub

Private Sub txtdateseminar1_GotFocus()
txtdateseminar1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdateseminar1
End Sub


Private Sub txtdateseminar1_LostFocus()
txtdateseminar1.BackColor = vbWhite
End Sub

Private Sub txtdateseminar2_GotFocus()
txtdateseminar2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdateseminar2
End Sub


Private Sub txtdateseminar2_LostFocus()
txtdateseminar2.BackColor = vbWhite
End Sub

Private Sub txtdateseminar3_GotFocus()
txtdateseminar3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdateseminar3
End Sub


Private Sub txtdateseminar3_LostFocus()
txtdateseminar3.BackColor = vbWhite
End Sub

Private Sub txtdep1_GotFocus()
txtdep1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdep1
End Sub

Private Sub txtdep1_LostFocus()
txtdep1.BackColor = vbWhite
End Sub

Private Sub txtdep2_GotFocus()
txtdep2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdep2
End Sub


Private Sub txtdep2_LostFocus()
txtdep2.BackColor = vbWhite
End Sub

Private Sub txtdep3_GotFocus()
txtdep3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdep3
End Sub


Private Sub txtdep3_LostFocus()
txtdep3.BackColor = vbWhite
End Sub
Private Sub txtdep5_LostFocus()
txtdep5.BackColor = vbWhite
End Sub

Private Sub txtdep4_GotFocus()
txtdep4.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdep4
End Sub
Private Sub txtdep5_GotFocus()
txtdep5.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdep5
End Sub

Private Sub txtdep4_LostFocus()
txtdep4.BackColor = vbWhite
End Sub

Private Sub txtdialect1_GotFocus()
txtdialect1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdialect1
End Sub

Private Sub txtdialect1_LostFocus()
txtdialect1.BackColor = vbWhite
End Sub

Private Sub txtdialect2_GotFocus()
txtdialect2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdialect2
End Sub


Private Sub txtdialect2_LostFocus()
txtdialect2.BackColor = vbWhite
End Sub

Private Sub txtdialect3_GotFocus()
txtdialect3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtdialect3
End Sub


Private Sub txtdialect3_LostFocus()
txtdialect3.BackColor = vbWhite
End Sub




Private Sub txtemergency_GotFocus()
txtemergency.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtemergency
End Sub

Private Sub txtemergency_LostFocus()
txtemergency.BackColor = vbWhite
End Sub

Private Sub txteraddress_GotFocus()
txteraddress.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txteraddress
End Sub

Private Sub txteraddress_LostFocus()
txteraddress.BackColor = vbWhite
End Sub

Private Sub txtermobile_GotFocus()
txtermobile.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtermobile
End Sub

Private Sub txtermobile_LostFocus()
txtermobile.BackColor = vbWhite
End Sub

Private Sub txtertel_GotFocus()
txtertel.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtertel
End Sub

Private Sub txtertel_LostFocus()
txtertel.BackColor = vbWhite
End Sub

Private Sub txtexam_GotFocus()
txtexam.BackColor = &HFFFFC0
Set oCommonFunction.ValidInteger = txtexam
End Sub

Private Sub txtexam_LostFocus()
txtexam.BackColor = vbWhite
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txteyecolor_GotFocus()
txteyecolor.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txteyecolor
End Sub

Private Sub txteyecolor_LostFocus()
txteyecolor.BackColor = vbWhite
End Sub

Private Sub txtfaddress_GotFocus()
txtfaddress.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtfaddress
End Sub

Private Sub txtfaddress_LostFocus()
txtfaddress.BackColor = vbWhite
End Sub

Private Sub txtfathername_GotFocus()
txtfathername.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtfathername
End Sub

Private Sub txtfathername_LostFocus()
txtfathername.BackColor = vbWhite
End Sub

Private Sub txtfirstchoice_GotFocus()
txtfirstchoice.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtfirstchoice
End Sub

Private Sub txtfirstchoice_LostFocus()
txtfirstchoice.BackColor = vbWhite
End Sub

Private Sub txtfirstname_GotFocus()
txtfirstname.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtfirstname
Call txtHLigths(txtfirstname)
End Sub

Private Sub txtfirstname_LostFocus()
If IsExist = True Then txtfirstname.SetFocus: Exit Sub
txtfirstname.BackColor = vbWhite
End Sub

Private Sub txtfjob_GotFocus()
txtfjob.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtfjob
End Sub

Private Sub txtfjob_LostFocus()
txtfjob.BackColor = vbWhite
End Sub

Private Sub txtfrom1_GotFocus()
txtfrom1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtfrom1
End Sub

Private Sub txtfrom1_LostFocus()
txtfrom1.BackColor = vbWhite
End Sub

Private Sub txtfrom2_GotFocus()
txtfrom2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtfrom2
End Sub

Private Sub txtfrom2_LostFocus()
txtfrom2.BackColor = vbWhite
End Sub

Private Sub txtfrom3_GotFocus()
txtfrom3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtfrom3
End Sub

Private Sub txtfrom3_LostFocus()
txtfrom3.BackColor = vbWhite
End Sub

Private Sub txthaircolor_GotFocus()
txthaircolor.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txthaircolor
End Sub

Private Sub txthaircolor_LostFocus()
txthaircolor.BackColor = vbWhite
End Sub

Private Sub txtheight_GotFocus()
txtheight.BackColor = &HFFFFC0
'Set oCommonFunction.ValidAlpha = txtheight
End Sub

Private Sub txtheight_LostFocus()
txtheight.BackColor = vbWhite
End Sub

Private Sub txthprelate_LostFocus()
txthprelate.BackColor = vbWhite
End Sub

Private Sub txthrswork_GotFocus()
txthrswork.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txthrswork
End Sub

Private Sub txthrswork_LostFocus()
txthrswork.BackColor = vbWhite
End Sub

Private Sub txtinterview_GotFocus()
txtinterview.BackColor = &HFFFFC0
Set oCommonFunction.ValidInteger = txtinterview
End Sub

Private Sub txtinterview_LostFocus()
txtinterview.BackColor = vbWhite
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtissueat_GotFocus()
txtissueat.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtissueat
Call txtHLigths(txtissueat)
End Sub

Private Sub txtissueat_LostFocus()
txtissueat.BackColor = vbWhite
End Sub

Private Sub txtmaddress_GotFocus()
txtmaddress.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtmaddress
End Sub

Private Sub txtmaddress_LostFocus()
txtmaddress.BackColor = vbWhite
End Sub

Private Sub txtmedical_GotFocus()
txtmedical.BackColor = &HFFFFC0
'Set oCommonFunction.ValidAlpha = txtmedical
End Sub

Private Sub txtmedical_LostFocus()
txtmedical.BackColor = vbWhite
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtmiddlename_GotFocus()
txtmiddlename.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtmiddlename
Call txtHLigths(txtmiddlename)
End Sub

Private Sub txtmiddlename_LostFocus()
txtmiddlename.BackColor = vbWhite
End Sub

Private Sub txtmjob_GotFocus()
txtmjob.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtmjob
End Sub

Private Sub txtmjob_LostFocus()
txtmjob.BackColor = vbWhite
End Sub

Private Sub txtmothername_GotFocus()
txtmothername.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtmothername
End Sub

Private Sub txtmothername_LostFocus()
txtmothername.BackColor = vbWhite
End Sub

Private Sub txtnameofcompany1_GotFocus()
txtnameofcompany1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtnameofcompany1
End Sub

Private Sub txtnameofcompany1_LostFocus()
txtnameofcompany1.BackColor = vbWhite
End Sub

Private Sub txtnameofcompany2_GotFocus()
txtnameofcompany2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtnameofcompany2
End Sub

Private Sub txtnameofcompany2_LostFocus()
txtnameofcompany2.BackColor = vbWhite
End Sub

Private Sub txtnameofcompany3_GotFocus()
txtnameofcompany3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtnameofcompany3
End Sub

Private Sub txtnameofcompany3_LostFocus()
txtnameofcompany3.BackColor = vbWhite
End Sub

Private Sub txtnbi_GotFocus()
txtnbi.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtnbi
End Sub

Private Sub txtnbi_LostFocus()
txtnbi.BackColor = vbWhite
End Sub

Private Sub txtnickname_GotFocus()
txtnickname.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtnickname
Call txtHLigths(txtnickname)
End Sub

Private Sub txtnickname_LostFocus()
txtnickname.BackColor = vbWhite
'If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtnolate_GotFocus()
txtnolate.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtnolate
End Sub

Private Sub txtnolate_LostFocus()
txtnolate.BackColor = vbWhite
End Sub

Private Sub txtnoot_GotFocus()
txtnoot.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtnoot
End Sub

Private Sub txtnoot_LostFocus()
txtnoot.BackColor = vbWhite
End Sub

Private Sub txtnoundertime_GotFocus()
txtnoundertime.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtnoundertime
End Sub

Private Sub txtnoundertime_LostFocus()
txtnoundertime.BackColor = vbWhite
End Sub

Private Sub txtnumdep_GotFocus()
txtnumdep.BackColor = &HFFFFC0
'Set oCommonFunction.ValidAlpha = txtnumdep
Call txtHLigths(txtnumdep)
End Sub

Private Sub txtnumdep_LostFocus()
txtnumdep.BackColor = vbWhite
If Trim(txtnumdep) = "" Then txtnumdep = "0"
End Sub

Private Sub txtpagibig_GotFocus()
txtpagibig.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtpagibig
End Sub

Private Sub txtpagibig_LostFocus()
txtpagibig.BackColor = vbWhite
End Sub

Private Sub txtphilhealth_GotFocus()
txtphilhealth.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtphilhealth
End Sub

Private Sub txtphilhealth_LostFocus()
txtphilhealth.BackColor = vbWhite
End Sub

Private Sub txtpos1_GotFocus()
txtpos1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtpos1
End Sub

Private Sub txtpos1_LostFocus()
txtpos1.BackColor = vbWhite
End Sub

Private Sub txtpos2_GotFocus()
txtpos2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtpos2
End Sub

Private Sub txtpos2_LostFocus()
txtpos2.BackColor = vbWhite
End Sub

Private Sub txtpos3_GotFocus()
txtpos3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtpos3
End Sub

Private Sub txtpos3_LostFocus()
txtpos3.BackColor = vbWhite
End Sub

Private Sub txtpostcourse_GotFocus()
txtpostcourse.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtpostcourse
End Sub

Private Sub txtpostcourse_LostFocus()
txtpostcourse.BackColor = vbWhite
End Sub

Private Sub txtpostsch_GotFocus()
txtpostsch.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtpostsch
End Sub


Private Sub txtpostsch_LostFocus()
txtpostsch.BackColor = vbWhite
End Sub

Private Sub txtprimarysch_GotFocus()
txtprimarysch.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtprimarysch
End Sub

Private Sub txtprimarysch_LostFocus()
txtprimarysch.BackColor = vbWhite
End Sub

Private Sub txtprovinceaddress_GotFocus()
txtprovinceaddress.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtprovinceaddress
End Sub

Private Sub txtprovinceaddress_LostFocus()
txtprovinceaddress.BackColor = vbWhite
End Sub

Private Sub txtreason1_GotFocus()
txtreason1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtreason1
End Sub

Private Sub txtreason1_LostFocus()
txtreason1.BackColor = vbWhite
End Sub

Private Sub txtreason2_GotFocus()
txtreason2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtreason2
End Sub

Private Sub txtreason2_LostFocus()
txtreason2.BackColor = vbWhite
End Sub

Private Sub txtreason3_GotFocus()
txtreason3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtreason3
End Sub

Private Sub txtreason3_LostFocus()
txtreason3.BackColor = vbWhite
End Sub

Private Sub txtreasonresign_GotFocus()
txtreasonresign.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtreasonresign
End Sub

Private Sub txtreasonresign_LostFocus()
txtreasonresign.BackColor = vbWhite
End Sub

Private Sub txtrefaddress1_GotFocus()
txtrefaddress1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefaddress1
End Sub

Private Sub txtrefaddress1_LostFocus()
txtrefaddress1.BackColor = vbWhite
End Sub

Private Sub txtrefaddress2_GotFocus()
txtrefaddress2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefaddress2
End Sub


Private Sub txtrefaddress2_LostFocus()
txtrefaddress2.BackColor = vbWhite
End Sub

Private Sub txtrefaddress3_GotFocus()
txtrefaddress3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefaddress3
End Sub

Private Sub txtrefaddress3_LostFocus()
txtrefaddress3.BackColor = vbWhite
End Sub

Private Sub txtrefjob1_GotFocus()
txtrefjob1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefjob1
End Sub


Private Sub txtrefjob1_LostFocus()
txtrefjob1.BackColor = vbWhite
End Sub

Private Sub txtrefjob2_GotFocus()
txtrefjob2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefjob2
End Sub


Private Sub txtrefjob2_LostFocus()
txtrefjob2.BackColor = vbWhite
End Sub

Private Sub txtrefjob3_GotFocus()
txtrefjob3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefjob3
End Sub


Private Sub txtrefjob3_LostFocus()
txtrefjob3.BackColor = vbWhite
End Sub

Private Sub txtrefname1_GotFocus()
txtrefname1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefname1
End Sub

Private Sub txtrefname1_LostFocus()
txtrefname1.BackColor = vbWhite
End Sub

Private Sub txtrefname2_GotFocus()
txtrefname2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefname2
End Sub

Private Sub txtrefname2_LostFocus()
txtrefname2.BackColor = vbWhite
End Sub

Private Sub txtrefname3_GotFocus()
txtrefname3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefname3
End Sub


Private Sub txtrefname3_LostFocus()
txtrefname3.BackColor = vbWhite
End Sub

Private Sub txtrefphone1_GotFocus()
txtrefphone1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefphone1
End Sub

Private Sub txtrefphone1_LostFocus()
txtrefphone1.BackColor = vbWhite
End Sub

Private Sub txtrefphone2_GotFocus()
txtrefphone2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefphone2
End Sub

Private Sub txtrefphone2_LostFocus()
txtrefphone2.BackColor = vbWhite
End Sub

Private Sub txtrefphone3_GotFocus()
txtrefphone3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefphone3
End Sub

Private Sub txtrefphone3_LostFocus()
txtrefphone3.BackColor = vbWhite
End Sub

Private Sub txtrefrel1_GotFocus()
txtrefrel1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefrel1
End Sub


Private Sub txtrefrel1_LostFocus()
txtrefrel1.BackColor = vbWhite
End Sub

Private Sub txtrefrel2_GotFocus()
txtrefrel2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefrel2
End Sub

Private Sub txtrefrel2_LostFocus()
txtrefrel2.BackColor = vbWhite
End Sub

Private Sub txtrefrel3_GotFocus()
txtrefrel3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrefrel3
End Sub

Private Sub txtrefrel3_LostFocus()
txtrefrel3.BackColor = vbWhite
End Sub

Private Sub txtreligion_GotFocus()
txtreligion.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtreligion
End Sub

Private Sub txtreligion_LostFocus()
txtreligion.BackColor = vbWhite
End Sub

Private Sub txtsecondchoice_GotFocus()
txtsecondchoice.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtsecondchoice
End Sub

Private Sub txtsecondchoice_LostFocus()
txtsecondchoice.BackColor = vbWhite
End Sub

Private Sub txtsis1name_GotFocus()
txtsis1name.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtsis1name
End Sub

Private Sub txtsis1name_LostFocus()
txtsis1name.BackColor = vbWhite
End Sub

Private Sub txtsis2name_GotFocus()
txtsis2name.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtsis2name
End Sub

Private Sub txtsis2name_LostFocus()
txtsis2name.BackColor = vbWhite
End Sub

Private Sub txtrescer_GotFocus()
txtrescer.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtrescer
Call txtHLigths(txtrescer)
End Sub

Private Sub txtrescer_LostFocus()
txtrescer.BackColor = vbWhite
End Sub

Private Sub txtsecsch_GotFocus()
txtsecsch.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtsecsch
End Sub

Private Sub txtsecsch_LostFocus()
txtsecsch.BackColor = vbWhite
End Sub

Private Sub txtseminar1_GotFocus()
txtseminar1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtseminar1
End Sub


Private Sub txtseminar1_LostFocus()
txtseminar1.BackColor = vbWhite
End Sub

Private Sub txtseminar2_GotFocus()
txtseminar2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtseminar2
End Sub


Private Sub txtseminar2_LostFocus()
txtseminar2.BackColor = vbWhite
End Sub

Private Sub txtseminar3_GotFocus()
txtseminar3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtseminar3
End Sub

Private Sub txtseminar3_LostFocus()
txtseminar3.BackColor = vbWhite
End Sub

Private Sub txtsjob_GotFocus()
txtsjob.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtsjob
End Sub


Private Sub txtsjob_LostFocus()
txtsjob.BackColor = vbWhite
End Sub

Private Sub txtspouseaddress_GotFocus()
txtspouseaddress.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtspouseaddress
End Sub


Private Sub txtspouseaddress_LostFocus()
txtspouseaddress.BackColor = vbWhite
End Sub

Private Sub txtspousename_GotFocus()
txtspousename.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtspousename
End Sub

Private Sub txtspousename_LostFocus()
txtspousename.BackColor = vbWhite
End Sub

Private Sub TxtContactPerson_GotFocus()
TxtContactPerson.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = TxtContactPerson
End Sub

Private Sub TxtContactPerson_LostFocus()
TxtContactPerson.BackColor = vbWhite
End Sub

Private Sub txtContactAddress_GotFocus()
txtContactAddress.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtContactAddress
End Sub

Private Sub txtContactAddress_LostFocus()
txtContactAddress.BackColor = vbWhite
End Sub

Private Sub txtContactPhone_GotFocus()
txtContactPhone.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtContactPhone
End Sub

Private Sub txtContactPhone_LostFocus()
txtContactPhone.BackColor = vbWhite
End Sub

Private Sub txtsuperior1_GotFocus()
txtsuperior1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtsuperior1
End Sub


Private Sub txtsuperior1_LostFocus()
txtsuperior1.BackColor = vbWhite
End Sub

Private Sub txtsuperior2_GotFocus()
txtsuperior2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtsuperior2
End Sub

Private Sub txtsuperior2_LostFocus()
txtsuperior2.BackColor = vbWhite
End Sub

Private Sub txtsuperior3_GotFocus()
txtsuperior3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtsuperior3
End Sub

Private Sub txtsuperior3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SSTab1.Tab = 4
    txtdialect1.SetFocus
End If
End Sub

Private Sub txtsuperior3_LostFocus()
txtsuperior3.BackColor = vbWhite
End Sub

Private Sub txtsupervisor_GotFocus()
txtsupervisor.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtsupervisor
End Sub

Private Sub txtsupervisor_LostFocus()
txtsupervisor.BackColor = vbWhite
End Sub

Private Sub txtsurname_GotFocus()
txtsurname.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtsurname
Call txtHLigths(txtsurname)
End Sub

Private Sub txtsurname_LostFocus()
txtsurname.BackColor = vbWhite
End Sub

Private Sub txttel1_GotFocus()
txttel1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txttel1
End Sub

Private Sub txttel1_LostFocus()
txttel1.BackColor = vbWhite
End Sub

Private Sub txttel2_GotFocus()
txttel2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txttel2
End Sub


Private Sub txttel2_LostFocus()
txttel2.BackColor = vbWhite
End Sub

Private Sub txttel3_GotFocus()
txttel3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txttel3
End Sub


Private Sub txttel3_LostFocus()
txttel3.BackColor = vbWhite
End Sub

Private Sub txttel4_GotFocus()
txttel4.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txttel4
End Sub

Private Sub txttel4_LostFocus()
txttel4.BackColor = vbWhite
End Sub

Private Sub txtto1_GotFocus()
txtto1.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtto1
End Sub

Private Sub txtto1_LostFocus()
txtto1.BackColor = vbWhite
End Sub

Private Sub txtto2_GotFocus()
txtto2.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtto2
End Sub

Private Sub txtto2_LostFocus()
txtto2.BackColor = vbWhite
End Sub

Private Sub txtto3_GotFocus()
txtto3.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtto3
End Sub


Private Sub txtto3_LostFocus()
txtto3.BackColor = vbWhite
End Sub

Private Sub txtweight_GotFocus()
txtweight.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtweight
End Sub

Private Sub txtweight_LostFocus()
txtweight.BackColor = vbWhite
End Sub

Public Sub CommandPass(ByVal strCommand As String)

Select Case strCommand
    Case "Close"
        Unload Me
End Select

End Sub

