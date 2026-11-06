VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmHRFinalEmployees 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Information"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   Icon            =   "frmHRFinalEmployees.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   10590
   Begin TabDlg.SSTab SSTab1 
      Height          =   4125
      Left            =   135
      TabIndex        =   5
      Top             =   1620
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
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Contract"
      TabPicture(1)   =   "frmHRFinalEmployees.frx":05A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame15"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame12"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame14"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Family"
      TabPicture(2)   =   "frmHRFinalEmployees.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(2)=   "Frame3"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Educational"
      TabPicture(3)   =   "frmHRFinalEmployees.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5(1)"
      Tab(3).Control(1)=   "Frame8"
      Tab(3).Control(2)=   "Frame10"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "References"
      TabPicture(4)   =   "frmHRFinalEmployees.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame11"
      Tab(4).Control(1)=   "Frame9"
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame14 
         Height          =   870
         Left            =   135
         TabIndex        =   273
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
            TabIndex        =   295
            Top             =   495
            Width           =   2640
         End
         Begin VB.ComboBox cmbempstatus 
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
            Left            =   1890
            List            =   "frmHRFinalEmployees.frx":0618
            Style           =   2  'Dropdown List
            TabIndex        =   37
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
            TabIndex        =   296
            Top             =   540
            Width           =   630
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Employee's Status"
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
            TabIndex        =   274
            Top             =   180
            Width           =   1320
         End
      End
      Begin VB.Frame Frame12 
         Height          =   1995
         Left            =   135
         TabIndex        =   241
         Top             =   1125
         Width           =   5055
         Begin VB.CommandButton cmdAddDept 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   285
            Left            =   4635
            TabIndex        =   260
            Top             =   1305
            Width           =   330
         End
         Begin VB.CommandButton cmdAddBranch 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   285
            Left            =   4635
            TabIndex        =   259
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
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1305
            Width           =   2685
         End
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
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   945
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
            TabIndex        =   42
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
            TabIndex        =   39
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
            TabIndex        =   38
            Top             =   225
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
            TabIndex        =   257
            Top             =   1395
            Width           =   825
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Branch"
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
            TabIndex        =   256
            Top             =   1035
            Width           =   525
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Supervisor"
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
            TabIndex        =   244
            Top             =   1710
            Width           =   795
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
            TabIndex        =   243
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
            TabIndex        =   242
            Top             =   315
            Width           =   1125
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1590
         Left            =   -66990
         TabIndex        =   278
         Top             =   2160
         Width           =   2115
         Begin VB.TextBox txtmedical 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1305
            TabIndex        =   88
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
            TabIndex        =   87
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
            TabIndex        =   86
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
            TabIndex        =   281
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
            TabIndex        =   280
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
            TabIndex        =   279
            Top             =   315
            Width           =   930
         End
      End
      Begin VB.Frame Frame15 
         Height          =   690
         Left            =   5310
         TabIndex        =   275
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
            Left            =   1170
            TabIndex        =   47
            Top             =   270
            Width           =   1185
            _ExtentX        =   2090
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
            Left            =   3510
            TabIndex        =   48
            Top             =   270
            Width           =   1185
            _ExtentX        =   2090
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
            TabIndex        =   277
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
            TabIndex        =   276
            Top             =   315
            Width           =   1275
         End
      End
      Begin VB.Frame Frame13 
         Height          =   825
         Left            =   5310
         TabIndex        =   264
         Top             =   990
         Width           =   4875
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
            Left            =   3735
            TabIndex        =   51
            Text            =   "15"
            Top             =   405
            Width           =   645
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
            Left            =   2115
            TabIndex        =   50
            Text            =   "15"
            Top             =   405
            Width           =   1005
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
            ItemData        =   "frmHRFinalEmployees.frx":061A
            Left            =   270
            List            =   "frmHRFinalEmployees.frx":0627
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   405
            Width           =   1485
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
            Left            =   2880
            TabIndex        =   265
            Text            =   "15"
            Top             =   1035
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
            Index           =   13
            Left            =   540
            TabIndex        =   299
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
            TabIndex        =   298
            Top             =   405
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
            TabIndex        =   297
            Top             =   405
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
            Left            =   3465
            TabIndex        =   284
            Top             =   180
            Width           =   915
         End
         Begin VB.Label lblRAte 
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
            Left            =   1935
            TabIndex        =   267
            Top             =   180
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
            Left            =   135
            TabIndex        =   266
            Top             =   180
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2190
         Index           =   0
         Left            =   5310
         TabIndex        =   268
         Top             =   1800
         Width           =   4920
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
            Left            =   1215
            TabIndex        =   58
            Text            =   "0.00"
            Top             =   1260
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
            Left            =   1260
            TabIndex        =   54
            Text            =   "0.00"
            Top             =   720
            Width           =   915
         End
         Begin VB.CheckBox ChkEcola 
            Caption         =   "Ecola payment"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   180
            TabIndex        =   286
            Top             =   180
            Width           =   1410
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
            Left            =   180
            TabIndex        =   57
            Text            =   "0.00"
            Top             =   1260
            Width           =   915
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
            Left            =   1710
            TabIndex        =   52
            Text            =   "0.00"
            Top             =   195
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
            Left            =   135
            TabIndex        =   53
            Text            =   "0.00"
            Top             =   720
            Width           =   915
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
            Left            =   1260
            TabIndex        =   62
            Text            =   "0.00"
            Top             =   1800
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
            TabIndex        =   61
            Text            =   "0.00"
            Top             =   1800
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
            Left            =   2430
            TabIndex        =   63
            Top             =   1800
            Width           =   1155
            _ExtentX        =   2037
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
            Left            =   3690
            TabIndex        =   64
            Top             =   1800
            Width           =   1155
            _ExtentX        =   2037
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
            Left            =   2430
            TabIndex        =   59
            Top             =   1260
            Width           =   1155
            _ExtentX        =   2037
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
            Left            =   3690
            TabIndex        =   60
            Top             =   1260
            Width           =   1155
            _ExtentX        =   2037
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
            Left            =   2430
            TabIndex        =   55
            Top             =   720
            Width           =   1155
            _ExtentX        =   2037
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
            Left            =   3690
            TabIndex        =   56
            Top             =   720
            Width           =   1155
            _ExtentX        =   2037
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
            Caption         =   "PF Pmt."
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   1305
            TabIndex        =   293
            Top             =   495
            Width           =   1035
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Misc. Pmt."
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   1305
            TabIndex        =   292
            Top             =   1035
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
            Index           =   8
            Left            =   2430
            TabIndex        =   291
            Top             =   495
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
            Index           =   7
            Left            =   3735
            TabIndex        =   290
            Top             =   495
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
            Index           =   6
            Left            =   2430
            TabIndex        =   289
            Top             =   1035
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
            Index           =   0
            Left            =   3705
            TabIndex        =   288
            Top             =   1035
            Width           =   735
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Misc. Amt."
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   283
            Top             =   1035
            Width           =   1035
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "PF Amt."
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   180
            TabIndex        =   282
            Top             =   495
            Width           =   1035
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
            Left            =   3705
            TabIndex        =   272
            Top             =   1575
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
            Index           =   33
            Left            =   2475
            TabIndex        =   271
            Top             =   1575
            Width           =   825
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "CB payment"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            Left            =   1290
            TabIndex        =   270
            Top             =   1575
            Width           =   1035
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "CB total"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   269
            Top             =   1575
            Width           =   825
         End
      End
      Begin VB.Frame Frame4 
         Height          =   915
         Left            =   135
         TabIndex        =   231
         Top             =   3060
         Width           =   5070
         Begin MSMask.MaskEdBox mskdateres 
            Height          =   285
            Left            =   1215
            TabIndex        =   45
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
            Left            =   1215
            TabIndex        =   43
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
         Begin MSMask.MaskEdBox mskdatereg 
            Height          =   285
            Left            =   3825
            TabIndex        =   46
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
         Begin MSMask.MaskEdBox mskendof 
            Height          =   285
            Left            =   7830
            TabIndex        =   65
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
            Left            =   3825
            TabIndex        =   44
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
            Left            =   2835
            TabIndex        =   236
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
            TabIndex        =   235
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
            Left            =   2835
            TabIndex        =   234
            Top             =   570
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
            Left            =   135
            TabIndex        =   233
            Top             =   210
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
            TabIndex        =   232
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
         TabIndex        =   237
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
            TabIndex        =   103
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
            TabIndex        =   101
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
            TabIndex        =   102
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
            TabIndex        =   240
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
            TabIndex        =   239
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
            TabIndex        =   238
            Top             =   270
            Width           =   660
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Dependencies"
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
         TabIndex        =   222
         Top             =   2160
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
            TabIndex        =   83
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
            TabIndex        =   75
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
            TabIndex        =   77
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
            TabIndex        =   79
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
            TabIndex        =   81
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
            TabIndex        =   85
            Top             =   1170
            Width           =   435
         End
         Begin MSMask.MaskEdBox mskdep2 
            Height          =   285
            Left            =   2610
            TabIndex        =   78
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
            TabIndex        =   76
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
            TabIndex        =   80
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
            TabIndex        =   82
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
            TabIndex        =   84
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
            TabIndex        =   230
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
            TabIndex        =   229
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
            TabIndex        =   228
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
            TabIndex        =   227
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
            TabIndex        =   226
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
            TabIndex        =   225
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
            TabIndex        =   224
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
            TabIndex        =   223
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
         Height          =   1500
         Left            =   -74865
         TabIndex        =   216
         Top             =   540
         Width           =   10050
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
            MaxLength       =   20
            TabIndex        =   74
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
            MaxLength       =   20
            TabIndex        =   71
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
            MaxLength       =   20
            TabIndex        =   68
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
            TabIndex        =   72
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
            TabIndex        =   69
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
            TabIndex        =   66
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
            TabIndex        =   67
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
            TabIndex        =   70
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
            Left            =   3345
            MaxLength       =   50
            TabIndex        =   73
            Top             =   1095
            Width           =   4140
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
            Left            =   8460
            TabIndex        =   258
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
            Left            =   2340
            TabIndex        =   221
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
            TabIndex        =   220
            Top             =   1215
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
            TabIndex        =   219
            Top             =   855
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
            TabIndex        =   218
            Top             =   495
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
            Left            =   5340
            TabIndex        =   217
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
         TabIndex        =   207
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
            TabIndex        =   100
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
            TabIndex        =   95
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
            TabIndex        =   91
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
            TabIndex        =   99
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
            TabIndex        =   97
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
            TabIndex        =   94
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
            TabIndex        =   90
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
            TabIndex        =   89
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
            TabIndex        =   92
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
            TabIndex        =   96
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
            TabIndex        =   98
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
            TabIndex        =   215
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
            TabIndex        =   214
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
            TabIndex        =   213
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
            TabIndex        =   212
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
            TabIndex        =   211
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
            TabIndex        =   210
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
            TabIndex        =   209
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
            TabIndex        =   208
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
         TabIndex        =   200
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
            TabIndex        =   121
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
            TabIndex        =   115
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
            TabIndex        =   109
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
            TabIndex        =   120
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
            TabIndex        =   114
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
            TabIndex        =   108
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
            TabIndex        =   119
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
            TabIndex        =   113
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
            TabIndex        =   107
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
            TabIndex        =   118
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
            TabIndex        =   112
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
            TabIndex        =   106
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
            TabIndex        =   117
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
            TabIndex        =   111
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
            TabIndex        =   105
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
            TabIndex        =   116
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
            TabIndex        =   110
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
            TabIndex        =   104
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
            TabIndex        =   206
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
            TabIndex        =   205
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
            TabIndex        =   204
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
            TabIndex        =   203
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
            TabIndex        =   202
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
            TabIndex        =   201
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
         TabIndex        =   193
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
            TabIndex        =   131
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
            TabIndex        =   132
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
            TabIndex        =   133
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
            TabIndex        =   134
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
            TabIndex        =   135
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
            TabIndex        =   136
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
            TabIndex        =   137
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
            TabIndex        =   138
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
            TabIndex        =   139
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
            TabIndex        =   140
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
            TabIndex        =   141
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
            TabIndex        =   142
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
            TabIndex        =   143
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
            TabIndex        =   144
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
            TabIndex        =   145
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
            TabIndex        =   199
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
            TabIndex        =   198
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
            TabIndex        =   197
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
            TabIndex        =   196
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
            TabIndex        =   195
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
            TabIndex        =   194
            Top             =   150
            Width           =   945
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1680
         Left            =   -74820
         TabIndex        =   183
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
            Left            =   6405
            MaxLength       =   40
            TabIndex        =   130
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
            Left            =   6405
            MaxLength       =   40
            TabIndex        =   127
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
            Left            =   6405
            MaxLength       =   40
            TabIndex        =   124
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
            TabIndex        =   122
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
            TabIndex        =   125
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
            TabIndex        =   128
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
            TabIndex        =   129
            Top             =   1125
            Width           =   3765
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
            TabIndex        =   126
            Top             =   735
            Width           =   3765
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
            TabIndex        =   123
            Top             =   375
            Width           =   3765
         End
         Begin VB.Label Label22 
            Caption         =   "(3)"
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   71
            Left            =   180
            TabIndex        =   192
            Top             =   1155
            Width           =   255
         End
         Begin VB.Label Label22 
            Caption         =   "(2)"
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   70
            Left            =   180
            TabIndex        =   191
            Top             =   795
            Width           =   255
         End
         Begin VB.Label Label22 
            Caption         =   "(1)"
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   69
            Left            =   180
            TabIndex        =   190
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
            TabIndex        =   189
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
            Left            =   7155
            TabIndex        =   188
            Top             =   135
            Width           =   1230
         End
         Begin VB.Label Label22 
            Caption         =   "(3)"
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   62
            Left            =   2310
            TabIndex        =   187
            Top             =   1155
            Width           =   255
         End
         Begin VB.Label Label22 
            Caption         =   "(2)"
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   61
            Left            =   2310
            TabIndex        =   186
            Top             =   795
            Width           =   255
         End
         Begin VB.Label Label22 
            Caption         =   "(1)"
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   59
            Left            =   2310
            TabIndex        =   185
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
            TabIndex        =   184
            Top             =   135
            Width           =   3510
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   -74865
         TabIndex        =   150
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
            Left            =   8760
            TabIndex        =   34
            Top             =   1980
            Width           =   1215
         End
         Begin VB.ComboBox cmbcivilstatus 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   8775
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1575
            Width           =   1245
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
            Left            =   5670
            TabIndex        =   17
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
            Left            =   5670
            TabIndex        =   18
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
            Left            =   6585
            TabIndex        =   30
            Top             =   2715
            Width           =   1125
         End
         Begin VB.TextBox txtrescer 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3840
            TabIndex        =   29
            Top             =   2745
            Width           =   1575
         End
         Begin VB.TextBox txtnbi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1080
            TabIndex        =   28
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
            TabIndex        =   22
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
            TabIndex        =   26
            Top             =   2355
            Width           =   1575
         End
         Begin VB.TextBox txthaircolor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1080
            TabIndex        =   25
            Top             =   2355
            Width           =   1695
         End
         Begin VB.TextBox txtacctno 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   8760
            TabIndex        =   35
            Top             =   2355
            Width           =   1215
         End
         Begin VB.TextBox txthouseno 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   945
            TabIndex        =   6
            Top             =   180
            Width           =   705
         End
         Begin VB.TextBox txtheight 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7155
            MaxLength       =   8
            TabIndex        =   16
            Top             =   1215
            Width           =   540
         End
         Begin VB.TextBox txtweight 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7170
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
            Left            =   8775
            MaxLength       =   15
            TabIndex        =   32
            Top             =   1260
            Width           =   1215
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
            Left            =   8760
            MaxLength       =   15
            TabIndex        =   31
            Top             =   900
            Width           =   1215
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
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   12
            Top             =   900
            Width           =   3015
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
            Left            =   5355
            MaxLength       =   50
            TabIndex        =   11
            Top             =   540
            Width           =   4575
         End
         Begin VB.TextBox txtbirthplace 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   19
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
            TabIndex        =   7
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
            TabIndex        =   8
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
            TabIndex        =   10
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
            Left            =   5760
            MaxLength       =   50
            TabIndex        =   13
            Top             =   900
            Width           =   1950
         End
         Begin VB.TextBox txtdistrict 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7425
            TabIndex        =   9
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
            Left            =   6570
            TabIndex        =   27
            Top             =   2355
            Width           =   1125
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
            TabIndex        =   15
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
            TabIndex        =   20
            Top             =   1635
            Width           =   1560
         End
         Begin MSMask.MaskEdBox mskbday 
            Height          =   285
            Left            =   1080
            TabIndex        =   14
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
            Left            =   6585
            TabIndex        =   24
            Top             =   1980
            Width           =   1125
            _ExtentX        =   1984
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
            TabIndex        =   23
            Top             =   1980
            Width           =   1575
            _ExtentX        =   2778
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
            Left            =   8730
            TabIndex        =   36
            Top             =   2745
            Width           =   1215
            _ExtentX        =   2143
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
         Begin VB.TextBox txtIssOn 
            Height          =   285
            Left            =   8775
            TabIndex        =   181
            Top             =   2745
            Width           =   1095
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
            Left            =   7800
            TabIndex        =   180
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
            Left            =   5640
            TabIndex        =   179
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
            TabIndex        =   178
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
            TabIndex        =   177
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
            TabIndex        =   176
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
            TabIndex        =   175
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
            TabIndex        =   174
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
            Left            =   7800
            TabIndex        =   173
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
            Left            =   7800
            TabIndex        =   172
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
            TabIndex        =   171
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
            Left            =   6525
            TabIndex        =   170
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
            Left            =   6525
            TabIndex        =   169
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
            Left            =   7800
            TabIndex        =   168
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
            TabIndex        =   167
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
            Left            =   7800
            TabIndex        =   166
            Top             =   1320
            Width           =   555
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Tel. No."
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
            TabIndex        =   165
            Top             =   960
            Width           =   570
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
            TabIndex        =   164
            Top             =   1020
            Width           =   1305
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Boarding House"
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
            Left            =   4065
            TabIndex        =   163
            Top             =   630
            Width           =   1125
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
            TabIndex        =   162
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
            TabIndex        =   161
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
            TabIndex        =   160
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
            TabIndex        =   159
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
            Left            =   5640
            TabIndex        =   158
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
            TabIndex        =   157
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
            TabIndex        =   156
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
            TabIndex        =   155
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
            TabIndex        =   154
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
            TabIndex        =   153
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
            Left            =   5640
            TabIndex        =   152
            Top             =   2445
            Width           =   705
         End
      End
   End
   Begin VB.Frame Frame7 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   4320
      TabIndex        =   262
      Top             =   45
      Width           =   6210
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
         Left            =   270
         TabIndex        =   285
         Top             =   45
         Width           =   5820
      End
   End
   Begin VB.CommandButton cmdlast 
      Height          =   255
      Left            =   10215
      Picture         =   "frmHRFinalEmployees.frx":0643
      Style           =   1  'Graphical
      TabIndex        =   254
      ToolTipText     =   "Last record"
      Top             =   585
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Height          =   255
      Left            =   9975
      Picture         =   "frmHRFinalEmployees.frx":078D
      Style           =   1  'Graphical
      TabIndex        =   253
      ToolTipText     =   "Next record"
      Top             =   585
      Width           =   255
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   255
      Left            =   9735
      Picture         =   "frmHRFinalEmployees.frx":08D7
      Style           =   1  'Graphical
      TabIndex        =   252
      ToolTipText     =   "Previous record"
      Top             =   585
      Width           =   255
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   255
      Left            =   9495
      Picture         =   "frmHRFinalEmployees.frx":0A21
      Style           =   1  'Graphical
      TabIndex        =   251
      ToolTipText     =   "Top record"
      Top             =   585
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
      MouseIcon       =   "frmHRFinalEmployees.frx":0B6B
      MousePointer    =   99  'Custom
      TabIndex        =   182
      Text            =   "Hide"
      Top             =   5625
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
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
         Left            =   3510
         Picture         =   "frmHRFinalEmployees.frx":0E75
         Style           =   1  'Graphical
         TabIndex        =   263
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
         Left            =   2925
         Picture         =   "frmHRFinalEmployees.frx":0F77
         Style           =   1  'Graphical
         TabIndex        =   261
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
         Left            =   2340
         Picture         =   "frmHRFinalEmployees.frx":10C1
         Style           =   1  'Graphical
         TabIndex        =   151
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
         Left            =   1170
         Picture         =   "frmHRFinalEmployees.frx":172B
         Style           =   1  'Graphical
         TabIndex        =   148
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
         Left            =   1755
         Picture         =   "frmHRFinalEmployees.frx":1C5D
         Style           =   1  'Graphical
         TabIndex        =   149
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
         Picture         =   "frmHRFinalEmployees.frx":1D5F
         Style           =   1  'Graphical
         TabIndex        =   147
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdnew 
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
         Picture         =   "frmHRFinalEmployees.frx":1EA9
         Style           =   1  'Graphical
         TabIndex        =   146
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
      Height          =   1320
      Left            =   135
      TabIndex        =   245
      Top             =   585
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   2328
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmHRFinalEmployees.frx":23DB
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
      Tab(0).Control(7)=   "txtfirstname"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtmiddlename"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtnickname"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtsurname"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtControlNo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtID"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdCopy"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtLastDate"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
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
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   300
         Text            =   "Date"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   ">>"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   294
         Top             =   180
         Width           =   420
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
         Left            =   2925
         MaxLength       =   20
         TabIndex        =   287
         Top             =   180
         Visible         =   0   'False
         Width           =   450
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
         MaxLength       =   12
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
         Height          =   285
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   1
         Top             =   540
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8940
         MaxLength       =   10
         TabIndex        =   4
         Top             =   540
         Width           =   1215
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
         Height          =   285
         Left            =   6420
         MaxLength       =   20
         TabIndex        =   3
         Top             =   540
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
         Height          =   285
         Left            =   3825
         MaxLength       =   20
         TabIndex        =   2
         Top             =   540
         Width           =   1575
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
         Left            =   7080
         TabIndex        =   301
         Top             =   120
         Width           =   930
      End
      Begin VB.Label Label1 
         Height          =   240
         Left            =   4545
         TabIndex        =   255
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
         Left            =   8100
         TabIndex        =   250
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
         Left            =   5460
         TabIndex        =   249
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
         Left            =   2970
         TabIndex        =   248
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
         TabIndex        =   247
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
         TabIndex        =   246
         Top             =   225
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmHRFinalEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CnnstrCopy As String
Public rs As ADODB.Recordset
Public prevControlNo As String
Public PrevCompany As String
Public PrevStartContract As String
Public PrevEndContract As String
Public Cid  As Integer
Public BrCode  As Integer
Public PaymentBasis As String
Public MonthlyRate
Public DailyRate
Public HourlyRate
Public WithEcolaPay   As Integer
Public EcolaAmount
Public DeptCode  As Integer
 
Public rsout As ADODB.Recordset
Public rsoutEdit As ADODB.Recordset
Dim rsdetails As ADODB.Recordset
Dim rsref As ADODB.Recordset
Dim rsdialect As ADODB.Recordset
Dim rssem As ADODB.Recordset
Dim rsemprec As ADODB.Recordset
Dim dep1bdate, dep2bdate, dep3bdate, dep4bdate, dep5bdate
Public TmpSurname As String
Public TmpFirstname As String
Public TmpMiddlename As String

Private Sub ChkEcola_Click()
    If cmdupdate.Enabled = True Then
        If ChkEcola.VALUE = 1 Then
            TxtEcolaAmount.Enabled = True
            TxtEcolaAmount.SetFocus
        Else
            TxtEcolaAmount.Enabled = False
            TxtEcolaAmount.Text = "0.00"
        End If
    End If
End Sub

Private Sub cmbempstatus_Click()
    Status = UCase(Mid(cmbempstatus.Text, 1, 1))
    If Status = "T" Or Status = "W" Or Status = "X" Or Status = "Y" Or Status = "Z" Then
        cmdCopy.Enabled = True
    Else
        cmdCopy.Enabled = False
    End If
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

Private Sub cmdCopy_Click()
If cmdupdate.Enabled = True Then
MsgBox "Please update or save first your task before you transfer.", vbInformation + vbInformation
Exit Sub
End If

Dim CName As String
    On Error GoTo b:
    If CompanyCode = 2 Then
        StrDbpath = GetSetting("DataSearch", "SE", "Path2")
        CName = "SAPI Manpower & Mgt. Services"
    ElseIf CompanyCode = 1 Then
        CName = "Jobs and Careers Inc."
        StrDbpath = GetSetting("DataSearch", "JCI", "Path2")
    End If
    If StrDbpath = "" Then MsgBox "Unable to connect to " & CName & ". Run the system for this company to establish the connection.": Exit Sub
    If MsgBox("Are you sure you want to move [" & Me.txtsurname.Text & "," & Me.txtfirstname.Text & "] to " & CName & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
     
    CnnstrCopy = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrDbpath
   'CnnstrCopy = "Driver={Microsoft Access Driver (*.mdb)};" & _
                        "Dbq=" & StrDbpath & ";" & _
                        "Uid=;Pwd=alpha;"
    intretval = OpenDatabaseAdo(cnn, Cnnstr)
    If intretval <> 0 Then
        MsgBox "Unable To Establish the Connection. ", vbCritical
    Exit Sub
    End If
    
    strSQL = "Select * from hremployees where firstname ='" & RTrim(txtfirstname.Text) & "'  and surname= '" & RTrim(txtsurname.Text) & "'"
    Dim rsemp As New ADODB.Recordset
    Set rsemp = ReturnRst(strSQL, CnnstrCopy, mstatic)
    If rsemp.EOF = False Then
        If MsgBox("CANNOT TRANFER THIS EMPLOYEE because the name [" & Me.txtsurname.Text & "," & Me.txtfirstname.Text & "]" & " is ALREADY EXIST in " & CName & " database." & vbCrLf & "" & vbCrLf & "Do you want to update this record in " & CName & "", vbYesNo + vbQuestion) = vbYes Then
          Call SaveHistory(True)
        Exit Sub
        Else
        
        Exit Sub
        End If
    End If
    
    PaymentBasis = IIf(cmbPayBasis.Text = "", " ", cmbPayBasis.Text)
    Cid = 0
    HourlyRate = "0.00"
    DailyRate = "0.00"
    MonthlyRate = "0.00"
         
    If ChkEcola.VALUE = 1 Then
        WithEcolaPay = 1
        EcolaAmount = IIf(TxtEcolaAmount.Text = "", "0.00", Format(TxtEcolaAmount.Text, "##0.00"))
    Else
        WithEcolaPay = 0
        EcolaAmount = 0
    End If
       
    DeptCode = 0
    BrCode = 0
     
    Dim rsCopy As New ADODB.Recordset
    Dim str As String
    If rsCopy.State = 1 Then rsCopy.Close
    'str = "Select * from hremployees where  controlno= '" & TxtControlNo.Text & "'"
    'Note we should use the firstname and lastname since control number is not reliable because it is editable by user.
    str = "Select * from hremployees where firstname = '" & txtfirstname.Text & "' and surname = '" & txtsurname.Text & "'"
    Set rsCopy = ReturnRst(str, CnnstrCopy, mstatic)
    If rsCopy.EOF = True Then
'              sqlString = "Insert into hremployees (ControlNo , FirstName, Empposition, age, middlename, Surname, nickname, BIRTHDATE, Address1, houseno, street, barangay,district, city,province1, province, contact, caddress,cphone,gender,Height,Weight,Telnum1,Telnum2,civilstatus,examgrade,interviewgrade,medicalgrade,sss,philhealth, pagibig,AcctNo, citizen, birthplace,haircolor, eyecolor, Religion, taxid, [Primary],dateprimary, secondary, datesecondary,college, datecollege, degree, Major, postgraduate," & _
'             " datepostgraduate, postgraduatecourse, fathername, faddress, fwork, mothername, maddress, mwork, nbi, rescer, issueat, issueon, datehired, dateprob, dateregular, dateresign, employmentstatus, supervisor, reasonres, endofcontract, ClientCode, brcode, " & _
'             " PaymentBasis,HourlyRate,DailyRate, MonthlyRate, StartContract, EndContract, NoofMonthsHired, cbamount, cbpayment,cbstartpayment, cbendpayment,  MiscFee,MFPayment,MFStartPayment,MFEndPayment, PlacementFee,PFpayment,PFstartpayment,PFendpayment, WithEcolaPay,  ecolaamount, [USERID], deptcode) " & _
'             " values ( '" & TxtControlNo.Text & "'," & ChkVal(txtfirstname) & "," & ChkVal(txtfirstchoice) & "," & IIf(txtage.Text = "", 0, txtage.Text) & "," & ChkVal(txtmiddlename) & "," & ChkVal(txtsurname) & "," & ChkVal(txtnickname) & "," & ChkMsk(mskbday) & "," & ChkVal(txtaddress1) & "," & ChkVal(txthouseno) & "," & ChkVal(txtstreet) & "," & ChkVal(txtbarangay) & "," & ChkVal(txtdistrict) & "," & ChkVal(txtcity) & "," & ChkVal(txtprovince) & "" & _
'             "," & ChkVal(txtprovinceaddress) & "," & ChkVal(txtemergency) & "," & ChkVal(txteraddress) & "," & ChkVal(txtertel) & ",'" & IIf(optmale.Value, "M", "F") & "'," & ChkVal(txtheight) & "," & ChkVal(txtweight) & "," & ChkVal(txttel1) & "," & ChkVal(txttel2) & ",'" & IIf(cmbcivilstatus.Text = "", "S", cmbcivilstatus.Text) & "'," & ChkVal(txtexam) & "," & ChkVal(txtinterview) & "," & ChkVal(txtmedical) & "," & ChkMsk(msksss) & "," & ChkVal(txtphilhealth) & "," & ChkVal(txtpagibig) & "," & ChkVal(txtacctno) & "" & _
'             "," & ChkVal(txtcitizen) & "," & ChkVal(txtbirthplace) & "," & ChkVal(txthaircolor) & "," & ChkVal(txteyecolor) & "," & ChkVal(txtreligion) & "," & ChkMsk(msktax) & "," & ChkVal(txtprimarysch) & "," & ChkVal(txtdateprimary) & "," & ChkVal(txtsecsch) & "," & ChkVal(txtdatesec) & "," & ChkVal(txtcollegesch) & "," & ChkVal(txtdatecollege) & "," & ChkVal(txtcollegecourse) & "," & ChkVal(txtcollegemajor) & "," & ChkVal(txtpostsch) & "," & ChkVal(txtdatepost) & "" & _
'             "," & ChkVal(txtpostcourse) & "," & ChkVal(txtfathername) & "," & ChkVal(txtfjob) & "," & ChkVal(txtfaddress) & "," & ChkVal(txtmothername) & "," & ChkVal(txtmaddress) & "," & ChkVal(txtmjob) & "," & ChkVal(txtnbi) & "," & ChkVal(txtrescer) & "," & ChkVal(txtissueat) & "," & ChkMsk(mskissue) & "," & ChkMsk(MskDateHired) & "," & ChkMsk(mskdateprob) & "," & ChkMsk(mskdatereg) & "" & _
'             "," & ChkMsk(mskdateres) & ",'" & Left(cmbempstatus.Text, 1) & "'," & ChkVal(txtsupervisor) & "," & ChkVal(txtreasonresign) & "," & ChkMsk(mskendof) & "," & 0 & "," & BrCode & ",'" & PaymentBasis & "'," & HourlyRate & "," & DailyRate & "," & MonthlyRate & "," & ChkMsk(mskStartContract) & "," & ChkMsk(mskEndContract) & "," & IIf(txtMonthsHired.Text = "", 5, txtMonthsHired.Text) & "" & _
'             "," & IIf(txtCBtotal = "", 0, Format(txtCBtotal, "##0.00")) & "," & IIf(txtCBpayment = "", 0, Format(txtCBpayment, "##0.00")) & "," & ChkMsk(mskStartCB) & "," & ChkMsk(mskEndCB) & "" & _
'             "," & IIf(TxtMFAmount = "", 0, Format(TxtMFAmount, "##0.00")) & "," & IIf(TxtMFPayment = "", 0, Format(TxtMFPayment, "##0.00")) & "," & ChkMsk(MskMFStartPayment) & "," & ChkMsk(MskMFEndPayment) & "" & _
'              "," & IIf(TxtMFAmount = "", 0, Format(TxtMFAmount, "##0.00")) & "," & IIf(txtPFPayment = "", 0, Format(txtPFPayment, "##0.00")) & "," & ChkMsk(MskPFStartPayment) & "," & ChkMsk(MskPFEndPayment) & "" & _
'             ",'" & WithEcolaPay & "'," & EcolaAmount & ",null," & DeptCode & ")"

'COMMENTED BY Ramil 11052006
             ' sqlString = "Insert into hremployees (ControlNo , FirstName, Empposition, age, middlename, Surname, nickname, BIRTHDATE, Address1, houseno, street, barangay,district, city,province1, province, contact, caddress,cphone,gender,Height,Weight,Telnum1,Telnum2,civilstatus,examgrade,interviewgrade,medicalgrade,sss,philhealth, pagibig,AcctNo, citizen, birthplace,haircolor, eyecolor, Religion, taxid, [Primary],dateprimary, secondary, datesecondary,college, datecollege, degree, Major, postgraduate," & _
             " datepostgraduate, postgraduatecourse, fathername, faddress, fwork, mothername, maddress, mwork, nbi, rescer, issueat, issueon, datehired, dateprob, dateregular, dateresign, employmentstatus, supervisor, reasonres, endofcontract, ClientCode, brcode, " & _
             " PaymentBasis,HourlyRate,DailyRate, MonthlyRate, StartContract, EndContract, NoofMonthsHired, cbamount, cbpayment,cbstartpayment, cbendpayment,  MiscFee,MFPayment,MFStartPayment,MFEndPayment, PlacementFee,PFpayment,PFstartpayment,PFendpayment, WithEcolaPay,  ecolaamount, [USERID], deptcode) " & _
             " values ( '" & TxtControlNo.Text & "'," & ChkVal(txtfirstname) & "," & ChkVal(txtfirstchoice) & "," & IIf(txtage.Text = "", 0, txtage.Text) & "," & ChkVal(txtmiddlename) & "," & ChkVal(txtsurname) & "," & ChkVal(txtnickname) & "," & ChkMsk(mskbday) & "," & ChkVal(txtaddress1) & "," & ChkVal(txthouseno) & "," & ChkVal(txtstreet) & "," & ChkVal(txtbarangay) & "," & ChkVal(txtdistrict) & "," & ChkVal(txtcity) & "," & ChkVal(txtprovince) & "" & _
             "," & ChkVal(txtprovinceaddress) & "," & ChkVal(txtemergency) & "," & ChkVal(txteraddress) & "," & ChkVal(txtertel) & ",'" & IIf(optmale.VALUE, "M", "F") & "'," & ChkVal(txtheight) & "," & ChkVal(txtweight) & "," & ChkVal(txttel1) & "," & ChkVal(txttel2) & ",'" & IIf(cmbcivilstatus.Text = "", "S", cmbcivilstatus.Text) & "'," & ChkVal(txtexam) & "," & ChkVal(txtinterview) & "," & ChkVal(txtmedical) & "," & ChkMsk(msksss) & "," & ChkVal(txtphilhealth) & "," & ChkVal(txtpagibig) & "," & ChkVal(txtacctno) & "" & _
             "," & ChkVal(txtcitizen) & "," & ChkVal(txtbirthplace) & "," & ChkVal(txthaircolor) & "," & ChkVal(txteyecolor) & "," & ChkVal(txtreligion) & "," & ChkMsk(msktax) & "," & ChkVal(txtprimarysch) & "," & ChkVal(txtdateprimary) & "," & ChkVal(txtsecsch) & "," & ChkVal(txtdatesec) & "," & ChkVal(txtcollegesch) & "," & ChkVal(txtdatecollege) & "," & ChkVal(txtcollegecourse) & "," & ChkVal(txtcollegemajor) & "," & ChkVal(txtpostsch) & "," & ChkVal(txtdatepost) & "" & _
             "," & ChkVal(txtpostcourse) & "," & ChkVal(txtfathername) & "," & ChkVal(txtfjob) & "," & ChkVal(txtfaddress) & "," & ChkVal(txtmothername) & "," & ChkVal(txtmaddress) & "," & ChkVal(txtmjob) & "," & ChkVal(txtnbi) & "," & ChkVal(txtrescer) & "," & ChkVal(txtissueat) & "," & ChkMsk(mskissue) & "," & ChkMsk(MskDateHired) & "," & ChkMsk(mskdateprob) & "," & ChkMsk(mskdatereg) & "" & _
             "," & ChkMsk(mskdateres) & ",'" & Left(cmbempstatus.Text, 1) & "', " & ChkVal(txtsupervisor) & "," & ChkVal(txtreasonresign) & "," & ChkMsk(mskendof) & "," & 0 & "," & BrCode & ",'" & PaymentBasis & "', " & HourlyRate & "," & DailyRate & "," & MonthlyRate & ",null,null," & IIf(txtMonthsHired.Text = "", 5, txtMonthsHired.Text) & "" & _
             "," & IIf(txtCBtotal = "", 0, Format(txtCBtotal, "##0.00")) & "," & IIf(txtCBpayment = "", 0, Format(txtCBpayment, "##0.00")) & "," & ChkMsk(mskStartCB) & "," & ChkMsk(mskEndCB) & "" & _
             "," & IIf(TxtMFAmount = "", 0, Format(TxtMFAmount, "##0.00")) & ", " & IIf(TxtMFPayment = "", 0, Format(TxtMFPayment, "##0.00")) & "," & ChkMsk(MskMFStartPayment) & "," & ChkMsk(MskMFEndPayment) & "" & _
              "," & IIf(TxtMFAmount = "", 0, Format(TxtMFAmount, "##0.00")) & "," & IIf(txtPFPayment = "", 0, Format(txtPFPayment, "##0.00")) & "," & ChkMsk(MskPFStartPayment) & "," & ChkMsk(MskPFEndPayment) & "" & _
             "," & IIf(WithEcolaPay = 1, True, False) & ", " & EcolaAmount & ",null," & DeptCode & ")"

  sqlString = "Insert into hremployees (ControlNo , FirstName, Empposition, age, middlename, Surname, nickname, BIRTHDATE, Address1, houseno, street, barangay,district, city,province1, province, contact, caddress,cphone,gender,Height,Weight,Telnum1,Telnum2,civilstatus,examgrade,interviewgrade,medicalgrade,sss,philhealth, pagibig,AcctNo, citizen, birthplace,haircolor, eyecolor, Religion, taxid, [Primary],dateprimary, secondary, datesecondary,college, datecollege, degree, Major, postgraduate," & _
       " datepostgraduate, postgraduatecourse, fathername, faddress, fwork, mothername, maddress, mwork, nbi, rescer, issueat, issueon, datehired, dateprob, dateregular, dateresign, employmentstatus, supervisor, reasonres, endofcontract, ClientCode, brcode, " & _
       " PaymentBasis,HourlyRate,DailyRate, MonthlyRate, StartContract, EndContract, NoofMonthsHired, cbamount, cbpayment,cbstartpayment, cbendpayment,  MiscFee,MFPayment,MFStartPayment,MFEndPayment, PlacementFee,PFpayment,PFstartpayment,PFendpayment, WithEcolaPay,  ecolaamount, [USERID], deptcode,remarks, LastUpdate) " & _
       " values ( '" & TxtControlNo.Text & "'," & ChkVal(txtfirstname) & "," & ChkVal(txtfirstchoice) & "," & IIf(txtage.Text = "", 0, txtage.Text) & "," & ChkVal(txtmiddlename) & "," & ChkVal(txtsurname) & "," & ChkVal(txtnickname) & "," & ChkMsk(mskbday) & "," & ChkVal(txtaddress1) & "," & ChkVal(txthouseno) & "," & ChkVal(txtstreet) & "," & ChkVal(txtbarangay) & "," & ChkVal(txtdistrict) & "," & ChkVal(txtcity) & "," & ChkVal(txtprovince) & "" & _
       "," & ChkVal(txtprovinceaddress) & "," & ChkVal(txtemergency) & "," & ChkVal(txteraddress) & "," & ChkVal(txtertel) & ",'" & IIf(optmale.VALUE, "M", "F") & "','" & Replace(txtheight, "'", """") & "'," & ChkVal(txtweight) & "," & ChkVal(txttel1) & "," & ChkVal(txttel2) & ",'" & IIf(cmbcivilstatus.Text = "", "S", cmbcivilstatus.Text) & "'," & ChkVal(txtexam) & "," & ChkVal(txtinterview) & "," & ChkVal(txtmedical) & "," & ChkMsk(msksss) & "," & ChkVal(txtphilhealth) & "," & ChkVal(txtpagibig) & "," & ChkVal(txtacctno) & "" & _
       "," & ChkVal(txtcitizen) & "," & ChkVal(txtbirthplace) & "," & ChkVal(txthaircolor) & "," & ChkVal(txteyecolor) & "," & ChkVal(txtreligion) & "," & ChkMsk(msktax) & "," & ChkVal(txtprimarysch) & "," & ChkVal(txtdateprimary) & "," & ChkVal(txtsecsch) & "," & ChkVal(txtdatesec) & "," & ChkVal(txtcollegesch) & "," & ChkVal(txtdatecollege) & "," & ChkVal(txtcollegecourse) & "," & ChkVal(txtcollegemajor) & "," & ChkVal(txtpostsch) & "," & ChkVal(txtdatepost) & "" & _
       "," & ChkVal(txtpostcourse) & "," & ChkVal(txtfathername) & "," & ChkVal(txtfjob) & "," & ChkVal(txtfaddress) & "," & ChkVal(txtmothername) & "," & ChkVal(txtmaddress) & "," & ChkVal(txtmjob) & "," & ChkVal(txtnbi) & "," & ChkVal(txtrescer) & "," & ChkVal(txtissueat) & "," & ChkMsk(mskissue) & "," & ChkMsk(MskDateHired) & "," & ChkMsk(mskdateprob) & "," & ChkMsk(mskdatereg) & "" & _
       "," & ChkMsk(mskdateres) & ",'" & Left(cmbempstatus.Text, 1) & "'," & ChkVal(txtsupervisor) & "," & ChkVal(txtreasonresign) & "," & ChkMsk(mskendof) & "," & Cid & "," & BrCode & ",'" & PaymentBasis & "'," & HourlyRate & "," & DailyRate & "," & MonthlyRate & "," & ChkMsk(mskStartContract) & "," & ChkMsk(mskEndContract) & "," & IIf(txtMonthsHired.Text = "", 5, txtMonthsHired.Text) & "" & _
       "," & IIf(txtCBtotal = "", 0, Format(txtCBtotal, "##0.00")) & "," & IIf(txtCBpayment = "", 0, Format(txtCBpayment, "##0.00")) & "," & ChkMsk(mskStartCB) & "," & ChkMsk(mskEndCB) & "" & _
       "," & IIf(TxtMFAmount = "", 0, Format(TxtMFAmount, "##0.00")) & "," & IIf(TxtMFPayment = "", 0, Format(TxtMFPayment, "##0.00")) & "," & ChkMsk(MskMFStartPayment) & "," & ChkMsk(MskMFEndPayment) & "" & _
        "," & IIf(txtPFAmount = "", 0, Format(txtPFAmount, "##0.00")) & "," & IIf(txtPFPayment = "", 0, Format(txtPFPayment, "##0.00")) & "," & ChkMsk(MskPFStartPayment) & "," & ChkMsk(MskPFEndPayment) & "" & _
       ",'" & WithEcolaPay & "'," & EcolaAmount & ",'" & operatorID & "'," & DeptCode & ",'" & IIf(Me.txtreasonresign.Text = "", " ", txtreasonresign.Text) & "', '" & Format(Now, "mm/dd/yyyy") & "')"
       
             Call Action_Query(sqlString, CnnstrCopy)
             'edited by ramil 10042006
             On Error GoTo bb:
           Call SaveHistory(False)
bb:
        Else
            MsgBox "CANNOT TRANSFER THIS EMPLOYEE because the control no. [" & TxtControlNo.Text & "] is ALREADY USED in " & CName, vbInformation
        Exit Sub
     End If
        
'OTHERS
    Dim rsOthers As New ADODB.Recordset
    
    Dim dept1Bdate, dept2Bdate, dept3Bdate, dept4Bdate, dept5Bdate
    dep1bdate = IIf(mskdep1.Text = "__/__/____", ChkMsk(mskdep1), Format(mskdep1.Text, "mm/dd/yyyy"))
    dep2bdate = IIf(mskdep2.Text = "__/__/____", ChkMsk(mskdep2), Format(mskdep2.Text, "mm/dd/yyyy"))
    dep3bdate = IIf(mskdep3.Text = "__/__/____", ChkMsk(mskdep3), Format(mskdep3.Text, "mm/dd/yyyy"))
    dep4bdate = IIf(mskdep4.Text = "__/__/____", ChkMsk(mskdep4), Format(mskdep4.Text, "mm/dd/yyyy"))
    dep5bdate = IIf(mskdep5.Text = "__/__/____", ChkMsk(mskdep5), Format(mskdep5.Text, "mm/dd/yyyy"))

    If rsOthers.State = 1 Then rsOthers.Close 'HRDETAILS
    str = "Select * from hrdetails where  employeenum= '" & TxtControlNo.Text & "'"
    Set rsOthers = ReturnRst(str, CnnstrCopy, mstatic)
    If rsOthers.EOF = True Then
       sqlString = "INSERT INTO hrdetails " & _
      "(employeenum,spousename,spouseaddress,spousework,Dependents ,dep1name,dep2name,dep3name,dep4name,dep5name,dep1bdate,dep2bdate,dep3bdate,dep4bdate,dep5bdate )" & _
      " values ('" & TxtControlNo.Text & "'," & ChkVal(txtspousename) & "," & ChkVal(txtspouseaddress) & "," & ChkVal(txtsjob) & "," & ChkVal(txtnumdep) & "," & ChkVal(txtdep1) & "," & ChkVal(txtdep2) & "," & ChkVal(txtdep3) & "," & ChkVal(txtdep4) & "," & ChkVal(txtdep5) & "," & dep1bdate & "," & dep2bdate & "," & dep3bdate & "," & dep4bdate & "," & dep5bdate & ")"
      Call Action_Query(sqlString, CnnstrCopy)
    End If
    
    If rsOthers.State = 1 Then rsOthers.Close 'HRCHARREFERENCE
    str = "Select * from hrcharreference where  employeenum= '" & TxtControlNo.Text & "'"
    Set rsOthers = ReturnRst(str, CnnstrCopy, mstatic)
    If rsOthers.EOF = True Then
        sqlString = "INSERT INTO hrcharreference " & _
          "(employeenum,name,occupation,address,telnum,relation,name1,occupation1,address1,telnum1 " & _
          ", relation1,name2,occupation2,address2,telnum2,relation2 )" & _
          " values ('" & TxtControlNo.Text & "'," & ChkVal(txtrefname1) & "," & ChkVal(txtrefjob1) & "," & ChkVal(txtrefaddress1) & "," & ChkVal(txtrefphone1) & "," & ChkVal(txtrefrel1) & "," & ChkVal(txtrefname2) & "," & ChkVal(txtrefjob2) & "," & ChkVal(txtrefaddress2) & "," & ChkVal(txtrefphone2) & "," & ChkVal(txtrefrel2) & "," & ChkVal(txtrefname3) & "," & ChkVal(txtrefjob3) & "," & ChkVal(txtrefjob3) & "," & ChkVal(txtrefphone3) & "," & ChkVal(txtrefrel3) & ")"
          Call Action_Query(sqlString, CnnstrCopy)
     End If
    
    If rsOthers.State = 1 Then rsOthers.Close 'EMPLOYMENT RECORD
    str = "Select * from hrEmploymentRecord where  employeenum= '" & TxtControlNo.Text & "'"
    Set rsOthers = ReturnRst(str, CnnstrCopy, mstatic)
    If rsOthers.EOF = True Then
        sqlString = "INSERT INTO hrEmploymentRecord " & _
          "(employeenum,company,[from],[to],[position],reason,superior,company1,from1,to1," & _
          " position1,reason1,superior1,company2,from2,to2,position2,reason2,superior2)" & _
          " values ('" & TxtControlNo.Text & "'," & ChkVal(txtnameofcompany1) & "," & ChkVal(txtfrom1) & "," & ChkVal(txtto1) & "," & ChkVal(txtpos1) & "," & ChkVal(txtreason1) & "," & ChkVal(txtsuperior1) & "," & _
                                            " " & ChkVal(txtnameofcompany2) & "," & ChkVal(txtfrom2) & "," & ChkVal(txtto2) & "," & ChkVal(txtpos2) & "," & ChkVal(txtreason2) & "," & ChkVal(txtsuperior2) & "," & _
                                            " " & ChkVal(txtnameofcompany3) & "," & ChkVal(txtfrom3) & "," & ChkVal(txtto3) & "," & ChkVal(txtpos3) & "," & ChkVal(txtreason3) & "," & ChkVal(txtsuperior3) & ")"
               Call Action_Query(sqlString, CnnstrCopy)
     End If

    If rsOthers.State = 1 Then rsOthers.Close 'HR SEMINAR
    str = "Select * from hrseminar where  employeenum= '" & TxtControlNo.Text & "'"
    Set rsOthers = ReturnRst(str, CnnstrCopy, mstatic)
    If rsOthers.EOF = True Then
    sqlString = "INSERT INTO hrseminar " & _
      "(employeenum,seminar1,seminar1date,seminar2,seminar2date,seminar3,seminar3date)" & _
      " values ('" & TxtControlNo.Text & "'," & ChkVal(txtseminar1) & "," & ChkVal(txtdateseminar1) & "," & _
                                        " " & ChkVal(txtseminar2) & "," & ChkVal(txtdateseminar2) & "," & _
                                        " " & ChkVal(txtseminar3) & "," & ChkVal(txtdateseminar3) & ")"
     Call Action_Query(sqlString, CnnstrCopy)
    End If

    If rsOthers.State = 1 Then rsOthers.Close 'HRDIALECTS
    str = "Select * from hrdialects where  employeenum= '" & TxtControlNo.Text & "'"
    Set rsOthers = ReturnRst(str, CnnstrCopy, mstatic)
    If rsOthers.EOF = True Then
    sqlString = "INSERT INTO hrdialects " & _
      "(employeenum,dialect1,dialect2,dialect3) " & _
      " values ('" & TxtControlNo.Text & "'," & ChkVal(txtdialect1) & "," & ChkVal(txtdialect2) & "," & ChkVal(txtdialect3) & ")"
           Call Action_Query(sqlString, CnnstrCopy)
    End If
   'frmStatus.Show 1
   MsgBox "[" & Me.txtsurname.Text & "," & Me.txtfirstname.Text & "] successfully moved to " & CName & " " & vbCrLf & " " & vbCrLf & "Pls. update the control no., client name and contract date of this employee.", vbInformation
 Exit Sub
b:
 MsgBox "Unable to tranfer this employee... Pls check the employee's information!", vbInformation
End Sub

Private Sub cmdCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If cmdCopy.ToolTipText = "" Then
'    If CompanyCode = 2 Then
'            CompanyName = "SAPI Manpower & Mgt. Services"
'    ElseIf CompanyCode = 1 Then
'        CompanyName = "Jobs and Careers Inc."
'    End If
'
'    cmdCopy.ToolTipText = "Move to " & CompanyName
'    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Load()
Call TextEnabled(False)
'Me.Visible = False
On Error Resume Next
   DoEvents
   SSTab1.Tab = 1
    FeMP = True
    '  Me.Top = 0 '(mdiMain.Height - Me.Height) / 8
    '  Me.Left = (mdiMain.Width - Me.Width) / 2
    Me.Left = 0
    Me.Top = 0
    Dim rstclient As ADODB.Recordset
    Dim rstBranch As ADODB.Recordset
    Dim rstDept As ADODB.Recordset
    Set rsout = New ADODB.Recordset
    Set rsdetails = New ADODB.Recordset
    Set rsref = New ADODB.Recordset
    Set rsdialect = New ADODB.Recordset
    Set rsemprec = New ADODB.Recordset
    Set rscivil = New ADODB.Recordset
    Set rssem = New ADODB.Recordset
    Set rsempstatus = New ADODB.Recordset
    Set rscivil = cnDCPI.Execute("Select * from TaxExemptioncode")
    Do While Not rscivil.EOF
        DoEvents
        cmbcivilstatus.AddItem rscivil!TaxExemptioncode
        rscivil.MoveNext
        Loop
        rscivil.Close
    Set rsempstatus = cnDCPI.Execute("Select * from HRemploymentstatus")
    Do While Not rsempstatus.EOF
        DoEvents
        cmbempstatus.AddItem UCase(rsempstatus!employmentstatus & "-" & rsempstatus!EmploymentStatusdesc)
        rsempstatus.MoveNext
    Loop
    rsempstatus.Close

    Set rstclient = cnDCPI.Execute("select * from clienttable order by clientname")
    Do While Not rstclient.EOF
        DoEvents
        cmbCompany.AddItem rstclient!ClientName
        cmbCompany.ItemData(cmbCompany.NewIndex) = rstclient!ClientCode
        rstclient.MoveNext
    Loop
    rstclient.Close
    
    Set rstBranch = cnDCPI.Execute("select * from HRBRANCH order by brname")
    Do While Not rstBranch.EOF
    DoEvents
        cmbBranch.AddItem UCase(IIf(IsNull(rstBranch!brname), "", rstBranch!brname))
        cmbBranch.ItemData(cmbBranch.NewIndex) = rstBranch!BrCode
        rstBranch.MoveNext
    Loop
     rstBranch.Close
        
    Set rstDept = cnDCPI.Execute("select * from HRDEPT order by deptname")
    Do While Not rstDept.EOF
        DoEvents
        cmbDept.AddItem UCase(rstDept!deptname)
        cmbDept.ItemData(cmbDept.NewIndex) = rstDept!DeptCode
        rstDept.MoveNext
    Loop
        rstDept.Close
        
    Set rsout = New ADODB.Recordset
    Set rsdetails = New ADODB.Recordset
    Set rsref = New ADODB.Recordset
    Set rsdialect = New ADODB.Recordset
    Set rsemprec = New ADODB.Recordset
    Set rsem = New ADODB.Recordset
    
    rsout.Open "select * from HREmployees order by  surname", cnDCPI, adOpenKeyset, adLockOptimistic
    rsdetails.Open "Select * from HRDetails", cnDCPI, adOpenKeyset, adLockOptimistic
    rsref.Open "Select * from HRCharReference", cnDCPI, adOpenKeyset, adLockOptimistic
    rsdialect.Open "Select * from HRDialects", cnDCPI, adOpenKeyset, adLockOptimistic
    rsemprec.Open "Select * from HREmploymentRecord", cnDCPI, adOpenKeyset, adLockOptimistic
    rssem.Open "select * from HRSeminar", cnDCPI, adOpenKeyset, adLockOptimistic
    
    If rsout.RecordCount > 0 Then
        cmdundo.Enabled = False
        cmdupdate.Enabled = False
    Else
       Call clearscreen
       cmdfirst.Enabled = False
       cmdPrev.Enabled = False
       cmdNext.Enabled = False
       cmdlast.Enabled = False
       cmdedit.Enabled = False
       cmddelete.Enabled = False
       cmdundo.Enabled = False
       cmdupdate.Enabled = False
    End If
    empno = TxtControlNo.Text
    
    Dim strSQL As String
        If AddFlag = True Then
            Call cmdnew_Click
        End If
        If EditFlag = True Then
             rsout.Find "employeenum =" & frmHRMainEmployee.TxtEmployeeNum.Text
             Call cmdEdit_Click
                Showdata rsout
            End If
 
    Set rstclient = Nothing
    Set rstBranch = Nothing
    Set rstDept = Nothing
    Set rsdetails = Nothing
    Set rsref = Nothing
    Set rsdialect = Nothing
    Set rsemprec = Nothing
    Set rscivil = Nothing
    Set rssem = Nothing
    Set rsempstatus = Nothing
End Sub

Public Sub Showdata(rec As ADODB.Recordset)
    
    On Error Resume Next
    DoEvents
    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    DoEvents
    If rec.EOF = True Then
        Call clearscreen
        cmdedit.Enabled = False
        cmddelete.Enabled = False
    Else
        txtID.Text = rec!employeenum
        TxtControlNo.Text = IIf(IsNull(rec!ControlNo), "<controno>", rec!ControlNo)
        txtfirstchoice.Text = IIf(IsNull(rec!Empposition), "", rec!Empposition)
        txtnickname.Text = IIf(IsNull(rec!nickname), "", rec!nickname)
        txtsurname.Text = IIf(IsNull(rec!Surname), "", rec!Surname)
        txtfirstname.Text = IIf(IsNull(rec!FirstName), "", rec!FirstName)
     
        txtLastDate.Text = IIf(IsNull(rec!LastUpdate), "__/__/____", rec!LastUpdate)
        txtmiddlename.Text = IIf(IsNull(rec!middlename), "", rec!middlename)
        TmpMiddlename = txtmiddlename.Text
        TmpSurname = Me.txtsurname.Text
        TmpFirstname = Me.txtfirstname.Text
    
        txtaddress1.Text = IIf(IsNull(rec!Address1), "", rec!Address1)
        txthouseno.Text = IIf(IsNull(rec!houseno), "", rec!houseno)
        txtstreet.Text = IIf(IsNull(rec!street), "", rec!street)
        txtbarangay.Text = IIf(IsNull(rec!barangay), "", rec!barangay)
        txtdistrict.Text = IIf(IsNull(rec!district), "", rec!district)
        txtprovince.Text = IIf(IsNull(rec!province1), "", rec!province1)
        txtcity.Text = IIf(IsNull(rec!city), "", rec!city)
        txtprovinceaddress.Text = IIf(IsNull(rec!province), "", rec!province)
        mskbday.Text = IIf(IsNull(rec!BIRTHDATE) = True Or rec!BIRTHDATE = "", "__/__/____", Format(rec!BIRTHDATE, "mm/dd/yyyy"))
        If mskbday.Text = "__/__/____" Or IsNull(rec!BIRTHDATE) = True Then
            txtage.Text = "" 'IIf(IsNull(rec!age), 0, rec!age)
        Else
            
            txtage.Text = DateDiff("yyyy", Format(mskbday, "mm/dd/yyyy"), Date)
            If rec!age <> txtage.Text Then
             sqlString = "Update hremployees set age=" & txtage.Text & "" & _
                          " where  employeenum = " & txtID.Text & " "
                    Call Action_Query(sqlString, Cnnstr)
            End If
            
        End If
        
        txtemergency.Text = IIf(IsNull(rec!contact), "", rec!contact)
        txteraddress.Text = IIf(IsNull(rec!caddress), "", rec!caddress)
        txtertel.Text = IIf(IsNull(rec!cphone), "", rec!cphone)
        Me.Caption = "EMPLOYEES INFORMATION   [Last update by: " + UCase(GETuSERNAME(IIf(IsNull(rec!USERID), "0", rec!USERID))) + "]"
        If (rec!gender) = "M" Then
            optmale.VALUE = True
            optfemale.VALUE = False
        ElseIf (rec!gender) = "F" Then
            optmale.VALUE = False
            optfemale.VALUE = True
        Else
            optmale.VALUE = False
            optfemale.VALUE = False
        End If
        txtheight.Text = IIf(IsNull(rec!Height), "", Replace(rec!Height, """", "'"))
        txtweight.Text = IIf(IsNull(rec!Weight), "", rec!Weight)
        txttel1.Text = IIf(IsNull(rec!Telnum1), "", rec!Telnum1)
        txttel2.Text = IIf(IsNull(rec!Telnum2), "", rec!Telnum2)
        txtexam.Text = IIf(IsNull(rec!examgrade), "", rec!examgrade)
        txtinterview.Text = IIf(IsNull(rec!interviewgrade), "", rec!interviewgrade)
        txtmedical.Text = IIf(IsNull(rec!medicalgrade), "", rec!medicalgrade)
        msksss.Text = IIf(IsNull(rec!sss), "__-_______-_", rec!sss)
        txtphilhealth = IIf(IsNull(rec!PhilHealth), "", rec!PhilHealth)
        txtpagibig.Text = IIf(IsNull(rec!pagibig), "", rec!pagibig)
        txtacctno.Text = IIf(IsNull(rec!AcctNo), "", rec!AcctNo)
        mskissue.Text = IIf(IsNull(rec!issueon), "__/__/____", rec!issueon)
        txtbirthplace.Text = IIf(IsNull(rec!birthplace), "", rec!birthplace)
        txthaircolor.Text = IIf(IsNull(rec!haircolor), "", rec!haircolor)
        txteyecolor.Text = IIf(IsNull(rec!eyecolor), "", rec!eyecolor)
        txtreligion.Text = IIf(IsNull(rec!Religion), "", rec!Religion)
             
        msktax.Text = IIf(IsNull(rec!taxid), "___-___-___", rec!taxid)
        txtprimarysch.Text = IIf(IsNull(rec!Primary), "", rec!Primary)
        txtdateprimary.Text = IIf(IsNull(rec!dateprimary), "", rec!dateprimary)
        txtsecsch.Text = IIf(IsNull(rec!secondary), "", rec!secondary)
        txtdatesec.Text = IIf(IsNull(rec!datesecondary), "", rec!datesecondary)
        txtcollegesch.Text = IIf(IsNull(rec!college), "", rec!college)
        txtdatecollege.Text = IIf(IsNull(rec!datecollege), "", rec!datecollege)
        txtcitizen.Text = IIf(IsNull(rec!citizen), "", rec!citizen)
        txtcollegecourse.Text = IIf(IsNull(rec!degree), "", rec!degree)
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
        txtnbi.Text = IIf(IsNull(rec!nbi), "", rec!nbi)
        txtrescer.Text = IIf(IsNull(rec!rescer), "", rec!rescer)
        txtissueat.Text = IIf(IsNull(rec!issueat), "", rec!issueat)
        Me.txtIssOn.Text = IIf(IsNull(rec!issueon), "", rec!issueon)
        mskdateprob.Text = IIf(IsNull(rec!dateprob), "__/__/____", Format(rec!dateprob, "mm/dd/yyyy"))
        MskDateHired.Text = IIf(IsNull(rec!datehired), "__/__/____", Format(rec!datehired, "mm/dd/yyyy"))
        mskdatereg.Text = IIf(IsNull(rec!dateregular), "__/__/____", Format(rec!dateregular, "mm/dd/yyyy"))
        mskdateres.Text = IIf(IsNull(rec!dateresign), "__/__/____", Format(rec!dateresign, "mm/dd/yyyy"))
        txtsupervisor.Text = IIf(IsNull(rec!supervisor), "", rec!supervisor)
        txtreasonresign.Text = IIf(IsNull(rec!reasonres), "", rec!reasonres)
        Me.mskendof.Text = IIf(IsNull(rec!endofcontract) = True Or rec!endofcontract = "", "__/__/____", Format(rec!endofcontract, "mm/dd/yyyy"))
    
    
       If rec!PaymentBasis = "Hourly" Then
        cmbPayBasis.ListIndex = 0
        txtRate = IIf(IsNull(rec!HourlyRate), "", Format(rec!HourlyRate, "##0.000"))
       ElseIf rec!PaymentBasis = "Daily" Then
        cmbPayBasis.ListIndex = 1
        txtRate = IIf(IsNull(rec!DailyRate), "", Format(rec!DailyRate, "##0.00"))
       ElseIf rec!PaymentBasis = "Monthly" Then
        cmbPayBasis.ListIndex = 2
        txtRate = IIf(IsNull(rec!MonthlyRate), "", Format(rec!MonthlyRate, "##0.00"))
       Else
        cmbPayBasis.ListIndex = -1
        txtRate = ""
       End If
       
       mskStartContract.Text = IIf(IsNull(rec!StartContract) = True Or rec!StartContract = "__/__/____", "__/__/____", Format(rec!StartContract, "mm/dd/yyyy"))
       mskEndContract.Text = IIf(IsNull(rec!endcontract) = True Or rec!endcontract = "__/__/____", "__/__/____", Format(rec!endcontract, "mm/dd/yyyy"))
       txtMonthsHired.Text = IIf(IsNull(rec!NoofMonthsHired) = True, "", rec!NoofMonthsHired)
       txtCBtotal.Text = IIf(IsNull(rec!Cbamount) = True, "", Format(rec!Cbamount, "##0.00"))
       txtCBpayment.Text = IIf(IsNull(rec!cbpayment) = True, "", Format(rec!cbpayment, "##0.00"))
       mskStartCB.Text = IIf(IsNull(rec!cbstartpayment) = True Or rec!cbstartpayment = "__/__/____", "__/__/____", Format(rec!cbstartpayment, "mm/dd/yyyy"))
       mskEndCB.Text = IIf(IsNull(rec!cbendpayment) = True Or rec!cbendpayment = "__/__/____", "__/__/____", Format(rec!cbendpayment, "mm/dd/yyyy"))
       'txtMiscFee.Text = IIf(IsNull(rec!MiscFee), "", Format(rec!MiscFee, "##0.00"))
       'txtPlacementFee.Text = IIf(IsNull(rec!PlacementFee), "", Format(rec!PlacementFee, "##0.00"))
        TxtMFAmount = IIf(IsNull(rec!MiscFee), "0.00", Format(rec!MiscFee, "##0.00"))
        txtPFAmount.Text = IIf(IsNull(rec!PlacementFee), "0.00", Format(rec!PlacementFee, "##0.00"))
        
        TxtMFPayment = IIf(IsNull(rec!MFpayment), "0.00", Format(rec!MFpayment, "##0.00"))
        txtPFPayment = IIf(IsNull(rec!PFpayment), "0.00", Format(rec!PFpayment, "##0.00"))
                
        MskMFStartPayment.Text = IIf(IsNull(rec!MFstartpayment) = True Or rec!MFstartpayment = "__/__/____", "__/__/____", Format(rec!MFstartpayment, "mm/dd/yyyy"))
        MskMFEndPayment.Text = IIf(IsNull(rec!MFendpayment) = True Or rec!MFendpayment = "__/__/____", "__/__/____", Format(rec!MFendpayment, "mm/dd/yyyy"))
        
        MskPFStartPayment.Text = IIf(IsNull(rec!PFstartpayment) = True Or rec!PFstartpayment = "__/__/____", "__/__/____", Format(rec!PFstartpayment, "mm/dd/yyyy"))
        MskPFEndPayment.Text = IIf(IsNull(rec!PFendpayment) = True Or rec!PFendpayment = "__/__/____", "__/__/____", Format(rec!PFendpayment, "mm/dd/yyyy"))
        
        'txtrate.Text = IIf(IsNull(rec!ntFee), "", rec!PlacementFee)
     If rec!WithEcolaPay = 1 Or rec!WithEcolaPay = True Then
        ChkEcola.VALUE = 1
'        TxtEcolaAmount.Enabled = True
    Else
        ChkEcola.VALUE = 0
 '       TxtEcolaAmount.Enabled = False
    End If
        TxtEcolaAmount.Text = IIf(IsNull(rec!EcolaAmount) = True, "", Format(rec!EcolaAmount, "##0.00"))
    
    End If
     
    Set rs = cnDCPI.Execute("Select taxexemptioncode From taxexemptioncode Where taxexemptioncode ='" & IIf(IsNull(rec!civilstatus), "", rec!civilstatus) & "'")
        If Not rs.EOF Then
            DoEvents
            cmbcivilstatus.Text = rs.Fields(0) '& "-" & rs.Fields(1)
        Else
            cmbcivilstatus.ListIndex = -1
        End If
        rs.Close
    
     Set rs = cnDCPI.Execute("Select * From clienttable Where clientcode =" & IIf(IsNull(rec!ClientCode), 0, rec!ClientCode))
            If Not rs.EOF Then
                cmbCompany.Text = rs.Fields(1)
                lblCompany.Caption = UCase(rs.Fields(1))
                If SuperUser = False And rs.Fields("ClientFlag2") = True Then
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
            rs.Close
            
     Set rs = cnDCPI.Execute("Select * From HRBRANCH Where brcode =" & IIf(IsNull(rec!BrCode), 0, rec!BrCode))
            If Not rs.EOF Then
                cmbBranch.Text = rs.Fields(1)
            Else
                cmbBranch.ListIndex = -1
            End If
            rs.Close
            
      Set rs = cnDCPI.Execute("Select * From HRdept Where deptcode =" & IIf(IsNull(rec!DeptCode), 0, rec!DeptCode))
           If Not rs.EOF Then
                cmbDept.Text = rs.Fields(1)
            Else
                cmbDept.ListIndex = -1
            End If
      rs.Close
            
      Set rs1 = cnDCPI.Execute("Select * From HREmploymentstatus Where employmentstatus ='" & IIf(IsNull(rec!employmentstatus), "", rec!employmentstatus) & "'")
            If Not rs1.EOF Then
                cmbempstatus.Text = rs1.Fields(0) & "-" & rs1.Fields(1)
            Else
                cmbempstatus.ListIndex = -1
            End If
            rs1.Close
    
      Set rsdialect = cnDCPI.Execute("Select * from HRDialects where employeenum='" & TxtControlNo.Text & "'")
         If rsdialect.EOF = False Then
             Showdialect rsdialect
             rsdialect.MoveNext
         Else
             txtdialect1 = ""
             txtdialect2 = ""
             txtdialect3 = ""
         End If
        rsdialect.Close
        
      Set rsdetails = cnDCPI.Execute("Select * from HRDetails where employeenum='" & TxtControlNo.Text & "'")
        If rsdetails.EOF = False Then
            Showdetails rsdetails
            rsdetails.MoveNext
        Else
             txtspousename.Text = ""
             txtspouseaddress.Text = ""
             txtsjob.Text = ""
             txtnumdep.Text = ""
             txtdep1.Text = ""
             txtdep2.Text = ""
             txtdep3.Text = ""
             txtdep4.Text = ""
             mskdep1.Text = "__/__/____"
             mskdep2.Text = "__/__/____"
             mskdep3.Text = "__/__/____"
             mskdep4.Text = "__/__/____"
             mskdep5.Text = "__/__/____" ' IIf(IsNull(Rec2!dep5bdate), "__/__/____", Format(Rec2!dep5bdate, "mm/dd/yyyy"))
        End If
        rsdetails.Close
    
      Set rsref = cnDCPI.Execute("Select * from HRCharReference where employeenum='" & TxtControlNo.Text & "'")
         If rsref.EOF = False Then
            ShowRef rsref
            rsref.MoveNext
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
        rsref.Close
        Set rsemprec = cnDCPI.Execute("Select * from HREmploymentRecord where employeenum='" & TxtControlNo.Text & "'")
        If rsemprec.EOF = False Then
            ShowEmpRec rsemprec
            rsemprec.MoveNext
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
        rsemprec.Close
        
        Set rssem = cnDCPI.Execute("Select * from HRSeminar where employeenum='" & TxtControlNo.Text & "'")
        If rssem.EOF = False Then
            ShowSem rssem
        Else
            txtseminar1.Text = ""
            txtdateseminar1.Text = ""
            txtseminar2.Text = ""
            txtdateseminar2.Text = ""
            txtseminar3.Text = ""
            txtdateseminar3.Text = ""
        End If
       
        rssem.Close
End Sub

Private Sub Edit()
    cmdnew.Enabled = False
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
    btnEmployeeSearch.Enabled = False
End Sub

Private Sub cmbcivil_GotFocus()
    cmbcivil.BackColor = &HFFFFC0
    Set oCommonFunction.ValidAlpha = cmbcivil
End Sub

Private Sub cmbcivil_LostFocus()
    cmbcivil.BackColor = vbWhite
End Sub

Public Sub btnEmpRefresh_Click()
On Error Resume Next
'   Dim rsout As New ADODB.Recordset
    If rsout.State = 1 Then rsout.Close
    rsout.Open "Select * from HREmployees order by surname, firstname", cnDCPI, adOpenKeyset, adLockOptimistic
         If rsout.EOF = False Then
             ' rsout.MoveFirst
              Showdata rsout
              cmdNext.Enabled = True
              cmdlast.Enabled = True
         Else
             MsgBox "All records deleted, the form will be close.", vbInformation
             Unload Me
             Exit Sub
         End If
btnEmpRefresh.Enabled = True
btnEmployeeSearch.Enabled = True
End Sub

Private Sub cmbbranch_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbcivilstatus_GotFocus()
cmbcivilstatus.BackColor = &HFFFFC0
End Sub

Private Sub cmbcivilstatus_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbcivilstatus_LostFocus()
cmbcivilstatus.BackColor = vbWhite
End Sub

Private Sub cmbCompany_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbdept_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
Private Sub cmbempstatus_GotFocus()
    cmbempstatus.BackColor = &HFFFFC0
End Sub

Private Sub cmbempstatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmbempstatus_LostFocus()
    cmbempstatus.BackColor = vbWhite
End Sub
Private Sub cmbPayBasis_Click()
    If cmbPayBasis.ListIndex = 0 Then lblRAte.Caption = "Hourly rate"
    If cmbPayBasis.ListIndex = 1 Then lblRAte.Caption = "Daily rate"
    If cmbPayBasis.ListIndex = 2 Then lblRAte.Caption = "Monthly rate"
    
  '  Dim rsDelete As ADODB.Recordset
   ' Set rsDelete = New ADODB.Recordset
   ' rsDelete.Open "Select HOURLYRATE,DAILYRATE,MONTHLYRATE from HREMPLOYEES where employeenum=" & GetEmployeenum() & "", cnDCPI, adOpenKeyset, adLockOptimistic
    'If rsDelete.EOF = False Then
   If addflagship <> True Then
         If cmbPayBasis.ListIndex = 0 Then txtRate.Text = Format(rsout("hourlyrate"), "#00.000")
        If cmbPayBasis.ListIndex = 1 Then txtRate.Text = Format(rsout("dailyrate"), "#00.00")
        If cmbPayBasis.ListIndex = 2 Then txtRate.Text = Format(rsout("monthlyrate"), "#00.00")
    'Else
    'End If
    End If
If txtRate.Enabled = True Then txtRate.SetFocus
End Sub
Private Sub cmdAddBranch_Click()
    frmBranch.Show 1
End Sub

Private Sub cmdAddDept_Click()
    frmDept.Show
End Sub

Private Sub cmdDelete_Click()
Dim oldControlNo As String
oldControlNo = TxtControlNo.Text
On Error Resume Next
    Dim res
    Dim rsDelete As ADODB.Recordset
    Set rsDelete = New ADODB.Recordset
    rsDelete.Open "Select * from employeebasicsalary where employeenum=" & Me.txtID.Text & "  and employeestatus=true", cnDCPI, adOpenKeyset, adLockOptimistic
    If rsDelete.EOF = False Then
        MsgBox "This employee is included in Payroll List, cannot be deleted. Mark it inactive in Employee Basic Salary maintenance first.", vbInformation, PayrollAlert
        Exit Sub
    Else
        Call DELEMP
        If mHRSEARCH = True Then
              Call frmHRSearch.lvwdisplay.ListItems.Clear
              Call frmHRSearch.chkStatus_Click(1)
            '  Call frmHRSearch.Loadlist("")
            '  frmHRSearch.cmbClient.ListIndex = 0
        End If
        
        If oldControlNo = TxtControlNo.Text Then
        Call btnEmpRefresh_Click
        End If
      
    End If
    rsDelete.Close
    Set rsDelete = Nothing
End Sub

Private Function GetEmployeenum() As Integer
On Error GoTo err2
    Dim rsemp As ADODB.Recordset
    Set rsemp = New ADODB.Recordset
    rsemp.Open "Select employeenum from hremployees where controlno='" & Me.TxtControlNo.Text & "'", cnDCPI, adOpenKeyset, adLockOptimistic
    If rsemp.EOF = False Then
        GetEmployeenum = rsemp.Fields(0)
    Else
        GetEmployeenum = 0
    End If
    rsemp.Close
    Set rsemp = Nothing
    Exit Function
err2:
    Err.Clear
    GetEmployeenum = 0
    
    
End Function

Private Sub DELEMP()
On Error Resume Next
 Dim sqlString As String
  res = MsgBox("You are about to delete 1 record." & vbCrLf & " " & vbCrLf & "If you click Yes, You won't able to undo this delete operation. " & vbCrLf & "Are you sure you want to delete the Employee No: " & TxtControlNo & "?", vbQuestion + vbYesNo, CompanyName)
  If res = vbYes Then
     sqlString = "Delete from HREmployees where employeenum=" & Me.txtID.Text & "":        Call Action_Query(sqlString, Cnnstr)
    
    If Trim(TxtControlNo.Text) <> "" Then
    sqlString = "Delete from HRDetails where employeenum='" & TxtControlNo.Text & "'":       Call Action_Query(sqlString, Cnnstr)
    sqlString = "Delete from HRCharReference where employeenum='" & TxtControlNo.Text & "'":         Call Action_Query(sqlString, Cnnstr)
    sqlString = "Delete from HRDialects where employeenum='" & TxtControlNo.Text & "'":          Call Action_Query(sqlString, Cnnstr)
    sqlString = "Delete from HREmploymentRecord where employeenum='" & TxtControlNo.Text & "'":      Call Action_Query(sqlString, Cnnstr)
    sqlString = "Delete from HRSeminar where employeenum='" & TxtControlNo.Text & "'":         Call Action_Query(sqlString, Cnnstr)
    End If
    
    sqlString = "Delete from HREmployeesAssignment where employeenum='" & GetEmployeenum() & "'":        Call Action_Query(sqlString, Cnnstr)
    sqlString = "Delete from HREmployeesLeaves where employeenum='" & GetEmployeenum() & "'":         Call Action_Query(sqlString, Cnnstr)
    sqlString = "Delete from HREmployeesOffenses where employeenum='" & GetEmployeenum() & "'":       Call Action_Query(sqlString, Cnnstr)
    sqlString = "Delete from HREmployeesTraining where employeenum='" & GetEmployeenum() & "'":      Call Action_Query(sqlString, Cnnstr)
    sqlString = "Delete from HREmployeeBenefits where employeenum='" & GetEmployeenum() & "'":       Call Action_Query(sqlString, Cnnstr)
    sqlString = "Delete from HREmployeesAppraisal where employeenum='" & GetEmployeenum() & "'":       Call Action_Query(sqlString, Cnnstr)
    'frmStatus.Show 1
    MsgBox "The record of [" & txtsurname.Text & ", " & txtfirstname.Text & "] was deleted.", vbInformation
    Call btnEmpRefresh_Click
End If
End Sub

Public Sub cmdEdit_Click()
    DoEvents
    TmpSurname = Me.txtsurname.Text
    TmpFirstname = Me.txtfirstname.Text
    TmpMiddlename = Me.txtmiddlename.Text
    
    prevControlNo = Me.TxtControlNo.Text
    PrevCompany = Me.cmbCompany.Text
    PrevStartContract = Me.mskStartContract.Text
    PrevEndContract = Me.mskEndContract.Text
    
    editflagship = True
    addflagship = False
    Call Edit
    Call TextEnabled(True)
    Call CheckUser1(operatorID)
End Sub

Private Sub cmdfirst_Click()
    rsout.MoveFirst
    cmdfirst.Enabled = False
    cmdPrev.Enabled = False
    cmdlast.Enabled = True
    cmdNext.Enabled = True
    Showdata rsout
End Sub

Private Sub cmdlast_Click()
    rsout.MoveLast
    cmdlast.Enabled = False
    cmdNext.Enabled = False
    cmdPrev.Enabled = True
    cmdfirst.Enabled = True
    Showdata rsout
End Sub

Private Sub cmdnew_Click()
    addflagship = True
    editflagship = False
    'LblUpdateBy.Caption = UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID)))
    Me.Caption = "EMPLOYEES INFORMATION   [Last update by: " + UCase(GETuSERNAME(IIf(IsNull(operatorID), "0", operatorID))) + "]"
    Call Add
    Call clearscreen
    'cmbPayBasis.ListIndex = 1
    TxtControlNo.Text = Format(Date, "yy") & "-" & Format(Date, "mm") & "-" & Format(NEWID + 1, "00000")
    Call TextEnabled(True)
End Sub

Private Sub cmdNext_Click()
    'On Error GoTo X
    'rsout.Requery
 rsout.MoveNext
 
 cmdPrev.Enabled = True
 cmdfirst.Enabled = True
   If rsout.EOF Then
        rsout.MoveLast
        cmdNext.Enabled = False
        cmdlast.Enabled = False
        MsgBox "End of file encountered", vbInformation, CompanyName
        Exit Sub
End If
Showdata rsout
End Sub
Private Sub refreshappno()
'On Error GoTo X
 rsout.MoveNext
 cmdPrev.Enabled = True
 cmdfirst.Enabled = True
   If rsout.EOF Then
        rsout.MoveLast
        cmdNext.Enabled = False
        cmdlast.Enabled = False
End If
    Showdata rsout
X:
    Exit Sub

End Sub

Private Sub cmdPrev_Click()
' On Error GoTo X
 rsout.MovePrevious
  cmdNext.Enabled = True
      cmdlast.Enabled = True
   If rsout.BOF Then
       rsout.MoveFirst
      cmdPrev.Enabled = False
      cmdfirst.Enabled = False
      MsgBox "Beginning of file encountered", vbInformation, CompanyName
    End If
    
   employeenum = rsout!employeenum
    Showdata rsout
X:
    Exit Sub
End Sub

Private Sub cmdsearch_Click()
    frmHRSearch.Show vbModal
End Sub

Private Sub cmdundo_Click()
   If AddFlag = True Then
        Unload Me
        Exit Sub
    End If

    addflagship = False
    editflagship = False
      Call TextEnabled(False)
    Showdata rsout
      btnEmpRefresh.Enabled = True
      btnEmployeeSearch.Enabled = True
    cmdundo.Enabled = False
    cmdupdate.Enabled = False
    cmdnew.Enabled = True
    cmddelete.Enabled = True
    cmdedit.Enabled = True
    cmdfirst.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdlast.Enabled = True
 cmdAddBranch.Enabled = False
 cmdAddDept.Enabled = False
End Sub
Public Sub Cancel()
    Call TextEnabled(False)
End Sub

Private Sub TextEnabled(ByVal vlfalse As Boolean)
    cmdAddBranch.Enabled = vlfalse
    cmdAddDept.Enabled = vlfalse
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
    cmbCompany.Enabled = vlfalse
    cmbDept.Enabled = vlfalse
    cmbBranch.Enabled = vlfalse
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
    cmbempstatus.Enabled = vlfalse
    txtsupervisor.Enabled = vlfalse
    txtreasonresign.Enabled = vlfalse
    mskendof.Enabled = vlfalse
    cmbPayBasis.Enabled = vlfalse
    txtRate.Enabled = vlfalse
    mskStartContract.Enabled = vlfalse
    mskEndContract.Enabled = vlfalse
    txtMonthsHired.Enabled = vlfalse
    TxtEcolaAmount.Enabled = vlfalse
    txtCBtotal.Enabled = vlfalse
    txtCBpayment.Enabled = vlfalse
    mskStartCB.Enabled = vlfalse
    mskEndCB.Enabled = vlfalse
    TxtMFAmount.Enabled = vlfalse
    txtPFAmount.Enabled = vlfalse
    
    TxtMFPayment.Enabled = vlfalse
    txtPFPayment.Enabled = vlfalse
    
    MskMFStartPayment.Enabled = vlfalse
    MskMFEndPayment.Enabled = vlfalse
    MskPFStartPayment.Enabled = vlfalse
    MskPFEndPayment.Enabled = vlfalse
    Frame11.Enabled = vlfalse
    Frame9.Enabled = vlfalse
    ChkEcola.Enabled = vlfalse
End Sub

Public Function ValidateTextbox() As Boolean

    If TxtControlNo.Text = "" Then MsgBox "Employee number is required", vbInformation: TxtControlNo.SetFocus: GoTo last
    If txtsurname.Text = "" Then MsgBox "Surname is required", vbInformation: txtsurname.SetFocus: GoTo last
    If txtfirstname.Text = "" Then MsgBox "Firstname is required", vbInformation: txtfirstname.SetFocus: GoTo last
    'If txtmiddlename.Text = "" Then MsgBox "Middle name/Middle initial is required", vbInformation: txtmiddlename.SetFocus: GoTo last
    If cmbCompany.Text = "" Then MsgBox "Pls select company for this employee.", vbInformation: SSTab1.Tab = 1: cmbCompany.SetFocus: GoTo last
    
    If Me.cmbPayBasis.ListIndex = -1 Then MsgBox "Pls select Payment basis for this employee.", vbInformation: SSTab1.Tab = 1: Me.cmbPayBasis.SetFocus: GoTo last
         If Me.txtRate.Text = 0 Or Me.txtRate.Text = "" Then MsgBox "Pls enter rate for this employee.", vbInformation: SSTab1.Tab = 1: Me.txtRate.SetFocus: GoTo last
     
      If Me.cmbcivilstatus.ListIndex = -1 Then MsgBox "Pls select civil status for this employee.", vbInformation: SSTab1.Tab = 0: Me.cmbcivilstatus.SetFocus: GoTo last
  
     Dim date1 As Date
     Dim date2 As Date
     
    If MskPFStartPayment <> "__/__/____" And IsDate(MskPFStartPayment) = True And MskPFEndPayment <> "__/__/____" And IsDate(MskPFEndPayment) = True Then
     date1 = MskPFStartPayment.Text
     date2 = Me.MskPFEndPayment.Text
        
        If date1 > date2 Then
             MsgBox "PF end date must not be earlier than start date.", vbInformation
             MskPFEndPayment.SetFocus
             GoTo last
        End If
     End If
    
        If MskMFStartPayment <> "__/__/____" And IsDate(MskMFStartPayment) = True And MskMFEndPayment <> "__/__/____" And IsDate(MskMFEndPayment) = True Then
     date1 = MskMFStartPayment.Text
     date2 = MskMFEndPayment.Text
     
        If date1 > date2 Then
             MsgBox "MF end date must not be earlier than start date.", vbInformation
             MskMFEndPayment.SetFocus
             GoTo last
        End If
     End If


  If mskStartCB <> "__/__/____" And IsDate(mskStartCB) = True And mskEndCB <> "__/__/____" And IsDate(mskEndCB) = True Then
     date1 = mskStartCB.Text
     date2 = mskEndCB.Text
        
        If date1 > date2 Then  ' should be less than
             MsgBox "CB end date must not be earlier than start date.", vbInformation
             mskEndCB.SetFocus
             GoTo last
        End If
     End If
   
    If RTrim(txtRate.Text) = "" Then txtRate = "0.000"
    ValidateTextbox = True
    Exit Function

last:
    
    ValidateTextbox = False

End Function

Private Sub SaveHistory(ByVal blUpdate As Boolean)

Dim cnnStrJCI As String
Dim cnnStrSAPI As String
   
   On Error GoTo bb:
  
  'SAPI
        StrDbpath = GetSetting("DataSearch", "SE", "Path2")
        CName = "SAPI Manpower & Mgt. Services"
        If StrDbpath = "" Then MsgBox "Cant save this data to history,unable to connect to " & CName & ". connection not established": Exit Sub
        cnnStrSAPI = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrDbpath

  'JCI
        CName = "Jobs and Careers Inc."
        StrDbpath = GetSetting("DataSearch", "JCI", "Path2")
        If StrDbpath = "" Then MsgBox "Cant save this data to history,unable to connect to " & CName & ". connection not established": Exit Sub
         cnnStrJCI = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrDbpath

    intretval = OpenDatabaseAdo(cnn, cnnStrSAPI)
    intretval1 = OpenDatabaseAdo(cnn, cnnStrJCI)
    If intretval <> 0 Or intretval1 <> 0 Then
        'MsgBox "Unable To Establish the Connection. ", vbCritical
    Exit Sub
    Else
    
            strSQL = "Select * from hremploymenthistory where controlno ='" & RTrim(TxtControlNo.Text) & "'"
            Dim rsemp As New ADODB.Recordset
            Set rsemp = ReturnRst(strSQL, cnnStrJCI, mstatic)
             Dim rsEmp1 As New ADODB.Recordset
            Set rsEmp1 = ReturnRst(strSQL, cnnStrSAPI, mstatic)
            
            If rsemp.EOF = False Or rsEmp1.EOF = False Then
            'already saved
               MsgBox "Employee number [" & Me.TxtControlNo.Text & "]" & " is ALREADY SAVED  in history.", vbInformation
                Exit Sub
            End If
    End If
    

    ' If PrevControlno <> Me.TxtControlNo.Text And PrevCompany <> Me.cmbCompany.Text And PrevStartContract <> Me.mskStartContract.Text And PrevEndContract <> Me.mskEndContract.Text Then
           Call ComputeValue
            'sqlString = "Insert into hremploymenthistory (employeenum,ControlNo ,Surname,middlename, FirstName,ClientCODE,ClientName, Empposition,employmentstatus, StartContract, EndContract, PaymentBasis,HourlyRate,DailyRate, MonthlyRate, deptcode,deptname,brcode,brname, [USERID]) " & _
            " values (" & txtID.Text & ",'" & TxtControlNo.Text & "'," & ChkVal(txtsurname) & "," & ChkVal(txtmiddlename) & "," & ChkVal(txtfirstname) & "," & Cid & ",'" & cmbCompany.Text & "'," & ChkVal(txtfirstchoice) & "," & ChkCmb(cmbempstatus) & "," & ChkMsk(mskStartContract) & "," & ChkMsk(mskEndContract) & "" & _
            "," & ChkCmb(cmbPayBasis) & "," & HourlyRate & "," & DailyRate & "," & MonthlyRate & "," & DeptCode & "," & ChkCmb(cmbDept) & "," & BrCode & "," & ChkCmb(cmbBranch) & "" & _
            "," & operatorID & ")"
            
     sqlString = "Insert into hremploymenthistory (employeenum,COMPANYID,ControlNo ,Surname,middlename, FirstName,ClientCODE,ClientName, Empposition,employmentstatus, StartContract, EndContract, PaymentBasis,HourlyRate,DailyRate, MonthlyRate, deptcode,deptname,brcode,brname, [USERID]) " & _
                    " values (" & txtID.Text & ",'" & CompanyCode & "','" & TxtControlNo.Text & "'," & ChkVal(txtsurname) & "," & ChkVal(txtmiddlename) & "," & ChkVal(txtfirstname) & "," & Cid & ",'" & cmbCompany.Text & "'," & ChkVal(txtfirstchoice) & "," & ChkCmb(cmbempstatus) & "," & ChkMsk(mskStartContract) & "," & ChkMsk(mskEndContract) & "" & _
                    "," & ChkCmb(cmbPayBasis) & "," & HourlyRate & "," & DailyRate & "," & MonthlyRate & "," & DeptCode & "," & ChkCmb(cmbDept) & "," & BrCode & "," & ChkCmb(cmbBranch) & "" & _
                    "," & operatorID & ")"
                    
      If (blUpdate = False) Then
      
      Else
                        
                      sqlString = "Update hremployees set ControlNo='" & TxtControlNo.Text & "'" & _
       ", FirstName=" & ChkVal(txtfirstname) & ", Empposition =" & ChkVal(txtfirstchoice) & ", age=" & IIf(txtage.Text = "", 0, txtage.Text) & ", middlename=" & ChkVal(txtmiddlename) & ", Surname=" & ChkVal(txtsurname) & ", nickname=" & ChkVal(txtnickname) & ", BIRTHDATE=" & ChkMsk(mskbday) & ", Address1=" & ChkVal(txtaddress1) & ", houseno=" & ChkVal(txthouseno) & ", street=" & ChkVal(txtstreet) & ", barangay=" & ChkVal(txtbarangay) & ",district=" & ChkVal(txtdistrict) & ", city=" & ChkVal(txtcity) & ",province1=" & ChkVal(txtprovince) & ", province=" & ChkVal(txtprovinceaddress) & ", contact=" & ChkVal(txtemergency) & ", caddress=" & ChkVal(txteraddress) & ",cphone=" & ChkVal(txtertel) & ",gender='" & IIf(optmale.VALUE, "M", "F") & "',Height='" & Replace(txtheight, "'", """") & "',Weight=" & ChkVal(txtweight) & "" & _
       " ,Telnum1=" & ChkVal(txttel1) & ",Telnum2=" & ChkVal(txttel2) & ",civilstatus='" & IIf(cmbcivilstatus.Text = "", "S", cmbcivilstatus.Text) & "',examgrade=" & ChkVal(txtexam) & ",interviewgrade=" & ChkVal(txtinterview) & ",medicalgrade=" & ChkVal(txtmedical) & ",sss=" & ChkMsk(msksss) & ",philhealth=" & ChkVal(txtphilhealth) & ", pagibig=" & ChkVal(txtpagibig) & ",AcctNo=" & ChkVal(txtacctno) & "" & _
       ", citizen=" & ChkVal(txtcitizen) & ", birthplace=" & ChkVal(txtbirthplace) & ",haircolor=" & ChkVal(txthaircolor) & ", eyecolor=" & ChkVal(txteyecolor) & ", Religion=" & ChkVal(txtreligion) & ", taxid=" & ChkMsk(msktax) & ", [Primary]=" & ChkVal(txtprimarysch) & ",dateprimary=" & ChkVal(txtdateprimary) & ", secondary=" & ChkVal(txtsecsch) & ", datesecondary=" & ChkVal(txtdatesec) & ",college=" & ChkVal(txtcollegesch) & ", datecollege=" & ChkVal(txtdatecollege) & ", degree=" & ChkVal(txtcollegecourse) & ", Major=" & ChkVal(txtcollegemajor) & ", postgraduate=" & ChkVal(txtpostsch) & "," & _
       " datepostgraduate= " & ChkVal(txtdatepost) & ", postgraduatecourse=" & ChkVal(txtpostcourse) & " , fathername=" & ChkVal(txtfathername) & ", faddress=" & ChkVal(txtfaddress) & ", fwork=" & ChkVal(txtfjob) & ", mothername=" & ChkVal(txtmothername) & ", maddress=" & ChkVal(txtmaddress) & ", mwork=" & ChkVal(txtmjob) & ", nbi=" & ChkVal(txtnbi) & ", rescer=" & ChkVal(txtrescer) & ", issueat=" & ChkVal(txtissueat) & ", issueon=" & ChkMsk(mskissue) & ", datehired=" & ChkMsk(MskDateHired) & ", dateprob=" & ChkMsk(mskdateprob) & ", dateregular=" & ChkMsk(mskdatereg) & ", dateresign=" & ChkMsk(mskdateres) & ", employmentstatus='" & Left(cmbempstatus.Text, 1) & "'" & _
       ", supervisor=" & ChkVal(txtsupervisor) & ", reasonres=" & ChkVal(txtreasonresign) & ", endofcontract=" & ChkMsk(mskendof) & ", ClientCode=" & Cid & ", brcode=" & BrCode & "" & _
       ", PaymentBasis='" & PaymentBasis & "',HourlyRate=" & HourlyRate & ",DailyRate=" & DailyRate & ", MonthlyRate=" & MonthlyRate & ", StartContract=" & ChkMsk(mskStartContract) & ", EndContract=" & ChkMsk(mskEndContract) & ", NoofMonthsHired=" & IIf(txtMonthsHired.Text = "", 5, txtMonthsHired.Text) & ", cbamount=" & IIf(txtCBtotal = "", 0, Format(txtCBtotal, "##0.00")) & ", cbpayment=" & IIf(txtCBpayment = "", 0, Format(txtCBpayment, "##0.00")) & ",cbstartpayment=" & ChkMsk(mskStartCB) & ", cbendpayment=" & ChkMsk(mskEndCB) & "" & _
       ", MiscFee=" & IIf(TxtMFAmount.Text = "", 0, Format(TxtMFAmount, "##0.00")) & ", MFPayment=" & IIf(TxtMFPayment = "", 0, Format(TxtMFPayment, "##0.00")) & " , MFstartpayment=" & ChkMsk(MskMFStartPayment) & ", MFendpayment=" & ChkMsk(MskMFEndPayment) & "" & _
       ", PlacementFee=" & IIf(txtPFAmount = "", 0, Format(txtPFAmount, "##0.00")) & ", PFPayment=" & IIf(txtPFPayment = "", 0, Format(txtPFPayment, "##0.00")) & ", PFSTARTPAYMENT=" & ChkMsk(MskPFStartPayment) & ", PFendpayment=" & ChkMsk(MskPFEndPayment) & "" & _
       ", WithEcolaPay='" & WithEcolaPay & "',  ecolaamount=" & EcolaAmount & ", [USERID]='" & operatorID & "', deptcode=" & DeptCode & ", remarks='" & IIf(Me.txtreasonresign.Text = "", " ", txtreasonresign.Text) & "', LastUpdate = '" & Format(Now, "mm/dd/yyyy") & "'" & _
      " where firstname ='" & RTrim(txtfirstname.Text) & "'  and surname= '" & RTrim(txtsurname.Text) & "'"
  
       




    End If
    
                    
                Call Action_Query(sqlString, cnnStrJCI)
              Call Action_Query(sqlString, cnnStrSAPI)
              MsgBox "Record succesfuly saved to history table.", vbInformation
       ' End If
bb:

End Sub
Private Sub cmdupdate_Click()
        If ValidateTextbox = False Then Exit Sub
        If AddFlag = True Or addflagship = True Then 'SAVE FOR NEW EMPLOYEES
            If saveclick = False Then Exit Sub
            Call SetButtonsave
            Call TextEnabled(False)
            AddFlag = False
            addflagship = False
            EditFlag = False
        End If
        
        If EditFlag = True Or editflagship = True Then 'UPDATE EMPLOYEES
            If UpdateClick = False Then Exit Sub
            'Call SaveHistory
            MsgBox "[" & txtsurname.Text & ", " & txtfirstname.Text & "] was successfully updated!", vbInformation
            Call Edit
            Call TextEnabled(False)
        End If
        'Call btnEmpRefresh_Click
        Call SetButtonsave
        addflagship = False
        editflagship = False
        If mHRSEARCH = True Then 'REFRESH HR SEARCH FORM
        If TmpMiddlename <> txtmiddlename.Text Or TmpSurname <> Me.txtsurname.Text Or TmpFirstname <> Me.txtfirstname.Text Then Call frmHRSearch.lvwdisplay.ListItems.Clear
        Call frmHRSearch.chkStatus_Click(1)
        End If
End Sub

Private Function UpdateClick() As Boolean

Dim rsoutEdit As ADODB.Recordset
Set rsoutEdit = New ADODB.Recordset
rsoutEdit.Open "Select * from HREmployees where controlno='" & TxtControlNo & "' and employeenum<>" & txtID.Text & " ", cnDCPI, adOpenKeyset, adLockOptimistic
With rsoutEdit
    If .EOF = False Then
    MsgBox "CANNOT UPDATE because employee no. is already used by [" & rsoutEdit!Surname & ", " & rsoutEdit!FirstName & "]" & vbCrLf & "" & vbCrLf & "Please enter unique employee no. or change the existing.", vbInformation
    TxtControlNo.SetFocus
    UpdateClick = False
    Exit Function
    End If
End With


If IsContractSave = True Then Exit Function
  Call ComputeValue
  
'start 'check data in other database
If IsContractSave = True Then Exit Function
    
'end
  
  sqlString = "Update hremployees set ControlNo='" & TxtControlNo.Text & "'" & _
       ", FirstName=" & ChkVal(txtfirstname) & ", Empposition =" & ChkVal(txtfirstchoice) & ", age=" & IIf(txtage.Text = "", 0, txtage.Text) & ", middlename=" & ChkVal(txtmiddlename) & ", Surname=" & ChkVal(txtsurname) & ", nickname=" & ChkVal(txtnickname) & ", BIRTHDATE=" & ChkMsk(mskbday) & ", Address1=" & ChkVal(txtaddress1) & ", houseno=" & ChkVal(txthouseno) & ", street=" & ChkVal(txtstreet) & ", barangay=" & ChkVal(txtbarangay) & ",district=" & ChkVal(txtdistrict) & ", city=" & ChkVal(txtcity) & ",province1=" & ChkVal(txtprovince) & ", province=" & ChkVal(txtprovinceaddress) & ", contact=" & ChkVal(txtemergency) & ", caddress=" & ChkVal(txteraddress) & ",cphone=" & ChkVal(txtertel) & ",gender='" & IIf(optmale.VALUE, "M", "F") & "',Height='" & Replace(txtheight, "'", """") & "',Weight=" & ChkVal(txtweight) & "" & _
       " ,Telnum1=" & ChkVal(txttel1) & ",Telnum2=" & ChkVal(txttel2) & ",civilstatus='" & IIf(cmbcivilstatus.Text = "", "S", cmbcivilstatus.Text) & "',examgrade=" & ChkVal(txtexam) & ",interviewgrade=" & ChkVal(txtinterview) & ",medicalgrade=" & ChkVal(txtmedical) & ",sss=" & ChkMsk(msksss) & ",philhealth=" & ChkVal(txtphilhealth) & ", pagibig=" & ChkVal(txtpagibig) & ",AcctNo=" & ChkVal(txtacctno) & "" & _
       ", citizen=" & ChkVal(txtcitizen) & ", birthplace=" & ChkVal(txtbirthplace) & ",haircolor=" & ChkVal(txthaircolor) & ", eyecolor=" & ChkVal(txteyecolor) & ", Religion=" & ChkVal(txtreligion) & ", taxid=" & ChkMsk(msktax) & ", [Primary]=" & ChkVal(txtprimarysch) & ",dateprimary=" & ChkVal(txtdateprimary) & ", secondary=" & ChkVal(txtsecsch) & ", datesecondary=" & ChkVal(txtdatesec) & ",college=" & ChkVal(txtcollegesch) & ", datecollege=" & ChkVal(txtdatecollege) & ", degree=" & ChkVal(txtcollegecourse) & ", Major=" & ChkVal(txtcollegemajor) & ", postgraduate=" & ChkVal(txtpostsch) & "," & _
       " datepostgraduate= " & ChkVal(txtdatepost) & ", postgraduatecourse=" & ChkVal(txtpostcourse) & " , fathername=" & ChkVal(txtfathername) & ", faddress=" & ChkVal(txtfaddress) & ", fwork=" & ChkVal(txtfjob) & ", mothername=" & ChkVal(txtmothername) & ", maddress=" & ChkVal(txtmaddress) & ", mwork=" & ChkVal(txtmjob) & ", nbi=" & ChkVal(txtnbi) & ", rescer=" & ChkVal(txtrescer) & ", issueat=" & ChkVal(txtissueat) & ", issueon=" & ChkMsk(mskissue) & ", datehired=" & ChkMsk(MskDateHired) & ", dateprob=" & ChkMsk(mskdateprob) & ", dateregular=" & ChkMsk(mskdatereg) & ", dateresign=" & ChkMsk(mskdateres) & ", employmentstatus='" & Left(cmbempstatus.Text, 1) & "'" & _
       ", supervisor=" & ChkVal(txtsupervisor) & ", reasonres=" & ChkVal(txtreasonresign) & ", endofcontract=" & ChkMsk(mskendof) & ", ClientCode=" & Cid & ", brcode=" & BrCode & "" & _
       ", PaymentBasis='" & PaymentBasis & "',HourlyRate=" & HourlyRate & ",DailyRate=" & DailyRate & ", MonthlyRate=" & MonthlyRate & ", StartContract=" & ChkMsk(mskStartContract) & ", EndContract=" & ChkMsk(mskEndContract) & ", NoofMonthsHired=" & IIf(txtMonthsHired.Text = "", 5, txtMonthsHired.Text) & ", cbamount=" & IIf(txtCBtotal = "", 0, Format(txtCBtotal, "##0.00")) & ", cbpayment=" & IIf(txtCBpayment = "", 0, Format(txtCBpayment, "##0.00")) & ",cbstartpayment=" & ChkMsk(mskStartCB) & ", cbendpayment=" & ChkMsk(mskEndCB) & "" & _
       ", MiscFee=" & IIf(TxtMFAmount.Text = "", 0, Format(TxtMFAmount, "##0.00")) & ", MFPayment=" & IIf(TxtMFPayment = "", 0, Format(TxtMFPayment, "##0.00")) & " , MFstartpayment=" & ChkMsk(MskMFStartPayment) & ", MFendpayment=" & ChkMsk(MskMFEndPayment) & "" & _
       ", PlacementFee=" & IIf(txtPFAmount = "", 0, Format(txtPFAmount, "##0.00")) & ", PFPayment=" & IIf(txtPFPayment = "", 0, Format(txtPFPayment, "##0.00")) & ", PFSTARTPAYMENT=" & ChkMsk(MskPFStartPayment) & ", PFendpayment=" & ChkMsk(MskPFEndPayment) & "" & _
       ", WithEcolaPay='" & WithEcolaPay & "',  ecolaamount=" & EcolaAmount & ", [USERID]='" & operatorID & "', deptcode=" & DeptCode & ", remarks='" & IIf(Me.txtreasonresign.Text = "", " ", txtreasonresign.Text) & "', LastUpdate = '" & Format(Now, "mm/dd/yyyy") & "' where  employeenum = " & txtID.Text & " "
       
On Error GoTo endSave
       Call Action_Query(sqlString, Cnnstr)
       
         'frmStatus.Show 1
         
'       If Left(cmbempstatus.Text, 1) <> "B" And Left(cmbempstatus.Text, 1) <> "P" And Left(cmbempstatus.Text, 1) <> "C" Then
'        MarkAsInActive
'       End If
       
  Call editdetails
  Call editemprec
  Call editdialect
  Call editChar
  Call editsem

'check employee basic salary
'Stop
 
   sqlString = "Update employeebasicsalary set controlno= '" & TxtControlNo.Text & "' where employeenum=" & txtID.Text & " "
   Call Action_Query(sqlString, Cnnstr)
   
   If Left(cmbempstatus.Text, 1) <> "B" And Left(cmbempstatus.Text, 1) <> "P" And Left(cmbempstatus.Text, 1) <> "C" And Left(cmbempstatus.Text, 1) <> "R" Then
    MarkAsInActive
   End If

UpdateClick = True
rsout.Requery
rsout.Find " employeenum = " & txtID.Text & ""

'Showdata rsout
Exit Function
endSave:
MsgBox "Encountered Exception in updating. The data you want to save might be invalid. Please review your entry.", vbInformation
UpdateClick = False
End Function

Private Function IsContractSave() As Boolean
Dim CompanyName2 As String
    If CompanyCode = 2 Then
        StrDbpath = GetSetting("DataSearch", "SE", "Path2")
        CompanyName2 = "SAPI Manpower & Mgt. Services"
    ElseIf CompanyCode = 1 Then
        CompanyName2 = "Jobs and Careers Inc."
        StrDbpath = GetSetting("DataSearch", "JCI", "Path2")
    End If
 
    If StrDbpath <> "" And mskStartContract.Text <> "__/__/____" And mskEndContract.Text <> "__/__/____" Then
        CnnstrCopy = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrDbpath
        '   CnnstrCopy = "Driver={Microsoft Access Driver (*.mdb)};" & _
                        "Dbq=" & StrDbpath & ";" & _
                        "Uid=;Pwd=alpha;"
        intretval = OpenDatabaseAdo(cnn, Cnnstr)
        If intretval <> 0 Then
            'MsgBox "Unable To Establish the Connection. ", vbCritical
            'Exit Sub
        Else
            strSQL = "Select * from HREmployees where controlno='" & TxtControlNo & "' and startcontract=#" & Me.mskStartContract.Text & "# and endcontract=#" & Me.mskEndContract.Text & "# and EmploymentStatus ='" & Mid(cmbempstatus.Text, 1, 1) & "'"
            Dim rsemp As New ADODB.Recordset
            Set rsemp = ReturnRst(strSQL, CnnstrCopy, mstatic)
            If rsemp.EOF = False Then
                MsgBox "CANNOT SAVE." & vbCrLf & " " & vbCrLf & " Same Employee No. and employee status with same contract found  in " & CompanyName2 & " database. Please update the following: Employees status, contract and employee no. ", vbInformation
              IsContractSave = True
              Exit Function
            End If
        End If
    End If
        
    IsContractSave = False
End Function

Private Sub refreshB()
     If rsout.State = 1 Then rsout.Close
    Set rsout = New ADODB.Recordset
    rsout.Open "Select * from HREmployees order by  surname,firstname", cnDCPI, adOpenKeyset, adLockOptimistic

End Sub
Private Sub editsem()
Dim rssem As ADODB.Recordset
Set rssem = New ADODB.Recordset
rssem.Open "Select * from HRSeminar where employeenum='" & TxtControlNo.Text & "'", cnDCPI, adOpenKeyset, adLockOptimistic

With rssem
If .EOF = True Then
    .AddNew
    !employeenum = TxtControlNo.Text
End If
       !seminar1 = txtseminar1.Text
       !seminar1date = txtdateseminar1.Text
       !seminar2 = txtseminar2.Text
       !seminar2date = txtdateseminar2.Text
       !seminar3 = txtseminar3.Text
       !seminar3date = txtdateseminar3.Text
     .Update
End With
rssem.Close
Set rssem = Nothing

End Sub
Private Sub editemprec()
'On Error Resume Next
Dim rsemprec As ADODB.Recordset
Set rsemprec = New ADODB.Recordset
rsemprec.Open "Select * from HREmploymentRecord where employeenum = '" & TxtControlNo.Text & "'", cnDCPI, adOpenKeyset, adLockOptimistic
With rsemprec
If .EOF = True Then
    .AddNew
    !employeenum = TxtControlNo.Text
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
rsemprec.Close
Set rsemprec = Nothing

End Sub

Private Sub editdialect()
Dim rsdialect As ADODB.Recordset
Set rsdialect = New ADODB.Recordset
rsdialect.Open "Select * from HRDialects where employeenum='" & TxtControlNo.Text & "'", cnDCPI, adOpenKeyset, adLockOptimistic

With rsdialect
If .EOF = True Then
    .AddNew
    !employeenum = TxtControlNo.Text
End If

       !dialect1 = txtdialect1.Text
       !dialect2 = txtdialect2.Text
       !dialect3 = txtdialect3.Text
    .Update
End With

rsdialect.Close
Set rsdialect = Nothing

End Sub
Private Sub editdetails()
Dim rsdetails As ADODB.Recordset
Set rsdetails = New ADODB.Recordset
rsdetails.Open "Select * from HRDetails where employeenum='" & TxtControlNo.Text & "'", cnDCPI, adOpenKeyset, adLockOptimistic
With rsdetails
If .EOF = True Then
    .AddNew
    !employeenum = TxtControlNo.Text
End If
       !spousename = IIf(txtspousename.Text <> "", txtspousename.Text, "")
       !spouseaddress = txtspouseaddress.Text
       !spousework = txtsjob.Text
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
rsdetails.Close
Set rsdetails = Nothing
End Sub
Private Sub editChar()
Dim rsref As ADODB.Recordset
Set rsref = New ADODB.Recordset
rsref.Open "Select * from HRCharReference where employeenum='" & TxtControlNo.Text & "'", cnDCPI, adOpenKeyset, adLockOptimistic

With rsref
If .EOF = True Then
    .AddNew
    !employeenum = TxtControlNo.Text
End If
       
       !Name = txtrefname1.Text
       !occupation = txtrefjob1.Text
       !Address = txtrefaddress1.Text
       !telnum = txtrefphone1.Text
       !Relation = txtrefrel1.Text
       !Name1 = txtrefname2.Text
       !occupation1 = txtrefjob2.Text
       !Address1 = txtrefaddress2.Text
       !Telnum1 = txtrefphone2.Text
       !Relation1 = txtrefrel2.Text
       !Name2 = txtrefname3.Text
       !occupation2 = txtrefjob3.Text
       !Address2 = txtrefaddress3.Text
       !Telnum2 = txtrefphone3.Text
       !Relation2 = txtrefrel3.Text
    .Update

End With
rsref.Close
Set rsref = Nothing

End Sub

Private Function ChkVal(ByRef mTBox As TextBox) As String
Dim mAmt As String

If Trim(mTBox.Text) = "" Then
    mAmt = Trim(Replace(mTBox, "'", "''"))
    mAmt = "null"
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
   Cid = cmbCompany.ItemData(cmbCompany.ListIndex)
    If cmbBranch.Text = "" Then
       BrCode = 0
    Else
       BrCode = Me.cmbBranch.ItemData(Me.cmbBranch.ListIndex)
    End If
    PaymentBasis = IIf(cmbPayBasis.Text = "", " ", cmbPayBasis.Text)
    If cmbPayBasis.ListIndex = 0 Then 'hourly
            HourlyRate = Format(Val(txtRate.Text), "##0.000")
            DailyRate = 8 * txtRate.Text
            MonthlyRate = (8 * txtRate.Text) * getField("NoofDaysperMonth", Cid)
    ElseIf cmbPayBasis.ListIndex = 1 Then 'daily
            HourlyRate = txtRate.Text / 8
            DailyRate = Format(Val(txtRate.Text), "##0.00")
            MonthlyRate = txtRate.Text * getField("NoofDaysperMonth", Cid)
    ElseIf cmbPayBasis.ListIndex = 2 Then 'monthly
            MonthlyRate = Format(Val(txtRate.Text), "##0.00")
            DailyRate = txtRate.Text * 12 / getField("WorkingdaysperyEAR", Cid)
            HourlyRate = (txtRate.Text) * 12 / getField("WorkingdaysperyEAR", Cid) / 8
    End If
    
    If MonthlyRate = "" Then MonthlyRate = 0
    If DailyRate = "" Then DailyRate = 0
    If HourlyRate = "" Then HourlyRate = 0
        
    If ChkEcola.VALUE = 1 Then
        WithEcolaPay = 1
        EcolaAmount = IIf(TxtEcolaAmount.Text = "", "0.00", Format(TxtEcolaAmount.Text, "##0.00"))
    Else
        WithEcolaPay = 0
        EcolaAmount = 0
    End If
        
    If cmbDept.Text = "" Then
        DeptCode = 0
    Else
        DeptCode = Me.cmbDept.ItemData(Me.cmbDept.ListIndex)
    End If
End Function

Public Function MarkAsInActive()
On Error Resume Next
sqlString = "update employeebasicsalary set employeestatus=0 where employeenum=" & Me.txtID.Text & ""
Call Action_Query(sqlString, Cnnstr)
End Function

Public Function saveclick() As Boolean
Dim res
Set rsout = New ADODB.Recordset
rsout.Open "Select * from HREmployees where controlno='" & TxtControlNo & "'", cnDCPI, adOpenKeyset, adLockOptimistic

With rsout
    If .EOF = False Then
    MsgBox "CANNOT SAVE because employee no. is already used by [" & rsout!Surname & ", " & rsout!FirstName & "]" & vbCrLf & "" & vbCrLf & "Please enter unique employee no. or change the existing.", vbInformation
    TxtControlNo.SetFocus
    saveclick = False
    Exit Function
    End If
End With

 Call ComputeValue
  sqlString = "Insert into hremployees (ControlNo , FirstName, Empposition, age, middlename, Surname, nickname, BIRTHDATE, Address1, houseno, street, barangay,district, city,province1, province, contact, caddress,cphone,gender,Height,Weight,Telnum1,Telnum2,civilstatus,examgrade,interviewgrade,medicalgrade,sss,philhealth, pagibig,AcctNo, citizen, birthplace,haircolor, eyecolor, Religion, taxid, [Primary],dateprimary, secondary, datesecondary,college, datecollege, degree, Major, postgraduate," & _
       " datepostgraduate, postgraduatecourse, fathername, faddress, fwork, mothername, maddress, mwork, nbi, rescer, issueat, issueon, datehired, dateprob, dateregular, dateresign, employmentstatus, supervisor, reasonres, endofcontract, ClientCode, brcode, " & _
       " PaymentBasis,HourlyRate,DailyRate, MonthlyRate, StartContract, EndContract, NoofMonthsHired, cbamount, cbpayment,cbstartpayment, cbendpayment,  MiscFee,MFPayment,MFStartPayment,MFEndPayment, PlacementFee,PFpayment,PFstartpayment,PFendpayment, WithEcolaPay,  ecolaamount, [USERID], deptcode,remarks, LastUpdate) " & _
       " values ( '" & TxtControlNo.Text & "'," & ChkVal(txtfirstname) & "," & ChkVal(txtfirstchoice) & "," & IIf(txtage.Text = "", 0, txtage.Text) & "," & ChkVal(txtmiddlename) & "," & ChkVal(txtsurname) & "," & ChkVal(txtnickname) & "," & ChkMsk(mskbday) & "," & ChkVal(txtaddress1) & "," & ChkVal(txthouseno) & "," & ChkVal(txtstreet) & "," & ChkVal(txtbarangay) & "," & ChkVal(txtdistrict) & "," & ChkVal(txtcity) & "," & ChkVal(txtprovince) & "" & _
       "," & ChkVal(txtprovinceaddress) & "," & ChkVal(txtemergency) & "," & ChkVal(txteraddress) & "," & ChkVal(txtertel) & ",'" & IIf(optmale.VALUE, "M", "F") & "','" & Replace(txtheight, "'", """") & "'," & ChkVal(txtweight) & "," & ChkVal(txttel1) & "," & ChkVal(txttel2) & ",'" & IIf(cmbcivilstatus.Text = "", "S", cmbcivilstatus.Text) & "'," & ChkVal(txtexam) & "," & ChkVal(txtinterview) & "," & ChkVal(txtmedical) & "," & ChkMsk(msksss) & "," & ChkVal(txtphilhealth) & "," & ChkVal(txtpagibig) & "," & ChkVal(txtacctno) & "" & _
       "," & ChkVal(txtcitizen) & "," & ChkVal(txtbirthplace) & "," & ChkVal(txthaircolor) & "," & ChkVal(txteyecolor) & "," & ChkVal(txtreligion) & "," & ChkMsk(msktax) & "," & ChkVal(txtprimarysch) & "," & ChkVal(txtdateprimary) & "," & ChkVal(txtsecsch) & "," & ChkVal(txtdatesec) & "," & ChkVal(txtcollegesch) & "," & ChkVal(txtdatecollege) & "," & ChkVal(txtcollegecourse) & "," & ChkVal(txtcollegemajor) & "," & ChkVal(txtpostsch) & "," & ChkVal(txtdatepost) & "" & _
       "," & ChkVal(txtpostcourse) & "," & ChkVal(txtfathername) & "," & ChkVal(txtfjob) & "," & ChkVal(txtfaddress) & "," & ChkVal(txtmothername) & "," & ChkVal(txtmaddress) & "," & ChkVal(txtmjob) & "," & ChkVal(txtnbi) & "," & ChkVal(txtrescer) & "," & ChkVal(txtissueat) & "," & ChkMsk(mskissue) & "," & ChkMsk(MskDateHired) & "," & ChkMsk(mskdateprob) & "," & ChkMsk(mskdatereg) & "" & _
       "," & ChkMsk(mskdateres) & ",'" & Left(cmbempstatus.Text, 1) & "'," & ChkVal(txtsupervisor) & "," & ChkVal(txtreasonresign) & "," & ChkMsk(mskendof) & "," & Cid & "," & BrCode & ",'" & PaymentBasis & "'," & HourlyRate & "," & DailyRate & "," & MonthlyRate & "," & ChkMsk(mskStartContract) & "," & ChkMsk(mskEndContract) & "," & IIf(txtMonthsHired.Text = "", 5, txtMonthsHired.Text) & "" & _
       "," & IIf(txtCBtotal = "", 0, Format(txtCBtotal, "##0.00")) & "," & IIf(txtCBpayment = "", 0, Format(txtCBpayment, "##0.00")) & "," & ChkMsk(mskStartCB) & "," & ChkMsk(mskEndCB) & "" & _
       "," & IIf(TxtMFAmount = "", 0, Format(TxtMFAmount, "##0.00")) & "," & IIf(TxtMFPayment = "", 0, Format(TxtMFPayment, "##0.00")) & "," & ChkMsk(MskMFStartPayment) & "," & ChkMsk(MskMFEndPayment) & "" & _
        "," & IIf(txtPFAmount = "", 0, Format(txtPFAmount, "##0.00")) & "," & IIf(txtPFPayment = "", 0, Format(txtPFPayment, "##0.00")) & "," & ChkMsk(MskPFStartPayment) & "," & ChkMsk(MskPFEndPayment) & "" & _
       ",'" & WithEcolaPay & "'," & EcolaAmount & ",'" & operatorID & "'," & DeptCode & ",'" & IIf(Me.txtreasonresign.Text = "", " ", txtreasonresign.Text) & "', '" & Format(Now, "mm/dd/yyyy") & "')"
       
On Error GoTo endSave
    Call Action_Query(sqlString, Cnnstr)
  
    'frmStatus.Show 1
       
   MsgBox "[" & txtsurname.Text & ", " & txtfirstname.Text & "] was successfully saved", vbInformation
   txtID.Text = GetEmployeenum()
   
On Error Resume Next

Dim dept1Bdate, dept2Bdate, dept3Bdate, dept4Bdate, dept5Bdate
dep1bdate = IIf(mskdep1.Text = "__/__/____", ChkMsk(mskdep1), Format(mskdep1.Text, "mm/dd/yyyy"))
 dep2bdate = IIf(mskdep2.Text = "__/__/____", ChkMsk(mskdep2), Format(mskdep2.Text, "mm/dd/yyyy"))
dep3bdate = IIf(mskdep3.Text = "__/__/____", ChkMsk(mskdep3), Format(mskdep3.Text, "mm/dd/yyyy"))
dep4bdate = IIf(mskdep4.Text = "__/__/____", ChkMsk(mskdep4), Format(mskdep4.Text, "mm/dd/yyyy"))
dep5bdate = IIf(mskdep5.Text = "__/__/____", ChkMsk(mskdep5), Format(mskdep5.Text, "mm/dd/yyyy"))


  sqlString = "INSERT INTO hrdetails " & _
      "(employeenum,spousename,spouseaddress,spousework,Dependents ,dep1name,dep2name,dep3name,dep4name,dep5name,dep1bdate,dep2bdate,dep3bdate,dep4bdate,dep5bdate )" & _
      " values ('" & TxtControlNo.Text & "'," & ChkVal(txtspousename) & "," & ChkVal(txtspouseaddress) & "," & ChkVal(txtsjob) & "," & ChkVal(txtnumdep) & "," & ChkVal(txtdep1) & "," & ChkVal(txtdep2) & "," & ChkVal(txtdep3) & "," & ChkVal(txtdep4) & "," & ChkVal(txtdep5) & "," & dep1bdate & "," & dep2bdate & "," & dep3bdate & "," & dep4bdate & "," & dep5bdate & ")"
      
           Call Action_Query(sqlString, Cnnstr)
  
  
  sqlString = "INSERT INTO hrcharreference " & _
      "(employeenum,name,occupation,address,telnum,relation,name1,occupation1,address1,telnum1 " & _
      ", relation1,name2,occupation2,address2,telnum2,relation2 )" & _
      " values ('" & TxtControlNo.Text & "'," & ChkVal(txtrefname1) & "," & ChkVal(txtrefjob1) & "," & ChkVal(txtrefaddress1) & "," & ChkVal(txtrefphone1) & "," & ChkVal(txtrefrel1) & "," & ChkVal(txtrefname2) & "," & ChkVal(txtrefjob2) & "," & ChkVal(txtrefaddress2) & "," & ChkVal(txtrefphone2) & "," & ChkVal(txtrefrel2) & "," & ChkVal(txtrefname3) & "," & ChkVal(txtrefjob3) & "," & ChkVal(txtrefjob3) & "," & ChkVal(txtrefphone3) & "," & ChkVal(txtrefrel3) & ")"
      
           Call Action_Query(sqlString, Cnnstr)


sqlString = "INSERT INTO hrEmploymentRecord " & _
      "(employeenum,company,[from],[to],[position],reason,superior,company1,from1,to1," & _
      " position1,reason1,superior1,company2,from2,to2,position2,reason2,superior2)" & _
      " values ('" & TxtControlNo.Text & "'," & ChkVal(txtnameofcompany1) & "," & ChkVal(txtfrom1) & "," & ChkVal(txtto1) & "," & ChkVal(txtpos1) & "," & ChkVal(txtreason1) & "," & ChkVal(txtsuperior1) & "," & _
                                        " " & ChkVal(txtnameofcompany2) & "," & ChkVal(txtfrom2) & "," & ChkVal(txtto2) & "," & ChkVal(txtpos2) & "," & ChkVal(txtreason2) & "," & ChkVal(txtsuperior2) & "," & _
                                        " " & ChkVal(txtnameofcompany3) & "," & ChkVal(txtfrom3) & "," & ChkVal(txtto3) & "," & ChkVal(txtpos3) & "," & ChkVal(txtreason3) & "," & ChkVal(txtsuperior3) & ")"
           Call Action_Query(sqlString, Cnnstr)
           

sqlString = "INSERT INTO hrseminar " & _
      "(employeenum,seminar1,seminar1date,seminar2,seminar2date,seminar3,seminar3date)" & _
      " values ('" & TxtControlNo.Text & "'," & ChkVal(txtseminar1) & "," & ChkVal(txtdateseminar1) & "," & _
                                        " " & ChkVal(txtseminar2) & "," & ChkVal(txtdateseminar2) & "," & _
                                        " " & ChkVal(txtseminar3) & "," & ChkVal(txtdateseminar3) & ")"
           Call Action_Query(sqlString, Cnnstr)
           
sqlString = "INSERT INTO hrdialects " & _
      "(employeenum,dialect1,dialect2,dialect3) " & _
      " values ('" & TxtControlNo.Text & "'," & ChkVal(txtdialect1) & "," & ChkVal(txtdialect2) & "," & ChkVal(txtdialect3) & ")"
           Call Action_Query(sqlString, Cnnstr)
           

If rsout.State = 1 Then rsout.Close
    Set rsout = New ADODB.Recordset
    rsout.Open "Select * from HREmployees order by  surname", cnDCPI, adOpenKeyset, adLockOptimistic


saveclick = True
Exit Function
endSave:
MsgBox "Encountered Exception in saving. The data you want to save might be invalid. Please review your entry.", vbInformation
saveclick = False
End Function

Public Function getField(ByVal fldname As String, ByVal ClientID As Integer) As Integer
    Dim rs As ADODB.Recordset
    Dim str As String
'
    str = "Select " & fldname & " From clienttable where clientCODE=" & ClientID & ""
    Set rs = cnDCPI.Execute(str)
    If Not rs.EOF Then
       getField = rs(0)
    End If
End Function


Public Function output() As String
    Dim rs As ADODB.Recordset
    Dim str As String
'
    str = "Select max(employeenum) From HREmployees"
    Set rs = cnDCPI.Execute(str)
    If Not rs.EOF Then
         output = Format(rs.Fields(0) + 1, "000000")
    Else
       output = "000001"
    End If
End Function
Private Sub SetButtonsave()
cmdnew.Enabled = True
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
End Sub

Public Sub RefreshBranch()

Set rstBranch = cnDCPI.Execute("select * from HRBRANCH")
cmbBranch.Clear
 Do While Not rstBranch.EOF
        cmbBranch.AddItem UCase(IIf(IsNull(rstBranch!brname), "", rstBranch!brname))
       cmbBranch.ItemData(cmbBranch.NewIndex) = rstBranch!BrCode
        rstBranch.MoveNext
    Loop
    rstBranch.Close
End Sub
Public Sub RefreshDept()
cmbDept.Clear
Set rstDept = cnDCPI.Execute("select * from HRDEPT")
 Do While Not rstDept.EOF
        cmbDept.AddItem UCase(rstDept!deptname)
       cmbDept.ItemData(cmbDept.NewIndex) = rstDept!DeptCode
        rstDept.MoveNext
    Loop
    rstDept.Close
End Sub

Public Function CheckUser1(ByVal operatorID As Integer)
 Dim RsUsers As ADODB.Recordset
 Set RsUsers = New ADODB.Recordset
 RsUsers.Open "select * from userslist where operatorid=" & operatorID & "", Cnnstr, adOpenDynamic, adLockOptimistic
 
 With RsUsers
     If .Fields("maintenancerights") = True Then
          cmdAddBranch.Enabled = True
          cmdAddDept.Enabled = True
       '   Else
        '  cmdAddBranch.Enabled = False
        '  cmdAddDept.Enabled = False
          
     End If
 End With
End Function

Public Sub Showdialect(rec1 As ADODB.Recordset)
    Dim rsdialect As ADODB.Recordset
    txtdialect1 = IIf(IsNull(rec1!dialect1), "", rec1!dialect1)
    txtdialect2 = IIf(IsNull(rec1!dialect2), "", rec1!dialect2)
    txtdialect3 = IIf(IsNull(rec1!dialect3), "", rec1!dialect3)
End Sub
Public Sub Showdetails(Rec2 As ADODB.Recordset)
Dim rsdetails As ADODB.Recordset
   txtspousename.Text = IIf(IsNull(Rec2!spousename), "", Rec2!spousename)
   txtspouseaddress.Text = IIf(IsNull(Rec2!spouseaddress), "", Rec2!spouseaddress)
   txtsjob.Text = IIf(IsNull(Rec2!spousework), "", Rec2!spousework)
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
Dim rsref As ADODB.Recordset
    txtrefname1.Text = IIf(IsNull(rec3!Name), "", rec3!Name)
    txtrefjob1.Text = IIf(IsNull(rec3!occupation), "", rec3!occupation)
    txtrefaddress1.Text = IIf(IsNull(rec3!Address), "", rec3!Address)
    txtrefphone1.Text = IIf(IsNull(rec3!telnum), "", rec3!telnum)
    txtrefrel1.Text = IIf(IsNull(rec3!Relation), "", rec3!Relation)

    txtrefname2.Text = IIf(IsNull(rec3!Name1), "", rec3!Name1)
    txtrefjob2.Text = IIf(IsNull(rec3!occupation1), "", rec3!occupation1)
    txtrefaddress2.Text = IIf(IsNull(rec3!Address1), "", rec3!Address1)
    txtrefphone2.Text = IIf(IsNull(rec3!Telnum1), "", rec3!Telnum1)
    txtrefrel2.Text = IIf(IsNull(rec3!Relation1), "", rec3!Relation1)

    txtrefname3.Text = IIf(IsNull(rec3!Name2), "", rec3!Name2)
    txtrefjob3.Text = IIf(IsNull(rec3!occupation2), "", rec3!occupation2)
    txtrefaddress3.Text = IIf(IsNull(rec3!Address2), "", rec3!Address2)
    txtrefphone3.Text = IIf(IsNull(rec3!Telnum2), "", rec3!Telnum2)
    txtrefrel3.Text = IIf(IsNull(rec3!Relation2), "", rec3!Relation2)

End Sub
Public Sub ShowEmpRec(rec4 As ADODB.Recordset)
Dim rsemprec As ADODB.Recordset
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
'Dim rssem As ADODB.Recordset
    txtseminar1.Text = IIf(IsNull(rec5!seminar1), "", rec5!seminar1)
    txtdateseminar1.Text = IIf(IsNull(rec5!seminar1date), "", rec5!seminar1date)

    txtseminar2.Text = IIf(IsNull(rec5!seminar2), "", rec5!seminar2)
    txtdateseminar2.Text = IIf(IsNull(rec5!seminar2date), "", rec5!seminar2date)

    txtseminar3.Text = IIf(IsNull(rec5!seminar3), "", rec5!seminar3)
    txtdateseminar3.Text = IIf(IsNull(rec5!seminar3date), "", rec5!seminar3date)

End Sub

Private Sub clearscreen()
    TxtControlNo.Text = ""
    txtprimarysch.Text = ""
    txtdateprimary.Text = ""
    txtnickname.Text = ""
    txtsurname.Text = ""
    txtfirstname.Text = ""
    txtmiddlename.Text = ""
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
    cmbcivilstatus.ListIndex = -1
    cmbempstatus.ListIndex = -1
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
    txtnbi.Text = ""
    txtrescer.Text = ""
    txtissueat.Text = ""
    'dtissueon= ""
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
    txtto1.Text = ""
    txtto1.Text = ""
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
    cmbempstatus.ListIndex = 1
    txtsupervisor.Text = ""
    txtreasonresign.Text = ""
    mskendof.Text = "__/__/____"
    
    cmbPayBasis.ListIndex = -1
    txtRate.Text = "0.00"
    mskStartContract.Text = "__/__/____"
    mskEndContract.Text = "__/__/____"
    txtMonthsHired.Text = "5"
    TxtEcolaAmount.Text = "0.00"
    txtCBtotal.Text = "0.00"
    txtCBpayment.Text = "0.00"
    mskStartCB.Text = "__/__/____"
    mskEndCB.Text = "__/__/____"
    TxtMFAmount.Text = "0.00"
    TxtMFPayment.Text = "0.00"
    txtPFPayment.Text = "0.00"
    txtPFAmount.Text = "0.00"
    
    If Not Trim(projectlong) = "" Then
        cmbCompany.Text = projectlong
    Else
        cmbCompany.ListIndex = -1
    End If
    cmbBranch.ListIndex = -1
    cmbDept.ListIndex = -1
End Sub

Private Sub mskb1birth_GotFocus()
mskb1birth.BackColor = &HFFFFC0
End Sub

Private Sub mskb1birth_LostFocus()
mskb1birth.BackColor = vbWhite
End Sub

Private Sub mskb2birth_GotFocus()
mskb2birth.BackColor = &HFFFFC0
End Sub

Private Sub mskb2birth_LostFocus()
mskb2birth.BackColor = vbWhite
End Sub

Private Sub Form_Unload(Cancel As Integer)
If TxtControlNo.Enabled = True Then
If MsgBox("Are you sure you want to close the form without saving?", vbYesNo + vbQuestion) = vbYes Then
    Set rsout = Nothing
    Set rsdetails = Nothing
    Set rsref = Nothing
    Set rsdialect = Nothing
    Set rsemprec = Nothing
    Set rscivil = Nothing
    Set rssem = Nothing
    Set rsempstatus = Nothing
    Set rsemprec = Nothing
    Set rsDelete = Nothing
    Set rsemp = Nothing
    FeMP = False
    Else
    Cancel = True
    End If
  End If
  FeMP = False
End Sub

Private Sub mskbday_GotFocus()
mskbday.BackColor = &HFFFFC0
Call mskHLigths(mskbday)
End Sub


Private Sub mskbday_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub mskbday_LostFocus()
mskbday.BackColor = vbWhite
If mskbday.Text <> "__/__/____" Then
txtage.Text = DateDiff("yyyy", Format(mskbday, "mm/dd/yyyy"), Date)
End If
End Sub

Private Sub mskc1birth_GotFocus()
mskc1birth.BackColor = &HFFFFC0

End Sub

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

Private Sub mskEndCB_GotFocus()
mskEndCB.BackColor = &HFFFFC0
Call mskHLigths(mskEndCB)
End Sub

Private Sub mskEndCB_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SSTab1.Tab = 2
txtfathername.SetFocus
End If

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

Private Sub mskissue_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 SSTab1.Tab = 1
 cmbempstatus.SetFocus
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

Private Sub SSTab1_Click(PreviousTab As Integer)
'Call TabResizing
'If Me.Height = 2400 Then
'Me.Height = 6345
'Else
'Me.Height = 2400
'End If

End Sub
Private Sub Add()
cmdnew.Enabled = False
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

Private Sub txtcontrolno_GotFocus()
Set oCommonFunction.ValidAlpha = TxtControlNo
Call txtHLigths(TxtControlNo)
End Sub


Private Sub TxtEcolaAmount_GotFocus()

  ' TxtEcolaAmount.SelStart = 0
  'TxtEcolaAmount.SelLength = Len(TxtEcolaAmount.Text)
 'TxtEcolaAmount.SetFocus
    
ChkEcola.VALUE = 1
TxtEcolaAmount.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = TxtEcolaAmount
Call txtHLigths(TxtEcolaAmount)


'SendKeys "{HOME}+{END}"
End Sub

Private Sub TxtEcolaAmount_LostFocus()
TxtEcolaAmount.BackColor = vbWhite
Call ToZeroText(TxtEcolaAmount)
''TxtEcolaAmount.Text = Format(TxtEcolaAmount.Text, "#00.00")
If Val(TxtEcolaAmount.Text) <= 0 Then
ChkEcola.VALUE = 0
End If


End Sub

Private Sub txtheight_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtbirthplace.SetFocus

End Sub

Private Sub txtHide_Click()
Me.Height = 2400
End Sub

Private Sub txthouseno_GotFocus()
txthouseno.BackColor = &HFFFFC0
'Set oCommonFunction.ValidInteger = txthouseno
End Sub

Private Sub txthouseno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtIssOn_GotFocus()
txtIssOn.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtIssOn
Call txtHLigths(txtIssOn)
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

Private Sub TxtMFAmount_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then SendKeys "{tab}"
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

Private Sub TxtMFPayment_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then SendKeys "{tab}"
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

Private Sub txtnickname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SSTab1.Tab = 0
txthouseno.SetFocus
End If

End Sub

Private Sub txtnumdep_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub txtPFAMOUNT_GotFocus()
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

Private Sub txtPFAmount_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then SendKeys "{tab}"
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

Private Sub txtPFPayment_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtPFPayment_LostFocus()
txtPFPayment.BackColor = vbWhite
Call ToZeroText(txtPFPayment)
End Sub

Private Sub txtRate_GotFocus()
'txtRate.Text = ""
txtRate.BackColor = &HFFFFC0
Set oCommonFunction.ValidDouble = txtRate
Call txtHLigths(txtRate)
'SendKeys "{HOME}+{END}"
End Sub

Private Sub txtRate_LostFocus()
txtRate.BackColor = vbWhite
Call ToZeroText(txtRate)
'If Val(txtRate.Text) <= 0 And cmbPayBasis.ListIndex >= 0 Then
'MsgBox "Pls enter " & lblRAte.Caption & " greater than zero.", vbInformation
'txtRate.SetFocus
'Exit Sub
'End If

End Sub

Private Sub txtreasonresign_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
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

Private Function IsExist() As Boolean
If editflagship = False Then
Set rsIsExist = New ADODB.Recordset
rsIsExist.Open "Select * from HREmployees where firstname='" & RTrim(Me.txtfirstname.Text) & "' and surname='" & RTrim(Me.txtsurname.Text) & "'", cnDCPI, adOpenKeyset, adLockOptimistic
With rsIsExist
    If .EOF = False Then
        If MsgBox("Employee name [ " & txtsurname & "," & txtfirstname.Text & "] already in the database. Do you want to edit this employee?", vbYesNo + vbInformation, vbInformation) = vbYes Then
        Dim rsouttemp As New ADODB.Recordset
        rsouttemp.Open "select * from HREmployees where firstname='" & RTrim(Me.txtfirstname.Text) & "' and surname='" & RTrim(Me.txtsurname.Text) & "'", cnDCPI, adOpenKeyset, adLockOptimistic
             If rsouttemp.EOF = False Then
                 editflagship = True
                addflagship = False
                 Showdata rsouttemp
             End If
        End If
       IsExist = True
'       rsouttemp.Close
       Set rsouttemp = Nothing
    Exit Function
    End If
End With
End If
IsExist = False
End Function
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
End Sub

Private Sub txtissueat_GotFocus()
txtissueat.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtissueat
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
Call txtHLigths(txtnickname)
Set oCommonFunction.ValidAlpha = txtnickname

End Sub

Private Sub txtnickname_LostFocus()
txtnickname.BackColor = vbWhite
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
'Set oCommonFunction.ValidAlpha = txtreasonresign
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
'Set oCommonFunction.ValidAlpha = txtrefrel3
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

