VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmEmployees 
   BackColor       =   &H80000009&
   Caption         =   "HR Employees"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   Picture         =   "frmEmployees.frx":0000
   ScaleHeight     =   10155
   ScaleWidth      =   10710
   WindowState     =   2  'Maximized
   Begin VB.Frame frmEmployee 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   120
      TabIndex        =   260
      Top             =   720
      Width           =   10455
      Begin VB.TextBox txtappstatus 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4920
         TabIndex        =   272
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtemployeeno 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         TabIndex        =   271
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtsurname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         MaxLength       =   20
         TabIndex        =   270
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtfirstname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   269
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtmiddlename 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   268
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtfirstchoice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   267
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtnickname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   266
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton btnTop 
         Height          =   255
         Left            =   8400
         Picture         =   "frmEmployees.frx":15F942
         Style           =   1  'Graphical
         TabIndex        =   264
         ToolTipText     =   "Top record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton btnPrev 
         Height          =   255
         Left            =   8640
         Picture         =   "frmEmployees.frx":15FA8C
         Style           =   1  'Graphical
         TabIndex        =   263
         ToolTipText     =   "Previous record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton btnNext 
         Height          =   255
         Left            =   8880
         Picture         =   "frmEmployees.frx":15FBD6
         Style           =   1  'Graphical
         TabIndex        =   262
         ToolTipText     =   "Next record"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton btnLast 
         Height          =   255
         Left            =   9120
         Picture         =   "frmEmployees.frx":15FD20
         Style           =   1  'Graphical
         TabIndex        =   261
         ToolTipText     =   "Last record"
         Top             =   0
         Width           =   255
      End
      Begin MSMask.MaskEdBox mskdateapplied 
         Height          =   315
         Left            =   3240
         TabIndex        =   265
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSComCtl2.DTPicker txtdateapplied 
         Height          =   315
         Left            =   3240
         TabIndex        =   273
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60227585
         CurrentDate     =   37581
      End
      Begin VB.TextBox txtDtApp1 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3240
         TabIndex        =   274
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   22
         Left            =   4920
         TabIndex        =   282
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   17
         Left            =   120
         TabIndex        =   281
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Surname"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   18
         Left            =   120
         TabIndex        =   280
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   20
         Left            =   1560
         TabIndex        =   279
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   21
         Left            =   3240
         TabIndex        =   278
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Applied"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   24
         Left            =   3240
         TabIndex        =   277
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   19
         Left            =   1560
         TabIndex        =   276
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nick Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   25
         Left            =   4920
         TabIndex        =   275
         Top             =   840
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdundo 
      Caption         =   "Undo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3840
      Picture         =   "frmEmployees.frx":15FE6A
      Style           =   1  'Graphical
      TabIndex        =   259
      ToolTipText     =   "Undo"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   9960
      Picture         =   "frmEmployees.frx":1604D4
      Style           =   1  'Graphical
      TabIndex        =   258
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdlast 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   8640
      Picture         =   "frmEmployees.frx":16061E
      Style           =   1  'Graphical
      TabIndex        =   257
      ToolTipText     =   "Last"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7680
      Picture         =   "frmEmployees.frx":160B50
      Style           =   1  'Graphical
      TabIndex        =   256
      ToolTipText     =   "Next"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6720
      Picture         =   "frmEmployees.frx":160C4A
      Style           =   1  'Graphical
      TabIndex        =   255
      ToolTipText     =   "Previous"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   5760
      Picture         =   "frmEmployees.frx":160D4C
      Style           =   1  'Graphical
      TabIndex        =   254
      ToolTipText     =   "First"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2880
      Picture         =   "frmEmployees.frx":16127E
      Style           =   1  'Graphical
      TabIndex        =   253
      ToolTipText     =   "Update"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1920
      Picture         =   "frmEmployees.frx":1617B0
      Style           =   1  'Graphical
      TabIndex        =   252
      ToolTipText     =   "Delete"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   960
      Picture         =   "frmEmployees.frx":1618B2
      Style           =   1  'Graphical
      TabIndex        =   251
      ToolTipText     =   "Edit"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   0
      Picture         =   "frmEmployees.frx":1619FC
      Style           =   1  'Graphical
      TabIndex        =   250
      ToolTipText     =   "New"
      Top             =   0
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   131
      Top             =   2400
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
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
      TabPicture(0)   =   "frmEmployees.frx":161F2E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label22(46)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label22(43)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label22(93)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label22(88)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label22(23)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label22(10)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label22(12)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label22(11)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label22(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label22(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label22(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label22(7)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label22(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label22(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label23(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label22(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label23(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label22(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label22(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label22(9)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label22(98)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label22(99)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label22(100)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label22(101)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label22(102)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label22(44)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtIssOn"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtDtBrth"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "msksss"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "msktax"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtnbi"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Frame5"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Frame1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtreligion"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Frame3"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtacctno"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmbcivilstatus"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txttel4"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txthouseno"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtheight"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtweight"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtpagibig"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txttel2"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txttel1"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtprovinceaddress"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtaddress1"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtage"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtcitizen"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtbirthplace"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txttel3"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Frame4"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Frame7"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtstreet"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtbarangay"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtcity"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtprovince"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtdistrict"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "mskbday"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "dtissueon"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "mskbirthdate"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "mskdatehired"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "mskissue"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txtphilhealth"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).ControlCount=   64
      TabCaption(1)   =   "Family/Educational Background"
      TabPicture(1)   =   "frmEmployees.frx":161F4A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label22(33)"
      Tab(1).Control(1)=   "Label22(32)"
      Tab(1).Control(2)=   "Label22(71)"
      Tab(1).Control(3)=   "Label22(70)"
      Tab(1).Control(4)=   "Label22(69)"
      Tab(1).Control(5)=   "Label22(68)"
      Tab(1).Control(6)=   "Label22(67)"
      Tab(1).Control(7)=   "Label22(62)"
      Tab(1).Control(8)=   "Label22(61)"
      Tab(1).Control(9)=   "Label22(59)"
      Tab(1).Control(10)=   "Label22(58)"
      Tab(1).Control(11)=   "Label22(57)"
      Tab(1).Control(12)=   "Label22(56)"
      Tab(1).Control(13)=   "Label22(55)"
      Tab(1).Control(14)=   "Label22(54)"
      Tab(1).Control(15)=   "Label22(53)"
      Tab(1).Control(16)=   "Label22(51)"
      Tab(1).Control(17)=   "Label22(50)"
      Tab(1).Control(18)=   "Label22(49)"
      Tab(1).Control(19)=   "Label22(31)"
      Tab(1).Control(20)=   "Label22(30)"
      Tab(1).Control(21)=   "Label22(28)"
      Tab(1).Control(22)=   "Label22(27)"
      Tab(1).Control(23)=   "Label22(34)"
      Tab(1).Control(24)=   "Label22(36)"
      Tab(1).Control(25)=   "Label22(37)"
      Tab(1).Control(26)=   "Label22(38)"
      Tab(1).Control(27)=   "Label22(26)"
      Tab(1).Control(28)=   "Label22(29)"
      Tab(1).Control(29)=   "Label22(39)"
      Tab(1).Control(30)=   "Label22(40)"
      Tab(1).Control(31)=   "Label22(41)"
      Tab(1).Control(32)=   "Label22(42)"
      Tab(1).Control(33)=   "Label22(35)"
      Tab(1).Control(34)=   "Label22(103)"
      Tab(1).Control(35)=   "mskdep5"
      Tab(1).Control(36)=   "mskhpbdate"
      Tab(1).Control(37)=   "mskdep2"
      Tab(1).Control(38)=   "mskdep1"
      Tab(1).Control(39)=   "mskdep4"
      Tab(1).Control(40)=   "mskdep3"
      Tab(1).Control(41)=   "txtspouseaddress"
      Tab(1).Control(42)=   "txtmaddress"
      Tab(1).Control(43)=   "txtfaddress"
      Tab(1).Control(44)=   "txtdateseminar3"
      Tab(1).Control(45)=   "txtdateseminar2"
      Tab(1).Control(46)=   "Frame6"
      Tab(1).Control(47)=   "txtdateseminar1"
      Tab(1).Control(48)=   "txtdialect1"
      Tab(1).Control(49)=   "txtdialect2"
      Tab(1).Control(50)=   "txtdialect3"
      Tab(1).Control(51)=   "txtseminar3"
      Tab(1).Control(52)=   "txtseminar2"
      Tab(1).Control(53)=   "txtseminar1"
      Tab(1).Control(54)=   "txtpostsch"
      Tab(1).Control(55)=   "txtcollegesch"
      Tab(1).Control(56)=   "txtsecsch"
      Tab(1).Control(57)=   "txtprimarysch"
      Tab(1).Control(58)=   "txtpostcourse"
      Tab(1).Control(59)=   "txtcollegecourse"
      Tab(1).Control(60)=   "txtcollegemajor"
      Tab(1).Control(61)=   "txthprelate"
      Tab(1).Control(62)=   "txthfdep"
      Tab(1).Control(63)=   "txtnumdep"
      Tab(1).Control(64)=   "txtdep4"
      Tab(1).Control(65)=   "txtdep3"
      Tab(1).Control(66)=   "txtdep2"
      Tab(1).Control(67)=   "txtfathername"
      Tab(1).Control(68)=   "txtmothername"
      Tab(1).Control(69)=   "txtspousename"
      Tab(1).Control(70)=   "txtfjob"
      Tab(1).Control(71)=   "txtmjob"
      Tab(1).Control(72)=   "txtsjob"
      Tab(1).Control(73)=   "txtdep1"
      Tab(1).Control(74)=   "txtdep5"
      Tab(1).Control(75)=   "txtdatepost"
      Tab(1).Control(76)=   "txtdatecollege"
      Tab(1).Control(77)=   "txtdatesec"
      Tab(1).Control(78)=   "txtdateprimary"
      Tab(1).Control(79)=   "msdep1"
      Tab(1).Control(80)=   "msdep2"
      Tab(1).Control(81)=   "msdep3"
      Tab(1).Control(82)=   "msdep4"
      Tab(1).Control(83)=   "msdep5"
      Tab(1).Control(84)=   "mshpbdate"
      Tab(1).ControlCount=   85
      TabCaption(2)   =   "Employment/References Background"
      TabPicture(2)   =   "frmEmployees.frx":161F66
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label22(72)"
      Tab(2).Control(1)=   "Label22(73)"
      Tab(2).Control(2)=   "Label22(77)"
      Tab(2).Control(3)=   "Label22(78)"
      Tab(2).Control(4)=   "Label22(80)"
      Tab(2).Control(5)=   "Label22(81)"
      Tab(2).Control(6)=   "Label22(84)"
      Tab(2).Control(7)=   "Label22(74)"
      Tab(2).Control(8)=   "Label22(75)"
      Tab(2).Control(9)=   "Label22(76)"
      Tab(2).Control(10)=   "Label22(79)"
      Tab(2).Control(11)=   "Label22(82)"
      Tab(2).Control(12)=   "Label22(83)"
      Tab(2).Control(13)=   "txtpos1"
      Tab(2).Control(14)=   "txtpos2"
      Tab(2).Control(15)=   "txtpos3"
      Tab(2).Control(16)=   "txtreason1"
      Tab(2).Control(17)=   "txtreason2"
      Tab(2).Control(18)=   "txtreason3"
      Tab(2).Control(19)=   "txtsuperior1"
      Tab(2).Control(20)=   "txtsuperior2"
      Tab(2).Control(21)=   "txtsuperior3"
      Tab(2).Control(22)=   "Frame2"
      Tab(2).Control(23)=   "txtrefname1"
      Tab(2).Control(24)=   "txtrefjob1"
      Tab(2).Control(25)=   "txtrefaddress1"
      Tab(2).Control(26)=   "txtrefphone1"
      Tab(2).Control(27)=   "txtrefrel1"
      Tab(2).Control(28)=   "txtrefname2"
      Tab(2).Control(29)=   "txtrefjob2"
      Tab(2).Control(30)=   "txtrefaddress2"
      Tab(2).Control(31)=   "txtrefphone2"
      Tab(2).Control(32)=   "txtrefrel2"
      Tab(2).Control(33)=   "txtrefname3"
      Tab(2).Control(34)=   "txtrefjob3"
      Tab(2).Control(35)=   "txtrefaddress3"
      Tab(2).Control(36)=   "txtrefphone3"
      Tab(2).Control(37)=   "txtrefrel3"
      Tab(2).Control(38)=   "txtfrom2"
      Tab(2).Control(39)=   "txtfrom3"
      Tab(2).Control(40)=   "txtto1"
      Tab(2).Control(41)=   "txtto2"
      Tab(2).Control(42)=   "txtto3"
      Tab(2).Control(43)=   "txtfrom1"
      Tab(2).Control(44)=   "txtnameofcompany3"
      Tab(2).Control(45)=   "txtnameofcompany2"
      Tab(2).Control(46)=   "txtnameofcompany1"
      Tab(2).ControlCount=   47
      TabCaption(3)   =   "Salary Information"
      TabPicture(3)   =   "frmEmployees.frx":161F82
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.TextBox txtphilhealth 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4320
         TabIndex        =   21
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txtnameofcompany1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   93
         Top             =   1170
         Width           =   2295
      End
      Begin VB.TextBox txtnameofcompany2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   99
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtnameofcompany3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   105
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtfrom1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -72360
         MaxLength       =   50
         TabIndex        =   94
         Top             =   1170
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskissue 
         Height          =   315
         Left            =   9960
         TabIndex        =   28
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox mskdatehired 
         Height          =   315
         Left            =   4560
         TabIndex        =   32
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox mskbirthdate 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox mshpbdate 
         Height          =   315
         Left            =   -64560
         TabIndex        =   61
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox msdep5 
         Height          =   315
         Left            =   -64560
         TabIndex        =   71
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox msdep4 
         Height          =   315
         Left            =   -68520
         TabIndex        =   69
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox msdep3 
         Height          =   315
         Left            =   -68520
         TabIndex        =   67
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox msdep2 
         Height          =   315
         Left            =   -72480
         TabIndex        =   65
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox msdep1 
         Height          =   315
         Left            =   -72480
         TabIndex        =   63
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSComCtl2.DTPicker dtissueon 
         Height          =   315
         Left            =   9960
         TabIndex        =   246
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60227585
         CurrentDate     =   37581
      End
      Begin MSComCtl2.DTPicker mskbday 
         Height          =   315
         Left            =   1320
         TabIndex        =   243
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60227585
         CurrentDate     =   38171.6113773148
      End
      Begin VB.TextBox txtto3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70680
         MaxLength       =   50
         TabIndex        =   107
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtto2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70680
         MaxLength       =   50
         TabIndex        =   101
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtto1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70680
         MaxLength       =   50
         TabIndex        =   95
         Top             =   1170
         Width           =   1455
      End
      Begin VB.TextBox txtfrom3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -72360
         MaxLength       =   50
         TabIndex        =   106
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtfrom2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -72360
         MaxLength       =   50
         TabIndex        =   100
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtdateprimary 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   82
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtdatesec 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   80
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtdatecollege 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   76
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox txtdatepost 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   73
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtdep5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -66600
         MaxLength       =   50
         TabIndex        =   70
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtdistrict 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7080
         TabIndex        =   3
         Top             =   660
         Width           =   1455
      End
      Begin VB.TextBox txtprovince 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1380
         Width           =   2535
      End
      Begin VB.TextBox txtcity 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   1020
         Width           =   1455
      End
      Begin VB.TextBox txtbarangay 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4920
         TabIndex        =   2
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox txtstreet 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3120
         TabIndex        =   1
         Top             =   660
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Height          =   1455
         Left            =   5280
         TabIndex        =   218
         Top             =   5400
         Width           =   6255
         Begin VB.TextBox txtActype 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   4080
            TabIndex        =   48
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtnoot 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   4080
            TabIndex        =   49
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtnoundertime 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   4080
            TabIndex        =   47
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtnolate 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1320
            TabIndex        =   46
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txthrswork 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1320
            TabIndex        =   45
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtdtrexempt 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1320
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "No OverTime"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   97
            Left            =   2760
            TabIndex        =   224
            Top             =   1080
            Width           =   1065
         End
         Begin VB.Label txtacttype 
            AutoSize        =   -1  'True
            Caption         =   "ActType"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   96
            Left            =   2760
            TabIndex        =   223
            Top             =   720
            Width           =   675
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "No Undertime"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   87
            Left            =   2760
            TabIndex        =   222
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "No Late"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   86
            Left            =   240
            TabIndex        =   221
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Hrs. Work"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   85
            Left            =   240
            TabIndex        =   220
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "DTRExempt"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   66
            Left            =   240
            TabIndex        =   219
            Top             =   360
            Width           =   930
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   3120
         TabIndex        =   209
         Top             =   3840
         Width           =   8415
         Begin MSMask.MaskEdBox mskendof 
            Height          =   315
            Left            =   6720
            TabIndex        =   39
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSMask.MaskEdBox mskdateres 
            Height          =   315
            Left            =   4080
            TabIndex        =   35
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSMask.MaskEdBox mskdatereg 
            Height          =   315
            Left            =   1440
            TabIndex        =   34
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSMask.MaskEdBox mskdateprob 
            Height          =   315
            Left            =   1440
            TabIndex        =   33
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
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
         Begin VB.ComboBox cmbdepartment 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   6720
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtreasonresign 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   6720
            MaxLength       =   40
            TabIndex        =   38
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtsupervisor 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   4080
            MaxLength       =   40
            TabIndex        =   37
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox cmbempstatus 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   600
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtendcontract 
            Height          =   315
            Left            =   6720
            TabIndex        =   130
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   60227585
            CurrentDate     =   37581
         End
         Begin MSComCtl2.DTPicker dtdateresign 
            Height          =   315
            Left            =   4080
            TabIndex        =   129
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   60227585
            CurrentDate     =   37581
         End
         Begin MSComCtl2.DTPicker dtdateprob 
            Height          =   315
            Left            =   1440
            TabIndex        =   127
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   60227585
            CurrentDate     =   37581
         End
         Begin MSComCtl2.DTPicker dtdatereg 
            Height          =   315
            Left            =   1440
            TabIndex        =   128
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   60227585
            CurrentDate     =   37581
         End
         Begin MSComCtl2.DTPicker txtdatehired 
            Height          =   315
            Left            =   1440
            TabIndex        =   233
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   60227585
            CurrentDate     =   38171.6113773148
         End
         Begin VB.TextBox txtDtHr 
            Height          =   285
            Left            =   1440
            TabIndex        =   236
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtDtPr 
            Height          =   285
            Left            =   1440
            TabIndex        =   237
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtDtReg 
            Height          =   285
            Left            =   1440
            TabIndex        =   238
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtDtRes 
            Height          =   285
            Left            =   4080
            TabIndex        =   239
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtECt 
            Height          =   285
            Left            =   6720
            TabIndex        =   240
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Department"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   96
            Left            =   5640
            TabIndex        =   225
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "End Contract"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   65
            Left            =   5640
            TabIndex        =   217
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Reason Res."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   64
            Left            =   5640
            TabIndex        =   216
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Supervisor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   63
            Left            =   3000
            TabIndex        =   215
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Emp. Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   60
            Left            =   3000
            TabIndex        =   214
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Date Probation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   52
            Left            =   120
            TabIndex        =   213
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Date Regular"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   48
            Left            =   120
            TabIndex        =   212
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Date Resign"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   47
            Left            =   3000
            TabIndex        =   211
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Date Hired"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   16
            Left            =   120
            TabIndex        =   210
            Top             =   360
            Width           =   840
         End
      End
      Begin VB.TextBox txttel3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1320
         Width           =   1560
      End
      Begin VB.TextBox txtbirthplace 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2160
         Width           =   1680
      End
      Begin VB.TextBox txtcitizen 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   18
         Top             =   2100
         Width           =   480
      End
      Begin VB.TextBox txtage 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   17
         Top             =   1740
         Width           =   495
      End
      Begin VB.TextBox txtaddress1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1020
         Width           =   4935
      End
      Begin VB.TextBox txtprovinceaddress 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1380
         Width           =   3015
      End
      Begin VB.TextBox txttel1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txttel2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtpagibig 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox txtdep1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74280
         MaxLength       =   50
         TabIndex        =   62
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtsjob 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   58
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtmjob 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   55
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtfjob 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   52
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtspousename 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   56
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtmothername 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   53
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtfathername 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   50
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtdep2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74280
         MaxLength       =   50
         TabIndex        =   64
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtweight 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6720
         MaxLength       =   8
         TabIndex        =   25
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtheight 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6720
         MaxLength       =   8
         TabIndex        =   24
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txthouseno 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox txttel4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   9960
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox cmbcivilstatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   9960
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtacctno 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   9960
         TabIndex        =   27
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtdep3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70560
         MaxLength       =   50
         TabIndex        =   66
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtdep4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70560
         MaxLength       =   50
         TabIndex        =   68
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtnumdep 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -64920
         MaxLength       =   50
         TabIndex        =   72
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   5640
         TabIndex        =   126
         Top             =   1800
         Width           =   2175
         Begin VB.OptionButton optmale 
            Caption         =   "Male"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optfemale 
            Caption         =   "Female"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1080
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtreligion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   360
         TabIndex        =   138
         Top             =   3840
         Width           =   2655
         Begin VB.TextBox txtmedical 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1560
            TabIndex        =   31
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtinterview 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1560
            TabIndex        =   30
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtexam 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1560
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Medical Grade"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   91
            Left            =   120
            TabIndex        =   141
            Top             =   1080
            Width           =   1170
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Interview Grade"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   90
            Left            =   120
            TabIndex        =   140
            Top             =   720
            Width           =   1320
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Exam. Grade"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   89
            Left            =   120
            TabIndex        =   139
            Top             =   360
            Width           =   1020
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "In case of Emergency"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1455
         Left            =   360
         TabIndex        =   134
         Top             =   5400
         Width           =   4815
         Begin VB.TextBox txtertel 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   43
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox txtemergency 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1920
            MaxLength       =   40
            TabIndex        =   41
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox txteraddress 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   42
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Telephone#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   15
            Left            =   360
            TabIndex        =   137
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   14
            Left            =   360
            TabIndex        =   136
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Contact Person"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Index           =   13
            Left            =   360
            TabIndex        =   135
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.TextBox txthfdep 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -64560
         MaxLength       =   50
         TabIndex        =   59
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txthprelate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -64560
         MaxLength       =   50
         TabIndex        =   60
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtcollegemajor 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -65400
         TabIndex        =   79
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox txtcollegecourse 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -65400
         TabIndex        =   78
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox txtpostcourse 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -65400
         MaxLength       =   20
         TabIndex        =   75
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox txtprimarysch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -71760
         MaxLength       =   50
         TabIndex        =   83
         Top             =   4680
         Width           =   6015
      End
      Begin VB.TextBox txtsecsch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -71760
         MaxLength       =   50
         TabIndex        =   81
         Top             =   4320
         Width           =   6015
      End
      Begin VB.TextBox txtcollegesch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -71760
         MaxLength       =   50
         TabIndex        =   77
         Top             =   3960
         Width           =   6015
      End
      Begin VB.TextBox txtpostsch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -71760
         MaxLength       =   50
         TabIndex        =   74
         Top             =   3600
         Width           =   6015
      End
      Begin VB.TextBox txtseminar1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -71400
         MaxLength       =   40
         TabIndex        =   87
         Top             =   5400
         Width           =   3855
      End
      Begin VB.TextBox txtseminar2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -71400
         MaxLength       =   40
         TabIndex        =   89
         Top             =   5760
         Width           =   3855
      End
      Begin VB.TextBox txtseminar3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -71400
         MaxLength       =   40
         TabIndex        =   91
         Top             =   6120
         Width           =   3855
      End
      Begin VB.TextBox txtdialect3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74520
         MaxLength       =   25
         TabIndex        =   86
         Top             =   6120
         Width           =   2655
      End
      Begin VB.TextBox txtdialect2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74520
         MaxLength       =   25
         TabIndex        =   85
         Top             =   5760
         Width           =   2655
      End
      Begin VB.TextBox txtdialect1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74520
         MaxLength       =   25
         TabIndex        =   84
         Top             =   5400
         Width           =   2655
      End
      Begin VB.TextBox txtdateseminar1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67320
         MaxLength       =   40
         TabIndex        =   88
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Height          =   135
         Left            =   -74880
         TabIndex        =   133
         Top             =   3120
         Width           =   11775
      End
      Begin VB.TextBox txtdateseminar2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67320
         MaxLength       =   40
         TabIndex        =   90
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox txtdateseminar3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67320
         MaxLength       =   40
         TabIndex        =   92
         Top             =   6120
         Width           =   1575
      End
      Begin VB.TextBox txtrefrel3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -65640
         MaxLength       =   30
         TabIndex        =   125
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox txtrefphone3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67440
         MaxLength       =   35
         TabIndex        =   124
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox txtrefaddress3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -69960
         MaxLength       =   50
         TabIndex        =   123
         Top             =   3840
         Width           =   2415
      End
      Begin VB.TextBox txtrefjob3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -72360
         MaxLength       =   30
         TabIndex        =   122
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox txtrefname3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   121
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox txtrefrel2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -65640
         MaxLength       =   30
         TabIndex        =   120
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox txtrefphone2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67440
         MaxLength       =   35
         TabIndex        =   119
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox txtrefaddress2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -69960
         MaxLength       =   50
         TabIndex        =   118
         Top             =   3480
         Width           =   2415
      End
      Begin VB.TextBox txtrefjob2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -72360
         MaxLength       =   30
         TabIndex        =   117
         Top             =   3480
         Width           =   2295
      End
      Begin VB.TextBox txtrefname2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   116
         Top             =   3480
         Width           =   2295
      End
      Begin VB.TextBox txtrefrel1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -65640
         MaxLength       =   30
         TabIndex        =   115
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtrefphone1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67440
         MaxLength       =   35
         TabIndex        =   114
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox txtrefaddress1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -69960
         MaxLength       =   50
         TabIndex        =   113
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox txtrefjob1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -72360
         MaxLength       =   30
         TabIndex        =   112
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox txtrefname1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   111
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   -74760
         TabIndex        =   132
         Top             =   2280
         Width           =   11055
      End
      Begin VB.TextBox txtsuperior3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -65640
         MaxLength       =   50
         TabIndex        =   110
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtsuperior2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -65640
         MaxLength       =   50
         TabIndex        =   104
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtsuperior1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -65640
         MaxLength       =   50
         TabIndex        =   98
         Top             =   1170
         Width           =   1935
      End
      Begin VB.TextBox txtreason3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67320
         MaxLength       =   50
         TabIndex        =   109
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtreason2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67320
         MaxLength       =   50
         TabIndex        =   103
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtreason1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67320
         MaxLength       =   50
         TabIndex        =   97
         Top             =   1170
         Width           =   1575
      End
      Begin VB.TextBox txtpos3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -69000
         MaxLength       =   50
         TabIndex        =   108
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtpos2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -69000
         MaxLength       =   50
         TabIndex        =   102
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtpos1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -69000
         MaxLength       =   50
         TabIndex        =   96
         Top             =   1170
         Width           =   1575
      End
      Begin VB.TextBox txtfaddress 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   -70920
         MaxLength       =   50
         TabIndex        =   51
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtmaddress 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   -70920
         MaxLength       =   50
         TabIndex        =   54
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtspouseaddress 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   -70920
         MaxLength       =   50
         TabIndex        =   57
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtnbi 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1320
         TabIndex        =   15
         Top             =   2880
         Width           =   1695
      End
      Begin MSMask.MaskEdBox msktax 
         Height          =   315
         Left            =   4320
         TabIndex        =   20
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   12582912
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
         Height          =   315
         Left            =   4320
         TabIndex        =   19
         Top             =   2460
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   12582912
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
      Begin VB.TextBox txtDtBrth 
         Height          =   285
         Left            =   1320
         TabIndex        =   234
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtIssOn 
         Height          =   285
         Left            =   9960
         TabIndex        =   235
         Top             =   3000
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker mskdep3 
         Height          =   315
         Left            =   -68520
         TabIndex        =   241
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60227585
         CurrentDate     =   38171.6113773148
      End
      Begin MSComCtl2.DTPicker mskdep4 
         Height          =   315
         Left            =   -68520
         TabIndex        =   242
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60227585
         CurrentDate     =   38171.6113773148
      End
      Begin MSComCtl2.DTPicker mskdep1 
         Height          =   315
         Left            =   -72480
         TabIndex        =   244
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60227585
         CurrentDate     =   38171.6113773148
      End
      Begin MSComCtl2.DTPicker mskdep2 
         Height          =   315
         Left            =   -72480
         TabIndex        =   245
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60227585
         CurrentDate     =   38171.6113773148
      End
      Begin MSComCtl2.DTPicker mskhpbdate 
         Height          =   315
         Left            =   -64560
         TabIndex        =   247
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60227585
         CurrentDate     =   38171.6113773148
      End
      Begin MSComCtl2.DTPicker mskdep5 
         Height          =   315
         Left            =   -64560
         TabIndex        =   248
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60227585
         CurrentDate     =   38171.6113773148
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "PHILHEALTH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   44
         Left            =   3240
         TabIndex        =   249
         Top             =   3292
         Width           =   990
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   103
         Left            =   -64560
         TabIndex        =   232
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Dep5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   35
         Left            =   -67080
         TabIndex        =   231
         Top             =   2400
         Width           =   405
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "District"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   102
         Left            =   6360
         TabIndex        =   230
         Top             =   780
         Width           =   585
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Province"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   101
         Left            =   5160
         TabIndex        =   229
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   100
         Left            =   360
         TabIndex        =   228
         Top             =   1140
         Width           =   315
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Brgy."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   99
         Left            =   4440
         TabIndex        =   227
         Top             =   780
         Width           =   420
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Street"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   98
         Left            =   2520
         TabIndex        =   226
         Top             =   780
         Width           =   510
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Tel. No.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   9
         Left            =   8760
         TabIndex        =   208
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "TAX ID#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   3
         Left            =   3240
         TabIndex        =   207
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "SSS#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   2
         Left            =   3240
         TabIndex        =   206
         Top             =   2580
         Width           =   405
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Birth Place"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   205
         Top             =   2220
         Width           =   870
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Citizenship"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   5
         Left            =   3240
         TabIndex        =   204
         Top             =   2220
         Width           =   930
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   3240
         TabIndex        =   203
         Top             =   1860
         Width           =   330
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Boarding House"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   2280
         TabIndex        =   202
         Top             =   1140
         Width           =   1305
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Provincial Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   201
         Top             =   1500
         Width           =   1560
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Tel. No.1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   7
         Left            =   8760
         TabIndex        =   200
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Tel. No.2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   8
         Left            =   8760
         TabIndex        =   199
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "BirthDate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   4
         Left            =   360
         TabIndex        =   198
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "PAGIBIG#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   6
         Left            =   360
         TabIndex        =   197
         Top             =   3292
         Width           =   750
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Dep4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   42
         Left            =   -71040
         TabIndex        =   196
         Top             =   2760
         Width           =   405
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Dep3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   41
         Left            =   -71040
         TabIndex        =   195
         Top             =   2400
         Width           =   405
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Dep2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   40
         Left            =   -74760
         TabIndex        =   194
         Top             =   2760
         Width           =   405
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Dep1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   39
         Left            =   -74760
         TabIndex        =   193
         Top             =   2400
         Width           =   405
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   29
         Left            =   -66720
         TabIndex        =   192
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   26
         Left            =   -72600
         TabIndex        =   191
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Spouse Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   38
         Left            =   -74760
         TabIndex        =   190
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Mother Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   37
         Left            =   -74760
         TabIndex        =   189
         Top             =   1320
         Width           =   1110
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Father Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   36
         Left            =   -74760
         TabIndex        =   188
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   11
         Left            =   6120
         TabIndex        =   187
         Top             =   2760
         Width           =   525
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   12
         Left            =   6120
         TabIndex        =   186
         Top             =   3120
         Width           =   570
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   34
         Left            =   -68400
         TabIndex        =   185
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "No. of House"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   10
         Left            =   360
         TabIndex        =   184
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Tel. No.4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   23
         Left            =   8760
         TabIndex        =   183
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Civil Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   88
         Left            =   8760
         TabIndex        =   182
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Acct No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   93
         Left            =   8760
         TabIndex        =   181
         Top             =   2572
         Width           =   660
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Total No. of Dependents"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   27
         Left            =   -67080
         TabIndex        =   180
         Top             =   2760
         Width           =   1965
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Religion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   360
         TabIndex        =   179
         Top             =   2615
         Width           =   660
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "HFDep"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   28
         Left            =   -65280
         TabIndex        =   178
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "HPBDate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   30
         Left            =   -65280
         TabIndex        =   177
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "HPRel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   31
         Left            =   -65280
         TabIndex        =   176
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Major In - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   49
         Left            =   -65400
         TabIndex        =   175
         Top             =   4440
         Width           =   810
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Course/Degree"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   50
         Left            =   -65160
         TabIndex        =   174
         Top             =   3360
         Width           =   1260
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Name of School"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   51
         Left            =   -69480
         TabIndex        =   173
         Top             =   3360
         Width           =   1275
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Year Graduated"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   53
         Left            =   -73440
         TabIndex        =   172
         Top             =   3360
         Width           =   1275
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Primary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   54
         Left            =   -74760
         TabIndex        =   171
         Top             =   4800
         Width           =   645
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Secondary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   55
         Left            =   -74760
         TabIndex        =   170
         Top             =   4440
         Width           =   870
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "College"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   56
         Left            =   -74760
         TabIndex        =   169
         Top             =   4080
         Width           =   630
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Post Graduate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   57
         Left            =   -74760
         TabIndex        =   168
         Top             =   3720
         Width           =   1170
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trainings/Seminars Attended"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   58
         Left            =   -71640
         TabIndex        =   167
         Top             =   5160
         Width           =   2145
      End
      Begin VB.Label Label22 
         Caption         =   "(1)"
         Height          =   255
         Index           =   59
         Left            =   -71640
         TabIndex        =   166
         Top             =   5460
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "(2)"
         Height          =   255
         Index           =   61
         Left            =   -71640
         TabIndex        =   165
         Top             =   5820
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "(3)"
         Height          =   255
         Index           =   62
         Left            =   -71640
         TabIndex        =   164
         Top             =   6180
         Width           =   255
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dates"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   67
         Left            =   -66840
         TabIndex        =   163
         Top             =   5160
         Width           =   435
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Languages/Dialects"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   68
         Left            =   -74760
         TabIndex        =   162
         Top             =   5160
         Width           =   1470
      End
      Begin VB.Label Label22 
         Caption         =   "(1)"
         Height          =   255
         Index           =   69
         Left            =   -74760
         TabIndex        =   161
         Top             =   5460
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "(2)"
         Height          =   255
         Index           =   70
         Left            =   -74760
         TabIndex        =   160
         Top             =   5820
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "(3)"
         Height          =   255
         Index           =   71
         Left            =   -74760
         TabIndex        =   159
         Top             =   6180
         Width           =   255
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Relation to Applicant"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   83
         Left            =   -65520
         TabIndex        =   158
         Top             =   2880
         Width           =   1680
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Business Phone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   82
         Left            =   -67200
         TabIndex        =   157
         Top             =   2880
         Width           =   1350
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Business Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   79
         Left            =   -69600
         TabIndex        =   156
         Top             =   2880
         Width           =   1545
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   76
         Left            =   -71640
         TabIndex        =   155
         Top             =   2880
         Width           =   915
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   75
         Left            =   -73920
         TabIndex        =   154
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "References"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   74
         Left            =   -74760
         TabIndex        =   153
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Immediate Superior"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   84
         Left            =   -65520
         TabIndex        =   152
         Top             =   900
         Width           =   1650
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Reason for Leaving"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   81
         Left            =   -67320
         TabIndex        =   151
         Top             =   900
         Width           =   1590
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   80
         Left            =   -68640
         TabIndex        =   150
         Top             =   900
         Width           =   675
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   78
         Left            =   -70200
         TabIndex        =   149
         Top             =   960
         Width           =   210
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   77
         Left            =   -72000
         TabIndex        =   148
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Name of Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   73
         Left            =   -74400
         TabIndex        =   147
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Employment Record (List most recent first)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   72
         Left            =   -74760
         TabIndex        =   146
         Top             =   600
         Width           =   3600
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   32
         Left            =   -69600
         TabIndex        =   145
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   33
         Left            =   -72120
         TabIndex        =   144
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "NBI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   43
         Left            =   360
         TabIndex        =   143
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Issue On"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   46
         Left            =   8760
         TabIndex        =   142
         Top             =   3052
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsout As ADODB.Recordset
Dim rsdetails As ADODB.Recordset
Dim rsref As ADODB.Recordset
Dim rsdialect As ADODB.Recordset
Dim rssem As ADODB.Recordset
Dim rsemprec As ADODB.Recordset
Dim dep1bdate, dep2bdate, dep3bdate, dep4bdate, dep5bdate As String
Dim bdt, bdt1, appdt, hiredt, probdt, regdt, resdt, eocdt, isson As Date
Dim depcd, txage As Integer
Dim gendr, civst, empst, usrid  As String
Dim dep1dt, dep2dt, dep3dt, dep4dt, dep5dt As String
'****09/28/04***** makiss
'Dim dep1dt, dep2dt, dep3dt, dep4dt, dep5dt As Date
Dim numdep As Integer

Private Sub edit()
    cmdnew.Enabled = False
    cmdedit.Enabled = False
    cmdupdate.Enabled = True
   ' cmdhired.Enabled = False
    cmdundo.Enabled = True
    cmddelete.Enabled = False
 '   cmdsearch.Enabled = False
    cmdfirst.Enabled = False
    cmdprev.Enabled = False
    cmdnext.Enabled = False
    cmdlast.Enabled = False
    cmdundo.Enabled = True
End Sub

Private Sub cmbcivil_GotFocus()
cmbcivil.BackColor = &H80000016
End Sub

Private Sub cmbcivil_LostFocus()
cmbcivil.BackColor = vbWhite
End Sub

Private Sub cmbcivilstatus_GotFocus()
cmbcivilstatus.BackColor = &H80000016
End Sub

Private Sub cmbcivilstatus_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub cmbcivilstatus_LostFocus()
cmbcivilstatus.BackColor = vbWhite
End Sub

Private Sub cmbempstatus_GotFocus()
cmbempstatus.BackColor = &H80000016
End Sub

Private Sub cmbempstatus_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub cmbempstatus_LostFocus()
cmbempstatus.BackColor = vbWhite
End Sub

Private Sub cmdclose_Click()
Dim res As String
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim res
Dim str, str1, str2, str3, str4, str5 As String
Dim applicantnum As String


res = MsgBox("Are you sure you want to delete Employee No: " & txtEmployeeNo & "?", vbQuestion + vbYesNo, "DCPI Information System")
If res = vbYes Then
    cnDCPI.Execute "EXEC DelHREmp '" & txtEmployeeNo & "'"
    res = MsgBox("Record Deleted", vbCritical, "DCPI Information System")
    rsout.Close
    rsout.Open "Select * from HREmployees", cnDCPI, adOpenKeyset, adLockOptimistic
    rsout.MoveLast
         showdata rsout
          cmdnext.Enabled = True
          cmdlast.Enabled = True
End If

End Sub

Private Sub cmdEdit_Click()
    editflagship = True
    addflagship = False
    Call edit
    
If UL = "" And DEP = 8 And HR = "SU" Then
    txtEmployeeNo.Enabled = True
Else
    txtEmployeeNo.Enabled = False
End If

txthfdep.Enabled = True
'makiss 09/29/04 *****************
'mskhpbdate.Enabled = True
mshpbdate.Enabled = True

txthprelate.Enabled = True
txtnickname.Enabled = True
txtsurname.Enabled = True
txtfirstname.Enabled = True
txtmiddlename.Enabled = True
txtfirstchoice.Enabled = True
'09/29/04 ******************
'mskbday.Enabled = True
mskbirthdate.Enabled = True

txtage.Enabled = True
txtaddress1.Enabled = True
cmbdepartment.Enabled = True
txthouseno.Enabled = True
txtstreet.Enabled = True
txtbarangay.Enabled = True
txtdistrict.Enabled = True
txtcity.Enabled = True
txtprovince.Enabled = True
txtprovinceaddress.Enabled = True
txtemergency.Enabled = True
txteraddress.Enabled = True
txtertel.Enabled = True
optMale.Enabled = True
optFemale.Enabled = True
txtheight.Enabled = True

txtweight.Enabled = True
txttel1.Enabled = True
txttel2.Enabled = True
txttel3.Enabled = True
txttel4.Enabled = True
cmbcivilstatus.Enabled = True
cmbempstatus.Enabled = True
' 09/29/04 **********************
'txtdateapplied.Enabled = True
' 11/03/04 ************
txtphilhealth.Enabled = True
' 11/03/04 ************
mskdateapplied.Enabled = True

txtappstatus.Enabled = True
txtexam.Enabled = True
txtinterview.Enabled = True
txtmedical.Enabled = True
msksss.Enabled = True
'txtphcode.Enabled = True
txtpagibig.Enabled = True
txtacctno.Enabled = True
txtcitizen.Enabled = True
txtbirthplace.Enabled = True
'txthaircolor.Enabled = True
'txteyecolor.Enabled = True
txtreligion.Enabled = True
msktax.Enabled = True
txtpostsch.Enabled = True
txtdatepost.Enabled = True
txtsecsch.Enabled = True
txtdatesec.Enabled = True
txtcollegesch.Enabled = True
txtdatecollege.Enabled = True
txtcollegecourse.Enabled = True
txtcollegemajor.Enabled = True
txtdatepost.Enabled = True
txtpostcourse.Enabled = True
txtprimarysch.Enabled = True
txtdateprimary.Enabled = True
txtfathername.Enabled = True
txtfaddress.Enabled = True
txtfjob.Enabled = True
txtmothername.Enabled = True
txtmaddress.Enabled = True
txtmjob.Enabled = True
txtspousename.Enabled = True
txtspouseaddress.Enabled = True
txtsjob.Enabled = True
txtnbi.Enabled = True
'txtrescer.Enabled = True
'txtissueat.Enabled = True
'dtissueon.Enabled = True
mskissue.Enabled = True

txtdep1.Enabled = True
txtdep2.Enabled = True
txtdep3.Enabled = True
txtdep4.Enabled = True
txtdep5.Enabled = True
mskdep1.Enabled = True
mskdep2.Enabled = True
mskdep3.Enabled = True
mskdep4.Enabled = True
mskdep5.Enabled = True
txtnumdep.Enabled = True
txtdialect1.Enabled = True
txtdialect2.Enabled = True
txtdialect3.Enabled = True
txtseminar1.Enabled = True
txtseminar2.Enabled = True
txtseminar3.Enabled = True
txtdateseminar1.Enabled = True
txtdateseminar2.Enabled = True
txtdateseminar3.Enabled = True
txtnameofcompany1.Enabled = True
txtnameofcompany2.Enabled = True
txtnameofcompany3.Enabled = True
txtfrom1.Enabled = True
txtfrom2.Enabled = True
txtfrom3.Enabled = True
txtto1.Enabled = True
txtto2.Enabled = True
txtto3.Enabled = True
txtpos1.Enabled = True
txtpos2.Enabled = True
txtpos3.Enabled = True
txtreason1.Enabled = True
txtreason2.Enabled = True
txtreason3.Enabled = True
txtsuperior1.Enabled = True
txtsuperior2.Enabled = True
txtsuperior3.Enabled = True
txtrefname1.Enabled = True
txtrefname2.Enabled = True
txtrefname3.Enabled = True
txtrefjob1.Enabled = True
txtrefjob2.Enabled = True
txtrefjob3.Enabled = True
txtrefaddress1.Enabled = True
txtrefaddress2.Enabled = True
txtrefaddress3.Enabled = True
txtrefphone1.Enabled = True
txtrefphone2.Enabled = True
txtrefphone3.Enabled = True
txtrefrel1.Enabled = True
txtrefrel2.Enabled = True
txtrefrel3.Enabled = True
' 09/29/04 *******************
'txtdatehired.Enabled = True
mskdatehired.Enabled = True
'dtdateprob.Enabled = True
mskdateprob.Enabled = True
'dtdatereg.Enabled = True
mskdatereg.Enabled = True
'dtdateresign.Enabled = True
mskdateres.Enabled = True

cmbempstatus.Enabled = True
txtsupervisor.Enabled = True
txtreasonresign.Enabled = True
'dtendcontract.Enabled = True
mskendof.Enabled = True
msdep1.Enabled = True
msdep2.Enabled = True
msdep3.Enabled = True
msdep4.Enabled = True
msdep5.Enabled = True
txtdtrexempt.Enabled = True
txthrswork.Enabled = True
txtnolate.Enabled = True
txtnoundertime.Enabled = True
txtActype.Enabled = True
txtnoot.Enabled = True
End Sub

Private Sub cmdfirst_Click()
rsout.MoveFirst
cmdfirst.Enabled = False
cmdprev.Enabled = False
cmdlast.Enabled = True
cmdnext.Enabled = True
showdata rsout
End Sub

Private Sub cmdlast_Click()
rsout.MoveLast
cmdlast.Enabled = False
cmdnext.Enabled = False
cmdprev.Enabled = True
cmdfirst.Enabled = True
showdata rsout
End Sub

Private Sub cmdnew_Click()
addflagship = True
editflagship = False

Call add


txtDtApp1.Visible = False
'txtdateapplied.Visible = True

txtDtBrth.Visible = False
'mskbday.Visible = True

txtIssOn.Visible = False
'dtissueon.Visible = True

'txtdatehired.Visible = True
txtDtHr.Visible = False

'dtdateprob.Visible = True
txtDtPr.Visible = False

'dtdatereg.Visible = True
txtDtReg.Visible = False

'dtdateresign.Visible = True
txtDtRes.Visible = False

'dtendcontract.Visible = True
txtECt.Visible = False

txtEmployeeNo.Enabled = True
txthfdep.Enabled = True
'makiss 09/29/04 ****************
'mskhpbdate.Enabled = True
mshpbdate.Enabled = True

txthprelate.Enabled = True
txtnickname.Enabled = True
txtsurname.Enabled = True
txtfirstname.Enabled = True
txtmiddlename.Enabled = True
txtfirstchoice.Enabled = True
' 09/29/04 *********************
'mskbday.Enabled = True
mskbirthdate.Enabled = True

txtage.Enabled = True
txtaddress1.Enabled = True
txthouseno.Enabled = True
txtstreet.Enabled = True
txtbarangay.Enabled = True
txtdistrict.Enabled = True
txtcity.Enabled = True
txtprovince.Enabled = True
txtprovinceaddress.Enabled = True
cmbdepartment.Enabled = True
txtemergency.Enabled = True
txteraddress.Enabled = True
txtertel.Enabled = True
optMale.Enabled = True
optFemale.Enabled = True
txtheight.Enabled = True
txtweight.Enabled = True
txttel1.Enabled = True
txttel2.Enabled = True
txttel3.Enabled = True
txttel4.Enabled = True
' 11/03/04 ************
txtphilhealth.Enabled = True
' 11/03/04 *************
cmbcivilstatus.Enabled = True
' 09/29/04 ********************
'txtdateapplied.Enabled = True
mskdateapplied.Enabled = True

txtappstatus.Enabled = True
txtexam.Enabled = True
txtinterview.Enabled = True
txtmedical.Enabled = True
msksss.Enabled = True
'txtphcode.Enabled = True
txtpagibig.Enabled = True
txtacctno.Enabled = True
txtcitizen.Enabled = True
txtbirthplace.Enabled = True
'txthaircolor.Enabled = True
'txteyecolor.Enabled = True
txtreligion.Enabled = True
msktax.Enabled = True
txtpostsch.Enabled = True
txtdatepost.Enabled = True
txtsecsch.Enabled = True
txtdatesec.Enabled = True
txtcollegesch.Enabled = True
txtdatecollege.Enabled = True
txtcollegecourse.Enabled = True
txtcollegemajor.Enabled = True
txtdatepost.Enabled = True
txtpostcourse.Enabled = True
txtprimarysch.Enabled = True
txtdateprimary.Enabled = True
txtfathername.Enabled = True
txtfaddress.Enabled = True
txtfjob.Enabled = True
txtmothername.Enabled = True
txtmaddress.Enabled = True
txtmjob.Enabled = True
txtspousename.Enabled = True
txtspouseaddress.Enabled = True
txtsjob.Enabled = True
txtnbi.Enabled = True
'txtrescer.Enabled = True
'txtissueat.Enabled = True
'dtissueon.Enabled = True
mskissue.Enabled = True

txtdep1.Enabled = True
txtdep2.Enabled = True
txtdep3.Enabled = True
txtdep4.Enabled = True
txtdep5.Enabled = True
'makiss 09/28/04**************
msdep1.Enabled = True
msdep2.Enabled = True
msdep3.Enabled = True
msdep4.Enabled = True
msdep5.Enabled = True
'mskdep1.Enabled = True
'mskdep2.Enabled = True
'mskdep3.Enabled = True
'mskdep4.Enabled = True
'mskdep5.Enabled = True
'******makiss 09/28/04********
txtnumdep.Enabled = True
txtdialect1.Enabled = True
txtdialect2.Enabled = True
txtdialect3.Enabled = True
txtseminar1.Enabled = True
txtseminar2.Enabled = True
txtseminar3.Enabled = True
txtdateseminar1.Enabled = True
txtdateseminar2.Enabled = True
txtdateseminar3.Enabled = True
txtnameofcompany1.Enabled = True
txtnameofcompany2.Enabled = True
txtnameofcompany3.Enabled = True
txtfrom1.Enabled = True
txtfrom2.Enabled = True
txtfrom3.Enabled = True
txtto1.Enabled = True
txtto2.Enabled = True
txtto3.Enabled = True
txtpos1.Enabled = True
txtpos2.Enabled = True
txtpos3.Enabled = True
txtreason1.Enabled = True
txtreason2.Enabled = True
txtreason3.Enabled = True
txtsuperior1.Enabled = True
txtsuperior2.Enabled = True
txtsuperior3.Enabled = True
txtrefname1.Enabled = True
txtrefname2.Enabled = True
txtrefname3.Enabled = True
txtrefjob1.Enabled = True
txtrefjob2.Enabled = True
txtrefjob3.Enabled = True
txtrefaddress1.Enabled = True
txtrefaddress2.Enabled = True
txtrefaddress3.Enabled = True
txtrefphone1.Enabled = True
txtrefphone2.Enabled = True
txtrefphone3.Enabled = True
txtrefrel1.Enabled = True
txtrefrel2.Enabled = True
txtrefrel3.Enabled = True
' 09/29/04 ************************
'txtdatehired.Enabled = True
mskdatehired.Enabled = True
'dtdateprob.Enabled = True
mskdateprob.Enabled = True
'dtdatereg.Enabled = True
mskdatereg.Enabled = True
'dtdateresign.Enabled = True
mskdateres.Enabled = True

cmbempstatus.Enabled = True
txtsupervisor.Enabled = True
txtreasonresign.Enabled = True
'dtendcontract.Enabled = True
mskendof.Enabled = True

txtdtrexempt.Enabled = True
txthrswork.Enabled = True
txtnolate.Enabled = True
txtnoundertime.Enabled = True
txtActype.Enabled = True
txtnoot.Enabled = True
'txtfirstchoice.SetFocus

Call clearscreen

End Sub

Private Sub cmdnext_Click()
On Error GoTo X
 rsout.MoveNext
 cmdprev.Enabled = True
 cmdfirst.Enabled = True
   If rsout.EOF Then
        rsout.MoveLast
        cmdnext.Enabled = False
        cmdlast.Enabled = False
        'MsgBox "End of file encountered", vbInformation, "DCPI Information System"
End If

    
    showdata rsout
X:
    Exit Sub
End Sub
Private Sub refreshappno()
On Error GoTo X
 rsout.MoveNext
 cmdprev.Enabled = True
 cmdfirst.Enabled = True
   If rsout.EOF Then
        rsout.MoveLast
        cmdnext.Enabled = False
        cmdlast.Enabled = False
'        MsgBox "End of file encountered", vbInformation, "DCPI Information System"
End If

    
    showdata rsout
X:
    Exit Sub

End Sub
Private Sub cmdprev_Click()
 On Error GoTo X
 rsout.MovePrevious
  cmdnext.Enabled = True
      cmdlast.Enabled = True
   If rsout.BOF Then
       rsout.MoveFirst
      cmdprev.Enabled = False
      cmdfirst.Enabled = False
 '     MsgBox "Beginning of file encountered", vbInformation, "DCPI Infromation System"
    End If
    
   employeenum = rsout!employeenum
    showdata rsout
X:
    Exit Sub
End Sub

Private Sub cmdsearch_Click()
frmHRSearch.Show vbModal
End Sub

Private Sub cmdundo_Click()
If addflag = True Then
        Unload Me
        Exit Sub
    End If

    addflagship = False
    editflagship = False
    showdata rsout
Call Cancel
End Sub
Private Sub Cancel()
cmdnew.Enabled = True
cmdedit.Enabled = True
cmdupdate.Enabled = False
cmdundo.Enabled = False
'cmdhired.Enabled = True
'cmdsearch.Enabled = True
cmddelete.Enabled = True
cmdfirst.Enabled = True
cmdprev.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True
'cmdExit.Enabled = True
txtEmployeeNo.Enabled = False
txthfdep.Enabled = False
cmbdepartment.Enabled = False
'makiss 09/29/04 *********
'mskhpbdate.Enabled = False
mshpbdate.Enabled = True

txthprelate.Enabled = False
txtprimarysch.Enabled = False
txtdateprimary.Enabled = False
txtnickname.Enabled = False
txtsurname.Enabled = False
txtfirstname.Enabled = False
txtmiddlename.Enabled = False
txtfirstchoice.Enabled = False
' 09/29/04 *****************
'mskbday.Enabled = False
mskbirthdate.Enabled = False

txtage.Enabled = False
txtaddress1.Enabled = False
txthouseno.Enabled = False
txtstreet.Enabled = False
txtbarangay.Enabled = False
txtdistrict.Enabled = False
txtcity.Enabled = False
txtprovince.Enabled = False
txtprovinceaddress.Enabled = False
txtemergency.Enabled = False
txteraddress.Enabled = False
txtertel.Enabled = False
optMale.Enabled = False
optFemale.Enabled = False
txtheight.Enabled = False
txtweight.Enabled = False
txttel1.Enabled = False
txttel2.Enabled = False
txttel3.Enabled = False
txttel4.Enabled = False
cmbcivilstatus.Enabled = False
' 09/29/04 ***************************
'txtdateapplied.Enabled = False
' 11/03/04 ************
txtphilhealth.Enabled = False
' 11/03/04 ************
mskdateapplied.Enabled = False

txtappstatus.Enabled = False
txtexam.Enabled = False
txtinterview.Enabled = False
txtmedical.Enabled = False
msksss.Enabled = False
'txtphcode.Enabled = False
txtpagibig.Enabled = False
txtacctno.Enabled = False
txtcitizen.Enabled = False
txtbirthplace.Enabled = False
'txthaircolor.Enabled = False
'txteyecolor.Enabled = False
txtreligion.Enabled = False
msktax.Enabled = False
txtpostsch.Enabled = False
txtdatepost.Enabled = False
txtsecsch.Enabled = False
txtdatesec.Enabled = False
txtcollegesch.Enabled = False
txtdatecollege.Enabled = False
txtcollegecourse.Enabled = False
txtcollegemajor.Enabled = False
txtdatepost.Enabled = False
txtpostcourse.Enabled = False
txtfathername.Enabled = False
txtfaddress.Enabled = False
txtfjob.Enabled = False
txtmothername.Enabled = False
txtmaddress.Enabled = False
txtmjob.Enabled = False
txtspousename.Enabled = False
txtspouseaddress.Enabled = False
txtsjob.Enabled = False
txtnbi.Enabled = False
'txtrescer.Enabled = False
'txtissueat.Enabled = False
'dtissueon.Enabled = False
mskissue.Enabled = False

txtdep1.Enabled = False
txtdep2.Enabled = False
txtdep3.Enabled = False
txtdep4.Enabled = False
txtdep5.Enabled = False
mskdep1.Enabled = False
mskdep2.Enabled = False
mskdep3.Enabled = False
mskdep4.Enabled = False
mskdep5.Enabled = False
txtnumdep.Enabled = False
txtdialect1.Enabled = False
txtdialect2.Enabled = False
txtdialect3.Enabled = False
txtseminar1.Enabled = False
txtseminar2.Enabled = False
txtseminar3.Enabled = False
txtdateseminar1.Enabled = False
txtdateseminar2.Enabled = False
txtdateseminar3.Enabled = False
txtnameofcompany1.Enabled = False
txtnameofcompany2.Enabled = False
txtnameofcompany3.Enabled = False
txtfrom1.Enabled = False
txtfrom2.Enabled = False
txtfrom3.Enabled = False
txtto1.Enabled = False
txtto2.Enabled = False
txtto3.Enabled = False
txtpos1.Enabled = False
txtpos2.Enabled = False
txtpos3.Enabled = False
txtreason1.Enabled = False
txtreason2.Enabled = False
txtreason3.Enabled = False
txtsuperior1.Enabled = False
txtsuperior2.Enabled = False
txtsuperior3.Enabled = False
txtrefname1.Enabled = False
txtrefname2.Enabled = False
txtrefname3.Enabled = False
txtrefjob1.Enabled = False
txtrefjob2.Enabled = False
txtrefjob3.Enabled = False
txtrefaddress1.Enabled = False
txtrefaddress2.Enabled = False
txtrefaddress3.Enabled = False
txtrefphone1.Enabled = False
txtrefphone2.Enabled = False
txtrefphone3.Enabled = False
txtrefrel1.Enabled = False
txtrefrel2.Enabled = False
txtrefrel3.Enabled = False
'09/29/04 *********************
'txtdatehired.Enabled = False
mskdatehired.Enabled = False
'dtdateprob.Enabled = False
mskdateprob.Enabled = False
'dtdatereg.Enabled = False
mskdatereg.Enabled = False
mskdateres.Enabled = False
'dtdateresign.Enabled = False

cmbempstatus.Enabled = False
txtsupervisor.Enabled = False
txtreasonresign.Enabled = False
'dtendcontract.Enabled = False
mskendof.Enabled = False

txtdtrexempt.Enabled = False
txthrswork.Enabled = False
txtnolate.Enabled = False
txtnoundertime.Enabled = False
txtActype.Enabled = False
txtnoot.Enabled = False

End Sub

Private Sub cmdupdate_Click()
If Trim(txtEmployeeNo.Text = "") Or (Val(txtEmployeeNo.Text) = 0) Then
    MsgBox "Employee Number Cannot Be Zero or Blank. Please Check!!!", vbCritical, "DCPI Information System"
    txtEmployeeNo.SetFocus
ElseIf Not IsDate(mskdateapplied.Text) Then
    MsgBox "Date Applied Cannot Be Blank. Please check!", vbCritical, "DCPI Information System"
    mskdateapplied.SetFocus
ElseIf Trim(txtfirstchoice.Text = "") Then
    MsgBox "Position Applied Cannot Be Blank. Please check!", vbCritical, "DCPI Information System"
    txtfirstchoice.SetFocus
ElseIf Trim(txtsurname.Text = "") Then
    MsgBox "Surname Cannot Be Blank. Please check!", vbCritical, "DCPI Information System"
    txtsurname.SetFocus
ElseIf Trim(txtfirstname.Text = "") Then
    MsgBox "FirstName Cannot Be Blank. Please check!", vbCritical, "DCPI Information System"
    txtfirstname.SetFocus
ElseIf Trim(txtmiddlename.Text = "") Then
    MsgBox "MiddleName Cannot Be Blank. Please check!", vbCritical, "DCPI Information System"
    txtmiddlename.SetFocus
ElseIf Trim(txtnickname.Text = "") Then
    MsgBox "Nickname Cannot Be Blank. Please check!", vbCritical, "DCPI Information System"
    txtnickname.SetFocus
Else
    txtEmployeeNo.Enabled = False
    txthfdep.Enabled = False
    
    '09/29/04 *****************************
    'mskhpbdate.Enabled = False
    mshpbdate.Enabled = False
    txthprelate.Enabled = False
    txtprimarysch.Enabled = False
    txtdateprimary.Enabled = False
    txtnickname.Enabled = False
    txtsurname.Enabled = False
    txtfirstname.Enabled = False
    txtmiddlename.Enabled = False
    txtfirstchoice.Enabled = False
    '09/29/04 *****************************
    'mskbday.Enabled = False
    mskbirthdate.Enabled = False
    ' 11/03/04 **********
    txtphilhealth.Enabled = False
    ' 11/03/04 **********
    txtage.Enabled = False
    txtaddress1.Enabled = False
    txthouseno.Enabled = False
    txtstreet.Enabled = False
    txtbarangay.Enabled = False
    txtdistrict.Enabled = False
    txtcity.Enabled = False
    txtprovince.Enabled = False
    txtprovinceaddress.Enabled = False
    txtemergency.Enabled = False
    txteraddress.Enabled = False
    txtertel.Enabled = False
    cmbdepartment.Enabled = False
    optMale.Enabled = False
    optFemale.Enabled = False
    txtheight.Enabled = False
    txtweight.Enabled = False
    txttel1.Enabled = False
    txttel2.Enabled = False
    txttel3.Enabled = False
    txttel4.Enabled = False
    cmbcivilstatus.Enabled = False
    ' 11/03/04 ************
    txtphilhealth.Enabled = False
    ' 11/03/04 ************
    '09/29/04 *************************
    'txtdateapplied.Enabled = False
    mskdateapplied.Enabled = False
    
    txtappstatus.Enabled = False
    txtexam.Enabled = False
    txtinterview.Enabled = False
    txtmedical.Enabled = False
    msksss.Enabled = False
    'txtphcode.Enabled = False
    txtpagibig.Enabled = False
    txtacctno.Enabled = False
    txtcitizen.Enabled = False
    txtbirthplace.Enabled = False
    'txthaircolor.Enabled = False
    'txteyecolor.Enabled = False
    txtreligion.Enabled = False
    msktax.Enabled = False
    txtpostsch.Enabled = False
    txtdatepost.Enabled = False
    txtsecsch.Enabled = False
    txtdatesec.Enabled = False
    txtcollegesch.Enabled = False
    txtdatecollege.Enabled = False
    txtcollegecourse.Enabled = False
    txtcollegemajor.Enabled = False
    txtdatepost.Enabled = False
    txtpostcourse.Enabled = False
    txtfathername.Enabled = False
    txtfaddress.Enabled = False
    txtfjob.Enabled = False
    txtmothername.Enabled = False
    txtmaddress.Enabled = False
    txtmjob.Enabled = False
    txtspousename.Enabled = False
    txtspouseaddress.Enabled = False
    txtsjob.Enabled = False
    txtnbi.Enabled = False
    'txtrescer.Enabled = False
    'txtissueat.Enabled = False
    'dtissueon.Enabled = False
    mskissue.Enabled = False
    
    txtdep1.Enabled = False
    txtdep2.Enabled = False
    txtdep3.Enabled = False
    txtdep4.Enabled = False
    txtdep5.Enabled = False
    mskdep1.Enabled = False
    mskdep2.Enabled = False
    mskdep3.Enabled = False
    mskdep4.Enabled = False
    mskdep5.Enabled = False
    txtnumdep.Enabled = False
    txtdialect1.Enabled = False
    txtdialect2.Enabled = False
    txtdialect3.Enabled = False
    txtseminar1.Enabled = False
    txtseminar2.Enabled = False
    txtseminar3.Enabled = False
    txtdateseminar1.Enabled = False
    txtdateseminar2.Enabled = False
    txtdateseminar3.Enabled = False
    txtnameofcompany1.Enabled = False
    txtnameofcompany2.Enabled = False
    txtnameofcompany3.Enabled = False
    txtfrom1.Enabled = False
    txtfrom2.Enabled = False
    txtfrom3.Enabled = False
    txtto1.Enabled = False
    txtto2.Enabled = False
    txtto3.Enabled = False
    txtpos1.Enabled = False
    txtpos2.Enabled = False
    txtpos3.Enabled = False
    txtreason1.Enabled = False
    txtreason2.Enabled = False
    txtreason3.Enabled = False
    txtsuperior1.Enabled = False
    txtsuperior2.Enabled = False
    txtsuperior3.Enabled = False
    txtrefname1.Enabled = False
    txtrefname2.Enabled = False
    txtrefname3.Enabled = False
    txtrefjob1.Enabled = False
    txtrefjob2.Enabled = False
    txtrefjob3.Enabled = False
    txtrefaddress1.Enabled = False
    txtrefaddress2.Enabled = False
    txtrefaddress3.Enabled = False
    txtrefphone1.Enabled = False
    txtrefphone2.Enabled = False
    txtrefphone3.Enabled = False
    txtrefrel1.Enabled = False
    txtrefrel2.Enabled = False
    txtrefrel3.Enabled = False
    ' 09/29/04 ***********************
    'txtdatehired.Enabled = False
    mskdatehired.Enabled = False
    'dtdateprob.Enabled = False
    mskdateprob.Enabled = False
    'dtdatereg.Enabled = False
    mskdatereg.Enabled = False
    mskdateres.Enabled = False
    'dtdateresign.Enabled = False
    
    cmbempstatus.Enabled = False
    txtsupervisor.Enabled = False
    txtreasonresign.Enabled = False
    'dtendcontract.Enabled = False
    mskendof.Enabled = False
    
    txtdtrexempt.Enabled = False
    txthrswork.Enabled = False
    txtnolate.Enabled = False
    txtnoundertime.Enabled = False
    txtActype.Enabled = False
    txtnoot.Enabled = False
    
    ' 09/29/04 *************
    '       bdt1 = IIf(IsNull(mskbday.Value), "01/01/1900", mskbday.Value)
            bdt1 = IIf(mskbirthdate = "__/__/____", Null, Format(mskbirthdate, "mm/dd/yyyy"))
            
            'bdt = IIf(bdt1 = "__/__/__", "01/01/1900", bdt1)
            'appdt = IIf(IsNull(txtdateapplied), "01/01/1900", Format(txtdateapplied, "m/d/yyyy"))
            appdt = IIf(mskdateapplied = "__/__/____", Null, Format(mskdateapplied, "mm/dd/yyyy"))
            'hiredt = IIf(IsNull(txtdatehired), "01/01/1900", txtdatehired)
            hiredt = IIf(mskdatehired = "__/__/____", Null, Format(mskdatehired, "mm/dd/yyyy"))
            'probdt = IIf(IsNull(dtdateprob.Value), Null, dtdateprob.Value)
            'probdt = IIf(IsNull(mskdateprob) Or (mskdateprob) = "", "__/__/____", Format(mskdateprob, "mm/dd/yyyy"))
            probdt = IIf(mskdateprob = "__/__/____", Null, Format(mskdateprob, "mm/dd/yyyy"))
            'regdt = IIf(IsNull(dtdatereg.Value), "01/01/1900", dtdatereg.Value)
            'regdt = IIf(IsNull(dtdatereg.Value), Null, dtdatereg.Value)
            regdt = IIf(mskdatereg = "__/__/____", Null, Format(mskdatereg, "mm/dd/yyyy"))
            'resdt = IIf(IsNull(dtdateresign.Value), Null, dtdateresign.Value)
            resdt = IIf(mskdateres = "__/__/____", Null, Format(mskdateres, "mm/dd/yyyy"))
            'eocdt = IIf(IsNull(dtendcontract.Value), "01/01/1900", dtendcontract.Value)
            eocdt = IIf(mskendof = "__/__/____", Null, Format(mskendof, "mm/dd/yyyy"))
            'isson = IIf(IsNull(dtissueon.Value), "01/01/1900", dtissueon.Value)
            isson = IIf(IsDate(mskissue), Format(mskissue, "mm/dd/yyyy"), "01/01/1900")
            'isson = IIf(mskissue = "__/__/____", "01/01/1900", Format(mskissue, "mm/dd/yyyy"))
            gendr = IIf(optMale.Value, "M", "F")
            civst = Left(cmbcivilstatus.Text, 2)
            empst = Left(cmbempstatus.Text, 1)
            usrid = UN
            txage = IIf(IsNull(txtage), 0, txtage)
            'dep1dt = IIf(IsNull(mskdep1.Value), "01/01/1900", mskdep1.Value)
            'dep1dt = IIf(IsNull(mskdep1.Value), "01/01/1900", Format(mskdep1.Value, "m/d/yyyy"))
            dep2dt = IIf(IsNull(mskdep2.Value), "01/01/1900", mskdep2.Value)
            dep3dt = IIf(IsNull(mskdep3.Value), "01/01/1900", mskdep3.Value)
            dep4dt = IIf(IsNull(mskdep4.Value), "01/01/1900", mskdep4.Value)
            dep5dt = IIf(IsNull(mskdep5.Value), "01/01/1900", mskdep5.Value)
    'dep1dt = IIf(mskdep1.Value = "__/__/__", "", mskdep1.Value)
    'dep2dt = IIf(mskdep2.Value = "__/__/__", "", mskdep2.Value)
    'dep3dt = IIf(mskdep3.Value = "__/__/__", "", mskdep3.Value)
    'dep4dt = IIf(mskdep4.Value = "__/__/__", "", mskdep4.Value)
    'dep5dt = IIf(mskdep5.Value = "__/__/__", "", mskdep5.Value)
    'numdep = IIf(IsNull(txtnumdep.Text), 0, txtnumdep.Text)
    
    If addflag = True Or addflagship = True Then
    
               Call save
               Call saveclick
               addflag = False
               addflagship = False
             End If
            
            If editflag = True Or editflagship = True Then
                Call edit
                Call editsave
            End If
            'Call save
            Call update
            
            addflagship = False
            editflagship = False

End If

End Sub

Private Sub editsave()
Dim rstemp As ADODB.Recordset
Dim str As String
Dim strdialect As String
       
       If cmbdepartment.Text = "" Then
            depcd = 0
        Else
            depcd = Me.cmbdepartment.ItemData(Me.cmbdepartment.ListIndex)
        End If
'ElseIf Not IsDate(mskdateprob.Text) Then
'    mskdateprob.Text = "__/__/____"

'Saving Hr Employees
cnDCPI.Execute "EXEC UpdHREmployees '" & txtEmployeeNo & "', '" & _
txtfirstname & "', '" & txtmiddlename & "', '" & txtsurname & "', '" & _
txtnickname & "', '" & txtfirstchoice & "', " & txtage & ", '" & _
IIf(IsDate(appdt), Format(appdt, "mm/dd/yyyy"), Null) & "', '" & txtappstatus & "', '" & IIf(IsDate(bdt1), Format(bdt1, "mm/dd/yyyy"), Null) & "', '" & IIf(InStr(txtaddress1, "'"), Replace(txtaddress1, "'", "''"), txtaddress1) & "', '" & _
txthouseno & "', '" & IIf(InStr(txtstreet, "'"), Replace(txtstreet, "'", "''"), txtstreet) & "', '" & txtbarangay & "', '" & txtdistrict & "', '" & _
txtcity & "', '" & txtprovince & "', '" & IIf(InStr(txtprovinceaddress, "'"), Replace(txtprovinceaddress, "'", "''"), txtprovinceaddress) & "', '" & _
IIf(InStr(txtemergency, "'"), Replace(txtemergency, "'", "''"), txtemergency) & "', '" & txtertel & "', '" & IIf(InStr(txteraddress, "'"), Replace(txteraddress, "'", "''"), txteraddress) & "', '" & _
gendr & "', '" & IIf(InStr(txtweight, "'"), Replace(txtweight, "'", "''"), txtweight) & "', '" & IIf(InStr(txtheight, "'") <> 0, Replace(txtheight, "'", "''"), txtheight) & "', '" & txtexam & "', '" & _
txtinterview & "', '" & txttel1 & "', '" & txttel2 & "', '" & txttel3 & "', '" & _
txttel4 & "', '" & civst & "', '" & IIf(IsDate(hiredt), Format(hiredt, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(resdt), Format(resdt, "mm/dd/yyyy"), Null) & "', '" & msksss & "', '" & _
txtpagibig & "', '" & empst & "', '" & txtacctno & "', '" & _
txtcitizen & "', '" & txtbirthplace & "', '" & _
IIf(InStr(txtreligion, "'"), Replace(txtreligion, "'", "''"), txtreligion) & "', '" & IIf(IsDate(regdt), Format(regdt, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(probdt), Format(probdt, "mm/dd/yyyy"), Null) & "', '" & msktax & "', '" & txtmedical & "', '" & _
IIf(InStr(txtprimarysch, "'"), Replace(txtprimarysch, "'", "''"), txtprimarysch) & "', '" & txtdateprimary & "', '" & IIf(InStr(txtsecsch, "'"), Replace(txtsecsch, "'", "''"), txtsecsch) & "', '" & txtdatesec & "', '" & _
IIf(InStr(txtcollegesch, "'"), Replace(txtcollegesch, "'", "''"), txtcollegesch) & "', '" & txtdatecollege & "', '" & IIf(InStr(txtcollegecourse, "'"), Replace(txtcollegecourse, "'", "''"), txtcollegecourse) & "', '" & txtcollegemajor & "', '" & _
IIf(InStr(txtpostsch, "'"), Replace(txtpostsch, "'", "''"), txtpostsch) & "', '" & txtdatepost & "', '" & IIf(InStr(txtpostcourse, "'"), Replace(txtpostcourse, "'", "''"), txtpostcourse) & "', '" & txtfathername & "', '" & _
IIf(InStr(txtfaddress, ",") <> 0, Replace(txtfaddress, ",", ""), txtfaddress) & "', '" & IIf(InStr(txtfjob, "'"), Replace(txtfjob, "'", "''"), txtfjob) & "', '" & IIf(InStr(txtmothername, "'"), Replace(txtmothername, "'", "''"), txtmothername) & "', '" & IIf(InStr(txtmaddress, "'") <> 0, Replace(txtmaddress, "'", "''"), txtmaddress) & "', '" & _
IIf(InStr(txtmjob, "'"), Replace(txtmjob, "'", "''"), txtmjob) & "', '" & txtnbi & "', '" & IIf(IsDate(isson), Format(isson, "mm/dd/yyyy"), Null) & "', '" & _
txtsupervisor & "', '" & txtreasonresign & "', '" & txtdtrexempt & "', '" & IIf(IsDate(eocdt), Format(eocdt, "mm/dd/yyyy"), Null) & "', '" & _
txthrswork & "', '" & txtnolate & "', '" & txtnoundertime & "', '" & txtActype & "', '" & _
txtnoot & "', " & depcd & ", '" & usrid & "', '" & txtphilhealth & "'"



'09/30/04*******************
'Saving HR Details

Set rstemp = cnDCPI.Execute("Select * from HRDetails where convert(int, Employeenum) = '" & Val(txtEmployeeNo.Text) & "'")
If rstemp.EOF = True Then
    cnDCPI.Execute "EXEC INSHRDetails '" & txtEmployeeNo & "', '" & _
    txtspousename & "', '" & txtspouseaddress & "', '" & txtsjob & "', '" & _
    txtnumdep & "', '" & txtdep1 & "', '" & txtdep2 & "', '" & txtdep3 & "', '" & _
    txtdep4 & "', '" & txtdep5 & "', '" & IIf(IsDate(msdep1.Text), Format(msdep1.Text, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(msdep2.Text), Format(msdep2.Text, "mm/dd/yyyy"), Null) & "', '" & _
    IIf(IsDate(msdep3.Text), Format(msdep3.Text, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(msdep4.Text), Format(msdep4.Text, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(msdep5.Text), Format(msdep5.Text, "mm/dd/yyyy"), Null) & "', '" & txthfdep & "', '" & _
    IIf(IsDate(mshpbdate.Text), Format(mshpbdate.Text, "mm/dd/yyyy"), Null) & "', '" & txthprelate & "'"
Else
    cnDCPI.Execute "EXEC UpdHRDetails '" & txtEmployeeNo & "', '" & _
    txtspousename & "', '" & txtspouseaddress & "', '" & IIf(InStr(txtsjob, "'"), Replace(txtsjob, "'", "''"), txtsjob) & "', '" & _
    txtnumdep & "', '" & txtdep1 & "', '" & txtdep2 & "', '" & txtdep3 & "', '" & _
    txtdep4 & "', '" & txtdep5 & "', '" & IIf(IsDate(msdep1.Text), Format(msdep1.Text, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(msdep2.Text), Format(msdep2.Text, "mm/dd/yyyy"), Null) & "', '" & _
    IIf(IsDate(msdep3.Text), Format(msdep3.Text, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(msdep4.Text), Format(msdep4.Text, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(msdep5.Text), Format(msdep5.Text, "mm/dd/yyyy"), Null) & "', '" & txthfdep & "', '" & _
    IIf(IsDate(mshpbdate.Text), Format(mshpbdate.Text, "mm/dd/yyyy"), Null) & "', '" & txthprelate & "'"
End If
rstemp.Close
Set rstemp = Nothing

'saving Character References
Set rstemp = cnDCPI.Execute("Select * from HRCharReference where convert(int, Employeenum) = '" & Val(txtEmployeeNo.Text) & "'")
If rstemp.EOF = True Then
    cnDCPI.Execute "EXEC INSHRCharRef '" & txtEmployeeNo & "', '" & _
    IIf(InStr(txtrefname1, "'"), Replace(txtrefname1, "'", "''"), txtrefname1) & "', '" & txtrefjob1 & "', '" & txtrefaddress1 & "', '" & _
    txtrefphone1 & "', '" & txtrefrel1 & "', '" & txtrefname2 & "', '" & _
    txtrefjob2 & "', '" & txtrefaddress2 & "', '" & txtrefphone2 & "', '" & _
    txtrefrel2 & "', '" & txtrefname3 & "', '" & txtrefaddress3 & "', '" & _
    txtrefjob3 & "', '" & txtrefphone3 & "', '" & txtrefrel3 & "'"
Else
    cnDCPI.Execute "EXEC UpdHRCharRef '" & txtEmployeeNo & "', '" & _
    IIf(InStr(txtrefname1, "'"), Replace(txtrefname1, "'", "''"), txtrefname1) & "', '" & txtrefjob1 & "', '" & IIf(InStr(txtrefaddress1, "'"), Replace(txtrefaddress1, "'", "''"), txtrefaddress1) & "', '" & _
    txtrefphone1 & "', '" & txtrefrel1 & "', '" & IIf(InStr(txtrefname2, "'"), Replace(txtrefname2, "'", "''"), txtrefname2) & "', '" & _
    IIf(InStr(txtrefjob2, "'") <> 0, Replace(txtrefjob2, "'", "''"), txtrefjob2) & "', '" & IIf(InStr(txtrefaddress2, "'"), Replace(txtrefaddress2, "'", "''"), txtrefaddress2) & "', '" & txtrefphone2 & "', '" & _
    txtrefrel2 & "', '" & IIf(InStr(txtrefname3, "'"), Replace(txtrefname3, "'", "''"), txtrefname3) & "', '" & IIf(InStr(txtrefaddress3, "'") <> 0, Replace(txtrefaddress3, "'", "''"), txtrefaddress3) & "', '" & _
    txtrefjob3 & "', '" & txtrefphone3 & "', '" & IIf(InStr(txtrefrel3, "'"), Replace(txtrefrel3, "'", "''"), txtrefrel3) & "'"
End If
rstemp.Close
Set rstemp = Nothing

'Saving Dialects
Set rstemp = cnDCPI.Execute("Select * from HRDialects where convert(int, Employeenum) = '" & Val(txtEmployeeNo.Text) & "'")
If rstemp.EOF = True Then
    cnDCPI.Execute "EXEC INSHRDialects '" & txtEmployeeNo & "', '" & _
    txtdialect1 & "', '" & txtdialect2 & "', '" & txtdialect3 & "'"
Else
    cnDCPI.Execute "EXEC UpdHRDialects '" & txtEmployeeNo & "', '" & _
    txtdialect1 & "', '" & txtdialect2 & "', '" & txtdialect3 & "'"
End If
rstemp.Close
Set rstemp = Nothing

'Saving HR Employee Records
Set rstemp = cnDCPI.Execute("Select * from HREmploymentRecord where convert(int, Employeenum) = '" & Val(txtEmployeeNo.Text) & "'")
If rstemp.EOF = True Then
    cnDCPI.Execute "EXEC INSHREmpRec '" & txtEmployeeNo & "', '" & _
    txtnameofcompany1 & "', '" & txtfrom1 & "', '" & txtto1 & "', '" & _
    txtpos1 & "', '" & txtreason1 & "', '" & txtsuperior1 & "', '" & _
    txtnameofcompany2 & "', '" & txtfrom2 & "', '" & txtto2 & "', '" & _
    txtpos2 & "', '" & txtreason2 & "', '" & txtsuperior2 & "', '" & _
    txtnameofcompany3 & "', '" & txtfrom3 & "', '" & txtto3 & "', '" & _
    txtpos3 & "', '" & txtreason3 & "', '" & txtsuperior3 & "'"
Else
    cnDCPI.Execute "EXEC UpdHREmpRec '" & txtEmployeeNo & "', '" & _
    IIf(InStr(txtnameofcompany1, "'"), Replace(txtnameofcompany1, "'", "''"), txtnameofcompany1) & "', '" & txtfrom1 & "', '" & txtto1 & "', '" & _
    IIf(InStr(txtpos1, "'"), Replace(txtpos1, "'", "''"), txtpos1) & "', '" & IIf(InStr(txtreason1, "'"), Replace(txtreason1, "'", "''"), txtreason1) & "', '" & txtsuperior1 & "', '" & _
    IIf(InStr(txtnameofcompany2, "'"), Replace(txtnameofcompany2, "'", "''"), txtnameofcompany2) & "', '" & txtfrom2 & "', '" & txtto2 & "', '" & _
    txtpos2 & "', '" & txtreason2 & "', '" & IIf(InStr(txtsuperior2, "'"), Replace(txtsuperior2, "'", "''"), txtsuperior2) & "', '" & _
    IIf(InStr(txtnameofcompany3, "'"), Replace(txtnameofcompany3, "'", "''"), txtnameofcompany3) & "', '" & txtfrom3 & "', '" & txtto3 & "', '" & _
    IIf(InStr(txtpos3, "'"), Replace(txtpos3, "'", "''"), txtpos3) & "', '" & txtreason3 & "', '" & txtsuperior3 & "'"
End If
rstemp.Close
Set rstemp = Nothing


'Saving Hr Seminar
Set rstemp = cnDCPI.Execute("Select * from HRSeminar where convert(int, Employeenum) = '" & Val(txtEmployeeNo.Text) & "'")
If rstemp.EOF = True Then
    cnDCPI.Execute "EXEC INSHRSeminar '" & txtEmployeeNo & "', '" & _
    IIf(InStr(txtseminar1, "'"), Replace(txtseminar1, "'", "''"), txtseminar1) & "', '" & txtdateseminar1 & "', '" & _
    IIf(InStr(txtseminar2, "'"), Replace(txtseminar2, "'", "''"), txtseminar2) & "', '" & txtdateseminar2 & "', '" & _
    IIf(InStr(txtseminar3, "'"), Replace(txtseminar2, "'", "''"), txtseminar3) & "', '" & txtdateseminar3 & "'"
Else
    cnDCPI.Execute "EXEC UpdHRSeminar '" & txtEmployeeNo & "', '" & _
    IIf(InStr(txtseminar1, "'"), Replace(txtseminar1, "'", "''"), txtseminar1) & "', '" & txtdateseminar1 & "', '" & _
    IIf(InStr(txtseminar2, "'"), Replace(txtseminar2, "'", "''"), txtseminar2) & "', '" & txtdateseminar2 & "', '" & _
    IIf(InStr(txtseminar3, "'"), Replace(txtseminar3, "'", "''"), txtseminar3) & "', '" & txtdateseminar3 & "'"
End If
rstemp.Close
Set rstemp = Nothing


rsout.Requery
rsdetails.Requery
rsdialect.Requery
rsemprec.Requery
rsref.Requery
rssem.Requery

'rsout.MoveLast
'rsdetails.MoveLast
'rsdialect.MoveLast
'rsemprec.MoveLast
'rsref.MoveLast
'rssem.MoveLast


'09/28/04
Call Form_Load


'Showdata rsout

'Call editdetails
'Call editChar
'Call editdialect
'Call editemprec
'Call editsem

End Sub

Private Sub saveclick()
      
        If Trim(cmbdepartment.Text) = "" Then
          depcd = 0
      Else
        If Me.cmbdepartment.ListIndex = -1 Then
            depcd = Deptcode
        Else
            depcd = Me.cmbdepartment.ItemData(Me.cmbdepartment.ListIndex)
        End If
         End If
'Saving HR Employees
cnDCPI.Execute "EXEC AddHREmployees '" & txtEmployeeNo & "', '" & _
txtfirstname & "', '" & txtmiddlename & "', '" & txtsurname & "', '" & _
txtnickname & "', '" & txtfirstchoice & "', " & txage & ", '" & _
IIf(IsDate(appdt), Format(appdt, "mm/dd/yyyy"), Null) & "', '" & txtappstatus & "', '" & IIf(IsDate(bdt1), Format(bdt1, "mm/dd/yyyy"), Null) & "', '" & txtaddress1 & "', '" & _
txthouseno & "', '" & txtstreet & "', '" & txtbarangay & "', '" & txtdistrict & "', '" & _
txtcity & "', '" & txtprovince & "', '" & txtprovinceaddress & "', '" & _
txtemergency & "', '" & txtertel & "', '" & txteraddress & "', '" & _
gendr & "', '" & txtweight & "', '" & IIf(InStr(txtheight, "'") <> 0, Replace(txtheight, "'", "''"), txtheight) & "', '" & txtexam & "', '" & _
txtinterview & "', '" & txttel1 & "', '" & txttel2 & "', '" & txttel3 & "', '" & _
txttel4 & "', '" & civst & "', '" & IIf(IsDate(hiredt), Format(hiredt, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(resdt), Format(resdt, "mm/dd/yyyy"), Null) & "', '" & msksss & "', '" & _
txtpagibig & "', '" & empst & "', '" & txtacctno & "', '" & _
txtcitizen & "', '" & txtbirthplace & "', '" & _
txtreligion & "', '" & IIf(IsDate(regdt), Format(regdt, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(probdt), Format(probdt, "mm/dd/yyyy"), Null) & "', '" & msktax & "', '" & txtmedical & "', '" & _
IIf(InStr(txtprimarysch, "'"), Replace(txtprimarysch, "'", "''"), txtprimarysch) & "', '" & txtdateprimary & "', '" & IIf(InStr(txtsecsch, "'"), Replace(txtsecsch, "'", "''"), txtsecsch) & "', '" & txtdatesec & "', '" & _
IIf(InStr(txtcollegesch, "'"), Replace(txtcollegesch, "'", "''"), txtcollegesch) & "', '" & txtdatecollege & "', '" & txtcollegecourse & "', '" & txtcollegemajor & "', '" & _
IIf(InStr(txtpostsch, "'"), Replace(txtpostsch, "'", "''"), txtpostsch) & "', '" & txtdatepost & "', '" & IIf(InStr(txtpostcourse, "'"), Replace(txtpostcourse, "'", "''"), txtpostcourse) & "', '" & txtfathername & "', '" & _
txtfaddress & "', '" & txtfjob & "', '" & txtmothername & "', '" & txtmaddress & "', '" & _
txtmjob & "', '" & txtnbi & "', '" & isson & "', '" & _
txtsupervisor & "', '" & txtreasonresign & "', '" & txtdtrexempt & "', '" & IIf(IsDate(eocdt), Format(eocdt, "mm/dd/yyyy"), Null) & "', '" & _
txthrswork & "', '" & txtnolate & "', '" & txtnoundertime & "', '" & txtActype & "', '" & _
txtnoot & "', " & depcd & ", '" & usrid & "', '" & txtphilhealth & "'"

rsout.Requery
rsout.MoveLast

'Saving HR Details
cnDCPI.Execute "EXEC AddHRDetails '" & txtEmployeeNo & "', '" & _
txtspousename & "', '" & txtspouseaddress & "', '" & txtsjob & "', " & _
numdep & ", '" & txtdep1 & "', '" & txtdep2 & "', '" & txtdep3 & "', '" & _
txtdep4 & "', '" & txtdep5 & "', '" & IIf(IsDate(msdep1.Text), Format(msdep1.Text, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(msdep2.Text), Format(msdep2.Text, "mm/dd/yyyy"), Null) & "', '" & _
IIf(IsDate(msdep3.Text), Format(msdep3.Text, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(msdep4.Text), Format(msdep4.Text, "mm/dd/yyyy"), Null) & "', '" & IIf(IsDate(msdep5.Text), Format(msdep5.Text, "mm/dd/yyy"), Null) & "', '" & txthfdep & "', '" & _
IIf(IsDate(mshpbdate.Text), Format(mshpbdate.Text, "mm/dd/yyyy"), Null) & "', '" & txthprelate & "'"

rsdetails.Requery
rsdetails.MoveLast

'Saving Character Reference
cnDCPI.Execute "EXEC AddHRCharRef '" & txtEmployeeNo & "', '" & _
txtrefname1 & "', '" & txtrefjob1 & "', '" & txtrefaddress1 & "', '" & _
txtrefphone1 & "', '" & txtrefrel1 & "', '" & txtrefname2 & "', '" & _
txtrefjob2 & "', '" & txtrefaddress2 & "', '" & txtrefphone2 & "', '" & _
txtrefrel2 & "', '" & txtrefname3 & "', '" & txtrefaddress3 & "', '" & _
txtrefjob3 & "', '" & txtrefphone3 & "', '" & txtrefrel3 & "'"

rsref.Requery
rsref.MoveLast

'Saving HR Dialects
cnDCPI.Execute "EXEC AddHRDialects '" & txtEmployeeNo & "', '" & _
txtdialect1 & "', '" & txtdialect2 & "', '" & txtdialect3 & "'"

rsdialect.Requery
rsdialect.MoveLast


'Saving Hr Employee Record
cnDCPI.Execute "EXEC AddHREmpRec '" & txtEmployeeNo & "', '" & _
txtnameofcompany1 & "', '" & txtfrom1 & "', '" & txtto1 & "', '" & _
txtpos1 & "', '" & txtreason1 & "', '" & txtsuperior1 & "', '" & _
txtnameofcompany2 & "', '" & txtfrom2 & "', '" & txtto2 & "', '" & _
txtpos2 & "', '" & txtreason2 & "', '" & txtsuperior2 & "', '" & _
txtnameofcompany3 & "', '" & txtfrom3 & "', '" & txtto3 & "', '" & _
txtpos3 & "', '" & txtreason3 & "', '" & txtsuperior3 & "'"

rsemprec.Requery
rsemprec.MoveLast

'Saving HR Seminar
cnDCPI.Execute "EXEC AddHRSeminar '" & txtEmployeeNo & "', '" & _
IIf(InStr(txtseminar1, "'"), Replace(txtseminar1, "'", "''"), txtseminar1) & "', '" & txtdateseminar1 & "', '" & _
IIf(InStr(txtseminar2, "'"), Replace(txtseminar2, "'", "''"), txtseminar2) & "', '" & txtdateseminar2 & "', '" & _
IIf(InStr(txtseminar3, "'"), Replace(txtseminar3, "'", "''"), txtseminar3) & "', '" & txtdateseminar3 & "'"

rssem.Requery
rssem.MoveLast

End Sub
Public Function output() As String
    Dim rs As ADODB.Recordset
    Dim str As String

    str = "Select max(employeenum) From HREmployees"
    Set rs = cnDCPI.Execute(str)
    If Not rs.EOF Then
         output = Format(rs.Fields(0) + 1, "000000")
    Else
         output = "000001"
    End If
End Function
Private Sub save()
cmdnew.Enabled = True
cmdedit.Enabled = True
cmdupdate.Enabled = False
'cmdsearch.Enabled = True
'cmdhired.Enabled = True
cmddelete.Enabled = True
cmdfirst.Enabled = True
cmdprev.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True
cmdundo.Enabled = False
End Sub
Private Sub TabResizing()
If SSTab1.Tab = 0 Then
    SSTab1.Height = 7000
   ' frmHRApplication.Height = 9000
ElseIf SSTab1.Tab = 1 Then
    SSTab1.Height = 7000
'    frmHRApplication.Height = 9000
ElseIf SSTab1.Tab = 2 Then
    SSTab1.Height = 4500
 '   frmHRApplication.Height = 7000
'ElseIf SSTab1.Tab = 3 Then
'    SSTab1.Height = 6615
'    frmHRApplication.Height = 9000
End If

End Sub

Private Sub Form_Load()



Dim rsdept As ADODB.Recordset
Set rsout = New ADODB.Recordset
Set rsdetails = New ADODB.Recordset
Set rsref = New ADODB.Recordset
Set rsdialect = New ADODB.Recordset
Set rsemprec = New ADODB.Recordset
Set rscivil = New ADODB.Recordset
Set rssem = New ADODB.Recordset
Set rsempstatus = New ADODB.Recordset
Call CheckUser
Call TabResizing

'addflag = False
'editflag = False

txtEmployeeNo.Enabled = False
txthfdep.Enabled = False
'09/29/04 *****************************
'mskhpbdate.Enabled = False
mshpbdate.Enabled = False

txthprelate.Enabled = False
txtprimarysch.Enabled = False
txtdateprimary.Enabled = False
txtnickname.Enabled = False
txtsurname.Enabled = False
txtfirstname.Enabled = False
txtmiddlename.Enabled = False
txtfirstchoice.Enabled = False
' 09/29/04 *******************
'mskbday.Enabled = False
mskbirthdate.Enabled = False
' 11/03/04 **********
txtphilhealth.Enabled = False
' 11/03/04 **********
txtage.Enabled = False
txtaddress1.Enabled = False
txthouseno.Enabled = False
txtstreet.Enabled = False
txtbarangay.Enabled = False
txtdistrict.Enabled = False
txtcity.Enabled = False
txtprovince.Enabled = False
txtprovinceaddress.Enabled = False
txtemergency.Enabled = False
txteraddress.Enabled = False
txtertel.Enabled = False
optMale.Enabled = False
optFemale.Enabled = False
txtheight.Enabled = False
txtweight.Enabled = False
txttel1.Enabled = False
txttel2.Enabled = False
txttel3.Enabled = False
txttel4.Enabled = False
cmbcivilstatus.Enabled = False
'DTdateapplied.Enabled = False
' 09/29/04 **********************
'txtdateapplied.Enabled = False
mskdateapplied.Enabled = False

txtappstatus.Enabled = False
txtexam.Enabled = False
txtinterview.Enabled = False
txtmedical.Enabled = False
msksss.Enabled = False
'txtphcode.Enabled = False
txtpagibig.Enabled = False
txtacctno.Enabled = False
txtcitizen.Enabled = False
txtbirthplace.Enabled = False
'txthaircolor.Enabled = False
'txteyecolor.Enabled = False
txtreligion.Enabled = False
msktax.Enabled = False
txtpostsch.Enabled = False
txtdatepost.Enabled = False
txtsecsch.Enabled = False
txtdatesec.Enabled = False
txtcollegesch.Enabled = False
txtdatecollege.Enabled = False
txtcollegecourse.Enabled = False
txtcollegemajor.Enabled = False
txtdatepost.Enabled = False
txtpostcourse.Enabled = False
txtfathername.Enabled = False
txtfaddress.Enabled = False
txtfjob.Enabled = False
txtmothername.Enabled = False
txtmaddress.Enabled = False
txtmjob.Enabled = False
txtspousename.Enabled = False
txtspouseaddress.Enabled = False
txtsjob.Enabled = False
txtnbi.Enabled = False
'txtrescer.Enabled = False
'txtissueat.Enabled = False
'dtissueon.Enabled = False
mskissue.Enabled = False

txtdep1.Enabled = False
txtdep2.Enabled = False
txtdep3.Enabled = False
txtdep4.Enabled = False
txtdep5.Enabled = False
mskdep1.Enabled = False
mskdep2.Enabled = False
mskdep3.Enabled = False
mskdep4.Enabled = False
mskdep5.Enabled = False
msdep1.Enabled = False
msdep2.Enabled = False
msdep3.Enabled = False
msdep4.Enabled = False
msdep5.Enabled = False
mshpbdate.Enabled = False
txtnumdep.Enabled = False
txtdialect1.Enabled = False
txtdialect2.Enabled = False
txtdialect3.Enabled = False
txtseminar1.Enabled = False
txtseminar2.Enabled = False
txtseminar3.Enabled = False
txtdateseminar1.Enabled = False
txtdateseminar2.Enabled = False
txtdateseminar3.Enabled = False
txtnameofcompany1.Enabled = False
txtnameofcompany2.Enabled = False
txtnameofcompany3.Enabled = False
txtfrom1.Enabled = False
txtfrom2.Enabled = False
txtfrom3.Enabled = False
txtto1.Enabled = False
txtto2.Enabled = False
txtto3.Enabled = False
txtpos1.Enabled = False
txtpos2.Enabled = False
txtpos3.Enabled = False
txtreason1.Enabled = False
txtreason2.Enabled = False
txtreason3.Enabled = False
txtsuperior1.Enabled = False
txtsuperior2.Enabled = False
txtsuperior3.Enabled = False
txtrefname1.Enabled = False
txtrefname2.Enabled = False
txtrefname3.Enabled = False
txtrefjob1.Enabled = False
txtrefjob2.Enabled = False
txtrefjob3.Enabled = False
txtrefaddress1.Enabled = False
txtrefaddress2.Enabled = False
txtrefaddress3.Enabled = False
txtrefphone1.Enabled = False
txtrefphone2.Enabled = False
txtrefphone3.Enabled = False
txtrefrel1.Enabled = False
txtrefrel2.Enabled = False
txtrefrel3.Enabled = False
' 09/29/04 ********************
'dtdateprob.Enabled = False
mskdateprob.Enabled = False
'txtdatehired.Enabled = False
mskdatehired.Enabled = False
'dtdatereg.Enabled = False
mskdatereg.Enabled = False
mskdateres.Enabled = False
'dtdateresign.Enabled = False

cmbempstatus.Enabled = False
txtsupervisor.Enabled = False
txtreasonresign.Enabled = False
'dtendcontract.Enabled = False
mskendof.Enabled = False

txtdtrexempt.Enabled = False
txthrswork.Enabled = False
txtnolate.Enabled = False
txtnoundertime.Enabled = False
txtActype.Enabled = False
txtnoot.Enabled = False
cmbdepartment.Enabled = False
mskdep1.Value = 0


Set rscivil = cnDCPI.Execute("Select * from HRcivilstatus")
Do While Not rscivil.EOF
    cmbcivilstatus.AddItem rscivil!civilstatus & "-" & rscivil!civilstatusdesc
    rscivil.MoveNext
    Loop
    rscivil.Close
    
Set rsempstatus = cnDCPI.Execute("Select * from HRemploymentstatus")
Do While Not rsempstatus.EOF
    cmbempstatus.AddItem rsempstatus!employmentstatus & "-" & rsempstatus!employmentstatusdesc
    rsempstatus.MoveNext
    Loop
    rsempstatus.Close

Set rsdept = cnDCPI.Execute("select * from DCDepartment")
 Do While Not rsdept.EOF
        cmbdepartment.AddItem rsdept!Department
        cmbdepartment.ItemData(cmbdepartment.NewIndex) = rsdept!Deptcode
        rsdept.MoveNext
    Loop
    rsdept.Close

'Set rsdept = cnDCPI.Execute("select * from DCDepartment")
' Do While Not rsdept.EOF
'        cmbdepartment.AddItem rsdept!DeptCode & "-" & rsdept!Department
'        cmbdepartment.ItemData(cmbdepartment.NewIndex) = rsdept!DeptCode
'        rsdept.MoveNext
'    Loop
'    rsdept.Close


rsout.Open "select * from HREmployees order by employeenum", cnDCPI, adOpenKeyset, adLockOptimistic
rsdetails.Open "Select * from HRDetails", cnDCPI, adOpenKeyset, adLockOptimistic
rsref.Open "Select * from HRCharReference", cnDCPI, adOpenKeyset, adLockOptimistic
rsdialect.Open "Select * from HRDialects", cnDCPI, adOpenKeyset, adLockOptimistic
rsemprec.Open "Select * from HREmploymentRecord", cnDCPI, adOpenKeyset, adLockOptimistic
rssem.Open "select * from HRSeminar", cnDCPI, adOpenKeyset, adLockOptimistic

If rsout.RecordCount > 0 Then
'    Showdata rsout
    cmdundo.Enabled = False
    cmdupdate.Enabled = False
Else
   Call clearscreen
   cmdfirst.Enabled = False
   cmdprev.Enabled = False
   cmdnext.Enabled = False
   cmdlast.Enabled = False
   cmdedit.Enabled = False
   cmddelete.Enabled = False
   'cmdsearch.Enabled = False
   cmdhired.Enabled = False
   cmdundo.Enabled = False
   cmdupdate.Enabled = False
End If
empno = txtEmployeeNo.Text

Dim strSQL As String
    '************** Add Click in TrackShipment ***************
    'Opens table
    If addflag = True Then
        Call cmdnew_Click
    End If
    If editflag = True Then
         rsout.Find "employeenum =" & frmHRMainEmployee.txtEmpNo.Text
         showdata rsout
         Call cmdEdit_Click
        End If


    If editflag = False And addflag = False Then
        
        rsout.Find "employeenum =" & frmHRMainEmployee.txtEmpNo.Text
        cmdnew.Enabled = False
        cmdedit.Enabled = False
        cmdupdate.Enabled = False
        cmddelete.Enabled = False
        cmdundo.Enabled = False
        cmdfirst.Enabled = False
        cmdprev.Enabled = False
        cmdnext.Enabled = False
        cmdlast.Enabled = False
        showdata rsout
    End If
        
End Sub

Public Sub showdata(rec As ADODB.Recordset)
Dim rs As ADODB.Recordset

    txtEmployeeNo.Text = IIf(IsNull(rec!employeenum), "", rec!employeenum)
    txtfirstchoice.Text = IIf(IsNull(rec!Position), "", rec!Position)
    'txtdateapplied.Value = IIf(IsNull(Rec!applicationdate), "01/01/1990", Rec!applicationdate)
    txtnickname.Text = IIf(IsNull(rec!nickname), "", rec!nickname)
    txtsurname.Text = IIf(IsNull(rec!surname), "", rec!surname)
    txtfirstname.Text = IIf(IsNull(rec!FirstName), "", rec!FirstName)
    txtmiddlename.Text = IIf(IsNull(rec!middlename), "", rec!middlename)
    txtaddress1.Text = IIf(IsNull(rec!address1), "", rec!address1)
    txthouseno.Text = IIf(IsNull(rec!houseno), "", rec!houseno)
    txtstreet.Text = IIf(IsNull(rec!Street), "", rec!Street)
    txtbarangay.Text = IIf(IsNull(rec!barangay), "", rec!barangay)
    txtdistrict.Text = IIf(IsNull(rec!district), "", rec!district)
    txtprovince.Text = IIf(IsNull(rec!province1), "", rec!province1)
    txtcity.Text = IIf(IsNull(rec!city), "", rec!city)
    txtprovinceaddress.Text = IIf(IsNull(rec!province), "", rec!province)
    txtage.Text = IIf(IsNull(rec!age), 0, rec!age)
    txtphilhealth.Text = IIf(IsNull(rec!philhealth), "", rec!philhealth)
    'mskdateapplied.Text = IIf((Format(Rec!ApplicationDate, "m/d/yyyy") = "1/1/1900") Or (IsNull(Rec!ApplicationDate)) Or (Rec!ApplicationDate = ""), "__/__/____", Format(Rec!ApplicationDate, "mm/dd/yyyy"))
    mskdateapplied.Text = IIf(IsDate(rec!applicationdate), Format(rec!applicationdate, "mm/dd/yyyy"), "__/__/____")
    'mskbirthdate.Text = IIf(IsDate(rec!birthdate), Format(rec!birthdate, "mm/dd/yyyy"), "__/__/____")
    'mskbirthdate.Text = IIf(IsDate(rec!birthdate), Format(rec!birthdate, "mm/dd/yyyy"), "__/__/____")
    mskbirthdate.Text = IIf((Format(rec!birthdate, "m/d/yyyy") = "1/1/1900") Or (Format(rec!birthdate, "mm/dd/yyyy") = 1 / 1 / 1900) Or (IsNull(rec!birthdate)) Or (rec!birthdate = ""), "__/__/____", Format(rec!birthdate, "mm/dd/yyyy"))
    mskdatehired.Text = IIf((Format(rec!datehired, "mm/dd/yyyy") = "01/01/1900") Or (IsNull(rec!datehired)) Or (rec!datehired = ""), "__/__/____", Format(rec!datehired, "mm/dd/yyyy"))
    'mskdatehired.Text = IIf(IsDate(Rec!datehired), Format(Rec!datehired, "mm/dd/yyyy"), "__/__/____")
    mskdateprob.Text = IIf((Format(rec!dateprob, "m/d/yyyy") = "1/1/1900") Or (IsNull(rec!dateprob)) Or (rec!dateprob = ""), "__/__/____", Format(rec!dateprob, "mm/dd/yyyy"))
    'mskdateprob.Text = IIf(IsDate(Rec!dateprob), Format(Rec!dateprob, "mm/dd/yyyy"), "__/__/____")
    mskdatereg.Text = IIf((Format(rec!dateregular, "m/d/yyyy") = "1/1/1900") Or (Format(rec!dateregular, "mm/dd/yyy") = "01/01/1900") Or (IsNull(rec!dateregular)) Or (rec!dateregular = ""), "__/__/____", Format(rec!dateregular, "mm/dd/yyyy"))
    'mskdatereg.Text = IIf(IsDate(Rec!dateregular), Format(Rec!dateregular, "mm/dd/yyyy"), "__/__/____")
    mskissue.Text = IIf((Format(rec!issueon, "m/d/yyyy") = "1/1/1900") Or (Format(rec!issueon, "mm/dd/yyyy") = "01/01/1900") Or (IsNull(rec!issueon)) Or (rec!issueon = ""), "__/__/____", Format(rec!issueon, "mm/dd/yyyy"))
    'mskissue.Text = IIf(IsDate(Rec!issueon), Format(Rec!issueon, "mm/dd/yyyy"), "__/__/____")
    mskdateres.Text = IIf((Format(rec!dateresign, "m/d/yyyy") = "1/1/1900") Or (IsNull(rec!dateresign)) Or (rec!dateresign = ""), "__/__/____", Format(rec!dateresign, "mm/dd/yyyy"))
    'mskdateres.Text = IIf(IsDate(Rec!dateresign), Format(Rec!dateresign, "mm/dd/yyyy"), "__/__/____")
    mskendof.Text = IIf((Format(rec!endofcontract, "m/d/yyyy") = "1/1/1900") Or (Format(rec!endofcontract, "mm/dd/yyyy") = "01/01/1900") Or (IsNull(rec!endofcontract)) Or (rec!endofcontract = ""), "__/__/____", Format(rec!endofcontract, "mm/dd/yyyy"))
    'mskendof.Text = IIf(IsDate(Rec!endofcontract), Format(Rec!endofcontract, "mm/dd/yyyy"), "__/__/____")

' 09/29/04 *****************
    'mskbday.Text = Format(IIf(IsNull(Rec!birthdate), "01/01/1900", Rec!birthdate), "mm/dd/yyyy")
    'mskbday.Value = IIf(IsNull(Rec!birthdate), "01/01/1900", Rec!birthdate)
   
'    If IsNull(Rec!ApplicationDate) Or Rec!ApplicationDate = "1/1/1900" Then
'        txtdateapplied.Visible = False
'       txtDtApp1.Visible = True
'    Else
'        txtdateapplied.Visible = True
'        txtDtApp1.Visible = False
'    End If
    'mskdateapplied.Text = IIf(IsNull(Rec!applicationdate) Or (Rec!applicationdate) = "", "__/__/____", Format(Rec!applicationdate, "mm/dd/yyyy"))
    
    'mskissue.Text = IIf((Rec!issueon = "1/1/1900") Or (IsNull(Rec!issueon)), "__/__/____", Format(Rec!issueon, "mm/dd/yyyy"))
    'msdep1.Text = IIf(IsNull(Rec2!dep1bdate) Or (Rec2!dep1bdate) = "", "__-__-____", Format(Rec2!dep1bdate, "mm-dd-yyyy"))
    
    
'    If IsNull(Rec!BirthDate) Or Rec!BirthDate = "1/1/1900" Or Rec!BirthDate = "Jan  1 1900 12:00AM" Then
'       mskbday.Visible = False
'       txtDtBrth.Visible = True
'    Else
'        mskbday.Visible = True
'        txtDtBrth.Visible = False
'        mskbday.Value = Rec!BirthDate
'    End If
    'mskbirthdate.Text = IIf(IsNull(Rec!birthdate) Or (Rec!birthdate) = "", "__/__/____", Format(Rec!birthdate, "mm/dd/yyyy"))

'    If IsNull(Rec!IssueOn) Or Rec!IssueOn = "1/1/1900" Then
'        txtIssOn.Visible = True
'       dtissueon.Visible = False
'    Else
'        dtissueon.Visible = True
'        txtIssOn.Visible = False
'        dtissueon.Value = Rec!IssueOn
'    End If

'    If IsNull(Rec!datehired) Or Rec!datehired = "1/1/1900" Or Rec!datehired = "Jan  1 1900 12:00AM" Then
'       txtDtHr.Visible = True
'       txtdatehired.Visible = False
'    Else
'        txtdatehired.Visible = True
'        txtDtHr.Visible = False
'        txtdatehired.Value = Rec!datehired
'    End If
    'mskdatehired.Text = IIf(IsNull(Rec!datehired) Or (Rec!datehired) = "", "__/__/____", Format(Rec!datehired, "mm/dd/yyyy"))
    
'    If IsNull(Rec!DateProb) Or Rec!DateProb = "1/1/1900" Then
'       txtDtPr.Visible = True
'       dtdateprob.Visible = False
'    Else
'        dtdateprob.Visible = True
'        txtDtPr.Visible = False
'        dtdateprob.Value = Rec!datehired
'    End If
            
'    If IsNull(Rec!DateRegular) Or Rec!DateRegular = "1/1/1900" Then
'       txtDtReg.Visible = True
'       dtdatereg.Visible = False
'    Else
'        dtdatereg.Visible = True
'        txtDtReg.Visible = False
'        dtdatereg = Rec!DateRegular
'    End If

    
'    If IsNull(Rec!DateResign) Or Rec!DateResign = "1/1/1900" Then
'       txtDtRes.Visible = True
'       dtdateresign.Visible = False
'    Else
'        dtdateresign.Visible = True
'        txtDtRes.Visible = False
'        dtdateresign = Rec!DateResign
'    End If

    
'    If IsNull(Rec!EndOfContract) Or Rec!EndOfContract = "1/1/1900" Then
'       txtECt.Visible = True
'       dtendcontract.Visible = False
'    Else
'        dtendcontract.Visible = True
'        txtECt.Visible = False
'        dtdateresign = Rec!EndOfContract
'    End If
        
    txtemergency.Text = IIf(IsNull(rec!contact), "", rec!contact)
    txteraddress.Text = IIf(IsNull(rec!CAddress), "", rec!CAddress)
    txtertel.Text = IIf(IsNull(rec!CPhone), "", rec!CPhone)
    If (rec!gender) = "M" Then
        optMale.Value = True
        optFemale.Value = False
    ElseIf (rec!gender) = "F" Then
        optMale.Value = False
        optFemale.Value = True
    Else
        optMale.Value = False
        optFemale.Value = False
    End If
    txtheight.Text = IIf(IsNull(rec!Height), "", rec!Height)
    txtweight.Text = IIf(IsNull(rec!Weight), "", rec!Weight)
    txttel1.Text = IIf(IsNull(rec!telnum1), "", rec!telnum1)
    txttel2.Text = IIf(IsNull(rec!telnum2), "", rec!telnum2)
    txttel3.Text = IIf(IsNull(rec!telnum3), "", rec!telnum3)
    txttel4.Text = IIf(IsNull(rec!telnum4), "", rec!telnum4)
   ' cmbcivilstatus.Text = IIf(IsNull(Rec!civilstatus), "", Rec!civilstatus)
    txtappstatus.Text = IIf(IsNull(rec!applicantstatus), "", rec!applicantstatus)
    txtexam.Text = IIf(IsNull(rec!examgrade), "", rec!examgrade)
    txtinterview.Text = IIf(IsNull(rec!interviewgrade), "", rec!interviewgrade)
    txtmedical.Text = IIf(IsNull(rec!medicalgrade), "", rec!medicalgrade)
    msksss.Text = IIf(IsNull(rec!sss), "__-_______-_", rec!sss)
    'txtphcode.Text = IIf(IsNull(Rec!phcode), "", Rec!phcode)
    txtpagibig.Text = IIf(IsNull(rec!pagibig) Or (rec!pagibig = "0"), "", rec!pagibig)
    txtacctno.Text = IIf(IsNull(rec!acctno), "", rec!acctno)
    txtbirthplace.Text = IIf(IsNull(rec!birthplace), "", rec!birthplace)
    'txthaircolor.Text = IIf(IsNull(Rec!haircolor), "", Rec!haircolor)
    'txteyecolor.Text = IIf(IsNull(Rec!eyecolor), "", Rec!eyecolor)
    txtreligion.Text = IIf(IsNull(rec!religion), "", rec!religion)
    msktax.Text = IIf(IsNull(rec!taxid), "___-___-___", rec!taxid)
    txtprimarysch.Text = IIf(IsNull(rec!Primary), "", rec!Primary)
    txtdateprimary.Text = IIf(IsNull(rec!DatePrimary), "", rec!DatePrimary)
    txtsecsch.Text = IIf(IsNull(rec!Secondary), "", rec!Secondary)
    txtdatesec.Text = IIf(IsNull(rec!DateSecondary), "", rec!DateSecondary)
    txtcollegesch.Text = IIf(IsNull(rec!College), "", rec!College)
    txtdatecollege.Text = IIf(IsNull(rec!DateCollege), "", rec!DateCollege)
    txtcitizen.Text = IIf(IsNull(rec!citizen), "", rec!citizen)
    txtcollegecourse.Text = IIf(IsNull(rec!degree), "", rec!degree)
    txtcollegemajor.Text = IIf(IsNull(rec!Major), "", rec!Major)
    txtpostsch.Text = IIf(IsNull(rec!PostGraduate), "", rec!PostGraduate)
    txtdatepost.Text = IIf(IsNull(rec!DatePostGraduate), "", rec!DatePostGraduate)
    txtpostcourse.Text = IIf(IsNull(rec!PostGraduateCourse), "", rec!PostGraduateCourse)
    txtfathername.Text = IIf(IsNull(rec!fathername), "", rec!fathername)
    txtfaddress.Text = IIf(IsNull(rec!faddress), "", rec!faddress)
    txtfjob.Text = IIf(IsNull(rec!fwork), "", rec!fwork)
    txtmothername.Text = IIf(IsNull(rec!mothername), "", rec!mothername)
    txtmaddress.Text = IIf(IsNull(rec!maddress), "", rec!maddress)
    txtmjob.Text = IIf(IsNull(rec!mwork), "", rec!mwork)
    txtnbi.Text = IIf(IsNull(rec!NBI), "", rec!NBI)
    'txtrescer.Text = IIf(IsNull(Rec!rescer), "", Rec!rescer)
    'txtissueat.Text = IIf(IsNull(Rec!issueat), "", Rec!issueat)
    'dtissueon.Value = IIf(IsNull(Rec!issueon), "", Rec!issueon)
    'txtdatehired.Value = IIf(IsNull(Rec!datehired), "", Rec!datehired)
    'dtdateprob.Value = IIf(IsNull(Rec!dateprob), "", Rec!dateprob)
    'dtdatereg.Value = IIf(IsNull(Rec!dateregular), "", Rec!dateregular)
    'dtdateresign.Value = IIf(IsNull(Rec!dateresign), "", Rec!dateresign)
'    cmbempstatus.Text = IIf(IsNull(Rec!employmentstatus), "", Rec!employmentstatus)
    txtsupervisor.Text = IIf(IsNull(rec!supervisor), "", rec!supervisor)
    txtreasonresign.Text = IIf(IsNull(rec!ReasonRes), "", rec!ReasonRes)
    'dtendcontract.Value = IIf(IsNull(Rec!endofcontract), "", Rec!endofcontract)
    txtdtrexempt.Text = IIf(IsNull(rec!DTRExempt), "", rec!DTRExempt)
    txthrswork.Text = IIf(IsNull(rec!HrsWork), "", rec!HrsWork)
    txtnolate.Text = IIf(IsNull(rec!NoLate), "", rec!NoLate)
    txtnoundertime.Text = IIf(IsNull(rec!NoUnder), "", rec!NoUnder)
    txtActype.Text = IIf(IsNull(rec!ActType), "", rec!ActType)
    txtnoot.Text = IIf(IsNull(rec!NoOT), "", rec!NoOT)
       
    Set rs = cnDCPI.Execute("Select * From HRCivilStatus Where civilstatus ='" & IIf(IsNull(rec!civilstatus), "", rec!civilstatus) & "'")
        If Not rs.EOF Then
            cmbcivilstatus.Text = rs.Fields(0) & "-" & rs.Fields(1)
        Else
            cmbcivilstatus.ListIndex = -1
        End If
        rs.Close

    Set rs = cnDCPI.Execute("Select * From DCDepartment Where deptCode =" & IIf(IsNull(rec!Deptcode), 0, rec!Deptcode))
        If Not rs.EOF Then
            cmbdepartment.Text = rs.Fields(1)
        Else
            cmbdepartment.ListIndex = -1
        End If
        rs.Close

'    Set rs = cnDCPI.Execute("Select * From DCDepartment Where deptcode =" & IIf(IsNull(Rec!DeptCode), "", Rec!DeptCode))
'        If Not rs.EOF Then
'            cmbdepartment.Text = rs.Fields(0) & "-" & rs.Fields(1)
'        Else
'            cmbdepartment.ListIndex = -1
'        End If
'        rs.Close

    Set rs = cnDCPI.Execute("Select * From HREmploymentstatus Where employmentstatus ='" & IIf(IsNull(rec!employmentstatus), "", rec!employmentstatus) & "'")
        If Not rs.EOF Then
            cmbempstatus.Text = rs.Fields(0) & "-" & rs.Fields(1)
        Else
            cmbempstatus.ListIndex = -1
        End If
        rs.Close
    
'showdata dialect
'    Set rs = cnDCPI.Execute("Select * from HRDialects where employeenum='" & txtEmployeeNo.Text & "'")
'    txtdialect1 = IIf(IsNull(rec1!dialect1), "", rec1!dialect1)
'    txtdialect2 = IIf(IsNull(rec1!dialect2), "", rec1!dialect2)
'    txtdialect3 = IIf(IsNull(rec1!dialect3), "", rec1!dialect3)


    Set rs = cnDCPI.Execute("Select * from HRDialects where convert(int, employeenum) ='" & Val(txtEmployeeNo.Text) & "'")
    Do While Not rs.EOF
        Showdialect rs
       rs.MoveNext
    Loop
    rs.Close
    
    Set rs = cnDCPI.Execute("Select * from HRDetails where convert(int, employeenum)='" & Val(txtEmployeeNo.Text) & "'")
    Do While Not rs.EOF
        Showdetails rs
        rs.MoveNext
    Loop
    rs.Close

    Set rs = cnDCPI.Execute("Select * from HRCharReference where convert(int, employeenum) ='" & Val(txtEmployeeNo.Text) & "'")
    Do While Not rs.EOF
        ShowRef rs
        rs.MoveNext
    Loop
    rs.Close

    Set rs = cnDCPI.Execute("Select * from HREmploymentRecord where convert(int, employeenum) ='" & Val(txtEmployeeNo.Text) & "'")
    Do While Not rs.EOF
        ShowEmpRec rs
        rs.MoveNext
    Loop
    rs.Close

    Set rs = cnDCPI.Execute("Select * from HRSeminar where convert(int, employeenum) ='" & Val(txtEmployeeNo.Text) & "'")
    Do While Not rs.EOF
        ShowSem rs
        rs.MoveNext
    Loop
    rs.Close
End Sub
Public Sub Showdialect(rec1 As ADODB.Recordset)
    txtdialect1 = IIf(IsNull(rec1!dialect1), "", rec1!dialect1)
    txtdialect2 = IIf(IsNull(rec1!dialect2), "", rec1!dialect2)
    txtdialect3 = IIf(IsNull(rec1!dialect3), "", rec1!dialect3)
End Sub
Public Sub Showdetails(Rec2 As ADODB.Recordset)

   txtspousename.Text = IIf(IsNull(Rec2!spousename), "", Rec2!spousename)
   txtspouseaddress.Text = IIf(IsNull(Rec2!spouseaddress), "", Rec2!spouseaddress)
   txtsjob.Text = IIf(IsNull(Rec2!spousework), "", Rec2!spousework)
   txtnumdep.Text = IIf(IsNull(Rec2!Dependents), 0, Rec2!Dependents)
   txtdep1.Text = IIf(IsNull(Rec2!dep1name), "", Rec2!dep1name)
   txtdep2.Text = IIf(IsNull(Rec2!dep2name), "", Rec2!dep2name)
   txtdep3.Text = IIf(IsNull(Rec2!dep3name), "", Rec2!dep3name)
   txtdep4.Text = IIf(IsNull(Rec2!dep4name), "", Rec2!dep4name)
   txtdep5.Text = IIf(IsNull(Rec2!dep5name), "", Rec2!dep5name)

'*****makiss***** 09/28/04
   'mskdep1.Value = IIf(IsNull(Rec2!dep1bdate), Null, Format(Rec2!dep1bdate, "mm/dd/yy"))
   
   'msdep1.Text = IIf(IsNull(Rec2!dep1bdate) Or (Rec2!dep1bdate) = "", "__/__/____", Format(Rec2!dep1bdate, "mm/dd/yyyy"))
   'msdep1.Text = IIf((Format(Rec2!dep1bdate, "m/d/yyyy") = "1/1/1900") Or (IsNull(Rec2!dep1bdate)) Or (Rec2!dep1bdate = "") Or (Rec2!dep1bdate = "__/__/__") Or (Rec2!dep1bdate = "__/__/____"), "__/__/____", Format(Rec2!dep1bdate, "mm/dd/yyyy"))
   msdep1.Text = IIf(IsDate(Rec2!dep1bdate), Format(Rec2!dep1bdate, "mm/dd/yyyy"), "__/__/____")
   'mskdep2.Value = IIf(IsNull(Rec2!dep2bdate), Null, Format(Rec2!dep2bdate, "mm/dd/yy"))
   'msdep2.Text = IIf(IsNull(Rec2!dep2bdate) Or (Rec2!dep2bdate) = "", "__/__/____", Format(Rec2!dep2bdate, "mm/dd/yyyy"))
   'msdep2.Text = IIf((Format(Rec2!dep2bdate, "m/d/yyyy") = "1/1/1900") Or (IsNull(Rec2!dep2bdate)) Or (Rec2!dep2bdate = "") Or (Rec2!dep2bdate = "__/__/__") Or (Rec2!dep2bdate = "__/__/____"), "__/__/____", Format(Rec2!dep2bdate, "mm/dd/yyyy"))
   msdep2.Text = IIf(IsDate(Rec2!dep2bdate), Format(Rec2!dep2bdate, "mm/dd/yyyy"), "__/__/____")
   
   'mskdep3.Value = IIf(IsNull(Rec2!dep3bdate), "01/01/1900", Format(Rec2!dep3bdate, "mm/dd/yy"))
   'msdep3.Text = IIf(IsNull(Rec2!dep3bdate), "__-__-____", Rec2!dep3bdate)
   'msdep3.Text = IIf(IsNull(Rec2!dep3bdate) Or (Rec2!dep3bdate) = "", "__/__/____", Format(Rec2!dep3bdate, "mm/dd/yyyy"))
   'msdep3.Text = IIf((Format(Rec2!dep3bdate, "m/d/yyyy") = "1/1/1900") Or (IsNull(Rec2!dep3bdate)) Or (Rec2!dep3bdate = "") Or (Rec2!dep3bdate = "__/__/__") Or (Rec2!dep3bdate = "__/__/____"), "__/__/____", Format(Rec2!dep3bdate, "mm/dd/yyyy"))
   msdep3.Text = IIf(IsDate(Rec2!dep3bdate), Format(Rec2!dep3bdate, "mm/dd/yyyy"), "__/__/____")
   'mskdep4.Value = IIf(IsNull(Rec2!dep4bdate), "01/01/1900", Format(Rec2!dep4bdate, "mm/dd/yy"))
   'msdep4.Text = IIf(IsNull(Rec2!dep4bdate) Or (Rec2!dep4bdate) = "", "__/__/____", Format(Rec2!dep4bdate, "mm/dd/yyyy"))
   'msdep4.Text = IIf((Format(Rec2!dep4bdate, "m/d/yyyy") = "1/1/1900") Or (IsNull(Rec2!dep4bdate)) Or (Rec2!dep4bdate = "") Or (Rec2!dep4bdate = "__/__/__") Or (Rec2!dep4bdate = "__/__/____"), "__/__/____", Format(Rec2!dep4bdate, "mm/dd/yyyy"))
   msdep4.Text = IIf(IsDate(Rec2!dep4bdate), Format(Rec2!dep4bdate, "mm/dd/yyyy"), "__/__/____")
   'mskdep5.Value = IIf(IsNull(Rec2!dep5bdate), "01/01/1900", Format(Rec2!dep5bdate, "mm/dd/yy"))
   'msdep5.Text = IIf(IsNull(Rec2!dep5bdate) Or (Rec2!dep5bdate) = "", "__/__/____", Format(Rec2!dep5bdate, "mm/dd/yyyy"))
   'msdep5.Text = IIf((Format(Rec2!dep5bdate, "m/d/yyyy") = "1/1/1900") Or (IsNull(Rec2!dep5bdate)) Or (Rec2!dep5bdate = "") Or (Rec2!dep5bdate = "__/__/__") Or (Rec2!dep5bdate = "__/__/____"), "__/__/____", Format(Rec2!dep5bdate, "mm/dd/yyyy"))
   msdep5.Text = IIf(IsDate(Rec2!dep5bdate), Format(Rec2!dep5bdate, "mm/dd/yyyy"), "__/__/____")
   
   txthfdep.Text = IIf(IsNull(Rec2!hfdep), "", Rec2!hfdep)
   'mskhpbdate.Value = IIf(IsNull(Rec2!hpbdate), "01/01/1900", Rec2!hpbdate)
   'mshpbdate.Text = IIf(IsNull(Rec2!hpbdate) Or (Rec2!hpbdate) = "", "__/__/____", Format(Rec2!hpbdate, "mm/dd/yyyy"))
   'mshpbdate.Text = IIf((Format(Rec2!hpbdate, "m/d/yyyy") = "1/1/1900") Or (IsNull(Rec2!hpbdate)) Or (Rec2!hpbdate = "") Or (Rec2!hpbdate = "__/__/__") Or (Rec2!hpbdate = "__/__/____"), "__/__/____", Format(Rec2!hpbdate, "mm/dd/yyyy"))
   mshpbdate.Text = IIf(IsDate(Rec2!hpbdate), Format(Rec2!hpbdate, "mm/dd/yyyy"), "__/__/____")
   txthprelate.Text = IIf(IsNull(Rec2!hprelate), "", Rec2!hprelate)
   
End Sub
Public Sub ShowRef(rec3 As ADODB.Recordset)

    txtrefname1.Text = IIf(IsNull(rec3!Name), "", rec3!Name)
    txtrefjob1.Text = IIf(IsNull(rec3!occupation), "", rec3!occupation)
    txtrefaddress1.Text = IIf(IsNull(rec3!Address), "", rec3!Address)
    txtrefphone1.Text = IIf(IsNull(rec3!telnum), "", rec3!telnum)
    txtrefrel1.Text = IIf(IsNull(rec3!Relation), "", rec3!Relation)

    txtrefname2.Text = IIf(IsNull(rec3!Name1), "", rec3!Name1)
    txtrefjob2.Text = IIf(IsNull(rec3!occupation1), "", rec3!occupation1)
    txtrefaddress2.Text = IIf(IsNull(rec3!address1), "", rec3!address1)
    txtrefphone2.Text = IIf(IsNull(rec3!telnum1), "", rec3!telnum1)
    txtrefrel2.Text = IIf(IsNull(rec3!Relation1), "", rec3!Relation1)

    txtrefname3.Text = IIf(IsNull(rec3!Name2), "", rec3!Name2)
    txtrefjob3.Text = IIf(IsNull(rec3!occupation2), "", rec3!occupation2)
    txtrefaddress3.Text = IIf(IsNull(rec3!Address2), "", rec3!Address2)
    txtrefphone3.Text = IIf(IsNull(rec3!telnum2), "", rec3!telnum2)
    txtrefrel3.Text = IIf(IsNull(rec3!Relation2), "", rec3!Relation2)
    
End Sub
Public Sub ShowEmpRec(rec4 As ADODB.Recordset)

    txtnameofcompany1.Text = IIf(IsNull(rec4!company), "", rec4!company)
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
Private Sub CheckUser()
    
    If UL = "AD" Then
        cmdnew.Enabled = True
        cmdedit.Enabled = True
        cmddelete.Enabled = True
        cmdupdate.Enabled = True
'        cmdhired.Enabled = True
        'cmdsearch.Enabled = True
        cmdundo.Enabled = True
        'DCListView.Enabled = True
        'DCFBDet.Enabled = True
        
        'menFileAdd.Enabled = True
        'menFileEdit.Enabled = True
        'menFileExport.Enabled = True
        'menRecordSearch.Enabled = True
        'menRecordFilter.Enabled = True
        
        
    Else
        Select Case HR
            Case "AD"
                cmdnew.Enabled = True
                cmdedit.Enabled = True
                cmddelete.Enabled = True
                cmdupdate.Enabled = True
                'cmdhired.Enabled = True
                'cmdsearch.Enabled = True
                cmdundo.Enabled = True
                'DCListView.Enabled = True
                'DCFBDet.Enabled = True
                
                'menFileAdd.Enabled = True
                'menFileEdit.Enabled = True
                'menFileExport.Enabled = True
                'menRecordSearch.Enabled = True
                'menRecordFilter.Enabled = True
        
            Case "EN"
                cmdnew.Enabled = True
                cmdedit.Enabled = True
                cmddelete.Enabled = False
                cmdupdate.Enabled = True
                'cmdhired.Enabled = False
                'cmdsearch.Enabled = True
                cmdundo.Enabled = True
                'DCListView.Enabled = True
                'DCFBDet.Enabled = True
                
                'menFileAdd.Enabled = True
                'menFileEdit.Enabled = True
               ' menFileDelete.Enabled = False
                'menFileExport.Enabled = False
                'menRecordSearch.Enabled = True
                'menRecordFilter.Enabled = True
                
            Case "MG"
                cmdnew.Enabled = False
                cmdedit.Enabled = False
                cmddelete.Enabled = False
                cmdupdate.Enabled = False
               ' cmdhired.Enabled = False
               ' cmdsearch.Enabled = True
                cmdundo.Enabled = True
                'DCListView.Enabled = False
                'DCFBDet.Enabled = False
                
                'menFileAdd.Enabled = False
                'menFileEdit.Enabled = False
                'menFileExport.Enabled = False
                'menRecordSearch.Enabled = True
                'menRecordFilter.Enabled = True
            Case "SU"
                cmdnew.Enabled = True
                cmdedit.Enabled = True
                cmddelete.Enabled = True
                cmdupdate.Enabled = True
               ' cmdhired.Enabled = True
              '  cmdsearch.Enabled = True
                cmdundo.Enabled = True
                'DCListView.Enabled = True
                'DCFBDet.Enabled = True
                
                'menFileAdd.Enabled = True
                'menFileEdit.Enabled = True
                'menFileExport.Enabled = True
                'menRecordSearch.Enabled = True
                'menRecordFilter.Enabled = True
            
        End Select
             
    End If
    
End Sub
Private Sub clearscreen()
txtEmployeeNo.Text = ""
txthfdep.Text = ""
'09/29/04 *************************
'mskhpbdate.Value = ""
mshpbdate.Text = "__/__/____"

txthprelate.Text = ""
txtprimarysch.Text = ""
txtdateprimary.Text = ""
txtnickname.Text = ""
txtsurname.Text = ""
txtfirstname.Text = ""
txtmiddlename.Text = ""
txtfirstchoice.Text = ""
'09/29/04
'mskbday.Value = ""
mskbirthdate.Text = "__/__/____"

txtage.Text = 0
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
optMale.Value = False
optFemale.Value = False
txtheight.Text = ""
txtweight.Text = ""
txttel1.Text = ""
txttel2.Text = ""
txttel3.Text = ""
txttel4.Text = ""
cmbcivilstatus.ListIndex = -1
cmbempstatus.ListIndex = -1
'09/29/04 *************************
'txtdateapplied.Value = ""
mskdateapplied.Text = "__/__/____"

txtappstatus.Text = ""
txtexam.Text = ""
txtinterview.Text = ""
txtmedical.Text = ""
msksss.Text = "__-_______-_"
'txtphcode.Text = ""
txtpagibig.Text = ""
txtacctno.Text = ""
txtcitizen.Text = ""
txtbirthplace.Text = ""
'txthaircolor.Text = ""
'txteyecolor.Text = ""
txtreligion.Text = ""
msktax.Text = "___-___-___"
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
'txtrescer.Text = ""
'txtissueat.Text = ""
'dtissueon.Value = ""
mskissue.Text = "__/__/____"
txtdep1.Text = ""
txtdep2.Text = ""
txtdep3.Text = ""
txtdep4.Text = ""
txtdep5.Text = ""

'makiss 09/28/04 *****************
'mskdep1.Value = ""
msdep1.Text = "__/__/____"
'mskdep2.Value = ""
msdep2.Text = "__/__/____"
'mskdep3.Value = ""
msdep3.Text = "__/__/____"
'mskdep4.Value = ""
msdep4.Text = "__/__/____"
'mskdep5.Value = ""
msdep5.Text = "__/__/____"
'makiss 09/28/04 ***************

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
' 09/29/04*********************
'dtdateprob.Value = ""
mskdateprob.Text = "__/__/____"
'txtdatehired.Value = ""
mskdatehired.Text = "__/__/____"
mskdatereg.Text = "__/__/____"
'dtdatereg.Value = ""
'dtdateresign.Value = ""
mskdateres.Text = "__/__/____"
cmbempstatus.ListIndex = -1
txtsupervisor.Text = ""
txtreasonresign.Text = ""
'dtendcontract.Value = ""
mskendof.Text = "__/__/____"

txtdtrexempt.Text = ""
txthrswork.Text = ""
txtnolate.Text = ""
txtnoundertime.Text = ""
txtActype.Text = ""
txtnoot.Text = ""
If Not Trim(projectlong) = "" Then
        cmbdepartment.Text = projectlong
    Else
        cmbdepartment.ListIndex = -1
    End If
End Sub

Private Sub mskb1birth_GotFocus()
mskb1birth.BackColor = &H80000016
End Sub

Private Sub mskb1birth_LostFocus()
mskb1birth.BackColor = vbWhite
End Sub

Private Sub mskb2birth_GotFocus()
mskb2birth.BackColor = &H80000016
End Sub

Private Sub mskb2birth_LostFocus()
mskb2birth.BackColor = vbWhite
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim res As String
If cmdupdate.Enabled = True Then
    If editflag = True Or addflag = True Then
        res = MsgBox("Save changes?", vbYesNo + vbQuestion, "DCPI Information System")
        If res = 6 Then
            cmdupdate_Click
        End If
    End If
End If
End Sub

Private Sub msdep1_GotFocus()
msdep1.BackColor = &H80000016
End Sub

Private Sub msdep1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub msdep1_LostFocus()
msdep1.BackColor = vbWhite
End Sub

Private Sub msdep2_GotFocus()
msdep2.BackColor = &H80000016
End Sub

Private Sub msdep2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub msdep2_LostFocus()
msdep2.BackColor = vbWhite
End Sub

Private Sub msdep3_GotFocus()
msdep3.BackColor = &H80000016
End Sub

Private Sub msdep3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub msdep3_LostFocus()
msdep3.BackColor = vbWhite
End Sub

Private Sub msdep4_GotFocus()
msdep4.BackColor = &H80000016
End Sub

Private Sub msdep4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub msdep4_LostFocus()
msdep4.BackColor = vbWhite
End Sub

Private Sub msdep5_GotFocus()
msdep5.BackColor = &H80000016
End Sub

Private Sub msdep5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub msdep5_LostFocus()
msdep5.BackColor = vbWhite
End Sub
Private Sub mshpbdate_GotFocus()
mshpbdate.BackColor = &H80000016
End Sub

Private Sub mshpbdate_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub mshpbdate_LostFocus()
mshpbdate.BackColor = vbWhite
End Sub

Private Sub mskbday_GotFocus()
'mskbday.BackColor = &H80000016
End Sub

Private Sub mskbday_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub mskbday_LostFocus()
'mskbday.BackColor = vbWhite
End Sub

Private Sub mskc1birth_GotFocus()
mskc1birth.BackColor = &H80000016
End Sub

Private Sub mskc1birth_LostFocus()
mskc1birth.BackColor = vbWhite
End Sub

Private Sub mskc2birth_GotFocus()
mskc2birth.BackColor = &H80000016
End Sub

Private Sub mskc2birth_LostFocus()
mskc2birth.BackColor = vbWhite
End Sub

Private Sub mskfbirth_GotFocus()
mskfbirth.BackColor = &H80000016
End Sub

Private Sub mskfbirth_LostFocus()
mskfbirth.BackColor = vbWhite
End Sub

Private Sub mskmbirth_GotFocus()
mskmbirth.BackColor = &H80000016
End Sub

Private Sub mskmbirth_LostFocus()
mskmbirth.BackColor = vbWhite
End Sub

Private Sub msks1birth_GotFocus()
msks1birth.BackColor = &H80000016
End Sub

Private Sub msks1birth_LostFocus()
msks1birth.BackColor = vbWhite
End Sub

Private Sub msks2birth_GotFocus()
msks2birth.BackColor = &H80000016
End Sub

Private Sub msks2birth_LostFocus()
msks2birth.BackColor = vbWhite
End Sub

Private Sub msksbirth_GotFocus()
msksbirth.BackColor = &H80000016
End Sub

Private Sub msksbirth_LostFocus()
msksbirth.BackColor = vbWhite
End Sub

Private Sub mskbirthdate_GotFocus()
mskbirthdate.BackColor = &H80000016
End Sub

Private Sub mskbirthdate_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub mskbirthdate_LostFocus()
mskbirthdate.BackColor = vbWhite
End Sub

Private Sub mskdateapplied_GotFocus()
mskdateapplied.BackColor = &H80000016
End Sub

Private Sub mskdateapplied_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub mskdateapplied_LostFocus()
mskdateapplied.BackColor = vbWhite
End Sub

Private Sub mskdatehired_GotFocus()
mskdatehired.BackColor = &H80000016
End Sub

Private Sub mskdatehired_LostFocus()
mskdatehired.BackColor = vbWhite
End Sub

Private Sub mskdateprob_GotFocus()
mskdateprob.BackColor = &H80000016
End Sub

Private Sub mskdateprob_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub mskdateprob_LostFocus()
mskdateprob.BackColor = vbWhite
End Sub

Private Sub mskdatereg_GotFocus()
mskdatereg.BackColor = &H80000016
End Sub

Private Sub mskdatereg_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub mskdatereg_LostFocus()
mskdatereg.BackColor = vbWhite
End Sub

Private Sub mskdateres_GotFocus()
mskdateres.BackColor = &H80000016
End Sub

Private Sub mskdateres_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub mskdateres_LostFocus()
mskdateres.BackColor = vbWhite
End Sub

Private Sub mskdep1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub mskdep2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub mskdep3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub mskdep4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub mskdep4_LostFocus()
'mskdep4.BackColor = vbWhite
End Sub
Private Sub mskendof_GotFocus()
mskendof.BackColor = &H80000016
End Sub

Private Sub mskendof_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub mskendof_LostFocus()
mskendof.BackColor = vbWhite
End Sub

Private Sub mskhpbdate_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub
Private Sub mskissue_GotFocus()
mskissue.BackColor = &H80000016
End Sub

Private Sub mskissue_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub mskissue_LostFocus()
mskissue.BackColor = vbWhite
End Sub

Private Sub msksss_GotFocus()
msksss.BackColor = &H80000016
End Sub

Private Sub msksss_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub msksss_LostFocus()
msksss.BackColor = vbWhite
End Sub

Private Sub msktax_GotFocus()
msktax.BackColor = &H80000016
End Sub

Private Sub msktax_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub msktax_LostFocus()
msktax.BackColor = vbWhite
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Call TabResizing
End Sub
Private Sub add()
cmdnew.Enabled = False
cmdedit.Enabled = False
cmdupdate.Enabled = True
cmdundo.Enabled = True
'cmdsearch.Enabled = False
'cmdhired.Enabled = False
cmddelete.Enabled = False
cmdfirst.Enabled = False
cmdprev.Enabled = False
cmdnext.Enabled = False
cmdlast.Enabled = False

End Sub


Private Sub txtacctno_GotFocus()
txtacctno.BackColor = &H80000016
End Sub

Private Sub txtacctno_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtacctno_LostFocus()
txtacctno.BackColor = vbWhite
End Sub

Private Sub txtActype_GotFocus()
txtActype.BackColor = &H80000016
End Sub

Private Sub txtActype_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtActype_LostFocus()
txtActype.BackColor = vbWhite
End Sub

Private Sub txtaddress1_GotFocus()
txtaddress1.BackColor = &H80000016
End Sub

Private Sub txtaddress1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtaddress1_LostFocus()
txtaddress1.BackColor = vbWhite
End Sub

Private Sub txtEmployeeNo_GotFocus()
txtEmployeeNo.BackColor = &H80000016
End Sub

Private Sub txtEmployeeNo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtEmployeeNo_LostFocus()
Dim rstemp As ADODB.Recordset
Dim str As String

If (IsNull(txtEmployeeNo.Text)) Or (txtEmployeeNo.Text = "") Or (Val(txtEmployeeNo.Text) = 0) Then
Else
    If addflag = True Then
        str = "Select * from HREmployees where convert(int, EmployeeNum) = " & Val(txtEmployeeNo.Text)
        Set rstemp = cnDCPI.Execute(str)
        If rstemp.EOF = False Then
            MsgBox "Existing Employee Number [" & rstemp!employeenum & "] " & vbCrLf & vbCrLf & "FIRSTNAME : " & rstemp!FirstName & vbCrLf & "SURNAME : " & rstemp!surname, vbInformation, "DCPI Information System"
            txtEmployeeNo.SetFocus
        End If
    End If
End If
End Sub

Private Sub txthouseno_GotFocus()
txthouseno.BackColor = &H80000016
End Sub

Private Sub txtphilhealth_GotFocus()
txtphilhealth.BackColor = &H80000016
End Sub

Private Sub txtphilhealth_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtphilhealth_LostFocus()
txtphilhealth.BackColor = vbWhite
End Sub

Private Sub txtstreet_GotFocus()
txtstreet.BackColor = &H80000016
End Sub
Private Sub txtcity_GotFocus()
txtcity.BackColor = &H80000016
End Sub
Private Sub txtcity_LostFocus()
txtcity.BackColor = vbWhite
End Sub

Private Sub txtstreet_LostFocus()
txtstreet.BackColor = vbWhite
End Sub
Private Sub txtbarangay_GotFocus()
txtbarangay.BackColor = &H80000016
End Sub
Private Sub txtdistrict_GotFocus()
txtdistrict.BackColor = &H80000016
End Sub
Private Sub txtprovince_GotFocus()
txtprovince.BackColor = &H80000016
End Sub
Private Sub txtprovince_LostFocus()
txtprovince.BackColor = vbWhite
End Sub

Private Sub txtdistrict_LostFocus()
txtdistrict.BackColor = vbWhite
End Sub
Private Sub txtdistrict_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub
Private Sub txtprovince_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtbarangay_LostFocus()
txtbarangay.BackColor = vbWhite
End Sub
Private Sub txtbarangay_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub
Private Sub txtcity_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txthouseno_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub
Private Sub txtstreet_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txthouseno_LostFocus()
txthouseno.BackColor = vbWhite
End Sub

Private Sub txtage_GotFocus()
txtage.BackColor = &H80000016
End Sub

Private Sub txtage_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtage_LostFocus()
txtage.BackColor = vbWhite
End Sub

Private Sub txtappstatus_GotFocus()
txtappstatus.BackColor = &H80000016
End Sub

Private Sub txtappstatus_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtappstatus_LostFocus()
txtappstatus.BackColor = vbWhite
End Sub

Private Sub txtbirthplace_GotFocus()
txtbirthplace.BackColor = &H80000016
End Sub

Private Sub txtbirthplace_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtbirthplace_LostFocus()
txtbirthplace.BackColor = vbWhite
End Sub

Private Sub txtbro1name_GotFocus()
txtbro1name.BackColor = &H80000016
End Sub

Private Sub txtbro1name_LostFocus()
txtbro1name.BackColor = vbWhite
End Sub

Private Sub txtbro2name_GotFocus()
txtbro2name.BackColor = &H80000016
End Sub

Private Sub txtbro2name_LostFocus()
txtbro2name.BackColor = vbWhite
End Sub

Private Sub txtchild1name_GotFocus()
txtchild1name.BackColor = &H80000016
End Sub

Private Sub txtchild1name_LostFocus()
txtchild1name.BackColor = vbWhite
End Sub

Private Sub txtchild2name_GotFocus()
txtchild2name.BackColor = &H80000016
End Sub

Private Sub txtchild2name_LostFocus()
txtchild2name.BackColor = vbWhite
End Sub

Private Sub txtcitizen_GotFocus()
txtcitizen.BackColor = &H80000016
End Sub

Private Sub txtcitizen_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtcitizen_LostFocus()
txtcitizen.BackColor = vbWhite
End Sub

Private Sub txtcityaddress_GotFocus()
txtcityaddress.BackColor = &H80000016
End Sub

Private Sub txtcityaddress_LostFocus()
txtcityaddress.BackColor = vbWhite
End Sub

Private Sub txtcollegecourse_GotFocus()
txtcollegecourse.BackColor = &H80000016
End Sub

Private Sub txtcollegecourse_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtcollegecourse_LostFocus()
txtcollegecourse.BackColor = vbWhite
End Sub

Private Sub txtcollegemajor_GotFocus()
txtcollegemajor.BackColor = &H80000016
End Sub

Private Sub txtcollegemajor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtcollegemajor_LostFocus()
txtcollegemajor.BackColor = vbWhite
End Sub

Private Sub txtcollegesch_GotFocus()
txtcollegesch.BackColor = &H80000016
End Sub

Private Sub txtcollegesch_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtcollegesch_LostFocus()
txtcollegesch.BackColor = vbWhite
End Sub

Private Sub txtdateapplied_GotFocus()
'txtdateapplied.BackColor = &H80000016
End Sub

Private Sub txtdateapplied_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtdateapplied_LostFocus()
'txtdateapplied.BackColor = vbWhite
End Sub

Private Sub txtdatecollege_GotFocus()
txtdatecollege.BackColor = &H80000016
End Sub

Private Sub txtdatecollege_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdatecollege_LostFocus()
txtdatecollege.BackColor = vbWhite
End Sub

Private Sub txtdatehired_GotFocus()
'txtdatehired.BackColor = &H80000016
End Sub

Private Sub txtdatehired_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdatehired_LostFocus()
'txtdatehired.BackColor = vbWhite
End Sub

Private Sub txtdatepost_GotFocus()
txtdatepost.BackColor = &H80000016
End Sub

Private Sub txtdatepost_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdatepost_LostFocus()
txtdatepost.BackColor = vbWhite
End Sub

Private Sub txtdateprimary_GotFocus()
txtdateprimary.BackColor = &H80000016
End Sub

Private Sub txtdateprimary_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdateprimary_LostFocus()
txtdateprimary.BackColor = vbWhite
End Sub

Private Sub txtdatesec_GotFocus()
txtdatesec.BackColor = &H80000016
End Sub

Private Sub txtdatesec_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdatesec_LostFocus()
txtdatesec.BackColor = vbWhite
End Sub

Private Sub txtdateseminar1_GotFocus()
txtdateseminar1.BackColor = &H80000016
End Sub

Private Sub txtdateseminar1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdateseminar1_LostFocus()
txtdateseminar1.BackColor = vbWhite
End Sub

Private Sub txtdateseminar2_GotFocus()
txtdateseminar2.BackColor = &H80000016
End Sub

Private Sub txtdateseminar2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdateseminar2_LostFocus()
txtdateseminar2.BackColor = vbWhite
End Sub

Private Sub txtdateseminar3_GotFocus()
txtdateseminar3.BackColor = &H80000016
End Sub

Private Sub txtdateseminar3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdateseminar3_LostFocus()
txtdateseminar3.BackColor = vbWhite
End Sub

Private Sub txtdep1_GotFocus()
txtdep1.BackColor = &H80000016
End Sub

Private Sub txtdep1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdep1_LostFocus()
txtdep1.BackColor = vbWhite
End Sub

Private Sub txtdep2_GotFocus()
txtdep2.BackColor = &H80000016
End Sub

Private Sub txtdep2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdep2_LostFocus()
txtdep2.BackColor = vbWhite
End Sub

Private Sub txtdep3_GotFocus()
txtdep3.BackColor = &H80000016
End Sub

Private Sub txtdep3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdep3_LostFocus()
txtdep3.BackColor = vbWhite
End Sub
Private Sub txtdep5_LostFocus()
txtdep5.BackColor = vbWhite
End Sub

Private Sub txtdep4_GotFocus()
txtdep4.BackColor = &H80000016
End Sub
Private Sub txtdep5_GotFocus()
txtdep5.BackColor = &H80000016
End Sub
Private Sub txtdep5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdep4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdep4_LostFocus()
txtdep4.BackColor = vbWhite
End Sub

Private Sub txtdialect1_GotFocus()
txtdialect1.BackColor = &H80000016
End Sub

Private Sub txtdialect1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdialect1_LostFocus()
txtdialect1.BackColor = vbWhite
End Sub

Private Sub txtdialect2_GotFocus()
txtdialect2.BackColor = &H80000016
End Sub

Private Sub txtdialect2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdialect2_LostFocus()
txtdialect2.BackColor = vbWhite
End Sub

Private Sub txtdialect3_GotFocus()
txtdialect3.BackColor = &H80000016
End Sub

Private Sub txtdialect3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdialect3_LostFocus()
txtdialect3.BackColor = vbWhite
End Sub

Private Sub txtdtrexempt_GotFocus()
txtdtrexempt.BackColor = &H80000016
End Sub

Private Sub txtdtrexempt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtdtrexempt_LostFocus()
txtdtrexempt.BackColor = vbWhite
End Sub

Private Sub txtemergency_GotFocus()
txtemergency.BackColor = &H80000016
End Sub

Private Sub txtemergency_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtemergency_LostFocus()
txtemergency.BackColor = vbWhite
End Sub

Private Sub txteraddress_GotFocus()
txteraddress.BackColor = &H80000016
End Sub

Private Sub txteraddress_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txteraddress_LostFocus()
txteraddress.BackColor = vbWhite
End Sub

Private Sub txtermobile_GotFocus()
txtermobile.BackColor = &H80000016
End Sub

Private Sub txtermobile_LostFocus()
txtermobile.BackColor = vbWhite
End Sub

Private Sub txtertel_GotFocus()
txtertel.BackColor = &H80000016
End Sub

Private Sub txtertel_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtertel_LostFocus()
txtertel.BackColor = vbWhite
End Sub

Private Sub txtexam_GotFocus()
txtexam.BackColor = &H80000016
End Sub

Private Sub txtexam_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtexam_LostFocus()
txtexam.BackColor = vbWhite
End Sub

'Private Sub txteyecolor_GotFocus()
'txteyecolor.BackColor = &H80000016
'End Sub
'
'Private Sub txteyecolor_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(Chr(KeyAscii))
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'    End If
'End Sub
'
'Private Sub txteyecolor_LostFocus()
'txteyecolor.BackColor = vbWhite
'End Sub

Private Sub txtfaddress_GotFocus()
txtfaddress.BackColor = &H80000016
End Sub

Private Sub txtfaddress_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtfaddress_LostFocus()
txtfaddress.BackColor = vbWhite
End Sub

Private Sub txtfathername_GotFocus()
txtfathername.BackColor = &H80000016
End Sub

Private Sub txtfathername_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtfathername_LostFocus()
txtfathername.BackColor = vbWhite
End Sub

Private Sub txtfirstchoice_GotFocus()
txtfirstchoice.BackColor = &H80000016
End Sub

Private Sub txtfirstchoice_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub txtfirstchoice_LostFocus()
txtfirstchoice.BackColor = vbWhite
UCase (txtfirstchoice.Text)
End Sub

Private Sub txtfirstname_GotFocus()
txtfirstname.BackColor = &H80000016
End Sub

Private Sub txtfirstname_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtfirstname_LostFocus()
txtfirstname.BackColor = vbWhite
End Sub

Private Sub txtfjob_GotFocus()
txtfjob.BackColor = &H80000016
End Sub

Private Sub txtfjob_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtfjob_LostFocus()
txtfjob.BackColor = vbWhite
End Sub

Private Sub txtfrom1_GotFocus()
txtfrom1.BackColor = &H80000016
End Sub

Private Sub txtfrom1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtfrom1_LostFocus()
txtfrom1.BackColor = vbWhite
End Sub

Private Sub txtfrom2_GotFocus()
txtfrom2.BackColor = &H80000016
End Sub

Private Sub txtfrom2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtfrom2_LostFocus()
txtfrom2.BackColor = vbWhite
End Sub

Private Sub txtfrom3_GotFocus()
txtfrom3.BackColor = &H80000016
End Sub

Private Sub txtfrom3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtfrom3_LostFocus()
txtfrom3.BackColor = vbWhite
End Sub

'Private Sub txthaircolor_GotFocus()
'txthaircolor.BackColor = &H80000016
'End Sub
'
'Private Sub txthaircolor_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(Chr(KeyAscii))
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'    End If
'End Sub
'
'Private Sub txthaircolor_LostFocus()
'txthaircolor.BackColor = vbWhite
'End Sub

Private Sub txtheight_GotFocus()
txtheight.BackColor = &H80000016
End Sub

Private Sub txtheight_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtheight_LostFocus()
txtheight.BackColor = vbWhite
End Sub

Private Sub txthfdep_GotFocus()
txthfdep.BackColor = &H80000016
End Sub

Private Sub txthfdep_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txthfdep_LostFocus()
txthfdep.BackColor = vbWhite
End Sub

Private Sub txthprelate_GotFocus()
txthprelate.BackColor = &H80000016
End Sub

Private Sub txthprelate_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txthprelate_LostFocus()
txthprelate.BackColor = vbWhite
End Sub

Private Sub txthrswork_GotFocus()
txthrswork.BackColor = &H80000016
End Sub

Private Sub txthrswork_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txthrswork_LostFocus()
txthrswork.BackColor = vbWhite
End Sub

Private Sub txtinterview_GotFocus()
txtinterview.BackColor = &H80000016
End Sub

Private Sub txtinterview_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtinterview_LostFocus()
txtinterview.BackColor = vbWhite
End Sub

'Private Sub txtissueat_GotFocus()
'txtissueat.BackColor = &H80000016
'End Sub
'
'Private Sub txtissueat_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(Chr(KeyAscii))
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'    End If
'End Sub
'
'Private Sub txtissueat_LostFocus()
'txtissueat.BackColor = vbWhite
'End Sub

Private Sub txtmaddress_GotFocus()
txtmaddress.BackColor = &H80000016
End Sub

Private Sub txtmaddress_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtmaddress_LostFocus()
txtmaddress.BackColor = vbWhite
End Sub

Private Sub txtmedical_GotFocus()
txtmedical.BackColor = &H80000016
End Sub

Private Sub txtmedical_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtmedical_LostFocus()
txtmedical.BackColor = vbWhite
End Sub

Private Sub txtmiddlename_GotFocus()
txtmiddlename.BackColor = &H80000016
End Sub

Private Sub txtmiddlename_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtmiddlename_LostFocus()
txtmiddlename.BackColor = vbWhite
End Sub

Private Sub txtmjob_GotFocus()
txtmjob.BackColor = &H80000016
End Sub

Private Sub txtmjob_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtmjob_LostFocus()
txtmjob.BackColor = vbWhite
End Sub

Private Sub txtmothername_GotFocus()
txtmothername.BackColor = &H80000016
End Sub

Private Sub txtmothername_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtmothername_LostFocus()
txtmothername.BackColor = vbWhite
End Sub

Private Sub txtnameofcompany1_GotFocus()
txtnameofcompany1.BackColor = &H80000016
End Sub

Private Sub txtnameofcompany1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtnameofcompany1_LostFocus()
txtnameofcompany1.BackColor = vbWhite
End Sub

Private Sub txtnameofcompany2_GotFocus()
txtnameofcompany2.BackColor = &H80000016
End Sub

Private Sub txtnameofcompany2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtnameofcompany2_LostFocus()
txtnameofcompany2.BackColor = vbWhite
End Sub

Private Sub txtnameofcompany3_GotFocus()
txtnameofcompany3.BackColor = &H80000016
End Sub

Private Sub txtnameofcompany3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtnameofcompany3_LostFocus()
txtnameofcompany3.BackColor = vbWhite
End Sub

Private Sub txtnbi_GotFocus()
txtnbi.BackColor = &H80000016
End Sub

Private Sub txtnbi_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtnbi_LostFocus()
txtnbi.BackColor = vbWhite
End Sub

Private Sub txtnickname_GotFocus()
txtnickname.BackColor = &H80000016
End Sub

Private Sub txtnickname_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtnickname_LostFocus()
txtnickname.BackColor = vbWhite
End Sub

Private Sub txtnolate_GotFocus()
txtnolate.BackColor = &H80000016
End Sub

Private Sub txtnolate_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtnolate_LostFocus()
txtnolate.BackColor = vbWhite
End Sub

Private Sub txtnoot_GotFocus()
txtnoot.BackColor = &H80000016
End Sub

Private Sub txtnoot_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtnoot_LostFocus()
txtnoot.BackColor = vbWhite
End Sub

Private Sub txtnoundertime_GotFocus()
txtnoundertime.BackColor = &H80000016
End Sub

Private Sub txtnoundertime_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtnoundertime_LostFocus()
txtnoundertime.BackColor = vbWhite
End Sub

Private Sub txtnumdep_GotFocus()
txtnumdep.BackColor = &H80000016
End Sub

Private Sub txtnumdep_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtnumdep_LostFocus()
txtnumdep.BackColor = vbWhite
End Sub

Private Sub txtpagibig_GotFocus()
txtpagibig.BackColor = &H80000016
End Sub

Private Sub txtpagibig_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtpagibig_LostFocus()
txtpagibig.BackColor = vbWhite
End Sub

'Private Sub txtphcode_GotFocus()
'txtphcode.BackColor = &H80000016
'End Sub
'
'Private Sub txtphcode_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(Chr(KeyAscii))
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'    End If
'End Sub
'
'Private Sub txtphcode_LostFocus()
'txtphcode.BackColor = vbWhite
'End Sub

Private Sub txtpos1_GotFocus()
txtpos1.BackColor = &H80000016
End Sub

Private Sub txtpos1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtpos1_LostFocus()
txtpos1.BackColor = vbWhite
End Sub

Private Sub txtpos2_GotFocus()
txtpos2.BackColor = &H80000016
End Sub

Private Sub txtpos2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtpos2_LostFocus()
txtpos2.BackColor = vbWhite
End Sub

Private Sub txtpos3_GotFocus()
txtpos3.BackColor = &H80000016
End Sub

Private Sub txtpos3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtpos3_LostFocus()
txtpos3.BackColor = vbWhite
End Sub

Private Sub txtpostcourse_GotFocus()
txtpostcourse.BackColor = &H80000016
End Sub

Private Sub txtpostcourse_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtpostcourse_LostFocus()
txtpostcourse.BackColor = vbWhite
End Sub

Private Sub txtpostsch_GotFocus()
txtpostsch.BackColor = &H80000016
End Sub

Private Sub txtpostsch_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtpostsch_LostFocus()
txtpostsch.BackColor = vbWhite
End Sub

Private Sub txtprimarysch_GotFocus()
txtprimarysch.BackColor = &H80000016
End Sub

Private Sub txtprimarysch_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtprimarysch_LostFocus()
txtprimarysch.BackColor = vbWhite
End Sub

Private Sub txtprovinceaddress_GotFocus()
txtprovinceaddress.BackColor = &H80000016
End Sub

Private Sub txtprovinceaddress_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtprovinceaddress_LostFocus()
txtprovinceaddress.BackColor = vbWhite
End Sub

Private Sub txtreason1_GotFocus()
txtreason1.BackColor = &H80000016
End Sub

Private Sub txtreason1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtreason1_LostFocus()
txtreason1.BackColor = vbWhite
End Sub

Private Sub txtreason2_GotFocus()
txtreason2.BackColor = &H80000016
End Sub

Private Sub txtreason2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtreason2_LostFocus()
txtreason2.BackColor = vbWhite
End Sub

Private Sub txtreason3_GotFocus()
txtreason3.BackColor = &H80000016
End Sub

Private Sub txtreason3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtreason3_LostFocus()
txtreason3.BackColor = vbWhite
End Sub

Private Sub txtreasonresign_GotFocus()
txtreasonresign.BackColor = &H80000016
End Sub

Private Sub txtreasonresign_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtreasonresign_LostFocus()
txtreasonresign.BackColor = vbWhite
End Sub

Private Sub txtrefaddress1_GotFocus()
txtrefaddress1.BackColor = &H80000016
End Sub

Private Sub txtrefaddress1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefaddress1_LostFocus()
txtrefaddress1.BackColor = vbWhite
End Sub

Private Sub txtrefaddress2_GotFocus()
txtrefaddress2.BackColor = &H80000016
End Sub

Private Sub txtrefaddress2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefaddress2_LostFocus()
txtrefaddress2.BackColor = vbWhite
End Sub

Private Sub txtrefaddress3_GotFocus()
txtrefaddress3.BackColor = &H80000016
End Sub

Private Sub txtrefaddress3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefaddress3_LostFocus()
txtrefaddress3.BackColor = vbWhite
End Sub

Private Sub txtrefjob1_GotFocus()
txtrefjob1.BackColor = &H80000016
End Sub

Private Sub txtrefjob1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefjob1_LostFocus()
txtrefjob1.BackColor = vbWhite
End Sub

Private Sub txtrefjob2_GotFocus()
txtrefjob2.BackColor = &H80000016
End Sub

Private Sub txtrefjob2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefjob2_LostFocus()
txtrefjob2.BackColor = vbWhite
End Sub

Private Sub txtrefjob3_GotFocus()
txtrefjob3.BackColor = &H80000016
End Sub

Private Sub txtrefjob3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefjob3_LostFocus()
txtrefjob3.BackColor = vbWhite
End Sub

Private Sub txtrefname1_GotFocus()
txtrefname1.BackColor = &H80000016
End Sub

Private Sub txtrefname1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefname1_LostFocus()
txtrefname1.BackColor = vbWhite
End Sub

Private Sub txtrefname2_GotFocus()
txtrefname2.BackColor = &H80000016
End Sub

Private Sub txtrefname2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefname2_LostFocus()
txtrefname2.BackColor = vbWhite
End Sub

Private Sub txtrefname3_GotFocus()
txtrefname3.BackColor = &H80000016
End Sub

Private Sub txtrefname3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefname3_LostFocus()
txtrefname3.BackColor = vbWhite
End Sub

Private Sub txtrefphone1_GotFocus()
txtrefphone1.BackColor = &H80000016
End Sub

Private Sub txtrefphone1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefphone1_LostFocus()
txtrefphone1.BackColor = vbWhite
End Sub

Private Sub txtrefphone2_GotFocus()
txtrefphone2.BackColor = &H80000016
End Sub

Private Sub txtrefphone2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefphone2_LostFocus()
txtrefphone2.BackColor = vbWhite
End Sub

Private Sub txtrefphone3_GotFocus()
txtrefphone3.BackColor = &H80000016
End Sub

Private Sub txtrefphone3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefphone3_LostFocus()
txtrefphone3.BackColor = vbWhite
End Sub

Private Sub txtrefrel1_GotFocus()
txtrefrel1.BackColor = &H80000016
End Sub

Private Sub txtrefrel1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefrel1_LostFocus()
txtrefrel1.BackColor = vbWhite
End Sub

Private Sub txtrefrel2_GotFocus()
txtrefrel2.BackColor = &H80000016
End Sub

Private Sub txtrefrel2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefrel2_LostFocus()
txtrefrel2.BackColor = vbWhite
End Sub

Private Sub txtrefrel3_GotFocus()
txtrefrel3.BackColor = &H80000016
End Sub

Private Sub txtrefrel3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtrefrel3_LostFocus()
txtrefrel3.BackColor = vbWhite
End Sub

Private Sub txtreligion_GotFocus()
txtreligion.BackColor = &H80000016
End Sub

Private Sub txtreligion_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtreligion_LostFocus()
txtreligion.BackColor = vbWhite
End Sub

Private Sub txtsecondchoice_GotFocus()
txtsecondchoice.BackColor = &H80000016
End Sub

Private Sub txtsecondchoice_LostFocus()
txtsecondchoice.BackColor = vbWhite
End Sub

Private Sub txtsis1name_GotFocus()
txtsis1name.BackColor = &H80000016
End Sub

Private Sub txtsis1name_LostFocus()
txtsis1name.BackColor = vbWhite
End Sub

Private Sub txtsis2name_GotFocus()
txtsis2name.BackColor = &H80000016
End Sub

Private Sub txtsis2name_LostFocus()
txtsis2name.BackColor = vbWhite
End Sub

'Private Sub txtrescer_GotFocus()
'txtrescer.BackColor = &H80000016
'End Sub
'
'Private Sub txtrescer_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(Chr(KeyAscii))
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'    End If
'End Sub
'
'Private Sub txtrescer_LostFocus()
'txtrescer.BackColor = vbWhite
'End Sub

Private Sub txtsecsch_GotFocus()
txtsecsch.BackColor = &H80000016
End Sub

Private Sub txtsecsch_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtsecsch_LostFocus()
txtsecsch.BackColor = vbWhite
End Sub

Private Sub txtseminar1_GotFocus()
txtseminar1.BackColor = &H80000016
End Sub

Private Sub txtseminar1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtseminar1_LostFocus()
txtseminar1.BackColor = vbWhite
End Sub

Private Sub txtseminar2_GotFocus()
txtseminar2.BackColor = &H80000016
End Sub

Private Sub txtseminar2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtseminar2_LostFocus()
txtseminar2.BackColor = vbWhite
End Sub

Private Sub txtseminar3_GotFocus()
txtseminar3.BackColor = &H80000016
End Sub

Private Sub txtseminar3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtseminar3_LostFocus()
txtseminar3.BackColor = vbWhite
End Sub

Private Sub txtsjob_GotFocus()
txtsjob.BackColor = &H80000016
End Sub

Private Sub txtsjob_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtsjob_LostFocus()
txtsjob.BackColor = vbWhite
End Sub

Private Sub txtspouseaddress_GotFocus()
txtspouseaddress.BackColor = &H80000016
End Sub

Private Sub txtspouseaddress_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtspouseaddress_LostFocus()
txtspouseaddress.BackColor = vbWhite
End Sub

Private Sub txtspousename_GotFocus()
txtspousename.BackColor = &H80000016
End Sub

Private Sub txtspousename_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtspousename_LostFocus()
txtspousename.BackColor = vbWhite
End Sub

Private Sub txtsuperior1_GotFocus()
txtsuperior1.BackColor = &H80000016
End Sub

Private Sub txtsuperior1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtsuperior1_LostFocus()
txtsuperior1.BackColor = vbWhite
End Sub

Private Sub txtsuperior2_GotFocus()
txtsuperior2.BackColor = &H80000016
End Sub

Private Sub txtsuperior2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtsuperior2_LostFocus()
txtsuperior2.BackColor = vbWhite
End Sub

Private Sub txtsuperior3_GotFocus()
txtsuperior3.BackColor = &H80000016
End Sub

Private Sub txtsuperior3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtsuperior3_LostFocus()
txtsuperior3.BackColor = vbWhite
End Sub

Private Sub txtsupervisor_GotFocus()
txtsupervisor.BackColor = &H80000016
End Sub

Private Sub txtsupervisor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtsupervisor_LostFocus()
txtsupervisor.BackColor = vbWhite
End Sub

Private Sub txtsurname_GotFocus()
txtsurname.BackColor = &H80000016
End Sub

Private Sub txtsurname_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtsurname_LostFocus()
txtsurname.BackColor = vbWhite
End Sub

Private Sub txttel1_GotFocus()
txttel1.BackColor = &H80000016
End Sub

Private Sub txttel1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txttel1_LostFocus()
txttel1.BackColor = vbWhite
End Sub

Private Sub txttel2_GotFocus()
txttel2.BackColor = &H80000016
End Sub

Private Sub txttel2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txttel2_LostFocus()
txttel2.BackColor = vbWhite
End Sub

Private Sub txttel3_GotFocus()
txttel3.BackColor = &H80000016
End Sub

Private Sub txttel3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txttel3_LostFocus()
txttel3.BackColor = vbWhite
End Sub

Private Sub txttel4_GotFocus()
txttel4.BackColor = &H80000016
End Sub

Private Sub txttel4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txttel4_LostFocus()
txttel4.BackColor = vbWhite
End Sub

Private Sub txtto1_GotFocus()
txtto1.BackColor = &H80000016
End Sub

Private Sub txtto1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtto1_LostFocus()
txtto1.BackColor = vbWhite
End Sub

Private Sub txtto2_GotFocus()
txtto2.BackColor = &H80000016
End Sub

Private Sub txtto2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtto2_LostFocus()
txtto2.BackColor = vbWhite
End Sub

Private Sub txtto3_GotFocus()
txtto3.BackColor = &H80000016
End Sub

Private Sub txtto3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtto3_LostFocus()
txtto3.BackColor = vbWhite
End Sub

Private Sub txtweight_GotFocus()
txtweight.BackColor = &H80000016
End Sub

Private Sub txtweight_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtweight_LostFocus()
txtweight.BackColor = vbWhite
End Sub

Private Sub update()
'***** 09/28/04 ******
txtEmployeeNo.Enabled = False
txthfdep.Enabled = False
mskhpbdate.Enabled = False
txthprelate.Enabled = False
txtprimarysch.Enabled = False
txtdateprimary.Enabled = False
txtnickname.Enabled = False
txtsurname.Enabled = False
txtfirstname.Enabled = False
txtmiddlename.Enabled = False
txtfirstchoice.Enabled = False
mskBday.Enabled = False
txtage.Enabled = False
txtaddress1.Enabled = False
txthouseno.Enabled = False
txtstreet.Enabled = False
txtbarangay.Enabled = False
txtdistrict.Enabled = False
txtcity.Enabled = False
txtprovince.Enabled = False
txtprovinceaddress.Enabled = False
txtemergency.Enabled = False
txteraddress.Enabled = False
txtertel.Enabled = False
optMale.Enabled = False
optFemale.Enabled = False
txtheight.Enabled = False
txtweight.Enabled = False
txttel1.Enabled = False
txttel2.Enabled = False
txttel3.Enabled = False
txttel4.Enabled = False
cmbcivilstatus.Enabled = False
'DTdateapplied.Enabled = False
'09/29/04 **************************
'txtdateapplied.Enabled = False
mskdateapplied.Enabled = False

' 11/03/04 **********
txtphilhealth.Enabled = False
' 11/03/04 **********
txtappstatus.Enabled = False
txtexam.Enabled = False
txtinterview.Enabled = False
txtmedical.Enabled = False
msksss.Enabled = False
'txtphcode.Enabled = False
txtpagibig.Enabled = False
txtacctno.Enabled = False
txtcitizen.Enabled = False
txtbirthplace.Enabled = False
'txthaircolor.Enabled = False
'txteyecolor.Enabled = False
txtreligion.Enabled = False
msktax.Enabled = False
txtpostsch.Enabled = False
txtdatepost.Enabled = False
txtsecsch.Enabled = False
txtdatesec.Enabled = False
txtcollegesch.Enabled = False
txtdatecollege.Enabled = False
txtcollegecourse.Enabled = False
txtcollegemajor.Enabled = False
txtdatepost.Enabled = False
txtpostcourse.Enabled = False
txtfathername.Enabled = False
txtfaddress.Enabled = False
txtfjob.Enabled = False
txtmothername.Enabled = False
txtmaddress.Enabled = False
txtmjob.Enabled = False
txtspousename.Enabled = False
txtspouseaddress.Enabled = False
txtsjob.Enabled = False
txtnbi.Enabled = False
'txtrescer.Enabled = False
'txtissueat.Enabled = False
'dtissueon.Enabled = False
mskissue.Enabled = False

txtdep1.Enabled = False
txtdep2.Enabled = False
txtdep3.Enabled = False
txtdep4.Enabled = False
txtdep5.Enabled = False

'makiss 09/28/04 **************
'mskdep1.Enabled = False
msdep1.Enabled = False
'mskdep2.Enabled = False
msdep2.Enabled = False
'mskdep3.Enabled = False
msdep3.Enabled = False
'mskdep4.Enabled = False
msdep4.Enabled = False
'mskdep5.Enabled = False
msdep5.Enabled = False
'makiss 09/28/04 **************

txtnumdep.Enabled = False
txtdialect1.Enabled = False
txtdialect2.Enabled = False
txtdialect3.Enabled = False
txtseminar1.Enabled = False
txtseminar2.Enabled = False
txtseminar3.Enabled = False
txtdateseminar1.Enabled = False
txtdateseminar2.Enabled = False
txtdateseminar3.Enabled = False
txtnameofcompany1.Enabled = False
txtnameofcompany2.Enabled = False
txtnameofcompany3.Enabled = False
txtfrom1.Enabled = False
txtfrom2.Enabled = False
txtfrom3.Enabled = False
txtto1.Enabled = False
txtto2.Enabled = False
txtto3.Enabled = False
txtpos1.Enabled = False
txtpos2.Enabled = False
txtpos3.Enabled = False
txtreason1.Enabled = False
txtreason2.Enabled = False
txtreason3.Enabled = False
txtsuperior1.Enabled = False
txtsuperior2.Enabled = False
txtsuperior3.Enabled = False
txtrefname1.Enabled = False
txtrefname2.Enabled = False
txtrefname3.Enabled = False
txtrefjob1.Enabled = False
txtrefjob2.Enabled = False
txtrefjob3.Enabled = False
txtrefaddress1.Enabled = False
txtrefaddress2.Enabled = False
txtrefaddress3.Enabled = False
txtrefphone1.Enabled = False
txtrefphone2.Enabled = False
txtrefphone3.Enabled = False
txtrefrel1.Enabled = False
txtrefrel2.Enabled = False
txtrefrel3.Enabled = False
'09/29/04 ******************
'dtdateprob.Enabled = False
mskdateprob.Enabled = False
mskdatereg.Enabled = False
mskdateres.Enabled = False
txtdatehired.Enabled = False
'dtdatereg.Enabled = False
'dtdateresign.Enabled = False
mskdateres.Enabled = False
mskissue.Enabled = False
mskendof.Enabled = False
cmbempstatus.Enabled = False
txtsupervisor.Enabled = False
txtreasonresign.Enabled = False
'dtendcontract.Enabled = False
mskendof.Enabled = False
txtdtrexempt.Enabled = False
txthrswork.Enabled = False
txtnolate.Enabled = False
txtnoundertime.Enabled = False
txtActype.Enabled = False
txtnoot.Enabled = False
cmbdepartment.Enabled = False
mskdep1.Value = 0
mskdatehired.Enabled = False
mskbirthdate.Enabled = False
mshpbdate.Enabled = False
cmdnew.Enabled = True
cmdedit.Enabled = True
cmddelete.Enabled = False
cmdfirst.Enabled = False
cmdnext.Enabled = False
cmdprev.Enabled = False
cmdlast.Enabled = False
cmdupdate.Enabled = False
cmdundo.Enabled = False
End Sub

