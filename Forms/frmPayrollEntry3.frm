VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPayrollEntry3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Transaction Entries"
   ClientHeight    =   7935
   ClientLeft      =   2205
   ClientTop       =   1800
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   10290
   Begin VB.TextBox txtGrossPay1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
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
      Height          =   225
      Left            =   3780
      Locked          =   -1  'True
      TabIndex        =   154
      Text            =   "0.00"
      Top             =   6750
      Width           =   1125
   End
   Begin VB.TextBox txtDaysWork 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Text            =   "0"
      Top             =   2430
      Width           =   1215
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   148
      Text            =   "0.00"
      Top             =   2430
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   3
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   132
      Text            =   "0.00"
      Top             =   3510
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   4
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   131
      Text            =   "0.00"
      Top             =   3870
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   5
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   130
      Text            =   "0.00"
      Top             =   4230
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   6
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   129
      Text            =   "0.00"
      Top             =   4590
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   7
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   128
      Text            =   "0.00"
      Top             =   4950
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   8
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   127
      Text            =   "0.00"
      Top             =   5310
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   9
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   126
      Text            =   "0.00"
      Top             =   5670
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   10
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   125
      Text            =   "0.00"
      Top             =   6030
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   11
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   124
      Text            =   "0.00"
      Top             =   6390
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancelTrans 
      Caption         =   "Cancel Transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   5
      Left            =   8820
      TabIndex        =   33
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveTrans 
      Caption         =   "Save Transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   4
      Left            =   8820
      TabIndex        =   32
      Top             =   6030
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   90
      TabIndex        =   109
      Top             =   7110
      Width           =   10095
      Begin VB.CommandButton cmdlast 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6300
         Picture         =   "frmPayrollEntry3.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   152
         ToolTipText     =   "Last record"
         Top             =   140
         Width           =   705
      End
      Begin VB.CommandButton cmdnext 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5580
         Picture         =   "frmPayrollEntry3.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   151
         ToolTipText     =   "Next record"
         Top             =   140
         Width           =   705
      End
      Begin VB.CommandButton cmdprev 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4860
         Picture         =   "frmPayrollEntry3.frx":062C
         Style           =   1  'Graphical
         TabIndex        =   150
         ToolTipText     =   "Previous record"
         Top             =   140
         Width           =   705
      End
      Begin VB.CommandButton cmdfirst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4140
         Picture         =   "frmPayrollEntry3.frx":0726
         Style           =   1  'Graphical
         TabIndex        =   149
         ToolTipText     =   "First record"
         Top             =   140
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find Employee"
         Height          =   555
         Left            =   3150
         TabIndex        =   37
         Top             =   140
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   555
         Left            =   9000
         TabIndex        =   40
         Top             =   140
         Width           =   975
      End
      Begin VB.CommandButton cmdmovetoMaster 
         Caption         =   "Move to Master"
         Height          =   555
         Left            =   8010
         TabIndex        =   39
         Top             =   140
         Width           =   975
      End
      Begin VB.CommandButton cmdPrintReports 
         Caption         =   "Print Reports"
         Height          =   555
         Left            =   7020
         TabIndex        =   38
         Top             =   140
         Width           =   975
      End
      Begin VB.CommandButton cmdEditTax 
         Caption         =   "Edit Tax"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   1080
         Picture         =   "frmPayrollEntry3.frx":0C58
         TabIndex        =   35
         Top             =   140
         Width           =   975
      End
      Begin VB.CommandButton cmdExcessOt 
         Caption         =   "Excess O.T."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   4
         Left            =   90
         TabIndex        =   34
         Top             =   140
         Width           =   975
      End
      Begin VB.CommandButton cmdFixedRates 
         Caption         =   "Configure Fixed Rates"
         Height          =   555
         Left            =   2070
         TabIndex        =   36
         Top             =   140
         Width           =   1065
      End
   End
   Begin VB.Frame frameExcessOT 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Excess Overtime"
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
      Height          =   1635
      Left            =   5130
      TabIndex        =   108
      Top             =   4320
      Width           =   4965
      Begin VB.TextBox txtExcessOTRegDay 
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
         Height          =   315
         Left            =   2160
         TabIndex        =   27
         Text            =   "0"
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         ForeColor       =   &H00FF0000&
         Height          =   250
         Index           =   12
         Left            =   3780
         TabIndex        =   153
         Text            =   "0.00"
         Top             =   270
         Width           =   1035
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         ForeColor       =   &H00FF0000&
         Height          =   250
         Index           =   13
         Left            =   3780
         TabIndex        =   118
         Text            =   "0.00"
         Top             =   540
         Width           =   1035
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         ForeColor       =   &H00FF0000&
         Height          =   250
         Index           =   14
         Left            =   3780
         TabIndex        =   117
         Text            =   "0.00"
         Top             =   810
         Width           =   1035
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         ForeColor       =   &H00FF0000&
         Height          =   250
         Index           =   15
         Left            =   3780
         TabIndex        =   116
         Text            =   "0.00"
         Top             =   1080
         Width           =   1035
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         ForeColor       =   &H00FF0000&
         Height          =   250
         Index           =   16
         Left            =   3780
         TabIndex        =   115
         Text            =   "0.00"
         Top             =   1350
         Width           =   1035
      End
      Begin VB.TextBox txtExcessOTSunLgl 
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
         Height          =   315
         Left            =   2160
         TabIndex        =   29
         Text            =   "0"
         Top             =   720
         Width           =   1035
      End
      Begin VB.TextBox txtExcessOTLegal 
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
         Height          =   315
         Left            =   2160
         TabIndex        =   30
         Text            =   "0"
         Top             =   990
         Width           =   1035
      End
      Begin VB.TextBox txtOTwoPremium 
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
         Height          =   315
         Left            =   2160
         TabIndex        =   31
         Text            =   "0"
         Top             =   1260
         Width           =   1035
      End
      Begin VB.TextBox txtExcessOTSplSun 
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
         Height          =   315
         Left            =   2160
         TabIndex        =   28
         Text            =   "0"
         Top             =   450
         Width           =   1035
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Amt:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   46
         Left            =   3330
         TabIndex        =   123
         Top             =   810
         Width           =   855
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Amt:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   47
         Left            =   3330
         TabIndex        =   122
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Amt:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   48
         Left            =   3330
         TabIndex        =   121
         Top             =   540
         Width           =   855
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Amt:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   49
         Left            =   3330
         TabIndex        =   120
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Amt:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   52
         Left            =   3330
         TabIndex        =   119
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime w/o Premium:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   90
         TabIndex        =   114
         Top             =   1350
         Width           =   2085
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Excess OT on Sun/Lgl:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   44
         Left            =   90
         TabIndex        =   113
         Top             =   780
         Width           =   2085
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Excess OT on Legal Hol:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   90
         TabIndex        =   112
         Top             =   1080
         Width           =   2085
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Excess OT on Spl/Sun."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   90
         TabIndex        =   111
         Top             =   510
         Width           =   2085
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Excess OT on Reg. Day:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   51
         Left            =   90
         TabIndex        =   110
         Top             =   240
         Width           =   2085
      End
   End
   Begin VB.TextBox txtPagibigLoan 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   9270
      Locked          =   -1  'True
      TabIndex        =   98
      Text            =   "0"
      Top             =   5400
      Width           =   765
   End
   Begin VB.TextBox txtSSSLoan 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   97
      Text            =   "0"
      Top             =   5400
      Width           =   675
   End
   Begin VB.TextBox txtSalAdjustment 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   17
      Text            =   "0"
      Top             =   6390
      Width           =   1215
   End
   Begin VB.TextBox txtPhilHealthYee 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   93
      Text            =   "0"
      Top             =   5130
      Width           =   675
   End
   Begin VB.TextBox txtPHealthYer 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   9270
      Locked          =   -1  'True
      TabIndex        =   92
      Text            =   "0"
      Top             =   5130
      Width           =   765
   End
   Begin VB.TextBox txtTAX 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   91
      Text            =   "0"
      Top             =   5670
      Width           =   675
   End
   Begin VB.TextBox txtEmpyeeSSS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   85
      Text            =   "0"
      Top             =   4590
      Width           =   675
   End
   Begin VB.TextBox txtEmpyerSSS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   9270
      Locked          =   -1  'True
      TabIndex        =   84
      Text            =   "0"
      Top             =   4590
      Width           =   765
   End
   Begin VB.TextBox txtEmpyeePagibig 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   83
      Text            =   "0"
      Top             =   4860
      Width           =   675
   End
   Begin VB.TextBox txtEmpyerPagibig 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   9270
      Locked          =   -1  'True
      TabIndex        =   82
      Text            =   "0"
      Top             =   4860
      Width           =   765
   End
   Begin VB.TextBox txtECOLA 
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
      Height          =   315
      Left            =   1080
      TabIndex        =   18
      Text            =   "0"
      Top             =   6660
      Width           =   855
   End
   Begin VB.TextBox txtOverPay 
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9180
      TabIndex        =   26
      Text            =   "0"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtMiscFee 
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      TabIndex        =   23
      Text            =   "0"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtPlacementFee 
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9180
      TabIndex        =   25
      Text            =   "0"
      Top             =   3690
      Width           =   855
   End
   Begin VB.TextBox txtVale 
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      TabIndex        =   22
      Text            =   "0"
      Top             =   3690
      Width           =   855
   End
   Begin VB.TextBox txtBDOCard 
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9180
      TabIndex        =   24
      Text            =   "0"
      Top             =   3420
      Width           =   855
   End
   Begin VB.TextBox txtDeductions 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   8820
      Locked          =   -1  'True
      TabIndex        =   74
      Text            =   "0.00"
      Top             =   2790
      Width           =   1125
   End
   Begin VB.TextBox txtDeductions 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   8820
      Locked          =   -1  'True
      TabIndex        =   73
      Text            =   "0.00"
      Top             =   2420
      Width           =   1125
   End
   Begin VB.TextBox txtSplOTHours 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Text            =   "0"
      Top             =   4950
      Width           =   1215
   End
   Begin VB.TextBox txtLglHolidayOT 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   14
      Text            =   "0"
      Top             =   5310
      Width           =   1215
   End
   Begin VB.TextBox txtSickLeave 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   15
      Text            =   "0"
      Top             =   5670
      Width           =   1215
   End
   Begin VB.TextBox txtVacLeave 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   16
      Text            =   "0"
      Top             =   6030
      Width           =   1215
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   63
      Text            =   "0.00"
      Top             =   3150
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   61
      Text            =   "0.00"
      Top             =   2790
      Width           =   1125
   End
   Begin VB.TextBox txtCashBond 
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      TabIndex        =   21
      Text            =   "0"
      Top             =   3420
      Width           =   855
   End
   Begin VB.TextBox txtHoursLate 
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6900
      TabIndex        =   20
      Text            =   "0"
      Top             =   2790
      Width           =   1215
   End
   Begin VB.TextBox txtDaysAbsent 
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6900
      TabIndex        =   19
      Text            =   "0"
      Top             =   2420
      Width           =   1215
   End
   Begin VB.TextBox txtLglHoliday 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   12
      Text            =   "0"
      Top             =   4590
      Width           =   1215
   End
   Begin VB.TextBox txtSplHoliday 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      Text            =   "0"
      Top             =   4230
      Width           =   1215
   End
   Begin VB.TextBox txtSundayOT 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   10
      Text            =   "0"
      Top             =   3870
      Width           =   1215
   End
   Begin VB.TextBox txtRestDay 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   9
      Text            =   "0"
      Top             =   3510
      Width           =   1215
   End
   Begin VB.TextBox txtNightDiff 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      Text            =   "0"
      Top             =   3150
      Width           =   1215
   End
   Begin VB.TextBox txtRegOT 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Text            =   "0"
      Top             =   2790
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.CommandButton cmdTransaction 
         Caption         =   "&Delete Transaction"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5730
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdTransaction 
         Caption         =   "&Edit Transaction"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   4
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdTransaction 
         Caption         =   "&Create Transaction"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2640
         TabIndex        =   3
         Top             =   1920
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvwEmployeeList 
         Height          =   1785
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3149
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483628
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Employee Name"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee No"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView lvwProcessEmp 
         Height          =   1785
         Left            =   7650
         TabIndex        =   2
         Top             =   390
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3149
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Processed Employee"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee No"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Contract:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2610
         TabIndex        =   147
         Top             =   1650
         Width           =   1785
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Pay:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2610
         TabIndex        =   146
         Top             =   1400
         Width           =   1785
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2610
         TabIndex        =   145
         Top             =   1140
         Width           =   1785
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2610
         TabIndex        =   144
         Top             =   900
         Width           =   1785
      End
      Begin VB.Label lblContract 
         BackStyle       =   0  'Transparent
         Caption         =   "Contract Month"
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
         Height          =   285
         Left            =   4395
         TabIndex        =   143
         Top             =   1620
         Width           =   3045
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "List of Employees"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   72
         Top             =   150
         Width           =   1485
      End
      Begin VB.Label lblBasicPay 
         BackStyle       =   0  'Transparent
         Caption         =   "PhP  100.00  (Daily)"
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
         Height          =   285
         Left            =   4395
         TabIndex        =   49
         Top             =   1400
         Width           =   3045
      End
      Begin VB.Label lblEmpName 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Height          =   285
         Left            =   4410
         TabIndex        =   48
         Top             =   1140
         Width           =   3045
      End
      Begin VB.Label lblEmpNo 
         BackStyle       =   0  'Transparent
         Caption         =   "00-00-00000"
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
         Height          =   285
         Left            =   4395
         TabIndex        =   47
         Top             =   900
         Width           =   3045
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Processed Employee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7740
         TabIndex        =   46
         Top             =   150
         Width           =   1755
      End
      Begin VB.Label lblProcessedEmp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9630
         TabIndex        =   45
         Top             =   150
         Width           =   375
      End
      Begin VB.Label lblTotalEmployee 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1980
         TabIndex        =   44
         Top             =   150
         Width           =   405
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Company:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2595
         TabIndex        =   43
         Top             =   270
         Width           =   4935
      End
      Begin VB.Label lblCutoff 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Period: September 1, 2005 - September 15, 2005"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2595
         TabIndex        =   42
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lists of Employees"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5670
         TabIndex        =   41
         Top             =   2250
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   5130
      TabIndex        =   101
      Top             =   5940
      Width           =   3525
      Begin VB.TextBox txtNetPay 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   107
         Text            =   "0.00"
         Top             =   720
         Width           =   1125
      End
      Begin VB.TextBox txtDeductions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   105
         Text            =   "0.00"
         Top             =   450
         Width           =   1125
      End
      Begin VB.TextBox txtGrossPay 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   104
         Text            =   "0.00"
         Top             =   180
         Width           =   1125
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Pay:"
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
         Height          =   195
         Left            =   270
         TabIndex        =   106
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Pay:"
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
         Height          =   195
         Left            =   270
         TabIndex        =   103
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deductions:"
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
         Height          =   195
         Left            =   270
         TabIndex        =   102
         Top             =   450
         Width           =   1545
      End
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Gross:"
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
      Height          =   195
      Index           =   54
      Left            =   3240
      TabIndex        =   142
      Top             =   6750
      Width           =   585
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   3360
      TabIndex        =   141
      Top             =   3510
      Width           =   945
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   3360
      TabIndex        =   140
      Top             =   3870
      Width           =   945
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   3360
      TabIndex        =   139
      Top             =   4230
      Width           =   945
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   3360
      TabIndex        =   138
      Top             =   4590
      Width           =   945
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   24
      Left            =   3360
      TabIndex        =   137
      Top             =   6390
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   25
      Left            =   3360
      TabIndex        =   136
      Top             =   6030
      Width           =   945
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   26
      Left            =   3360
      TabIndex        =   135
      Top             =   5670
      Width           =   945
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   27
      Left            =   3360
      TabIndex        =   134
      Top             =   5310
      Width           =   945
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   28
      Left            =   3360
      TabIndex        =   133
      Top             =   4950
      Width           =   945
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Pagibig Loan:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   43
      Left            =   7650
      TabIndex        =   100
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "SSS Loan:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   39
      Left            =   5220
      TabIndex        =   99
      Top             =   5400
      Width           =   1275
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee PHealth:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   42
      Left            =   5220
      TabIndex        =   96
      Top             =   5130
      Width           =   1635
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Employer PHealth:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   41
      Left            =   7650
      TabIndex        =   95
      Top             =   5130
      Width           =   1545
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "WH TAX:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   40
      Left            =   5220
      TabIndex        =   94
      Top             =   5670
      Width           =   1455
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee SSS:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   38
      Left            =   5220
      TabIndex        =   90
      Top             =   4590
      Width           =   1275
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Contributions:"
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
      Height          =   195
      Left            =   5220
      TabIndex        =   89
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Employer SSS:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   37
      Left            =   7650
      TabIndex        =   88
      Top             =   4590
      Width           =   1455
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Pagibig:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   5220
      TabIndex        =   87
      Top             =   4860
      Width           =   1635
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Employer Pagibig:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   7650
      TabIndex        =   86
      Top             =   4860
      Width           =   1545
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "ECOLA:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   270
      TabIndex        =   81
      Top             =   6690
      Width           =   555
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Over/Under Pay:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   7740
      TabIndex        =   80
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Agency/Misc. Fee:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5220
      TabIndex        =   79
      Top             =   3960
      Width           =   1635
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Placement Fee:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   36
      Left            =   7740
      TabIndex        =   78
      Top             =   3690
      Width           =   1455
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "VALE c/o Client:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   35
      Left            =   5220
      TabIndex        =   77
      Top             =   3690
      Width           =   1635
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Card Trans. Fee:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   34
      Left            =   7740
      TabIndex        =   76
      Top             =   3420
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Deductions:"
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
      Height          =   195
      Left            =   5220
      TabIndex        =   75
      Top             =   3150
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Spl. Holiday OT Hrs:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   33
      Left            =   240
      TabIndex        =   71
      Top             =   4980
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Legal Holiday OT Hr:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   240
      TabIndex        =   70
      Top             =   5340
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Sick Leave Days:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   240
      TabIndex        =   69
      Top             =   5700
      Width           =   1815
   End
   Begin VB.Label labelLeave 
      BackStyle       =   0  'Transparent
      Caption         =   "Vacation Leave Day:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   240
      TabIndex        =   68
      Top             =   6060
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Adjustment:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   240
      TabIndex        =   67
      Top             =   6420
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   8280
      TabIndex        =   66
      Top             =   2820
      Width           =   495
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   8280
      TabIndex        =   65
      Top             =   2460
      Width           =   495
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   3360
      TabIndex        =   64
      Top             =   3150
      Width           =   945
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   3360
      TabIndex        =   62
      Top             =   2790
      Width           =   945
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Bond Deposit:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   5220
      TabIndex        =   60
      Top             =   3420
      Width           =   1635
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Hours Late:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   5220
      TabIndex        =   59
      Top             =   2790
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Days Absent:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5220
      TabIndex        =   58
      Top             =   2430
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   3360
      TabIndex        =   57
      Top             =   2460
      Width           =   945
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   4
      X1              =   10170
      X2              =   10170
      Y1              =   2340
      Y2              =   7020
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   3
      X1              =   5040
      X2              =   5040
      Y1              =   2400
      Y2              =   7020
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   2340
      Y2              =   7020
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Legal Holiday:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   56
      Top             =   4620
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Spl. Holiday:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   55
      Top             =   4260
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Sunday OT Hours:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   54
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Rest Day Hours:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   53
      Top             =   3540
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Night Diff. Hours:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   52
      Top             =   3180
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Regular OT Hours:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   51
      Top             =   2820
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   1
      X1              =   180
      X2              =   10125
      Y1              =   7020
      Y2              =   7035
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Days Worked:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   50
      Top             =   2490
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   0
      X1              =   180
      X2              =   10125
      Y1              =   2340
      Y2              =   2355
   End
   Begin MSForms.Image Image1 
      Height          =   4695
      Left            =   90
      Top             =   2340
      Width           =   10095
      BackColor       =   12632256
      Size            =   "17806;8281"
   End
End
Attribute VB_Name = "frmPayrollEntry3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCompanyInfo As ADODB.Recordset
Dim rsClientInfo As ADODB.Recordset     ' client table
Dim rsEmployeeInfo As ADODB.Recordset
Dim rsProcessEmployee As ADODB.Recordset

Dim rsPayrollMast As ADODB.Recordset    ' master payroll table
Dim rsPayrollTrans As ADODB.Recordset   ' payroll transaction

Dim rsSSSTable As ADODB.Recordset
Dim rsTaxTable As ADODB.Recordset
Dim rsPHealthTable As ADODB.Recordset
Dim rsPagibigTable As ADODB.Recordset

Dim PayrollValidation As clsCommon
Dim strQuery As String

'Get Agency Information
Dim mAgencyCode As Integer

'Get Employee Information
Dim ProcessFlag As Boolean
Dim EmployeeName As String
Dim EmployeeNo As String
Dim ECOLARate As Double
Dim DailyRate As Double
Dim MonthlyRate As Double
Dim HourlyRate As Double
Dim PaymentBasis As String
Dim AddorEditFlag As Boolean
Dim MonthlyFlag As Boolean
Dim EmpTaxCode As String

'Variable declarations for deductions.
Dim BDOCardFee As Double
Dim CashBondFee As Double
Dim AgencyFee As Double

'Get Client Information
Dim ClientCutoff As String
Dim WSSSFlag As Boolean
Dim WPHealthFlag As Boolean
Dim WPagibigFlag As Boolean
Dim WTaxFlag As Boolean
Dim EditTaxFlag As Boolean
Dim NoofCutoffs As Integer
Dim mCutOffNo As Integer

Private Sub cmdCreateTrans_Click(Index As Integer)
statbar.Panels(1).Text = "Create Transaction"
End Sub

Private Sub cmdCancelTrans_Click(Index As Integer)
AddorEditFlag = False
Call DisableCmdButtons
End Sub

Private Sub cmdClose_Click()
Set rsCompanyInfo = Nothing
Set rsClientInfo = Nothing
Set rsEmployeeInfo = Nothing
Set rsProcessEmployee = Nothing
Unload Me
End Sub

Private Sub cmdEditTax_Click(Index As Integer)
If EditTaxFlag = True Then
    ' Enable tax text box for editing
    txtTAX.Enabled = True
    txtTAX.Locked = False
    txtTAX.SetFocus
Else
    txtTAX.Enabled = False
    txtTAX.Locked = True
End If
End Sub

Private Sub cmdExcessOt_Click(Index As Integer)

If cmdExcessOt(4).Caption = "Excess O.T." Then
    cmdExcessOt(4).Caption = "Hide Excess OT"
    frameExcessOT.Visible = True
Else
    cmdExcessOt(4).Caption = "Excess O.T."
    frameExcessOT.Visible = False
End If

End Sub

Private Sub cmdSaveTrans_Click(Index As Integer)
Call SaveTransaction
AddorEditFlag = False
Call DisableCmdButtons
End Sub

Private Sub cmdTransaction_Click(Index As Integer)
AddorEditFlag = True
Call ZeroOutAmount
Call DisableCmdButtons
End Sub

Private Sub Command1_Click(Index As Integer)
Dim strword As String
Dim amountinwords As Double
amountinwords = 343432.22
strword = ConvertAmountinWords(amountinwords)
End Sub

Private Function ConvertAmountinWords(ByVal mAmt As Double) As String
MsgBox oCommonFunction.Word(mAmt)
End Function

Private Sub Form_Load()

mAgencyCode = rsCompanyTable!agencycode
Call DisableTextBoxes(False)
Me.Caption = "Payroll Transaction Entries for " & Replace(mCompanyName, "&", "&&")
lblCompany.Caption = "Company: " & Replace(mCompanyName, "&", "&&")
lblCutoff.Caption = "Payroll Period: " & Format(frmSelectCutoff.txtCutOffDate(0).Text, "mmm dd, yyyy") & " to " & Format(frmSelectCutoff.txtCutOffDate(1).Text, "mmm dd, yyyy")
frameExcessOT.Visible = False
Call GetClientInformation
Call GetEmployeeBasic
' Open Payroll reference table.
Call OpenReferenceTable
If rsEmployeeInfo.RecordCount <> 0 Then
    Call GetBasicInformation
End If
End Sub

Private Sub Form_Initialize()
    Set PayrollValidation = New clsCommon
End Sub

Private Sub Form_Terminate()
    Set PayrollValidation = Nothing
End Sub

Private Sub GetEmployeeBasic()
'Get all the list of employees for processing
strQuery = "SELECT a.*, (trim(b.surname) + ', ' + trim(b.firstname) + ' ' + iif(b.middlename = '', '' , left(b.middlename, 1) + '.') ) AS fullname, " & _
            "b.sss, b.philhealth, b.pagibig, b.acctno, b.taxid, b.employmentstatus FROM employeebasicsalary AS a LEFT JOIN hremployees AS b ON a.employeenum = b.employeenum " & _
            "WHERE a.clientcode = " & mClientCode & " Order by b.surname, b.firstname"

Set rsEmployeeInfo = New ADODB.Recordset
Set rsEmployeeInfo = ReturnRst(strQuery, Cnnstr, mstatic)
If rsEmployeeInfo.RecordCount = 0 Then
    Call DisableAllButtons
End If
Call FillEmployeeList(lvwEmployeeList, lvwProcessEmp, rsEmployeeInfo)
End Sub

Private Sub GetClientInformation()
Dim rsCutOff As ADODB.Recordset

strQuery = "SELECT * from cutofftable WHERE clientcode = " & mClientCode & " ORDER BY recno"
Set rsCutOff = ReturnRst(strQuery, Cnnstr, mstatic)
rsCutOff.MoveLast
mCutOffNo = rsCutOff!cutoffno

WSSSFlag = False
WPHealthFlag = False
WPagibigFlag = False
WTaxFlag = False
With rsClientTable
    If .RecordCount = 0 Then
        Exit Sub
    End If
    If .EOF = True Then
        .MoveLast
    ElseIf .BOF = True Then .MoveFirst
    End If
    WSSSFlag = !clientsssflag
    WPHealthFlag = !ClientPhilHealthFlag
    WPagibigFlag = !ClientPagibigFlag
    WTaxFlag = !ClientTaxFlag
    EditTaxFlag = !ClientEditTaxFlag
    ClientCutoff = !ClientCutoff
    If ClientCutoff = "Weekly" Then
        NoofCutoffs = 4
    ElseIf ClientCutoff = "Semi-Monthly" Then
        NoofCutoffs = 2
    Else        ' for monthly cut-offs
        NoofCutoffs = 1
    End If
End With

End Sub

Private Function FillEmployeeList(ByRef mLsEmp As ListView, _
                        ByRef mLsProc As ListView, _
                        mRs As Recordset)
Dim litem1 As ListItem
Dim litem2 As ListItem
Dim mCtr1 As Integer
Dim mCtr2 As Integer

    mLsEmp.View = lvwReport
    mLsProc.View = lvwReport
    mCtr1 = mCtr2 = 0
    lblProcessedEmp.Caption = 0
    lblTotalEmployee.Caption = 0
    mLsEmp.ListItems.Clear
    mLsProc.ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
        If ProcessFlag = True Then ' add to process list view
            Set litem1 = mLsProc.ListItems.add(, , IIf(IsNull(mRs(47)), "", mRs(47)))
            litem1.SubItems(1) = mRs(0)
            mCtr1 = mCtr1 + 1
            lblProcessedEmp.Caption = mCtr1
        Else                        ' add to employee list view
            Set litem2 = mLsEmp.ListItems.add(, , IIf(IsNull(mRs(47)), "", mRs(47)))
            litem2.SubItems(1) = mRs(0)
            mCtr2 = mCtr2 + 1
            lblTotalEmployee.Caption = mCtr2
        End If
        mRs.MoveNext
    Wend
End Function

Private Sub OpenReferenceTable()
Dim strQuery As String

strQuery = "SELECT minwage, maxwage, er_share, ec_share, ee_share FROM SSStable order by id"
Set rsSSSTable = ReturnRst(strQuery, Cnnstr, mstatic)
strQuery = "SELECT minwage, maxwage, ph_er, ph_ee FROM PhilHealthtable order by id"
Set rsPHealthTable = ReturnRst(strQuery, Cnnstr, mstatic)
strQuery = "SELECT minwage, maxwage, pagibigee, pagibiger, pagibigmax FROM Pagibigtable order by id"
Set rsPagibigTable = ReturnRst(strQuery, Cnnstr, mstatic)

'Open tax table depending on cut-off.
Select Case ClientCutoff
    Case "Daily" ' daily cut-off used the daily table
        strQuery = "SELECT taxcode, salrate2, basetax, perover FROM DailyTaxTable order by id"
        Set rsTaxTable = ReturnRst(strQuery, Cnnstr, mstatic)
    Case "Weekly" ' weekly cut-off used the weekly table
        strQuery = "SELECT taxcode, salrate2, basetax, perover FROM WeeklyTaxTable order by id"
        Set rsTaxTable = ReturnRst(strQuery, Cnnstr, mstatic)
    Case "Semi-Monthly" ' semi-monthly cut-off used the semimonthly table
        strQuery = "SELECT taxcode, salrate2, basetax, perover FROM SemiMonthlyTaxTable order by id"
        Set rsTaxTable = ReturnRst(strQuery, Cnnstr, mstatic)
    Case "Monthly" ' Monthly cut-off used the monthly table
        strQuery = "SELECT taxcode, salrate2, basetax, perover FROM MonthlyTaxTable order by id"
        Set rsTaxTable = ReturnRst(strQuery, Cnnstr, mstatic)
    Case Else
        strQuery = "SELECT taxcode, salrate2, basetax, perover FROM SemiMonthlyTaxTable order by id"
        Set rsTaxTable = ReturnRst(strQuery, Cnnstr, mstatic)
End Select

End Sub

Private Sub SSSPayment(ByVal mBasicPay As Double)

txtEmpyeeSSS.Text = "0"
txtEmpyerSSS.Text = "0"
With rsSSSTable
    .MoveFirst
    Do While Not .EOF
        If mBasicPay >= !minwage And mBasicPay <= !maxwage Then
            txtEmpyeeSSS.Text = Format(!ee_share, "0.00")
            txtEmpyerSSS.Text = Format(!er_share + !ec_share, "0.00")
            Exit Sub
        End If
        .MoveNext
    Loop
End With
End Sub

Private Sub PHealthPayment(ByVal mBasicPay As Double)

txtPhilHealthYee.Text = "0"
txtPHealthYer.Text = "0"
With rsPHealthTable
    .MoveFirst
    Do While Not .EOF
        If mBasicPay >= !minwage And mBasicPay <= !maxwage Then
            txtPhilHealthYee.Text = Format(!ph_ee, "0.00")
            txtPHealthYer.Text = Format(!ph_er, "0.00")
        Exit Sub
        End If
        .MoveNext
    Loop
End With
End Sub

Private Sub PagibigPayment(ByVal mBasicPay As Double)

txtEmpyeePagibig.Text = "0"
txtEmpyerPagibig.Text = "0"
With rsPagibigTable
    .MoveFirst
    Do While Not .EOF
        If mBasicPay >= !minwage And mBasicPay <= !maxwage Then
            txtEmpyeePagibig.Text = Format((IIf(mBasicPay > 5000, !pagibigmax, !pagibigee) / NoofCutoffs), "0.00")
            txtEmpyerPagibig.Text = Format((IIf(mBasicPay > 5000, !pagibigmax, !pagibiger) / NoofCutoffs), "0.00")
        Exit Sub
        End If
        .MoveNext
    Loop
End With
End Sub

Private Sub GetContributions(ByVal mBasicPay As Double)

    If WSSSFlag = True Then
        Call SSSPayment(mBasicPay)
    End If
    If WPagibigFlag = True Then
        Call PagibigPayment(MonthlyRate)
    End If
    If WPHealthFlag = True Then
        Call PHealthPayment(mBasicPay)
    End If
    
End Sub

Private Sub ComputeTax(ByVal mGrossPay As Double)
Dim mBaseTax As Double
Dim mTaxRate As Double
Dim mSalRate As Double

With rsTaxTable
    '.Filter = "taxcode = '" & Trim(EmpTaxCode) & "'"
    .Filter = "taxcode" & " = " & "'" & Trim(EmpTaxCode) & "'"
    If .EOF Then
        MsgBox "Invalid tax exemption code! Not found in tax table!", vbInformation, PayrollAlert
        Exit Sub
    End If
    .MoveFirst
    mBaseTax = !BaseTax
    mTaxRate = !PerOver
    mSalRate = !SalRate2
    Do While Not .EOF 'And !taxcode = Trim(EmpTaxCode)
        If mGrossPay < !SalRate2 Then
            Exit Do
        Else
            mBaseTax = !BaseTax
            mTaxRate = !PerOver
            mSalRate = !SalRate2
        End If
        .MoveNext
    Loop
End With
If mGrossPay > 1 Then
    txtTAX.Text = Format(((mGrossPay - mSalRate) * mTaxRate) + mBaseTax, "0.00")
Else
    txtTAX.Text = "0.00"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmSelectCutoff.Show
End Sub


Private Sub lvwEmployeeList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ' Get the employee number
    EmployeeName = Trim(Item.Text)
    Call GetBasicInformation
End Sub

Private Sub lvwProcessEmp_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ' Get the employee number
    EmployeeName = Trim(Item.Text)
    Call GetBasicInformation
End Sub

Private Sub GetBasicInformation()   ' per employee
Dim PayBasis As String
Dim ContractMonth As String

    lblEmpName.Caption = Trim(EmployeeName)
    rsEmployeeInfo.MoveFirst
    rsEmployeeInfo.Find "Fullname = '" & EmployeeName & "'"
    If Not rsEmployeeInfo.EOF Then
        EmployeeNo = rsEmployeeInfo!employeenum
        PayBasis = IIf(rsEmployeeInfo!PaymentBasis = "Daily", "PhP " & Format(rsEmployeeInfo!DailyRate, "###,##0.00") & " (Daily)", "PhP " & Format(rsEmployeeInfo!MonthlyRate, "###,##0.00") & " (Monthly)")
        DailyRate = IIf(IsNull(rsEmployeeInfo!DailyRate), 0, rsEmployeeInfo!DailyRate)
        PaymentBasis = IIf(IsNull(rsEmployeeInfo!PaymentBasis) Or rsEmployeeInfo!PaymentBasis = "", "Daily", rsEmployeeInfo!PaymentBasis)
        MonthlyRate = IIf(IsNull(rsEmployeeInfo!MonthlyRate), 0, rsEmployeeInfo!MonthlyRate)
        HourlyRate = IIf(IsNull(rsEmployeeInfo!HourlyRate), 0, rsEmployeeInfo!HourlyRate)
        ContractMonth = IIf(Trim(rsEmployeeInfo!employmentstatus) <> "R", IIf(IsNull(rsEmployeeInfo!StartContract), "", Format(rsEmployeeInfo!StartContract, "mm/dd/yyyy")) & " to " & IIf(IsNull(rsEmployeeInfo!EndContract), "", Format(rsEmployeeInfo!EndContract, "mm/dd/yyyy")), "Regular Employee")
        EmpTaxCode = IIf(IsNull(rsEmployeeInfo!EmpTaxCode), "Z", rsEmployeeInfo!EmpTaxCode)
        'Get ECOLA Rate
        ECOLARate = IIf(rsEmployeeInfo!WithEcolaPay = True, rsEmployeeInfo!EcolaAmount, 0#)
        'Get Miscellaneous Fees
        BDOCardFee = Format(IIf(IsNull(rsEmployeeInfo!AgencyBDOCardFee), 0#, rsEmployeeInfo!AgencyBDOCardFee), "0.00")
        CashBondFee = Format(IIf(IsNull(rsEmployeeInfo!AgencyCashBond), 0#, rsEmployeeInfo!AgencyCashBond), "0.00")
        AgencyFee = Format(IIf(IsNull(rsEmployeeInfo!AgencyFee), 0#, rsEmployeeInfo!AgencyFee), "0.00")
    Else
        EmployeeNo = ""
        PayBasis = ""
    End If
    lblEmpNo.Caption = EmployeeNo
    lblBasicPay.Caption = PayBasis
    lblContract.Caption = IIf(Trim(ContractMonth) = "to", "Not Specified", ContractMonth)

End Sub

Private Sub DisableTextBoxes(ByVal TFalse As Boolean)
    txtDaysWork.Enabled = TFalse
    txtRegOT.Enabled = TFalse
    txtNightDiff.Enabled = TFalse
    txtRestDay.Enabled = TFalse
    txtSundayOT.Enabled = TFalse
    txtSplHoliday.Enabled = TFalse
    txtLglHoliday.Enabled = TFalse
    txtSplOTHours.Enabled = TFalse
    txtLglHolidayOT.Enabled = TFalse
    txtSickLeave.Enabled = TFalse
    txtVacLeave.Enabled = TFalse
    txtSalAdjustment.Enabled = TFalse
    txtECOLA.Enabled = TFalse
    txtDaysAbsent.Enabled = TFalse
    txtHoursLate.Enabled = TFalse
    txtCashBond.Enabled = TFalse
    txtVale.Enabled = TFalse
    txtMiscFee.Enabled = TFalse
    txtBDOCard.Enabled = TFalse
    txtPlacementFee.Enabled = TFalse
    txtOverPay.Enabled = TFalse
    txtExcessOTRegDay.Enabled = TFalse
    txtExcessOTSplSun.Enabled = TFalse
    txtExcessOTSunLgl.Enabled = TFalse
    txtExcessOTLegal.Enabled = TFalse
    txtOTwoPremium.Enabled = TFalse
End Sub

Private Sub ZeroOutAmount()
Dim i As Integer

    txtDaysWork.Text = ""
    txtRegOT.Text = ""
    txtNightDiff.Text = ""
    txtRestDay.Text = ""
    txtSundayOT.Text = ""
    txtSplHoliday.Text = ""
    txtLglHoliday.Text = ""
    txtSplOTHours.Text = ""
    txtLglHolidayOT.Text = ""
    txtSickLeave.Text = ""
    txtVacLeave.Text = ""
    txtSalAdjustment.Text = ""
    txtDaysAbsent.Text = ""
    txtHoursLate.Text = ""
    txtCashBond.Text = Format(CashBondFee, "0.00")
    txtVale.Text = ""
    txtMiscFee.Text = Format(AgencyFee, "0.00")
    txtBDOCard.Text = Format(BDOCardFee, "0.00")
    txtPlacementFee.Text = ""
    txtOverPay.Text = ""
    txtExcessOTRegDay.Text = ""
    txtExcessOTSplSun.Text = ""
    txtExcessOTSunLgl.Text = ""
    txtExcessOTLegal.Text = ""
    txtOTwoPremium.Text = ""
    txtECOLA.Text = Format(ECOLARate, "0.00")
    
    txtEmpyeeSSS.Text = "0"
    txtEmpyerSSS.Text = "0"
    txtEmpyeePagibig.Text = "0"
    txtEmpyerPagibig.Text = "0"
    txtPhilHealthYee.Text = "0"
    txtPHealthYer.Text = "0"
    txtSSSLoan.Text = "0"
    txtPagibigLoan.Text = "0"
    txtTAX.Text = "0"
    
    txtGrossPay.Text = "0.00"
    txtNetPay.Text = "0.00"
    txtDeductions(2).Text = "0.00"
    
    For i = 0 To 16
        txtAmount(i).Text = "0.00"
        txtAmount(i).BackColor = &HC0C0C0
    Next
    For i = 0 To 1
        txtDeductions(i).Text = "0.00"
        txtDeductions(i).BackColor = &HC0C0C0
    Next
End Sub

Private Sub DisableAllButtons()
    lvwEmployeeList.Enabled = False
    lvwProcessEmp.Enabled = False
    cmdSaveTrans(4).Visible = False
    cmdCancelTrans(5).Visible = False
    cmdTransaction(0).Enabled = False
    cmdTransaction(1).Enabled = False
    cmdTransaction(2).Enabled = False
    cmdFind.Enabled = False
    cmdfirst.Enabled = False
    cmdprev.Enabled = False
    cmdnext.Enabled = False
    cmdlast.Enabled = False
    cmdPrintReports.Enabled = False
    cmdmovetoMaster.Enabled = False
    cmdClose.Enabled = True
End Sub

Private Sub DisableCmdButtons()
    If AddorEditFlag = True Then
        Call DisableTextBoxes(True)
        lvwEmployeeList.Enabled = False
        lvwProcessEmp.Enabled = False
        cmdSaveTrans(4).Visible = True
        cmdCancelTrans(5).Visible = True
        cmdTransaction(0).Enabled = False
        cmdTransaction(1).Enabled = False
        cmdTransaction(2).Enabled = False
        cmdFind.Enabled = False
        cmdfirst.Enabled = False
        cmdprev.Enabled = False
        cmdnext.Enabled = False
        cmdlast.Enabled = False
        cmdPrintReports.Enabled = False
        cmdmovetoMaster.Enabled = False
        cmdClose.Enabled = False
    Else
        Call DisableTextBoxes(False)
        lvwEmployeeList.Enabled = True
        lvwProcessEmp.Enabled = True
        cmdSaveTrans(4).Visible = False
        cmdCancelTrans(5).Visible = False
        cmdTransaction(0).Enabled = True
        cmdTransaction(1).Enabled = True
        cmdTransaction(2).Enabled = True
        cmdFind.Enabled = True
        cmdfirst.Enabled = True
        cmdprev.Enabled = True
        cmdnext.Enabled = True
        cmdlast.Enabled = True
        cmdPrintReports.Enabled = True
        cmdmovetoMaster.Enabled = True
        cmdClose.Enabled = True
    End If
End Sub

Private Sub txtAmount_Change(Index As Integer)
Dim i As Integer
Dim TotValue As Double

    For i = 0 To 16
        TotValue = TotValue + IIf(Trim(txtAmount(i)) = "." Or txtAmount(i).Text = "", 0, txtAmount(i).Text)
        txtAmount(i).Refresh
    Next
    txtGrossPay1.Text = Format(TotValue, "###,##0.00")
    txtGrossPay.Text = Format(TotValue, "###,##0.00")
    txtAmount(1).ForeColor = &HFF0000
    txtNetPay = Format(txtGrossPay.Text - txtDeductions(2).Text, "###,##0.00")
    'Compute for tax.
    If WTaxFlag = True Then
        Call ComputeTax(txtGrossPay.Text)
    End If

End Sub

Private Sub txtBDOCard_GotFocus()
    SendKeys "{HOME}+{END}"
    txtBDOCard.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtBDOCard
End Sub

Private Sub txtBDOCard_LostFocus()
    txtBDOCard.BackColor = &HFFFFFF
    Call txtGrossPay_Change
End Sub

Private Sub txtCashBond_GotFocus()
    SendKeys "{HOME}+{END}"
    txtCashBond.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtCashBond
End Sub

Private Sub txtCashBond_LostFocus()
    txtCashBond.BackColor = &HFFFFFF
    Call txtGrossPay_Change
End Sub

Private Sub txtDaysAbsent_GotFocus()
    SendKeys "{HOME}+{END}"
    txtDaysAbsent.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtDaysAbsent
End Sub

Private Sub txtDaysAbsent_LostFocus()
    txtDeductions(0).BackColor = &HFFFFFF
    txtDeductions(0).Text = Format(ComputeAmount(txtDaysAbsent.Text, PaymentBasis, False, 300), "###,##0.00")
    txtDaysAbsent.BackColor = &HFFFFFF
    Call txtGrossPay_Change
End Sub

Private Sub txtDaysWork_GotFocus()
    SendKeys "{HOME}+{END}"
    txtDaysWork.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtDaysWork
End Sub

Private Sub txtDaysWork_LostFocus()
    MonthlyFlag = True
    txtAmount(0).BackColor = &HFFFFFF
    txtAmount(0).Text = Format(ComputeAmount(txtDaysWork.Text, PaymentBasis, False, 0), "###,##0.00")
    txtDaysWork.BackColor = &HFFFFFF
    MonthlyFlag = False
    Call txtECOLA_LostFocus
    Call GetContributions(txtAmount(0).Text)
End Sub

Private Function ComputeAmount(ByRef mAmt As String, mPay As String, mHrSw As Boolean, ByVal IndexNo As Integer) As Double
Dim TotValue As Double
Dim i As Integer
Dim DeciSw As Boolean
Dim strAmt As String

    DeciSw = False
    If IsNull(mAmt) Or mAmt = "" Or mAmt = "0" Then
    'If mAmt = 0# Then
        mAmt = "0.00"
        ComputeAmount = 0#
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
        mAmt = IIf(strAmt = ".", 0#, Format(strAmt, "0.00"))
        If mAmt > 99999.99 Then
            MsgBox "Entries should not exceed 99,999.99. Please correct it!", vbInformation, PayrollAlert
            ComputeAmount = 0#
            Exit Function
        End If
        If mPay = "Daily" Then
            If mHrSw = False Then
                ComputeAmount = mAmt * DailyRate
            Else
                ComputeAmount = mAmt * HourlyRate
            End If
        Else
            If mPay = "Monthly" Then
                If MonthlyFlag = False Then ' do not used the monthly salary as base
                    If mHrSw = False Then
                        ComputeAmount = mAmt * DailyRate
                    Else
                        ComputeAmount = mAmt * HourlyRate
                    End If
                Else        ' get the monthly salary base
                    If mHrSw = False Then
                        ComputeAmount = MonthlyRate / NoofCutoffs
                    Else
                        ComputeAmount = mAmt * HourlyRate
                    End If
                End If
            End If
        End If
    End If
'    If IndexNo > 299 Then
'        For i = 0 To 2
'            TotValue = TotValue + txtDeductions(i).Text
'        Next
'        txtDeductions(2).Text = Format(TotValue, "###,##0.00")
'    Else
'        For i = 0 To 16
'            TotValue = TotValue + txtAmount(i).Text
'        Next
'        txtGrossPay1.Text = Format(TotValue, "###,##0.00")
'        txtGrossPay.Text = Format(TotValue, "###,##0.00")
'    End If
    Call txtGrossPay_Change
    Call txtAmount_Change(0)
End Function

Private Sub txtECOLA_GotFocus()
    SendKeys "{HOME}+{END}"
    txtECOLA.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtECOLA
End Sub

Private Sub txtECOLA_LostFocus()
Dim i As Integer
Dim strAmt As String
Dim mAmt As String
Dim DeciSw As Boolean

    txtECOLA.BackColor = &HFFFFFF
    mAmt = Trim(txtECOLA.Text)
    For i = 1 To Len(mAmt)
        If DeciSw = True And Mid(mAmt, i, 1) = "." Then
        Else
            strAmt = strAmt & Mid(mAmt, i, 1)
        End If
        If DeciSw = False And Mid(mAmt, i, 1) = "." Then
            DeciSw = True
        End If
    Next
    mAmt = IIf(strAmt = ".", 0#, Format(strAmt, "0.00"))
    'Add Ecola payment to total netpay.
    txtNetPay.Text = Format(txtNetPay.Text + Val(mAmt), "###,##0.00")  ' add the ECOLA to adjustment field
End Sub

Private Sub txtExcessOTLegal_GotFocus()
    SendKeys "{HOME}+{END}"
    txtExcessOTLegal.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtExcessOTLegal
End Sub

Private Sub txtExcessOTLegal_LostFocus()
    txtAmount(15).BackColor = &HFFFFFF
    txtAmount(15).Text = Format(ComputeAmount(txtExcessOTLegal.Text, PaymentBasis, True, 15), "###,##0.00")
    txtExcessOTLegal.BackColor = &HFFFFFF
End Sub

Private Sub txtExcessOTRegDay_GotFocus()
    SendKeys "{HOME}+{END}"
    txtExcessOTRegDay.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtExcessOTRegDay
End Sub

Private Sub txtExcessOTRegDay_LostFocus()
    txtAmount(12).BackColor = &HFFFFFF
    txtAmount(12).Text = Format(ComputeAmount(txtExcessOTRegDay.Text, PaymentBasis, True, 12), "###,##0.00")
    txtExcessOTRegDay.BackColor = &HFFFFFF
End Sub

Private Sub txtExcessOTSplSun_GotFocus()
    SendKeys "{HOME}+{END}"
    txtExcessOTSplSun.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtExcessOTSplSun
End Sub

Private Sub txtExcessOTSplSun_LostFocus()
    txtAmount(13).BackColor = &HFFFFFF
    txtAmount(13).Text = Format(ComputeAmount(txtExcessOTSplSun.Text, PaymentBasis, True, 13), "###,##0.00")
    txtExcessOTSplSun.BackColor = &HFFFFFF
End Sub

Private Sub txtExcessOTSunLgl_GotFocus()
    SendKeys "{HOME}+{END}"
    txtExcessOTSunLgl.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtExcessOTSunLgl
End Sub

Private Sub txtExcessOTSunLgl_LostFocus()
    txtAmount(14).BackColor = &HFFFFFF
    txtAmount(14).Text = Format(ComputeAmount(txtExcessOTSunLgl.Text, PaymentBasis, True, 14), "###,##0.00")
    txtExcessOTSunLgl.BackColor = &HFFFFFF
End Sub

Private Sub txtGrossPay_Change()
Dim i As Integer
Dim TotValue As Double

    For i = 0 To 1
        TotValue = TotValue + IIf(txtDeductions(i).Text = "", 0, txtDeductions(i).Text)
        txtDeductions(i).Refresh
    Next
    TotValue = TotValue + IIf(txtBDOCard.Text = "", 0, txtBDOCard.Text) + IIf(txtCashBond.Text = "", 0, txtCashBond.Text) + _
        IIf(txtVale.Text = "", 0, txtVale.Text) + IIf(txtPlacementFee.Text = "", 0, txtPlacementFee.Text) + IIf(txtMiscFee.Text = "", 0, txtMiscFee.Text) + IIf(txtOverPay.Text = "", 0, txtOverPay.Text) + _
        IIf(txtEmpyeeSSS.Text = "", 0, txtEmpyeeSSS.Text) + IIf(txtEmpyeePagibig.Text = "", 0, txtEmpyeePagibig.Text) + IIf(txtPhilHealthYee.Text = "", 0, txtPhilHealthYee.Text) + _
        IIf(txtPagibigLoan.Text = "", 0, txtPagibigLoan.Text) + IIf(txtSSSLoan.Text = "", 0, txtSSSLoan.Text) + IIf(txtTAX.Text = "", 0, txtTAX.Text)
    txtDeductions(2).Text = Format(TotValue, "###,##0.00")
    txtNetPay = Format(txtGrossPay.Text - TotValue, "###,##0.00")
End Sub

Private Sub txtHoursLate_GotFocus()
    SendKeys "{HOME}+{END}"
    txtHoursLate.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtHoursLate
End Sub

Private Sub txtHoursLate_LostFocus()
    txtDeductions(1).BackColor = &HFFFFFF
    txtDeductions(1).Text = Format(ComputeAmount(txtHoursLate.Text, PaymentBasis, True, 301), "###,##0.00")
    txtHoursLate.BackColor = &HFFFFFF
    Call txtGrossPay_Change
End Sub

Private Sub txtLglHoliday_GotFocus()
    SendKeys "{HOME}+{END}"
    txtLglHoliday.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtLglHoliday
End Sub

Private Sub txtLglHoliday_LostFocus()
    txtAmount(6).BackColor = &HFFFFFF
    txtAmount(6).Text = Format(ComputeAmount(txtLglHoliday.Text, PaymentBasis, False, 6), "###,##0.00")
    txtLglHoliday.BackColor = &HFFFFFF
End Sub

Private Sub txtLglHolidayOT_GotFocus()
    SendKeys "{HOME}+{END}"
    txtLglHolidayOT.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtLglHolidayOT
End Sub

Private Sub txtLglHolidayOT_LostFocus()
    txtAmount(8).BackColor = &HFFFFFF
    txtAmount(8).Text = Format(ComputeAmount(txtLglHolidayOT.Text, PaymentBasis, True, 8), "###,##0.00")
    txtLglHolidayOT.BackColor = &HFFFFFF
End Sub

Private Sub txtMiscFee_GotFocus()
    SendKeys "{HOME}+{END}"
    txtMiscFee.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtMiscFee
End Sub

Private Sub txtMiscFee_LostFocus()
    txtMiscFee.BackColor = &HFFFFFF
    Call txtGrossPay_Change
End Sub

Private Sub txtNightDiff_GotFocus()
    SendKeys "{HOME}+{END}"
    Set oCommonFunction.ValidDouble = txtNightDiff
    txtNightDiff.BackColor = &HFFFFC0
End Sub

Private Sub txtNightDiff_LostFocus()
    txtAmount(2).BackColor = &HFFFFFF
    txtAmount(2).Text = Format(ComputeAmount(txtNightDiff.Text, PaymentBasis, True, 2), "###,##0.00")
    txtNightDiff.BackColor = &HFFFFFF
End Sub

Private Sub txtOTwoPremium_GotFocus()
    SendKeys "{HOME}+{END}"
    txtOTwoPremium.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtOTwoPremium
End Sub

Private Sub txtOTwoPremium_LostFocus()
    txtAmount(16).BackColor = &HFFFFFF
    txtAmount(16).Text = Format(ComputeAmount(txtOTwoPremium.Text, PaymentBasis, False, 16), "###,##0.00")
    txtOTwoPremium.BackColor = &HFFFFFF
End Sub

Private Sub txtOverPay_GotFocus()
    SendKeys "{HOME}+{END}"
    txtOverPay.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtOverPay
End Sub

Private Sub txtOverPay_LostFocus()
    txtOverPay.BackColor = &HFFFFFF
    Call txtGrossPay_Change
End Sub

Private Sub txtPlacementFee_GotFocus()
    SendKeys "{HOME}+{END}"
    txtPlacementFee.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtPlacementFee
End Sub

Private Sub txtPlacementFee_LostFocus()
    txtPlacementFee.BackColor = &HFFFFFF
    Call txtGrossPay_Change
End Sub

Private Sub txtRegOT_GotFocus()
    SendKeys "{HOME}+{END}"
    txtRegOT.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtRegOT
End Sub

Private Sub txtRegOT_LostFocus()
    txtAmount(1).BackColor = &HFFFFFF
    txtAmount(1).Text = Format(ComputeAmount(txtRegOT.Text, PaymentBasis, True, 1), "###,##0.00")
    txtRegOT.BackColor = &HFFFFFF
End Sub

Private Sub txtRestDay_GotFocus()
    SendKeys "{HOME}+{END}"
    txtRestDay.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtRestDay
End Sub

Private Sub txtRestDay_LostFocus()
    txtAmount(3).BackColor = &HFFFFFF
    txtAmount(3).Text = Format(ComputeAmount(txtRestDay.Text, PaymentBasis, True, 3), "###,##0.00")
    txtRestDay.BackColor = &HFFFFFF
End Sub

Private Sub txtSalAdjustment_GotFocus()
    SendKeys "{HOME}+{END}"
    txtSalAdjustment.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtSalAdjustment
End Sub

Private Sub txtSalAdjustment_LostFocus()
Dim i As Integer
Dim strAmt As String
Dim mAmt As String
Dim DeciSw As Boolean

    txtAmount(11).BackColor = &HFFFFFF
    mAmt = Trim(txtSalAdjustment)
    For i = 1 To Len(mAmt)
        If DeciSw = True And Mid(mAmt, i, 1) = "." Then
        Else
            strAmt = strAmt & Mid(mAmt, i, 1)
        End If
        If DeciSw = False And Mid(mAmt, i, 1) = "." Then
            DeciSw = True
        End If
    Next
    mAmt = IIf(strAmt = ".", 0#, Format(strAmt, "0.00"))
    
    txtAmount(11).Text = Format(mAmt, "0.00") 'Format(ComputeAmount(Format(IIf(txtSalAdjustment.Text = "", 0, txtSalAdjustment.Text), "0.00"), PaymentBasis, False, 11), "###,##0.00")
    txtSalAdjustment.BackColor = &HFFFFFF
End Sub

Private Sub txtSickLeave_GotFocus()
    SendKeys "{HOME}+{END}"
    txtSickLeave.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtSickLeave
End Sub

Private Sub txtSickLeave_LostFocus()
    txtAmount(9).BackColor = &HFFFFFF
    txtAmount(9).Text = Format(ComputeAmount(txtSickLeave.Text, PaymentBasis, False, 9), "###,##0.00")
    txtSickLeave.BackColor = &HFFFFFF
End Sub

Private Sub txtSplHoliday_GotFocus()
    SendKeys "{HOME}+{END}"
    txtSplHoliday.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtSplHoliday
End Sub

Private Sub txtSplHoliday_LostFocus()
    txtAmount(5).BackColor = &HFFFFFF
    txtAmount(5).Text = Format(ComputeAmount(txtSplHoliday.Text, PaymentBasis, False, 5), "###,##0.00")
    txtSplHoliday.BackColor = &HFFFFFF
End Sub

Private Sub txtSplOTHours_GotFocus()
    SendKeys "{HOME}+{END}"
    txtSplOTHours.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtSplOTHours
End Sub

Private Sub txtSplOTHours_LostFocus()
    txtAmount(7).BackColor = &HFFFFFF
    txtAmount(7).Text = Format(ComputeAmount(txtSplOTHours.Text, PaymentBasis, True, 7), "###,##0.00")
    txtSplOTHours.BackColor = &HFFFFFF
End Sub

Private Sub txtSundayOT_GotFocus()
    SendKeys "{HOME}+{END}"
    txtSundayOT.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtSundayOT
End Sub

Private Sub txtSundayOT_LostFocus()
    txtAmount(4).BackColor = &HFFFFFF
    txtAmount(4).Text = Format(ComputeAmount(txtSundayOT.Text, PaymentBasis, True, 4), "###,##0.00")
    txtSundayOT.BackColor = &HFFFFFF
End Sub


Private Sub txtTAX_GotFocus()
    SendKeys "{HOME}+{END}"
    txtTAX.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtTAX
End Sub

Private Sub txtTAX_LostFocus()
    txtTAX.BackColor = &HC0C0C0
    txtTAX.ForeColor = &HFF0000
    Call txtGrossPay_Change
    txtTAX.Locked = True
    cmdEditTax(3).SetFocus
End Sub

Private Sub txtVacLeave_GotFocus()
    SendKeys "{HOME}+{END}"
    txtVacLeave.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtVacLeave
End Sub

Private Sub txtVacLeave_LostFocus()
    txtAmount(10).BackColor = &HFFFFFF
    txtAmount(10).Text = Format(ComputeAmount(txtVacLeave.Text, PaymentBasis, False, 10), "###,##0.00")
    txtVacLeave.BackColor = &HFFFFFF
End Sub

Private Sub txtVale_GotFocus()
    SendKeys "{HOME}+{END}"
    txtVale.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtVale
End Sub

Private Sub txtVale_LostFocus()
    txtVale.BackColor = &HFFFFFF
    Call txtGrossPay_Change
End Sub

Public Sub SaveTransaction()
Dim mECCAmt As Double
Dim mOtherLoan As Double

mOtherLoan = 0#
MsgBox Format(frmSelectCutoff.txtCutOffDate(0).Text, "m")
        
strQuery = "INSERT INTO payrollmaster (employeenum, agencycode, clientcode, dayspresent, basicpay, " & _
        "othours, otpay, excessothours, excessotpay, nightdiffot, nightdiffpay, " & _
        "restdayothrs, restdayotpay, sundayot, sundayotpay, nosplholiday, splholidaypay, " & _
        "splholidayothours, splholidayotpay, excesssplholothours, excesssplholotpay, " & _
        "nolglholiday, lglholidaypay, lglholidayothours, lglholidayotpay, excesslglholothours, excesslglholotpay, otwithoutpremium, otwithoutpremiumpay, novacationleave, " & _
        "vacationleavepay, nosickleave, sickleavepay, salaryadjust, ecolapay, grosspay, daysabsent, absentdeduction, hourstardiness, tardinessdeduction, " & _
        "cashbond, cardtransfee, valewithclient, placementfee, miscfee, overunderpay, " & _
        "sssemployee, sssemployer, eccemployer, phealthemployee, phealthemployer, pagibigemployee, " & _
        "pagibigemployer, wtaxamount, sssloan, pagibigloan, otherloan, deductions, netpay, " & _
        "startpaydate, endpaydate, monthno, cutoffno, payflag1, payflag2,paytext1, paytext2, entryby, entrydate )" & _
        " VALUES ('" & EmployeeNo & "', " & mAgencyCode & ", " & mClientCode & ", " & ChkVal(txtDaysWork) & ", " & ChkVal(txtAmount(0)) & _
        ", " & ChkVal(txtRegOT) & ", " & ChkVal(txtAmount(1)) & ", " & ChkVal(txtExcessOTRegDay) & ", " & ChkVal(txtAmount(12)) & ", " & ChkVal(txtNightDiff) & ", " & ChkVal(txtAmount(2)) & _
        ", " & ChkVal(txtRestDay) & ", " & ChkVal(txtAmount(3)) & ", " & ChkVal(txtSundayOT) & ", " & ChkVal(txtAmount(4)) & ", " & ChkVal(txtSplHoliday) & ", " & ChkVal(txtAmount(5)) & _
        ", " & ChkVal(txtSplOTHours) & ", " & ChkVal(txtAmount(7)) & ", " & ChkVal(txtExcessOTSplSun) & ", " & ChkVal(txtAmount(13)) & _
        ", " & ChkVal(txtLglHoliday) & ", " & ChkVal(txtAmount(6)) & ", " & ChkVal(txtLglHolidayOT) & ", " & ChkVal(txtAmount(8)) & ", " & ChkVal(txtExcessOTLegal) & ", " & ChkVal(txtAmount(15)) & ", " & ChkVal(txtOTwoPremium) & ", " & ChkVal(txtAmount(16)) & ", " & ChkVal(txtVacLeave) & _
        ", " & ChkVal(txtAmount(10)) & ", " & ChkVal(txtSickLeave) & ", " & ChkVal(txtAmount(9)) & ", " & ChkVal(txtSalAdjustment) & ", " & ChkVal(txtECOLA) & ", " & ChkVal(txtGrossPay) & ", " & ChkVal(txtDaysAbsent) & ", " & ChkVal(txtDeductions(0)) & ", " & ChkVal(txtHoursLate) & ", " & ChkVal(txtDeductions(1)) & _
        ", " & ChkVal(txtCashBond) & ", " & ChkVal(txtBDOCard) & ", " & ChkVal(txtVale) & ", " & ChkVal(txtPlacementFee) & ", " & ChkVal(txtMiscFee) & ", " & ChkVal(txtOverPay) & _
        ", " & ChkVal(txtEmpyeeSSS) & ", " & ChkVal(txtEmpyerSSS) & ", " & Format(mECCAmt, "0.00") & ", " & ChkVal(txtPhilHealthYee) & ", " & ChkVal(txtPHealthYer) & ", " & ChkVal(txtEmpyeePagibig) & _
        ", " & ChkVal(txtEmpyerPagibig) & ", " & ChkVal(txtTAX) & ", " & ChkVal(txtSSSLoan) & ", " & ChkVal(txtPagibigLoan) & ", " & Format(mOtherLoan, "0.00") & ", " & ChkVal(txtDeductions(2)) & ", " & ChkVal(txtNetPay) & _
        ", '" & Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy") & "', '" & Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy") & "', " & Format(frmSelectCutoff.txtCutOffDate(0).Text, "m") & ", " & mCutOffNo & ", 0, 0,'','', '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & "')"

Call Action_Query(strQuery, Cnnstr)

'strQuery = "INSERT INTO payrollmaster (employeenum, agencycode, clientcode, dayspresent, basicpay, othours, otpay, excessothours, excessotpay, nightdiffot, nightdiffpay, " & _
'        "restdayothrs, restdayotpay, sundayot, sundayotpay, nosplholiday, splholidaypay, splholidayothours, splholidayotpay, excesssplholothours, excesssplholotpay, " & _
'        "nolglholiday, lglholidaypay, lglholidayothours, lglholidayotpay, excesslglholothours, excesslglholotpay, otwithoutpremium, otwithoutpremiumpay, novacationleave, " & _
'        "vacationleavepay, nosickleave, sickleavepay, salaryadjust, ecolapay, grosspay, daysabsent, absentdeduction, hourstardiness, tardinessdeduction, " & _
'        "cashbond, cardtransfee, valewithclient, placementfee, miscfee, overunderpay, " & _
'        "sssemployee, sssemployer, eccemployer, phealthemployee, phealthemployer, pagibigemployee, " & _
'        "pagibigemployer, wtaxamount, sssloan, pagibigloan, otherloan, deductions, netpay, " & _
'        "startpaydate, endpaydate, monthno, cutoffno, payflag1, payflag2,paytext1, paytext2, entryby, entrydate )" & _
'        " VALUES ('" & EmployeeNo & "', " & mAgencyCode & ", " & mClientCode & ", " & Format(txtDaysWork, "0.00") & ", " & Format(txtAmount(0), "0.00") & _
'        ", " & Format(txtRegOT, "0.00") & ", " & Format(txtAmount(1), "0.00") & ", " & Format(txtExcessOTRegDay, "0.00") & ", " & Format(txtAmount(12), "0.00") & ", " & Format(txtNightDiff, "0.00") & ", " & Format(txtAmount(2), "0.00") & _
'        ", " & Format(txtRestDay, "0.00") & ", " & Format(txtAmount(3), "0.00") & ", " & Format(txtSundayOT, "0.00") & ", " & Format(txtAmount(4), "0.00") & ", " & Format(txtSplHoliday, "0.00") & ", " & Format(txtAmount(5), "0.00") & "," & Format(txtExcessOTSplSun, "0.00") & ", " & Format(txtAmount(13), "0.00") & _
'        ", " & Format(txtLglHoliday, "0.00") & ", " & Format(txtAmount(6), "0.00") & ", " & Format(txtLglHolidayOT, "0.00") & ", " & Format(txtAmount(8), "0.00") & ", " & Format(txtExcessOTLegal, "0.00") & ", " & Format(txtAmount(15), "0.00") & ", " & Format(txtOTwoPremium, "0.00") & ", " & Format(txtAmount(16), "0.00") & ", " & Format(txtVacLeave, "0.00") & _
'        ", " & Format(txtAmount(10), "0.00") & ", " & Format(txtSickLeave, "0.00") & ", " & Format(txtAmount(9), "0.00") & ", " & Format(txtSalAdjustment, "0.00") & ", " & Format(txtECOLA, "0.00") & ", " & Format(txtGrossPay, "0.00") & ", " & Format(txtDaysAbsent, "0.00") & ", " & Format(txtDeductions(0), "0.00") & ", " & Format(txtHoursLate, "0.00") & ", " & Format(txtDeductions(1), "0.00") & _
'        ", " & Format(txtCashBond, "0.00") & ", " & Format(txtBDOCard, "0.00") & ", " & Format(txtVale, "0.00") & ", " & Format(txtPlacementFee, "0.00") & ", " & Format(txtMiscFee, "0.00") & ", " & Format(txtOverPay, "0.00") & _
'        ", " & Format(txtEmpyeeSSS, "0.00") & ", " & Format(txtEmpyerSSS, "0.00") & ", " & Format(mECCAmt, "0.00") & ", " & Format(txtPhilHealthYee, "0.00") & ", " & Format(txtPHealthYer, "0.00") & ", " & Format(txtEmpyeePagibig, "0.00") & _
'        ", " & Format(txtEmpyerPagibig, "0.00") & ", " & Format(txtTAX, "0.00") & ", " & Format(txtSSSLoan, "0.00") & ", " & Format(txtPagibigLoan, "0.00") & ", " & Format(mOtherLoan, "0.00") & ", " & Format(txtDeductions(2), "0.00") & ", " & Format(txtNetPay, "0.00") & _
'        ", '" & Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy") & "', '" & Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy") & "', " & Format(frmSelectCutoff.txtCutOffDate(0).Text, "m") & ", " & mCutOffNo & ", 0, 0, , , '" & mUserID & "', '" & Format(Now, "mm/dd/yyyy") & "')"


End Sub

Private Function ChkVal(ByRef mTBox As TextBox) As Double
Dim i As Integer
Dim strAmt As String
Dim mAmt As String
Dim DeciSw As Boolean

    mAmt = Trim(Replace(mTBox, ",", ""))
    For i = 1 To Len(mAmt)
        If DeciSw = True And Mid(mAmt, i, 1) = "." Then
        Else
            strAmt = strAmt & Mid(mAmt, i, 1)
        End If
        If DeciSw = False And Mid(mAmt, i, 1) = "." Then
            DeciSw = True
        End If
    Next
    mAmt = IIf(strAmt = "." Or strAmt = "0" Or strAmt = "", 0#, strAmt)
    ChkVal = Format(mAmt, "0.00")
End Function

