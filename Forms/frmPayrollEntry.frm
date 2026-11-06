VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPayrollEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Transaction Entries"
   ClientHeight    =   6720
   ClientLeft      =   2925
   ClientTop       =   2955
   ClientWidth     =   11430
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmPayrollEntry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   11430
   Begin VB.Frame FrameDeductions 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Other OT / Adjustment"
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
      Height          =   3735
      Left            =   5160
      TabIndex        =   174
      Top             =   2880
      Visible         =   0   'False
      Width           =   3645
      Begin VB.TextBox txtCalLoan 
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
         Left            =   1680
         TabIndex        =   36
         Text            =   "0"
         Top             =   3240
         Width           =   855
      End
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
         Left            =   1680
         TabIndex        =   29
         Text            =   "0"
         Top             =   720
         Width           =   855
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
         Height          =   285
         Index           =   12
         Left            =   2640
         TabIndex        =   204
         Text            =   "0.00"
         Top             =   720
         Width           =   885
      End
      Begin VB.TextBox txtAdj 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   201
         Text            =   "0.00"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtOtherAdj 
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
         Left            =   1680
         TabIndex        =   34
         Text            =   "0"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtHazardAdj 
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
         Left            =   1680
         TabIndex        =   33
         Text            =   "0"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtEcolaAdj 
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
         Left            =   1680
         TabIndex        =   32
         Text            =   "0"
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtUnpaidSL 
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
         Left            =   1680
         TabIndex        =   31
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtAdj 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   196
         Text            =   "0.00"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtAdj 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   195
         Text            =   "0.00"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtAdj 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   194
         Text            =   "0.00"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtAdj 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   193
         Text            =   "0.00"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtUnpaidVL 
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
         Left            =   1680
         TabIndex        =   30
         Text            =   "0"
         Top             =   1320
         Width           =   855
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
         Height          =   285
         Index           =   21
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   189
         Text            =   "0.00"
         Top             =   2760
         Width           =   885
      End
      Begin VB.TextBox txtAdjustments 
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
         Left            =   1680
         TabIndex        =   35
         Text            =   "0"
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtRDonLGL 
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
         Left            =   1680
         TabIndex        =   28
         Text            =   "0"
         Top             =   465
         Width           =   855
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
         Height          =   285
         Index           =   20
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   188
         Text            =   "0.00"
         Top             =   465
         Width           =   885
      End
      Begin VB.TextBox txtRDonSPL 
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
         Left            =   1680
         TabIndex        =   27
         Text            =   "0"
         Top             =   195
         Width           =   855
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
         Height          =   285
         Index           =   19
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   186
         Text            =   "0.00"
         Top             =   195
         Width           =   885
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
         Height          =   285
         Index           =   18
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   184
         Text            =   "0.00"
         Top             =   -360
         Width           =   885
      End
      Begin VB.TextBox txtRDSplLgl 
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
         Left            =   1680
         TabIndex        =   26
         Text            =   "0"
         Top             =   -360
         Width           =   855
      End
      Begin MSComctlLib.ListView ListDeductions 
         Height          =   345
         Left            =   120
         TabIndex        =   175
         Top             =   3960
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   609
         View            =   3
         LabelEdit       =   1
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
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "SSS Cal. Loan:"
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
         Index           =   68
         Left            =   120
         TabIndex        =   206
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Excess Overtime:"
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
         Index           =   67
         Left            =   120
         TabIndex        =   203
         Top             =   750
         Width           =   1575
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Others:"
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
         Index           =   66
         Left            =   120
         TabIndex        =   202
         Top             =   2310
         Width           =   1575
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Adj:"
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
         Index           =   65
         Left            =   120
         TabIndex        =   199
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Hazard Adj:"
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
         Index           =   64
         Left            =   120
         TabIndex        =   198
         Top             =   2070
         Width           =   1575
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Ecola Adj:"
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
         Index           =   63
         Left            =   120
         TabIndex        =   197
         Top             =   1830
         Width           =   1575
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Unpaid SL:"
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
         Index           =   62
         Left            =   120
         TabIndex        =   192
         Top             =   1590
         Width           =   1575
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Unpaid VL:"
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
         Index           =   61
         Left            =   120
         TabIndex        =   191
         Top             =   1350
         Width           =   1575
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustments:"
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
         Index           =   60
         Left            =   120
         TabIndex        =   190
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "RD on LGL Hours:"
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
         Index           =   59
         Left            =   120
         TabIndex        =   187
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "RD on SPL Hours:"
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
         Index           =   58
         Left            =   120
         TabIndex        =   185
         Top             =   195
         Width           =   1575
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "RD Spl/Lgl OT Hrs:"
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
         Index           =   57
         Left            =   120
         TabIndex        =   182
         Top             =   -360
         Width           =   1575
      End
   End
   Begin VB.TextBox txtHDMFCal 
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
      TabIndex        =   183
      Text            =   "0"
      Top             =   5400
      Width           =   765
   End
   Begin VB.TextBox txtCalamityLoan 
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
      TabIndex        =   181
      Text            =   "0"
      Top             =   5400
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   7050
      TabIndex        =   107
      Top             =   5565
      Width           =   3045
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
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   113
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
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   111
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
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   110
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
         Left            =   150
         TabIndex        =   112
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
         Left            =   150
         TabIndex        =   109
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
         Left            =   150
         TabIndex        =   108
         Top             =   450
         Width           =   1545
      End
   End
   Begin VB.TextBox txtNoDays 
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
      Left            =   2040
      TabIndex        =   180
      Text            =   "0"
      Top             =   2310
      Width           =   495
   End
   Begin VB.TextBox txtUnionDues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   38
      Text            =   "0"
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtHazardPay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   37
      Text            =   "0"
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdDTR 
      Caption         =   "DTR"
      Height          =   285
      Left            =   1560
      TabIndex        =   177
      ToolTipText     =   "Read DTR Entries"
      Top             =   2310
      Width           =   495
   End
   Begin VB.TextBox txtECC 
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
      Left            =   6120
      TabIndex        =   176
      Text            =   "0"
      Top             =   6720
      Visible         =   0   'False
      Width           =   435
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
      Left            =   5760
      TabIndex        =   114
      Top             =   7440
      Width           =   4965
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
         Left            =   2160
         TabIndex        =   40
         Text            =   "0"
         Top             =   420
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
         TabIndex        =   124
         Text            =   "0.00"
         Top             =   450
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
         TabIndex        =   123
         Text            =   "0.00"
         Top             =   720
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
         TabIndex        =   122
         Text            =   "0.00"
         Top             =   960
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
         TabIndex        =   121
         Text            =   "0.00"
         Top             =   1260
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
         Left            =   2160
         TabIndex        =   42
         Text            =   "0"
         Top             =   690
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
         Left            =   2160
         TabIndex        =   43
         Text            =   "0"
         Top             =   960
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
         Left            =   2160
         TabIndex        =   44
         Text            =   "0"
         Top             =   1230
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
         Index           =   49
         Left            =   3330
         TabIndex        =   156
         Top             =   180
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
         Height          =   195
         Index           =   46
         Left            =   3330
         TabIndex        =   128
         Top             =   720
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
         Height          =   195
         Index           =   47
         Left            =   3330
         TabIndex        =   127
         Top             =   980
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
         Height          =   195
         Index           =   48
         Left            =   3330
         TabIndex        =   126
         Top             =   450
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
         Height          =   195
         Index           =   52
         Left            =   3330
         TabIndex        =   125
         Top             =   1250
         Width           =   495
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
         TabIndex        =   120
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
         TabIndex        =   119
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
         TabIndex        =   118
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
         TabIndex        =   117
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
         TabIndex        =   116
         Top             =   240
         Width           =   2085
      End
   End
   Begin VB.CheckBox chkOtherDed 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Others"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8880
      TabIndex        =   173
      Top             =   2950
      Width           =   1215
   End
   Begin VB.CheckBox chkClient 
      Caption         =   "Check1"
      Height          =   255
      Left            =   5400
      TabIndex        =   170
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkUpdate 
      Caption         =   "Check1"
      Height          =   255
      Left            =   5040
      TabIndex        =   169
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtOverPay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9180
      TabIndex        =   25
      Text            =   "0"
      Top             =   3750
      Width           =   855
   End
   Begin VB.TextBox txtMiscFee 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6840
      TabIndex        =   22
      Text            =   "0"
      Top             =   3750
      Width           =   855
   End
   Begin VB.TextBox txtProcessDate 
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   164
      Text            =   "Date"
      Top             =   6390
      Width           =   1305
   End
   Begin VB.TextBox txtProcessBy 
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   163
      Text            =   "By"
      Top             =   6360
      Width           =   1305
   End
   Begin VB.TextBox txtIncentiveRemarks 
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
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   17
      Top             =   5460
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox txt13thMonth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   6090
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
      Height          =   285
      Index           =   17
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   158
      Text            =   "0.00"
      Top             =   3255
      Width           =   1125
   End
   Begin VB.TextBox txtECOLA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Text            =   "0"
      Top             =   3255
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1680
      TabIndex        =   155
      Text            =   "0"
      Top             =   6000
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox txtSplOtHours 
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2400
      TabIndex        =   14
      Text            =   "0"
      Top             =   5460
      Width           =   855
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
      Left            =   2040
      TabIndex        =   5
      Text            =   "0"
      Top             =   2940
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
      Height          =   285
      Index           =   2
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   153
      Text            =   "0.00"
      Top             =   2940
      Width           =   1125
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
      Left            =   2040
      TabIndex        =   4
      Text            =   "0"
      Top             =   2625
      Width           =   1215
   End
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
      Height          =   195
      Left            =   3800
      Locked          =   -1  'True
      TabIndex        =   151
      Text            =   "0.00"
      Top             =   6405
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
      Left            =   2520
      TabIndex        =   3
      Text            =   "0"
      Top             =   2310
      Width           =   735
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
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
      Height          =   285
      Index           =   0
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   150
      Text            =   "0.00"
      Top             =   2310
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
      Height          =   285
      Index           =   3
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   137
      Text            =   "0.00"
      Top             =   5780
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
      Height          =   285
      Index           =   4
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   136
      Text            =   "0.00"
      Top             =   3570
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
      Height          =   285
      Index           =   5
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   135
      Text            =   "0.00"
      Top             =   3885
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
      Height          =   285
      Index           =   6
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   134
      Text            =   "0.00"
      Top             =   4200
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
      Height          =   285
      Index           =   7
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   133
      Text            =   "0.00"
      Top             =   4515
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
      Height          =   285
      Index           =   8
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   132
      Text            =   "0.00"
      Top             =   4830
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
      Height          =   285
      Index           =   9
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   131
      Text            =   "0.00"
      Top             =   5145
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
      Height          =   285
      Index           =   10
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   130
      Text            =   "0.00"
      Top             =   5460
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
      Height          =   285
      Index           =   11
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   129
      Text            =   "0.00"
      Top             =   6090
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   10270
      TabIndex        =   115
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton cmdRefreshRecords 
         Caption         =   "Refresh Records"
         Height          =   675
         Left            =   65
         TabIndex        =   200
         ToolTipText     =   "Configure Rates"
         Top             =   2280
         Width           =   945
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   555
         Index           =   3
         Left            =   65
         Picture         =   "frmPayrollEntry.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   168
         ToolTipText     =   "Last"
         Top             =   1740
         Width           =   945
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   555
         Index           =   2
         Left            =   65
         Picture         =   "frmPayrollEntry.frx":0DFC
         Style           =   1  'Graphical
         TabIndex        =   167
         ToolTipText     =   "Next"
         Top             =   1200
         Width           =   945
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   555
         Index           =   1
         Left            =   65
         Picture         =   "frmPayrollEntry.frx":0EF6
         Style           =   1  'Graphical
         TabIndex        =   166
         ToolTipText     =   "Previous"
         Top             =   660
         Width           =   945
      End
      Begin VB.CommandButton cmdNavigation 
         Height          =   555
         Index           =   0
         Left            =   65
         Picture         =   "frmPayrollEntry.frx":0FF0
         Style           =   1  'Graphical
         TabIndex        =   165
         ToolTipText     =   "First"
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Configure Client"
         Height          =   555
         Left            =   65
         TabIndex        =   47
         Top             =   4305
         Width           =   945
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   555
         Left            =   65
         TabIndex        =   50
         ToolTipText     =   "Close"
         Top             =   5985
         Width           =   945
      End
      Begin VB.CommandButton cmdmovetoMaster 
         Caption         =   "Move to Master"
         Height          =   555
         Left            =   65
         TabIndex        =   49
         ToolTipText     =   "Move to Master"
         Top             =   5415
         Width           =   945
      End
      Begin VB.CommandButton cmdPrintReports 
         Caption         =   "Payroll Reports"
         Height          =   555
         Left            =   65
         TabIndex        =   48
         ToolTipText     =   "Payroll Reports"
         Top             =   4860
         Width           =   945
      End
      Begin VB.CommandButton cmdEditTax 
         Caption         =   "Edit OT && ND"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   3
         Left            =   65
         Picture         =   "frmPayrollEntry.frx":1522
         TabIndex        =   45
         ToolTipText     =   "Edit Agency Charges"
         Top             =   2955
         Width           =   945
      End
      Begin VB.CommandButton cmdFixedRates 
         Caption         =   "Configure Fixed Rates"
         Height          =   675
         Left            =   65
         TabIndex        =   46
         ToolTipText     =   "Configure Rates"
         Top             =   3630
         Width           =   945
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
      TabIndex        =   104
      Text            =   "0"
      Top             =   5100
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
      TabIndex        =   103
      Text            =   "0"
      Top             =   5100
      Width           =   675
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
      TabIndex        =   99
      Text            =   "0"
      Top             =   4830
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
      TabIndex        =   98
      Text            =   "0"
      Top             =   4830
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
      TabIndex        =   97
      Text            =   "0"
      Top             =   5370
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
      TabIndex        =   91
      Text            =   "0"
      Top             =   4290
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
      TabIndex        =   90
      Text            =   "0"
      Top             =   4290
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
      TabIndex        =   89
      Text            =   "0"
      Top             =   4560
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
      TabIndex        =   88
      Text            =   "0"
      Top             =   4560
      Width           =   765
   End
   Begin VB.TextBox txtPlacementFee 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9180
      TabIndex        =   24
      Text            =   "0"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtVale 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6840
      TabIndex        =   21
      Text            =   "0"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtBDOCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9180
      TabIndex        =   23
      Text            =   "0"
      Top             =   3210
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
      Height          =   285
      Index           =   1
      Left            =   8820
      Locked          =   -1  'True
      TabIndex        =   81
      Text            =   "0.00"
      Top             =   2625
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
      Height          =   285
      Index           =   0
      Left            =   8820
      Locked          =   -1  'True
      TabIndex        =   80
      Text            =   "0.00"
      Top             =   2310
      Width           =   1125
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
      Left            =   2040
      TabIndex        =   10
      Text            =   "0"
      Top             =   4515
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
      Left            =   2040
      TabIndex        =   11
      Text            =   "0"
      Top             =   4830
      Width           =   1215
   End
   Begin VB.TextBox txtSalAdjustment 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Text            =   "0"
      Top             =   5145
      Width           =   1215
   End
   Begin VB.TextBox txtIncentive 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Text            =   "0"
      Top             =   5460
      Width           =   855
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
      Height          =   285
      Index           =   1
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   70
      Text            =   "0.00"
      Top             =   2625
      Width           =   1125
   End
   Begin VB.TextBox txtTASAesla 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6840
      TabIndex        =   20
      Text            =   "0"
      Top             =   3210
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6900
      TabIndex        =   19
      Text            =   "0"
      Top             =   2625
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6900
      TabIndex        =   18
      Text            =   "0"
      Top             =   2310
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
      Left            =   2040
      TabIndex        =   9
      Text            =   "0"
      Top             =   4200
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
      Left            =   2040
      TabIndex        =   8
      Text            =   "0"
      Top             =   3885
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
      Left            =   2040
      TabIndex        =   7
      Text            =   "0"
      Top             =   3570
      Width           =   1215
   End
   Begin VB.TextBox txtTaxRefund 
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
      Left            =   2040
      TabIndex        =   15
      Text            =   "0"
      Top             =   5780
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H8000000F&
      Height          =   2175
      Left            =   120
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
         Index           =   1
         Left            =   5010
         TabIndex        =   41
         Top             =   1800
         Width           =   2025
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
         Left            =   3000
         TabIndex        =   39
         Top             =   1800
         Width           =   2025
      End
      Begin MSComctlLib.ListView lvwEmployeeList 
         Height          =   1665
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   2937
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
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lvwProcessEmp 
         Height          =   1665
         Left            =   7650
         TabIndex        =   2
         Top             =   390
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   2937
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
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Hired:"
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
         TabIndex        =   149
         Top             =   1530
         Width           =   1185
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
         TabIndex        =   148
         Top             =   1275
         Width           =   1185
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Emp. Name:"
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
         TabIndex        =   147
         Top             =   1020
         Width           =   1185
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Emp. No:"
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
         TabIndex        =   146
         Top             =   780
         Width           =   1185
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
         Left            =   3675
         TabIndex        =   145
         Top             =   1530
         Width           =   3885
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
         TabIndex        =   79
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
         Left            =   3675
         TabIndex        =   59
         Top             =   1275
         Width           =   3885
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
         Left            =   3690
         TabIndex        =   58
         Top             =   1020
         Width           =   3885
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
         Left            =   3675
         TabIndex        =   57
         Top             =   780
         Width           =   3885
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
         Left            =   7650
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
         Top             =   150
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
         TabIndex        =   52
         Top             =   480
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
         TabIndex        =   51
         Top             =   2250
         Width           =   2175
      End
   End
   Begin VB.CheckBox chkEcolaHol 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3000
      TabIndex        =   171
      ToolTipText     =   "Include/Exclude ECOLA payment"
      Top             =   3890
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkEcolaLegal 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3000
      TabIndex        =   172
      ToolTipText     =   "Include/Exclude ECOLA payment"
      Top             =   4200
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labelLeave 
      BackStyle       =   0  'Transparent
      Caption         =   "BSC:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   205
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Union Dues:"
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
      Index           =   56
      Left            =   5160
      TabIndex        =   179
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Hazard Pay:"
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
      Index           =   55
      Left            =   5160
      TabIndex        =   178
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   1
      X1              =   135
      X2              =   10170
      Y1              =   6630
      Y2              =   6645
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
      TabIndex        =   162
      Top             =   5490
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
      Height          =   285
      Index           =   24
      Left            =   3360
      TabIndex        =   161
      Top             =   6135
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
      Height          =   285
      Index           =   30
      Left            =   3360
      TabIndex        =   160
      Top             =   5820
      Width           =   855
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Surety Fund:"
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
      Index           =   54
      Left            =   7650
      TabIndex        =   159
      Top             =   5370
      Width           =   1455
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Refund:"
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
      TabIndex        =   157
      Top             =   5820
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
      Index           =   14
      Left            =   3360
      TabIndex        =   154
      Top             =   2970
      Width           =   945
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
      Left            =   255
      TabIndex        =   152
      Top             =   2685
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
      Index           =   15
      Left            =   3360
      TabIndex        =   144
      Top             =   3285
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
      TabIndex        =   143
      Top             =   3600
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
      TabIndex        =   142
      Top             =   3915
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
      TabIndex        =   141
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
      Index           =   26
      Left            =   3360
      TabIndex        =   140
      Top             =   5175
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
      TabIndex        =   139
      Top             =   4860
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
      TabIndex        =   138
      Top             =   4545
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
      TabIndex        =   106
      Top             =   5100
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
      TabIndex        =   105
      Top             =   5100
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
      TabIndex        =   102
      Top             =   4830
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
      TabIndex        =   101
      Top             =   4830
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
      TabIndex        =   100
      Top             =   5370
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
      TabIndex        =   96
      Top             =   4290
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
      TabIndex        =   95
      Top             =   4020
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
      TabIndex        =   94
      Top             =   4290
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
      TabIndex        =   93
      Top             =   4560
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
      TabIndex        =   92
      Top             =   4560
      Width           =   1545
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Uniform:"
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
      TabIndex        =   87
      Top             =   3750
      Width           =   1455
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Planters Loan:"
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
      TabIndex        =   86
      Top             =   3750
      Width           =   1635
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Advance:"
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
      TabIndex        =   85
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Vale:"
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
      TabIndex        =   84
      Top             =   3480
      Width           =   1635
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "LMC Deductions:"
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
      TabIndex        =   83
      Top             =   3210
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions:"
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
      TabIndex        =   82
      Top             =   2940
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
      Index           =   33
      Left            =   255
      TabIndex        =   78
      Top             =   4575
      Width           =   1815
   End
   Begin VB.Label lblEntries 
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
      Index           =   32
      Left            =   255
      TabIndex        =   77
      Top             =   4890
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Income:"
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
      Left            =   255
      TabIndex        =   76
      Top             =   5205
      Width           =   1815
   End
   Begin VB.Label labelLeave 
      BackStyle       =   0  'Transparent
      Caption         =   "Allowance:"
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
      Left            =   255
      TabIndex        =   75
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "13th Month Pay:"
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
      Left            =   255
      TabIndex        =   74
      Top             =   6135
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
      TabIndex        =   73
      Top             =   2685
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
      TabIndex        =   72
      Top             =   2370
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
      Index           =   13
      Left            =   3360
      TabIndex        =   71
      Top             =   2655
      Width           =   945
   End
   Begin VB.Label lblEntries 
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
      Height          =   255
      Index           =   12
      Left            =   5220
      TabIndex        =   69
      Top             =   3210
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
      TabIndex        =   68
      Top             =   2685
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
      TabIndex        =   67
      Top             =   2370
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
      TabIndex        =   66
      Top             =   2340
      Width           =   945
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   4
      X1              =   10170
      X2              =   10170
      Y1              =   2220
      Y2              =   6580
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   3
      X1              =   5040
      X2              =   5040
      Y1              =   2280
      Y2              =   6580
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   2250
      Y2              =   6580
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Spl. Holiday Hours:"
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
      Left            =   255
      TabIndex        =   65
      Top             =   4230
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg. Holiday Hours:"
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
      Left            =   255
      TabIndex        =   64
      Top             =   3915
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Restday Hours:"
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
      Left            =   255
      TabIndex        =   63
      Top             =   3630
      Width           =   1815
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
      Index           =   3
      Left            =   255
      TabIndex        =   62
      Top             =   3315
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
      Left            =   255
      TabIndex        =   61
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Hours:"
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
      Left            =   255
      TabIndex        =   60
      Top             =   2370
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   0
      X1              =   135
      X2              =   10120
      Y1              =   2220
      Y2              =   2235
   End
End
Attribute VB_Name = "frmPayrollEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TotalAmt As Double

Dim ComputationBase As Boolean
Dim rsCompanyInfo As ADODB.Recordset
Dim rsClientInfo As ADODB.Recordset     ' client table
Dim rsEmployeeInfo As ADODB.Recordset
Dim rsProcessEmployee As ADODB.Recordset
Dim rsPreviousEmployee As ADODB.Recordset   ' get previous employee payroll
Dim OrigMonthlyRate As Double   ' Save OrigMonthlyRate
Dim AddLeaveFlag As Boolean
Dim HazardFlag As Boolean

Dim rsSSSTable As ADODB.Recordset
Dim rsTaxTable As ADODB.Recordset   ' for monthly paid employee
Dim rsTaxTable2 As ADODB.Recordset  ' for daily paid employee
Dim rsPHealthTable As ADODB.Recordset
Dim rsPagibigTable As ADODB.Recordset
Dim rsEmpDeductions As ADODB.Recordset  ' recordset for deductions

Dim rsLoanDed As ADODB.Recordset            ' to get the deductions from table Bernard 1/21/2011
Dim arrLoanPayments(1 To 20) As Double      ' to save the payments per loan type
Dim StartCutoffDate, EndCutoffDate          ' make it available in getloan function 1/21/2011

Dim PayrollValidation As clsCommon
Dim StrQuery As String
Dim RetSw As Boolean
Dim retClientSw As Boolean

Dim ListViewFlag As Boolean
Dim AddFlag As Boolean
Dim EditFlag As Boolean
Dim mCtr1 As Integer            ' ctr for employee list
Dim mCtr2 As Integer            ' ctr for process employee
Dim mBilledFlag As Boolean      ' to determine whether the payroll period was already billed
Dim NewCutoffFlag As Boolean    ' flag for new cut-off
Dim InsertedFlag As Boolean     ' flag to determine whether new cut-off was inserted
Dim WithSSSDedFlag As Boolean   ' flag whether SSS deduction will apply
Dim WithPHDedFlag As Boolean    ' flag whether philhealth deduction will apply
Dim WithPBIGDedFlag As Boolean  ' flag whether pagibig deduction will apply

'Get Agency Information.
Dim mAgencyCode As Integer
Dim AgencySw As Boolean

Dim DailyRate As Double

'Get Employee Information.
Dim EmployeeName As String
Dim EmployeeNo As Integer
Dim EmpNo As String
Dim mECOLARate As Double
Dim HourlyRate As Double
Dim MonthlyRate As Double
Dim PaymentBasis As String
Dim MonthlyFlag As Boolean
Dim EmpTaxCode As String

Dim BDOCardFee As Double
Dim TASAeslaFee As Double
Dim EmployeeAgencyRate As Double
Dim mEndContract As Date
Dim OTNoTax As Boolean
Dim EcolaInOTRate As Boolean
Dim SSSDedIsMonthly As Boolean

'Loans with scheduled date.
Dim mSSSLoan As Double
Dim mPBigLoan As Double
Dim mCalamityLoan As Double, mAllowance As Double, NewPHRate As Double
Dim mSSSCalLoan As Double   ' for SSS Calamity
Dim mHDMFCalLoan As Double  ' for HDMF Calamity
Dim mMFAmount As Double
Dim mCBAmount As Double
Dim mPFAmount As Double, mBSCAllowRate As Double ' for BSC Allowance 2
Dim mUniform As Double
Dim mCashAdvance As Double
Dim mVale As Double             ' updated by bernard 1/21/2011

'Get Client Information.
Dim ClientCutoff As String
Dim WSSSFlag As Boolean
Dim WPHealthFlag As Boolean
Dim WPagibigFlag As Boolean
Dim mPagibigPay As Double
Dim WTaxFlag As Boolean, EmpTaxFlag As Boolean, UseTax2018Flag As Boolean
Dim EditTaxFlag As Boolean
Dim TaxPercentage As Double
Dim ClientTaxPercentageFlag As Boolean
Dim NoofCutoffs As Integer
Dim mCutOffNo As Integer
Dim WEcolaFlag As Boolean
Dim WTaxableEcola As Boolean, WTaxableAllowance As Boolean
Dim WEcolaDeduction As Boolean
Dim DoubleEcolaonSpl As Boolean
Dim DoubleEcolaonLgl As Boolean
Dim EcolaPayonRestDay As Boolean
Dim EcolaPayonHoliday As Boolean
Dim WCalamityLoan As Boolean
Dim mExcludedTax As Boolean
Dim mOTExcluded As Boolean
Dim mHalfSSSPay As Boolean
Dim AgencyRate As Double
Dim MonthlyEcola As Double

Dim ClientSSSRate As Double
Dim ClientPHRate As Double
Dim ClientPagibigRate As Double

Dim WorkHrPerDay As Double
Dim NoofDaysperMonth As Double
Dim WorkDaysperYear As Double

Dim W13MonthPay As Boolean
Dim Hourly13thMonthRate As Double

Dim OverTimeRate As Double
Dim SundayRate As Double
Dim TardinessRate As Double
Dim NDiffRate As Double
Dim SplHolidayRate As Double
Dim LglHolidayRate As Double
Dim SplHolidayOTRate As Double
Dim LglHolidayOTRate As Double
Dim ExcessOTRegDayRate As Double
Dim ExcessOTSplSunRate As Double
Dim ExcessOTSunLglRate As Double
Dim ExcessOTLegalRate As Double
Dim OTwoPremiumRate As Double
Dim RestDayRate As Double
Dim mTotalDeductions As Double
Dim InsField
Dim InsValue
Dim EdtField
Dim StartDed As String
Dim LMCDeduction As String ' updated Feb. 18, 2012
Dim HazardPayRate As Double, AllowancePayRate As Double
Dim UnionDuesRate As Double, AllowanceBasis As String

Private Sub cmdCreateTrans_Click(Index As Integer)
'statbar.Panels(1).Text = "Create Transaction"
End Sub

Private Sub cmdCancelTransaction()
Dim res

res = MsgBox("Cancel Entries?", vbQuestion + vbYesNo, PayrollAlert)
If res = vbNo Then
    Exit Sub
End If
If frmBasicRates.Visible = True Then
    frmBasicRates.Visible = False
End If
AddFlag = False
EditFlag = False
Call DisableCmdButtons
If ListViewFlag = True Then
    cmdTransaction(0).Enabled = True
    cmdTransaction(1).Enabled = False
    cmdTransaction(0).Caption = "&Create Transaction"
Else
    cmdTransaction(0).Enabled = True
    cmdTransaction(1).Enabled = True
    cmdTransaction(0).Caption = "&Edit Transaction"
End If
cmdTransaction(1).Caption = "&Delete Transaction"
End Sub

Private Sub chkEcolaHol_Click()
    'Call ComputeForECOLA(2)
    Call AccumulateAmount
End Sub

Private Sub chkEcolaHol_LostFocus()
    'Call ComputeForECOLA(2)
    Call AccumulateAmount
End Sub

Private Sub chkEcolaLegal_Click()
    'Call ComputeForECOLA(2)
    Call AccumulateAmount
End Sub

Private Sub chkEcolaLegal_LostFocus()
    'Call ComputeForECOLA(2)
    Call AccumulateAmount
End Sub

Private Sub chkOtherDed_Click()
' Updated by Bernard aug. 24, 2013.
    If chkOtherDed.VALUE = 1 Then
        FrameDeductions.Visible = True
        'Call FillDeductionList(ListDeductions, rsEmpDeductions)
    Else
        FrameDeductions.Visible = False
    End If
'lblEntries(58).Caption = "SSS Calamity"
'lblEntries(57).Caption = "HDMF Calamity"

End Sub

Private Function FillDeductionList(ByRef mList As ListView, mRs As Recordset)
Dim litem1 As ListItem
    
    InsField = ""
    InsValue = ""
    EdtField = ""
    ' Comment by Bernard.
'    mList.ListItems.Clear
'    Set litem1 = mList.ListItems.Add(, , "Planters Loan: " & Format(mCBAmount, "#,##0.00"))
'        litem1.SubItems(1) = 0
'    InsField = "Vale_IOU "
'    InsValue = mCBAmount
'    EdtField = "Vale_IOU = " & mCBAmount

End Function

'Private Function FillDeductionListOrig(ByRef mList As ListView, mRs As Recordset)
'
'Dim litem1 As ListItem
'Dim LSubItem As ListSubItem
'Dim mTotal As Double
'Dim IncludeFlag As Boolean
'Dim StartCutoffDate, EndCutoffDate, StartDate, LastDate
'Dim mRem
'Dim mDed As String
'
'    InsField = ""
'    InsValue = ""
'    EdtField = ""
'    mRem = "Elpsd."
'    With mRs
'        .Filter = "EmployeeID = " & EmployeeNo
'    End With
'
'    mList.ListItems.Clear
'    While Not (mRs.EOF Or mRs.BOF)
'        IncludeFlag = True
'        If mRs!DeductionStatus = False Then
'            IncludeFlag = False
'            mRem = "InAct."
''        ElseIf mRs!DeductionFlag = True Then ' check if payment was based from no of payments
''            'If mRs!DeductionCounter < mRs!NoOfDeductions Then
''            If mRs!paymentsmade >= mRs!TotalDeduction Then  ' Fully Paid
''                IncludeFlag = False
''                If mRs!DeductionCounter < mRs!NoOfDeductions Then
''                    mRem = "Check." ' entries in no of deductions not tally
''                Else
''                    mRem = "Paid."
''                End If
''            Else
''                IncludeFlag = True
''            End If
'        Else    ' Deductions was based from End Date.
'            StartCutoffDate = Val(Format(frmSelectCutoff.txtCutOffDate(0).Text, "yyyymmdd"))
'            EndCutoffDate = Val(Format(frmSelectCutoff.txtCutOffDate(1).Text, "yyyymmdd"))
'            StartDate = Val(Format(mRs!DedStartpayment, "yyyymmdd"))
'            LastDate = Val(Format(mRs!DedEndPayment, "yyyymmdd"))
'            If StartDate = 19000101 Then StartDate = 99999999
'            If LastDate = 19000101 Then LastDate = 99999999
'            ' Validate whether entry dates fall on cut-off dates.
'            If (StartCutoffDate >= StartDate And StartCutoffDate <= LastDate) Or _
'            (EndCutoffDate >= StartDate And EndCutoffDate <= LastDate) Then
'                IncludeFlag = True
'            Else
'                IncludeFlag = False
'                mRem = "Elpsd."
'            End If
'        End If
'
'        If IncludeFlag = True Then
'            'Set litem1 = mList.ListItems.Add(, , IIf(IsNull(mRs.Fields(1)), "New Ded", mRs.Fields(1)) & Space(25 - Len(LTrim(RTrim(mRs.Fields(1))))) & Format(mRs.Fields(8), "##,#00.00"))
'            Set litem1 = mList.ListItems.Add(, , IIf(IsNull(mRs.Fields("DeductionDescription")), "New Ded", mRs.Fields("DeductionDescription")) & Space(25 - Len(LTrim(RTrim(mRs.Fields("DeductionDescription"))))) & Format(mRs.Fields("DeductionPayment"), "##,#00.00"))
'                litem1.SubItems(1) = mRs("EmployeeID")  ' mRs(0)
'            mTotal = mTotal + mRs.Fields("DeductionPayment")
'            InsField = IIf(InsField = "", " ", InsField & ", ") & mRs.Fields("DeductionFieldName") ' 24
'            InsValue = IIf(InsValue = "", " ", InsValue & ", ") & mRs.Fields("DeductionPayment")  ' 8
'            EdtField = IIf(EdtField = "", " ", EdtField & ", ") & mRs.Fields("DeductionFieldName") & " = " & mRs.Fields("DeductionPayment") ' 22
'        Else
'            ' Updated by Bernard July 2009.
'            If IsNull(mRs.Fields("DeductionDescription")) = False Then
'                mDed = mRs.Fields("DeductionDescription")
'            Else
'                mDed = "New Ded"
'            End If
'            'Set litem1 = mList.ListItems.Add(, , IIf(IsNull(mRs.Fields("DeductionDescription")), "New Ded", mRs.Fields("DeductionDescription")) & Space(25 - Len(LTrim(RTrim(mRs.Fields("DeductionDescription"))))) & mRem)
'            Set litem1 = mList.ListItems.Add(, , mDed & Space(25 - Len(mDed)) & mRem)
'                litem1.SubItems(1) = mRs("EmployeeID")
'        End If
'        mRs.MoveNext
'    Wend
'    Set litem1 = mList.ListItems.Add(, , "Total: " & Space(23) & Format(mTotal, "##,#00.00"))
'    mTotalDeductions = mTotal
'
'End Function

Private Sub cmdClose_Click()
Set rsCompanyInfo = Nothing
Set rsClientInfo = Nothing
Set rsEmployeeInfo = Nothing
Set rsProcessEmployee = Nothing

Call ProtectDatabaseRecords(CompanyCode, mClientCode, False, True)  ' Unprotect the Employee deduction table for entries

Unload Me
End Sub

Private Sub cmdDTR_Click()
Dim res
Dim rsDTR As ADODB.Recordset
Dim mDTRKEY As String
Dim CheckCutOff As Date, mYr As String, mMon As String

res = MsgBox("Read DTR entries of " & lblEmpName, vbQuestion + vbYesNo, PayrollAlert)
If res = vbYes Then
    If mCutOffNo = 1 Then
'        ' Updated for new Payroll Cut-off.
'        CheckCutOff = Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy")
'        If month(CheckCutOff) = 1 Then
'            mMon = "12"
'            mYr = Trim$(str(Val(Year(CheckCutOff) - 1)))
'        Else
'            mMon = Right("0" + Trim$(str(Val(month(CheckCutOff) + 1))), 2)
'            mYr = Trim$(str(Year(CheckCutOff)))
'        End If
'        mDTRKEY = mYr + mMon + "01"
        mDTRKEY = Format(frmSelectCutoff.txtCutOffDate(1).Text, "yyyy") & Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm") + "01"
    Else
        mDTRKEY = Format(frmSelectCutoff.txtCutOffDate(1).Text, "yyyy") & Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm") + "02"
    End If
    
    If FMSFlag = False Then
        StrQuery = "SELECT TOP 1 * FROM vwDTRFinalPayroll WHERE EmpNo = '" & EmpNo & "'" & _
            " AND DTRKEY = '" & mDTRKEY & "'"
    Else
        StrQuery = "SELECT TOP 1 * FROM vwDTRFinalConductor WHERE EmpNo = '" & EmpNo & "'" & _
            " AND DTRKEY = '" & mDTRKEY & "'"
    End If
    Set rsDTR = ReturnRst(StrQuery, cnnstr, mstatic)
    If rsDTR.RecordCount = 0 Then
        MsgBox "No DTR Records processed from Time Keeping for " & lblEmpName & Chr(13) & _
            " for Cut-Off " & frmSelectCutoff.txtCutOffDate(0).Text & " to " & frmSelectCutoff.txtCutOffDate(1).Text, vbInformation, PayrollAlert
    Else
        Call ProcessDtrEntries(rsDTR, mDTRKEY)
        MsgBox "DTR entries was updated!", vbInformation, PayrollAlert
    End If
End If
Set rsDTR = Nothing
txtDaysWork.Enabled = True
txtDaysWork.SetFocus
Call AutoClickText
'MsgBox "Payroll was automatically computed!", vbInformation, PayrollAlert

End Sub

Private Sub ProcessDtrEntries(ByRef rsDTR As ADODB.Recordset, ByVal mDTRKEY As String)
Dim rsFiled As ADODB.Recordset
Dim mSLHours As Double, mVLHours As Double
Dim mDTRID As String

With rsDTR
    mDTRID = !DTRID
'mSLHours = GetLeaveHours(!DTRID, mDTRKEY, "S")
'mVLHours = GetLeaveHours(!DTRID, mDTRKEY, "L")
    txtDaysWork.Text = IIf(IsNull(!WorkHours), 0, !WorkHours)  ' - mSLHours - mVLHours)
    txtRegOT.Text = IIf(IsNull(!OT), 0, !OT)
'    txtSundayOT.Text = IIf(IsNull(!RestOT), 0, !RestOT)
    txtDaysAbsent.Text = IIf(IsNull(!Absent), 0, !Absent)
    txtHoursLate.Text = IIf(IsNull(!Lates), 0, !Lates) + IIf(IsNull(!Undertime), 0, !Undertime)
    txtNightDiff.Text = IIf(IsNull(!NDiff), 0, !NDiff)
    'txtSickLeave.Text = mSLHours
    'txtVacLeave.Text = mVLHours
    If PaymentBasis = "Monthly" Then
        txtDaysWork.Text = "104"
    End If
End With
    
StrQuery = "SELECT * FROM vwFiledLeave WHERE DTRID = '" & mDTRID & "'" & _
        " AND DTRKEY = '" & mDTRKEY & "'"
Set rsFiled = ReturnRst(StrQuery, cnnstr, mstatic)

If rsFiled.RecordCount > 0 Then
    With rsFiled
        .MoveFirst
        Do While Not .EOF
            Select Case Trim$(!leavetype)
                Case "H"    ' Spl Holiday reverse variable name
                    txtLglHoliday.Text = IIf(IsNull(!LeaveHours), 0, !LeaveHours)
                Case "L"    ' Legal Holiday
                    txtSplHoliday.Text = IIf(IsNull(!LeaveHours), 0, !LeaveHours)
                Case "M", "S"    ' Sickness
                    txtSickLeave.Text = IIf(IsNull(!LeaveHours), 0, !LeaveHours)
                Case "P", "V", "Y"    ' Vacation leave
                    txtVacLeave.Text = IIf(IsNull(!LeaveHours), 0, !LeaveHours)
                Case "R"    ' Rest day ot
                    txtSundayOT.Text = IIf(IsNull(!LeaveHours), 0, !LeaveHours)
                    txtDaysWork.Text = Val(txtDaysWork.Text) - !LeaveHours
           End Select
            .MoveNext
        Loop
    End With
End If

Set rsFiled = Nothing
If FMSFlag = False Then
    StrQuery = "SELECT COUNT(reqworkhours) ReqWork FROM DTRPayrollRecords WHERE DTRID = '" & mDTRID & "'" & _
        " AND DTRKEY = '" & mDTRKEY & "' AND ReqWorkHours <> 0.00"
Else
    StrQuery = "SELECT COUNT(reqworkhours) ReqWork FROM DTRPayrollConductors WHERE DTRID = '" & mDTRID & "'" & _
        " AND DTRKEY = '" & mDTRKEY & "' AND ReqWorkHours <> 0.00"
End If
Set rsFiled = ReturnRst(StrQuery, cnnstr, mstatic)

If rsFiled.RecordCount > 0 Then
    txtNoDays = rsFiled!ReqWork
Else
    txtNoDays = (Val(txtNoDays.Text) / 8)
End If

Set rsFiled = Nothing

End Sub

'Private Function GetLeaveHours(ByVal mDTRID As String, _
'    ByVal mDTRKEY As String, ByVal mType As String) As Double
'Dim rsLeave As ADODB.Recordset
'
''If mType = "S" Then ' get sick leave
''    StrQuery = "SELECT TOP 1 * FROM vwFiledSickLeave WHERE DTRID = '" & mDTRID & "'" & _
''        " AND DTRKEY = '" & mDTRKEY & "'"
''Else                ' get vacation leave
''    StrQuery = "SELECT TOP 1 * FROM vwFiledVacationLeave WHERE DTRID = '" & mDTRID & "'" & _
''        " AND DTRKEY = '" & mDTRKEY & "'"
''End If
'StrQuery = "SELECT TOP 1 * FROM vwFiledVacationLeave WHERE DTRID = '" & mDTRID & "'" & _
'    " AND DTRKEY = '" & mDTRKEY & "' AND LeaveType = '" & mType & "'"
'
'
'Set rsLeave = ReturnRst(StrQuery, cnnstr, mstatic)
'If rsLeave.RecordCount <> 0 Then
'    GetLeaveHours = Val(rsLeave!LeaveHours)
'Else
'    GetLeaveHours = 0
'End If
'Set rsLeave = Nothing
'
'End Function

Private Sub AutoClickText()
txtRegOT.SetFocus
txtNightDiff.SetFocus
'txtECOLA.SetFocus
'txtSundayOT.SetFocus
'txtSplHoliday.SetFocus
'txtLglHoliday.SetFocus
'txtSickLeave.SetFocus
'txtVacLeave.SetFocus
'Updated April 1, 2011
If PaymentBasis = "Monthly" Then
    AddLeaveFlag = True
    OrigMonthlyRate = TxtAmount(0).Text
    TxtAmount(0).Text = Format(TxtAmount(0).Text - (Val(TxtAmount(7).Text) + Val(TxtAmount(8).Text)), "###,#00.00")
Else
    AddLeaveFlag = False
End If
txtSalAdjustment.SetFocus
txtIncentive.SetFocus
'txtDaysAbsent.SetFocus
'txtHoursLate.SetFocus
txtTASAesla.SetFocus
txtVale.SetFocus
txtMiscFee.SetFocus
txtBDOCard.SetFocus
txtPlacementFee.SetFocus
txtOverPay.SetFocus

If WithHazardPay = True And EditFlag = False Then
    If txtHazardPay.Visible = True Then
        txtHazardPay.SetFocus   'for hazard pay
    End If
    If txtUnionDues.Visible = True Then
        txtUnionDues.SetFocus   'for union dues
    End If
    'txtHazardPay.Text = Format(HazardPayRate * (Val(txtDaysWork.Text) / 8), "##0.00")
    'txtHazardPay.Text = Format(HazardPayRate * Val(txtNoDays.Text), "##0.00") ' place in Ecola computation
    txtUnionDues.Text = Format(UnionDuesRate, "##0.00")
End If

'' Auto-compute allowance if rate is non-zero. Updated July 17, 2016.
'If AllowancePayRate > 0# Then
'    txtIncentive.Text = Format(AllowancePayRate * Val(txtNoDays.Text), "##0.00")
'End If

'Auto-compute Allowance if rate is non-zero.
If AllowanceBasis = "Per Day" Then
    txtIncentive.Text = Format(AllowancePayRate * Val(txtNoDays.Text), "##0.00")
ElseIf AllowanceBasis = "All Cutoff" Then
    txtIncentive.Text = Format(AllowancePayRate, "##,##0.00")
ElseIf AllowanceBasis = "1st Cutoff" And mCutOffNo = 1 Then
    txtIncentive.Text = Format(AllowancePayRate, "##,##0.00")
ElseIf AllowanceBasis = "2nd Cutoff" And mCutOffNo = 2 Then
    txtIncentive.Text = Format(AllowancePayRate, "##,##0.00")
End If

cmdTransaction(0).SetFocus
End Sub

Private Sub cmdEditTax_Click(Index As Integer)
frmBasicRates.TxtDRate.Text = Format(DailyRate, "##0.00")
frmBasicRates.Show
End Sub

'Private Sub cmdEditTax_Click(Index As Integer)
'frmBasicRates.Show
'
''If EditTaxFlag = True Then
'    ' Enable tax text box for editing
''    AgencySw = True
''    txtCalamityLoan.Enabled = True
''    txtCalamityLoan.Locked = False
''    If txtCalamityLoan.Enabled = True Then
''        txtCalamityLoan.SetFocus
''    End If
''Else
''    txtCalamityLoan.Enabled = False
''    txtCalamityLoan.Locked = True
''End If
'End Sub

'Private Sub cmdExcessOt_Click(Index As Integer)
'    rsClientTable.Requery
'    Call GetClientInformation
'    RetSw = True
'    retClientSw = True
'    'Call GetBasicInformation ' comment by bernard Dec. 6, 2016 to lessen the task
'
'    'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ClientCode = " & mClientCode & " and ProtectedFlag = 0"
'    StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
'    Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
'    Call FillDeductionList(ListDeductions, rsEmpDeductions)
'    Call GetEmployeeBasic
'
'    Call ReComputePayroll ' Bernard Updates 8/7/2010
'    Call SaveTransaction
'
'    'MsgBox "Data was updated!", vbInformation, PayrollAlert
'End Sub

Private Sub cmdFind_Click()
    chkClient.VALUE = 1
    frmClientTable.Show
    frmClientTable.SetFocus
End Sub

Private Sub cmdFixedRates_Click()
    empBasicSw = True
    frmBasicSalary.Show
    frmBasicSalary.SetFocus
    RetSw = True
End Sub

Private Sub cmdmovetoMaster_Click()
Dim res

' Move transaction to Master.
res = MsgBox("This will move the Transaction Records to Master and close the Payroll Cutoff. Do you want to proceed?", vbQuestion + vbYesNo, PayrollAlert)
If res = vbYes Then
    Call MoveToMaster
    Call ProtectDatabaseRecords(CompanyCode, mClientCode, False, True)  ' Unprotect the Employee deduction table for entries
    Call AdjustDeductionTable
End If
End Sub

Private Sub MoveToMaster()
Dim StartDate
Dim EndDate
Dim mFortheMonth As Integer
Dim mTotalEmp As Integer

On Error GoTo Errh
'StartDate = Format(frmSelectCutoff.txtCutOffDate(0).Text, "yyyymmdd") ' Original code
StartDate = Format(Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy"), "yyyymmdd")

'EndDate = Format(frmSelectCutoff.txtCutOffDate(1).Text, "yyyymmdd") ' Original code
EndDate = Format(Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy"), "yyyymmdd")
mFortheMonth = Format(frmSelectCutoff.txtCutOffDate(0).Text, "m")
mTotalEmp = rsPayrollTrans.RecordCount

StrQuery = "INSERT INTO PayrollMaster SELECT * FROM PayrollTrans WHERE clientcode = " & mClientCode
Call Action_Query(StrQuery, cnnstr)
StrQuery = "INSERT INTO MasterAdjust SELECT * FROM PayrollAdjust WHERE clientcode2 = " & mClientCode
Call Action_Query(StrQuery, cnnstr)

'Call UpdateEmployeeDeductions  ' update in deductions was done during payroll entries.

''Updated July 2009 for Billing purposes.
'StrQuery = "INSERT INTO BillingCutoffTable SELECT * FROM qryPayrollTransCutoff WHERE clientCode = " & mClientCode & " and StartCutoffText = '" & _
'    Format(frmSelectCutoff.txtCutOffDate(0).Text, "yyyymmdd") & "' and EndCutoffText = '" & _
'    Format(frmSelectCutoff.txtCutOffDate(1).Text, "yyyymmdd") & "'"
'Call Action_Query(StrQuery, cnnstr)

'Updated for Automated Deduction recomputation. December 7, 2013.
'StrQuery = "EXEC UpdateDeductionEntries " & mClientCode & ", '" & frmSelectCutoff.txtCutOffDate(0).Text & "', '" & frmSelectCutoff.txtCutOffDate(1).Text & "'"
'Call Action_Query(StrQuery, cnnstr)

StrQuery = "DELETE FROM PayrollTrans WHERE clientcode = " & mClientCode
Call Action_Query(StrQuery, cnnstr)
StrQuery = "DELETE FROM PayrollAdjust WHERE clientcode2 = " & mClientCode
Call Action_Query(StrQuery, cnnstr)

'Update Cutoff Table.
StrQuery = "UPDATE CutoffTable SET BilledFlag = 1, ViewFlag = 1, TotalEmployee = " & mTotalEmp & ", ClientSeriesNo = '" & frmSelectCutoff.txtSeriesNo & "', Processedby = '" & mUserID & "' WHERE " & _
        "AgencyCode = " & CompanyCode & " and ClientCode = " & mClientCode & " and StartCutOffText = '" & StartDate & "' and EndCutOffText = '" & EndDate & "'"

'strQuery = "UPDATE cutofftable SET BilledFlag = true, ViewFlag = true, TotalEmployee = " & mTotalEmp & ", ClientSeriesNo = '" & frmSelectCutoff.txtSeriesNo & "', Processedby = '" & mUserID & "' WHERE " & _
'        "AgencyCode = " & mAgencyCode & " and ClientCode = " & mClientCode & " and StartCutOffText = '" & StartDate & "' and EndCutOffText = '" & EndDate & "'"

Call Action_Query(StrQuery, cnnstr)
' Set the view flag to false, if payroll cutoff is 1 year older. Original code.
'StartDate = Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy")
'EndDate = Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy")

StartDate = Format(Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy"), "mm/dd/yyyy")
EndDate = Format(Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy"), "mm/dd/yyyy")


'StrQuery = "UPDATE cutofftable SET ViewFlag = 0 WHERE BilledFlag = 1 and " & _
'        "AgencyCode = " & mAgencyCode & " and ClientCode = " & mClientCode & _
'        " and FortheMonth = " & mFortheMonth & " and FortheYear = " & Format(EndDate, "yyyy") - 1

' Updated 2/20/2017, use the EndCutoffDate.
StrQuery = "UPDATE cutofftable SET ViewFlag = 0 WHERE BilledFlag = 1 and " & _
        "AgencyCode = " & mAgencyCode & " and ClientCode = " & mClientCode & _
        " and MONTH(EndCutoffDate) = " & Format(EndDate, "m") & " and YEAR(EndCutoffDate) = " & Format(EndDate, "yyyy") - 1

Call Action_Query(StrQuery, cnnstr)

'' Updated 2/20/2017, use the EndCutoffDate.
'StrQuery = "UPDATE cutofftable SET ViewFlag = 1 WHERE BilledFlag = 1 and " & _
'        "AgencyCode = " & mAgencyCode & " and ClientCode = " & mClientCode & _
'        " and YEAR(EndCutoffDate) >= " & Format(now, "yyyy") - 1
'
'Call Action_Query(StrQuery, cnnstr)

'Update Basic Salary Table.
StrQuery = "UPDATE employeebasicsalary SET processedflag = 0 WHERE clientcode = " & mClientCode
Call Action_Query(StrQuery, cnnstr)
'Update client series number.
StrQuery = "UPDATE clienttable SET clientseriesno = " & ClientSeriesNo & " WHERE clientcode = " & mClientCode
Call Action_Query(StrQuery, cnnstr)

MsgBox "Transaction was successfully moved to Master!!!", vbInformation, PayrollAlert

InsertedFlag = False    ' set inserted to false, to make it available for new cut-off
Unload Me
frmSelectCutoff.Show
Exit Sub

Errh:

    MsgBox "Error: Cannot move transaction to Master Table!", vbInformation, PayrollAlert
    Err.Clear
    Resume Next
    
End Sub

Private Sub AdjustDeductionTable()
Dim rsLoanAmt As Recordset
Dim rsPayment As Recordset
Dim mDeduction As Double, mPayment As Double, RunBal As Double, RunBal2 As Double, DedPay As Double
Dim StartDate, DedStart
Dim EndDate, mmm As Integer, ddd As Integer, yyy As Integer
Dim mFortheMonth As Integer
Dim mTotalEmp As Integer
Dim mRem As String

On Error GoTo ExitRtn2

RunBal = 0: RunBal2 = 0: DedPay = 0
StartDate = Format(frmSelectCutoff.txtCutOffDate(0).Text, "yyyymmdd")
EndDate = Format(frmSelectCutoff.txtCutOffDate(1).Text, "yyyymmdd")

' Get those payments not equal to schedule.
StrQuery = "SELECT * FROM EmployeeDedSchedules " & _
    " WHERE DeductionStatus = 1 AND PaymentsMade > 0.00 AND PaymentsMade > DeductionPayment" & _
    " AND Remarks like 'Paid%'"
Set rsLoanAmt = ReturnRst(StrQuery, cnnstr, mstatic)
If rsLoanAmt.RecordCount > 0 Then
    Do While Not rsLoanAmt.EOF
        mDeduction = rsLoanAmt!DeductionPayment
        mPayment = rsLoanAmt!PaymentsMade
        ' Get last deduction schedule.
        StrQuery = "SELECT TOP 1 * FROM EmployeeDedSchedules " & _
            " WHERE DeductionStatus = 1 AND DeductionCode = " & rsLoanAmt!DeductionCode & " AND EmployeeID = " & rsLoanAmt!EmployeeID & _
            " AND DeductionPayment = " & mDeduction & " ORDER BY DedStartPayment DESC "
        Set rsPayment = ReturnRst(StrQuery, cnnstr, mstatic)
        If rsPayment.RecordCount > 0 Then
            With rsPayment
'            If mDeduction > mPayment Then  ' understated payment, Add new schedule
'                'Insert the new schedules.
'                mRem = "'Add. " & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'"
'                If Day(!DedStartPayment) > 1 Then ' every 16th
'                    mmm = month(!DedStartPayment) + 1
'                    If mmm > 12 Then
'                        mmm = 1: yyy = Year(!DedStartPayment) + 1
'                    Else
'                        yyy = Year(!DedStartPayment)
'                    End If
'                    DedStart = Right("00" & Trim$(mmm), 2) & "/01/" & Trim$(yyy)
'                Else
'                    mmm = month(!DedStartPayment) + 1
'                    If mmm > 12 Then
'                        mmm = 1: yyy = Year(!DedStartPayment) + 1
'                    Else
'                        yyy = Year(!DedStartPayment)
'                    End If
'                    DedStart = Right("00" & Trim$(mmm), 2) & "/16/" & Trim$(yyy)
'                End If
'                StrQuery = "INSERT INTO EmployeeDedSchedules (Remarks, EmployeeName, EmployeeID, EmpNo, DeductionDescription, AgencyCode, ClientCode, DeductionStatus, DeductionCode, DeductionPayment, RunningBalance, DeductionAmount, VoucherNo, DedStartPayment, DedEndPayment, TotalInterest, TotalDeduction, NoOfDeductions, DeductionCounter, DeductionFlag, EntryBy, EntryDate, DeductionID) " & _
'                    " VALUES ( " & mRem & ", '" & !EmployeeName & "', " & !EmployeeID & ", '" & !EmpNo & "', '" & !DeductionDescription & "', " & !AgencyCode & ", " & !ClientCode & ", " & IIf(!DeductionStatus = True, 1, 0) & ", " & !DeductionCode & ", " & (mDeduction - mPayment) & ", " & (mDeduction - mPayment) & ", " & (mDeduction - mPayment) & ", '" & !VoucherNo & _
'                    "', '" & DedStart & "', '" & DedStart & _
'                    "', " & !TotalInterest & ", " & !TotalDeduction & ", " & !NoOfDeductions & ", " & (!DeductionCounter + 1) & ", " & IIf(!DeductionFlag = True, 1, 0) & ", '" & !EntryBy & "', '" & Format(Now, "mm/dd/yyyy") & "', " & !DeductionID & ")"
'            Else
                ' Check if Running Balance is greater than excess payment.
                If !DeductionPayment > (mPayment - mDeduction) Then
                    DedPay = !DeductionPayment - (mPayment - mDeduction)
                    RunBal = IIf(IsNull(!PaymentsMade), 0, !PaymentsMade) + (mPayment - mDeduction)
                Else
                    DedPay = 0#
                    RunBal = IIf(IsNull(!PaymentsMade), 0, !PaymentsMade) + (mPayment - mDeduction)
                End If
                mRem = "'Ded. (" & RunBal & ") " & Format(Now, "mm/dd/yyyy") & "'"
                StrQuery = "UPDATE EmployeeDedSchedules SET DeductionPayment = " & DedPay & ", PaymentsMade = " & RunBal & ", PaymentDate = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "', " & _
                    " Remarks = " & mRem & _
                    " WHERE DedStartpayment = '" & !DedStartpayment & "' AND " & _
                    " DedEndPayment = '" & !DedEndPayment & "'" & _
                    " AND DeductionStatus = 1 AND DeductionCode = " & !DeductionCode & " AND EmployeeID = " & !EmployeeID
'            End If
            If Action_Query(StrQuery, cnnstr) = True Then
                ' Update the remarks to signifies adjustments.
                mRem = "Replace(Remarks, 'Paid', 'Adjt')"
                StrQuery = "UPDATE EmployeeDedSchedules SET Remarks = " & mRem & _
                    " WHERE DedStartpayment = '" & rsLoanAmt!DedStartpayment & "' AND " & _
                    " DedEndPayment = '" & rsLoanAmt!DedEndPayment & "'" & _
                    " AND DeductionStatus = 1 AND DeductionCode = " & rsLoanAmt!DeductionCode & " AND EmployeeID = " & rsLoanAmt!EmployeeID
                Call Action_Query(StrQuery, cnnstr)
            End If
            End With
        End If
        Set rsPayment = Nothing
        rsLoanAmt.MoveNext
    Loop
End If

'' Check if actual payment is not equal to payroll, then subject for adjustment in Ded table.
'' Updated January 20, 2012.
'If mActualAmt <> arrLoanPayments(o) And arrLoanPayments(o) <> arrLoanAmount(o) Then
'    StrQuery = "SELECT TOP 1 * FROM EmployeeDedSchedules " & _
'        " WHERE DeductionStatus = 1 AND DeductionCode = " & o & " AND EmployeeID = " & EmployeeNo & _
'        " ORDER BY dedstartpayment DESC "
'    Set rsLoanAmt = ReturnRst(StrQuery, cnnstr, mstatic)
'    If rsLoanAmt.RecordCount > 0 Then
'        With rsLoanAmt
'        If mActualAmt > arrLoanPayments(o) Then  ' understated payment, Add new schedule
'            'Insert the new schedules.
'            mRem = "'Adj. " & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'"
'            StrQuery = "INSERT INTO EmployeeDedSchedules (Remarks, EmployeeName, EmployeeID, EmpNo, DeductionDescription, AgencyCode, ClientCode, DeductionStatus, DeductionCode, DeductionPayment, RunningBalance, DeductionAmount, VoucherNo, DedStartPayment, DedEndPayment, TotalInterest, TotalDeduction, NoOfDeductions, DeductionCounter, DeductionFlag, EntryBy, EntryDate, DeductionID) " & _
'                " VALUES ( " & mRem & ", '" & !EmployeeName & "', " & !EmployeeID & ", '" & !EmpNo & "', '" & !DeductionDescription & "', " & !Agencycode & ", " & !ClientCode & ", " & IIf(!DeductionStatus = True, 1, 0) & ", " & !DeductionCode & ", " & (mActualAmt - arrLoanPayments(o)) & ", " & (mActualAmt - arrLoanPayments(o)) & ", " & (mActualAmt - arrLoanPayments(o)) & ", '" & !VoucherNo & _
'                "', '" & !DedStartpayment & "', '" & !DedEndPayment & _
'                "', " & !TotalInterest & ", " & !TotalDeduction & ", " & !NoOfDeductions & ", " & !DeductionCounter & ", " & IIf(!DeductionFlag = True, 1, 0) & ", '" & !EntryBy & "', '" & Format(Now, "mm/dd/yyyy") & "', " & !DeductionID & ")"
'            Call Action_Query(StrQuery, CnnstrHR)
'
'        Else
'
'        End If
'        End With
'    End If
'    Set rsLoanAmt = Nothing
'End If

Exit Sub

ExitRtn2:

    MsgBox Err.Description, vbInformation, PayrollAlert
    Err.Clear

End Sub

' Update the Employee Deduction table for payments and count. Updated Sept. 6, 2008.
Private Sub UpdateEmployeeDeductionTable(ByVal mEmpno As Integer, ByVal ForPayment As Boolean)

Dim rsEmpDed As ADODB.Recordset
Dim mPayment As String
Dim mPayment2 As Double

rsPayrollTrans.Filter = ""  ' clear the filter of payroll transactions
'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ClientCode = " & mClientCode
StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE EmployeeID = " & mEmpno & " and ProtectedFlag = 0 "  ' ProtectedFlag are used to preserved old deduction entries that was updated
Set rsEmpDeductions = ReturnRst(StrQuery, cnnstr, mstatic)
If rsEmpDeductions.RecordCount <> 0 Then
    With rsEmpDeductions
    .MoveFirst
    Do While Not .EOF
        mPayment2 = 0#
        If ForPayment = False Then   ' Deduc deduction payments
            'Comment by Bernard Feb. 24, 2014
'            StrQuery = "UPDATE EmployeeDeductions SET DeductionCounter = DeductionCounter - 1, PaymentsMade = PaymentsMade - LastDeduction, LastDeduction = 0 " & _
'                " WHERE DeductionCode = " & !DeductionCode & " and EmployeeID = " & !EmployeeID
'                '" WHERE DeductionCode = " & !DeductionCode & " and EmployeeID = " & !EmployeeID & " and EmpNo = '" & !EmpNo & "' and AgencyCode = " & !agencycode & " and ClientCode = " & !ClientCode
'            Call Action_Query(StrQuery, CnnstrHR)
        Else
            rsPayrollTrans.Filter = ""  ' clear the filter of payroll transactions
            rsPayrollTrans.Filter = "EmployeeID = " & !EmployeeID & " and EmpNo = '" & !EmpNo & "' and AgencyCode = " & !agencycode & " and ClientCode = " & !ClientCode
            If rsPayrollTrans.EOF = False Then
                If UCase(Trim(!DeductionFieldName)) = "MEDICAL" Then
                    mPayment2 = rsPayrollTrans!Medical
                ElseIf UCase(Trim(!DeductionFieldName)) = "SALARYLOAN" Then
                    mPayment2 = rsPayrollTrans!SalaryLoan
                ElseIf UCase(Trim(!DeductionFieldName)) = "UNIFORM" Then
                    mPayment2 = rsPayrollTrans!Uniform
                ElseIf UCase(Trim(!DeductionFieldName)) = "VARIANCES" Then
                    mPayment2 = rsPayrollTrans!Variances
                ElseIf UCase(Trim(!DeductionFieldName)) = "VALE_IOU" Then
                    mPayment2 = rsPayrollTrans!Vale_IOU
                Else
                    mPayment2 = 0
                End If
                'Comment by Bernard Feb. 24, 2014.
'                If mPayment2 <> 0 Then
'                    If AddFlag = False Then ' the transaction record was edited, we need to delete the previous updates in deduction table
'                        StrQuery = "UPDATE EmployeeDeductions SET DeductionCounter = DeductionCounter - 1, PaymentsMade = PaymentsMade - LastDeduction, LastDeduction = 0 " & _
'                            " WHERE DeductionCode = " & !DeductionCode & " and EmployeeID = " & !EmployeeID
'                            '" WHERE DeductionCode = " & !DeductionCode & " and EmployeeID = " & !EmployeeID & " and EmpNo = '" & !EmpNo & "' and AgencyCode = " & !agencycode & " and ClientCode = " & !ClientCode
'                        Call Action_Query(StrQuery, CnnstrHR)
'                    End If
'                    StrQuery = "UPDATE EmployeeDeductions SET DeductionCounter = DeductionCounter + 1, LastDeduction = " & mPayment2 & ", PaymentsMade = PaymentsMade + " & mPayment2 & _
'                        " WHERE DeductionCode = " & !DeductionCode & " and EmployeeID = " & !EmployeeID
'                        '" WHERE DeductionCode = " & !DeductionCode & " and EmployeeID = " & !EmployeeID & " and EmpNo = '" & !EmpNo & "' and AgencyCode = " & !agencycode & " and ClientCode = " & !ClientCode
'                    Call Action_Query(StrQuery, CnnstrHR)
'                End If
            End If
        End If
        .MoveNext
    Loop
    End With
End If
rsPayrollTrans.Filter = ""  ' clear the filter of payroll transactions
'Requery the Deduction table for updates.
'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ClientCode = " & mClientCode & " and ProtectedFlag = 0"
StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
End Sub

'Private Sub UpdateEmployeeDeductions()
'
'Dim mPayment As String
'Dim mPayment2 As Double
'
'rsPayrollTrans.Filter = ""  ' clear the filter of payroll transactions
''strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ClientCode = " & mClientCode & " and ProtectedFlag = 0"
'StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
'
'Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
'If rsEmpDeductions.RecordCount <> 0 Then
'    With rsEmpDeductions
'    .MoveFirst
'    Do While Not .EOF
'        mPayment2 = 0#
'        rsPayrollTrans.Filter = ""  ' clear the filter of payroll transactions
'        'rsPayrollTrans.Filter = "EmployeeID = " & !EmployeeID & " and EmpNo = '" & !EmpNo & "' and " & Trim(!DeductionFieldName) & " <> 0.00"
'        rsPayrollTrans.Filter = "EmployeeID = " & !EmployeeID & " and EmpNo = '" & !EmpNo & "' and AgencyCode = " & !agencycode & " and ClientCode = " & !ClientCode
'        If rsPayrollTrans.EOF = False Then
'            mPayment = "rsPayrollTrans" & "!" & Trim(!DeductionFieldName)
'            If UCase(Trim(!DeductionFieldName)) = "MEDICAL" Then
'                mPayment2 = rsPayrollTrans!Medical
'            ElseIf UCase(Trim(!DeductionFieldName)) = "SALARYLOAN" Then
'                mPayment2 = rsPayrollTrans!SalaryLoan
'            ElseIf UCase(Trim(!DeductionFieldName)) = "UNIFORM" Then
'                mPayment2 = rsPayrollTrans!Uniform
'            ElseIf UCase(Trim(!DeductionFieldName)) = "VARIANCES" Then
'                mPayment2 = rsPayrollTrans!Variances
'            ElseIf UCase(Trim(!DeductionFieldName)) = "VALE_IOU" Then
'                mPayment2 = rsPayrollTrans!Vale_IOU
'            Else
'                mPayment2 = 0
'            End If
'           'mPayment2 = mPayment
'           If mPayment2 <> 0 Then
'                'strQuery = "UPDATE EmployeeDeductions SET DeductionCounter = DeductionCounter + 1, PaymentsMade = PaymentsMade + " & mPayment2 & " WHERE DeductionCode = " & !DeductionCode & " and EmployeeID = " & !EmployeeID & " and EmpNo = '" & !EmpNo & "' and AgencyCode = " & !AgencyCode & " and ClientCode = " & !ClientCode
'                StrQuery = "UPDATE EmployeeDeductions SET DeductionCounter = DeductionCounter + 1, PaymentsMade = PaymentsMade + " & mPayment2 & " WHERE DeductionCode = " & !DeductionCode & " and EmployeeID = " & !EmployeeID
'                Call Action_Query(StrQuery, CnnstrHR)
'            End If
'        End If
'        .MoveNext
'    Loop
'    End With
'End If
'rsPayrollTrans.Filter = ""  ' clear the filter of payroll transactions
'End Sub

Private Sub cmdNavigation_Click(Index As Integer)
'Call Navigation(rsPayrollTrans, Me.cmdNavigation(Index), _
'    Me.cmdNavigation(0), Me.cmdNavigation(1), Me.cmdNavigation(2), Me.cmdNavigation(3))
Call DisplayAmount
Call ScrollRec
End Sub

Private Sub ScrollRec()
Dim strFindMe As String
Dim itmFound As ListItem
   
    strFindMe = EmployeeNo
    Set itmFound = lvwProcessEmp.FindItem(strFindMe, 1, , lvwPartial)
    If itmFound Is Nothing Then
    Else
        itmFound.EnsureVisible
        itmFound.SELECTED = True
    End If
End Sub

Private Sub cmdPrintReports_Click()
frmPrintPayslip.Show 'vbModal
End Sub

Private Sub cmdSaveTransaction()
Call SaveTransaction
AddFlag = False
EditFlag = False
cmdEditTax(3).Enabled = False
Call DisableCmdButtons
If ListViewFlag = True Then
    cmdTransaction(0).Enabled = True
    cmdTransaction(1).Enabled = False
    cmdTransaction(0).Caption = "&Create Transaction"
Else
    cmdTransaction(0).Enabled = True
    cmdTransaction(1).Enabled = True
    cmdTransaction(0).Caption = "&Edit Transaction"
End If
End Sub

Private Sub cmdRefreshRecords_Click()
    rsClientTable.Requery
    Call GetClientInformation
    RetSw = True
    retClientSw = True
    
    StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"
    Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
    Call FillDeductionList(ListDeductions, rsEmpDeductions)
    Call GetEmployeeBasic
    
    Call ReComputePayroll ' Bernard Updates 8/7/2010
    Call SaveTransaction
End Sub

Private Sub cmdTransaction_Click(Index As Integer)
Dim res, tt

cmdDTR.Caption = "DTR"
cmdDTR.Enabled = False
If lblEmpName = "" Then
    MsgBox "Select employee first!", vbInformation, PayrollAlert
    Exit Sub
End If
If Index = 0 Then ' check whether create, edit or save transaction
    If cmdTransaction(Index).Caption = "&Create Transaction" Then
        ' Create Transaction
        Call ZeroOutAmount
        AddFlag = True
        chkEcolaHol = 0
        chkEcolaLegal = 0
        cmdEditTax(3).Enabled = True
        cmdTransaction(0).Caption = "&Save Transaction"
        cmdTransaction(1).Caption = "&Cancel Transaction"
        Call DisableCmdButtons
        If PaymentBasis = "Monthly" Then
            txtDaysWork.Text = "104" '(NoofDaysperMonth / 2) * 8   ' convert to hours
            TxtAmount(0).Text = Format(MonthlyRate / 2, "##,##0.00")
        ElseIf PaymentBasis = "Daily" Or PaymentBasis = "Hourly" Then
            txtDaysWork.Text = "104" '(NoofDaysperMonth / 2) * 8   ' convert to hours
            TxtAmount(0).Text = Format((DailyRate / 8) * IIf(txtDaysWork.Text = "", 0#, txtDaysWork.Text), "##,##0.00")
        ElseIf PaymentBasis = "Commission" Then ' for commission base employees
            'Call GetCommissionEntries
            cmdDTR.Caption = "FMS"
        End If
        cmdDTR.Enabled = True
        Call AutoClickText  ' Updated July 17, 2016
        
'        ' Get Hazard and Union Dues.
'        If WithHazardPay = True Then
'            'txtHazardPay.Text = Format(HazardPayRate * (Val(txtDaysWork.Text) / 8), "##0.00")
'            txtHazardPay.Text = Format(HazardPayRate * Val(txtNoDays.Text), "##0.00")
'            txtUnionDues.Text = Format(UnionDuesRate, "##0.00")
'        End If
'        ' Auto-compute allowance if rate is non-zero.
'        If AllowancePayRate > 0# Then
'            txtIncentive.Text = Format(AllowancePayRate * Val(txtNoDays.Text), "##0.00")
'        End If
        
        'Call GetDeductionsFromTable
        txtDaysWork.SetFocus
    ElseIf cmdTransaction(0).Caption = "&Edit Transaction" Then
        ' Edit Transaction
        EditFlag = True
        cmdEditTax(3).Enabled = True
        cmdTransaction(0).Caption = "&Save Transaction"
        cmdTransaction(1).Caption = "&Cancel Transaction"
        Call DisableCmdButtons
        cmdDTR.Enabled = True
        'Call AutoClickText  ' Updated July 17, 2016 disable first
        
'        ' Get Hazard and Union Dues.
''        If WithHazardPay = True And (Val(txtHazardPay.Text) = 10) And (txtDaysWork.Text / 8) > 1 Then
''            txtHazardPay.Text = Format(Val(txtHazardPay.Text) * (txtDaysWork.Text / 8), "##0.00")
''        End If
'        If WithHazardPay = True Then
'            'txtHazardPay.Text = Format(HazardPayRate * (txtDaysWork.Text / 8), "##0.00")
'            txtHazardPay.Text = Format(HazardPayRate * Val(txtNoDays.Text), "##0.00")
'        End If
        
        'txtDaysWork.SetFocus
    ElseIf cmdTransaction(0).Caption = "&Save Transaction" Then
        ' Save Transaction
        Call cmdSaveTransaction
        cmdTransaction(1).Caption = "&Delete Transaction"
        lvwProcessEmp_Click
        ' Save payments back to deduction table.
        tt = LoanDeductionAmount(1, StartCutoffDate, EndCutoffDate, True, 0, False)
        cmdDTR.Enabled = False
    End If
Else    ' Index is 1
    If cmdTransaction(Index).Caption = "&Cancel Transaction" Then
        ' Cancel Transaction
        Call cmdCancelTransaction
        cmdEditTax(3).Enabled = False
    ElseIf cmdTransaction(Index).Caption = "&Delete Transaction" Then
        ' Delete transaction
        res = MsgBox("Delete employee " & EmployeeName & "?", vbQuestion + vbYesNo + vbDefaultButton2, PayrollAlert)
        If res = vbYes Then
            Call DeleteTransaction
            ' Need to delete all Save payments to deduction table.
            tt = LoanDeductionAmount(1, StartCutoffDate, EndCutoffDate, True, 1, False)
        End If
    End If
End If
End Sub

'Private Sub GetDeductionsFromTable()
''Get the entries from client.
'txtTASAesla.Text = Format(mCBAmount, "##,##0.00")
'txtMiscFee.Text = Format(mMFAmount, "##,##0.00")
'txtOverPay.Text = Format(BDOCardFee, "##,##0.00")
'txtPlacementFee.Text = Format(mPFAmount, "##,##0.00")
'txtSSSLoan.Text = Format(mSSSLoan, "##,##0.00")
'txtPagibigLoan.Text = Format(mPBigLoan, "##,##0.00")
'txtCalamityLoan.Text = Format(mCalamityLoan, "##,##0.00")
'txtIncentive = Format(mAllowance, "##,##0.00")
'TxtAmount(10) = Format(mAllowance, "##,##0.00")
'End Sub

Private Sub GetCommissionEntries(ByVal mEmpId As Integer, _
    ByVal mPos As String, ByVal CutoffDate As Integer)

Dim rsCommission As ADODB.Recordset
 
StrQuery = "SELECT * FROM vwDriverCommissionRecords WHERE EmployeeID = " & _
    mEmpId & " AND format(TransDate, ""yyyymmdd"") = " & CutoffDate
Set rsCommission = ReturnRst(StrQuery, cnnstr, mstatic)
If Not rsCommission.EOF Then
      'rsCommission = rsCommission!AgencyName
 Else
    MsgBox "No Transaction records found for " & lblEmpName.Caption, vbInformation, PayrollAlert
End If
Set rsCommission = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim strword As String
Dim AmountInWords As Double
AmountInWords = 343432.22
strword = ConvertAmountinWords(AmountInWords)
End Sub

Private Function ConvertAmountinWords(ByVal mAmt As Double) As String
MsgBox oCommonFunction.Word(mAmt)
End Function

'Private Sub Form_Activate()
''Call lvwEmployeeList_Click
'End Sub

Private Sub Form_Load()

'If AccountType = "FMS" Then
'    cmdDTR.Caption = "FMS"
'    cmdDTR.ToolTipText = "Read FMS Entries..."
'Else
'    cmdDTR.Caption = "DTR"
'    cmdDTR.ToolTipText = "Read DTR Entries..."
'End If

cmdDTR.Enabled = False
cmdEditTax(3).Enabled = False
InsertedFlag = False
Me.Top = (mdiMain.Height - Me.Height) / 8
Me.Left = (mdiMain.Width - Me.Width) / 2

TotalAmt = 0#   ' variable for total gross payroll

'cmdSaveTrans(4).Visible = False
'cmdCancelTrans(5).Visible = False

mUserID = GETuSERNAME(operatorID)
empBasicSw = False
mClientSw = False
mAgencyCode = rsCompanyTable!agencycode
Call DisableTextBoxes(False)
Me.Caption = "Payroll Transaction Entries for " & Replace(mCompanyName, "&", "&&")
lblCompany.Caption = "Company: " & Replace(mCompanyName, "&", "&&")
lblCutoff.Caption = "Payroll Period: " & Format(frmSelectCutoff.txtCutOffDate(0).Text, "mmm dd, yyyy") & " to " & Format(frmSelectCutoff.txtCutOffDate(1).Text, "mmm dd, yyyy")
frameExcessOT.Visible = False
FrameDeductions.Visible = False

' Other Deductions, Updated Feb. 29, 2008. Bernard
chkOtherDed.Enabled = True
'strQuery = "SELECT * FROM qryEmployeeDeductions WHERE AgencyCode = " & CompanyCode & " and ClientCode = " & mClientCode & " and ProtectedFlag = 0"
StrQuery = "SELECT * FROM qryEmployeeDeductions WHERE ProtectedFlag = 0"

Set rsEmpDeductions = ReturnRst(StrQuery, CnnstrHR, mstatic)
Call FillDeductionList(ListDeductions, rsEmpDeductions)

'Get all Client settings.
Call GetClientInformation
' Get previous payroll.
Call GetPreviousPayroll

'Check whether new cutoff or the previous cut-off is open.
If frmSelectCutoff.chkNewCutoff.VALUE = 1 _
  Or frmSelectCutoff.Image2.Visible = True Then
    NewCutoffFlag = True
    Call GetEmployeeBasic
    'Open Payroll reference table.
    Call OpenReferenceTable
    If rsEmployeeInfo.RecordCount <> 0 Then
        Call GetBasicInformation
    End If
    Call PayrollTransRec
Else ' The cut-off selected is for viewing only, retrieve it from master table.
    NewCutoffFlag = False
    cmdTransaction(0).Enabled = False
    cmdTransaction(1).Enabled = False
    Call GetPayrollFromMaster
    Call PayrollMasterRec
End If

'If PaymentBasis = "Monthly" And Val(Format(frmSelectCutoff.txtCutoffDate(1).Text, "dd")) < 17 Then
'PaymentBasis = "Monthly"  ' Comment by Bernard
If Val(Format(frmSelectCutoff.txtCutOffDate(1).Text, "dd")) < 17 Then
    '        If MsgBox("Apply Monthly SSS Deduction?", vbYesNo, "SSS Deductions") = vbYes Then
                WithSSSDedFlag = True
                WithPHDedFlag = True   ' flag whether philhealth deduction will apply
                WithPBIGDedFlag = False  ' flag whether pagibig deduction will apply
                WTaxFlag = True     ' updated Jan. 14, 2012 to compute tax for every payday
                mCutOffNo = 1
            Else
                WithSSSDedFlag = False
                WithPHDedFlag = False   ' flag whether philhealth deduction will apply
                WithPBIGDedFlag = True  ' flag whether pagibig deduction will apply
                WTaxFlag = True  ' flag whether tax deduction will apply
                mCutOffNo = 2
    '        End If
'ElseIf LMCDeduction <> 0 And Val(Format(frmSelectCutoff.txtCutoffDate(1).Text, "dd")) = LMCDeduction Then
'        WithSSSDedFlag = True
'Else
'        WithSSSDedFlag = False
End If

Call GetBasicInformation ' updated Dec. 6, 2016.
'Call lvwEmployeeList_Click

End Sub

Private Sub Form_Initialize()
    Set PayrollValidation = New clsCommon
End Sub

Private Sub Form_Terminate()
    Set PayrollValidation = Nothing
End Sub

Private Sub GetEmployeeBasic()

'Get all the list of employees for processing.

'strQuery = "SELECT a.*, (trim(b.surname) + ', ' + trim(b.firstname) + ' ' + iif(b.middlename = '', '' , left(b.middlename, 1) + '.') ) AS fullname, " & _
'            "b.sss, b.philhealth, b.pagibig, b.acctno, b.TIN, b.employmentstatus FROM employeebasicsalary AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID " & _
'            "WHERE a.clientcode = " & mClientCode & " and a.employeestatus = 1 ORDER BY b.surname, b.firstname"

' Updated by Bernard to exclude all finished contract employees. January 27, 2007.
'StrQuery = "SELECT a.*, " & _
'            "b.sss, b.philhealth, b.pagibig, b.acctno, b.TIN, b.employmentstatus FROM qryEmployeeBasics AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID " & _
'            "WHERE a.clientcode = " & mClientCode & " and a.employeestatus = 1 and a.payrollstatus not in ('T','W','X','Y','Z') ORDER BY b.surname, b.firstname"

' Updated Dec. 6, 2016.
StrQuery = "SELECT a.*, " & _
    "b.sss, b.philhealth, b.pagibig, b.acctno, b.TIN, b.employmentstatus, b.DateHired, b.Town AS AllowanceBase, " & _
    "CASE WHEN b.EcolaAmount <= 0 then a.EcolaAmount ELSE b.EcolaAmount END AS EcolaAmount1 " & _
    "FROM qryEmployeeBasics AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID " & _
    "WHERE a.clientcode = " & mClientCode & " and a.employeestatus = 1 and a.payrollstatus not in ('T','W','X','Y','Z') ORDER BY b.surname, b.firstname"

Set rsEmployeeInfo = New ADODB.Recordset
Set rsEmployeeInfo = ReturnRst(StrQuery, cnnstr, mstatic)
If rsEmployeeInfo.RecordCount = 0 Then
    Call DisableAllButtons
End If
Call FillEmployeeList(lvwEmployeeList, lvwProcessEmp, rsEmployeeInfo)

End Sub

Private Sub GetPreviousPayroll()

StrQuery = "SELECT a.EmployeeID, a.EmpNo, a.sssEmployee, a.PHealthEmployee, a.PagibigEmployee " & _
            "FROM payrollmaster a WHERE a.clientcode = " & mClientCode & _
            " and CONVERT(int, RIGHT(a.startpaydate,4)) = " & Year(frmSelectCutoff.txtCutOffDate(0).Text) & _
            " and CONVERT(int, LEFT(a.startpaydate,2)) = " & month(frmSelectCutoff.txtCutOffDate(0).Text)

Set rsPreviousEmployee = New ADODB.Recordset
Set rsPreviousEmployee = ReturnRst(StrQuery, cnnstr, mstatic)

End Sub

Private Sub GetPayrollFromMaster()
Dim StartDate As String
Dim EndDate As String
Dim litem1 As ListItem

StartDate = frmSelectCutoff.txtCutOffDate(0).Text
EndDate = frmSelectCutoff.txtCutOffDate(1).Text

'strQuery = "SELECT a.EmployeeID, a.EmpNo, a.employeename AS fullname, a.startpaydate, a.endpaydate, b.paymentbasis, b.startcontract, b.endcontract, b.dailyrate, b.monthlyrate " & _
'            "FROM payrollmaster AS a LEFT JOIN employeebasicsalary AS b ON a.EmployeeID = b.EmployeeID " & _
'            "WHERE a.clientcode = " & mclientcode & " and a.startpaydate = '" & StartDate & "' and a.endpaydate = '" & EndDate & "' ORDER BY employeename"

'StrQuery = "SELECT a.EmployeeID, a.EmpNo, a.employeename AS fullname, a.startpaydate, a.endpaydate, b.paymentbasis, b.startcontract, b.endcontract, b.hourlyrate, b.dailyrate, b.monthlyrate " & _
'            "FROM payrollmaster AS a LEFT JOIN qryEmployeeBasics AS b ON a.EmployeeID = b.EmployeeID " & _
'            "WHERE a.clientcode = " & mClientCode & " and a.startpaydate = '" & StartDate & "' and a.endpaydate = '" & EndDate & "' ORDER BY employeename"

StrQuery = "SELECT a.EmployeeID, a.EmpNo, a.employeename AS fullname, a.startpaydate, a.endpaydate, b.paymentbasis, b.startcontract, b.endcontract, b.hourlyrate, b.dailyrate, b.monthlyrate, c.DateHired " & _
            "FROM payrollmaster AS a LEFT JOIN qryEmployeeBasics AS b ON a.EmployeeID = b.EmployeeID " & _
            "LEFT JOIN HREmployees AS C ON a.EmployeeID = c.EmployeeID " & _
            "WHERE a.clientcode = " & mClientCode & " and a.startpaydate = '" & StartDate & "' and a.endpaydate = '" & EndDate & "' ORDER BY employeename"

Set rsEmployeeInfo = New ADODB.Recordset
Set rsEmployeeInfo = ReturnRst(StrQuery, cnnstr, mstatic)
If rsEmployeeInfo.RecordCount = 0 Then
    Call DisableAllButtons
    Exit Sub
End If

lvwProcessEmp.View = lvwReport
mCtr1 = 0
lblProcessedEmp.Caption = 0
lvwProcessEmp.ListItems.Clear
With rsEmployeeInfo
    While Not (.EOF Or .BOF)
        Set litem1 = lvwProcessEmp.ListItems.Add(, , IIf(IsNull(!FullName), "", !FullName))
        litem1.SubItems(1) = !EmployeeID
        mCtr1 = mCtr1 + 1
        lblProcessedEmp.Caption = mCtr1
        .MoveNext
    Wend
End With

End Sub

Private Sub GetClientInformation()
Dim rsCutOff As ADODB.Recordset

' We need to determine the last cut-off.
StrQuery = "SELECT * FROM cutofftable WHERE clientcode = " & mClientCode & " ORDER BY recno"
Set rsCutOff = ReturnRst(StrQuery, cnnstr, mstatic)
If rsCutOff.RecordCount <> 0 Then
    rsCutOff.MoveLast
    'mCutOffNo = rsCutOff!CutOffNo
    mBilledFlag = rsCutOff!BilledFlag
Else
    'mCutOffNo = 0
    mBilledFlag = False
End If

WSSSFlag = False
WPHealthFlag = False
WPagibigFlag = False
'WTaxFlag = False
With rsClientTable
    If .RecordCount = 0 Then
        Exit Sub
    End If
    .MoveFirst
    .Find "clientcode = " & mClientCode, 0, adSearchForward ' goto selected client record
    
    If .EOF = True Then
        .MoveLast
    ElseIf .BOF = True Then .MoveFirst
    End If
    
    'FMS Flag Updated June 1, 2012
    FMSFlag = IIf(IsNull(!ClientCode2), False, IIf(!ClientCode2 = "FMS", True, False))
    
    ' Get work hours.
    WorkHrPerDay = IIf(IsNull(!WorkHrsPerDay), 8, !WorkHrsPerDay)
    NoofDaysperMonth = IIf(IsNull(!NoofDaysperMonth), 26, !NoofDaysperMonth)
    WorkDaysperYear = IIf(IsNull(!WorkingDaysperYear), 303, !WorkingDaysperYear)
    
    ' Get SSS Info
    WSSSFlag = !ClientSSSFlag
    ClientSSSRate = IIf(IsNull(!ClientSSSRate), 0#, !ClientSSSRate)
    ' Get PhilHealth Info
    WPHealthFlag = !ClientPhilHealthFlag
    ClientPHRate = IIf(IsNull(!ClientPHRateperDay), 0#, !ClientPHRateperDay)
    ' Get Pagibig info
    WPagibigFlag = !ClientPagibigFlag
    ClientPagibigRate = IIf(IsNull(!ClientPagibigRate), 0#, !ClientPagibigRate)
    mPagibigPay = IIf(WPagibigFlag = True, !PagibigPayments, 0#)
    ' Get Tax info
    'WTaxFlag = !ClientTaxFlag
    EditTaxFlag = !ClientEditTaxFlag
    'If Client Tax Percentage is not zero, the percentage tax computation will apply
    ClientTaxPercentageFlag = !ClientTaxPercentFlag
    'If ClientTaxPercentageFlag = True Then ' Updated by Bernard
        TaxPercentage = IIf(IsNull(!ClientTaxPercentage), 0#, !ClientTaxPercentage)
    'End If
    
    ' Get ECOLA Rate.
    WEcolaFlag = !ClientEcolaFlag
    WTaxableEcola = !ClientEcolaTaxFlag
    ' Check if Allowance is taxable, updated Feb. 7, 2015.
    WTaxableAllowance = !EcolawLateDeduction
    
    WEcolaDeduction = False  '!EcolawLateDeduction
    DoubleEcolaonSpl = !EcolaDoubleonSpl
    DoubleEcolaonLgl = !EcolaDoubleonLegal
    EcolaPayonRestDay = !EcolaPayonRestDay
    EcolaPayonHoliday = !EcolaPayOnHolidays
    'mECOLARate = IIf(WEcolaFlag = True, !ecolarate, 0#) ' updated to PFAmount from HREmployees
    
'    ' Get SSS and PHealth Add on Rate. Updated Nov. 10, 2009.
'    HazardPayRate = IIf(!HazardPayFlag = True, !HazardPayRate, 0#)
'    UnionDuesRate = IIf(!UnionDuesFlag = True, !UnionDuesRate, 0#)
    
    ' 5/19/15: Check for VarPercentRate, if not-zero it means allowance is auto-computed similar to hazard.
    'AllowancePayRate = Format(IIf(!VatPercentRate = "", "0.00", !VatPercentRate), "##0.00") ' change to HREmployees PFPayment
    
    'New Rate of Phealth, updated 1/6/2018
    NewPHRate = Format(IIf(!VatPercentRate = "", "0.00", !VatPercentRate), "##0.00") ' change to HREmployees PFPayment
    UseTax2018Flag = !EcolaPayonRestDay ' use tax table 2018 when checked
    
    If NewPHRate <> 0 Then
        lblEntries(41).ForeColor = &HFF& ' red
        lblEntries(42).ForeColor = &HFF&
    Else
        lblEntries(41).ForeColor = &H80000012 ' blank
        lblEntries(42).ForeColor = &H80000012
    End If
    
    ' Get Hazard Pay and Union Dues. Updated May 16, 2011. Bernard.
    If !HazardPayFlag = True Then
        HazardPayRate = Format(IIf(!HazardPayRate = "", "0.00", !HazardPayRate), "##0.00")
        lblEntries(55).Visible = True
        txtHazardPay.Visible = True
        WithHazardPay = True
    Else
        lblEntries(55).Visible = False
        txtHazardPay.Visible = False
        HazardPayRate = 0#
        WithHazardPay = False
    End If
    'Get Union Dues.
    If !UnionDuesFlag = True Then
        UnionDuesRate = Format(IIf(!UnionDuesRate = "", "0.00", !UnionDuesRate), "##0.00")
        lblEntries(56).Visible = True
        txtUnionDues.Visible = True
        WithHazardPay = True
    Else
        lblEntries(56).Visible = False
        txtUnionDues.Visible = False
        UnionDuesRate = 0#
        WithHazardPay = False
    End If
    ' Make union dues deduction first half of the month.
    If Val(Format(frmSelectCutoff.txtCutOffDate(1).Text, "dd")) > 15 Then
        UnionDuesRate = 0
    End If
   
    'MonthlyEcola = mECOLARate * NoofDaysperMonth
    'Key in excluding sss, pagibig, philhealth from tax computation
    mExcludedTax = !withvatcharges
    ' key in excluding Ecola from SSS, Philhealth Computation
    ' mExcludedEcola = !ClientFlag3  ' previously for Ecola, updated jan 17, 2015
    mOTExcluded = !ClientFlag3
    mHalfSSSPay = !ClientFlag4
    
    ' Change deduction caption updated Feb. 18, 2012.
    If IsNull(!BillingReportType) Or !BillingReportType = "" Or !BillingReportType = "Other Clients" Then
        LMCDeduction = "LMC Deductions"
        'lblEntries(34).Caption = "LMC Deductions:"
        lblEntries(34).Caption = "Associations:"
    Else
        LMCDeduction = !BillingReportType
        lblEntries(34).Caption = IIf(!BillingReportType = "Assoctn", "Associations:", !BillingReportType) & ":"
    End If
    
    Call UpdateLabel
    
    ' Get Fix Rates.
    OverTimeRate = IIf(IsNull(!OverTimeRate), 0#, !OverTimeRate)
    SundayRate = IIf(IsNull(!SundayRate), 0#, !SundayRate)
    TardinessRate = IIf(IsNull(!TardinessRate), 0#, !TardinessRate)
    NDiffRate = IIf(IsNull(!NDiffRate), 0#, !NDiffRate)
    SplHolidayRate = IIf(IsNull(!SpecialHolidayRate), 0#, !SpecialHolidayRate)
    LglHolidayRate = IIf(IsNull(!LegalHolidayRate), 0#, !LegalHolidayRate)
    SplHolidayOTRate = IIf(IsNull(!SplHolidayOTRate), 0#, !SplHolidayOTRate)
    LglHolidayOTRate = IIf(IsNull(!legalholidayotrate), 0#, !legalholidayotrate)
    ExcessOTRegDayRate = IIf(IsNull(!ExcessOTRegDayRate), 0#, !ExcessOTRegDayRate)
    ExcessOTSplSunRate = IIf(IsNull(!ExcessOTSplSunRate), 0#, !ExcessOTSplSunRate)
    ExcessOTSunLglRate = IIf(IsNull(!TardinessRate), 0#, !TardinessRate)    ' RD SPL / LGL Rate
    ExcessOTLegalRate = IIf(IsNull(!LateChargesinMin), 0#, !LateChargesinMin)   ' RD on Legal Rate
    OTwoPremiumRate = IIf(IsNull(!OTwoPremiumRate), 0#, !OTwoPremiumRate)   ' RD on SPL Rate
    RestDayRate = IIf(IsNull(!RestDayRate), 0#, !RestDayRate)
    
    ' Display 13th Month Pay or Not
    Display13thPay = !display13monthpay
    
    ' Get Cutoff.
    ClientCutoff = !ClientCutoff
    If ClientCutoff = "Weekly" Then
'        If mCutOffNo = 4 Then
'            mCutOffNo = 1
'        Else
'            mCutOffNo = mCutOffNo + 1
'        End If
        NoofCutoffs = 1  '4
    ElseIf ClientCutoff = "Semi-Monthly" Then
'        If mCutOffNo = 2 Then
'            mCutOffNo = 1
'            'WithSSSDedFlag = False
'        Else
'            mCutOffNo = mCutOffNo + 1
'            'If mCutOffNo = 2 Then WithSSSDedFlag = True ' end of the month, apply deductions to sss, HDMP, philhealth
'        End If
        NoofCutoffs = 1  '2
    Else        ' for monthly cut-offs
        NoofCutoffs = 1
    End If
    ' Get Agency Rate
    WCalamityLoan = !ClientAgencyFlag
    If WCalamityLoan = True Then
        AgencyRate = IIf(IsNull(!ClientAgencyRate), 0#, !ClientAgencyRate)
        EmployeeAgencyRate = 0#
    Else
        AgencyRate = 0#
        EmployeeAgencyRate = IIf(IsNull(!clientchargeperday), 0#, !clientchargeperday)
    End If
End With

End Sub

Private Sub UpdateLabel()
If FMSFlag = True Then
    lblEntries(12).Caption = "Damage/Accident:"
    'lblEntries(35).Caption = "Vale:"
    lblEntries(7).Caption = "Shortages:"
    'lblEntries(34).Caption = "Associations:"
    'lblEntries(36).Caption = "Surety Fund:"
    'lblEntries(8).Caption = "Uniform:"
    lblEntries(54).Caption = "Surety Fund:"
Else
    lblEntries(12).Caption = "Other Deductions:"
    'lblEntries(35).Caption = "Vale:"
    lblEntries(7).Caption = "Planters Loan:"
    'lblEntries(34).Caption = "Associations:"
    'lblEntries(36).Caption = "Cash Advance:"
    'lblEntries(8).Caption = "Uniform:"
    lblEntries(54).Caption = "HDMF Calamity:" ' updated Jan 17, 2015 from Calamity Loan:
End If

End Sub

Private Sub FillEmployeeList(ByRef mLsEmp As ListView, _
                        ByRef mLsProc As ListView, _
                        mRs As Recordset)
Dim litem1 As ListItem
Dim litem2 As ListItem

    mLsEmp.View = lvwReport
    mLsProc.View = lvwReport
    mCtr1 = 0
    mCtr2 = 0
    lblProcessedEmp.Caption = 0
    lblTotalEmployee.Caption = 0
    mLsEmp.ListItems.Clear
    mLsProc.ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
        If mRs!ProcessedFlag = True Then ' add to process list view
'            Set litem1 = mLsProc.ListItems.Add(, , IIf(IsNull(mRs(47)), "", mRs(47)))
            Set litem1 = mLsProc.ListItems.Add(, , IIf(IsNull(mRs!FullName), "", mRs!FullName))
            litem1.SubItems(1) = mRs(0)
            mCtr1 = mCtr1 + 1
            lblProcessedEmp.Caption = mCtr1
        Else                        ' add to employee list view
            'Set litem2 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs(47)), "", mRs(47)))
            Set litem2 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!FullName), "", mRs!FullName))
            litem2.SubItems(1) = mRs(0)
            mCtr2 = mCtr2 + 1
            lblTotalEmployee.Caption = mCtr2
        End If
        mRs.MoveNext
    Wend
End Sub

Private Sub OpenReferenceTable()
Dim StrQuery As String

StrQuery = "SELECT minwage, maxwage, er_share, ec_share, ee_share FROM SSStable order by id"
Set rsSSSTable = ReturnRst(StrQuery, cnnstr, mstatic)
StrQuery = "SELECT minwage, maxwage, ph_er, ph_ee FROM PhilHealthtable order by id"
Set rsPHealthTable = ReturnRst(StrQuery, cnnstr, mstatic)
StrQuery = "SELECT minwage, maxwage, pagibigee, pagibiger, pagibigmax FROM Pagibigtable order by id"
Set rsPagibigTable = ReturnRst(StrQuery, cnnstr, mstatic)

''Open tax table depending on cut-off.
'Select Case ClientCutoff
'    Case "Daily" ' daily cut-off used the daily table
'        StrQuery = "SELECT taxcode, salrate2, basetax, perover FROM DailyTaxTable order by id"
'        Set rsTaxTable = ReturnRst(StrQuery, cnnstr, mstatic)
'    Case "Weekly" ' weekly cut-off used the weekly table
'        StrQuery = "SELECT taxcode, salrate2, basetax, perover FROM WeeklyTaxTable order by id"
'        Set rsTaxTable = ReturnRst(StrQuery, cnnstr, mstatic)
'    Case "Semi-Monthly" ' semi-monthly cut-off used the semimonthly table
'        StrQuery = "SELECT taxcode, salrate2, basetax, perover FROM SemiMonthlyTaxTable order by id"
'        Set rsTaxTable = ReturnRst(StrQuery, cnnstr, mstatic)
'        StrQuery = "SELECT taxcode, salrate2, basetax, perover FROM MonthlyTaxTable order by id"
'        Set rsTaxTable2 = ReturnRst(StrQuery, cnnstr, mstatic)
'    Case "Monthly" ' Monthly cut-off used the monthly table
'        StrQuery = "SELECT taxcode, salrate2, basetax, perover FROM MonthlyTaxTable order by id"
'        Set rsTaxTable = ReturnRst(StrQuery, cnnstr, mstatic)
'    Case Else
'        StrQuery = "SELECT taxcode, salrate2, basetax, perover FROM SemiMonthlyTaxTable order by id"
'        Set rsTaxTable = ReturnRst(StrQuery, cnnstr, mstatic)
'End Select

' Commented January 14, 2012. Bernard
'    StrQuery = "SELECT taxcode, salrate2, basetax, perover FROM DailyTaxTable order by id"
'    Set rsTaxTable2 = ReturnRst(StrQuery, cnnstr, mstatic)
'    StrQuery = "SELECT taxcode, salrate2, basetax, perover FROM MonthlyTaxTable order by id"
'    Set rsTaxTable = ReturnRst(StrQuery, cnnstr, mstatic)

    If UseTax2018Flag = True Then
        'New tax table 2018.
        StrQuery = "SELECT taxcode, salrate2, basetax, perover FROM SemiMonthlyTax2018 order by id"
    Else
        StrQuery = "SELECT taxcode, salrate2, basetax, perover FROM SemiMonthlyTaxTable order by id"
    End If
    Set rsTaxTable = ReturnRst(StrQuery, cnnstr, mstatic)

End Sub

Private Sub SSSPayment(ByVal mBasicPay As Double)
Dim sssee As Double
Dim ssser As Double

txtEmpyeeSSS.Text = "0"
txtEmpyerSSS.Text = "0"
txtECC.Text = "0"

'Check whether SSS deductions is monthly. Written April 15, 2006.
'If PaymentBasis = "Monthly" And WithSSSDedFlag = False Then
If WithSSSDedFlag = False Then
    Exit Sub
End If

'If mExcludedEcola = False Then 'Ecola is excluded from sss, phealth computation
'    mBasicPay = mBasicPay + MonthlyEcola
'End If

With rsSSSTable
    .MoveFirst
    Do While Not .EOF
        If mBasicPay >= !minwage And mBasicPay <= !maxwage Then
            txtEmpyeeSSS.Text = !ee_share ' + HazardPayRate ' add on rate, updated 11/10/2009
            txtEmpyerSSS.Text = !er_share + !ec_share
            txtECC.Text = !ec_share ' Save ECC amount 8/6/2009
'            sssee = Format(txtEmpyeeSSS / NoofDaysperMonth, "##0.00")
'            ssser = Format(txtEmpyerSSS / NoofDaysperMonth, "##0.00")
'            If PaymentBasis = "Hourly" Then
'                txtEmpyeeSSS.Text = Format((sssee / 8) * txtDaysWork, "##,##0.00")
'                txtEmpyerSSS.Text = Format((ssser / 8) * txtDaysWork, "##,##0.00")
'            ElseIf PaymentBasis = "Monthly" Then
                txtEmpyeeSSS.Text = Format(txtEmpyeeSSS, "##,##0.00")
                txtEmpyerSSS.Text = Format(txtEmpyerSSS, "##,##0.00")
'            Else
'                txtEmpyeeSSS.Text = Format(sssee * txtDaysWork, "##,##0.00")
'                txtEmpyerSSS.Text = Format(ssser * txtDaysWork, "##,##0.00")
'            End If
'            ' SSS is half computed.
'            If mHalfSSSPay = True Then
'                'txtEmpyeeSSS.Text = Format(!ee_share / 2, "##,##0.00")
'                'txtEmpyerSSS.Text = Format(!er_share / 2, "##,##0.00")
'                ' Updated 8/6/2009
'                txtEmpyeeSSS.Text = Format(txtEmpyeeSSS / 2, "##,##0.00")
'                txtEmpyerSSS.Text = Format(txtEmpyerSSS / 2, "##,##0.00")
'            End If
            Exit Sub
        End If
        .MoveNext
    Loop
End With
End Sub

Private Sub PHealthPayment(ByVal mBasicPay As Double)
Dim medee As Double
Dim meder As Double

txtPhilHealthYee.Text = "0"
txtPHealthYer.Text = "0"

' Check whether SSS deductions is monthly. Written April 15, 2006.
'If PaymentBasis = "Monthly" And WithSSSDedFlag = False Then
If WithPHDedFlag = False Then
    Exit Sub
End If

'If mExcludedEcola = False Then 'Ecola is excluded from sss, phealth computation
'    mBasicPay = mBasicPay + MonthlyEcola
'End If

If NewPHRate <> 0 Then ' apply the newphealth rates.
'    txtPhilHealthYee.Text = Format(NewPHealthRates(mBasicPay), "##,##0.00")
'    txtPHealthYer.Text = txtPhilHealthYee.Text

    txtPhilHealthYee.Text = Format(NewPHealthRates(mBasicPay), "##,##0.00")
    txtPHealthYer.Text = Format(NewPHealthRates(mBasicPay, True), "##,##0.00")

Else
    With rsPHealthTable
    .MoveFirst
    Do While Not .EOF
        If mBasicPay >= !minwage And mBasicPay <= !maxwage Then
            txtPhilHealthYee.Text = !ph_ee ' + UnionDuesRate    ' Phealth add on rate
            txtPHealthYer.Text = !ph_er
'            medee = Format((txtPhilHealthYee / NoofDaysperMonth), "##0.00")
'            meder = Format((txtPHealthYer / NoofDaysperMonth), "##0.00")
'            If PaymentBasis = "Hourly" Then
'                txtPhilHealthYee.Text = Format((medee / 8) * txtDaysWork, "##,##0.00")
'                txtPHealthYer.Text = Format((meder / 8) * txtDaysWork, "##,##0.00")
'            ElseIf PaymentBasis = "Monthly" Then
                txtPhilHealthYee.Text = Format(txtPhilHealthYee, "##,##0.00")
                txtPHealthYer.Text = Format(txtPHealthYer, "##,##0.00")
'            Else
'                txtPhilHealthYee.Text = Format((txtPhilHealthYee / NoofDaysperMonth) * txtDaysWork, "##,##0.00")
'                txtPHealthYer.Text = Format((txtPHealthYer / NoofDaysperMonth) * txtDaysWork, "##,##0.00")
'            End If
'            ' SSS is half computed.
'            If mHalfSSSPay = True Then
'                txtPhilHealthYee.Text = Format(!ph_ee / 2, "##,##0.00")
'                txtPHealthYer.Text = Format(!ph_er / 2, "##,##0.00")
'            End If
            Exit Sub
        End If
        .MoveNext
    Loop
    End With
End If

End Sub

Function NewPHealthRates(ByVal mBasic As Double, Optional ForYer As Boolean = False) As Double
Dim TotMed As Double
Dim MedYer As Double

If ClientPHRate <> 0 Then
    NewPHealthRates = ClientPHRate
ElseIf mBasic <= 10000 Then
    NewPHealthRates = 150#   '137.5 updated 2/16/2020
Else
    'NewPHealthRates = ((mBasic * (NewPHRate / 100)) / 2)
    TotMed = mBasic * (NewPHRate / 100)
    TotMed = Format(TotMed, "###,##0.00")
    If (Right(TotMed, 1) Mod 2) <> 0 Then
        If ForYer = True Then
            MedYer = Format((TotMed / 2), "###,##0.00")
        Else
            MedYer = Format((TotMed / 2), "###,##0.00")
            MedYer = Format(MedYer - 0.01, "###,##0.00")
        End If
        NewPHealthRates = Format(MedYer, "###,##0.00")
    Else
        NewPHealthRates = Format(TotMed / 2, "###,##0.00")
    End If
End If

End Function

'
'Function NewPHealthRates(ByVal mBasic As Double) As Double
'Dim TotMed As Double
'Dim MedYer As Double
'
'If mBasic <= 10000 Then
'    NewPHealthRates = 137.5
'Else
'    'NewPHealthRates = ((mBasic * (NewPHRate / 100)) / 2)
'    ' Updated 7/22/2018
'
'    TotMed = Format((mBasic * (NewPHRate / 100)), "###,##0.00")
'    If (TotMed Mod 2) <> 0 Then
'        If ForYer = True Then
'            MedYer = Format((TotMed / 2), "###,##0.00")
'        Else
'            MedYer = Format((TotMed / 2), "###,##0.00")
'            MedYer = MedYer - 0.01
'        End If
'        NewPHealthRates = MedYer
'    Else
'        NewPHealthRates = TotMed / 2
'    End If
'End If
'
'End Function

Private Sub PagibigPayment(ByVal mBasicPay As Double)

txtEmpyeePagibig.Text = "0"
txtEmpyerPagibig.Text = "0"

' Check whether SSS deductions is monthly. Written April 15, 2006.
'If PaymentBasis = "Monthly" And WithSSSDedFlag = False Then
If WithPBIGDedFlag = False Then
    Exit Sub
End If

'If mExcludedEcola = False Then 'Ecola is excluded from sss, phealth computation
'    mBasicPay = mBasicPay + MonthlyEcola
'End If

If mPagibigPay <> 0 Then ' use the pagibig entries
    txtEmpyeePagibig.Text = Format(mPagibigPay, "##,##0.00")
    txtEmpyerPagibig.Text = Format(mPagibigPay, "##,##0.00")
Else    ' if pagibig contribution default is 0, use the table in computation
    With rsPagibigTable
    .MoveFirst
    Do While Not .EOF
        If mBasicPay >= !minwage And mBasicPay <= !maxwage Then
            'txtEmpyeePagibig.Text = Format((IIf(mBasicPay > 5000, !pagibigmax, !pagibigee) / NoofCutoffs), "##,##0.00")
            'txtEmpyerPagibig.Text = Format((IIf(mBasicPay > 5000, !pagibigmax, !pagibiger) / NoofCutoffs), "##,##0.00")
            txtEmpyeePagibig.Text = Format((mBasicPay * !pagibigee), "##,##0.00")
            txtEmpyerPagibig.Text = Format((mBasicPay * !pagibiger), "##,##0.00")
            If txtEmpyerPagibig.Text = "0.02" Then
                txtEmpyerPagibig.Text = "0.00"
                txtEmpyeePagibig.Text = "0.00"
            End If
            ' SSS is half computed.
            If mHalfSSSPay = True Then
                txtEmpyerPagibig.Text = "50.00"
                txtEmpyeePagibig.Text = "50.00"
            ElseIf txtEmpyeePagibig.Text < 50# Then
                txtEmpyerPagibig.Text = "0.00"
                txtEmpyeePagibig.Text = "0.00"
            End If
        Exit Sub
        End If
        .MoveNext
    Loop
    End With
End If

End Sub

Private Sub ComputeForECOLA(ByVal EcolaSw As Integer, Optional rdHours As Double)
Dim i As Integer
Dim strAmt As String
Dim mAmt As String
Dim DeciSw As Boolean
Dim mEcolaPay As Double, mAllowance2 As Double
Dim mEcolaRate2 As Double
Dim EcolaDeduc As Double
Dim NoofHoursLate As Double, HourlyEcola As Double, mHazard As Double, DailyEcola As Double

    mHazard = HazardPayRate * Val(txtNoDays.Text)
    mAllowance2 = mBSCAllowRate * Val(txtNoDays.Text)
    
'If EcolaSw = 5 Then
'    If IIf(txtSplHoliday.Text = "", 0, txtSplHoliday.Text) > 0 Then  ' DoubleEcolaonSpl = True And
'        mHazard = mHazard + (Val(txtSplHoliday / 8) * HazardPayRate)
'    End If
'    If IIf(txtLglHoliday.Text = "", 0, txtLglHoliday.Text) > 0 Then ' DoubleEcolaonLgl = True And
'        mHazard = mHazard + (Val(txtLglHoliday / 8) * HazardPayRate)
'    End If
'    If IIf(txtSickLeave.Text = "", 0, txtSickLeave.Text) > 0 Then ' for sickleave
'        mHazard = mHazard + (Val(txtSickLeave) * HazardPayRate)
'    End If
'    If IIf(txtVacLeave.Text = "", 0, txtVacLeave.Text) > 0 Then ' for vacation leave
'        mHazard = mHazard + (Val(txtVacLeave) * HazardPayRate)
'    End If
'    GoTo ExitHazard
'End If
    ' If ECOLA flag is false, make ecola 0.
    If WEcolaFlag = False Then
    Else
        If PaymentBasis = "Monthly" Then ' if monthly paid get the fixed Ecola rate
            'mECOLARate = mECOLARate / 2
            mEcolaPay = mECOLARate '* IIf(txtDaysWork.Text = "", 0, Val(txtNoDays.Text))
            HourlyEcola = (((mECOLARate * 2 * 12) / WorkDaysperYear) / 8)
            DailyEcola = ((mECOLARate * 2 * 12) / WorkDaysperYear)
        Else
            mEcolaPay = mECOLARate * IIf(txtDaysWork.Text = "", 0, Val(txtNoDays.Text))
            HourlyEcola = mECOLARate / 8
            DailyEcola = mECOLARate
        End If
        mEcolaRate2 = mECOLARate
        
        ' Check if ecola is taxable.
        'If WEcolaDeduction = True And (txtDeductions(1) > 0 Or txtDeductions(0) > 0) Then  ' deduc lates to Ecola payment
        If (txtDeductions(0) > 0) Then    ' deduc lates to Ecola payment
            If PaymentBasis = "Monthly" Then ' if monthly paid get the fixed Ecola rate per day
                'EcolaDeduc = (((mECOLARate * 2 * 12) / WorkDaysperYear) / 8)
'                NoofHoursLate = (IIf(txtHoursLate.Text = "", 0, txtHoursLate.Text) + (IIf(txtDaysAbsent.Text = "", 0, txtDaysAbsent.Text) * WorkHrPerDay))
                NoofHoursLate = (IIf(txtDaysAbsent.Text = "", 0, txtDaysAbsent.Text) * WorkHrPerDay)
                'EcolaDeduc = EcolaDeduc * NoofHoursLate
                EcolaDeduc = HourlyEcola * NoofHoursLate
            Else
                NoofHoursLate = (IIf(txtHoursLate.Text = "", 0, txtHoursLate.Text) + (IIf(txtDaysAbsent.Text = "", 0, txtDaysAbsent.Text) * WorkHrPerDay))
                EcolaDeduc = HourlyEcola * NoofHoursLate
            End If
            mEcolaPay = mEcolaPay - EcolaDeduc
        End If
        'Double Ecola on Special Holiday
        If DoubleEcolaonLgl = True And IIf(txtSplHoliday.Text = "", 0, txtSplHoliday.Text) > 0 Then  ' DoubleEcolaonSpl = True And
            'mEcolaPay = mEcolaPay + HourlyEcola
            mEcolaPay = mEcolaPay + (Val(txtSplHoliday) * HourlyEcola)
            mHazard = mHazard + (Val(txtSplHoliday / 8) * HazardPayRate)
            mAllowance2 = mAllowance2 + (Val(txtSplHoliday) * (mBSCAllowRate / 8))
        End If
        'Double Ecola on Regular (Legal) Holiday
        If DoubleEcolaonSpl = True And IIf(txtLglHoliday.Text = "", 0, txtLglHoliday.Text) > 0 Then ' DoubleEcolaonLgl = True And
            'mEcolaPay = mEcolaPay + HourlyEcola 'mEcolaRate2
            mEcolaPay = mEcolaPay + (Val(txtLglHoliday) * HourlyEcola)
            mHazard = mHazard + (Val(txtLglHoliday / 8) * HazardPayRate)
            mAllowance2 = mAllowance2 + (Val(txtLglHoliday) * (mBSCAllowRate / 8))
        End If

        If PaymentBasis <> "Monthly" Then ' if monthly paid get the fixed Ecola rate per day
            If IIf(txtSickLeave.Text = "", 0, txtSickLeave.Text) > 0 Then ' for sickleave
                mEcolaPay = mEcolaPay + (Val(txtSickLeave) * DailyEcola)
                mHazard = mHazard + (Val(txtSickLeave) * HazardPayRate)
                mAllowance2 = mAllowance2 + (Val(txtSickLeave) * mBSCAllowRate)
            End If
            If IIf(txtVacLeave.Text = "", 0, txtVacLeave.Text) > 0 Then ' for vacation leave
                mEcolaPay = mEcolaPay + (Val(txtVacLeave) * DailyEcola)
                mHazard = mHazard + (Val(txtVacLeave) * HazardPayRate)
                mAllowance2 = mAllowance2 + (Val(txtVacLeave) * mBSCAllowRate)
            End If
        End If

        ' Add Ecola for Restday.
        If (Val(txtSundayOT) > 0) Then
            If PaymentBasis = "Monthly" Then ' if monthly paid get the fixed Ecola rate per day
                'EcolaDeduc = (((mECOLARate * 2 * 12) / WorkDaysperYear) / 8)
                EcolaDeduc = HourlyEcola * Val(txtSundayOT)
            Else
                EcolaDeduc = (IIf(txtSundayOT.Text = "", 0, txtSundayOT.Text) * HourlyEcola)  '(HourlyEcola / WorkHrPerDay))
                'EcolaDeduc = mECOLARate * NoofHoursLate
            End If
            mHazard = mHazard + (Val(txtSundayOT / 8) * HazardPayRate)
            mEcolaPay = mEcolaPay + EcolaDeduc
        End If
        
        'Updated 6/17/2018 Bernard Add RD hours in ecola.
        If ChkVal(txtRDonSPL) <> 0 Then     ' Ecola for spl RD
            mEcolaPay = mEcolaPay + (HourlyEcola * ChkVal(txtRDonSPL))
        End If
        If ChkVal(txtRDonLGL) <> 0 Then     ' Ecola for lgl RD
            mEcolaPay = mEcolaPay + ((HourlyEcola * ChkVal(txtRDonLGL)) * 2)
        End If
        
        ' Total Ecola Amount.
        txtECOLA.Text = Format(mEcolaPay, "##,##0.00")
        TxtAmount(17) = Format(mEcolaPay, "##,##0.00")
        txtSplOtHours.Text = Format(mAllowance2, "##,##0.00")
    End If
    
ExitHazard:

    If EcolaSw <> 3 Then
        txtHazardPay.Text = Format(mHazard, "##,##0.00")
    End If
End Sub

Private Function GetPreviousBasicPay() As Double ' Get the basic pay from previous payroll for SSS computation
Dim mMonth As Integer
Dim mYear As Integer
Dim mPrevPay As Recordset

mMonth = Val(Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm"))
mYear = Val(Format(frmSelectCutoff.txtCutOffDate(0).Text, "yyyy"))

' If cut-off month is the same, change the month.
If mMonth = Val(Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm")) Then
    If mMonth = 1 Then
        mMonth = 12
        mYear = mYear - 1
    Else
        mMonth = mMonth - 1
    End If
    StrQuery = "SELECT GrossPay FROM PayrollMaster WHERE clientcode = " & mClientCode & _
            " AND EmployeeID = " & EmployeeNo & " AND (DAY(StartPayDate) = 16 OR DAY(StartPayDate) = 11 OR DAY(StartPayDate) = 25) AND " & _
            " MONTH(StartPaydate) = " & mMonth & " AND YEAR(StartPayDate) = " & mYear
Else
    StrQuery = "SELECT GrossPay FROM PayrollMaster WHERE clientcode = " & mClientCode & _
            " AND EmployeeID = " & EmployeeNo & " AND (DAY(EndPayDate) = 28 OR DAY(EndPayDate) = 29 OR DAY(EndPayDate) = 30 OR DAY(EndPayDate) = 31 OR DAY(EndPayDate) = 25) AND " & _
            " MONTH(EndPaydate) = " & mMonth & " AND YEAR(EndPayDate) = " & mYear
End If
Set mPrevPay = ReturnRst(StrQuery, cnnstr, mstatic)
With mPrevPay
    If .EOF = True Then
        GetPreviousBasicPay = 0#
    Else
        GetPreviousBasicPay = !GrossPay
    End If
End With
Set mPrevPay = Nothing

End Function

Private Sub GetContributions(ByVal mBasicPay As Double)

    If WSSSFlag = True Then
        ' Better get the monthly rate for SSS Contribution
        'Call SSSPayment(MonthlyRate)        ' mBasicPay ' updated by Bernard to use Basic pay in SSS
        Call SSSPayment(mBasicPay + GetPreviousBasicPay)
    End If
    If WPHealthFlag = True Then
        Call PHealthPayment(MonthlyRate)    ' mBasicPay
    End If
    If WPagibigFlag = True Then
        Call PagibigPayment(MonthlyRate)
    End If
    ' Get 13th month pay, if monthly paid, 13th month is fixed.
    If W13MonthPay = True Then
        ' Get 13th month pay, just divide rate by 12. June 30, 2012
        txt13thMonth.Text = Format((DailyRate / 12), "##,##0.00")
        TxtAmount(11).Text = Format((DailyRate / 12), "##,##0.00")
            
'        If PaymentBasis = "Monthly" Then
            
            
'            If Hourly13thMonthRate <> 0 Then ' get the default 13th month pay
''                txt13thMonth.Text = Format(IIf(IsNull(Hourly13thMonthRate), 0#, Hourly13thMonthRate - txtDeductions(1)), "##,##0.00")
''                TxtAmount(11).Text = Format(IIf(IsNull(Hourly13thMonthRate), 0#, Hourly13thMonthRate - txtDeductions(1)), "##,##0.00")
'                txt13thMonth.Text = Format(IIf(IsNull(Hourly13thMonthRate), 0#, Hourly13thMonthRate), "##,##0.00")
'                TxtAmount(11).Text = Format(IIf(IsNull(Hourly13thMonthRate), 0#, Hourly13thMonthRate), "##,##0.00")
'            Else
'                ' Deduc absences and lates, Oct. 14, 2006. Bernard
'                txt13thMonth.Text = Format((TxtAmount(0) - txtDeductions(1) + Val(Replace(txtECOLA.Text, ",", ""))) / 12, "##,##0.00")
'                TxtAmount(11).Text = Format((TxtAmount(0) - txtDeductions(1) + Val(Replace(txtECOLA.Text, ",", ""))) / 12, "##,##0.00")
'            End If


'        ElseIf PaymentBasis = "Hourly" Then ' updated Nov. 1, 2006
'            If Hourly13thMonthRate <> 0 Then ' get the default 13th month rate per hour
'                txt13thMonth.Text = Format(IIf(IsNull(Hourly13thMonthRate), 0#, Hourly13thMonthRate - txtDeductions(1)), "##,##0.00")
'                TxtAmount(11).Text = Format(IIf(IsNull(txt13thMonth.Text), 0#, (txt13thMonth * txtDaysWork.Text)), "##,##0.00")
'            Else
'                txt13thMonth.Text = Format(((((DailyRate * txtDaysWork.Text) - txtDeductions(1)) / 12) / 8), "##,##0.00")
'                TxtAmount(11).Text = Format(((((DailyRate * txtDaysWork.Text) - txtDeductions(1)) / 12) / 8), "##,##0.00")
'            End If
'        Else
'            If Hourly13thMonthRate <> 0 Then ' get the default 13th month rate per hour
'                txt13thMonth.Text = Format(IIf(IsNull(Hourly13thMonthRate), 0#, Hourly13thMonthRate - txtDeductions(1)), "##,##0.00")
'                TxtAmount(11).Text = Format(IIf(IsNull(txt13thMonth.Text), 0#, (txt13thMonth * txtDaysWork.Text)), "##,##0.00")
'            Else
'                txt13thMonth.Text = Format((((DailyRate * txtDaysWork.Text) - txtDeductions(1)) / 12), "##,##0.00")
'                TxtAmount(11).Text = Format((((DailyRate * txtDaysWork.Text) - txtDeductions(1)) / 12), "##,##0.00")
'            End If
'        End If
        lblEntries(29).Visible = True
        txt13thMonth.Visible = True
        lblEntries(24).Visible = True
        TxtAmount(11).Visible = True
    Else
        lblEntries(29).Visible = False
        txt13thMonth.Visible = False
        lblEntries(24).Visible = False
        TxtAmount(11).Visible = False
    End If
    
End Sub

Private Sub ComputeTax(ByVal mGrossPay As Double, ByRef rsTax As ADODB.Recordset, _
                Optional DeducFlag As Integer = 1)
                
Dim mBaseTax As Double
Dim mTaxRate As Double
Dim mSalRate As Double

'If PaymentBasis = "Monthly" Then ' do not include the ecola/allowance in tax computation
'   mGrossPay = mGrossPay - IIf(txtECOLA.Text = "", 0, txtECOLA.Text)
'End If

'    If WTaxFlag = True And DeducFlag = 2 Then   ' deduc absences
'        Call ComputeTax(txtGrossPay1.Text - txtDeductions(0).Text) ' bernard
'    End If
'
'    If WTaxFlag = True And DeducFlag = 3 Then   ' deduc lates
'        Call ComputeTax(txtGrossPay1.Text - txtDeductions(1).Text) ' bernard
'    End If

' Do not include incentive in tax computation always.
'mGrossPay = mGrossPay - TxtAmount(10).Text
'
'If OTNoTax = True Then  ' Ot, Ecola, Incentive, vale non-taxable
'    mGrossPay = mGrossPay - TxtAmount(1).Text - TxtAmount(4).Text ' deduc reg ot and sun ot
'    mGrossPay = mGrossPay - IIf(txtIncentive.Text = "", 0, txtIncentive.Text) ' deduc incentive
'    mGrossPay = mGrossPay - txtVale.Text    ' deduc vale
'   'mGrossPay = mGrossPay - IIf(txtECOLA.Text = "", 0, txtECOLA.Text) ' updated may 5, 2007
'End If

'If ClientTaxPercentageFlag = True Then
''If TaxPercentage <> 0 Then  ' apply the tax percentage computation
'    txtTAX.Text = Format(mGrossPay * (TaxPercentage / 100), "##,#00.00")
'Else
    With rsTax
    '.Filter = "taxcode = '" & Trim(EmpTaxCode) & "'"
    If EmpTaxCode = "" Then
        MsgBox "No Tax Exemption Code entries! Please check HRIS!", vbInformation, PayrollAlert
        txtTax.Text = 0#
        Exit Sub
    Else
        If UseTax2018Flag = False Then ' apply new tax 2018
            .Filter = "taxcode" & " = " & "'" & Trim(EmpTaxCode) & "'"
            If .EOF Then
                MsgBox "Invalid Tax Exemption Code (" & Trim(EmpTaxCode) & ")! Not found in tax table!", vbInformation, PayrollAlert
                txtTax.Text = 0#
                Exit Sub
            End If
        Else
            .Filter = ""
        End If
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
        txtTax.Text = Format(((mGrossPay - mSalRate) * mTaxRate) + mBaseTax, "##,##0.00")
    Else
        txtTax.Text = "0.00"
    End If
'End If
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsEmpDeductions = Nothing
Set rsPreviousEmployee = Nothing
InsertedFlag = False
frmSelectCutoff.Show
Call ProtectDatabaseRecords(CompanyCode, mClientCode, False, True)  ' Unprotect the Employee deduction table for entries
End Sub

' Commented by Bernard dec. 6, 2016
'Private Sub lvwEmployeeList_Click()
'    Call GetBasicInformation
'     'Check if contract already ends
'    'If mEndContract <> #1/1/1900# And mEndContract < frmSelectCutoff.txtCutOffDate(0) Then
'    '    MsgBox EmployeeName & ", already end of contract!", vbInformation, PayrollAlert
'    'End If
'
'    Call ZeroOutAmount
'    If mCtr2 <> 0 Then
'        cmdTransaction(0).Enabled = True
'        cmdTransaction(1).Enabled = False
'        cmdTransaction(0).Caption = "&Create Transaction"
'        ListViewFlag = True
'    Else
'        cmdTransaction(0).Enabled = True
'        cmdTransaction(1).Enabled = True
'        cmdTransaction(0).Caption = "&Edit Transaction"
'        ListViewFlag = False
'    End If
'    ' Populate deductions.
'    Call FillDeductionList(ListDeductions, rsEmpDeductions)
'
'End Sub

Private Sub lvwEmployeeList_DblClick()
Call cmdTransaction_Click(0)
End Sub

Private Sub lvwEmployeeList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ' Get the employee number
    EmployeeName = Trim(Item.Text)
'    Call GetBasicInformation
'     'Check if contract already ends
'    'If mEndContract <> #1/1/1900# And mEndContract < frmSelectCutoff.txtCutOffDate(0) Then
'    '    MsgBox EmployeeName & ", already end of contract!", vbInformation, PayrollAlert
'    'End If
'    cmdTransaction(0).Enabled = True
'    cmdTransaction(1).Enabled = False
'    cmdTransaction(0).Caption = "&Create Transaction"
'    ListViewFlag = True
'    ' Populate deductions.
'    Call FillDeductionList(ListDeductions, rsEmpDeductions)
    
    ' Start of code transfer.
    
    Call GetBasicInformation
    Call ZeroOutAmount
    If mCtr2 <> 0 Then
        cmdTransaction(0).Enabled = True
        cmdTransaction(1).Enabled = False
        cmdTransaction(0).Caption = "&Create Transaction"
        ListViewFlag = True
    Else
        cmdTransaction(0).Enabled = True
        cmdTransaction(1).Enabled = True
        cmdTransaction(0).Caption = "&Edit Transaction"
        ListViewFlag = False
    End If
    ' Populate deductions.
    Call FillDeductionList(ListDeductions, rsEmpDeductions)

    
End Sub

Private Sub lvwProcessEmp_Click()
    If NewCutoffFlag = True Then
        Call GetBasicInformation
        If mCtr1 <> 0 Then
            cmdTransaction(0).Enabled = True
            cmdTransaction(1).Enabled = True
            cmdTransaction(0).Caption = "&Edit Transaction"
            ListViewFlag = False
        Else
            cmdTransaction(0).Enabled = True
            cmdTransaction(1).Enabled = False
            cmdTransaction(0).Caption = "&Create Transaction"
            ListViewFlag = True
        End If
    Else
        Call GetPayrollMasterInfo
        cmdTransaction(0).Enabled = False
        cmdTransaction(1).Enabled = False
        'We need not allow edits here.
        lvwEmployeeList.Enabled = False
        cmdFind.Enabled = True
        cmdNavigation(0).Enabled = False
        cmdNavigation(1).Enabled = False
        cmdNavigation(2).Enabled = False
        cmdNavigation(3).Enabled = False
        cmdPrintReports.Enabled = True
        cmdmovetoMaster.Enabled = False
        cmdClose.Enabled = True
    End If
    Call DisplayAmount
    ' Populate deductions.
    Call FillDeductionList(ListDeductions, rsEmpDeductions)
End Sub

Private Sub lvwProcessEmp_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ' Get the employee number
    EmployeeName = Trim(Item.Text)
    If NewCutoffFlag = True Then
        Call GetBasicInformation
        cmdTransaction(0).Enabled = True
        cmdTransaction(1).Enabled = True
        cmdTransaction(0).Caption = "&Edit Transaction"
    Else
        Call GetPayrollMasterInfo
        cmdTransaction(0).Enabled = False
        cmdTransaction(1).Enabled = False
        'We need not allow edits here.
        lvwEmployeeList.Enabled = False
        cmdFind.Enabled = True
        cmdNavigation(0).Enabled = False
        cmdNavigation(1).Enabled = False
        cmdNavigation(2).Enabled = False
        cmdNavigation(3).Enabled = False
        cmdPrintReports.Enabled = True
        cmdmovetoMaster.Enabled = False
        cmdClose.Enabled = True
    End If
    Call DisplayAmount
    ListViewFlag = False
    ' Populate deductions.
    Call FillDeductionList(ListDeductions, rsEmpDeductions)
End Sub

Private Sub GetPayrollMasterInfo()
Dim PayBasis As String
Dim ContractMonth As String
Dim mRs As Recordset
Dim mRSFlag As Boolean

    mRSFlag = False
    lblEmpName.Caption = Trim(EmployeeName)
    
DisplayInfo:

    If mRSFlag = False Then
        With rsEmployeeInfo
        If .RecordCount = 0 Then
            Exit Sub
        End If
    
        .MoveFirst
        .Find "Fullname = '" & Trim$(EmployeeName) & "'"
        If Not .EOF Then
            EmployeeNo = !EmployeeID
            EmpNo = !EmpNo
            If PaymentBasis = "Hourly" Then
                'PayBasis = "PhP " & Format(!HourlyRate, "###,##0.000") & " (Hourly)"
                PayBasis = Format(!HourlyRate, "###,##0.000") & " (Hourly)"
                lblEntries(0).Caption = "No. of Hours:"
            ElseIf PaymentBasis = "Monthly" Then
                'PayBasis = "PhP " & Format(!MonthlyRate, "###,##0.00") & " (Monthly)"
                PayBasis = Format(!MonthlyRate, "###,##0.00") & " (Monthly)"
                lblEntries(0).Caption = "No. of Days Worked:"
                txtDaysWork.Text = NoofDaysperMonth
            Else
                'PayBasis = IIf(!PaymentBasis = "Daily", "PhP " & Format(!DailyRate, "###,##0.00") & " (Daily)", "PhP " & Format(!MonthlyRate, "###,##0.00") & " (Monthly)")
                PayBasis = IIf(!PaymentBasis = "Daily", Format(!DailyRate, "###,##0.00") & " (Daily)", Format(!MonthlyRate, "###,##0.00") & " (Monthly)")
                lblEntries(0).Caption = "No. of Days Worked:"
            End If
'            HourlyRate = IIf(IsNull(!HourlyRate), 0, !HourlyRate)
'            DailyRate = IIf(IsNull(!DailyRate), 0, !DailyRate)
'            'PaymentBasis = IIf(IsNull(!PaymentBasis) Or !PaymentBasis = "", "Daily", !PaymentBasis)
'            MonthlyRate = IIf(IsNull(!MonthlyRate), 0, !MonthlyRate)
            'ContractMonth = IIf(IsNull(!StartContract), "", Format(!StartContract, "mm/dd/yyyy")) & " to " & IIf(IsNull(!endcontract), "", Format(!endcontract, "mm/dd/yyyy"))
            ContractMonth = IIf(IsNull(!datehired), "", Format(!datehired, "mm/dd/yyyy")) & " Cola Rate: " & Format(mECOLARate, "#0.00")
        Else
            'EmployeeNo = ""
            EmpNo = ""
            PayBasis = ""
            If mRSFlag = False Then
                ' Employee might be transfered in different agency.
                StrQuery = "SELECT UPPER(Surname + ', ' + Firstname + ' ' + Left(Middlename,1) + '.') AS Fullname, * " & _
                    " FROM HREmployees " & _
                    " WHERE UPPER(Surname + ', ' + Firstname + ' ' + Left(Middlename,1) + '.') = '" & EmployeeName & "'"
                Set mRs = ReturnRst(StrQuery, cnnstr, mstatic)
                If mRs.RecordCount <> 0 Then
                    mRSFlag = True
                    GoTo DisplayInfo
                    Exit Sub
                End If
            End If
        End If
        End With
    Else
        With mRs
            If .RecordCount = 0 Then
                Exit Sub
            End If
            .MoveFirst
            .Find "Fullname = '" & Trim$(EmployeeName) & "'"
            If Not .EOF Then
                EmployeeNo = !EmployeeID
                EmpNo = !EmpNo
                If PaymentBasis = "Hourly" Then
                    'PayBasis = "PhP " & Format(!HourlyRate, "###,##0.000") & " (Hourly)"
                    PayBasis = Format(!HourlyRate, "###,##0.000") & " (Hourly)"
                    lblEntries(0).Caption = "No. of Hours:"
                ElseIf PaymentBasis = "Monthly" Then
                    'PayBasis = "PhP " & Format(!MonthlyRate, "###,##0.00") & " (Monthly)"
                    PayBasis = Format(!MonthlyRate, "###,##0.00") & " (Monthly)"
                    lblEntries(0).Caption = "No. of Days Worked:"
                    txtDaysWork.Text = NoofDaysperMonth
                Else
                    'PayBasis = IIf(!PaymentBasis = "Daily", "PhP " & Format(!DailyRate, "###,##0.00") & " (Daily)", "PhP " & Format(!MonthlyRate, "###,##0.00") & " (Monthly)")
                    PayBasis = IIf(!PaymentBasis = "Daily", Format(!DailyRate, "###,##0.00") & " (Daily)", Format(!MonthlyRate, "###,##0.00") & " (Monthly)")
                    lblEntries(0).Caption = "No. of Days Worked:"
                End If
                HourlyRate = IIf(IsNull(!HourlyRate), 0, !HourlyRate)
                DailyRate = IIf(IsNull(!DailyRate), 0, !DailyRate)
                'PaymentBasis = IIf(IsNull(!PaymentBasis) Or !PaymentBasis = "", "Daily", !PaymentBasis)
                MonthlyRate = IIf(IsNull(!MonthlyRate), 0, !MonthlyRate)
                'ContractMonth = IIf(IsNull(!StartContract), "", Format(!StartContract, "mm/dd/yyyy")) & " to " & IIf(IsNull(!endcontract), "", Format(!endcontract, "mm/dd/yyyy"))
                ContractMonth = IIf(IsNull(!datehired), "", Format(!datehired, "mm/dd/yyyy")) & " Cola Rate: " & Format(mECOLARate, "#0.00")
            Else
                EmpNo = ""
                PayBasis = ""
            End If
        End With
        Set mRs = Nothing
    End If
    lblEmpNo.Caption = EmpNo 'EmployeeNo
    lblBasicPay.Caption = PayBasis
    lblContract.Caption = IIf(Trim(ContractMonth) = "to", "Not Specified", ContractMonth)
End Sub

Private Sub GetBasicInformation()     ' per Employee
Dim PayBasis As String
Dim empPosition As String
Dim ContractMonth As String, tt
Dim StartDate, LastDate

'    If RetSw = True Then    ' refresh the recordset
'        strQuery = "SELECT a.*, (trim(b.surname) + ', ' + trim(b.firstname)) AS fullname, " & _
'            "b.sss, b.philhealth, b.pagibig, b.acctno, b.TIN, b.employmentstatus FROM employeebasicsalary AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID " & _
'            "WHERE a.clientcode = " & mclientcode & " and a.employeestatus = true ORDER BY b.surname, b.firstname"
    
    ' Updated by bernard, to exclude end of contract employees. January 27, 2007.
    'If AddFlag = True Or EditFlag = True Then
        StrQuery = "SELECT a.*, " & _
            "b.sss, b.philhealth, b.pagibig, b.acctno, b.TIN, b.employmentstatus, b.DateHired, b.Town AS AllowanceBase, " & _
            "CASE WHEN b.EcolaAmount <= 0 then a.EcolaAmount ELSE b.EcolaAmount END AS EcolaAmount1 " & _
            "FROM qryEmployeeBasics AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID " & _
            "WHERE a.clientcode = " & mClientCode & " and a.employeestatus = 1 and a.payrollstatus not in ('T','W','X','Y','Z') ORDER BY b.surname, b.firstname"
        'Set rsEmployeeInfo = Nothing
        Set rsEmployeeInfo = ReturnRst(StrQuery, cnnstr, mstatic)
    'End If
'    End If
    
    
    'Default caption.
    'lblEntries(0).Caption = "No. of Days Worked:"
    lblEntries(0).Caption = "No. of Hours:"
    lblEntries(1).Caption = "Regular OT Hours:  "
    lblEntries(2).Caption = "Night Diff. Hours: "
    lblEntries(3).Caption = "ECOLA:             "   ' rest day hours
    lblEntries(4).Caption = "Restday Hours:     "   ' excess rest day OT
    lblEntries(31).Caption = "Other Income:      "  ' Salary Adjustment
    If FMSFlag = False Then
        labelLeave(30).Caption = "Allowance:         "  ' Incentive Pay
    Else
        labelLeave(30).Caption = "Allowance:         "  ' Bonus for Driver / Conductor
    End If
    lblEntries(19).Caption = "Tax Refund:        "  ' Ecola Allowance
    
    RetSw = False
    lblEmpName.Caption = Trim(EmployeeName)
    
    With rsEmployeeInfo
    
    If .RecordCount = 0 Then
        Exit Sub
    End If
    .MoveFirst
    .Find "Fullname = '" & EmployeeName & "'"
    If Not .EOF Then
        EmployeeNo = !EmployeeID
        EmpNo = !EmpNo
        empPosition = !EmpNo & " (" & !NotedBy & ")"  ' Edited by Bernard July 12, 2011
        'empPosition = !EmpNo
        If !PaymentBasis = "Hourly" Then
            PayBasis = "PhP " & Format(!HourlyRate, "###,##0.000") & " (Hourly)   " & IIf(IsNull(!EmpTaxCode), "Z", !EmpTaxCode)
            lblEntries(0).Caption = "No. of Hours:"
        ElseIf !PaymentBasis = "Commission" Then
            Label2.Caption = "Commission Base:"
            PayBasis = Format(!MonthlyRate, "#0.00") & " (% Commission Rate) " & IIf(IsNull(!EmpTaxCode), "Z", !EmpTaxCode)
            lblEntries(0).Caption = "Commission:        "
            lblEntries(1).Caption = "Gross Amount:      "
            lblEntries(2).Caption = "Partials:          "
            lblEntries(3).Caption = "Funds/Bonus:       "
            'lblEntries(4).Caption = "Bonus:             "
            lblEntries(4).Caption = "_____________________________________________"   ' excess rest day OT
            lblEntries(31).Caption = "Toll Fee:          "
            labelLeave(30).Caption = "Fuel:              "
            lblEntries(19).Caption = "Miscellaneous:     "
        Else
            PayBasis = IIf(!PaymentBasis = "Daily", "PhP " & Format(!DailyRate, "###,##0.00") & " (Daily)   " & IIf(IsNull(!EmpTaxCode), "Z", !EmpTaxCode), "PhP " & Format(!MonthlyRate, "###,##0.00") & " (Monthly)   " & IIf(IsNull(!EmpTaxCode), "Z", !EmpTaxCode))
            'lblEntries(0).Caption = "No. of Days Worked:"
            lblEntries(0).Caption = "No. of Hours:"
        End If
        HourlyRate = IIf(IsNull(!HourlyRate), 0, !HourlyRate)
        DailyRate = IIf(IsNull(!DailyRate), 0, !DailyRate)
        PaymentBasis = IIf(IsNull(!PaymentBasis) Or !PaymentBasis = "", "Hourly", !PaymentBasis)
        If PaymentBasis <> "Monthly" Then
            PaymentBasis = "Hourly"
            mECOLARate = IIf(WEcolaFlag = True, IIf(IsNull(!EcolaAmount1), 0, !EcolaAmount1), 0#)
        Else
            'mECOLARate = IIf(WEcolaFlag = True, IIf(IsNull(!EcolaAmount), 0, !EcolaAmount), 0#)
            'mECOLARate = (IIf(WEcolaFlag = True, IIf(IsNull(!EcolaAmount1), 0, !EcolaAmount1), 0#) / NoofDaysperMonth)
            mECOLARate = (IIf(WEcolaFlag = True, IIf(IsNull(!EcolaAmount1), 0, !EcolaAmount1), 0#) / 2)
        End If
        'Recompute the rates. Updated July 13, 2016.
        PaymentBasis = IIf(IsNull(!PaymentBasis) Or !PaymentBasis = "", "Hourly", !PaymentBasis)
        If PaymentBasis = "Monthly" Then
            MonthlyRate = IIf(IsNull(!MonthlyRate), 0, !MonthlyRate)
            DailyRate = (MonthlyRate * 12) / WorkDaysperYear 'IIf(IsNull(!DailyRate), 0, !DailyRate)
            HourlyRate = DailyRate / 8 'IIf(IsNull(!HourlyRate), 0, !HourlyRate)
        ElseIf PaymentBasis = "Daily" Then
            DailyRate = IIf(IsNull(!DailyRate), 0, !DailyRate)
            MonthlyRate = (DailyRate * WorkDaysperYear) / 12
            HourlyRate = DailyRate / 8
        Else    ' for Hourly Rate
            HourlyRate = IIf(IsNull(!HourlyRate), 0, !HourlyRate)
            HourlyRate = IIf(IsNull(!HourlyRate), 0, !HourlyRate)
            DailyRate = IIf(IsNull(!DailyRate), 0, !DailyRate)
        End If
        
        
        'MonthlyRate = IIf(IsNull(!MonthlyRate), 0, !MonthlyRate)
        'HourlyRate = IIf(IsNull(!HourlyRate), 0, !HourlyRate)
        'AllowanceBase = IIf(IsNull(!Town), "", !Town)
        AllowanceBasis = IIf(IsNull(!AllowanceBase), "", !AllowanceBase)
        
        ' Get ECOLA Rate, if monthly paid ECOLA is fixed.
        WEcolaFlag = !WithEcolaPay
        
'        If PaymentBasis = "Monthly" Then ' if monthly paid get the fixed Ecola rate
'            mECOLARate = (mEcolaMonth / NoofDaysperMonth) '* IIf(txtDaysWork.Text = "", 0, Val(txtNoDays.Text))
'        Else
'            mECOLARate = mEcolaMonth
'        End If
        
        'ContractMonth = IIf(Trim(!employmentstatus) <> "R", IIf(IsNull(!StartContract), "", Format(!StartContract, "mm/dd/yyyy")) & " to " & IIf(IsNull(!endcontract), "", Format(!endcontract, "mm/dd/yyyy")), "Regular Employee")
        ContractMonth = IIf(IsNull(!datehired), "", Format(!datehired, "mm/dd/yyyy")) & " Cola Rate: " & Format(mECOLARate, "#0.00")
        
        EmpTaxCode = IIf(IsNull(!EmpTaxCode), "Z", !EmpTaxCode)
        mEndContract = IIf(IsNull(!endcontract), #1/1/1900#, !endcontract)
        
        'Get Miscellaneous Fees.
        'BDOCardFee = Format(IIf(IsNull(!BillingRateAgencyBDOCardFee), 0#, !AgencyBDOCardFee), "##,##0.00")
        'BDOCardFee = Format(IIf(IsNull(!AgencyBDOCardFee), 0#, !AgencyBDOCardFee), "##,##0.00")
        'mCBAmount = Format(IIf(IsNull(!AgencyTASAesla), 0#, !AgencyTASAesla), "##,##0.00")
        'EmployeeAgencyRate = Format(IIf(IsNull(!AgencyFee), 0#, !AgencyFee), "0.00")
        
        ' Get SSS and Pagibig Loan and Calamity Loan (previously Vale).
        StartCutoffDate = Val(Format(frmSelectCutoff.txtCutOffDate(0).Text, "yyyymmdd"))
        EndCutoffDate = Val(Format(frmSelectCutoff.txtCutOffDate(1).Text, "yyyymmdd"))
        StartDate = Val(Format(!StartSSSPayment, "yyyymmdd"))
        If StartDate = 19000101 Then StartDate = 99999999
        LastDate = Val(Format(!LastSSSPayment, "yyyymmdd"))
        If LastDate = 19000101 Then LastDate = 99999999
        
        ' Initialized the recordset for deductions.
        tt = LoanDeductionAmount(6, StartCutoffDate, EndCutoffDate, False, 0, True)
    
        ' Check if last SSS payment date elapsed. Commented Feb. 14, 2017
'        If (StartCutoffDate >= StartDate And StartCutoffDate <= LastDate) Or _
'            (EndCutoffDate >= StartDate And EndCutoffDate <= LastDate) Then
'            mSSSLoan = Format(IIf(IsNull(!SSSLoanPayment), 0#, !SSSLoanPayment), "##,##0.00")
'        Else
            mSSSLoan = LoanDeductionAmount(5, StartCutoffDate, EndCutoffDate)
'        End If
        
        ' Check if last Pagibig Payment Date elapsed.
        StartDate = Val(Format(!StartPbigPayment, "yyyymmdd"))
        If StartDate = 19000101 Then StartDate = 99999999
        LastDate = Val(Format(!LastPbigPayment, "yyyymmdd"))
        If LastDate = 19000101 Then LastDate = 99999999
        
        ' Check if Pagibig Payment Date elapsed.
'        If (StartCutoffDate >= StartDate And StartCutoffDate <= LastDate) Or _
'            (EndCutoffDate >= StartDate And EndCutoffDate <= LastDate) Then
'            mPBigLoan = Format(IIf(IsNull(!PbigLoanPayment), 0#, !PbigLoanPayment), "##,##0.00")
'        Else
            mPBigLoan = LoanDeductionAmount(6, StartCutoffDate, EndCutoffDate)
'        End If
        
        ' Check if last Calamity Loan (previously Vale) Payment date elapsed.
        StartDate = Val(Format(!StartValePayment, "yyyymmdd"))
        If StartDate = 19000101 Then StartDate = 99999999
        LastDate = Val(Format(!LastValePayment, "yyyymmdd"))
        If LastDate = 19000101 Then LastDate = 99999999
            
        ' For Cut-off Allowance of employees, Updated May 4, 2016.
        'mAllowance = Format(IIf(IsNull(!ValeLoanPayment), 0#, !ValeLoanPayment), "##,##0.00")
        
'        If (StartCutoffDate >= StartDate And StartCutoffDate <= LastDate) Or _
'            (EndCutoffDate >= StartDate And EndCutoffDate <= LastDate) Then
'            'mAllowance = Format(IIf(IsNull(!ValeLoanPayment), 0#, !ValeLoanPayment), "##,##0.00")
'            'mCalamityLoan = Format(IIf(IsNull(!ValeLoanPayment), 0#, !ValeLoanPayment), "##,##0.00")
'        Else
'            'mCalamityLoan = LoanDeductionAmount(9, StartCutoffDate, EndCutoffDate)
'            ' Surety Fund updated June 14, 2012
            mCalamityLoan = LoanDeductionAmount(12, StartCutoffDate, EndCutoffDate)
'        End If
        ' HDMF Calamity
        mHDMFCalLoan = LoanDeductionAmount(16, StartCutoffDate, EndCutoffDate)
        ' SSS Calamity Loan 9/22/2018 Bernard
        mSSSCalLoan = LoanDeductionAmount(17, StartCutoffDate, EndCutoffDate)
        
        ' Check if Misc Fee date elapsed.
        StartDate = Val(Format(!MFstartpayment, "yyyymmdd"))
        If StartDate = 19000101 Then StartDate = 99999999
        LastDate = Val(Format(!MFendpayment, "yyyymmdd"))
        If LastDate = 19000101 Then LastDate = 99999999
         
'        If (StartCutoffDate >= StartDate And StartCutoffDate <= LastDate) Or _
'            (EndCutoffDate >= StartDate And EndCutoffDate <= LastDate) Then
'           mMFAmount = Format(IIf(IsNull(!MFpayment), 0#, !MFpayment), "##,##0.00")
'        Else
            If lblEntries(34).Caption = "SANAWU:" Then '10
                mMFAmount = LoanDeductionAmount(10, StartCutoffDate, EndCutoffDate)   ' LMC ded
            Else
                mMFAmount = LoanDeductionAmount(11, StartCutoffDate, EndCutoffDate)   ' LMC ded
            End If
'        End If
'New
'1   Uniform
'2   LMC Ded
'3   Cash Advance
'4   Planters Loan
'5   SSS Loan
'6   Pagibig Loan
'7   Vale Cash
'8   Other Deductions
'9   Calamity Loan
'15 sss calamity
'16 hdmf calamity
'17 SSS Calamity

        ' Check if Cash Bond date elapsed.
        StartDate = Val(Format(!CBStartPayment, "yyyymmdd"))
        If StartDate = 19000101 Then StartDate = 99999999
        LastDate = Val(Format(!CBEndPayment, "yyyymmdd"))
        If LastDate = 19000101 Then LastDate = 99999999
        
'        If (StartCutoffDate >= StartDate And StartCutoffDate <= LastDate) Or _
'            (EndCutoffDate >= StartDate And EndCutoffDate <= LastDate) Then
'            mCBAmount = Format(IIf(IsNull(!CBPayment), 0#, !CBPayment), "##,##0.00")
'        Else
            mCBAmount = LoanDeductionAmount(4, StartCutoffDate, EndCutoffDate) ' Planters Loan
'        End If
       
        ' Check if Cash/Advance date elapsed. Bernard 8/7/2010.
        StartDate = Val(Format(!CaStartPayment, "yyyymmdd"))
        If StartDate = 19000101 Then StartDate = 99999999
        LastDate = Val(Format(!CaEndPayment, "yyyymmdd"))
        If LastDate = 19000101 Then LastDate = 99999999
        
'        If (StartCutoffDate >= StartDate And StartCutoffDate <= LastDate) Or _
'            (EndCutoffDate >= StartDate And EndCutoffDate <= LastDate) Then
'            mCashAdvance = Format(IIf(IsNull(!CashAdvance), 0#, !CashAdvance), "##,##0.00")  ' Cash Advance
'        Else
            mCashAdvance = LoanDeductionAmount(3, StartCutoffDate, EndCutoffDate)   ' cash advance
'        End If
       
        ' Check if Uniform date elapsed. Bernard 8/7/2010.
        StartDate = Val(Format(!UNIStartPayment, "yyyymmdd"))
        If StartDate = 19000101 Then StartDate = 99999999
        LastDate = Val(Format(!UNIEndPayment, "yyyymmdd"))
        If LastDate = 19000101 Then LastDate = 99999999
        
'        If (StartCutoffDate >= StartDate And StartCutoffDate <= LastDate) Or _
'            (EndCutoffDate >= StartDate And EndCutoffDate <= LastDate) Then
'            mUniform = Format(IIf(IsNull(!MiscFee), 0#, !MiscFee), "##,##0.00")  ' Uniform
'        Else
            mUniform = LoanDeductionAmount(1, StartCutoffDate, EndCutoffDate)   ' uniform
'        End If
        
        mVale = LoanDeductionAmount(7, StartCutoffDate, EndCutoffDate)   ' Vale
       
        ' Check for Other Deductions.
        StartDate = Val(Format(!PFstartpayment, "yyyymmdd"))
        If StartDate = 19000101 Then StartDate = 99999999
        LastDate = Val(Format(!PFendpayment, "yyyymmdd"))
        If LastDate = 19000101 Then LastDate = 99999999
        
        mBSCAllowRate = Format(IIf(IsNull(!PFAmount), 0#, !PFAmount), "##,##0.00")
        If mBSCAllowRate <> 0 Then ' show the entry for allowance 2
            With txtIncentive
                .Width = 855
                .Height = 285
                .Left = 1200
                .Top = 5460
            End With
            labelLeave(0).Visible = True
            txtSplOtHours.Visible = True
            txtSplOtHours.Text = Format(mBSCAllowRate, "##,##0.00") ' BSC Allowance 2
        Else ' set to default
            With txtIncentive
                .Width = 1215
                .Height = 285
                .Left = 2040
                .Top = 5460
            End With
            labelLeave(0).Visible = False
            txtSplOtHours.Visible = False
        End If
        'AllowancePayRate = Format(IIf(IsNull(!PFpayment), 0#, !PFpayment), "##,##0.00") ' get allowance rate from HR profile
        AllowancePayRate = Format(IIf(Val(!PFPayment) = 0, !ValeLoanPayment, !PFPayment), "##,##0.00")
        
'        If (StartCutoffDate >= StartDate And StartCutoffDate <= LastDate) Or _
'            (EndCutoffDate >= StartDate And EndCutoffDate <= LastDate) Then
'            mPFAmount = Format(IIf(IsNull(!PFpayment), 0#, !PFpayment), "##,##0.00")
'        Else
            mPFAmount = LoanDeductionAmount(8, StartCutoffDate, EndCutoffDate)  ' Other ded
'        End If
        ' Re-Initialized the recordset for deductions.
        tt = LoanDeductionAmount(6, StartCutoffDate, EndCutoffDate, False, 0, True)
       
        'Refresh those with auto-deduction entries.
        txtTASAesla.Text = Format(mPFAmount, "##,##0.00") ' Other deductions
        txtBDOCard.Text = Format(mMFAmount, "##,##0.00") 'LMC
        txtVale.Text = Format(mVale, "##,##0.00") ' VaLe Cash
        txtMiscFee.Text = Format(mCBAmount, "##,##0.00") ' Planters loan
        'txtOverPay.Text = Format(IIf(IsNull(!MiscFee), 0#, !MiscFee), "##,##0.00")  ' Uniform
        'txtPlacementFee.Text = Format(IIf(IsNull(!CashAdvance), 0#, !CashAdvance), "##,##0.00")  ' Cash Advance
        txtOverPay.Text = Format(mUniform, "##,##0.00")  ' Uniform Bernard 8/7/2010
        txtPlacementFee.Text = Format(mCashAdvance, "##,##0.00")  ' Cash Advance Bernard 8/7/2010
        txtSSSLoan.Text = Format(mSSSLoan, "##,##0.00")
        txtPagibigLoan.Text = Format(mPBigLoan, "##,##0.00")
        txtCalamityLoan.Text = Format(mCalamityLoan, "##,##0.00")
        txtHDMFCal.Text = Format(mHDMFCalLoan, "##,##0.00")
        txtCalLoan.Text = Format(mSSSCalLoan, "##,##0.00")
        
        ' For Allowance Computation. Updated April 4, 2016.
        'txtIncentive.Text = Format(mAllowance, "##,##0.00")
        'txtAmount(10) = Format(mAllowance, "##,##0.00")
        
        ' There are cases were employee is not subject for any contributions.
        ' Get SSS Info
        WSSSFlag = !WithSSSDeduction
        WPHealthFlag = !WithPHealthDeduction
        
        ' Updated 2/16/2020 Bernard.
        If !PHealthRateperDay <> 0 Then
            ClientPHRate = !PHealthRateperDay
        Else
            ClientPHRate = 0#
        End If

        ' Get Pagibig info.
        WPagibigFlag = !WithPbigDeduction
        If ClientCutoff = "Weekly" Then ' use the pagibig in employee settings
            mPagibigPay = IIf(WPagibigFlag = True, IIf(IsNull(!PagibigRateperDay), 0#, !PagibigRateperDay), 0#)
        Else
            mPagibigPay = IIf(WPagibigFlag = True, IIf(IsNull(!PagibigRateperDay), 0#, !PagibigRateperDay), 0#)
'            mPagibigPay = IIf(WPagibigFlag = True, IIf(IsNull(!PagibigPayments), 0#, !PagibigPayments), 0#)
        End If
        ' Get Tax info. Updated January 14, 2012
        EmpTaxFlag = !WithTaxDeduction
        'If WTaxFlag = False Then
        '    TaxPercentage = 0#
        'End If
        
        'If OT, Ecola and Allowance is exempted from tax computation. Written April 14, 2006.
        OTNoTax = !BasicSalaryFlag3    ' exclude OT, Ecola from tax
        EcolaInOTRate = !BasicSalaryFlag4
        SSSDedIsMonthly = !BasicSalaryFlag5
        ComputationBase = True '!BasicSalaryFlag6

        'If !PaymentBasis = "Monthly" Then
'        If PaymentBasis = "Monthly" Then
'            MonthlyEcola = mECOLARate * 2
'        Else
'            MonthlyEcola = mECOLARate * NoofDaysperMonth
'        End If
        ' Get 13th month pay.
        W13MonthPay = !With13thMonthPay
        Hourly13thMonthRate = IIf(IsNull(!Hourly13thMonthRate), 0#, !Hourly13thMonthRate)
        If W13MonthPay = False Then
            lblEntries(29).Visible = False
            txt13thMonth.Visible = False
            lblEntries(24).Visible = False
            TxtAmount(11).Visible = False
        Else
            lblEntries(29).Visible = True
            txt13thMonth.Visible = True
            lblEntries(24).Visible = True
            TxtAmount(11).Visible = True
        End If
    Else
        'EmployeeNo = ""
        EmpNo = ""
        PayBasis = ""
    End If
    End With
    lblEmpNo.Caption = empPosition 'EmployeeNo
    lblBasicPay.Caption = PayBasis
    lblContract.Caption = IIf(Trim(ContractMonth) = "to", "Not Specified", ContractMonth)
End Sub

Private Function LoanDeductionAmount(ByVal DedCode As Integer, ByVal StartCutoff As Long, _
    ByVal EndCutoff As Long, Optional GetDedFlag As Boolean = False, _
    Optional mPayments As Double = 0, Optional QuerySw As Boolean = False) As Double

Dim o As Integer
Dim mRem As String
Dim mActualAmt As Double

LoanDeductionAmount = 0#
If QuerySw = True Then  ' Refresh the recordset of deductions
'    If rsLoanDed.State = 1 Then
'        rsLoanDed.Close
'    End If
    StrQuery = "SELECT DeductionPayment, DeductionCode FROM EmployeeDedSchedules " & _
        " WHERE CONVERT(varchar(10), DedStartpayment, 112) >= " & StartCutoff & " AND " & _
        " CONVERT(varchar(10), DedEndPayment, 112) <= " & EndCutoff & _
        " AND DeductionStatus = 1 AND EmployeeID = " & EmployeeNo
    Set rsLoanDed = ReturnRst(StrQuery, cnnstr, mstatic)
ElseIf GetDedFlag = False Then  ' get the deductions if found
    If DedCode = 8 Then ' for other deductions updated July 28, 2012
        ' Get the sum of query.
        StrQuery = "SELECT ClientCode, EmployeeID, DeductionCode, SUM(DeductionPayment) AS DeductionPayment FROM EmployeeDedSchedules " & _
            " WHERE CONVERT(varchar(10), DedStartpayment, 112) >= " & StartCutoff & " AND " & _
            " CONVERT(varchar(10), DedEndPayment, 112) <= " & EndCutoff & _
            " AND DeductionStatus = 1 AND EmployeeID = " & EmployeeNo & _
            " GROUP BY ClientCode, EmployeeID, DeductionCode "
        Set rsLoanDed = ReturnRst(StrQuery, cnnstr, mstatic)
    End If
    
    With rsLoanDed
        .Filter = ""
        .Filter = " DeductionCode = " & DedCode
        If rsLoanDed.RecordCount <> 0 Then
            LoanDeductionAmount = IIf(IsNull(rsLoanDed!DeductionPayment), 0, rsLoanDed!DeductionPayment)
            arrLoanPayments(DedCode) = LoanDeductionAmount
        End If
    End With
'    StrQuery = "SELECT TOP 1 DeductionPayment FROM EmployeeDedSchedules " & _
'        " WHERE CONVERT(varchar(10), DedStartpayment, 112) >= " & StartCutoff & " AND " & _
'        " CONVERT(varchar(10), DedEndPayment, 112) <= " & EndCutoff & _
'        " AND DeductionStatus = 1 AND DeductionCode = " & DedCode & " AND EmployeeID = " & EmployeeNo
'    Set rsLoanDed = ReturnRst(StrQuery, cnnstr, mstatic)

ElseIf GetDedFlag = True And mPayments = 0 Then   ' Update the Deductions table with payments.
    For o = 1 To 17 ' 16 add SSS Calamity Loan 10/21/2018
        If o = 8 Then ' for Deductions
            'mRem = " LTRIM(RTRIM(Remarks)) + ' Pd " & Format(Now, "mm/dd/yyyy") & "'"
            mRem = " LTRIM(RTRIM(VoucherNo))" ' + ' Reset to 0. " & Format(Now, "mm/dd/yyyy") & "'"
        Else
            mRem = "'Paid. " & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'"
        End If
    
        If arrLoanPayments(o) <> 0 Then
            'Save original payment, January 21, 2012.
            mActualAmt = arrLoanPayments(o)
            Select Case o
                Case 1: ' Uniform
                    arrLoanPayments(o) = txtOverPay.Text
                Case 2, 10, 11: 'LMC Ded, SANAWU, Association
                    arrLoanPayments(o) = txtBDOCard.Text
                Case 3: 'Cash Advance
                    arrLoanPayments(o) = txtPlacementFee.Text
                Case 4, 14: 'Planters Loan / Shortages
                    arrLoanPayments(o) = txtMiscFee.Text
                Case 5: 'SSS Loan
                    arrLoanPayments(o) = txtSSSLoan.Text
                Case 6: 'Pagibig loan
                    arrLoanPayments(o) = txtPagibigLoan.Text
                Case 7: 'Vale Cash
                    arrLoanPayments(o) = txtVale.Text
                Case 8, 13: ' Other Deductions / Damages
                    arrLoanPayments(o) = txtTASAesla.Text
                Case 12: 'Calamity Loan 9
                    arrLoanPayments(o) = txtCalamityLoan.Text
                Case 16: 'HDMF Calamity Loan
                    arrLoanPayments(o) = txtHDMFCal.Text
                Case 17: 'SSS Calamity Loan
                    arrLoanPayments(o) = txtCalLoan.Text
            End Select
            
'10  SANAWU SANAWU
'11  Association Association
'12  Surety Fund SuretyFund
'13  Damages/Incident    Damages
'14  Shortages Shortages
'16  HDMF Calamity Loan

            StrQuery = "UPDATE EmployeeDedSchedules SET PaymentsMade = " & arrLoanPayments(o) & ", PaymentDate = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "', " & _
                " Remarks = " & mRem & _
            " WHERE CONVERT(varchar(10), DedStartpayment, 112) >= " & StartCutoff & " AND " & _
            " CONVERT(varchar(10), DedEndPayment, 112) <= " & EndCutoff & _
            " AND DeductionStatus = 1 AND DeductionCode = " & o & " AND EmployeeID = " & EmployeeNo
            Call Action_Query(StrQuery, cnnstr)
        End If
    Next o
ElseIf GetDedFlag = True And mPayments = 1 Then   ' need to reset to 0 all payments save in deduction table.
    For o = 1 To 10
        If o = 8 Then ' Other Deductions
            'mRem = " LTRIM(RTRIM(Remarks)) "
            mRem = " LTRIM(RTRIM(VoucherNo)) " ' + ' Deleted.'"
        Else
            mRem = " 'Deleted back to 0. " & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'"
        End If
        arrLoanPayments(o) = 0#
        StrQuery = "UPDATE EmployeeDedSchedules SET PaymentsMade = " & arrLoanPayments(o) & ", PaymentDate = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "', " & _
                " Remarks = " & mRem & _
            " WHERE CONVERT(varchar(10), DedStartpayment, 112) >= " & StartCutoff & " AND " & _
            " CONVERT(varchar(10), DedEndPayment, 112) <= " & EndCutoff & _
            " AND DeductionStatus = 1 AND DeductionCode = " & o & " AND EmployeeID = " & EmployeeNo
        Call Action_Query(StrQuery, cnnstr)
    Next o
End If

End Function

Private Sub ReComputePayroll()
txtNoDays.Refresh
txtDaysWork.Refresh
txtRegOT.Refresh
txtNightDiff.Refresh
txtECOLA.Refresh
txtSplHoliday.Refresh
txtLglHoliday.Refresh
txtSplOtHours.Refresh
txtLglHolidayOT.Refresh
txtSickLeave.Refresh
txtVacLeave.Refresh
txtSalAdjustment.Refresh
txtAdjustments.Refresh
txtIncentive.Refresh

If WithHazardPay = True And EditFlag = False Then
    txtUnionDues.Refresh
    txtHazardPay.Refresh
End If

txt13thMonth.Enabled = False
txtDaysAbsent.Refresh
txtHoursLate.Refresh
txtTASAesla.Refresh
txtVale.Refresh
txtMiscFee.Refresh
txtBDOCard.Refresh
txtPlacementFee.Refresh
txtOverPay.Refresh
txtExcessOTRegDay.Refresh
txtExcessOTSplSun.Refresh
txtExcessOTSunLgl.Refresh
txtExcessOTLegal.Refresh
'txtOTwoPremium.Refresh

End Sub

Private Sub DisableTextBoxes(ByVal TFalse As Boolean)
    chkEcolaHol.Enabled = TFalse
    chkEcolaLegal.Enabled = TFalse
    
    txtDaysWork.Enabled = TFalse
    txtNoDays.Enabled = TFalse
    txtRegOT.Enabled = TFalse
    txtNightDiff.Enabled = TFalse
    txtTaxRefund.Enabled = TFalse
    txtSundayOT.Enabled = TFalse
    txtSplHoliday.Enabled = TFalse
    txtLglHoliday.Enabled = TFalse
    txtSplOtHours.Enabled = TFalse
    txtLglHolidayOT.Enabled = TFalse
    txtSickLeave.Enabled = TFalse
    txtVacLeave.Enabled = TFalse
    txtSalAdjustment.Enabled = TFalse
    txtAdjustments.Enabled = TFalse ' Updated Feb. 5, 2015
    txtCalLoan.Enabled = TFalse ' Updated 9/22/2018
    
    txtIncentive.Enabled = TFalse
    txt13thMonth.Enabled = False
    txtECOLA.Enabled = TFalse
    txtDaysAbsent.Enabled = TFalse
    txtHoursLate.Enabled = TFalse
    txtTASAesla.Enabled = TFalse
    txtVale.Enabled = TFalse
    txtMiscFee.Enabled = TFalse
    txtBDOCard.Enabled = TFalse
    txtPlacementFee.Enabled = TFalse
    txtOverPay.Enabled = TFalse
    txtExcessOTRegDay.Enabled = TFalse
    txtExcessOTSplSun.Enabled = TFalse
    txtExcessOTSunLgl.Enabled = TFalse
    txtExcessOTLegal.Enabled = TFalse
    'txtOTwoPremium.Enabled = TFalse
    If WithHazardPay = True Then
        txtHazardPay.Enabled = TFalse
        txtUnionDues.Enabled = TFalse
    End If
    ' Updated Jan 17, 2015.
    txtRDSplLgl.Enabled = TFalse
    txtRDonSPL.Enabled = TFalse
    txtRDonLGL.Enabled = TFalse
    
    txtUnpaidVL.Enabled = TFalse
    txtUnpaidSL.Enabled = TFalse
    txtEcolaAdj.Enabled = TFalse
    txtHazardAdj.Enabled = TFalse
    txtOtherAdj.Enabled = TFalse
    
End Sub

Private Sub ZeroOutAmount()
Dim i As Integer

    txtNoDays.Text = ""
    txtDaysWork.Text = ""
    txtRegOT.Text = ""
    txtNightDiff.Text = ""
    txtTaxRefund.Text = ""
    txtSundayOT.Text = ""
    txtSplHoliday.Text = ""
    txtLglHoliday.Text = ""
    'txtSplOTHours.Text = ""
    'txtLglHolidayOT.Text = ""
    txtSickLeave.Text = ""
    txtVacLeave.Text = ""
    txtSalAdjustment.Text = ""
    txtAdjustments.Text = ""
    txtIncentive.Text = ""
    txtDaysAbsent.Text = ""
    txtHoursLate.Text = ""
    ' Updated Jan 17, 2015
    txtRDSplLgl.Text = ""
    txtRDonSPL.Text = ""
    txtRDonLGL.Text = ""
    txtUnpaidVL.Text = ""
    txtUnpaidSL.Text = ""
    txtEcolaAdj.Text = ""
    txtHazardAdj.Text = ""
    txtOtherAdj.Text = ""
    
    'txtBDOCard.Text = "0.00" ' Format(BDOCardFee, "##,##0.00") ' orig
    
    ' Get the entries from client.
'    txtTASAesla.Text = Format(mCBAmount, "##,##0.00")
'    txtMiscFee.Text = Format(mMFAmount, "##,##0.00")
'    txtOverPay.Text = Format(BDOCardFee, "##,##0.00")
'    txtPlacementFee.Text = Format(mPFAmount, "##,##0.00")
'    txtSSSLoan.Text = Format(mSSSLoan, "##,##0.00")
'    txtPagibigLoan.Text = Format(mPBigLoan, "##,##0.00")
'    txtCalamityLoan.Text = Format(mCalamityLoan, "##,##0.00")
    
    txtExcessOTRegDay.Text = ""
    txtExcessOTSplSun.Text = ""
    txtExcessOTSunLgl.Text = ""
    txtExcessOTLegal.Text = ""
'    txtOTwoPremium.Text = ""
    txt13thMonth.Text = ""
    txtECOLA.Text = "" 'Format(mECOLARate, "0.00")
    
    txtEmpyeeSSS.Text = "0"
    txtEmpyerSSS.Text = "0"
    txtEmpyeePagibig.Text = "0"
    txtEmpyerPagibig.Text = "0"
    txtPhilHealthYee.Text = "0"
    txtPHealthYer.Text = "0"
    txtTax.Text = "0"
   
    For i = 0 To 21 '21 include adjustment feb. 5, 2015
        TxtAmount(i).Text = "0.00"
        TxtAmount(i).BackColor = &HC0C0C0
    Next
    For i = 0 To 1
        txtDeductions(i).Text = "0.00"
        txtDeductions(i).BackColor = &HC0C0C0
    Next
    For i = 0 To 3  ' for payroll adjustments
        txtAdj(i).Text = "0.00"
        txtAdj(i).BackColor = &HC0C0C0
    Next
    txtGrossPay.Text = "0.00"
    txtGrossPay1.Text = "0.00"
    txtNetPay.Text = "0.00"
    txtDeductions(2).Text = "0.00"
End Sub

Private Sub DisplayAmount()
Dim i As Integer
    
    With rsPayrollTrans
    If .RecordCount <> 0 Then
        .MoveFirst
        .Find "EmployeeID = " & EmployeeNo
        If Not .EOF Then    ' employee number was located
            For i = 0 To 21 '21
                TxtAmount(i).BackColor = &HFFFFFF
            Next
            For i = 0 To 1
                txtDeductions(i).BackColor = &HFFFFFF
            Next
            For i = 0 To 3  ' for adjustments
                txtAdj(i).BackColor = &HFFFFFF
            Next
            txtNoDays.Text = Format(!noofdays, "0.00")  ' Updated July 14, 2012
            txtDaysWork.Text = Format(!DaysPresent, "0.00")
            txtRegOT.Text = Format(!OTHours, "0.00")
            txtNightDiff.Text = Format(!NightDiffOT, "0.00")
            txtTaxRefund.Text = Format(!TaxRefundHours, "0.00")
            txtSundayOT.Text = Format(!SundayOT, "0.00")
            txtSplHoliday.Text = Format(!NoRegHoliday, "0.00")
            txtLglHoliday.Text = Format(!NoSplHoliday, "0.00")
            'txtLglHolidayOT.Text = Format(!lglholidayothours, "0.00")
            txtSickLeave.Text = Format(!NoSickLeave, "0.00")
            txtVacLeave.Text = Format(!NoVacationLeave, "0.00")
            txtSalAdjustment.Text = Format(!OtherIncome, "##,##0.00")
            txtAdjustments.Text = Format(!Variances, "##,##0.00")
            txtCalLoan.Text = Format(!Medical, "##,##0.00") ' updated 9/22/2018 Bernard
            
            'BSC Incentive 9/9/2017
            txtSplOtHours.Text = Format(!splholidayothours, "0.00")
            TxtAmount(10).Text = Format(!Allowance, "##,##0.00")
            txtIncentive.Text = Format(!OtherLoan, "##,##0.00")
            
            txtDaysAbsent.Text = Format(!DaysAbsent, "0.00")
            txtHoursLate.Text = Format(!HoursTardiness, "##,##0.00")
            
            'Automated Deductions.
            txtTASAesla.Text = Format(!TASAesla, "##,##0.00")
            txtVale.Text = Format(!ValeWithClient, "##,##0.00")
            'txtMiscFee.Text = Format(!MiscFee, "##,##0.00")
            txtMiscFee.Text = Format(!Vale_IOU, "##,##0.00") ' updated to Planters loan 1/22/2011
            txtBDOCard.Text = Format(!LMCDed, "##,##0.00") ' by bernard
            txtPlacementFee.Text = Format(!CashAdvance, "##,##0.00")
            'txtOverPay.Text = Format(!OverUnderPay, "##,##0.00") orig
            txtOverPay.Text = Format(!Uniform, "##,##0.00")
            
            txtExcessOTRegDay.Text = Format(!excessothours, "##,##0.00")
            txtExcessOTSplSun.Text = Format(!excesssplholothours, "##,##0.00")
            txtExcessOTLegal.Text = Format(!excesslglholothours, "##,##0.00")
            txtExcessOTSunLgl.Text = Format(!excessotsunlegalhrs, "##,##0.00")
            'txtOTwoPremium.Text = Format(!otwithoutpremium, "##,##0.00")
            txtECOLA.Text = Format(!EcolaPay, "##,##0.00")
            'txtIncentive.Text = Format(!Allowance, "##,##0.00") ' comment by bernard 9/9/2017
            txt13thMonth.Text = Format(!month13rate, "##,##0.00")

            txtEmpyeeSSS.Text = Format(!SSSEmployee, "##,##0.00")
            txtEmpyerSSS.Text = Format(!SSSEmployer, "##,##0.00")
            txtEmpyeePagibig.Text = Format(!pagibigemployee, "##,##0.00")
            txtEmpyerPagibig.Text = Format(!PagibigEmployer, "##,##0.00")
            txtPhilHealthYee.Text = Format(!PHealthEmployee, "##,##0.00")
            txtPHealthYer.Text = Format(!PHealthEmployer, "##,##0.00")
            txtSSSLoan.Text = Format(!SSSLoan, "##,##0.00")
            txtPagibigLoan.Text = Format(!PagibigLoan, "##,##0.00")
            txtTax.Text = Format(!WTaxAmount, "##,##0.00")
            txtCalamityLoan.Text = Format(!CalamityLoan, "##,##0.00")
        
'            'Get the entries from client.
'            txtTASAesla.Text = Format(mCBAmount, "##,##0.00")
'            txtMiscFee.Text = Format(mMFAmount, "##,##0.00")
'            txtOverPay.Text = Format(BDOCardFee, "##,##0.00")
'            txtPlacementFee.Text = Format(mPFAmount, "##,##0.00")
'            txtSSSLoan.Text = Format(mSSSLoan, "##,##0.00")
'            txtPagibigLoan.Text = Format(mPBigLoan, "##,##0.00")
'            txtCalamityLoan.Text = Format(mCalamityLoan, "##,##0.00")
        
            txtGrossPay.Text = Format(!GrossPay, "##,##0.00")
            txtGrossPay1.Text = Format(!GrossPay, "##,##0.00")
            txtNetPay.Text = Format(!NetPay, "##,##0.00")
            txtDeductions(2).Text = Format(!Deductions, "##,##0.00")
            
            TxtAmount(0).Text = Format(!BasicPay, "###,##0.00")
            TxtAmount(1).Text = Format(!OTPay, "###,##0.00")
            TxtAmount(2).Text = Format(!NightDiffPay, "###,##0.00")
            TxtAmount(3).Text = Format(!TaxRefundPay, "##,##0.00")
            TxtAmount(4).Text = Format(!SundayOTPay, "##,##0.00")
            TxtAmount(5).Text = Format(!RegHolidayPay, "##,##0.00")
            TxtAmount(6).Text = Format(!SplHolidayPay, "##,##0.00")
            'txtAmount(7).Text = Format(!splholidayotpay, "##,##0.00")
            'txtAmount(8).Text = Format(!lglholidayotpay, "##,##0.00")
            TxtAmount(7).Text = Format(!SickLeavePay, "##,##0.00")
            TxtAmount(8).Text = Format(!VacationLeavePay, "##,##0.00")
            TxtAmount(9).Text = Format(!OtherIncome, "##,##0.00")
            TxtAmount(21).Text = Format(!Variances, "##,##0.00")
            
            TxtAmount(10).Text = Format(!Allowance, "##,##0.00")
            
            ' 13th month pay
            TxtAmount(11).Text = Format(!Month13Payment, "##,##0.00")
            
            ' for HDMF Calamity
            txtHDMFCal.Text = Format(!ExcessLglHolOTPay, "##,##0.00")
            ' for SSS Calamity Loan 9/22/2018
            txtCalLoan.Text = Format(!Medical, "##,##0.00")
            
            TxtAmount(12).Text = Format(!excessotpay, "##,##0.00")
            'txtAmount(13).Text = Format(!excesssplholotpay, "##,##0.00")
            
            TxtAmount(14).Text = Format(!excessotsunlegalpay, "##,##0.00")
            'txtAmount(15).Text = Format(!excesslglholotpay, "##,##0.00")
            'TxtAmount(16).Text = Format(!otwithoutpremiumpay, "##,##0.00")
            
            ' UPdated Jan 17, 2015. Updated again 3/12/2017 use the excess OT
            'txtRDSplLgl.Text = Format(!excessotsunlegalhrs, "##,##0.00")
            txtExcessOTRegDay.Text = Format(!excessotsunlegalhrs, "##,##0.00")
            'TxtAmount(18).Text = Format(!excessotsunlegalpay, "##,##0.00")
            TxtAmount(12).Text = Format(!excessotsunlegalpay, "##,##0.00")
            
            txtRDonSPL.Text = Format(!excesssplholothours, "##,##0.00")
            TxtAmount(19).Text = Format(!excesssplholotpay, "##,##0.00")
            txtRDonLGL.Text = Format(!otwithoutpremium, "##,##0.00")
            TxtAmount(20).Text = Format(!otwithoutpremiumpay, "##,##0.00")
            
            ' Adjustments Updated Dec. 3, 2016
            txtAdj(0).Text = Format(!unpaidvl, "##,##0.00")
            txtUnpaidVL.Text = Format(!unpaidvldays, "##,##0.00")
            txtAdj(1).Text = Format(!unpaidsl, "##,##0.00")
            txtUnpaidSL.Text = Format(!unpaidsldays, "##,##0.00")
            txtAdj(2).Text = Format(!ecolaadj, "##,##0.00")
            txtEcolaAdj.Text = Format(!ecolaadj, "##,##0.00")
            txtAdj(3).Text = Format(!hazardadj, "##,##0.00")
            txtHazardAdj.Text = Format(!hazardadj, "##,##0.00")
            txtAdj(4).Text = Format(!RegRate, "##,##0.00")
            txtOtherAdj.Text = Format(!RegRate, "##,##0.00")
           
            
    ' Updates Jan 17, 2015
    ' excessotsunlegalhrs = " & ChkVal(txtRDSplLgl) & ", excessotsunlegalpay = " & ChkVal(txtAmount(18))
    ' excesssplholothours = " & ChkVal(txtRDonSPL) & ", excesssplholotpay = " & ChkVal(txtAmount(19))
    ' otwithoutpremium = " & ChkVal(txtRDonLGL) & ", otwithoutpremiumpay = " & ChkVal(txtAmount(20))
            
            TxtAmount(17).Text = Format(!EcolaPay, "##,##0.00")
            If WithHazardPay = True Then
                txtHazardPay.Text = Format(!HazardPay, "##,##0.00")
                txtUnionDues.Text = Format(!UnionDues, "##,##0.00")
            End If
            txtDeductions(0).Text = Format(!AbsentDeduction, "##,##0.00")
            txtDeductions(1).Text = Format(!TardinessDeduction, "##,##0.00")
            
            txtProcessBy.Text = "By: " & IIf(IsNull(!Entryby), "Unknown", !Entryby)
            txtProcessDate.Text = "Date: " & IIf(IsNull(!EntryDate), "__/__/____", Format(!EntryDate, "mm/dd/yyyy"))
    End If
    End If
    End With
End Sub

Private Sub DisableAllButtons()
    lvwEmployeeList.Enabled = False
    lvwProcessEmp.Enabled = False
    'cmdSaveTrans(4).Visible = False
    'cmdCancelTrans(5).Visible = False
    cmdTransaction(0).Enabled = False
    cmdTransaction(1).Enabled = False
'    cmdTransaction(2).Enabled = False
    cmdFind.Enabled = False
    cmdNavigation(0).Enabled = False
    cmdNavigation(1).Enabled = False
    cmdNavigation(2).Enabled = False
    cmdNavigation(3).Enabled = False
    cmdPrintReports.Enabled = False
    cmdmovetoMaster.Enabled = False
    cmdClose.Enabled = True
End Sub

Private Sub DisableCmdButtons()
    If AddFlag = True Or EditFlag = True Then
        Call DisableTextBoxes(True)
        lvwEmployeeList.Enabled = False
        lvwProcessEmp.Enabled = False
        cmdFixedRates.Enabled = False
        'cmdExcessOt(4).Enabled = False
        cmdRefreshRecords.Enabled = False
        'cmdSaveTrans(4).Visible = True
        'cmdCancelTrans(5).Visible = True
        cmdTransaction(0).Enabled = True
        cmdTransaction(1).Enabled = True
        'cmdTransaction(2).Enabled = False
        cmdFind.Enabled = False
        cmdNavigation(0).Enabled = False
        cmdNavigation(1).Enabled = False
        cmdNavigation(2).Enabled = False
        cmdNavigation(3).Enabled = False
        cmdPrintReports.Enabled = False
        cmdmovetoMaster.Enabled = False
        cmdClose.Enabled = False
    Else
        Call DisableTextBoxes(False)
        lvwEmployeeList.Enabled = True
        lvwProcessEmp.Enabled = True
        cmdFixedRates.Enabled = True
        'cmdExcessOt(4).Enabled = True
        cmdRefreshRecords.Enabled = True
        'cmdSaveTrans(4).Visible = False
        'cmdCancelTrans(5).Visible = False
        cmdTransaction(0).Enabled = True
        cmdTransaction(1).Enabled = True
        'cmdTransaction(2).Enabled = True
        cmdFind.Enabled = True
        cmdNavigation(0).Enabled = True
        cmdNavigation(1).Enabled = True
        cmdNavigation(2).Enabled = True
        cmdNavigation(3).Enabled = True
        cmdPrintReports.Enabled = True
        cmdmovetoMaster.Enabled = True
        cmdClose.Enabled = True
    End If
End Sub

Private Sub txt13thMonth_GotFocus()
    Sendkeys "{HOME}+{END}"
    txt13thMonth.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txt13thMonth
End Sub

Private Sub txt13thMonth_LostFocus()
Dim i As Integer
Dim strAmt As String
Dim mAmt As String
Dim DeciSw As Boolean

    txt13thMonth.BackColor = &HFFFFFF
    mAmt = Trim(txt13thMonth.Text)
    For i = 1 To Len(mAmt)
        If DeciSw = True And mID(mAmt, i, 1) = "." Then
        Else
            strAmt = strAmt & mID(mAmt, i, 1)
        End If
        If DeciSw = False And mID(mAmt, i, 1) = "." Then
            DeciSw = True
        End If
    Next
    mAmt = IIf(strAmt = ".", 0#, Format(strAmt, "0.00"))
    TxtAmount(11).BackColor = &HFFFFFF
    TxtAmount(11).Text = Format(mAmt, "##,##0.00")
    Call AccumulateAmount
    txt13thMonth.Text = Format(mAmt, "##,##0.00")
End Sub

Private Sub AccumulateAmount(Optional TaxSw As Boolean = False, _
    Optional DeducSw As Boolean = False, Optional Recompute As Boolean = False)
Dim i As Integer
Dim TotValue As Double
Dim TaxableAmt As Double
Dim TotSSS As Double, Tot_OT As Double

    Tot_OT = 0: TotValue = 0
    ' Refresh Ecola and Hazard pay
    'If HazardFlag = True Then  ' commented by Bernard 2/1/2014
        'Call ReComputeHazardEcola ' comment by bernard
    'End If
    'Accumulate Total Payments.
    For i = 0 To 21 ' previously 21
        If Display13thPay = False And i = 11 Then   ' do not include
        Else
            'TotValue = TotValue + IIf(Trim(TxtAmount(i)) = "." Or TxtAmount(i).Text = "", 0, TxtAmount(i).Text)
            'Updated Feb. 1, 2015
            TotValue = TotValue + ChkVal(TxtAmount(i))
        End If
        TxtAmount(i).Refresh
    Next
    
    'Add Hazard Pay.
    TotValue = TotValue + IIf(WithHazardPay = True, Val(txtHazardPay.Text), 0)
    
    TaxableAmt = TotValue - (Val(Replace(txtDeductions(0).Text, ",", "")) + Val(Replace(txtDeductions(1).Text, ",", "")))
    If WTaxableEcola = False Then ' ecola is non-taxable
        'TaxableAmt = TaxableAmt - IIf(Trim(TxtAmount(17)) = "." Or TxtAmount(17).Text = "", 0, TxtAmount(17).Text)
        ' Updated Feb. 1, 2015
        TaxableAmt = TaxableAmt - ChkVal(TxtAmount(17))
    'ElseIf Val(Replace(txtDeductions(1).Text, ",", "")) <> 0# And WTaxableEcola = False Then
    ElseIf ChkVal(txtDeductions(1)) <> 0# And WTaxableEcola = False Then
        'TaxableAmt = TaxableAmt - IIf(Trim(TxtAmount(17)) = "." Or TxtAmount(17).Text = "", 0, TxtAmount(17).Text)
        TaxableAmt = TaxableAmt - ChkVal(TxtAmount(17))
    End If
    
    ' Allowance is Taxable updated Feb. 7, 2015.
    If WTaxableAllowance = False Then ' allowance is non-taxable
        TaxableAmt = TaxableAmt - ChkVal(TxtAmount(10))
    End If
'    ' Check if Other Income is Taxable, Updated Feb. 16, 2015
'    If EcolaPayonRestDay = False Then ' other income is non-taxable
        TaxableAmt = TaxableAmt - ChkVal(TxtAmount(9))
'    End If
    
    If mExcludedTax = True Then 'sss, pagibig, phealth is excluded from tax computation
        'TaxableAmt = TaxableAmt - (Val(txtEmpyeeSSS.Text) + Val(txtEmpyeePagibig.Text) + Val(txtPhilHealthYee.Text))
        'TotSSS = GetPreviousSSS(1) + GetPreviousSSS(2) + GetPreviousSSS(3) + (Val(txtEmpyeeSSS.Text) + Val(txtEmpyeePagibig.Text) + Val(txtPhilHealthYee.Text))
        TotSSS = (Val(txtEmpyeeSSS.Text) + Val(txtEmpyeePagibig.Text) + Val(txtPhilHealthYee.Text))
        'TaxableAmt = MonthlyRate - TotSSS
        TaxableAmt = TaxableAmt - TotSSS ' Updated January 14, 2012
    End If
    
    If mOTExcluded = True Then  ' exclude Reg OT and Other Income from tax Feb. 5, 2015
        Tot_OT = ChkVal(TxtAmount(1)) + ChkVal(TxtAmount(9)) '+ Val(TxtAmount(5).Text) + Val(TxtAmount(6).Text))
        'Tot_OT = Tot_OT + (Val(TxtAmount(18).Text) + Val(TxtAmount(19).Text) + Val(TxtAmount(20).Text))
        TaxableAmt = TaxableAmt - Tot_OT ' Updated January 17, 2015
    End If
    txtGrossPay1.Text = Format(TotValue - (Val(Replace(txtDeductions(1).Text, ",", "")) + Val(Replace(txtDeductions(0).Text, ",", ""))), "###,##0.00")
    txtGrossPay1.Refresh
    txtGrossPay.Text = Format(TotValue - (Val(Replace(txtDeductions(1).Text, ",", "")) + Val(Replace(txtDeductions(0).Text, ",", ""))), "###,##0.00")
    txtGrossPay.Refresh
    
'    lblEntries(40).Caption = "WH TAX:"
    
If Val(Format(frmSelectCutoff.txtCutOffDate(1).Text, "dd")) < 17 Then
    WithSSSDedFlag = True
    WithPHDedFlag = True   ' flag whether philhealth deduction will apply
    WithPBIGDedFlag = False  ' flag whether pagibig deduction will apply
    WTaxFlag = True         ' updated Jan. 12, 2012, to reflect tax for every payroll
    mCutOffNo = 1
Else
    WithSSSDedFlag = False
    WithPHDedFlag = False   ' flag whether philhealth deduction will apply
    WithPBIGDedFlag = True  ' flag whether pagibig deduction will apply
    WTaxFlag = True  ' flag whether tax deduction will apply
    mCutOffNo = 2
End If
    
    If ClientTaxPercentageFlag = True And DailyRate <= TaxPercentage Then
        WTaxFlag = False
'        lblEntries(40).Caption = "Not Taxable:"
'        If MsgFlag = False Then
'            MsgBox "Daily rate " & DailyRate & " is less than or equal to Tax Bracket (" & TaxPercentage & ")." & Chr(13) & _
'                "No tax deduction will apply.", vbInformation, PayrollAlert
'            MsgFlag = True
'        End If
    ElseIf EmpTaxFlag = False Then
        WTaxFlag = False
    End If
    
    'Compute for tax.
    If WTaxFlag = True Then
        ' Do not include Tax refund in Tax computation. Updated March 2, 2013
        'TaxableAmt = TaxableAmt - Val(Replace(TxtAmount(3).Text, ",", ""))
        TaxableAmt = TaxableAmt - ChkVal(TxtAmount(3))
    '    If TaxSw = False Then
        If PaymentBasis = "Monthly" Then ' use monthly tax table
            'TaxableAmt = TaxableAmt - Val(Replace(txtDeductions(1).Text, ",", ""))
            Call ComputeTax(TaxableAmt, rsTaxTable) ' rstaxtable
        Else
            Call ComputeTax(TaxableAmt, rsTaxTable) ' taxtable2
        End If
    '        Call AccumulateDeductions            ' do not update the edited agency fee
    '    End If
    End If
    
    txtDeductions(2).Refresh
    txtNetPay = Format(txtGrossPay.Text - txtDeductions(2).Text, "###,##0.00")
    txtNetPay.Refresh
End Sub

'Private Sub ReComputeHazardEcola()
'Dim mMultiplier As Double
'Dim mEcolaPay As Double
'Dim mEcolaDeduc As Double
'Dim mHRs As Double
'
''mMultiplier = Val(txtNoDays.Text) + (Val(txtSundayOT.Text) / 8) + (Val(txtSplHoliday.Text) / 8) + _
''        (Val(txtLglHoliday.Text) / 8) + (Val(txtSickLeave.Text)) + (Val(txtVacLeave.Text))
'If PaymentBasis = "Monthly" Then ' if monthly
'    mMultiplier = Val(txtNoDays.Text)
'Else
'    mMultiplier = Val(txtNoDays.Text) + (Val(txtSundayOT.Text) / 8) + (Val(txtSplHoliday.Text) / 8) + _
'         (Val(txtSickLeave.Text)) + (Val(txtVacLeave.Text))
'End If
'    ' Add TxtSundayOT
'
'    ' Add txtSplHoliday
'
'    ' Add txtLglHoliday
'
'    ' Add txtSickLeave
'
'    ' Add txtVacLeave
'
'If WithHazardPay = True And HazardFlag = True Then
'    txtHazardPay.Text = Format(HazardPayRate * mMultiplier, "##,##0.00")
'End If
'
'''Get Ecola Rate per Day.
''If PaymentBasis = "Monthly" Then
''    mECOLARate = (mEcolaMonth / NoofDaysperMonth)
''Else
''    mECOLARate = mEcolaMonth
''End If
'
'If Val(txtDaysAbsent) > 0 Then      ' Deduc the lates and absences.
'    If PaymentBasis = "Monthly" Then ' if monthly paid get the fixed Ecola rate per day
'        mEcolaDeduc = (((mECOLARate * 2 * 12) / WorkDaysperYear) / 8) ' get ecola rate per hour
'        mHRs = (IIf(txtDaysAbsent.Text = "", 0, txtDaysAbsent.Text) * WorkHrPerDay)
'        mEcolaDeduc = mEcolaDeduc * mHRs
'        mEcolaPay = mECOLARate - mEcolaDeduc
'    Else
'        mHRs = (IIf(txtDaysAbsent.Text = "", 0, txtDaysAbsent.Text) * WorkHrPerDay)
'        mEcolaDeduc = (mECOLARate / 8) * mHRs
'        mEcolaPay = (mECOLARate * mMultiplier) - mEcolaDeduc
'    End If
'Else
'    ' Apply the daily way of computing COLA for monthly.
'    If PaymentBasis = "Monthly" Then ' if monthly paid get the fixed Ecola rate per day
'        mEcolaPay = mECOLARate
'    Else
'        mEcolaPay = mECOLARate * mMultiplier
'    End If
'End If
'
'' Total Ecola Amount.
'txtECOLA.Text = Format(mEcolaPay, "##,##0.00")
'TxtAmount(17) = Format(mEcolaPay, "##,##0.00")
'MonthlyEcola = Format(mEcolaPay, "##,##0.00")
'
'End Sub

Private Function GetPreviousSSS(ByVal mFlag As Integer) As Double

GetPreviousSSS = 0
If rsPreviousEmployee.RecordCount = 0 Then
    Exit Function
End If

With rsPreviousEmployee
    .MoveFirst
    .Find "EmployeeID = " & EmployeeNo
    If .EOF = False Then
        If mFlag = 1 Then ' return the previous sss
            GetPreviousSSS = !SSSEmployee
        ElseIf mFlag = 2 Then ' return previous phealth contributions
            GetPreviousSSS = !PHealthEmployee
        Else ' return previous pagibig contributions
            GetPreviousSSS = !pagibigemployee
        End If
    End If
End With
End Function

Private Sub txtCalamityLoan_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtCalamityLoan.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtCalamityLoan
End Sub

Private Sub txtBDOCard_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtBDOCard.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtBDOCard
End Sub

Private Sub txtBDOCard_LostFocus()
    txtBDOCard.BackColor = &HFFFFFF
    Call AccumulateDeductions
    txtBDOCard.Text = Format(txtBDOCard.Text, "##,##0.00")
End Sub

Private Sub txtOtherAdj_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtOtherAdj.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtOtherAdj
End Sub

Private Sub txtOtherAdj_LostFocus()
    txtAdj(4) = IIf(txtOtherAdj.Text = "", 0#, ChkVal(txtOtherAdj))
    Call SumAdjustment(txtAdj(4))
    txtAdj(4).BackColor = &HFFFFFF
    txtOtherAdj.BackColor = &HFFFFFF
    txtAdj(4) = Format(txtAdj(4), "##,##0.00")
End Sub

Private Sub txtCalLoan_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtCalLoan.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtCalLoan
End Sub

Private Sub txtCalLoan_LostFocus()
    txtCalLoan.BackColor = &HFFFFFF
    Call AccumulateDeductions
    txtCalLoan.Text = Format(txtCalLoan.Text, "##,##0.00")
End Sub

'Private Sub txtHDMFCal_GotFocus()
'    SendKeys "{HOME}+{END}"
'    txtHDMFCal.BackColor = &HFFFFC0
'    Set oCommonFunction.ValidDouble = txtHDMFCal
'End Sub
'
'Private Sub txtHDMFCal_LostFocus()
'    txtHDMFCal.BackColor = &HC0C0C0
'    txtHDMFCal.ForeColor = &HFF0000
'    Call AccumulateDeductions
'    txtHDMFCal.Text = Format(txtHDMFCal.Text, "##,##0.00")
'End Sub

Private Sub txtPagibigLoan_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtPagibigLoan.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtPagibigLoan
End Sub

Private Sub txtPagibigLoan_LostFocus()
    txtPagibigLoan.BackColor = &HC0C0C0
    txtPagibigLoan.ForeColor = &HFF0000
    Call AccumulateDeductions
    txtPagibigLoan.Text = Format(txtPagibigLoan.Text, "##,##0.00")
End Sub

Private Sub txtRDSplLgl_GotFocus()  ' updated Jan 17, 2015
    Sendkeys "{HOME}+{END}"
    txtRDSplLgl.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtRDSplLgl
End Sub

Private Sub txtRDSplLgl_LostFocus() ' updated Jan 17, 2015
    TxtAmount(18).BackColor = &HFFFFFF
    TxtAmount(18).Text = Format(ComputeAmount(txtRDSplLgl.Text, PaymentBasis, True, 18), "###,##0.00")
    txtRDSplLgl.BackColor = &HFFFFFF
    Call AccumulateAmount
End Sub

Private Sub txtRDonSPL_GotFocus() ' updated Jan 17, 2015
    Sendkeys "{HOME}+{END}"
    txtRDonSPL.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtRDonSPL
End Sub

Private Sub txtRDonSPL_LostFocus() ' updated June 26, 2018
    TxtAmount(19).BackColor = &HFFFFFF
    TxtAmount(19).Text = Format(ComputeAmount(txtRDonSPL.Text, PaymentBasis, True, 19), "###,##0.00")
    txtRDonSPL.BackColor = &HFFFFFF
    Call ComputeForECOLA(6, ChkVal(txtRDonSPL))
    Call AccumulateAmount
End Sub

Private Sub txtRDonLGL_GotFocus() ' updated Jan 17, 2015
    Sendkeys "{HOME}+{END}"
    txtRDonLGL.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtRDonLGL
End Sub

Private Sub txtRDonLGL_LostFocus() ' updated Jan 17, 2015
    TxtAmount(20).BackColor = &HFFFFFF
    TxtAmount(20).Text = Format(ComputeAmount(txtRDonLGL.Text, PaymentBasis, True, 20), "###,##0.00")
    txtRDonLGL.BackColor = &HFFFFFF
    Call ComputeForECOLA(7, ChkVal(txtRDonLGL))
    Call AccumulateAmount
End Sub

Private Sub txtSSSLoan_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtSSSLoan.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtSSSLoan
End Sub

Private Sub txtSSSLoan_LostFocus()
    txtSSSLoan.BackColor = &HC0C0C0
    txtSSSLoan.ForeColor = &HFF0000
    Call AccumulateDeductions
    txtSSSLoan.Text = Format(txtSSSLoan.Text, "##,##0.00")
End Sub

Private Sub txtTASAesla_GotFocus()
    'Call txtHoursLate_LostFocus
    Sendkeys "{HOME}+{END}"
    txtTASAesla.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtTASAesla
End Sub

Private Sub txtTASAesla_LostFocus()
    txtTASAesla.BackColor = &HFFFFFF
    Call AccumulateDeductions
    txtTASAesla.Text = Format(txtTASAesla.Text, "##,##0.00")
End Sub

Private Sub txtDaysAbsent_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtDaysAbsent.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtDaysAbsent
End Sub

Private Sub txtDaysAbsent_LostFocus()
    txtDeductions(0).BackColor = &HFFFFFF
    Call ComputeForECOLA(1) ' updated codes feb 26, 2006
    'txtDeductions(0).Text = Format(ComputeAmount(txtDaysAbsent.Text, PaymentBasis, False, 300), "###,##0.00")
    txtDeductions(0).Text = Format(ComputeAmount(txtDaysAbsent.Text, PaymentBasis, True, 302), "###,##0.00")
    txtDaysAbsent.BackColor = &HFFFFFF
    Call AccumulateDeductions
    ' If there are entries update deductions. updated feb 26, 2006.
    'Call GetContributions(txtGrossPay1.Text - Val(Replace(txtDeductions(0).Text, ",", "")))
    Call GetContributions(txtGrossPay1.Text)
'    If WTaxFlag = True Then
'        Call ComputeTax(txtGrossPay1.Text - txtDeductions(0).Text) ' bernard
'    End If
    Call AccumulateAmount
    Call AccumulateDeductions
End Sub

Private Sub txtDaysAbsent_Validate(Cancel As Boolean)
    txtDeductions(0).BackColor = &HFFFFFF
    txtDeductions(0).Text = Format(ComputeAmount(txtDaysAbsent.Text, PaymentBasis, True, 302), "###,##0.00")
    txtDaysAbsent.BackColor = &HFFFFFF
    Call AccumulateDeductions
    ' If there are entries update deductions. updated feb 26, 2006.
    'Call GetContributions(txtGrossPay1.Text - Val(Replace(txtDeductions(0).Text, ",", "")))
    Call GetContributions(txtGrossPay1.Text)
'    If WTaxFlag = True Then
'        Call ComputeTax(txtGrossPay1.Text - txtDeductions(0).Text) ' bernard
'    End If
    Call AccumulateDeductions
End Sub

Private Sub txtNoDays_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtNoDays.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtNoDays
End Sub

Private Sub txtNoDays_LostFocus()
    If IsNull(txtNoDays) Or txtNoDays = "" Then
        txtNoDays = "0"
    End If
    
'    ' Included in AccumulateAmount Procedure.
'    If WithHazardPay = True Then
'        'txtHazardPay.Text = Format(HazardPayRate * (Val(txtDaysWork.Text) / 8), "##0.00")
'        txtHazardPay.Text = Format(HazardPayRate * Val(txtNoDays.Text), "##0.00")
'        txtUnionDues.Text = Format(UnionDuesRate, "##0.00")
'    End If
    txtNoDays.BackColor = &HFFFFFF
    Call ComputeForECOLA(0)
    Call GetContributions(txtGrossPay1)    'Updated Bernard
    Call AccumulateAmount(False, False, True)
    Call AutoClickText
End Sub

Private Sub txtDaysWork_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtDaysWork.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtDaysWork
End Sub

Private Sub txtDaysWork_LostFocus()
    MonthlyFlag = True
    If IsNull(txtDaysWork) Or txtDaysWork = "" Then
        txtDaysWork = "0"
    Else
        txtNoDays.Text = Val(txtDaysWork) / 8
    End If
    
'    ' Included in AccumulateAmount Procedure
'    If WithHazardPay = True Then
'        'txtHazardPay.Text = Format(HazardPayRate * (Val(txtDaysWork.Text) / 8), "##0.00")
'        txtHazardPay.Text = Format(HazardPayRate * Val(txtNoDays.Text), "##0.00")
'        txtUnionDues.Text = Format(UnionDuesRate, "##0.00")
'    End If
'    'Auto-compute Allowance if rate is non-zero.
'    'If AllowancePayRate > 0# Then
'    If AllowanceBase = "Per Day" Then
'        txtIncentive.Text = Format(AllowancePayRate * Val(txtNoDays.Text), "##0.00")
'    Else
'        'txtIncentive.Text = Format(mAllowance, "##,##0.00")
'        txtIncentive.Text = Format(AllowancePayRate, "##,##0.00")
'    End If
    
    'Call AutoClickText
    
    TxtAmount(0).BackColor = &HFFFFFF
    TxtAmount(0).Text = Format(ComputeAmount(txtDaysWork.Text, PaymentBasis, True, 0), "###,##0.00")
    txtDaysWork.BackColor = &HFFFFFF
    Call ComputeForECOLA(0)
    Call GetContributions(txtGrossPay1)    'Updated Bernard
    MonthlyFlag = False
    Call AccumulateAmount(False, False, True)
    OrigMonthlyRate = TxtAmount(0).Text ' Updated April 2, 2011
    Call AutoClickText
End Sub

Private Function ComputeAmount(ByRef mAmt As String, mPay As String, mHrSw As Boolean, ByVal IndexNo As Integer, Optional Boxno As Integer) As Double
Dim TotValue As Double
Dim i As Integer
Dim DeciSw As Boolean
Dim strAmt As String

    DeciSw = False
    If IsNull(mAmt) Or mAmt = "" Or mAmt = "0" Then
        mAmt = "0.00"
        ComputeAmount = 0#
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
        mAmt = IIf(strAmt = ".", 0#, Format(strAmt, "0.00"))
        If mAmt > 99999.99 Then
            MsgBox "Entries should not exceed 99,999.99. Please correct it!", vbInformation, PayrollAlert
            ComputeAmount = 0#
            Exit Function
        End If
        
        'Recompute for Daily Rate, Updated Oct. 18, 2014. Bernard
        If DailyRate > 0 Then
            HourlyRate = DailyRate / 8
        End If
        
        If mPay = "Hourly" Then ' if hourly rate
            If mHrSw = False Then
                ComputeAmount = mAmt * (DailyRate / 8)   'HourlyRate
            Else
                If IndexNo = 301 Then ' compute for late in hours
                    ComputeAmount = mAmt * HourlyRate
                ElseIf IndexNo = 302 Then ' compute for Absent in days
                    ComputeAmount = mAmt * DailyRate
                Else
                    ComputeAmount = mAmt * HourlyRate
                End If
            End If
        ElseIf mPay = "Daily" Then
            If mHrSw = False Then
                ComputeAmount = mAmt * DailyRate
            Else
                If IndexNo = 301 Then ' compute for late in hours
                    ComputeAmount = mAmt * HourlyRate
                ElseIf IndexNo = 302 Then ' compute for Absent in days
                    ComputeAmount = mAmt * DailyRate
                Else
                    ComputeAmount = mAmt * HourlyRate
                End If
            End If
        ElseIf mPay = "Monthly" Then
                If MonthlyFlag = False Then ' do not used the monthly salary as base
                    If mHrSw = False Then
                        ComputeAmount = mAmt * DailyRate
                    ElseIf IndexNo = 302 Then ' compute for Absent in days
                        ComputeAmount = mAmt * DailyRate
                    Else
                        'ComputeAmount = mAmt * HourlyRate
                        'Updated October 14, 2006. Bernard
                        'ComputeAmount = mAmt * (((MonthlyRate * 12) / 314) / 8)
                        ComputeAmount = mAmt * (((MonthlyRate * 12) / WorkDaysperYear) / 8)
                    End If
                Else        ' get the monthly salary base
                    If mHrSw = False Then
                        ComputeAmount = MonthlyRate / 2 'NoofCutoffs
                    ElseIf IndexNo = 302 Then ' compute for Absent in days
                        ComputeAmount = mAmt * DailyRate
                    Else
                        ComputeAmount = MonthlyRate / 2
                        'ComputeAmount = mAmt * HourlyRate
                    End If
                End If
            'End If
        End If
    End If
    Select Case IndexNo
        Case 1 ' get the overtime rate
            If EcolaInOTRate = True Then    ' Include Ecola in OT Rate, April 15, 2006
                ComputeAmount = IIf(OverTimeRate = 0, 1, OverTimeRate) * GetOTRate(mAmt)
            Else
                ComputeAmount = IIf(OverTimeRate = 0, 1, OverTimeRate) * ComputeAmount
            End If
        Case 2 ' get the overtime night diff rate
            If EcolaInOTRate = True Then    ' Include Ecola in OT Rate, April 15, 2006
                ComputeAmount = IIf(NDiffRate = 0, 1, NDiffRate) * GetOTRate(mAmt)
            Else
                ComputeAmount = IIf(NDiffRate = 0, 1, NDiffRate) * ComputeAmount
            End If
        Case 4 ' get the Rest Day rate , previously 3
            ComputeAmount = IIf(RestDayRate = 0, 1, RestDayRate) * ComputeAmount
        Case 444 ' get the Sunday ot rate (updated to Excess Rest Day OT hours April 23, 2009)
            If EcolaInOTRate = True Then    ' Include Ecola in OT Rate, April 15, 2006
                ComputeAmount = IIf(SundayRate = 0, 1, SundayRate) * GetOTRate(mAmt)
            Else
                ComputeAmount = IIf(SundayRate = 0, 1, SundayRate) * ComputeAmount
            End If
        Case 5 ' get the SPL holiday rate
            ComputeAmount = IIf(LglHolidayRate = 0, 1, LglHolidayRate) * ComputeAmount
        Case 6 ' get the legal holiday rate
            ComputeAmount = IIf(SplHolidayRate = 0, 1, SplHolidayRate) * ComputeAmount
        Case 7 ' get the spl holiday ot rate
            If EcolaInOTRate = True Then    ' Include Ecola in OT Rate, April 15, 2006
                ComputeAmount = IIf(SplHolidayRate = 0, 1, SplHolidayRate) * GetOTRate(mAmt)
            Else
                ComputeAmount = IIf(SplHolidayOTRate = 0, 1, SplHolidayOTRate) * ComputeAmount
            End If
        Case 8 ' get the legal holiday rate
            ComputeAmount = IIf(LglHolidayOTRate = 0, 1, LglHolidayOTRate) * ComputeAmount
        Case 9  ' Sick Leave Pay
            ComputeAmount = mAmt * DailyRate
        Case 10 ' Vacation Leave
            ComputeAmount = mAmt * DailyRate
        Case 12 ' get the excess ot reg day rate updated 3/12/17
            If EcolaInOTRate = True Then    ' Include Ecola in OT Rate, April 15, 2006
                ComputeAmount = IIf(ExcessOTRegDayRate = 0, 1, ExcessOTRegDayRate) * GetOTRate(mAmt)
            Else
                ComputeAmount = IIf(ExcessOTRegDayRate = 0, 1, ExcessOTRegDayRate) * ComputeAmount
            End If
        Case 13 ' get the excess rate
            ComputeAmount = IIf(ExcessOTSplSunRate = 0, 1, ExcessOTSplSunRate) * ComputeAmount
        Case 14 ' get the excess rate
            ComputeAmount = IIf(ExcessOTSunLglRate = 0, 1, ExcessOTSunLglRate) * ComputeAmount
        Case 15 ' get the excess rate
            ComputeAmount = IIf(ExcessOTLegalRate = 0, 1, ExcessOTLegalRate) * ComputeAmount
        Case 16 ' get the excess rate
            ComputeAmount = IIf(OTwoPremiumRate = 0, 1, OTwoPremiumRate) * ComputeAmount
   
        ' Updated Jan 17, 2015 additional Restday OT on SPL / LGL Holidays.
        Case 18 ' get the RD Spl / Lgl OT Rate
            ComputeAmount = (HourlyRate * IIf(ExcessOTSunLglRate = 0, 0.3, ExcessOTSunLglRate)) * ChkVal(txtRDSplLgl)
        Case 19 ' get the RD on SPL Rate
            ComputeAmount = (HourlyRate * IIf(OTwoPremiumRate = 0, 1.5, OTwoPremiumRate)) * ChkVal(txtRDonSPL)
        Case 20 ' get the RD on LGL Rate
            ComputeAmount = (HourlyRate * IIf(ExcessOTLegalRate = 0, 2.6, ExcessOTLegalRate)) * ChkVal(txtRDonLGL)
    End Select
    'Call AccumulateDeductions
    'Call AccumulateAmount
End Function

Private Sub txtECOLA_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtECOLA.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtECOLA
End Sub

Private Sub txtECOLA_LostFocus()
Dim i As Integer
Dim strAmt As String
Dim mAmt As String
Dim DeciSw As Boolean

    TxtAmount(17).BackColor = &HFFFFFF
    If EditFlag = False Then ' Updated Nov. 19, 2016 do not re-update the data
        Call ComputeForECOLA(1)
    End If
    'TxtAmount(17).Text = Format(ComputeAmount(txtECOLA.Text, PaymentBasis, True, 17), "###,##0.00")
    txtECOLA.BackColor = &HFFFFFF
    Call AccumulateAmount

'    txtECOLA.BackColor = &HFFFFFF
'    mAmt = Trim(txtECOLA.Text)
'    For i = 1 To Len(mAmt)
'        If DeciSw = True And mID(mAmt, i, 1) = "." Then
'        Else
'            strAmt = strAmt & mID(mAmt, i, 1)
'        End If
'        If DeciSw = False And mID(mAmt, i, 1) = "." Then
'            DeciSw = True
'        End If
'    Next
'    mAmt = IIf(strAmt = ".", 0#, Format(strAmt, "0.00"))
'    TxtAmount(17).BackColor = &HFFFFFF
'    TxtAmount(17).Text = Format(mAmt, "##,##0.00")
'    MonthlyEcola = Val(TxtAmount(17).Text)
'    Call AccumulateAmount
'    Call AccumulateDeductions
'    txtECOLA.Text = Format(mAmt, "##,##0.00")
End Sub

Private Sub txtExcessOTLegal_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtExcessOTLegal.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtExcessOTLegal
End Sub

Private Sub txtExcessOTLegal_LostFocus()
    TxtAmount(15).BackColor = &HFFFFFF
    TxtAmount(15).Text = Format(ComputeAmount(txtExcessOTLegal.Text, PaymentBasis, True, 15), "###,##0.00")
    txtExcessOTLegal.BackColor = &HFFFFFF
    Call AccumulateAmount
End Sub

Private Sub txtExcessOTRegDay_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtExcessOTRegDay.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtExcessOTRegDay
End Sub

Private Sub txtExcessOTRegDay_LostFocus()
    TxtAmount(12).BackColor = &HFFFFFF
    'TxtAmount(12).Text = Format(ComputeAmount(txtExcessOTRegDay.Text, PaymentBasis, True, 12), "###,##0.00")
    TxtAmount(12).Text = Format(ChkVal(txtExcessOTRegDay), "###,##0.00")
    txtExcessOTRegDay.BackColor = &HFFFFFF
    Call AccumulateAmount
End Sub

Private Sub txtExcessOTSplSun_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtExcessOTSplSun.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtExcessOTSplSun
End Sub

Private Sub txtExcessOTSplSun_LostFocus()
    TxtAmount(13).BackColor = &HFFFFFF
    TxtAmount(13).Text = Format(ComputeAmount(txtExcessOTSplSun.Text, PaymentBasis, True, 13), "###,##0.00")
    txtExcessOTSplSun.BackColor = &HFFFFFF
    Call AccumulateAmount
End Sub

Private Sub txtExcessOTSunLgl_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtExcessOTSunLgl.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtExcessOTSunLgl
End Sub

Private Sub txtExcessOTSunLgl_LostFocus()
    TxtAmount(14).BackColor = &HFFFFFF
    TxtAmount(14).Text = Format(ComputeAmount(txtExcessOTSunLgl.Text, PaymentBasis, True, 14), "###,##0.00")
    txtExcessOTSunLgl.BackColor = &HFFFFFF
    Call AccumulateAmount
End Sub

Private Sub AccumulateDeductions(Optional UpdAgency As Boolean)
Dim i As Integer
Dim TotValue As Double
    
'    For i = 0 To 1
'        TotValue = TotValue + IIf(txtDeductions(i).Text = "", 0, txtDeductions(i).Text)
'        txtDeductions(i).Refresh
'    Next
    TotValue = IIf(txtDeductions(0).Text = "", 0, Replace(txtDeductions(0).Text, ",", ""))    ' No of days absent
    If PaymentBasis = "Monthly" And TotValue <> 0 Then        ' do not include absences in deductions
        AccumulateAmount
        TotValue = 0
    End If
'    TotValue = TotValue + IIf(txtBDOCard.Text = "", 0, txtBDOCard.Text) + IIf(txtTASAesla.Text = "", 0, txtTASAesla.Text) + IIf(txtUnionDues.Text = "", 0, txtUnionDues.Text) + _
'        IIf(txtVale.Text = "", 0, txtVale.Text) + IIf(txtPlacementFee.Text = "", 0, txtPlacementFee.Text) + IIf(txtMiscFee.Text = "", 0, txtMiscFee.Text) + IIf(txtOverPay.Text = "", 0, txtOverPay.Text) + _
'        IIf(txtEmpyeeSSS.Text = "", 0, txtEmpyeeSSS.Text) + IIf(txtEmpyeePagibig.Text = "", 0, txtEmpyeePagibig.Text) + IIf(txtPhilHealthYee.Text = "", 0, txtPhilHealthYee.Text) + _
'        IIf(txtPagibigLoan.Text = "", 0, txtPagibigLoan.Text) + IIf(txtSSSLoan.Text = "", 0, txtSSSLoan.Text) + IIf(txtCalamityLoan.Text = "", 0, txtCalamityLoan.Text) + IIf(txtHDMFCal.Text = "", 0, txtHDMFCal.Text) + IIf(txtTax.Text = "", 0, txtTax.Text) + _
'        mTotalDeductions

    ' Updated 9/19/2018 Bernard.
    TotValue = TotValue + ChkVal(txtBDOCard) + ChkVal(txtTASAesla) + ChkVal(txtUnionDues) + _
        ChkVal(txtVale) + ChkVal(txtPlacementFee) + ChkVal(txtMiscFee) + ChkVal(txtOverPay) + _
        ChkVal(txtEmpyeeSSS) + ChkVal(txtEmpyeePagibig) + ChkVal(txtPhilHealthYee) + _
        ChkVal(txtPagibigLoan) + ChkVal(txtSSSLoan) + ChkVal(txtCalamityLoan) + ChkVal(txtHDMFCal) + _
        ChkVal(txtTax) + ChkVal(txtCalLoan) + mTotalDeductions
    
    txtDeductions(2).Text = Format(TotValue, "###,##0.00")
    txtDeductions(2).Refresh
    ' Compute for Agency Charge.
    TotValue = 0
    ' Get employer contributions.
    TotValue = TotValue + IIf(txtEmpyerSSS.Text = "", 0, txtEmpyerSSS.Text) + IIf(txtEmpyerPagibig.Text = "", 0, txtEmpyerPagibig.Text) + IIf(txtPHealthYer.Text = "", 0, txtPHealthYer.Text) '+ _
        'IIf(txtPagibigLoan.Text = "" Or txtPagibigLoan = "0", 0, txtPagibigLoan.Text) + IIf(txtTAX.Text = "", 0, txtTAX.Text)
'    If AgencySw = False Then ' Compute for Agency Charge
'        If WCalamityLoan = False And EmployeeAgencyRate <> 0 Then ' if not 0, use employee rate
'            txtCalamityLoan.Text = Format(EmployeeAgencyRate * txtDaysWork.Text, "##,##0.00")
'        ElseIf WCalamityLoan = True Then    ' else use the client agency rate
'            txtCalamityLoan.Text = Format((AgencyRate / 100) * (TotValue + txtGrossPay1), "##,##0.00")
'        End If
'        'AgencySw = False
'    End If
    
    txtNetPay = Format(txtGrossPay.Text - txtDeductions(2).Text, "###,##0.00")

End Sub

Private Sub txtHoursLate_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtHoursLate.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtHoursLate
End Sub

Private Sub txtHoursLate_LostFocus()
    txtDeductions(1).BackColor = &HFFFFFF
    Call ComputeForECOLA(3)
    txtDeductions(1).Text = Format(ComputeAmount(txtHoursLate.Text, PaymentBasis, True, 301), "###,##0.00")
    txtHoursLate.BackColor = &HFFFFFF
    Call AccumulateDeductions
    'Call getdeductions
    'Call GetContributions(txtGrossPay1.Text - Val(Replace(txtDeductions(1).Text, ",", "")))
    Call GetContributions(txtGrossPay1.Text)

'    If WTaxFlag = True Then
'        Call ComputeTax(txtGrossPay1.Text - txtDeductions(1).Text) ' bernard
'    End If
    Call AccumulateAmount
    Call AccumulateDeductions
End Sub

Private Sub txtHoursLate_Validate(Cancel As Boolean)
    txtDeductions(1).BackColor = &HFFFFFF
    txtDeductions(1).Text = Format(ComputeAmount(txtHoursLate.Text, PaymentBasis, True, 301), "###,##0.00")
    txtHoursLate.BackColor = &HFFFFFF
    Call AccumulateDeductions
    Call GetContributions(txtGrossPay1.Text)

'    If WTaxFlag = True Then
'        Call ComputeTax(txtGrossPay1.Text - txtDeductions(1).Text) ' bernard
'    End If
    Call AccumulateDeductions
End Sub

Private Sub txtIncentive_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtIncentive.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtIncentive
End Sub

Private Sub txtIncentive_LostFocus()
Dim i As Integer
Dim strAmt As String
Dim mAmt As String
Dim DeciSw As Boolean

    TxtAmount(10).BackColor = &HFFFFFF
    TxtAmount(10).Text = Format(FormatedAmt(Trim(txtIncentive)) + FormatedAmt(Trim(txtSplOtHours)), "##,##0.00")
    txtIncentive.BackColor = &HFFFFFF
    Call AccumulateAmount
    txtIncentive.Text = Format(txtIncentive, "##,##0.00")
    
'    txtAmount(10).BackColor = &HFFFFFF
'    mAmt = Trim(txtIncentive)
'    For i = 1 To Len(mAmt)
'        If DeciSw = True And mID(mAmt, i, 1) = "." Then
'        Else
'            strAmt = strAmt & mID(mAmt, i, 1)
'        End If
'        If DeciSw = False And mID(mAmt, i, 1) = "." Then
'            DeciSw = True
'        End If
'    Next
'    mAmt = IIf(strAmt = "." Or strAmt = "", 0#, Format(strAmt, "0.00"))
'    txtAmount(10).Text = Format(mAmt, "##,##0.00")
'    txtIncentive.BackColor = &HFFFFFF
'    If mAmt > 0 Then
'        lblEntries(25).Caption = "Note"
'        'txtIncentiveRemarks.Visible = True
''        If txtIncentiveRemarks.Text = "" Then
''            txtIncentiveRemarks.Text = "May-June"
''        End If
'        'TxtAmount(10).Visible = False
''        If txtIncentiveRemarks.Enabled = True Then
''            txtIncentiveRemarks.SetFocus
''        End If
'    End If
'    Call AccumulateAmount
'    txtIncentive.Text = Format(mAmt, "##,##0.00")
End Sub

Function FormatedAmt(mAmt As String) As Double
Dim i As Integer
Dim strAmt As String
Dim DeciSw As Boolean

FormatedAmt = 0#
For i = 1 To Len(mAmt)
    If DeciSw = True And mID(mAmt, i, 1) = "." Then
    Else
        strAmt = strAmt & mID(mAmt, i, 1)
    End If
    If DeciSw = False And mID(mAmt, i, 1) = "." Then
        DeciSw = True
    End If
Next
FormatedAmt = IIf(strAmt = "." Or strAmt = "", 0#, Format(strAmt, "0.00"))
End Function

Private Sub txtIncentiveRemarks_GotFocus()
'    SendKeys "{HOME}+{END}"
'    txtIncentiveRemarks.BackColor = &HFFFFC0
'    Set oCommonFunction.ValidAlpha = txtIncentiveRemarks
End Sub

Private Sub txtIncentiveRemarks_LostFocus()
'    txtAmount(10).Visible = True
'    txtIncentiveRemarks.BackColor = &HFFFFFF
'    Call AccumulateAmount
End Sub

Private Sub txtLglHoliday_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtLglHoliday.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtLglHoliday
End Sub

Private Sub txtLglHoliday_LostFocus()
Dim mTot As Double

    TxtAmount(6).BackColor = &HFFFFFF
    TxtAmount(6).Text = Format(ComputeAmount(txtLglHoliday.Text, PaymentBasis, True, 6), "###,##0.00")
    txtLglHoliday.BackColor = &HFFFFFF
    'txtECOLA.Text = GetHolidayEcolaRate(txtECOLA.Text, 3)

'    ' Updated 5/1/2011
'    If PaymentBasis = "Monthly" Then
'        mTot = Val(Replace(TxtAmount(5).Text, ",", "")) + Val(Replace(TxtAmount(6).Text, ",", "")) + _
'            Val(Replace(TxtAmount(7).Text, ",", "")) + Val(Replace(TxtAmount(8).Text, ",", ""))
'        TxtAmount(0).Text = Format(OrigMonthlyRate - mTot, "###,#00.00")
'    End If
    
    Call ComputeForECOLA(1) ' updated Sep. 1, 2016 add Ecola rate for restday
    Call AccumulateAmount(False, False, True)
End Sub

Private Sub txtLglHolidayOT_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtLglHolidayOT.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtLglHolidayOT
End Sub

Private Sub txtLglHolidayOT_LostFocus()
'    TxtAmount(8).BackColor = &HFFFFFF
'    TxtAmount(8).Text = Format(ComputeAmount(txtLglHolidayOT.Text, PaymentBasis, True, 8), "###,##0.00")
'    txtLglHolidayOT.BackColor = &HFFFFFF
'    Call AccumulateAmount
End Sub

Private Sub txtMiscFee_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtMiscFee.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtMiscFee
End Sub

Private Sub txtMiscFee_LostFocus()
    txtMiscFee.BackColor = &HFFFFFF
    Call AccumulateDeductions
    txtMiscFee.Text = Format(txtMiscFee.Text, "##,##0.00")
End Sub

Private Sub txtNightDiff_GotFocus()
    Sendkeys "{HOME}+{END}"
    Set oCommonFunction.ValidDouble = txtNightDiff
    txtNightDiff.BackColor = &HFFFFC0
End Sub

Private Sub txtNightDiff_LostFocus()
    TxtAmount(2).BackColor = &HFFFFFF
    TxtAmount(2).Text = Format(ComputeAmount(txtNightDiff.Text, PaymentBasis, True, 2), "###,##0.00")
    txtNightDiff.BackColor = &HFFFFFF
    Call AccumulateAmount
End Sub

'Private Sub txtOTwoPremium_GotFocus()
'    SendKeys "{HOME}+{END}"
'    txtOTwoPremium.BackColor = &HFFFFC0
'    Set oCommonFunction.ValidDouble = txtOTwoPremium
'End Sub

'Private Sub txtOTwoPremium_LostFocus()
'    TxtAmount(16).BackColor = &HFFFFFF
'    TxtAmount(16).Text = Format(ComputeAmount(txtOTwoPremium.Text, PaymentBasis, True, 16), "###,##0.00")
'    txtOTwoPremium.BackColor = &HFFFFFF
'    Call AccumulateAmount
'End Sub

Private Sub txtOverPay_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtOverPay.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtOverPay
End Sub

Private Sub txtOverPay_LostFocus()
    txtOverPay.BackColor = &HFFFFFF
    Call AccumulateDeductions
    txtOverPay.Text = Format(txtOverPay.Text, "##,##0.00")
End Sub

Private Sub txtPlacementFee_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtPlacementFee.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtPlacementFee
End Sub

Private Sub txtPlacementFee_LostFocus()
    txtPlacementFee.BackColor = &HFFFFFF
    Call AccumulateDeductions
    txtPlacementFee.Text = Format(txtPlacementFee.Text, "##,##0.00")
End Sub

Private Sub txtRegOT_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtRegOT.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtRegOT
End Sub

Private Sub txtRegOT_LostFocus()
    TxtAmount(1).BackColor = &HFFFFFF
    TxtAmount(1).Text = Format(ComputeAmount(txtRegOT.Text, PaymentBasis, True, 1), "###,##0.00")
    txtRegOT.BackColor = &HFFFFFF
    Call AccumulateAmount
End Sub

Function GetOTRate(ByVal OTHours As Double) As Double  ' Returns the hourly rate of basic pay + ecola
Dim cEcolaRate As Double
    'Formula as per elena April 15, 2006
    'ROT = 8500 X 12 MO./314 DAYS/8HRS/DAY X 1.25

    If PaymentBasis = "Monthly" Then
        cEcolaRate = ((mECOLARate * 12) / WorkDaysperYear) / WorkHrPerDay
        GetOTRate = ((((MonthlyRate * 12) / WorkDaysperYear) / WorkHrPerDay) + cEcolaRate) * OTHours
    Else
        cEcolaRate = (mECOLARate / WorkHrPerDay)
        GetOTRate = ((DailyRate / WorkHrPerDay) + cEcolaRate) * OTHours
    End If
    
End Function

Private Sub txtTaxRefund_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtTaxRefund.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtTaxRefund
End Sub

Private Sub txtTaxRefund_LostFocus()
Dim i As Integer
Dim strAmt As String
Dim mAmt As String
Dim DeciSw As Boolean

    TxtAmount(3).BackColor = &HFFFFFF
    mAmt = Trim(txtTaxRefund)
    For i = 1 To Len(mAmt)
        If DeciSw = True And mID(mAmt, i, 1) = "." Then
        Else
            strAmt = strAmt & mID(mAmt, i, 1)
        End If
        If DeciSw = False And mID(mAmt, i, 1) = "." Then
            DeciSw = True
        End If
    Next
    mAmt = IIf(strAmt = "." Or strAmt = "", 0#, Format(strAmt, "0.00"))
    TxtAmount(3).Text = Format(mAmt, "##,##0.00")
    txtTaxRefund.BackColor = &HFFFFFF
    Call AccumulateAmount
    txtTaxRefund.Text = Format(mAmt, "##,##0.00")
End Sub

Private Sub txtSalAdjustment_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtSalAdjustment.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtSalAdjustment
End Sub

Private Sub txtSalAdjustment_LostFocus()
Dim i As Integer
Dim strAmt As String
Dim mAmt As String
Dim DeciSw As Boolean

    TxtAmount(9).BackColor = &HFFFFFF
    mAmt = Trim(txtSalAdjustment)
    For i = 1 To Len(mAmt)
        If DeciSw = True And mID(mAmt, i, 1) = "." Then
        Else
            strAmt = strAmt & mID(mAmt, i, 1)
        End If
        If DeciSw = False And mID(mAmt, i, 1) = "." Then
            DeciSw = True
        End If
    Next
    mAmt = IIf(strAmt = "." Or strAmt = "", 0#, Format(strAmt, "0.00"))
    TxtAmount(9).Text = Format(mAmt, "##,##0.00")
    txtSalAdjustment.BackColor = &HFFFFFF
    Call AccumulateAmount
    txtSalAdjustment.Text = Format(mAmt, "##,##0.00")
End Sub

'Private Sub txtAdjustments_GotFocus()
'    SendKeys "{HOME}+{END}"
'    txtAdjustments.BackColor = &HFFFFC0
'    Set oCommonFunction.ValidDouble = txtAdjustments
'End Sub
'
'Private Sub txtAdjustments_LostFocus()
'    txtAmount(21).BackColor = &HFFFFFF
'    txtAmount(21).Text = Format(FormatedAmt(Trim(txtAdjustments)), "##,##0.00")
'    txtAdjustments.BackColor = &HFFFFFF
'    Call AccumulateAmount
'    txtAdjustments.Text = Format(txtAmount(21), "##,##0.00")
'End Sub

'Function FormatedAmt(mAmt As String) As Double
'Dim i As Integer
'Dim strAmt As String
'Dim DeciSw As Boolean
'
'FormatedAmt = 0#
'For i = 1 To Len(mAmt)
'    If DeciSw = True And mID(mAmt, i, 1) = "." Then
'    Else
'        strAmt = strAmt & mID(mAmt, i, 1)
'    End If
'    If DeciSw = False And mID(mAmt, i, 1) = "." Then
'        DeciSw = True
'    End If
'Next
'FormatedAmt = IIf(strAmt = "." Or strAmt = "", 0#, Format(strAmt, "0.00"))
'End Function

Private Sub txtSickLeave_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtSickLeave.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtSickLeave
End Sub

Private Sub txtSickLeave_LostFocus()
Dim mTot As Double

    TxtAmount(7).BackColor = &HFFFFFF
    TxtAmount(7).Text = Format(ComputeAmount(txtSickLeave.Text, PaymentBasis, False, 9), "###,##0.00")
    txtSickLeave.BackColor = &HFFFFFF
    
    If PaymentBasis = "Monthly" Then
'        mTot = Val(Replace(TxtAmount(5).Text, ",", "")) + Val(Replace(TxtAmount(6).Text, ",", "")) + _
'            Val(Replace(TxtAmount(7).Text, ",", "")) + Val(Replace(TxtAmount(8).Text, ",", ""))
'        TxtAmount(0).Text = Format(OrigMonthlyRate - mTot, "###,#00.00")
        TxtAmount(0).Text = Format(OrigMonthlyRate - (Val(Replace(TxtAmount(7).Text, ",", "")) + Val(Replace(TxtAmount(8).Text, ",", ""))), "###,#00.00")
    End If
    Call ComputeForECOLA(1) ' updated Sep. 1, 2016 add Ecola rate for restday
    Call AccumulateAmount(False, False, True)
End Sub

Private Sub txtSplHoliday_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtSplHoliday.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtSplHoliday
End Sub

Private Sub txtSplHoliday_LostFocus()
Dim mTot As Double

    TxtAmount(5).BackColor = &HFFFFFF
    TxtAmount(5).Text = Format(ComputeAmount(txtSplHoliday.Text, PaymentBasis, True, 5), "###,##0.00")
    txtSplHoliday.BackColor = &HFFFFFF
    'txtECOLA.Text = GetHolidayEcolaRate(txtECOLA.Text, 2)
    
'    ' Updated 5/1/2011
'    If PaymentBasis = "Monthly" Then
'        mTot = Val(Replace(TxtAmount(5).Text, ",", "")) + Val(Replace(TxtAmount(6).Text, ",", "")) + _
'            Val(Replace(TxtAmount(7).Text, ",", "")) + Val(Replace(TxtAmount(8).Text, ",", ""))
'        TxtAmount(0).Text = Format(OrigMonthlyRate - mTot, "###,#00.00")
'    End If
    
    Call ComputeForECOLA(1) ' updated Sep. 1, 2016 add Ecola rate for restday
    Call AccumulateAmount(False, False, True)
End Sub

Private Sub txtSplOTHours_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtSplOtHours.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtSplOtHours
End Sub

Private Sub txtSplOTHours_LostFocus()
    TxtAmount(10).BackColor = &HFFFFFF
    TxtAmount(10).Text = Format(FormatedAmt(Trim(txtIncentive)) + FormatedAmt(Trim(txtSplOtHours)), "##,##0.00")
    txtSplOtHours.BackColor = &HFFFFFF
    Call AccumulateAmount
    txtSplOtHours.Text = Format(txtSplOtHours, "##,##0.00")
End Sub

Private Sub txtSundayOT_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtSundayOT.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtSundayOT
End Sub

Private Sub txtSundayOT_LostFocus()
    TxtAmount(4).BackColor = &HFFFFFF
    TxtAmount(4).Text = Format(ComputeAmount(txtSundayOT.Text, PaymentBasis, True, 4), "###,##0.00")
    txtSundayOT.BackColor = &HFFFFFF
    Call ComputeForECOLA(1) ' updated Sep. 1, 2016 add Ecola rate for restday
    Call AccumulateAmount(False, False, True)
End Sub

Private Sub txtTAX_GotFocus()
'    SendKeys "{HOME}+{END}"
'    txtTAX.BackColor = &HFFFFC0
'    Set oCommonFunction.ValidDouble = txtTAX
End Sub

Private Sub txtCalamityLoan_LostFocus()
    txtCalamityLoan.BackColor = &HC0C0C0
    txtCalamityLoan.ForeColor = &HFF0000
    Call AccumulateAmount(True)
    txtCalamityLoan.Locked = True
    If cmdEditTax(3).Enabled = True Then
        cmdEditTax(3).SetFocus
    End If
End Sub

Private Sub txtUnpaidSL_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtUnpaidSL.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtUnpaidSL
End Sub

Private Sub txtUnpaidVL_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtUnpaidVL.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtUnpaidVL
End Sub

Private Sub txtUnpaidVL_LostFocus() ' updated Jan 17, 2015
    txtAdj(0) = ChkVal(txtUnpaidVL) * DailyRate
    Call SumAdjustment(0)
    txtAdj(0).BackColor = &HFFFFFF
    txtUnpaidVL.BackColor = &HFFFFFF
    txtAdj(0) = Format(txtAdj(0), "##,##0.00")
End Sub

Private Sub txtUnpaidSL_LostFocus() ' updated Jan 17, 2015
    txtAdj(1) = ChkVal(txtUnpaidSL) * DailyRate
    Call SumAdjustment(0)
    txtAdj(1).BackColor = &HFFFFFF
    txtUnpaidSL.BackColor = &HFFFFFF
    txtAdj(1) = Format(txtAdj(1), "##,##0.00")
End Sub

Private Sub txtEcolaAdj_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtEcolaAdj.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtEcolaAdj
End Sub

Private Sub txtEcolaAdj_LostFocus() ' updated Dec. 3, 2016
Dim DailyEcola As Double

    If EditFlag = True Then ' Updated Nov. 19, 2016 do not re-update the data
        If PaymentBasis = "Monthly" Then ' if monthly paid get the fixed Ecola rate
            DailyEcola = ((mECOLARate * 2 * 12) / WorkDaysperYear)
        Else
            DailyEcola = mECOLARate
        End If
    End If
    
    txtAdj(2) = IIf(txtEcolaAdj.Text = "", 0, ChkVal(txtEcolaAdj)) * DailyEcola
    Call SumAdjustment(txtAdj(2))
    txtAdj(2).BackColor = &HFFFFFF
    txtEcolaAdj.BackColor = &HFFFFFF
    txtAdj(2) = Format(txtAdj(2), "##,##0.00")
End Sub

Private Sub txtHazardAdj_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtHazardAdj.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtHazardAdj
End Sub

Private Sub txtHazardAdj_LostFocus() ' updated Dec. 3, 2016
    txtAdj(3) = IIf(txtHazardAdj.Text = "", 0, ChkVal(txtHazardAdj)) * HazardPayRate
    Call SumAdjustment(txtAdj(3))
    txtAdj(3).BackColor = &HFFFFFF
    txtHazardAdj.BackColor = &HFFFFFF
    txtAdj(3) = Format(txtAdj(3), "##,##0.00")
End Sub

Private Sub SumAdjustment(ByVal mAmt As Double)
Dim DailyEcola As Double

    If PaymentBasis = "Monthly" Then ' if monthly paid get the fixed Ecola rate
        DailyEcola = ((mECOLARate * 2 * 12) / WorkDaysperYear)
    Else
        DailyEcola = mECOLARate
    End If
    txtAdj(2) = (Val(txtUnpaidVL) + Val(txtUnpaidSL)) * DailyEcola
    txtAdj(3) = (Val(txtUnpaidVL) + Val(txtUnpaidSL)) * HazardPayRate
    
    txtAdj(3) = txtHazardAdj.Text
    
    TxtAmount(21) = ChkVal(txtAdj(0)) + ChkVal(txtAdj(1)) + ChkVal(txtAdj(2)) + ChkVal(txtAdj(3)) + ChkVal(txtAdj(4)) '+ mAmt
    txtAdj(2) = Format(txtAdj(2), "##,##0.00")
    txtAdj(3) = Format(txtAdj(3), "##,##0.00")
    txtAdj(4) = Format(txtAdj(4), "##,##0.00")
    
    ' Update the running total.
    Call AccumulateAmount
End Sub

Private Sub txtVacLeave_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtVacLeave.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtVacLeave
End Sub

Private Sub txtVacLeave_LostFocus()
Dim mTot As Double

    TxtAmount(8).BackColor = &HFFFFFF
    TxtAmount(8).Text = Format(ComputeAmount(txtVacLeave.Text, PaymentBasis, False, 10), "###,##0.00")
    txtVacLeave.BackColor = &HFFFFFF
    
    If PaymentBasis = "Monthly" Then
'        mTot = Val(Replace(TxtAmount(5).Text, ",", "")) + Val(Replace(TxtAmount(6).Text, ",", "")) + _
'            Val(Replace(TxtAmount(7).Text, ",", "")) + Val(Replace(TxtAmount(8).Text, ",", ""))
'        TxtAmount(0).Text = Format(OrigMonthlyRate - mTot, "###,#00.00")
        TxtAmount(0).Text = Format(OrigMonthlyRate - (Val(Replace(TxtAmount(7).Text, ",", "")) + Val(Replace(TxtAmount(8).Text, ",", ""))), "###,#00.00")
    End If
    Call ComputeForECOLA(1) ' updated Sep. 1, 2016 add Ecola rate for restday
    Call AccumulateAmount(False, False, True)
End Sub

Private Sub txtHazardPay_GotFocus()
    HazardFlag = True
    Sendkeys "{HOME}+{END}"
    txtHazardPay.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtHazardPay
End Sub

Private Sub txtHazardPay_LostFocus()
    txtHazardPay.BackColor = &HFFFFFF
    txtHazardPay.Text = Format(txtHazardPay.Text, "###,##0.00")
    HazardFlag = False
    Call AccumulateAmount
End Sub

Private Sub txtUnionDues_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtUnionDues.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtUnionDues
End Sub

Private Sub txtUnionDues_LostFocus()
    txtUnionDues.BackColor = &HFFFFFF
    Call AccumulateDeductions
    txtUnionDues.Text = Format(txtUnionDues.Text, "###,##0.00")
End Sub

Private Sub txtVale_GotFocus()
    Sendkeys "{HOME}+{END}"
    txtVale.BackColor = &HFFFFC0
    Set oCommonFunction.ValidDouble = txtVale
End Sub

Private Sub txtVale_LostFocus()
    txtVale.BackColor = &HFFFFFF
    Call AccumulateDeductions
    txtVale.Text = Format(txtVale.Text, "##,##0.00")
End Sub

Private Function GetDedAmount(mRs As Recordset, mDedType As String) As Double
Dim mDedList As String
Dim IncludeFlag As Boolean
Dim StartCutoffDate, EndCutoffDate, StartDate, LastDate

    GetDedAmount = 0#
    With mRs
        .Filter = "EmployeeID = " & EmployeeNo
    End With
    StartCutoffDate = Val(Format(frmSelectCutoff.txtCutOffDate(0).Text, "yyyymmdd"))
    EndCutoffDate = Val(Format(frmSelectCutoff.txtCutOffDate(1).Text, "yyyymmdd"))
    While Not (mRs.EOF Or mRs.BOF)
        IncludeFlag = True
        If mRs!DeductionFlag = True Then 'check if payment was based from no of payments
            If mRs!DeductionCounter < mRs!NoOfDeductions Then
                IncludeFlag = True
            Else
                IncludeFlag = False
            End If
        Else    ' Deductions was based from End Date
            StartDate = Val(Format(mRs!DedStartpayment, "yyyymmdd"))
            LastDate = Val(Format(mRs!DedEndPayment, "yyyymmdd"))
            If StartDate = 19000101 Then StartDate = 99999999
            If LastDate = 19000101 Then LastDate = 99999999
            ' Validate whether entry dates fall on cut-off dates.
            If (StartCutoffDate >= StartDate And StartCutoffDate <= LastDate) Or _
            (EndCutoffDate >= StartDate And EndCutoffDate <= LastDate) Then
                IncludeFlag = True
            Else
                IncludeFlag = False
            End If
        End If
        
        If IncludeFlag = True Then  ' deduction is included
        
        End If
    Wend
    
End Function

Public Sub SaveTransaction()
Dim rsPayrollCheck As ADODB.Recordset
Dim mECCAmt As Double
Dim mOtherLoan As Double
Dim StartDate ' As Date ' remove by bernard
Dim EndDate  ' As Date
Dim StartDate1
Dim EndDate1
Dim mFortheMonth As Integer
Dim mFortheYear As Integer
Dim mTotalEmp As Integer
Dim ResetFlag As Boolean
Dim WithDuplicate As Boolean
Dim strQuery1 As String

    On Error GoTo Errh
    
    WithDuplicate = False
    'Call txtHoursLate_LostFocus     ' check for possible entries in late deductions  commented Feb. 15, 2017
    If GetWindowsVersion() = "Windows 10" Then
        StartDate = frmSelectCutoff.txtCutOffDate(0).Text
        EndDate = frmSelectCutoff.txtCutOffDate(1).Text
    End If
    ' Updated 3/1/2020 Bernard.
    StartDate = Format(Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy"), "mm/dd/yyyy")
    EndDate = Format(Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy"), "mm/dd/yyyy")
    StartDate1 = Format(Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy"), "yyyymmdd")
    EndDate1 = Format(Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy"), "yyyymmdd")
    'mFortheMonth = Format(frmSelectCutoff.txtCutOffDate(0).Text, "m")
    ' Use the end cutoff date 2/20/2017.
    'mFortheMonth = Format(frmSelectCutoff.txtCutOffDate(1).Text, "m") ' original code
    mFortheMonth = Format(Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy"), "m")
    
    mFortheYear = Format(frmSelectCutoff.txtCutOffDate(1).Text, "yyyy")
    mTotalEmp = rsPayrollTrans.RecordCount

    txtNetPay = Format(txtGrossPay.Text - txtDeductions(2).Text, "###,##0.00")
    mOtherLoan = Val(txtIncentive) ' for BSC Allowance
    ' Save ECC Amount 8/6/2009.
    mECCAmt = Val(txtECC.Text)
    
    ' Validate first if employee already exists with the same Client Code.
    StrQuery = "SELECT EmployeeID, ClientCode FROM PayrollTrans WHERE EmployeeID = " & EmployeeNo
    Set rsPayrollCheck = ReturnRst(StrQuery, cnnstr, mstatic)
    If rsPayrollCheck.EOF = False Then
        WithDuplicate = True
        ResetFlag = AddFlag
        AddFlag = False ' proceed with updates
    End If
    'OverUnderPay was changed to Uniform 2/1/2010.
    ' MiscFee was changed to Vale_IOU. 4/2/2011
    ' Updated Feb. 15, 2015 to include Variances for Adjustment.
    ' Updates 9/22/2018 use Medical for SSS CAlamity Loan.
    If AddFlag = True Then  ' save new entries
        StrQuery = "INSERT INTO PayrollTrans (EmployeeID, EmpNo, employeename, agencycode, clientcode, NoofDays, dayspresent, basicpay, " & _
            "othours, otpay, excessothours, excessotpay, nightdiffot, nightdiffpay, month13rate, month13payment, " & _
            "TaxRefundHours, TaxRefundPay, sundayot, sundayotpay, NoRegHoliday, RegHolidayPay, Variances, Medical, " & _
            "splholidayothours, splholidayotpay, excesssplholothours, excesssplholotpay, " & _
            "NoSplHoliday, SplHolidayPay, lglholidayothours, lglholidayotpay, excesslglholothours, excesslglholotpay, excessotsunlegalhrs, excessotsunlegalpay, otwithoutpremium, otwithoutpremiumpay, novacationleave, " & _
            "vacationleavepay, nosickleave, sickleavepay, OtherIncome, Allowance, incentiveremarks, ecolapay, grosspay, daysabsent, absentdeduction, hourstardiness, tardinessdeduction, " & _
            "TASAesla, LMCDed, ValeWithClient, CashAdvance, Vale_IOU, Uniform, HazardPay, UnionDues, " & _
            "sssemployee, sssemployer, eccemployer, phealthemployee, phealthemployer, pagibigemployee, " & _
            "pagibigemployer, wtaxamount, CalamityLoan, sssloan, pagibigloan, otherloan, deductions, netpay, " & _
            "startpaydate, endpaydate, monthno, cutoffno, payflag1, payflag2, paytext1, paytext2, entryby, " & IIf(InsField = "", "", InsField & ",") & " entrydate )" & _
            " VALUES (" & EmployeeNo & ", '" & EmpNo & "', '" & EmployeeName & "'," & CompanyCode & ", " & mClientCode & ", " & ChkVal(txtNoDays) & ", " & ChkVal(txtDaysWork) & ", " & ChkVal(TxtAmount(0)) & _
            ", " & ChkVal(txtRegOT) & ", " & ChkVal(TxtAmount(1)) & ", " & ChkVal(txtExcessOTRegDay) & ", " & ChkVal(TxtAmount(12)) & ", " & ChkVal(txtNightDiff) & ", " & ChkVal(TxtAmount(2)) & ", " & ChkVal(txt13thMonth) & ", " & ChkVal(TxtAmount(11)) & _
            ", " & ChkVal(txtTaxRefund) & ", " & ChkVal(TxtAmount(3)) & ", " & ChkVal(txtSundayOT) & ", " & ChkVal(TxtAmount(4)) & ", " & ChkVal(txtSplHoliday) & ", " & ChkVal(TxtAmount(5)) & ", " & ChkVal(TxtAmount(21)) & ", " & ChkVal(txtCalLoan) & _
            ", " & ChkVal(txtSplOtHours) & ", " & ChkVal(TxtAmount(4)) & ", " & ChkVal(txtRDonSPL) & ", " & ChkVal(TxtAmount(19)) & _
            ", " & ChkVal(txtLglHoliday) & ", " & ChkVal(TxtAmount(6)) & ", " & ChkVal(txtLglHolidayOT) & ", " & ChkVal(TxtAmount(4)) & ", " & ChkVal(txtExcessOTLegal) & ", " & ChkVal(txtHDMFCal) & ", " & ChkVal(txtExcessOTRegDay) & ", " & ChkVal(TxtAmount(12)) & ", " & ChkVal(txtRDonLGL) & ", " & ChkVal(TxtAmount(20)) & ", " & ChkVal(txtVacLeave) & _
            ", " & ChkVal(TxtAmount(8)) & ", " & ChkVal(txtSickLeave) & ", " & ChkVal(TxtAmount(7)) & ", " & ChkVal(txtSalAdjustment) & ", " & ChkVal(TxtAmount(10)) & ", '" & Trim(txtIncentiveRemarks) & "', " & ChkVal(txtECOLA) & ", " & ChkVal(txtGrossPay) & ", " & ChkVal(txtDaysAbsent) & ", " & ChkVal(txtDeductions(0)) & ", " & ChkVal(txtHoursLate) & ", " & ChkVal(txtDeductions(1)) & _
            ", " & ChkVal(txtTASAesla) & ", " & ChkVal(txtBDOCard) & ", " & ChkVal(txtVale) & ", " & ChkVal(txtPlacementFee) & ", " & ChkVal(txtMiscFee) & ", " & ChkVal(txtOverPay) & ", " & ChkVal(txtHazardPay) & ", " & ChkVal(txtUnionDues) & _
            ", " & ChkVal(txtEmpyeeSSS) & ", " & ChkVal(txtEmpyerSSS) & ", " & Format(mECCAmt, "0.00") & ", " & ChkVal(txtPhilHealthYee) & ", " & ChkVal(txtPHealthYer) & ", " & ChkVal(txtEmpyeePagibig) & _
            ", " & ChkVal(txtEmpyerPagibig) & ", " & ChkVal(txtTax) & ", " & ChkVal(txtCalamityLoan) & ", " & ChkVal(txtSSSLoan) & ", " & ChkVal(txtPagibigLoan) & ", " & Format(mOtherLoan, "0.00") & ", " & ChkVal(txtDeductions(2)) & ", " & ChkVal(txtNetPay) & _
            ", '" & StartDate & "', '" & EndDate & "', " & Format(frmSelectCutoff.txtCutOffDate(0).Text, "m") & ", " & mCutOffNo & ", 0, 0,'','', '" & mUserID & "', " & IIf(InsValue = "", " '", InsValue & ", '") & Format(Now, "mm/dd/yyyy") & "')"
    
            '", '" & Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy") & "', '" & Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy") & "', " & Format(frmSelectCutoff.txtCutOffDate(0).Text, "m") & ", " & mCutOffNo & ", 0, 0,'','', '" & mUserID & "', " & IIf(InsValue = "", " '", InsValue & ", '") & Format(Now, "mm/dd/yyyy") & "')"
    
    
        strQuery1 = "INSERT INTO PayrollAdjust (EmployeeID2, EmpNo2, agencycode2, clientcode2, unpaidvl, unpaidvldays, unpaidsl, unpaidsldays, ecolaadj, hazardadj, regrate, startpaydate2, endpaydate2)" & _
            " VALUES (" & EmployeeNo & ", '" & EmpNo & "', " & CompanyCode & ", " & mClientCode & ", " & ChkVal(txtAdj(0)) & ", " & ChkVal(txtUnpaidVL) & ", " & ChkVal(txtAdj(1)) & ", " & ChkVal(txtUnpaidSL) & ", " & ChkVal(txtAdj(2)) & ", " & ChkVal(txtAdj(3)) & ", " & ChkVal(txtAdj(4)) & _
            ", '" & StartDate & "', '" & EndDate & "')"
            
            '", '" & Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy") & "', '" & Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy") & "')"


'Note txtrdspllgl was replace by txtExcessOTRegDay 3/12/2017
    
    Else    ' Update Entries
        If EditFlag = True Then ' Updated Feb. 1, 2015 do not re-update the data
            StrQuery = "UPDATE PayrollTrans SET employeename = '" & EmployeeName & "', agencycode = " & CompanyCode & ", clientcode = " & mClientCode & ", NoOfDays = " & ChkVal(txtNoDays) & ", Dayspresent = " & ChkVal(txtDaysWork) & ", basicpay = " & ChkVal(TxtAmount(0)) & _
                ", othours = " & ChkVal(txtRegOT) & ", otpay = " & ChkVal(TxtAmount(1)) & ", excessothours = " & ChkVal(txtExcessOTRegDay) & ", excessotpay = " & ChkVal(TxtAmount(12)) & ", nightdiffot = " & ChkVal(txtNightDiff) & ", nightdiffpay = " & ChkVal(TxtAmount(2)) & ", month13rate = " & ChkVal(txt13thMonth) & ", month13payment = " & ChkVal(TxtAmount(11)) & _
                ", TaxRefundHours = " & ChkVal(txtTaxRefund) & ", TaxRefundPay = " & ChkVal(TxtAmount(3)) & ", sundayot = " & ChkVal(txtSundayOT) & ", sundayotpay = " & ChkVal(TxtAmount(4)) & ", NoRegHoliday = " & ChkVal(txtSplHoliday) & ", RegHolidayPay = " & ChkVal(TxtAmount(5)) & ", Variances = " & ChkVal(TxtAmount(21)) & ", Medical = " & ChkVal(txtCalLoan) & _
                ", splholidayothours = " & ChkVal(txtSplOtHours) & ", splholidayotpay = " & ChkVal(TxtAmount(4)) & ", excesssplholothours = " & ChkVal(txtRDonSPL) & ", excesssplholotpay = " & ChkVal(TxtAmount(19)) & _
                ", NoSplHoliday = " & ChkVal(txtLglHoliday) & ", SplHolidayPay = " & ChkVal(TxtAmount(6)) & ", lglholidayothours = " & ChkVal(txtLglHolidayOT) & ", lglholidayotpay = " & ChkVal(TxtAmount(4)) & ", excesslglholothours = " & ChkVal(txtExcessOTLegal) & ", excesslglholotpay = " & ChkVal(txtHDMFCal) & ", excessotsunlegalhrs = " & ChkVal(txtExcessOTRegDay) & ", excessotsunlegalpay = " & ChkVal(TxtAmount(12)) & ", otwithoutpremium = " & ChkVal(txtRDonLGL) & ", otwithoutpremiumpay = " & ChkVal(TxtAmount(20)) & ", novacationleave = " & ChkVal(txtVacLeave) & _
                ", vacationleavepay = " & ChkVal(TxtAmount(8)) & ", nosickleave = " & ChkVal(txtSickLeave) & ", sickleavepay = " & ChkVal(TxtAmount(7)) & ", OtherIncome = " & ChkVal(txtSalAdjustment) & ", Allowance = " & ChkVal(TxtAmount(10)) & ", incentiveremarks = '" & Trim(txtIncentiveRemarks) & "', ecolapay = " & ChkVal(txtECOLA) & ", grosspay = " & ChkVal(txtGrossPay) & ", daysabsent = " & ChkVal(txtDaysAbsent) & ", absentdeduction = " & ChkVal(txtDeductions(0)) & ", hourstardiness = " & ChkVal(txtHoursLate) & ", tardinessdeduction = " & ChkVal(txtDeductions(1)) & _
                ", TASAesla = " & ChkVal(txtTASAesla) & ", LMCDed = " & ChkVal(txtBDOCard) & ", valewithclient = " & ChkVal(txtVale) & ", CashAdvance = " & ChkVal(txtPlacementFee) & ", Vale_IOU = " & ChkVal(txtMiscFee) & ", Uniform = " & ChkVal(txtOverPay) & ", HazardPay = " & ChkVal(txtHazardPay) & ", UnionDues = " & ChkVal(txtUnionDues) & _
                ", sssemployee = " & ChkVal(txtEmpyeeSSS) & ", sssemployer = " & ChkVal(txtEmpyerSSS) & ", eccemployer = " & Format(mECCAmt, "0.00") & ", phealthemployee = " & ChkVal(txtPhilHealthYee) & ", phealthemployer = " & ChkVal(txtPHealthYer) & ", pagibigemployee = " & ChkVal(txtEmpyeePagibig) & _
                ", pagibigemployer = " & ChkVal(txtEmpyerPagibig) & ", wtaxamount = " & ChkVal(txtTax) & ", CalamityLoan = " & ChkVal(txtCalamityLoan) & ", sssloan = " & ChkVal(txtSSSLoan) & ", pagibigloan = " & ChkVal(txtPagibigLoan) & ", otherloan = " & Format(mOtherLoan, "0.00") & ", deductions = " & ChkVal(txtDeductions(2)) & ", netpay = " & ChkVal(txtNetPay) & IIf(EdtField = "", "", ", " & EdtField) & _
                ", startpaydate = '" & Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy") & "', endpaydate = '" & Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy") & "', monthno = " & Format(frmSelectCutoff.txtCutOffDate(0).Text, "m") & ", cutoffno = " & mCutOffNo & ", payflag1 = 0, payflag2 = 0, paytext1 = '', paytext2 = '', entryby = '" & mUserID & "', entrydate = '" & Format(Now, "mm/dd/yyyy") & "' WHERE EmployeeID = " & EmployeeNo
        
            strQuery1 = "UPDATE PayrollAdjust SET agencycode2 = " & CompanyCode & ", clientcode2 = " & mClientCode & ", unpaidvl = " & ChkVal(txtAdj(0)) & ", unpaidvldays = " & ChkVal(txtUnpaidVL) & ", unpaidsl = " & ChkVal(txtAdj(1)) & ", unpaidsldays = " & ChkVal(txtUnpaidSL) & ", ecolaadj = " & ChkVal(txtAdj(2)) & ", hazardadj = " & ChkVal(txtAdj(3)) & ", RegRate = " & ChkVal(txtAdj(4)) & _
                ", startpaydate2 = '" & Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy") & "', endpaydate2 = '" & Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy") & "' WHERE EmployeeID2 = " & EmployeeNo
                '", NoOfDays = " & ChkVal(txtNoDays) & ", startpaydate2 = '" & Format(frmSelectCutoff.txtCutOffDate(0).Text, "mm/dd/yyyy") & "', endpaydate2 = '" & Format(frmSelectCutoff.txtCutOffDate(1).Text, "mm/dd/yyyy") & "' WHERE EmployeeID2 = " & EmployeeNo

        Else
            MsgBox "Records was updated!", vbInformation, PayrollAlert
        End If
    End If
    
    ' Updates Jan 17, 2015
    ' excesssplholothours = " & ChkVal(txtRDonSPL) & ", excesssplholotpay = " & ChkVal(txtAmount(19))
    ' excessotsunlegalhrs = " & ChkVal(txtRDSplLgl) & ", excessotsunlegalpay = " & ChkVal(txtAmount(18))
    ' otwithoutpremium = " & ChkVal(txtRDonLGL) & ", otwithoutpremiumpay = " & ChkVal(txtAmount(20))
    
    ' txtamount(15) = excesslglholotpay updated August 24, 2013.
    'miscfee was replace by vale_iou 3/21/2011
    If WithDuplicate = True Then
        AddFlag = ResetFlag
    End If
    rsPayrollCheck.Close
    
    'Call Action_Query(strQuery, cnnStr)
    If (AddFlag Or EditFlag) Then
        If Action_Query(StrQuery, cnnstr) = True Then   ' successfully inserted/updated
            If strQuery1 <> "" Then
                Call Action_Query(strQuery1, cnnstr)  ' update payroll adjustments
            End If
            MsgBox "Entries Saved!!!", vbInformation, PayrollAlert
            If frmBasicRates.Visible = True Then
                frmBasicRates.Visible = False
            End If
            'Update Cutoff Table with new CutOff.
            If AddFlag = True And frmSelectCutoff.chkNewCutoff.VALUE = 1 And InsertedFlag = False Then
                StrQuery = "INSERT INTO cutofftable ( AgencyCode, ClientCode, ClientName, FortheMonth, FortheYear, StartCutOffDate, EndCutOffDate, " & _
                    "CutOffNo, BilledFlag, TotalEmployee, ClientSeriesNo, Processedby, StartCutOffText, EndCutOffText ) " & _
                    " VALUES ( " & CompanyCode & ", " & mClientCode & ", '" & mCompanyName & "', " & mFortheMonth & ", " & mFortheYear & _
                    ", '" & Format(StartDate, "mm/dd/yyyy") & "', '" & Format(EndDate, "mm/dd/yyyy") & "', " & mCutOffNo & ", 0, " & mTotalEmp & ", '" & frmSelectCutoff.txtSeriesNo & "', '" & mUserID & "', '" & StartDate1 & "', '" & EndDate1 & "')"
                    Call Action_Query(StrQuery, cnnstr)
                    InsertedFlag = True
            End If
        'Update Employee Basic Salary table processed flag.
        StrQuery = "UPDATE employeebasicsalary SET processedflag = 1 WHERE EmployeeID = " & EmployeeNo
        Call Action_Query(StrQuery, cnnstr)
        Call PayrollTransRec
        Call GetEmployeeBasic
        Call UpdateEmployeeDeductionTable(EmployeeNo, True)  ' update the employeedeductiontable
        Else
            MsgBox "Unable to Save the transaction record!!!", vbInformation, PayrollAlert
        End If
    End If
    AgencySw = False
    'rsEmployeeInfo.Requery
    Exit Sub
    
Errh:

    If AddFlag = True Then
        MsgBox "Error encountered cannot save new entry!", vbInformation, PayrollAlert
    Else
        MsgBox "Error encountered cannot save edits!", vbInformation, PayrollAlert
    End If
    Err.Clear
End Sub

Private Sub DeleteTransaction()
Dim strQuery1 As String

'Delete the record from the Payroll master.
StrQuery = "DELETE FROM PayrollTrans WHERE EmployeeID = " & EmployeeNo
' Delete records from Payroll Adjustment table.
strQuery1 = "DELETE FROM PayrollAdjust WHERE EmployeeID2 = " & EmployeeNo
'Call Action_Query(strQuery, cnnStr)
If Action_Query(StrQuery, cnnstr) = True Then
    Call Action_Query(strQuery1, cnnstr)
    rsPayrollTrans.Requery
    'Uncheck Processed flag in Employee Basic Salary table.
    StrQuery = "UPDATE employeebasicsalary SET processedflag = 0 WHERE EmployeeID = " & EmployeeNo
    Call Action_Query(StrQuery, cnnstr)
    Call UpdateEmployeeDeductionTable(EmployeeNo, False)  ' Delete the deductions made from employeedeductiontable
Else
    MsgBox "Unable to delete transaction record!!!", vbInformation, PayrollAlert
End If
'rsEmployeeInfo.Requery
Call GetEmployeeBasic
End Sub

Private Function ChkVal(ByRef mTBox As TextBox) As Double
Dim i As Integer
Dim strAmt As String
Dim mAmt As String
Dim DeciSw As Boolean

    mAmt = Trim(Replace(mTBox, ",", ""))
    For i = 1 To Len(mAmt)
        If DeciSw = True And mID(mAmt, i, 1) = "." Then
        Else
            strAmt = strAmt & mID(mAmt, i, 1)
        End If
        If DeciSw = False And mID(mAmt, i, 1) = "." Then
            DeciSw = True
        End If
    Next
    mAmt = IIf(strAmt = "." Or strAmt = "0" Or strAmt = "", 0#, strAmt)
    ChkVal = Format(mAmt, "0.00")

End Function

Private Sub PayrollTransRec()   ' Update payroll transaction record
Dim StartDate As String
Dim EndDate As String

Set rsPayrollTrans = New ADODB.Recordset
StartDate = frmSelectCutoff.txtCutOffDate(0).Text
EndDate = frmSelectCutoff.txtCutOffDate(1).Text
'StrQuery = "SELECT * FROM PayrollTrans WHERE clientcode = " & mClientCode & _
'        " and startpaydate = '" & StartDate & "' and endpaydate = '" & EndDate & "'"
' Updated Dec. 3, 2016
StrQuery = "SELECT a.*, b.* FROM PayrollTrans a " & _
        " LEFT OUTER JOIN PayrollAdjust b ON a.EmpNo = b.EmpNo2 AND a.ClientCode = b.ClientCode2 AND a.StartPaydate = b.StartPayDate2 AND a.EndPaydate = b.EndPaydate2 " & _
        " WHERE clientcode = " & mClientCode & _
        " and startpaydate = '" & StartDate & "' and endpaydate = '" & EndDate & "'"

Set rsPayrollTrans = ReturnRst(StrQuery, cnnstr, mstatic)

End Sub

Private Sub PayrollMasterRec()   ' Get previous payroll transaction record
Dim StartDate As String
Dim EndDate As String

Set rsPayrollTrans = New ADODB.Recordset
StartDate = frmSelectCutoff.txtCutOffDate(0).Text
EndDate = frmSelectCutoff.txtCutOffDate(1).Text
'StrQuery = "SELECT * FROM PayrollMaster WHERE clientcode = " & mClientCode & _
'        " and startpaydate = '" & StartDate & "' and endpaydate = '" & EndDate & "'"
' Updated Dec. 3, 2016
StrQuery = "SELECT a.*, b.* FROM PayrollMaster a " & _
        " LEFT OUTER JOIN MasterAdjust b ON a.EmpNo = b.EmpNo2 AND a.ClientCode = b.ClientCode2 AND a.StartPaydate = b.StartPayDate2 AND a.EndPaydate = b.EndPaydate2 " & _
        " WHERE clientcode = " & mClientCode & _
        " and startpaydate = '" & StartDate & "' and endpaydate = '" & EndDate & "'"

Set rsPayrollTrans = ReturnRst(StrQuery, cnnstr, mstatic)

End Sub

