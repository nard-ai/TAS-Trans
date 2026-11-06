VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPrintPayslip 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Report"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmPrintPayslip.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5700
   Begin VB.Frame fraButton 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   375
      Left            =   1395
      TabIndex        =   39
      Top             =   4140
      Width           =   2895
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
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
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
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
         Left            =   1620
         TabIndex        =   40
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   3975
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   5370
      Begin VB.Frame Frame2 
         Caption         =   "Select Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   180
         Width           =   5145
         Begin VB.OptionButton optReport 
            Caption         =   "Payroll Summary Short"
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
            Index           =   9
            Left            =   2620
            TabIndex        =   44
            Top             =   1515
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.OptionButton optReport 
            Caption         =   "DTR Summary Report"
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
            Index           =   8
            Left            =   2610
            TabIndex        =   43
            Top             =   1335
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Employee Summary"
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
            Index           =   7
            Left            =   2610
            TabIndex        =   9
            Top             =   1200
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Employee Deductions"
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
            Index           =   6
            Left            =   2520
            TabIndex        =   8
            Top             =   690
            Width           =   2295
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Employee Payroll List"
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
            Index           =   5
            Left            =   180
            TabIndex        =   7
            Top             =   1185
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Mandatory Deductions"
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
            Index           =   4
            Left            =   2520
            TabIndex        =   6
            Top             =   270
            Width           =   2445
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Payroll Summary Report"
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
            Index           =   3
            Left            =   2520
            TabIndex        =   5
            Top             =   1095
            Width           =   2535
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Payroll Bank Copy"
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
            Index           =   2
            Left            =   180
            TabIndex        =   4
            Top             =   1095
            Width           =   2535
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Statement of Account"
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
            Index           =   1
            Left            =   180
            TabIndex        =   3
            Top             =   690
            Width           =   2295
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Payroll Slip"
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
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   270
            Value           =   -1  'True
            Width           =   2445
         End
      End
      Begin VB.Frame FraDebit 
         Height          =   1770
         Left            =   90
         TabIndex        =   23
         Top             =   2025
         Visible         =   0   'False
         Width           =   5145
         Begin VB.CheckBox chkDebit 
            Caption         =   "Show logo"
            Height          =   240
            Left            =   3960
            TabIndex        =   35
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.ComboBox cboBank 
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
            Height          =   315
            Left            =   3015
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   630
            Width           =   2070
         End
         Begin VB.TextBox TxtDMNum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3015
            TabIndex        =   29
            Text            =   "0"
            Top             =   1035
            Width           =   1275
         End
         Begin VB.TextBox TxtSANum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1170
            TabIndex        =   28
            Text            =   "35-012-4655"
            Top             =   990
            Width           =   1185
         End
         Begin VB.TextBox TxtAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1170
            TabIndex        =   26
            Top             =   585
            Width           =   1185
         End
         Begin MSMask.MaskEdBox txtCreditDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   1170
            TabIndex        =   30
            Top             =   1350
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   0
            AllowPrompt     =   -1  'True
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
         Begin MSMask.MaskEdBox txtCutOffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1170
            TabIndex        =   24
            Top             =   225
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   0
            AllowPrompt     =   -1  'True
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
         Begin MSMask.MaskEdBox txtCutOffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3060
            TabIndex        =   25
            Top             =   225
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   0
            AllowPrompt     =   -1  'True
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
         Begin VB.Label lblEntries 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
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
            Index           =   5
            Left            =   2520
            TabIndex        =   38
            Top             =   675
            Width           =   1725
         End
         Begin VB.Label lblEntries 
            BackStyle       =   0  'Transparent
            Caption         =   "CA/SA No."
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
            Index           =   1
            Left            =   135
            TabIndex        =   37
            Top             =   1035
            Width           =   1050
         End
         Begin VB.Label lblEntries 
            BackStyle       =   0  'Transparent
            Caption         =   "DM No."
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
            Index           =   3
            Left            =   2475
            TabIndex        =   36
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label lblEntries 
            BackStyle       =   0  'Transparent
            Caption         =   "Crediting Date"
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
            Index           =   4
            Left            =   90
            TabIndex        =   34
            Top             =   1395
            Width           =   1050
         End
         Begin VB.Label lblEntries 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount:"
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
            Index           =   8
            Left            =   135
            TabIndex        =   33
            Top             =   630
            Width           =   1185
         End
         Begin VB.Label lblEntries 
            BackStyle       =   0  'Transparent
            Caption         =   "Cut off From:"
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
            Index           =   7
            Left            =   135
            TabIndex        =   32
            Top             =   270
            Width           =   1185
         End
         Begin VB.Label lblEntries 
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
            Height          =   255
            Index           =   6
            Left            =   2520
            TabIndex        =   31
            Top             =   285
            Width           =   465
         End
      End
      Begin VB.Frame FraStatement 
         Height          =   1770
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   5145
         Begin VB.CheckBox chkStatement 
            Caption         =   "Show logo"
            Height          =   240
            Left            =   3960
            TabIndex        =   20
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.TextBox txtSOA 
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
            Left            =   1260
            TabIndex        =   17
            Text            =   "0"
            Top             =   990
            Width           =   2475
         End
         Begin VB.TextBox TxtAmount1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1260
            TabIndex        =   15
            Top             =   630
            Width           =   1185
         End
         Begin MSMask.MaskEdBox mskDueDate 
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
            Left            =   1260
            TabIndex        =   19
            Top             =   1350
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   0
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
         Begin MSMask.MaskEdBox txtCutOffDate1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1260
            TabIndex        =   12
            Top             =   270
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   0
            AllowPrompt     =   -1  'True
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
         Begin MSMask.MaskEdBox txtCutOffDate1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3150
            TabIndex        =   13
            Top             =   270
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   0
            AllowPrompt     =   -1  'True
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
         Begin VB.Label lblEntries 
            BackStyle       =   0  'Transparent
            Caption         =   "Due Date:"
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
            Index           =   0
            Left            =   225
            TabIndex        =   22
            Top             =   1395
            Width           =   1185
         End
         Begin VB.Label lblEntries 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount:"
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
            Index           =   9
            Left            =   225
            TabIndex        =   21
            Top             =   675
            Width           =   1185
         End
         Begin VB.Label lblEntries 
            BackStyle       =   0  'Transparent
            Caption         =   "Cut off From:"
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
            Index           =   10
            Left            =   225
            TabIndex        =   18
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label lblEntries 
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
            Height          =   255
            Index           =   11
            Left            =   2610
            TabIndex        =   16
            Top             =   330
            Width           =   465
         End
         Begin VB.Label lblEntries 
            BackStyle       =   0  'Transparent
            Caption         =   "SOA No."
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
            Index           =   12
            Left            =   225
            TabIndex        =   14
            Top             =   1035
            Width           =   1185
         End
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Employee:"
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
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   2790
         Width           =   1725
      End
   End
   Begin VB.ComboBox cmbEmployee 
      BackColor       =   &H00FFFFC0&
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
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   4230
      Visible         =   0   'False
      Width           =   2835
   End
End
Attribute VB_Name = "frmPrintPayslip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCutOffTable As ADODB.Recordset
Dim rsCurrPayrollTrans As ADODB.Recordset
Dim rsGetDTRreport As ADODB.Recordset
Dim OnProcessFlag As Boolean
Dim NoDaysperCutoff As Integer
Dim StrQuery As String
Dim PrevSOANo As String
Dim mSapi As String
Dim pageset As Boolean
Dim mSeriesNo As String

Private Sub cboBank_GotFocus()
  SendKeys "{HOME}+{END}"
  cboBank.BackColor = &HFFFFC0
End Sub

Private Sub cboBank_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cboBank_LostFocus()
cboBank.BackColor = &HFFFFFF
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strPayTrans As String
    Dim rsPay As ADODB.Recordset
    Dim StartDate As String
    Dim EndDate As String
    Dim OtherDed As Double
    Dim OtherDesc As String
    
    OtherDed = 0#
    OtherDesc = ""
    mSeriesNo = frmSelectCutoff.txtSeriesNo.Text    ' update the series no
    StartDate = frmSelectCutoff.txtCutoffDate(0).Text
    EndDate = frmSelectCutoff.txtCutoffDate(1).Text
    If optReport(0).VALUE = True Then               ' for payslip
'        strPayTrans = "Select * from PayrollReport a " & _
'                " WHERE a.clientcode = " & mClientCode & _
'                " and a.startpaydate = '" & StartDate & "' and a.endpaydate = '" & EndDate & "'" & _
'                " ORDER BY employeename"
        strPayTrans = "EXEC GetPayrollEntries 1, 1, " & mClientCode & ", '" & StartDate & "', '" & EndDate & "', '" & mUserID & "'"
        Call Action_Query(strPayTrans, cnnstr)
        
        strPayTrans = "SELECT * FROM tmpPayrollTemplate ORDER BY EmployeeName"
        Set rsPay = ReturnRst(strPayTrans, cnnstr, mstatic)
        If rsPay.EOF = True Then
'            ' Original code.
'            strPayTrans = "SELECT a.*, (Trim(b.surname)+', '+Trim(b.firstname)+' '+IIf(IsNull(b.middlename) Or b.middlename='','',Left(b.middlename,1)+'.')) AS fullname, " & _
'                "b.DailyRate, b.HourlyRate, a.TaxRefundHours+a.sundayot AS restdayhrs, a.TaxRefundPay+a.sundayotpay AS restdaypay " & _
'                "FROM payrollmaster AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID " & _
'                " WHERE a.clientcode = " & mClientCode & _
'                " and a.startpaydate = '" & StartDate & "' and a.endpaydate = '" & EndDate & "'" & _
'                "ORDER BY b.Surname, b.FirstName"
        
        'Updated by Bernard 08/4/2008.
'            strPayTrans = "SELECT a.*, (Trim(b.surname)+', '+Trim(b.firstname)+' '+IIf(IsNull(b.middlename) Or b.middlename='','',Left(b.middlename,1)+'.')) AS fullname, " & _
'                "b.DailyRate, b.HourlyRate, a.TaxRefundHours+a.sundayot AS restdayhrs, a.TaxRefundPay+a.sundayotpay AS restdaypay, " & _
'                "a.uniform + a.medical + a.variances + a.salaryloan + a.vale_iou AS otherded " & _
'                "FROM payrollMaster AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID " & _
'                "WHERE a.clientcode = " & mClientCode & _
'                " and a.startpaydate = '" & StartDate & "' and a.endpaydate = '" & EndDate & "'" & _
'                "ORDER BY b.Surname, b.FirstName"

            strPayTrans = "EXEC GetPayrollEntries 0, 1, " & mClientCode & ", '" & StartDate & "', '" & EndDate & "', '" & mUserID & "'"
            Call Action_Query(strPayTrans, cnnstr)
            strPayTrans = "SELECT * FROM tmpPayrollTemplate ORDER BY EmployeeName"
        
            Set rsPay = Nothing
            Set rsPay = New ADODB.Recordset
            Set rsPay = ReturnRst(strPayTrans, cnnstr, mstatic)
        End If
        
        'If CompanyCode = 1 Then 'SAPI PAYSLIP
        ' All Agency
        If rsPay.EOF = True Then
            GoTo ExitRtn1
        End If
            'With rptPayslipReport_ALL
            With rptPayslipTAS
            Set .DataSource = rsPay
            .DataMember = rsPay.DataMember
'            If rsPay!Allowance = 0 Then
'                .Sections("Section1").Controls("lblIncentive").Visible = True   ' false
'                .Sections("Section1").Controls("txtInceRemarks").Visible = True ' false
'                .Sections("Section1").Controls("txtIncentive").Visible = True   ' false
'            Else
'                .Sections("Section1").Controls("lblIncentive").Visible = True
'                .Sections("Section1").Controls("txtInceRemarks").Visible = True
'                .Sections("Section1").Controls("txtIncentive").Visible = True
'            End If
'
'            If rsPay!NoVacationLeave = 0 Then  ' for slvl updated march 22, 2008
'                .Sections("Section1").Controls("lblSLVL").Visible = False
'                .Sections("Section1").Controls("txtSLVLNo").Visible = False
'                .Sections("Section1").Controls("txtSLVLPay").Visible = False
'            Else
'                .Sections("Section1").Controls("lblSLVL").Visible = True
'                .Sections("Section1").Controls("txtSLVLNO").Visible = True
'                .Sections("Section1").Controls("txtSLVLPay").Visible = True
'            End If
'
''            If rsPay!OverUnderPay = 0 Then  ' for over/under payment updated March 23, 2008
''                .Sections("Section1").Controls("lblOverUnder").Visible = False
''                .Sections("Section1").Controls("txtOverUnder").Visible = False
''                .Sections("Section1").Controls("txtOverUnder").Visible = False
''            Else
'                .Sections("Section1").Controls("lblOverUnder").Visible = True
'                .Sections("Section1").Controls("txtOverUnder").Visible = True
'                .Sections("Section1").Controls("txtOverUnder").Visible = True
'            End If
'            If rsPay!Vale_IOU <> 0 Then
'                OtherDed = OtherDed + rsPay!Vale_IOU
'                OtherDesc = IIf(OtherDesc <> "", OtherDesc & "/IOU", "IOU")
'            End If
'            If rsPay!Uniform <> 0 Then
'                OtherDed = OtherDed + rsPay!Uniform
'                OtherDesc = IIf(OtherDesc <> "", OtherDesc & "/UN", "UN")
'            End If
'            If rsPay!Medical <> 0 Then
'                OtherDed = OtherDed + rsPay!Medical
'                OtherDesc = IIf(OtherDesc <> "", OtherDesc & "/ME", "ME")
'            End If
'            If rsPay!Variances <> 0 Then
'                OtherDed = OtherDed + rsPay!Variances
'                OtherDesc = IIf(OtherDesc <> "", OtherDesc & "/VA", "VA")
'            End If
'            If rsPay!SalaryLoan <> 0 Then
'                OtherDed = OtherDed + rsPay!SalaryLoan
'                OtherDesc = IIf(OtherDesc <> "", OtherDesc & "/SL", "SL")
'            End If
           ' .Sections("Section1").Controls("lblotherDed").Caption = Format(rsPay!OtherDed, "###,##0.00")
         '   .Sections("Section1").Controls("lblDed").Caption = "Others:"
         '   .Sections("Section1").Controls("lblDed").Visible = False
'            If OtherDed = 0 Then
'                .Sections("Section1").Controls("lblOthers").Visible = False
'                .Sections("Section1").Controls("lblotherDed").Visible = False
'                .Sections("Section1").Controls("lblDed").Visible = False
'            Else
'                .Sections("Section1").Controls("lblOthers").Visible = True
'                .Sections("Section1").Controls("lblotherDed").Visible = True
'                .Sections("Section1").Controls("lblDed").Visible = True
'            End If
            'If rsPay!Month13Payment = 0 Then
'            If Display13thPay = False Then
'                .Sections("Section1").Controls("lbl13thMonth").Visible = False
'                .Sections("Section1").Controls("txt13thPay").Visible = False
'            Else
'                .Sections("Section1").Controls("lbl13thMonth").Visible = True
'                .Sections("Section1").Controls("txt13thPay").Visible = True
'            End If
            DoEvents
            .Show vbModal
            End With
     
'        ElseIf CompanyCode = 2 Then 'JCI PAYSLIP
'            With rptPayslipReport_ALL
'            Set .DataSource = rsPay
'            .DataMember = rsPay.DataMember
'            .Sections("Section1").Controls("lblIncentive").Visible = False
'            If rsPay!Allowance = 0 Then
'                .Sections("Section1").Controls("lblIncentive").Visible = True
'                .Sections("Section1").Controls("txtInceRemarks").Visible = True
'                .Sections("Section1").Controls("txtIncentive").Visible = True
'            Else
'                .Sections("Section1").Controls("lblIncentive").Visible = True
'                .Sections("Section1").Controls("txtInceRemarks").Visible = True
'                .Sections("Section1").Controls("txtIncentive").Visible = True
'            End If
'            'If rsPay!Month13Payment = 0 Then
'            If Display13thPay = False Then
'                .Sections("Section1").Controls("lbl13thMonth").Visible = False
'                .Sections("Section1").Controls("txt13thPay").Visible = False
'            Else
'                .Sections("Section1").Controls("lbl13thMonth").Visible = True
'                .Sections("Section1").Controls("txt13thPay").Visible = True
'            End If
'
'            DoEvents
'            .Show vbModal
'            End With
'        End If
        
    ElseIf optReport(1).VALUE = True Then   ' statement of account
        If TxtAmount1.Text = "" Or mskDueDate.Text = "__/__/____" Or txtSOA.Text = "" Or txtCutOffDate1(0).Text = "__/__/____" Or txtCutOffDate1(1).Text = "__/__/____" Then
            MsgBox "Fill up the blank field/s.", vbInformation
            Exit Sub
        End If
        ShowStatement
        Exit Sub
    ElseIf optReport(3).VALUE = True Then   ' for Payroll Register
        'Call GetCurrentPayrollTransaction
        Set rsCurrPayrollTrans = GetPayrollTransactionRecords(True)
        If rsCurrPayrollTrans.RecordCount = 0 Then
            'rsPay.Close
            MsgBox "No employees selected for this payroll cut-off! Cannot proceed.", vbInformation, PayrollAlert
            Exit Sub
        End If
        'Call PayrollSummaryReportNew(rsCurrPayrollTrans, 4)
        Call PayrollSummaryRegister(rsCurrPayrollTrans)
    
    ElseIf optReport(9).VALUE = True Then   ' payroll summary report (Short)
        'Call GetCurrentPayrollTransaction
        Set rsCurrPayrollTrans = GetPayrollTransactionRecords(True)
        If rsCurrPayrollTrans.RecordCount = 0 Then
            'rsPay.Close
            MsgBox "No employees selected for this payroll cut-off! Cannot proceed.", vbInformation, PayrollAlert
            Exit Sub
        End If
        Call PayrollSummaryReportNew(rsCurrPayrollTrans, 4)
    
'    ElseIf optReport(4).VALUE = True Then   ' billing or payroll summary report
'        'Call GetCurrentPayrollTransaction
'        Set rsCurrPayrollTrans = GetPayrollTransactionRecords(True)
'        If rsCurrPayrollTrans.RecordCount = 0 Then
'            'rsPay.Close
'            MsgBox "No employees selected for this payroll cut-off! Cannot proceed.", vbInformation, PayrollAlert
'            Exit Sub
'        End If
'        Call BillingSummaryReport(rsCurrPayrollTrans, 4)
        
    ElseIf optReport(4).VALUE = True Then   ' Mandatory Reports
        Set rsCurrPayrollTrans = GetPayrollTransactionRecords(True)
        If rsCurrPayrollTrans.RecordCount = 0 Then
            MsgBox "No employees selected for this payroll cut-off! Cannot proceed.", vbInformation, PayrollAlert
            Exit Sub
        End If
        Call PayrollSummaryRegister(rsCurrPayrollTrans, "Payroll Contributions.rpt")
        
    ElseIf optReport(6).VALUE = True Then   ' Employee Deductions
        Set rsCurrPayrollTrans = GetPayrollTransactionRecords(True)
        If rsCurrPayrollTrans.RecordCount = 0 Then
            MsgBox "No employees selected for this payroll cut-off! Cannot proceed.", vbInformation, PayrollAlert
            Exit Sub
        End If
        Call PayrollSummaryRegister(rsCurrPayrollTrans, "Employee Deductions.rpt")
        
    ElseIf optReport(2).VALUE = True Then  ' billing or payroll summary report
        'Call GetCurrentPayrollTransaction
        Set rsCurrPayrollTrans = GetPayrollTransactionRecords(True)
        If rsCurrPayrollTrans.RecordCount = 0 Then
            'rsPay.Close
            MsgBox "No employees selected for this payroll cut-off! Cannot proceed.", vbInformation, PayrollAlert
            Exit Sub
        End If
        Call PayrollBankCopy(rsCurrPayrollTrans, 2)
'    ElseIf optReport(6).VALUE = True Then 'DEBIT MEMO
'        If cboBank.Text = "" Then MsgBox "Select bank", vbInformation: cboBank.SetFocus
'        Call showBank
'        ShowBankCopy
'        Exit Sub

   ElseIf optReport(8).VALUE = True Then   ' DTR REPORT
        Call GetDTRreport
        If rsGetDTRreport.RecordCount = 0 Then
            rsGetDTRreport.Close
            MsgBox "No employees selected for this payroll cut-off! Cannot proceed.", vbInformation, PayrollAlert
            Exit Sub
        End If
        Call ShowDtrReport(rsGetDTRreport, 4)
    Else
        MsgBox "Under Development!", vbInformation, PayrollAlert
        Exit Sub
    End If

    Exit Sub

ExitRtn1:

    MsgBox "No records selected!", vbInformation, PayrollAlert
    
End Sub

'Private Sub GetCurrentPayrollTransaction()
'Dim StartDate As String
'Dim EndDate As String
'
'StartDate = frmSelectCutoff.txtCutOffDate(0).Text
'EndDate = frmSelectCutoff.txtCutOffDate(1).Text
'
'Set rsCurrPayrollTrans = New ADODB.Recordset
'
'strQuery = "SELECT a.EmpNo as controlnum, a.*, b.acctno, c.* " & _
'    "FROM (payrolltrans AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID) LEFT JOIN qryEmployeeBasics AS c ON a.EmployeeID = c.EmployeeID " & _
'    "WHERE a.clientcode = " & mClientCode & _
'    " and a.startpaydate = '" & StartDate & "' and a.endpaydate = '" & EndDate & "'" & _
'    " ORDER BY a.employeename"
'
'Set rsCurrPayrollTrans = ReturnRst(strQuery, cnnStr, mstatic)
'
''Check if there were records.
'If rsCurrPayrollTrans.EOF = True Then
'    Set rsCurrPayrollTrans = Nothing
'    Set rsCurrPayrollTrans = New ADODB.Recordset
'
'    strQuery = "SELECT a.EmpNo as controlnum, a.*, b.acctno, c.* " & _
'        "FROM (payrollmaster AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID) LEFT JOIN qryEmployeeBasics AS c ON a.EmployeeID = c.EmployeeID " & _
'        "WHERE a.clientcode = " & mClientCode & _
'        " and a.startpaydate = '" & StartDate & "' and a.endpaydate = '" & EndDate & "'" & _
'        " ORDER BY a.employeename"
'
'    Set rsCurrPayrollTrans = ReturnRst(strQuery, cnnStr, mstatic)
'End If
'
'End Sub

Public Function GetPayrollTransactionRecords(ByVal PayrollSw As Boolean, _
    Optional ClientCd As Integer, Optional mMonth As String = "January", Optional mYear As Integer = 2008) As ADODB.Recordset
Dim rsPayTrans As ADODB.Recordset
Dim StartDate As String
Dim EndDate As String
Dim whereCon As String

If PayrollSw = True Then    ' for current payroll processing
    StartDate = frmSelectCutoff.txtCutoffDate(0).Text
    EndDate = frmSelectCutoff.txtCutoffDate(1).Text
    whereCon = "WHERE a.clientcode = " & mClientCode & _
        " and a.startpaydate = '" & StartDate & "' and a.endpaydate = '" & EndDate & "'" & _
        " ORDER BY a.employeename"
Else    ' for payroll summary per cut-off
    If ClientCd <> 0 Then ' if specific Clients
        mClientCode = ClientCd
        If CutOffPeriod = "START" Then
            whereCon = "WHERE a.clientcode = " & mClientCode & _
                " and format(a.startpaydate, ""mmmm"") = '" & mMonth & "'" & _
                " and format(a.startpaydate, ""yyyy"") = " & mYear & _
                " ORDER BY a.employeename"
        Else    ' Cut-off period is END
            whereCon = "WHERE a.clientcode = " & mClientCode & _
                " and format(a.endpaydate, ""mmmm"") = '" & mMonth & "'" & _
                " and format(a.endpaydate, ""yyyy"") = " & mYear & _
                " ORDER BY a.employeename"
        End If
    Else        ' ALL Clients
        If CutOffPeriod = "START" Then
            whereCon = "WHERE format(a.startpaydate, ""mmmm"") = '" & mMonth & "'" & _
                    " and format(a.startpaydate, ""yyyy"") = " & mYear & _
                    " ORDER BY a.employeename"
        Else    ' Cut-off period is END
            whereCon = "WHERE format(a.endpaydate, ""mmmm"") = '" & mMonth & "'" & _
                    " and format(a.endpaydate, ""yyyy"") = " & mYear & _
                    " ORDER BY a.employeename"
        End If
    End If
End If

Set rsPayTrans = New ADODB.Recordset

StrQuery = "SELECT a.EmpNo as controlnum, a.*, b.acctno, c.*, d.ClientName " & _
    "FROM ((payrolltrans AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID) " & _
    "LEFT JOIN qryEmployeeBasics AS c ON a.EmployeeID = c.EmployeeID) LEFT JOIN ClientTable AS d ON c.clientcode = d.clientcode " & whereCon

Set rsPayTrans = ReturnRst(StrQuery, cnnstr, mstatic)

'Check if there were records.
If rsPayTrans.EOF = True Then
    Set rsPayTrans = Nothing
    'Set rsPayTrans = New ADODB.Recordset
    
    StrQuery = "SELECT a.EmpNo as controlnum, a.*, b.acctno, c.*, d.ClientName " & _
        "FROM ((payrollmaster AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID) " & _
        "LEFT JOIN qryEmployeeBasics AS c ON a.EmployeeID = c.EmployeeID) LEFT JOIN ClientTable AS d ON c.clientcode = d.clientcode " & whereCon

    Set rsPayTrans = ReturnRst(StrQuery, cnnstr, mstatic)
End If
Set GetPayrollTransactionRecords = rsPayTrans
End Function

Private Sub GetDTRreport()
Dim StartDate As String
Dim EndDate As String

Set rsGetDTRreport = New ADODB.Recordset
StartDate = frmSelectCutoff.txtCutoffDate(0).Text
EndDate = frmSelectCutoff.txtCutoffDate(1).Text
'strQuery = "SELECT a.EmpNo as controlnum, a.*, b.acctno, c.* " & _
    "FROM (payrolltrans AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID) LEFT JOIN qryEmployeeBasics AS c ON a.EmployeeID = c.EmployeeID " & _
    "WHERE a.clientcode = " & mclientcode & _
    " and a.startpaydate = '" & StartDate & "' and a.endpaydate = '" & EndDate & "'" & _
    " ORDER BY a.employeename"
    
StrQuery = "SELECT a.EmpNo AS controlnum, a.EmployeeName, b.AcctNo, a.ClientCode, ClientTable.ClientName, HRbranch.BrName, a.StartPayDate, a.EndPayDate, a.TASAesla, a.CashAdvance, a.MiscFee" & _
      " FROM (((payrolltrans AS a LEFT JOIN hremployees AS b ON a.EmployeeID = b.EmployeeID) LEFT JOIN qryEmployeeBasics AS c ON a.EmployeeID = c.EmployeeID) LEFT JOIN ClientTable ON a.ClientCode = ClientTable.ClientCode) LEFT JOIN HRbranch ON b.BrCode = HRbranch.BrCode " & _
      " WHERE (((a.StartPayDate)='" & StartDate & "') AND ((a.EndPayDate)='" & EndDate & "')) " & _
      " ORDER BY a.EmployeeName"
      
Set rsGetDTRreport = ReturnRst(StrQuery, cnnstr, mstatic)

End Sub

Private Sub showBank()
If cboBank.ListCount > 0 Then Exit Sub
    cboBank.Clear
    Dim rstBank As ADODB.Recordset
    Set rstBank = cnn.Execute("select distinct BankName,BankID from TBLBANK")
    Do While Not rstBank.EOF
        cboBank.AddItem rstBank!bankname
        cboBank.ItemData(cboBank.NewIndex) = rstBank!BANKID
        rstBank.MoveNext
    Loop
    rstBank.Close
End Sub

Private Sub ShowBankCopy()
    Dim StartDate
    Dim EndDate
    Dim mFortheMonth As Integer
    Dim mTotalEmp As Integer
    Dim mAgencyCode As Integer
    mAgencyCode = 1
    StartDate = Format(frmSelectCutoff.txtCutoffDate(0).Text, "yyyymmdd")
    EndDate = Format(frmSelectCutoff.txtCutoffDate(1).Text, "yyyymmdd")
    mFortheMonth = Format(frmSelectCutoff.txtCutoffDate(0).Text, "m")
    mTotalEmp = rsPayrollTrans.RecordCount
    
    With rptDebitStatement
    If chkDebit.VALUE = 1 Then
          If CompanyCode = 1 Then
         .Sections("Section2").Controls("image1").Visible = True
         .Sections("Section2").Controls("image2").Visible = False
         ElseIf CompanyCode = 2 Then
         .Sections("Section2").Controls("image2").Visible = True
         .Sections("Section2").Controls("image1").Visible = False
         End If
      End If
      End With
    
    Dim rsItems As New ADODB.Recordset
    Set rsItems = ReturnRst("Select * from clienttable where clientcode=" & mClientCode, cnnstr, mstatic)
    Dim Cls As clsCommon
    Set Cls = New clsCommon
    
    
    Set rsItems = ReturnRst("Select * from clienttable where clientcode=" & mClientCode, cnnstr, mstatic)
    If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
    With rptDebitStatement
         Set .DataSource = rsItems
             .DataMember = rsItems.DataMember
                  .Sections("Section1").Controls("lblGenManager").Caption = GetAssistant(5, mAgencyCode)
              '           strQuery = "select sum(netpay) + sum(CalamityLoan) as Gross from payrolltrans WHERE " & _
                       "  ClientCode = " & mclientcode & "" & _
                      " and StartPayDate = '" & frmSelectCutoff.txtCutOffDate(0) & "' and endPayDate= '" & frmSelectCutoff.txtCutOffDate(1) & "'"
'              Dim BillingAmount As New ADODB.Recordset
'              Set BillingAmount = ReturnRst(strQuery, Cnnstr, mstatic)
'              If BillingAmount.EOF = False Then
'              Dim BillAmount As Double
'              BillAmount = IIf(IsNull(BillingAmount.Fields!gross) = True, "0", BillingAmount.Fields!gross)
'              BillAmount = Format(TxtAmount.Text, "##,##0.00")
'              End If
'            '  BillAmount = 12344.34
 
         .Sections("Section2").Controls("LblDmNum").Caption = "DM No. " & TxtDMNum.Text '"Php  " & Format(BillAmount, "#0,000.00")
         .Sections("Section1").Controls("Lblbankname").Caption = GetBankInfo(cboBank.ItemData(cboBank.ListIndex), 1)
          .Sections("Section1").Controls("Lbladdress").Caption = GetBankInfo(cboBank.ItemData(cboBank.ListIndex), 2)
           .Sections("Section1").Controls("LblContactPerson").Caption = GetBankInfo(cboBank.ItemData(cboBank.ListIndex), 3)
          .Sections("Section1").Controls("LblPosition").Caption = GetBankInfo(cboBank.ItemData(cboBank.ListIndex), 4)
         
              Dim billing As String
              billing = "This is to authorize "
              billing = billing & GetBankInfo(cboBank.ItemData(cboBank.ListIndex), 1)
              billing = billing & " to debit " & mCompanyName & ", INC.'s CA/SA No. "
'              If CompanyCode = 1 Then
'                   billing = billing & " to debit SEARCH ASIA, INC.'s CA/SA No. "
'              ElseIf CompanyCode = 2 Then
'                  billing = billing & " to debit JOBS & CAREERS, INC.'s CA/SA No. "
'              End If
              billing = billing & TxtSANum.Text & _
                         " with the bank's SHAW BOULEVARD branch for the amount of PESOS : "
                         billing = billing & UCase(Cls.Word(TxtAmount))
                         billing = billing & " corresponding to the payroll period "
                         billing = billing & Format(txtCutoffDate(0).Text, "mmm dd") & " - " & Format(txtCutoffDate(1).Text, "mmm dd, yyyy")
                         billing = billing & " for actual crediting on "
                         billing = billing & Format(txtCreditDate.Text, "mmmm dd, yyyy")
                         billing = billing & " with details as per attached Payroll proof list, and as per Payroll File in diskette or transmitted electronically."
              
'              billing = "      This  is to bill you the amount in PESOS: "
'              billing = billing & UCase(Cls.Word(BillAmount)) ' UCase(Cls.Word(11427.68)) 'amount
'              billing = billing & " representing the services rendered for the payroll period "
'              billing = billing & Format(frmSelectCutoff.txtCutOffDate(0).Text, "mmm dd") & " - " & Format(frmSelectCutoff.txtCutOffDate(1).Text, "mmm dd, yyyy")
'              billing = billing & "."
      .Sections("Section1").Controls("lblBillStatement").Caption = billing
             ' Unload frmPrintPayslip
             
             .Show vbModal
    End With
End Sub

Private Function GetBankInfo(ByVal indx As Integer, ByVal fld As Integer) As String
     Dim rstBank As ADODB.Recordset
             Set rstBank = cnn.Execute("select * from TBLBANK where bankid=" & indx & "")
             
             If rstBank.EOF = False Then
                    If fld = 1 Then GetBankInfo = rstBank.Fields!bankname
                    If fld = 2 Then GetBankInfo = rstBank.Fields!Address
                     If fld = 3 Then GetBankInfo = rstBank.Fields!Contactperson
                      If fld = 4 Then GetBankInfo = rstBank.Fields!Position
             End If
             
             rstBank.Close
End Function

Private Sub ShowStatement()
    Dim StartDate
    Dim EndDate
    Dim mFortheMonth As Integer
    Dim mTotalEmp As Integer
    Dim mAgencyCode As Integer
    
    On Error GoTo Error1
    
    mAgencyCode = 1
    StartDate = Format(txtCutoffDate(1).Text, "yyyymmdd")
    EndDate = Format(txtCutOffDate1(1).Text, "yyyymmdd")
    mFortheMonth = Format(txtCutOffDate1(0).Text, "m")
    mTotalEmp = rsPayrollTrans.RecordCount
    
     With rpStatement
     If chkStatement.VALUE = 1 Then
          If CompanyCode = 1 Then
         .Sections("Section2").Controls("image1").Visible = True
         .Sections("Section2").Controls("image2").Visible = False
         ElseIf CompanyCode = 2 Then
         .Sections("Section2").Controls("image2").Visible = True
         .Sections("Section2").Controls("image1").Visible = False
         End If
      End If
      End With
    
    Dim rsItems As New ADODB.Recordset
    Set rsItems = ReturnRst("Select * from clienttable where clientcode=" & mClientCode, cnnstr, mstatic)
    Dim Cls As clsCommon
    Set Cls = New clsCommon
    
    Set rsItems = ReturnRst("Select * from clienttable where clientcode=" & mClientCode, cnnstr, mstatic)
    If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
    With rpStatement
         Set .DataSource = rsItems
             .DataMember = rsItems.DataMember
             .Sections("Section1").Controls("lblCompany").Caption = GetAssistant(4, mAgencyCode)
             .Sections("Section1").Controls("txtcompanyheader").DataField = "ClientName"
             
             .Sections("Section1").Controls("txtClientAddress").DataField = "ClientAddress"
             .Sections("Section1").Controls("lblAgencyAssistant").Caption = GetAssistant(0, mAgencyCode)
             .Sections("Section1").Controls("lblAgencyAssistantPosition").Caption = GetAssistant(1, mAgencyCode)
             .Sections("Section1").Controls("lblGenManager").Caption = GetAssistant(2, mAgencyCode)
             .Sections("Section1").Controls("lblGenManCaption").Caption = GetAssistant(3, mAgencyCode)
             .Sections("Section1").Controls("lblAccounting").Caption = GetAccounting(0, mClientCode)
             .Sections("Section1").Controls("lblAccountingCaption").Caption = GetAccounting(1, mClientCode)

Dim BillAmount
                BillAmount = Format(TxtAmount1, "##,##0.00")
              .Sections("Section1").Controls("LblBillingAmount").Caption = "Php  " & Format(TxtAmount1, "##,##0.00")
              Dim billing As String
              billing = "      This  is to bill you the amount in PESOS: "
              billing = billing & UCase(Cls.Word(BillAmount)) ' UCase(Cls.Word(11427.68)) 'amount
              billing = billing & " representing the services rendered for the payroll period "
              billing = billing & Format(frmSelectCutoff.txtCutoffDate(0).Text, "mmm dd") & " - " & Format(frmSelectCutoff.txtCutoffDate(1).Text, "mmm dd, yyyy")
              billing = billing & "."
              .Sections("Section1").Controls("lblBillStatement").Caption = billing
              ' Print SOA and Due Date.
              .Sections("Section1").Controls("lblsoa").Caption = "SOA Control#       " & Trim(txtSOA.Text)
              .Sections("Section1").Controls("lblDueDate").Caption = "Due Date:        " & IIf(IsNull(mskDueDate), "", Format(mskDueDate, "mmmm dd, yyyy"))
             ' Unload frmPrintPayslip
             
             .Show vbModal
    End With
    
Exit Sub

Error1:
    
        MsgBox Err.Description, vbInformation, PayrollAlert
        Err.Clear
    
End Sub

Private Function BillAmount() As Double
    Dim BillingAmount As New ADODB.Recordset
    
    ' Do not include the Agency Charge.
    StrQuery = "select sum(netpay) as Gross from payrolltrans WHERE " & _
                "  ClientCode = " & mClientCode & "" & _
                " and StartPayDate = '" & frmSelectCutoff.txtCutoffDate(0) & "' and endPayDate= '" & frmSelectCutoff.txtCutoffDate(1) & "'"
'    strQuery = "select sum(netpay) + sum(CalamityLoan) as Gross from payrolltrans WHERE " & _
'                "  ClientCode = " & mclientcode & "" & _
'                " and StartPayDate = '" & frmSelectCutoff.txtCutOffDate(0) & "' and endPayDate= '" & frmSelectCutoff.txtCutOffDate(1) & "'"

    Set BillingAmount = ReturnRst(StrQuery, cnnstr, mstatic)
    If BillingAmount.EOF = False And BillingAmount.Fields!gross <> 0 Then
        BillAmount = IIf(IsNull(BillingAmount.Fields!gross) = True, "0", BillingAmount.Fields!gross)
        BillAmount = Format(BillAmount, "##0.00")
    Else    ' if 0, get it from payroll master table
        BillingAmount.Close
        StrQuery = "select sum(netpay) as Gross from payrollmaster WHERE " & _
                "  ClientCode = " & mClientCode & "" & _
                " and StartPayDate = '" & frmSelectCutoff.txtCutoffDate(0) & "' and endPayDate= '" & frmSelectCutoff.txtCutoffDate(1) & "'"
        Set BillingAmount = ReturnRst(StrQuery, cnnstr, mstatic)
        If BillingAmount.EOF = False And BillingAmount.Fields!gross <> 0 Then
            BillAmount = IIf(IsNull(BillingAmount.Fields!gross) = True, "0", BillingAmount.Fields!gross)
            BillAmount = Format(BillAmount, "##0.00")
        Else
            BillAmount = Format(0, "##0.00")
        End If
    End If
        
End Function

Public Function GetAssistant(ByVal WhatField As Integer, ByVal whereCon As Integer)
Dim rsAssistant As New ADODB.Recordset

StrQuery = "Select AgencyAssistant,AgencyAssitantPosition,AgencyContactPerson,AgencyContactPosition,agencyname,AgencyOwner from companytable "
Set rsAssistant = ReturnRst(StrQuery, cnnstr, mstatic)
'Set rsAssistant = ReturnRst("Select AgencyAssistant,AgencyAssitantPosition,AgencyContactPerson,AgencyContactPosition,agencyname,AgencyOwner from companytable " & whereCon, CnnstrHR, mstatic)
If rsAssistant.EOF = False Then
    If WhatField = 0 Then GetAssistant = rsAssistant.Fields!AgencyAssistant
    If WhatField = 1 Then GetAssistant = rsAssistant.Fields!AgencyAssitantPosition
    If WhatField = 2 Then GetAssistant = rsAssistant.Fields!AgencyContactPerson
    If WhatField = 3 Then GetAssistant = rsAssistant.Fields!AgencyContactPosition
    If WhatField = 4 Then GetAssistant = rsAssistant.Fields!AgencyName
     If WhatField = 5 Then GetAssistant = rsAssistant.Fields!AgencyOwner
End If
End Function

Public Function GetAccounting(ByVal WhatField As Integer, ByVal whereCon As Integer) As String
Dim rsAssistant As New ADODB.Recordset
Set rsAssistant = ReturnRst("Select ClientContactPerson,ClientContactPosition from clienttable where clientcode= " & whereCon, cnnstr, mstatic)
If rsAssistant.EOF = False Then
    If WhatField = 0 Then GetAccounting = IIf(IsNull(rsAssistant.Fields!ClientContactPerson) = True, "<Not declared>", rsAssistant.Fields!ClientContactPerson)
    If WhatField = 1 Then GetAccounting = IIf(IsNull(rsAssistant.Fields!ClientContactPosition) = True, "<Not declared>", rsAssistant.Fields!ClientContactPosition)
End If
End Function

Private Sub Form_Load()
Dim StrQuery As String
Dim rsCompany As ADODB.Recordset
Dim currSOANo As String
   
    Me.Height = 3330
    fraButton.Top = 2330 '2430
    Frame1.Height = 2130
    
    ' Center screen
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    ' Get Company Information
    StrQuery = "Select * from companytable where agencycode = 1"
    Set rsCompany = ReturnRst(StrQuery, CnnstrHR, mstatic)
    PrevSOANo = IIf(IsNull(rsCompany!soaControlNo), "yyyy-mm-000", rsCompany!soaControlNo)
    If CompanyCode = 1 Then
        mSapi = " (SAPI)"
    Else
        mSapi = " (JCI)"
    End If
    
    Set rsCompany = Nothing
    pageset = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsCurrPayrollTrans = Nothing
Unload Me
End Sub

Private Sub BillingSummaryReport(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim sumString As String
Dim ctr As Long
Dim i As Long, j As Long
Dim cPos As Integer

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

'Set xlSheet = book1.Worksheets.Add
Set xlSheet = xlBook1.Worksheets(1)

cPos = 5
ctr = 7
i = 1

'On Error GoTo Errh
'Print Header
'xlSheet.Range("A2") = mCompanyName & mSapi  '" (SAPI)"
xlSheet.Range("A2") = mCompanyName & " (" & CompanyShort & ")"

With xlSheet.Range("A2")
    .Font.Name = "Arial"  'Font.Name
    .Font.Size = 14 'font.size
    .Font.Bold = True ' Font.Bold
    .Font.Italic = False 'font.Italic
    '.HorizontalAlignment = xlHAlignLeft 'left
    '.HorizontalAlignment = xlHAlignCenter 'center
    '.HorizontalAlignment = xlHAlignRight 'right
End With
xlSheet.Range("A3") = "Pay Period ending " & Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "dd yyyy")
With xlSheet.Range("A3")
    .Font.Name = "Arial"  'Font.Name
    .Font.Size = 10 'Font.size
    .Font.Bold = False ' Font.Bold
    .Font.Italic = False 'font.Italic
    '.HorizontalAlignment = xlHAlignLeft 'left
    '.HorizontalAlignment = xlHAlignCenter 'center
    '.HorizontalAlignment = xlHAlignRight 'right
End With
''If pageset = False Then
''With Application.ActiveSheet.PageSetup
'With xlSheet.PageSetup
'    '.LeftHeader = FormatPage(Me.txtLeftHdr)
'    '.RightHeader = FormatPage(Me.txtRtHdr)
'   '.CenterHeader = "Run Date / Time: &date (&time)"
'    '.LeftFooter = FormatPage(Me.txtLeftFtr)
'    '.RightFooter = "Page: &page / &pages"
'    '.CenterFooter = FormatPage(Me.txtCenterFtr)
'    .PrintTitleRows = "$1:$4" ' & mfgHeading.Rows
'    .Orientation = xlLandscape 'IIf(Me.optPortrait, xlPortrait, xlLandscape)
'    .TopMargin = 0.75 ' Application.InchesToPoints(CDbl(cboTopMargin))
'    .LeftMargin = 0.75 'Application.InchesToPoints(CDbl(cboLeftMargin))
'    .RightMargin = 0.75 'Application.InchesToPoints(CDbl(cboRightMargin))
'    .BottomMargin = 0.75 'Application.InchesToPoints(CDbl(cboBottomMargin))
'    .HeaderMargin = 0.5 'Application.InchesToPoints(CDbl(cboHeaderMargin))
'    .FooterMargin = 0.5 'Application.InchesToPoints(CDbl(cboFooterMargin))
'    '.Order = xlOverThenDown
'End With

'pageset = True
'End If

'xlSheet.Cells(cPos, 1).Value = "ID"
xlSheet.Cells(cPos, 2).VALUE = "Employee Name"
xlSheet.Cells(cPos, 3).VALUE = "Tax"
xlSheet.Cells(cPos, 4).VALUE = "Daily"
xlSheet.Cells(cPos, 5).VALUE = "Reg."
xlSheet.Cells(cPos, 6).VALUE = "This Pay"
xlSheet.Cells(cPos, 7).VALUE = "Absent/"
xlSheet.Cells(cPos, 8).VALUE = "OT"
xlSheet.Cells(cPos, 9).VALUE = "OT"
xlSheet.Cells(cPos, 10).VALUE = "50.00"
xlSheet.Cells(cPos, 11).VALUE = "Total"
xlSheet.Cells(cPos, 12).VALUE = "13th"
xlSheet.Cells(cPos, 13).VALUE = "Employer"

xlSheet.Cells(cPos, 17).VALUE = "TOTAL"
xlSheet.Cells(cPos, 18).VALUE = "Agency"
xlSheet.Cells(cPos, 19).VALUE = "GROSS"
xlSheet.Cells(cPos, 20).VALUE = "plus"
xlSheet.Cells(cPos, 21).VALUE = "TOTAL"
xlSheet.Cells(cPos, 22).VALUE = "ChinaBank/BDO"

cPos = cPos + 1
xlSheet.Cells(cPos, 1).VALUE = "No."
xlSheet.Cells(cPos, 3).VALUE = "Status"
xlSheet.Cells(cPos, 4).VALUE = "Rate"
xlSheet.Cells(cPos, 5).VALUE = "Days"
xlSheet.Cells(cPos, 6).VALUE = "Period"
xlSheet.Cells(cPos, 7).VALUE = "Tardiness"
xlSheet.Cells(cPos, 8).VALUE = "Hours"
xlSheet.Cells(cPos, 9).VALUE = "Amount"
xlSheet.Cells(cPos, 10).VALUE = "ECOLA"
xlSheet.Cells(cPos, 11).VALUE = "Pay"
xlSheet.Cells(cPos, 12).VALUE = "Month"

xlSheet.Cells(cPos, 13).VALUE = "SSS"
xlSheet.Cells(cPos, 14).VALUE = "EC"
xlSheet.Cells(cPos, 15).VALUE = "PhilHealth"
xlSheet.Cells(cPos, 16).VALUE = "Pag-ibig"

xlSheet.Cells(cPos, 17).VALUE = "CHARGES"
xlSheet.Cells(cPos, 18).VALUE = "Charge"
xlSheet.Cells(cPos, 19).VALUE = "BILLING"
xlSheet.Cells(cPos, 20).VALUE = "VAT"
xlSheet.Cells(cPos, 21).VALUE = "BILLING"
xlSheet.Cells(cPos, 22).VALUE = "Acct. Number"
    
Do While Not rs.EOF
    'DoEvents
    xlSheet.Cells(ctr, 1).VALUE = i
    xlSheet.Cells(ctr, 2).VALUE = rs!EmployeeName
    xlSheet.Cells(ctr, 3).VALUE = rs!EmpTaxCode
    xlSheet.Cells(ctr, 4).VALUE = rs!DailyRate
    xlSheet.Cells(ctr, 5).VALUE = rs!DaysPresent
    xlSheet.Cells(ctr, 6).VALUE = Format(rs!BasicPay, "##,##0.00")
    xlSheet.Cells(ctr, 7).VALUE = Format(rs!AbsentDeduction + rs!TardinessDeduction, "##,#0.00")
    xlSheet.Cells(ctr, 8).VALUE = Format(rs!OTHours, "##,##0.00")
    xlSheet.Cells(ctr, 9).VALUE = Format(rs!OTPay, "##,##0.00")
    xlSheet.Cells(ctr, 10).VALUE = Format(rs!EcolaPay, "##,##0.00")
    ' Total pay.
    mTotalPay = rs!BasicPay + rs!OTHours + rs!OTPay + rs!EcolaPay - (rs!AbsentDeduction + rs!TardinessDeduction)
    xlSheet.Cells(ctr, 11).VALUE = Format(mTotalPay, "##,##0.00")
    xlSheet.Cells(ctr, 12).VALUE = Format(rs!WTaxAmount, "##,##0.00")
    
    If rs!SSSEmployer <> 0 Then
        ' Updated to get the actual ECC contribution, 8/6/2009.
        xlSheet.Cells(ctr, 13).VALUE = Format(rs!SSSEmployer - IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer), "##,##0.00")
        xlSheet.Cells(ctr, 14).VALUE = Format(IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer), "##,##0.00")
    Else
        xlSheet.Cells(ctr, 13).VALUE = Format(rs!SSSEmployer, "##,##0.00")
        xlSheet.Cells(ctr, 14).VALUE = Format(0, "##,##0.00")
    End If
    xlSheet.Cells(ctr, 15).VALUE = Format(rs!PHealthEmployer, "##,##0.00")
    xlSheet.Cells(ctr, 16).VALUE = Format(rs!PagibigEmployer, "##,##0.00")
    xlSheet.Cells(ctr, 17).FormulaR1C1 = "=RC[-6]+RC[-5]+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
    xlSheet.Cells(ctr, 18).VALUE = Format(rs!CalamityLoan, "##,##0.00")
    xlSheet.Cells(ctr, 19).FormulaR1C1 = "=RC[-2]+RC[-1]"
    xlSheet.Cells(ctr, 21).FormulaR1C1 = "=RC[-2]+RC[-1]"
    xlSheet.Cells(ctr, 22).VALUE = IIf(rs!AcctNo = "" Or IsNull(rs!AcctNo), "n/a", rs!AcctNo)
  
    'Total Deduction
    'mTotalDeduction = rs!WTaxAmount + rs!SSSEmployee + rs!PHealthEmployee + rs!PagibigEmployee + rs!TASAesla + rs!CardTransFee + rs!ValeWithClient + rs!MiscFee + rs!CashAdvance
    'xlSheet.Cells(ctr, 18).Value = Format(mTotalPay - mTotalDeduction, "###,##0.00")
 
    'xlSheet.Cells.EntireColumn.AutoFit
    'xlSheet.Cells.EntireRow.AutoFit
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext
Loop

'Compute for Grand Total.
sumString = "=SUM(R[" & i * -1 & "]C:R[-2]C)"
For j = Asc("E") To Asc("U")
    mRange = Chr(j) & i + 7
    xlSheet.Range(mRange).Select
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
Next

mRange = "A5:V" & i + 7
xlSheet.Range(mRange).Select
xlApp1.Selection.Columns.AutoFit
With xlApp1.Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .MergeCells = False
End With

'Format amount to 0.00.
mRange = "E7:U" & i + 7
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
With xlApp1.Selection
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .MergeCells = False
End With

mRange = "A5:V" & i + 7
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"  'Font.Name
    .Font.Size = 8 'font.size
    .Font.Bold = False ' Font.Bold
    .Font.Italic = False 'font.Italic
End With

mRange = "M5:P5"
xlSheet.Range(mRange).Select
xlApp1.Selection.Merge

xlSheet.Columns("E:E").ColumnWidth = 5.71

mRange = "A5:V" & i + 7
xlSheet.Range(mRange).Select
'Selection.Borders(xlDiagonalDown).LineStyle = xlContinuous
'Selection.Borders(xlDiagonalUp).LineStyle = xlContinuous
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

'START
With xlApp1
    .Range("A5:V6").Select
    .Range("V6").Activate
    .ActiveWindow.ScrollColumn = 14
    With xlApp1.Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
    End With
    .Range("A5:V6").Select
    .Range("V6").Activate
    .Selection.Font.Bold = True
    .ActiveWindow.LargeScroll ToRight:=-2
End With
'END

mRange = "B7:B" & i + 7
xlSheet.Range(mRange).Select
With xlApp1.Selection
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .MergeCells = False
End With
xlSheet.Range("A2").Select

errR:

    Set xlApp1 = Nothing
    Exit Sub
'    Range("C7").Select
'    Selection.Subtotal GroupBy:=3, Function:=xlCount, TotalList:=Array(7), _
'        Replace:=True, PageBreaks:=True, SummaryBelowData:=True
'    ActiveWindow.SelectedSheets.PrintPreview
    
Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
    GoTo errR
   
'xlSheet.Name = "Overtime"
End Sub

Private Sub ShowDtrReport(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim sumString As String
Dim ctr As Long
Dim i As Long, j As Long
Dim cPos As Integer

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)
cPos = 7
ctr = 9
i = 1

'On Error GoTo Errh
' Print Header
xlSheet.Range("A2") = "FOR THE PAYROL PERIOD OF:"
xlSheet.Range("A3") = "PAYROLL RELEASE DATE:"
xlSheet.Range("A4") = "COMPANY NAME:"

xlSheet.Range("D2") = Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd yyyy")
xlSheet.Range("D3") = Format(Date, "mmm dd yyyy")
xlSheet.Range("D4") = rs!ClientName

'With xlSheet.Range("A2")
'    .Font.Name = "Arial"  'Font.Name
'    .Font.Size = 12 'font.size
'    .Font.Bold = True ' Font.Bold
'    .Font.Italic = False 'font.Italic
'End With

xlSheet.Cells(cPos - 1, 1).VALUE = "No."
xlSheet.Cells(cPos - 1, 2).VALUE = "NAME"
xlSheet.Cells(cPos - 1, 3).VALUE = "BRANCH"
xlSheet.Cells(cPos - 1, 4).VALUE = "DTR"
xlSheet.Cells(cPos - 1, 5).VALUE = "OT"
xlSheet.Cells(cPos, 6).VALUE = "C/B"
xlSheet.Cells(cPos, 7).VALUE = "P/F"
xlSheet.Cells(cPos, 8).VALUE = "M/F"
xlSheet.Cells(cPos, 9).VALUE = "C/A"
xlSheet.Cells(cPos - 1, 10).VALUE = "ADD-ON"
xlSheet.Cells(cPos, 10).VALUE = "OTHERS"
xlSheet.Cells(cPos - 1, 11).VALUE = "ACCOUNT"
xlSheet.Cells(cPos, 11).VALUE = "NUMBER"
xlSheet.Cells(cPos, 12).VALUE = "REMARKS"
cPos = cPos + 1

Do While Not rs.EOF
    'DoEvents
    xlSheet.Cells(ctr, 1).VALUE = i
    xlSheet.Cells(ctr, 2).VALUE = rs!EmployeeName
    xlSheet.Cells(ctr, 3).VALUE = rs!BrName
    xlSheet.Cells(ctr, 6).VALUE = Format(rs!TASAesla, "##,##0.00")
    xlSheet.Cells(ctr, 7).VALUE = Format(rs!CashAdvance, "##,#0.00")
    xlSheet.Cells(ctr, 8).VALUE = Format(rs!MiscFee, "##,##0.00")
    xlSheet.Cells(ctr, 11).VALUE = IIf(rs!AcctNo = "", "n/a", rs!AcctNo)
 
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Cells.EntireRow.AutoFit
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext
Loop

    xlSheet.Columns("A:A").ColumnWidth = 3.43
    xlSheet.Columns("B:B").ColumnWidth = 24.57
    xlSheet.Range("F6:I6").Select
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlApp1.Selection.Merge
    xlApp1.ActiveCell.FormulaR1C1 = "DEDUCTIONS"
    Dim R
    R = "A6:L" & ctr
    xlSheet.Range("" & R & "").Select
    xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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
    xlSheet.Range("A6:L7").Select
    xlApp1.Selection.Font.Bold = True
    With xlApp1.Selection.Font
        .Name = "Tahoma"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    xlSheet.Columns("L:L").ColumnWidth = 17.86
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
    End With
    xlSheet.Columns("F:I").Select
    xlApp1.Selection.ColumnWidth = 6.14
    xlSheet.Columns("E:E").ColumnWidth = 4.29
    xlSheet.Columns("D:D").ColumnWidth = 4.29
    
    xlSheet.Columns("K:K").Select
    xlApp1.Selection.ColumnWidth = 14.43
    xlApp1.Selection.ColumnWidth = 14.29
    
    xlApp1.Range("A6:L7").Select
    With xlApp1.Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
    End With
    xlApp1.Range("A2:G4").Select
    xlApp1.Selection.Font.Bold = True
    'xlApp1.Range("F9:H12").Select
    R = "F9:H" & ctr - 1
    xlApp1.Range("" & R & "").Select
    xlApp1.Selection.NumberFormat = "0.0"
    xlApp1.Selection.NumberFormat = "0.00"

    Set xlApp1 = Nothing
    Exit Sub
    
Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
    Set xlApp1 = Nothing
    
End Sub

Private Sub mskDueDate_Validate(Cancel As Boolean)
If mskDueDate <> "__/__/____" And IsDate(mskDueDate) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskDueDate.SetFocus

End Sub

Private Sub PrintReportHeader(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal mRange As String, ByVal Head1 As String, Optional Head2 As String = "", Optional HorAlign As Variant = xlCenter, _
    Optional VerAlign As Variant = xlCenter, Optional WrapText As Boolean = True, _
    Optional ShrinkToFit As Boolean = False, Optional MergedCells As Boolean = False, _
    Optional ColorIndex As Variant = xlNone, Optional FontName As String = "Arial", _
    Optional FontSize As Integer = 8, Optional FontBold As Boolean = False, Optional ColumnWidth As Integer = 7.5)
    
    xlSheet.Range(mRange).Select
    If Head2 <> "" Then
        xlApp1.ActiveCell.FormulaR1C1 = Head1 & Chr(10) & Head2
    Else
        xlApp1.ActiveCell.FormulaR1C1 = Head1
    End If
    With xlApp1
        .Selection.Interior.ColorIndex = ColorIndex '15 or xlnone
        .Selection.HorizontalAlignment = HorAlign 'xlCenter
        .Selection.VerticalAlignment = VerAlign 'xlCenter
        .Selection.WrapText = WrapText 'True
        .Selection.Orientation = 0
        .Selection.AddIndent = False
        .Selection.ShrinkToFit = ShrinkToFit 'False
        .Selection.MergeCells = MergedCells 'False
    End With
    With xlSheet.Range(mRange)
        .Font.Name = FontName '"Arial"    'Font.Name
        .Font.Size = FontSize '14         'Font.size
        .Font.Bold = FontBold 'True       'Font.Bold
    End With
    With xlSheet.Columns.Range(mRange)
        .ColumnWidth = ColumnWidth '7.5
    End With
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

Private Sub PayrollSummaryReportNew(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim sumString As String
Dim ctr As Long
Dim i As Integer, j As Integer
Dim cPos As Integer
Dim EndSw As Boolean
Dim mHead As String

Dim SubBDO(50) As Double
Dim SubATM(50) As Double
Dim SubCash(50) As Double
Dim SubHold(50) As Double
Dim bdoctr, atmctr, cashctr, holdctr
Dim LineCtr As Integer
Dim bdopos As Integer
Dim atmpos As Integer
Dim cashpos As Integer
Dim holdpos As Integer

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

'Set xlSheet = book1.Worksheets.Add
Set xlSheet = xlBook1.Worksheets(1)

cPos = 5
ctr = 7
LineCtr = 0
For i = 1 To 50
    SubBDO(i) = 0
    SubATM(i) = 0
    SubCash(i) = 0
    SubHold(i) = 0
    bdoctr = 0
    atmctr = 0
    cashctr = 0
    holdctr = 0
Next
i = 1

'On Error GoTo Errh
'Print Header
'xlSheet.Range("A2") = mCompanyName & msapi  '" (SAPI)"

'Updated by Bernard May 13, 2008.
mHead = mCompanyName & " (" & CompanyShort & ")"
'xlSheet.Range("A2") = mCompanyName & " (" & CompanyShort & ")"
Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 14, True)
If Format(rs!StartPayDate, "mmm") = Format(rs!EndPayDate, "mmm") Then
    'xlSheet.Range("A3") = "Payroll Summary Report for Cut-Off " & Format(rs!startPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "dd, yyyy")
    mHead = "Payroll Summary Report for Cut-Off " & Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "dd, yyyy")
Else
    'xlSheet.Range("A3") = "Payroll Summary Report for Cut-Off " & Format(rs!startPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd, yyyy")
    mHead = "Payroll Summary Report for Cut-Off " & Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd, yyyy")
End If
Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, False)

'Updated by Bernard March 1, 2008.

'With xlSheet.Range("A3")
'    .Font.Name = "Arial"  'Font.Name
'    .Font.Size = 10 'font.size
'    .Font.Bold = False ' Font.Bold
'    .Font.Italic = False 'font.Italic
'    '.HorizontalAlignment = xlHAlignLeft 'left
'    '.HorizontalAlignment = xlHAlignCenter 'center
'    '.HorizontalAlignment = xlHAlignRight 'right
'End With

'xlSheet.Range("AO3") = mSeriesNo
Call PrintReportHeader(xlSheet, xlApp1, "AS3", mSeriesNo, "", xlLeft, , False, False, False, , , 8, False)
'xlSheet.Range("AO4") = Format(Now, "mm/dd/yyyy")
Call PrintReportHeader(xlSheet, xlApp1, "AS4", Format(Now, "mm/dd/yyyy"), "", xlLeft, xlBottom, False, False, False, , , 8, False)

'With xlSheet.Range("AK3")
'    .Font.Name = "Arial"  'Font.Name
'    .Font.Size = 8 'font.size
'    .Font.Bold = False ' Font.Bold
'    .Font.Italic = False 'font.Italic
'End With
'With xlSheet.Range("AK4")
'    .Font.Name = "Arial"  'Font.Name
'    .Font.Size = 8 'font.size
'    .Font.Bold = False ' Font.Bold
'    .Font.Italic = False 'font.Italic
'
'        .HorizontalAlignment = xlLeft
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .MergeCells = False
'
'End With

'xlSheet.Cells(cPos, 1).VALUE = "Control"
Call PrintReportHeader(xlSheet, xlApp1, "A5:A6", "Control", "No.", xlCenter, xlCenter, True, False, True, 15, , 8, False, 13)
Call PrintReportHeader(xlSheet, xlApp1, "B5:B6", "Employee Name", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 17)
Call PrintReportHeader(xlSheet, xlApp1, "C5:C6", "Tax", "Status", xlCenter, xlCenter, True, False, True, 15, , 8, False, 5)
Call PrintReportHeader(xlSheet, xlApp1, "D5:D6", "Daily", "Rate", xlCenter, xlCenter, True, False, True, 15, , 8, False, 5)
Call PrintReportHeader(xlSheet, xlApp1, "E5:E6", "Reg.", "Days", xlCenter, xlCenter, True, False, True, 15, , 8, False, 5)
Call PrintReportHeader(xlSheet, xlApp1, "F5:F6", "This Pay", "Period", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "G5:G6", "Absent", "Deduction", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)
Call PrintReportHeader(xlSheet, xlApp1, "H5:H6", "Tardiness", "Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)
Call PrintReportHeader(xlSheet, xlApp1, "I5:I6", "Tardiness", "Deduction", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)
Call PrintReportHeader(xlSheet, xlApp1, "J5:J6", "Reg OT", "Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 6)
Call PrintReportHeader(xlSheet, xlApp1, "K5:K6", "Reg OT", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 6)
Call PrintReportHeader(xlSheet, xlApp1, "L5:L6", "Night Diff", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)
Call PrintReportHeader(xlSheet, xlApp1, "M5:M6", "Night Diff", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.5)

Call PrintReportHeader(xlSheet, xlApp1, "N5:N6", "Rest Day", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "O5:O6", "Rest Day", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "P5:P6", "Sunday OT", "Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "Q5:Q6", "Sunday OT", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)

Call PrintReportHeader(xlSheet, xlApp1, "R5:R6", "Spl Hol.", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "S5:S6", "Spl Hol.", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "T5:T6", "Lgl Hol.", "OT Hours", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "U5:U6", "Lgl Hol.", "OT Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "V5:V6", "Salary", "Adjustment", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "W5:W6", "Incentive", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.3)
Call PrintReportHeader(xlSheet, xlApp1, "X5:X6", "E-cola/Allow.", "Payment", xlCenter, xlCenter, True, False, True, 15, , 8, False, 9.5)
Call PrintReportHeader(xlSheet, xlApp1, "Y5:Y6", "GROSS PAY", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 9)
Call PrintReportHeader(xlSheet, xlApp1, "Z5:Z6", "WTAX", "Amount", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AA5:AC5", "Employee", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 18)
Call PrintReportHeader(xlSheet, xlApp1, "AA6:AA6", "SSS", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 6)
Call PrintReportHeader(xlSheet, xlApp1, "AB6:AB6", "PhilHealth", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AC6:AC6", "Pag-ibig", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 6)
Call PrintReportHeader(xlSheet, xlApp1, "AD5:AD6", "Cash", "Bond", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AE5:AE6", "Placement", "Fee", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AF5:AF6", "Misc.", "Fee", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AG5:AG6", "Card", "Trans Fee", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AH5:AH6", "Salary", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AI5:AI6", "Vale", "(IOU)", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AJ5:AJ6", "Uniform", "Charges", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AK5:AK6", "Medical", "Charges", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AL5:AL6", "Other", "Ded.", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AM5:AM6", "Adjustment", "Audit", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AN5:AN6", "SSS", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AO5:AO6", "PBIG", "Loan", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AP5:AP6", "Vale", "w/ Client", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AQ5:AQ6", "TOTAL", "Deductions", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "AR5:AR6", "NET PAY", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "AS5:AS6", "ChinaBank/BDO", "Account Number", xlCenter, xlCenter, True, False, True, 15, , 8, False, 15.87)

Call PrintReportHeader(xlSheet, xlApp1, "AT5:AT6", "13th Month", "Payment", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.7)
Call PrintReportHeader(xlSheet, xlApp1, "AU5:AX5", "Employer", "", xlCenter, xlBottom, True, True, True, 15, , 8, False, 18)
Call PrintReportHeader(xlSheet, xlApp1, "AU6:AU6", "SSS", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AV6:AV6", "EC", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AW6:AW6", "PhilHealth", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AX6:AX6", "Pag-ibig", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7)
Call PrintReportHeader(xlSheet, xlApp1, "AY5:AY6", "TOTAL", "Charges", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "AZ5:AZ6", "Agency", "Charge", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8)
Call PrintReportHeader(xlSheet, xlApp1, "BA5:BA6", "GROSS", "BILLING", xlCenter, xlCenter, True, False, True, 15, , 8, False, 8.7)

' 2 Employee Name.
'xlSheet.Cells(cPos, 3).VALUE = "Tax"
'xlSheet.Cells(cPos, 4).VALUE = "Daily"
'xlSheet.Cells(cPos, 5).VALUE = "Reg."
'xlSheet.Cells(cPos, 6).VALUE = "This Pay"
'xlSheet.Cells(cPos, 7).VALUE = "Absent"
'
'xlSheet.Cells(cPos, 8).VALUE = "Tardiness"  ' Hrs
'xlSheet.Cells(cPos, 9).VALUE = "Tardiness"  ' deduction
'
'xlSheet.Cells(cPos, 10).VALUE = "Reg OT"
'xlSheet.Cells(cPos, 11).VALUE = "Reg OT"
'xlSheet.Cells(cPos, 12).VALUE = "Night Diff"
'xlSheet.Cells(cPos, 13).VALUE = "Night Diff"
'
''xlSheet.Cells(cPos, 14).VALUE = "Rest Day"
''xlSheet.Cells(cPos, 15).VALUE = "Rest Day"
'
''xlSheet.Cells(cPos, 15).VALUE = "Sunday OT"
''xlSheet.Cells(cPos, 16).VALUE = "Sunday OT"
'
'xlSheet.Cells(cPos, 14).VALUE = "Spl Hol."
'xlSheet.Cells(cPos, 15).VALUE = "Spl Hol."
'xlSheet.Cells(cPos, 16).VALUE = "Lgl Hol."
'xlSheet.Cells(cPos, 17).VALUE = "Lgl Hol."
'
''xlSheet.Cells(cPos, 21).VALUE = "Number of"
''xlSheet.Cells(cPos, 22).VALUE = "Sick Leave"
''xlSheet.Cells(cPos, 23).VALUE = "Number of"
''xlSheet.Cells(cPos, 24).VALUE = "Vac. Leave"
'
'xlSheet.Cells(cPos, 18).VALUE = "Salary"
'xlSheet.Cells(cPos, 19).VALUE = "Incentive"
'xlSheet.Cells(cPos, 20).VALUE = "E-cola/Allow."
'
''xlSheet.Cells(cPos, 17).Value = "GrossPay"
'xlSheet.Cells(cPos, 18).VALUE = "WTAX"
'xlSheet.Cells(cPos, 19).VALUE = "Employee"  ' 30
'xlSheet.Cells(cPos, 22).VALUE = "Cash"
'xlSheet.Cells(cPos, 23).VALUE = "Placement"
'xlSheet.Cells(cPos, 24).VALUE = "Misc."
'xlSheet.Cells(cPos, 25).VALUE = "Card"
'
'xlSheet.Cells(cPos, 26).VALUE = "Salary"    'Loan
'xlSheet.Cells(cPos, 27).VALUE = "VALE"
'xlSheet.Cells(cPos, 28).VALUE = "Uniform"
'xlSheet.Cells(cPos, 29).VALUE = "Medical"
'xlSheet.Cells(cPos, 30).VALUE = "Variances"
'xlSheet.Cells(cPos, 31).VALUE = "Adjustment"
'
'xlSheet.Cells(cPos, 32).VALUE = "SSS"  ' Retain position
'xlSheet.Cells(cPos, 33).VALUE = "PBIG"
'xlSheet.Cells(cPos, 34).VALUE = "Vale"
''xlSheet.Cells(cPos, 40).VALUE = "Over/Under"    ' over/under pay
'
'xlSheet.Cells(cPos, 35).VALUE = "TOTAL"
''xlSheet.Cells(cPos, 36).Value = "NET PAY"
'xlSheet.Cells(cPos, 37).VALUE = "ChinaBank/BDO"
'
''xlSheet.Cells(cPos, 38).VALUE = "13th Month"
''xlSheet.Cells(cPos, 39).VALUE = "Employer"
''xlSheet.Cells(cPos, 43).VALUE = "TOTAL"
''xlSheet.Cells(cPos, 44).VALUE = "Agency"
''xlSheet.Cells(cPos, 45).VALUE = "GROSS"
'
''xlSheet.Cells(cPos, 52).Value = "plus"
''xlSheet.Cells(cPos, 53).Value = "NET"
'
''xlSheet.Cells(cPos, 28).Value = "13th Month"
''xlSheet.Cells(cPos, 29).Value = "Cash Bond"
''xlSheet.Cells(cPos, 30).Value = "Vale"
''xlSheet.Cells(cPos, 31).Value = "Misc. Fee"
''xlSheet.Cells(cPos, 32).Value = "Card"
''xlSheet.Cells(cPos, 33).Value = "Placement"
''xlSheet.Cells(cPos, 34).Value = "Adjustment"
''xlSheet.Cells(cPos, 35).Value = "Employee"
''xlSheet.Cells(cPos, 38).Value = "Employer"
''xlSheet.Cells(cPos, 41).Value = "WTAX"
''xlSheet.Cells(cPos, 42).Value = "SSS Loan"
''xlSheet.Cells(cPos, 43).Value = "PBIG Loan"
''xlSheet.Cells(cPos, 45).Value = "TOTAL"
''xlSheet.Cells(cPos, 47).Value = "Agency"
''xlSheet.Cells(cPos, 48).Value = "ChinaBank/BDO"
'
'cPos = cPos + 1
'xlSheet.Cells(cPos, 1).VALUE = "No."
'xlSheet.Cells(cPos, 2).VALUE = "Employee Name"
'xlSheet.Cells(cPos, 3).VALUE = "Status"
'xlSheet.Cells(cPos, 4).VALUE = "Rate"
'xlSheet.Cells(cPos, 5).VALUE = "Days"
'xlSheet.Cells(cPos, 6).VALUE = "Period"
'xlSheet.Cells(cPos, 7).VALUE = "Deduction"
'
'xlSheet.Cells(cPos, 8).VALUE = "Hours"      ' Tardiness
'xlSheet.Cells(cPos, 9).VALUE = "Deduction"  ' Tardiness
'
'xlSheet.Cells(cPos, 10).VALUE = "Hours"     ' Reg OT
'xlSheet.Cells(cPos, 11).VALUE = "Amount"
'xlSheet.Cells(cPos, 12).VALUE = "OT Hours"
'xlSheet.Cells(cPos, 13).VALUE = "OT Amount"
'
''xlSheet.Cells(cPos, 14).VALUE = "OT Hours"
''xlSheet.Cells(cPos, 15).VALUE = "OT Amount"
'
''xlSheet.Cells(cPos, 16).VALUE = "Hours"     ' Sunday OT to be removed
''xlSheet.Cells(cPos, 17).VALUE = "Amount"    ' Sunday OT to be removed
'
''xlSheet.Cells(cPos, 16).VALUE = "OT Hours"
''xlSheet.Cells(cPos, 17).VALUE = "OT Amount"
''xlSheet.Cells(cPos, 18).VALUE = "OT Hours"
''xlSheet.Cells(cPos, 19).VALUE = "OT Amount"
'
''xlSheet.Cells(cPos, 22).VALUE = "Sick Leave"
''xlSheet.Cells(cPos, 23).VALUE = "Amount"
''xlSheet.Cells(cPos, 24).VALUE = "Vac. Leave"
''xlSheet.Cells(cPos, 25).VALUE = "Amount"
'
'xlSheet.Cells(cPos, 14).VALUE = "Adjustment"
'xlSheet.Cells(cPos, 15).VALUE = "Amount"
'xlSheet.Cells(cPos, 16).VALUE = "Payment"
'
'xlSheet.Cells(cPos, 17).VALUE = "GROSS PAY"
'xlSheet.Cells(cPos, 18).VALUE = "Amount"
'
'xlSheet.Cells(cPos, 19).VALUE = "SSS"
'xlSheet.Cells(cPos, 20).VALUE = "PhilHealth"
'xlSheet.Cells(cPos, 21).VALUE = "Pag-ibig"
'
'xlSheet.Cells(cPos, 22).VALUE = "Bond"
'xlSheet.Cells(cPos, 23).VALUE = "Fee"
'xlSheet.Cells(cPos, 24).VALUE = "Fee"
'xlSheet.Cells(cPos, 25).VALUE = "Trans Fee"
'
'xlSheet.Cells(cPos, 26).VALUE = "Loan"
'xlSheet.Cells(cPos, 27).VALUE = "(IOU)"
'xlSheet.Cells(cPos, 28).VALUE = "Charges"       ' Uniform
'xlSheet.Cells(cPos, 29).VALUE = "Charges"       ' Medical
'xlSheet.Cells(cPos, 30).VALUE = "Amount"
'xlSheet.Cells(cPos, 31).VALUE = "Audit"
'
'xlSheet.Cells(cPos, 32).VALUE = "Loan"          ' SSS
'xlSheet.Cells(cPos, 33).VALUE = "Loan"          ' PBIG
'xlSheet.Cells(cPos, 34).VALUE = "w/ Client"     ' VALE
'
'xlSheet.Cells(cPos, 35).VALUE = "Deductions"    ' Retain Position
'xlSheet.Cells(cPos, 36).VALUE = "NET PAY"
'xlSheet.Cells(cPos, 37).VALUE = "Account Number"
'xlSheet.Cells(cPos, 38).VALUE = "Payment"
'xlSheet.Cells(cPos, 39).VALUE = "SSS"
'xlSheet.Cells(cPos, 40).VALUE = "EC"
'xlSheet.Cells(cPos, 41).VALUE = "PhilHealth"
'xlSheet.Cells(cPos, 42).VALUE = "Pag-ibig"
'xlSheet.Cells(cPos, 43).VALUE = "Charges"
'xlSheet.Cells(cPos, 44).VALUE = "Charge"
'xlSheet.Cells(cPos, 45).VALUE = "BILLING"

Do While Not rs.EOF
    ' DoEvents
    xlSheet.Cells(ctr, 1).VALUE = rs!ControlNum
    xlSheet.Cells(ctr, 2).VALUE = rs!EmployeeName
    xlSheet.Cells(ctr, 3).VALUE = rs!EmpTaxCode
    xlSheet.Cells(ctr, 4).VALUE = rs!DailyRate
    xlSheet.Cells(ctr, 5).VALUE = rs!DaysPresent
    
    If rs!PaymentSystem = "BDO" Then
        bdoctr = bdoctr + 1
        SubBDO(1) = SubBDO(1) + rs!DaysPresent
        SubBDO(2) = SubBDO(2) + rs!BasicPay
        SubBDO(3) = SubBDO(3) + rs!AbsentDeduction
        SubBDO(4) = SubBDO(4) + rs!HoursTardiness   ' Add
        SubBDO(5) = SubBDO(5) + rs!TardinessDeduction
        SubBDO(6) = SubBDO(6) + rs!OTHours
        SubBDO(7) = SubBDO(7) + rs!OTPay
        SubBDO(8) = SubBDO(8) + rs!NightDiffOT
        SubBDO(9) = SubBDO(9) + rs!NightDiffPay
        SubBDO(10) = SubBDO(10) + rs!TaxRefundHours
        SubBDO(11) = SubBDO(11) + rs!TaxRefundPay
        SubBDO(12) = SubBDO(12) + rs!SundayOT
        SubBDO(13) = SubBDO(13) + rs!SundayOTPay
        
        SubBDO(14) = SubBDO(14) + rs!NoRegHoliday
        SubBDO(15) = SubBDO(15) + rs!RegHolidayPay
        SubBDO(16) = SubBDO(16) + rs!NoSplHoliday
        SubBDO(17) = SubBDO(17) + rs!SplHolidayPay
'        SubBDO(16) = SubBDO(16) + rs!NoSickLeave
'        SubBDO(18) = SubBDO(18) + rs!SickLeavePay
'        SubBDO(19) = SubBDO(19) + rs!NoVacationLeave
'        SubBDO(20) = SubBDO(20) + rs!VacationLeavePay
        SubBDO(18) = SubBDO(18) + rs!OtherIncome
        SubBDO(19) = SubBDO(19) + rs!Allowance
        SubBDO(20) = SubBDO(20) + rs!EcolaPay
        SubBDO(21) = SubBDO(21) + rs!GrossPay
        SubBDO(22) = SubBDO(22) + rs!WTaxAmount
        
        SubBDO(23) = SubBDO(23) + rs!SSSEmployee
        SubBDO(24) = SubBDO(24) + rs!PHealthEmployee
        SubBDO(25) = SubBDO(25) + rs!pagibigemployee
        SubBDO(26) = SubBDO(26) + rs!TASAesla
        SubBDO(27) = SubBDO(27) + rs!CashAdvance
        SubBDO(28) = SubBDO(28) + rs!MiscFee
        SubBDO(29) = SubBDO(29) + rs!LMCDed
        
        SubBDO(30) = SubBDO(30) + rs!SalaryLoan
        SubBDO(31) = SubBDO(31) + rs!Vale_IOU
        SubBDO(32) = SubBDO(32) + rs!Uniform
        SubBDO(33) = SubBDO(33) + rs!Medical
        SubBDO(34) = SubBDO(34) + rs!Variances
        SubBDO(35) = SubBDO(35) + rs!OverUnderPay
        
        SubBDO(36) = SubBDO(36) + rs!SSSLoan
        SubBDO(37) = SubBDO(37) + rs!PagibigLoan
        SubBDO(38) = SubBDO(38) + rs!ValeWithClient
        SubBDO(39) = SubBDO(39) + rs!Deductions
        SubBDO(40) = SubBDO(40) + rs!NetPay
        'SubBDO(33) = SubBDO(33) + rs!AcctNo    Account No
        SubBDO(42) = SubBDO(42) + rs!Month13Payment
        If rs!SSSEmployer <> 0 Then
            SubBDO(43) = SubBDO(43) + rs!SSSEmployer - IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer)  '10
            SubBDO(44) = SubBDO(44) + IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer)  '10
        Else
            SubBDO(43) = SubBDO(43) + 0
            SubBDO(44) = SubBDO(44) + 0
        End If
        SubBDO(45) = SubBDO(45) + rs!PHealthEmployer
        SubBDO(46) = SubBDO(46) + rs!PagibigEmployer
        SubBDO(47) = SubBDO(47) + rs!CalamityLoan
    ElseIf rs!PaymentSystem = "ATM" Then
        atmctr = atmctr + 1
        SubATM(1) = SubATM(1) + rs!DaysPresent
        SubATM(2) = SubATM(2) + rs!BasicPay
        SubATM(3) = SubATM(3) + rs!AbsentDeduction
        SubATM(4) = SubATM(4) + rs!HoursTardiness   ' Add
        SubATM(5) = SubATM(5) + rs!TardinessDeduction
        SubATM(6) = SubATM(6) + rs!OTHours
        SubATM(7) = SubATM(7) + rs!OTPay
        SubATM(8) = SubATM(8) + rs!NightDiffOT
        SubATM(9) = SubATM(9) + rs!NightDiffPay
        SubATM(10) = SubATM(10) + rs!TaxRefundHours
        SubATM(11) = SubATM(11) + rs!TaxRefundPay
        SubATM(12) = SubATM(12) + rs!SundayOT
        SubATM(13) = SubATM(13) + rs!SundayOTPay
        
        SubATM(14) = SubATM(14) + rs!NoRegHoliday
        SubATM(15) = SubATM(15) + rs!RegHolidayPay
        SubATM(16) = SubATM(16) + rs!NoSplHoliday
        SubATM(17) = SubATM(17) + rs!SplHolidayPay
'        subatm(16) = subatm(16) + rs!NoSickLeave
'        subatm(18) = subatm(18) + rs!SickLeavePay
'        subatm(19) = subatm(19) + rs!NoVacationLeave
'        subatm(20) = subatm(20) + rs!VacationLeavePay
        SubATM(18) = SubATM(18) + rs!OtherIncome
        SubATM(19) = SubATM(19) + rs!Allowance
        SubATM(20) = SubATM(20) + rs!EcolaPay
        SubATM(21) = SubATM(21) + rs!GrossPay
        SubATM(22) = SubATM(22) + rs!WTaxAmount
        
        SubATM(23) = SubATM(23) + rs!SSSEmployee
        SubATM(24) = SubATM(24) + rs!PHealthEmployee
        SubATM(25) = SubATM(25) + rs!pagibigemployee
        SubATM(26) = SubATM(26) + rs!TASAesla
        SubATM(27) = SubATM(27) + rs!CashAdvance
        SubATM(28) = SubATM(28) + rs!MiscFee
        SubATM(29) = SubATM(29) + rs!LMCDed
        
        SubATM(30) = SubATM(30) + rs!SalaryLoan
        SubATM(31) = SubATM(31) + rs!Vale_IOU
        SubATM(32) = SubATM(32) + rs!Uniform
        SubATM(33) = SubATM(33) + rs!Medical
        SubATM(34) = SubATM(34) + rs!Variances
        SubATM(35) = SubATM(35) + rs!OverUnderPay
        
        SubATM(36) = SubATM(36) + rs!SSSLoan
        SubATM(37) = SubATM(37) + rs!PagibigLoan
        SubATM(38) = SubATM(38) + rs!ValeWithClient
        SubATM(39) = SubATM(39) + rs!Deductions
        SubATM(40) = SubATM(40) + rs!NetPay
        'subatm(33) = subatm(33) + rs!AcctNo    Account No
        SubATM(42) = SubATM(42) + rs!Month13Payment
        If rs!SSSEmployer <> 0 Then
            SubATM(43) = SubATM(43) + rs!SSSEmployer - IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer)  '10
            SubATM(44) = SubATM(44) + IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer)  '10
        Else
            SubATM(43) = SubATM(43) + 0
            SubATM(44) = SubATM(44) + 0
        End If
        SubATM(45) = SubATM(45) + rs!PHealthEmployer
        SubATM(46) = SubATM(46) + rs!PagibigEmployer
        SubATM(47) = SubATM(47) + rs!CalamityLoan
    ElseIf rs!PaymentSystem = "Cash" Then
        cashctr = cashctr + 1
        SubCash(1) = SubCash(1) + rs!DaysPresent
        SubCash(2) = SubCash(2) + rs!BasicPay
        SubCash(3) = SubCash(3) + rs!AbsentDeduction
        SubCash(4) = SubCash(4) + rs!HoursTardiness   ' Add
        SubCash(5) = SubCash(5) + rs!TardinessDeduction
        SubCash(6) = SubCash(6) + rs!OTHours
        SubCash(7) = SubCash(7) + rs!OTPay
        SubCash(8) = SubCash(8) + rs!NightDiffOT
        SubCash(9) = SubCash(9) + rs!NightDiffPay
        SubCash(10) = SubCash(10) + rs!TaxRefundHours
        SubCash(11) = SubCash(11) + rs!TaxRefundPay
        SubCash(12) = SubCash(12) + rs!SundayOT
        SubCash(13) = SubCash(13) + rs!SundayOTPay
        
        SubCash(14) = SubCash(14) + rs!NoRegHoliday
        SubCash(15) = SubCash(15) + rs!RegHolidayPay
        SubCash(16) = SubCash(16) + rs!NoSplHoliday
        SubCash(17) = SubCash(17) + rs!SplHolidayPay
'        subcash(16) = subcash(16) + rs!NoSickLeave
'        subcash(18) = subcash(18) + rs!SickLeavePay
'        subcash(19) = subcash(19) + rs!NoVacationLeave
'        subcash(20) = subcash(20) + rs!VacationLeavePay
        SubCash(18) = SubCash(18) + rs!OtherIncome
        SubCash(19) = SubCash(19) + rs!Allowance
        SubCash(20) = SubCash(20) + rs!EcolaPay
        SubCash(21) = SubCash(21) + rs!GrossPay
        SubCash(22) = SubCash(22) + rs!WTaxAmount
        
        SubCash(23) = SubCash(23) + rs!SSSEmployee
        SubCash(24) = SubCash(24) + rs!PHealthEmployee
        SubCash(25) = SubCash(25) + rs!pagibigemployee
        SubCash(26) = SubCash(26) + rs!TASAesla
        SubCash(27) = SubCash(27) + rs!CashAdvance
        SubCash(28) = SubCash(28) + rs!MiscFee
        SubCash(29) = SubCash(29) + rs!LMCDed
        
        SubCash(30) = SubCash(30) + rs!SalaryLoan
        SubCash(31) = SubCash(31) + rs!Vale_IOU
        SubCash(32) = SubCash(32) + rs!Uniform
        SubCash(33) = SubCash(33) + rs!Medical
        SubCash(34) = SubCash(34) + rs!Variances
        SubCash(35) = SubCash(35) + rs!OverUnderPay
        
        SubCash(36) = SubCash(36) + rs!SSSLoan
        SubCash(37) = SubCash(37) + rs!PagibigLoan
        SubCash(38) = SubCash(38) + rs!ValeWithClient
        SubCash(39) = SubCash(39) + rs!Deductions
        SubCash(40) = SubCash(40) + rs!NetPay
        'subcash(33) = subcash(33) + rs!AcctNo    Account No
        SubCash(42) = SubCash(42) + rs!Month13Payment
        If rs!SSSEmployer <> 0 Then
            SubCash(43) = SubCash(43) + rs!SSSEmployer - IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer)  '10
            SubCash(44) = SubCash(44) + IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer)  '10
        Else
            SubCash(43) = SubCash(43) + 0
            SubCash(44) = SubCash(44) + 0
        End If
        SubCash(45) = SubCash(45) + rs!PHealthEmployer
        SubCash(46) = SubCash(46) + rs!PagibigEmployer
        SubCash(47) = SubCash(47) + rs!CalamityLoan
    Else 'If rs!PaymentSystem = "HOLD" Then
        holdctr = holdctr + 1
        SubHold(1) = SubHold(1) + rs!DaysPresent
        SubHold(2) = SubHold(2) + rs!BasicPay
        SubHold(3) = SubHold(3) + rs!AbsentDeduction
        SubHold(4) = SubHold(4) + rs!HoursTardiness   ' Add
        SubHold(5) = SubHold(5) + rs!TardinessDeduction
        SubHold(6) = SubHold(6) + rs!OTHours
        SubHold(7) = SubHold(7) + rs!OTPay
        SubHold(8) = SubHold(8) + rs!NightDiffOT
        SubHold(9) = SubHold(9) + rs!NightDiffPay
        SubHold(10) = SubHold(10) + rs!TaxRefundHours
        SubHold(11) = SubHold(11) + rs!TaxRefundPay
        SubHold(12) = SubHold(12) + rs!SundayOT
        SubHold(13) = SubHold(13) + rs!SundayOTPay
        
        SubHold(14) = SubHold(14) + rs!NoRegHoliday
        SubHold(15) = SubHold(15) + rs!RegHolidayPay
        SubHold(16) = SubHold(16) + rs!NoSplHoliday
        SubHold(17) = SubHold(17) + rs!SplHolidayPay
'        subhold(16) = subhold(16) + rs!NoSickLeave
'        subhold(18) = subhold(18) + rs!SickLeavePay
'        subhold(19) = subhold(19) + rs!NoVacationLeave
'        subhold(20) = subhold(20) + rs!VacationLeavePay
        SubHold(18) = SubHold(18) + rs!OtherIncome
        SubHold(19) = SubHold(19) + rs!Allowance
        SubHold(20) = SubHold(20) + rs!EcolaPay
        SubHold(21) = SubHold(21) + rs!GrossPay
        SubHold(22) = SubHold(22) + rs!WTaxAmount
        
        SubHold(23) = SubHold(23) + rs!SSSEmployee
        SubHold(24) = SubHold(24) + rs!PHealthEmployee
        SubHold(25) = SubHold(25) + rs!pagibigemployee
        SubHold(26) = SubHold(26) + rs!TASAesla
        SubHold(27) = SubHold(27) + rs!CashAdvance
        SubHold(28) = SubHold(28) + rs!MiscFee
        SubHold(29) = SubHold(29) + rs!LMCDed
        
        SubHold(30) = SubHold(30) + rs!SalaryLoan
        SubHold(31) = SubHold(31) + rs!Vale_IOU
        SubHold(32) = SubHold(32) + rs!Uniform
        SubHold(33) = SubHold(33) + rs!Medical
        SubHold(34) = SubHold(34) + rs!Variances
        SubHold(35) = SubHold(35) + rs!OverUnderPay
        
        SubHold(36) = SubHold(36) + rs!SSSLoan
        SubHold(37) = SubHold(37) + rs!PagibigLoan
        SubHold(38) = SubHold(38) + rs!ValeWithClient
        SubHold(39) = SubHold(39) + rs!Deductions
        SubHold(40) = SubHold(40) + rs!NetPay
        'subhold(33) = subhold(33) + rs!AcctNo    Account No
        SubHold(42) = SubHold(42) + rs!Month13Payment
        If rs!SSSEmployer <> 0 Then
            SubHold(43) = SubHold(43) + rs!SSSEmployer - IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer)  '10
            SubHold(44) = SubHold(44) + IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer)  '10
        Else
            SubHold(43) = SubHold(43) + 0
            SubHold(44) = SubHold(44) + 0
        End If
        SubHold(45) = SubHold(45) + rs!PHealthEmployer
        SubHold(46) = SubHold(46) + rs!PagibigEmployer
        SubHold(47) = SubHold(47) + rs!CalamityLoan
    End If
    xlSheet.Cells(ctr, 6).VALUE = Format(rs!BasicPay, "##,##0.00")
    xlSheet.Cells(ctr, 7).VALUE = Format(rs!AbsentDeduction, "##,#0.00")
    xlSheet.Cells(ctr, 8).VALUE = Format(rs!HoursTardiness, "##,#0.00")
    xlSheet.Cells(ctr, 9).VALUE = Format(rs!TardinessDeduction, "##,#0.00")
    
    xlSheet.Cells(ctr, 10).VALUE = Format(rs!OTHours, "##,##0.00")
    xlSheet.Cells(ctr, 11).VALUE = Format(rs!OTPay, "##,##0.00")
    xlSheet.Cells(ctr, 12).VALUE = Format(rs!NightDiffOT, "##,##0.00")
    xlSheet.Cells(ctr, 13).VALUE = Format(rs!NightDiffPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 14).VALUE = Format(rs!TaxRefundHours, "##,##0.00")
    xlSheet.Cells(ctr, 15).VALUE = Format(rs!TaxRefundPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 16).VALUE = Format(rs!SundayOT, "##,##0.00")
    xlSheet.Cells(ctr, 17).VALUE = Format(rs!SundayOTPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 18).VALUE = Format(rs!NoRegHoliday, "##,##0.00")
    xlSheet.Cells(ctr, 19).VALUE = Format(rs!RegHolidayPay, "##,##0.00")
    xlSheet.Cells(ctr, 20).VALUE = Format(rs!NoSplHoliday, "##,##0.00")
    xlSheet.Cells(ctr, 21).VALUE = Format(rs!SplHolidayPay, "##,##0.00")
    
'    xlSheet.Cells(ctr, 21).VALUE = Format(rs!NoSickLeave, "##,##0.00")
'    xlSheet.Cells(ctr, 22).VALUE = Format(rs!SickLeavePay, "##,##0.00")
'    xlSheet.Cells(ctr, 23).VALUE = Format(rs!NoVacationLeave, "##,##0.00")
'    xlSheet.Cells(ctr, 24).VALUE = Format(rs!VacationLeavePay, "##,##0.00")

    xlSheet.Cells(ctr, 22).VALUE = Format(rs!OtherIncome, "##,##0.00")
    xlSheet.Cells(ctr, 23).VALUE = Format(rs!Allowance, "##,##0.00")
    xlSheet.Cells(ctr, 24).VALUE = Format(rs!EcolaPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 25).VALUE = Format(rs!GrossPay, "##,##0.00")
    xlSheet.Cells(ctr, 26).VALUE = Format(rs!WTaxAmount, "##,##0.00")
    xlSheet.Cells(ctr, 27).VALUE = Format(rs!SSSEmployee, "##,##0.00")
    xlSheet.Cells(ctr, 28).VALUE = Format(rs!PHealthEmployee, "##,##0.00")
    xlSheet.Cells(ctr, 29).VALUE = Format(rs!pagibigemployee, "##,##0.00")
    
    xlSheet.Cells(ctr, 30).VALUE = Format(rs!TASAesla, "##,##0.00")
    xlSheet.Cells(ctr, 31).VALUE = Format(rs!CashAdvance, "##,##0.00")
    xlSheet.Cells(ctr, 32).VALUE = Format(rs!MiscFee, "##,##0.00")
    xlSheet.Cells(ctr, 33).VALUE = Format(rs!LMCDed, "##,##0.00")
    
    xlSheet.Cells(ctr, 34).VALUE = Format(rs!SalaryLoan, "##,##0.00")
    xlSheet.Cells(ctr, 35).VALUE = Format(rs!Vale_IOU, "##,##0.00")
    xlSheet.Cells(ctr, 36).VALUE = Format(rs!Uniform, "##,##0.00")
    xlSheet.Cells(ctr, 37).VALUE = Format(rs!Medical, "##,##0.00")
    xlSheet.Cells(ctr, 38).VALUE = Format(rs!Variances, "##,##0.00")
    xlSheet.Cells(ctr, 39).VALUE = Format(rs!OverUnderPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 40).VALUE = Format(rs!SSSLoan, "##,##0.00")
    xlSheet.Cells(ctr, 41).VALUE = Format(rs!PagibigLoan, "##,##0.00")
    xlSheet.Cells(ctr, 42).VALUE = Format(rs!ValeWithClient, "##,##0.00")
    xlSheet.Cells(ctr, 43).VALUE = Format(rs!Deductions, "##,##0.00")
    xlSheet.Cells(ctr, 44).VALUE = Format(rs!NetPay, "##,##0.00")
    xlSheet.Cells(ctr, 45).VALUE = IIf(rs!AcctNo = "", "n/a", rs!AcctNo)

    xlSheet.Cells(ctr, 46).VALUE = Format(rs!Month13Payment, "##,##0.00")
    If rs!SSSEmployer <> 0 Then
        xlSheet.Cells(ctr, 47).VALUE = Format(rs!SSSEmployer - IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer), "##,##0.00")
        xlSheet.Cells(ctr, 48).VALUE = Format(IIf(rs!ECCEmployer = 0, 10, rs!ECCEmployer), "##,##0.00")
    Else
        xlSheet.Cells(ctr, 47).VALUE = Format(0, "##,##0.00")
        xlSheet.Cells(ctr, 48).VALUE = Format(0, "##,##0.00")
    End If
    xlSheet.Cells(ctr, 49).VALUE = Format(rs!PHealthEmployer, "##,##0.00")
    xlSheet.Cells(ctr, 50).VALUE = Format(rs!PagibigEmployer, "##,##0.00")
    xlSheet.Cells(ctr, 51).VALUE = Format(rs!CalamityLoan, "##,##0.00")
    xlSheet.Cells(ctr, 52).FormulaR1C1 = "=RC[-21]+RC[-5]+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
    xlSheet.Cells(ctr, 53).FormulaR1C1 = "=RC[-2]+RC[-1]"
    
    'xlsheetActiveCell.FormulaR1C1 = "=RC[-21]+RC[-5]+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
    
    'xlSheet.Cells.EntireColumn.AutoFit
    'xlSheet.Cells.EntireRow.Font.Name = "Arial"
    'xlSheet.Cells.EntireRow.Font.Size = 8
    
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext

Loop

'Compute for Grand Total.
EndSw = False
sumString = "=SUM(R[" & i * -1 & "]C:R[-2]C)"
mRange = "B" & i + 7 & ":B" & i + 7
xlSheet.Range(mRange).Select
xlApp1.ActiveCell.FormulaR1C1 = "           TOTAL"

mRange = "A" & i + 7 & ":A" & i + 7
xlSheet.Range(mRange).Select
xlApp1.ActiveCell.FormulaR1C1 = i - 1

For j = Asc("E") To Asc("Z")
    If EndSw = False And Chr(j) = "Z" Then
        EndSw = True
        'xlApp1.ActiveCell.FormulaR1C1 = atmctr
    End If
    If EndSw = True Then
        If Chr(j) = "Z" Then
            mRange = Chr(j) & i + 7 & ":" & Chr(j) & i + 7
        Else
            mRange = "A" & Chr(j) & i + 7 & ":A" & Chr(j) & i + 7
        End If
    Else
        mRange = Chr(j) & i + 7 & ":" & Chr(j) & i + 7
    End If
    xlSheet.Range(mRange).Select
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
    xlSheet.Range(mRange).Select
    If xlApp1.ActiveCell.VALUE = 0 Then
        xlApp1.Selection.EntireColumn.Hidden = True
    End If
Next
For j = Asc("A") To Asc("R")    ' K
    mRange = "A" & Chr(j) & i + 7 & ":A" & Chr(j) & i + 7
    xlSheet.Range(mRange).Select
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
    xlSheet.Range(mRange).Select
    If xlApp1.ActiveCell.VALUE = 0 Then
        xlApp1.Selection.EntireColumn.Hidden = True
    End If
Next
For j = Asc("T") To Asc("Z")    ' K
    mRange = "A" & Chr(j) & i + 7 & ":A" & Chr(j) & i + 7
    xlSheet.Range(mRange).Select
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
    xlSheet.Range(mRange).Select
    If xlApp1.ActiveCell.VALUE = 0 Then
        xlApp1.Selection.EntireColumn.Hidden = True
    End If
Next
For j = Asc("A") To Asc("A")    ' K
    mRange = "B" & Chr(j) & i + 7 & ":B" & Chr(j) & i + 7
    xlSheet.Range(mRange).Select
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
    xlSheet.Range(mRange).Select
    If xlApp1.ActiveCell.VALUE = 0 Then
        xlApp1.Selection.EntireColumn.Hidden = True
    End If
Next

mRange = "A5:BA" & i + 7
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
End With

'mRange = "A7:AO" & i + 7
'xlSheet.Range(mRange).Select
'xlApp1.Selection.Columns.AutoFit
'With xlApp1.Selection
'    .HorizontalAlignment = xlCenter
'    .VerticalAlignment = xlBottom
'    .WrapText = False
'    .Orientation = 0
'    .AddIndent = False
'    .ShrinkToFit = False
'    .MergeCells = False
'End With

'Format amount to 0.00.
mRange = "E7:BA" & i + 7    ' E7:AP
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
With xlApp1.Selection
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .MergeCells = False
    .Style = "Comma"
End With

'Format amount to 0.00.
'mRange = "AL7:AK" & i + 7   ' AR7:AY
'xlSheet.Range(mRange).Select
'xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'With xlApp1.Selection
'    .HorizontalAlignment = xlRight
'    .VerticalAlignment = xlBottom
'    .WrapText = False
'    .Orientation = 0
'    .AddIndent = False
'    .ShrinkToFit = False
'    .MergeCells = False
'End With


'xlSheet.Range(mRange).Select
'xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'With xlApp1.Selection.Borders(xlEdgeLeft)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With
'With xlApp1.Selection.Borders(xlEdgeTop)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With
'With xlApp1.Selection.Borders(xlEdgeBottom)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With
'With xlApp1.Selection.Borders(xlEdgeRight)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With
'xlApp1.Selection.Borders(xlInsideVertical).LineStyle = xlNone
'xlApp1.Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

'For j = Asc("A") To Asc("Z")
'    If Chr(j) = "S" Or Chr(j) = "T" Or Chr(j) = "U" Then
'        mRange = Chr(j) & "6:" & Chr(j) & "6"
'    Else
'        mRange = Chr(j) & "5:" & Chr(j) & "6"
'    End If
'    xlApp1.Range(mRange).Select
'    xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    With xlApp1.Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlApp1.Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlApp1.Selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlApp1.Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    xlApp1.Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
'Next
'
'For j = Asc("A") To Asc("K")
'    If Chr(j) = "M" Or Chr(j) = "N" Or Chr(j) = "O" Or Chr(j) = "P" Then
'        mRange = "A" & Chr(j) & "6:A" & Chr(j) & "6"
'    Else
'        mRange = "A" & Chr(j) & "5:A" & Chr(j) & "6"
'    End If
'    xlApp1.Range(mRange).Select
'    xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    With xlApp1.Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlApp1.Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlApp1.Selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlApp1.Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    xlApp1.Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
'Next

'mRange = "S6:U6"
'    xlApp1.Range(mRange).Select
'    xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    With xlApp1.Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlApp1.Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlApp1.Selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlApp1.Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    xlApp1.Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

'mRange = "AO5:AP6"
'    xlApp1.Range(mRange).Select
'    xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    With xlApp1.Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlApp1.Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlApp1.Selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlApp1.Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    xlApp1.Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

'mRange = "S5:U5"
'xlSheet.Range(mRange).Select
'xlApp1.Selection.Merge
'mRange = "AM5:AP5"
'xlSheet.Range(mRange).Select
'xlApp1.Selection.Merge


'mRange = "A7:AK" & i + 7
'xlSheet.Range(mRange).Select
'xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'With xlApp1.Selection.Borders(xlEdgeLeft)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With
'With xlApp1.Selection.Borders(xlEdgeTop)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With
'With xlApp1.Selection.Borders(xlEdgeBottom)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With
'With xlApp1.Selection.Borders(xlEdgeRight)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With
'With xlApp1.Selection.Borders(xlInsideVertical)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With
'With xlApp1.Selection.Borders(xlInsideHorizontal)
'    .LineStyle = xlContinuous
'    .Weight = xlThin
'    .ColorIndex = xlAutomatic
'End With

'START
'    With xlApp1
'    .Range("A5:AK6").Select
'    .Range("S5").Activate
'    End With
'    With xlApp1.Selection.Interior
'        .ColorIndex = 15
'        .Pattern = xlSolid
'
'    '.Range("A5:AU6").Select
'    'Range("AU6").Activate
'   ' .Selection.Font.Bold = True
'    End With
''END

mRange = "A7:A" & i + 8
xlSheet.Range(mRange).Select
xlApp1.Selection.HorizontalAlignment = xlCenter
mRange = "C7:C" & i + 7
xlSheet.Range(mRange).Select
xlApp1.Selection.HorizontalAlignment = xlCenter
mRange = "D7:D" & i + 7
xlSheet.Range(mRange).Select
xlApp1.Selection.HorizontalAlignment = xlCenter
' Center the BDO
mRange = "AS7:AS" & i + 7
xlSheet.Range(mRange).Select
xlApp1.Selection.HorizontalAlignment = xlCenter

'mRange = "B7:B" & i + 7
'xlSheet.Range(mRange).Select
'With xlApp1.Selection
'    .HorizontalAlignment = xlLeft
'    .VerticalAlignment = xlBottom
'    .WrapText = False
'    .Orientation = 0
'    .AddIndent = False
'    .IndentLevel = 0
'    .ShrinkToFit = False
'    .MergeCells = False
'End With
'
'xlSheet.Columns("E:E").ColumnWidth = 6#
'xlSheet.Columns("F:F").ColumnWidth = 8.29
'xlSheet.Columns("R:R").ColumnWidth = 7.3
'xlSheet.Columns("V:V").ColumnWidth = 6.51
'xlSheet.Columns("X:X").ColumnWidth = 6.51
'xlSheet.Columns("Q:Q").ColumnWidth = 9.01
'xlSheet.Columns("Z:Z").ColumnWidth = 6.5
'xlSheet.Columns("AA:AA").ColumnWidth = 6.5
'xlSheet.Columns("AF:AF").ColumnWidth = 6.5
'xlSheet.Columns("AG:AG").ColumnWidth = 6.5
'xlSheet.Columns("AJ:AJ").ColumnWidth = 8.51

'Bold the total.
mRange = "B" & i + 7 & ":BA" & i + 7
xlSheet.Range(mRange).Select
xlApp1.ActiveCell.Font.Bold = True

' Box IT.
mRange = "A5:BA" & i + 7
Call PutBox(xlSheet, xlApp1, mRange)

'Print SubTotal.
bdopos = 0
atmpos = 0
cashpos = 0
holdpos = 0
ctr = 9

For j = 1 To 41
    If atmctr > 0 Then
        If atmpos = 0 Then
            atmpos = 1
        End If
        ctr = 8 + atmpos
        xlSheet.Cells(i + ctr, j + 4).VALUE = Format(SubATM(j), "##,##0.00")
        If j = 4 Then
           xlSheet.Cells(i + ctr, j).VALUE = atmctr
        End If
        If j = 41 Then
           xlSheet.Cells(i + ctr, j + 4).VALUE = "ATM"
        End If
    End If
    If bdoctr > 0 Then
        If bdopos = 0 Then
            bdopos = 1
        End If
        ctr = 8 + atmpos + bdopos
        xlSheet.Cells(i + ctr, j + 4).VALUE = Format(SubBDO(j), "##,##0.00")
        If j = 4 Then
           xlSheet.Cells(i + ctr, j).VALUE = bdoctr
        End If
        If j = 41 Then
           xlSheet.Cells(i + ctr, j + 4).VALUE = "BDO"
        End If
    End If
    If cashctr > 0 Then
        If cashpos = 0 Then
            cashpos = 1
        End If
        ctr = 8 + atmpos + bdopos + cashpos
        xlSheet.Cells(i + ctr, j + 4).VALUE = Format(SubCash(j), "##,##0.00")
        If j = 4 Then
           xlSheet.Cells(i + ctr, j).VALUE = cashctr
        End If
        If j = 41 Then
           xlSheet.Cells(i + ctr, j + 4).VALUE = "Cash"
        End If
    End If
    If holdctr > 1 Then
        If holdpos = 0 Then
            holdpos = 1
        End If
        ctr = 8 + atmpos + bdopos + cashpos + holdpos
        xlSheet.Cells(i + ctr, j + 4).VALUE = Format(SubHold(j), "##,##0.00")
        If j = 4 Then
           xlSheet.Cells(i + ctr, j).VALUE = holdctr
        End If
        If j = 41 Then
           xlSheet.Cells(i + ctr, j + 4).VALUE = "Hold"
        End If
    End If
Next

i = i + 5
'Print for Prepared by:, Checked by: and Approved by:
mRange = "B" & i + 10
xlSheet.Range(mRange) = "Prepared by:"
mRange = "F" & i + 10
xlSheet.Range(mRange) = "Checked by:"
mRange = "K" & i + 10
xlSheet.Range(mRange) = "Approved by:"

mRange = "B" & i + 12
xlSheet.Range(mRange) = "      CHERRY      "
xlSheet.Range(mRange).Font.Underline = xlUnderlineStyleSingle
mRange = "F" & i + 12
xlSheet.Range(mRange) = " HR Coordinator   "
xlSheet.Range(mRange).Font.Underline = xlUnderlineStyleSingle
mRange = "K" & i + 12
xlSheet.Range(mRange) = "   EOP / NLW      "
xlSheet.Range(mRange).Font.Underline = xlUnderlineStyleSingle

'Change the font.
mRange = "B" & i + 2 & ":BA" & i + 14
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = True       'Font.Bold
    .Font.Italic = False    'Font.Italic
End With

With xlApp1.Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .MergeCells = False
    .Style = "Comma"
End With

'mRange = "E" & i + 2 & ":AK" & i + 14
'xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"

xlSheet.Range("A2").Select

Call SaveUserTransactions("Generate Payroll Summary for " & mCompanyName, "Print Payroll Summary.")

errR:

    Set xlApp1 = Nothing
    Set xlBook1 = Nothing
    Set xlSheet = Nothing
    Exit Sub
    
'    Range("C7").Select
'    Selection.Subtotal GroupBy:=3, Function:=xlCount, TotalList:=Array(7), _
'        Replace:=True, PageBreaks:=True, SummaryBelowData:=True
'    ActiveWindow.SelectedSheets.PrintPreview
    
Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
    GoTo errR
    
'xlSheet.Name = "Overtime"
End Sub

Private Sub PayrollSummaryShort(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim sumString As String
Dim ctr As Long
Dim i As Integer, j As Integer
Dim cPos As Integer
Dim LineCtr As Integer
Dim EndSw As Boolean
Dim SubBDO(12) As Double
Dim SubATM(12) As Double
Dim SubCash(12) As Double
Dim SubHold(12) As Double
Dim bdoctr, atmctr, cashctr, holdctr

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

'Set xlSheet = book1.Worksheets.Add
Set xlSheet = xlBook1.Worksheets(1)

cPos = 5
ctr = 7
LineCtr = 0
i = 1
For i = 1 To 12
    SubBDO(i) = 0
    SubATM(i) = 0
    SubCash(i) = 0
    SubHold(i) = 0
    bdoctr = 0
    atmctr = 0
    cashctr = 0
    holdctr = 0
Next

'On Error GoTo Errh
' Print Header
'xlSheet.Range("A2") = mCompanyName & mSapi  '" (SAPI)"
xlSheet.Range("A2") = mCompanyName & " (" & CompanyShort & ")"

With xlSheet.Range("A2")
    .Font.Name = "Arial"  'Font.Name
    .Font.Size = 14 'font.size
    .Font.Bold = True ' Font.Bold
    .Font.Italic = False 'font.Italic
    '.HorizontalAlignment = xlHAlignLeft 'left
    '.HorizontalAlignment = xlHAlignCenter 'center
    '.HorizontalAlignment = xlHAlignRight 'right
End With
xlSheet.Range("A3") = "Payroll Summary Report for Cut-Off " & Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "dd, yyyy")
With xlSheet.Range("A3")
    .Font.Name = "Arial"  'Font.Name
    .Font.Size = 10 'font.size
    .Font.Bold = False ' Font.Bold
    .Font.Italic = False 'font.Italic
    '.HorizontalAlignment = xlHAlignLeft 'left
    '.HorizontalAlignment = xlHAlignCenter 'center
    '.HorizontalAlignment = xlHAlignRight 'right
End With
'With xlApp1.ActiveSheet.PageSetup
'With xlSheet.PageSetup
'    .Orientation = xlLandscape
'    .TopMargin = 0.75 ' Application.InchesToPoints(CDbl(cboTopMargin))
'    .LeftMargin = Application.InchesToPoints(0.29)
'    '.LeftMargin = 3.75 'Application.InchesToPoints(CDbl(cboLeftMargin))
'    .RightMargin = 0.75 'Application.InchesToPoints(CDbl(cboRightMargin))
'    .BottomMargin = 0.75 'Application.InchesToPoints(CDbl(cboBottomMargin))
'    .HeaderMargin = 0.5 'Application.InchesToPoints(CDbl(cboHeaderMargin))
'    .FooterMargin = 0.5 'Application.InchesToPoints(CDbl(cboFooterMargin))
'End With

xlSheet.Cells(cPos, 1).VALUE = "Head"
xlSheet.Cells(cPos, 2).VALUE = "Control"
xlSheet.Cells(cPos, 4).VALUE = "Tax"
xlSheet.Cells(cPos, 5).VALUE = "Daily"
xlSheet.Cells(cPos, 6).VALUE = "Reg."
xlSheet.Cells(cPos, 7).VALUE = "This Pay"

xlSheet.Cells(cPos, 8).VALUE = "Absent"
xlSheet.Cells(cPos, 9).VALUE = "Tardiness"
xlSheet.Cells(cPos, 10).VALUE = "Reg OT"
xlSheet.Cells(cPos, 11).VALUE = "Reg OT"
xlSheet.Cells(cPos, 12).VALUE = "Night Diff"
xlSheet.Cells(cPos, 13).VALUE = "Night Diff"
xlSheet.Cells(cPos, 14).VALUE = "Rest Day"
xlSheet.Cells(cPos, 15).VALUE = "Rest Day"
xlSheet.Cells(cPos, 16).VALUE = "Sunday OT"
xlSheet.Cells(cPos, 17).VALUE = "Sunday OT"
xlSheet.Cells(cPos, 18).VALUE = "Spl Holiday"
xlSheet.Cells(cPos, 19).VALUE = "Spl Holiday"
xlSheet.Cells(cPos, 20).VALUE = "Lgl Holiday"
xlSheet.Cells(cPos, 21).VALUE = "Lgl Holiday"
xlSheet.Cells(cPos, 22).VALUE = "Number of"
xlSheet.Cells(cPos, 23).VALUE = "Sick Leave"
xlSheet.Cells(cPos, 24).VALUE = "Number of"
xlSheet.Cells(cPos, 25).VALUE = "Vac. Leave"
xlSheet.Cells(cPos, 26).VALUE = "Salary"
xlSheet.Cells(cPos, 27).VALUE = "Incentive"
xlSheet.Cells(cPos, 28).VALUE = "ECOLA"

xlSheet.Cells(cPos, 30).VALUE = "WTAX"
xlSheet.Cells(cPos, 31).VALUE = "Employee"
xlSheet.Cells(cPos, 34).VALUE = "Cash Bond"
xlSheet.Cells(cPos, 35).VALUE = "Placement"
xlSheet.Cells(cPos, 36).VALUE = "Misc. Fee"
xlSheet.Cells(cPos, 37).VALUE = "Card"
xlSheet.Cells(cPos, 38).VALUE = "SSS Loan"
xlSheet.Cells(cPos, 39).VALUE = "PBIG Loan"
xlSheet.Cells(cPos, 40).VALUE = "Vale"
xlSheet.Cells(cPos, 41).VALUE = "Adjustment"
xlSheet.Cells(cPos, 42).VALUE = "TOTAL"
'xlSheet.Cells(cPos, 42).Value = "NET PAY"
xlSheet.Cells(cPos, 44).VALUE = "ChinaBank/BDO"

'xlSheet.Cells(cPos, 44).VALUE = "13th Month"
'xlSheet.Cells(cPos, 45).VALUE = "Employer"
'xlSheet.Cells(cPos, 49).VALUE = "TOTAL"
'xlSheet.Cells(cPos, 50).VALUE = "Agency"
'xlSheet.Cells(cPos, 51).VALUE = "GROSS"
'xlSheet.Cells(cPos, 52).Value = "plus"
'xlSheet.Cells(cPos, 53).Value = "NET"

'xlSheet.Cells(cPos, 28).Value = "13th Month"
'xlSheet.Cells(cPos, 29).Value = "Cash Bond"
'xlSheet.Cells(cPos, 30).Value = "Vale"
'xlSheet.Cells(cPos, 31).Value = "Misc. Fee"
'xlSheet.Cells(cPos, 32).Value = "Card"
'xlSheet.Cells(cPos, 33).Value = "Placement"
'xlSheet.Cells(cPos, 34).Value = "Adjustment"
'xlSheet.Cells(cPos, 35).Value = "Employee"
'xlSheet.Cells(cPos, 38).Value = "Employer"
'xlSheet.Cells(cPos, 41).Value = "WTAX"
'xlSheet.Cells(cPos, 42).Value = "SSS Loan"
'xlSheet.Cells(cPos, 43).Value = "PBIG Loan"
'xlSheet.Cells(cPos, 45).Value = "TOTAL"
'xlSheet.Cells(cPos, 47).Value = "Agency"
'xlSheet.Cells(cPos, 48).Value = "ChinaBank/BDO"

cPos = cPos + 1
xlSheet.Cells(cPos, 1).VALUE = "Count"
xlSheet.Cells(cPos, 2).VALUE = "No."
xlSheet.Cells(cPos, 3).VALUE = "Employee Name"
xlSheet.Cells(cPos, 4).VALUE = "Status"
xlSheet.Cells(cPos, 5).VALUE = "Rate"
xlSheet.Cells(cPos, 6).VALUE = "Days"
xlSheet.Cells(cPos, 7).VALUE = "Period"
xlSheet.Cells(cPos, 8).VALUE = "Deduction"
xlSheet.Cells(cPos, 9).VALUE = "Deduction"
xlSheet.Cells(cPos, 10).VALUE = "Hours"
xlSheet.Cells(cPos, 11).VALUE = "Amount"
xlSheet.Cells(cPos, 12).VALUE = "OT Hours"
xlSheet.Cells(cPos, 13).VALUE = "OT Amount"
xlSheet.Cells(cPos, 14).VALUE = "OT Hours"
xlSheet.Cells(cPos, 15).VALUE = "OT Amount"
xlSheet.Cells(cPos, 16).VALUE = "Hours"
xlSheet.Cells(cPos, 17).VALUE = "Amount"
xlSheet.Cells(cPos, 18).VALUE = "OT Hours"
xlSheet.Cells(cPos, 19).VALUE = "OT Amount"
xlSheet.Cells(cPos, 20).VALUE = "OT Hours"
xlSheet.Cells(cPos, 21).VALUE = "OT Amount"
xlSheet.Cells(cPos, 22).VALUE = "Sick Leave"
xlSheet.Cells(cPos, 23).VALUE = "Amount"
xlSheet.Cells(cPos, 24).VALUE = "Vac. Leave"
xlSheet.Cells(cPos, 25).VALUE = "Amount"
xlSheet.Cells(cPos, 26).VALUE = "Adjustment"
xlSheet.Cells(cPos, 27).VALUE = "Amount"
xlSheet.Cells(cPos, 28).VALUE = "Payment"

xlSheet.Cells(cPos, 29).VALUE = "GROSS PAY"
xlSheet.Cells(cPos, 30).VALUE = "Amount"
xlSheet.Cells(cPos, 31).VALUE = "SSS"
xlSheet.Cells(cPos, 32).VALUE = "PhilHealth"
xlSheet.Cells(cPos, 33).VALUE = "Pag-ibig"
xlSheet.Cells(cPos, 34).VALUE = "Amount"
xlSheet.Cells(cPos, 35).VALUE = "Fee"
xlSheet.Cells(cPos, 36).VALUE = "Amount"
xlSheet.Cells(cPos, 37).VALUE = "Trans Fee"
xlSheet.Cells(cPos, 38).VALUE = "Amount"
xlSheet.Cells(cPos, 39).VALUE = "Amount"
xlSheet.Cells(cPos, 40).VALUE = "w/ Client"
xlSheet.Cells(cPos, 41).VALUE = "Pay"
xlSheet.Cells(cPos, 42).VALUE = "Deductions"
xlSheet.Cells(cPos, 43).VALUE = "NET PAY"
xlSheet.Cells(cPos, 44).VALUE = "Account Number"
xlSheet.Cells(cPos, 45).VALUE = "Payment"
xlSheet.Cells(cPos, 46).VALUE = "SSS"
xlSheet.Cells(cPos, 47).VALUE = "EC"
xlSheet.Cells(cPos, 48).VALUE = "PhilHealth"
xlSheet.Cells(cPos, 49).VALUE = "Pag-ibig"
xlSheet.Cells(cPos, 50).VALUE = "Charges"
xlSheet.Cells(cPos, 51).VALUE = "Charge"
xlSheet.Cells(cPos, 52).VALUE = "BILLING"
'xlSheet.Cells(cPos, 52).Value = "VAT"
'xlSheet.Cells(cPos, 53).Value = "BILLING"

'xlSheet.Cells(cPos, 28).Value = "Payment"
'xlSheet.Cells(cPos, 29).Value = "Amount"
'xlSheet.Cells(cPos, 30).Value = "w/ Client"
'xlSheet.Cells(cPos, 31).Value = "Amount"
'xlSheet.Cells(cPos, 32).Value = "Trans Fee"
'xlSheet.Cells(cPos, 33).Value = "Fee"
'xlSheet.Cells(cPos, 34).Value = "Pay"
'xlSheet.Cells(cPos, 35).Value = "SSS"
'xlSheet.Cells(cPos, 36).Value = "PhilHealth"
'xlSheet.Cells(cPos, 37).Value = "Pag-ibig"
'xlSheet.Cells(cPos, 38).Value = "SSS"
'xlSheet.Cells(cPos, 39).Value = "PhilHealth"
'xlSheet.Cells(cPos, 40).Value = "Pag-ibig"
'xlSheet.Cells(cPos, 41).Value = "Amount"
'xlSheet.Cells(cPos, 42).Value = "Amount"
'xlSheet.Cells(cPos, 43).Value = "Amount"
'xlSheet.Cells(cPos, 44).Value = "GROSS PAY"
'xlSheet.Cells(cPos, 45).Value = "Deductions"
'xlSheet.Cells(cPos, 46).Value = "NET PAY"
'xlSheet.Cells(cPos, 47).Value = "Charge"
'xlSheet.Cells(cPos, 48).Value = "Account Number"

Do While Not rs.EOF
    'DoEvents
    LineCtr = LineCtr + 1
    xlSheet.Cells(ctr, 1).VALUE = LineCtr
    xlSheet.Cells(ctr, 2).VALUE = rs!ControlNum
    xlSheet.Cells(ctr, 3).VALUE = rs!EmployeeName
    xlSheet.Cells(ctr, 4).VALUE = rs!EmpCivilStatus
    xlSheet.Cells(ctr, 5).VALUE = rs!DailyRate
    xlSheet.Cells(ctr, 6).VALUE = rs!DaysPresent
    
    If rs!PaymentSystem = "BDO" Then
        bdoctr = 1
        SubBDO(1) = SubBDO(1) + rs!DaysPresent
        SubBDO(2) = SubBDO(2) + rs!BasicPay
        SubBDO(3) = SubBDO(3) + rs!GrossPay
        SubBDO(4) = SubBDO(4) + rs!WTaxAmount
        SubBDO(5) = SubBDO(5) + rs!SSSEmployee
        SubBDO(6) = SubBDO(6) + rs!PHealthEmployee
        SubBDO(7) = SubBDO(7) + rs!pagibigemployee
        SubBDO(8) = SubBDO(8) + rs!TASAesla
        SubBDO(9) = SubBDO(9) + rs!LMCDed
        SubBDO(10) = SubBDO(10) + rs!ValeWithClient
        SubBDO(11) = SubBDO(11) + rs!Deductions
        SubBDO(12) = SubBDO(12) + rs!NetPay
    ElseIf rs!PaymentSystem = "ATM" Then
        atmctr = 1
        SubATM(1) = SubATM(1) + rs!DaysPresent
        SubATM(2) = SubATM(2) + rs!BasicPay
        SubATM(3) = SubATM(3) + rs!GrossPay
        SubATM(4) = SubATM(4) + rs!WTaxAmount
        SubATM(5) = SubATM(5) + rs!SSSEmployee
        SubATM(6) = SubATM(6) + rs!PHealthEmployee
        SubATM(7) = SubATM(7) + rs!pagibigemployee
        SubATM(8) = SubATM(8) + rs!TASAesla
        SubATM(9) = SubATM(9) + rs!LMCDed
        SubATM(10) = SubATM(10) + rs!ValeWithClient
        SubATM(11) = SubATM(11) + rs!Deductions
        SubATM(12) = SubATM(12) + rs!NetPay
    ElseIf rs!PaymentSystem = "Cash" Then
        cashctr = 1
        SubCash(1) = SubCash(1) + rs!DaysPresent
        SubCash(2) = SubCash(2) + rs!BasicPay
        SubCash(3) = SubCash(3) + rs!GrossPay
        SubCash(4) = SubCash(4) + rs!WTaxAmount
        SubCash(5) = SubCash(5) + rs!SSSEmployee
        SubCash(6) = SubCash(6) + rs!PHealthEmployee
        SubCash(7) = SubCash(7) + rs!pagibigemployee
        SubCash(8) = SubCash(8) + rs!TASAesla
        SubCash(9) = SubCash(9) + rs!LMCDed
        SubCash(10) = SubCash(10) + rs!ValeWithClient
        SubCash(11) = SubCash(11) + rs!Deductions
        SubCash(12) = SubCash(12) + rs!NetPay
    ElseIf rs!PaymentSystem = "HOLD" Then
        holdctr = 1
        SubHold(1) = SubHold(1) + rs!DaysPresent
        SubHold(2) = SubHold(2) + rs!BasicPay
        SubHold(3) = SubHold(3) + rs!GrossPay
        SubHold(4) = SubHold(4) + rs!WTaxAmount
        SubHold(5) = SubHold(5) + rs!SSSEmployee
        SubHold(6) = SubHold(6) + rs!PHealthEmployee
        SubHold(7) = SubHold(7) + rs!pagibigemployee
        SubHold(8) = SubHold(8) + rs!TASAesla
        SubHold(9) = SubHold(9) + rs!LMCDed
        SubHold(10) = SubHold(10) + rs!ValeWithClient
        SubHold(11) = SubHold(11) + rs!Deductions
        SubHold(12) = SubHold(12) + rs!NetPay
    End If
    
'    xlSheet.Cells(ctr, 6).VALUE = Format(rs!BasicPay, "##,##0.00")
'    xlSheet.Cells(ctr, 7).VALUE = Format(rs!AbsentDeduction, "##,#0.00")
'    xlSheet.Cells(ctr, 8).VALUE = Format(rs!TardinessDeduction, "##,#0.00")
'    xlSheet.Cells(ctr, 9).VALUE = Format(rs!OTHours, "##,##0.00")
'    xlSheet.Cells(ctr, 10).VALUE = Format(rs!OTPay, "##,##0.00")
'    xlSheet.Cells(ctr, 11).VALUE = Format(rs!NightDiffOT, "##,##0.00")
'    xlSheet.Cells(ctr, 12).VALUE = Format(rs!NightDiffPay, "##,##0.00")
'    xlSheet.Cells(ctr, 13).VALUE = Format(rs!TaxRefundHours, "##,##0.00")
'    xlSheet.Cells(ctr, 14).VALUE = Format(rs!TaxRefundPay, "##,##0.00")
'    xlSheet.Cells(ctr, 15).VALUE = Format(rs!SundayOT, "##,##0.00")
'    xlSheet.Cells(ctr, 16).VALUE = Format(rs!SundayOTPay, "##,##0.00")
'    xlSheet.Cells(ctr, 17).VALUE = Format(rs!NoRegHoliday, "##,##0.00")
'    xlSheet.Cells(ctr, 18).VALUE = Format(rs!RegHolidayPay, "##,##0.00")
'    xlSheet.Cells(ctr, 19).VALUE = Format(rs!NoSplHoliday, "##,##0.00")
'    xlSheet.Cells(ctr, 20).VALUE = Format(rs!SplHolidayPay, "##,##0.00")
'    xlSheet.Cells(ctr, 21).VALUE = Format(rs!NoSickLeave, "##,##0.00")
'    xlSheet.Cells(ctr, 22).VALUE = Format(rs!SickLeavePay, "##,##0.00")
'    xlSheet.Cells(ctr, 23).VALUE = Format(rs!NoVacationLeave, "##,##0.00")
'    xlSheet.Cells(ctr, 24).VALUE = Format(rs!VacationLeavePay, "##,##0.00")
'    xlSheet.Cells(ctr, 25).VALUE = Format(rs!OtherIncome, "##,##0.00")
'    xlSheet.Cells(ctr, 26).VALUE = Format(rs!Allowance, "##,##0.00")
'    xlSheet.Cells(ctr, 27).VALUE = Format(rs!EcolaPay, "##,##0.00")
    
    xlSheet.Cells(ctr, 7).VALUE = Format(rs!BasicPay, "##,##0.00")
    xlSheet.Cells(ctr, 8).VALUE = Format(rs!GrossPay, "##,##0.00")
    xlSheet.Cells(ctr, 9).VALUE = Format(rs!WTaxAmount, "##,##0.00")
    xlSheet.Cells(ctr, 10).VALUE = Format(rs!SSSEmployee, "##,##0.00")
    xlSheet.Cells(ctr, 11).VALUE = Format(rs!PHealthEmployee, "##,##0.00")
    xlSheet.Cells(ctr, 12).VALUE = Format(rs!pagibigemployee, "##,##0.00")
    xlSheet.Cells(ctr, 13).VALUE = Format(rs!TASAesla, "##,##0.00")
    'xlSheet.Cells(ctr, 34).VALUE = Format(rs!CashAdvance, "##,##0.00")
    'xlSheet.Cells(ctr, 35).VALUE = Format(rs!MiscFee, "##,##0.00")
    xlSheet.Cells(ctr, 14).VALUE = Format(rs!LMCDed, "##,##0.00")
    'xlSheet.Cells(ctr, 37).VALUE = Format(rs!SSSLoan, "##,##0.00")
    'xlSheet.Cells(ctr, 38).VALUE = Format(rs!PagibigLoan, "##,##0.00")
    xlSheet.Cells(ctr, 15).VALUE = Format(rs!ValeWithClient, "##,##0.00")
    'xlSheet.Cells(ctr, 40).VALUE = Format(rs!OverUnderPay, "##,##0.00")
    xlSheet.Cells(ctr, 16).VALUE = Format(rs!Deductions, "##,##0.00")
    xlSheet.Cells(ctr, 17).VALUE = Format(rs!NetPay, "##,##0.00")
    xlSheet.Cells(ctr, 18).VALUE = IIf(rs!AcctNo = "", "n/a", rs!AcctNo)
    'xlSheet.Cells(ctr, 44).VALUE = Format(rs!Month13Payment, "##,##0.00")
'    If rs!sssemployer <> 0 Then
'        xlSheet.Cells(ctr, 45).VALUE = Format(rs!sssemployer - 10, "##,##0.00")
'        xlSheet.Cells(ctr, 46).VALUE = Format(10, "##,##0.00")
'    Else
'        xlSheet.Cells(ctr, 45).VALUE = Format(0, "##,##0.00")
'    End If
'    xlSheet.Cells(ctr, 47).VALUE = Format(rs!PHealthEmployer, "##,##0.00")
'    xlSheet.Cells(ctr, 48).VALUE = Format(rs!PagibigEmployer, "##,##0.00")
'    xlSheet.Cells(ctr, 50).VALUE = Format(rs!CalamityLoan, "##,##0.00")
'    xlSheet.Cells(ctr, 49).FormulaR1C1 = "=RC[-21]+RC[-5]+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
'    xlSheet.Cells(ctr, 51).FormulaR1C1 = "=RC[-2]+RC[-1]"
    'xlsheetActiveCell.FormulaR1C1 = "=RC[-21]+RC[-5]+RC[-4]+RC[-3]+RC[-2]+RC[-1]"
    
    'xlSheet.Cells.EntireColumn.AutoFit
    'xlSheet.Cells.EntireRow.Font.Name = "Arial"
    'xlSheet.Cells.EntireRow.Font.Size = 8
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext
Loop

'Compute for Grand Total.
EndSw = False
sumString = "=SUM(R[" & i * -1 & "]C:R[-2]C)"
xlSheet.Cells(i + 3, 2).VALUE = "Total"
xlSheet.Cells(i + 3, 2).Font.Name = "Arial"  'Font.Name
xlSheet.Cells(i + 3, 2).Font.Size = 14 'font.size
xlSheet.Cells(i + 3, 2).Font.Bold = True ' Font.Bold
For j = Asc("E") To Asc("Z")
    If EndSw = False And Chr(j) = "Z" Then
        EndSw = True
    End If
    If EndSw = True Then
        mRange = "A" & Chr(j) & i + 7 & ":A" & Chr(j) & i + 7
    Else
        mRange = Chr(j) & i + 7 & ":" & Chr(j) & i + 7
    End If
    xlSheet.Range(mRange).Select
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
Next
For j = Asc("A") To Asc("Y")
    mRange = "A" & Chr(j) & i + 7 & ":A" & Chr(j) & i + 7
    xlSheet.Range(mRange).Select
    xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"
Next

'Print SubTotal
i = i + 4 + 7
If atmctr = 1 Then
    i = i + 1
    xlSheet.Cells(i, 5).VALUE = Format(SubATM(1), "##,##0.00")
    xlSheet.Cells(i, 6).VALUE = Format(SubATM(2), "##,##0.00")
    xlSheet.Cells(i, 7).VALUE = Format(SubATM(3), "##,#0.00")
    xlSheet.Cells(i, 8).VALUE = Format(SubATM(4), "##,#0.00")
    xlSheet.Cells(i, 9).VALUE = Format(SubATM(5), "##,##0.00")
    xlSheet.Cells(i, 10).VALUE = Format(SubATM(6), "##,##0.00")
    xlSheet.Cells(i, 11).VALUE = Format(SubATM(7), "##,##0.00")
    xlSheet.Cells(i, 12).VALUE = Format(SubATM(8), "##,##0.00")
    xlSheet.Cells(i, 13).VALUE = Format(SubATM(9), "##,##0.00")
    xlSheet.Cells(i, 14).VALUE = Format(SubATM(10), "##,##0.00")
    xlSheet.Cells(i, 15).VALUE = Format(SubATM(11), "##,##0.00")
    xlSheet.Cells(i, 16).VALUE = Format(SubATM(12), "##,##0.00")
End If

If bdoctr = 1 Then
    i = i + 1
    xlSheet.Cells(i, 5).VALUE = Format(SubBDO(1), "##,##0.00")
    xlSheet.Cells(i, 6).VALUE = Format(SubBDO(2), "##,##0.00")
    xlSheet.Cells(i, 7).VALUE = Format(SubBDO(3), "##,#0.00")
    xlSheet.Cells(i, 8).VALUE = Format(SubBDO(4), "##,#0.00")
    xlSheet.Cells(i, 9).VALUE = Format(SubBDO(5), "##,##0.00")
    xlSheet.Cells(i, 10).VALUE = Format(SubBDO(6), "##,##0.00")
    xlSheet.Cells(i, 11).VALUE = Format(SubBDO(7), "##,##0.00")
    xlSheet.Cells(i, 12).VALUE = Format(SubBDO(8), "##,##0.00")
    xlSheet.Cells(i, 13).VALUE = Format(SubBDO(9), "##,##0.00")
    xlSheet.Cells(i, 14).VALUE = Format(SubBDO(10), "##,##0.00")
    xlSheet.Cells(i, 15).VALUE = Format(SubBDO(11), "##,##0.00")
    xlSheet.Cells(i, 16).VALUE = Format(SubBDO(12), "##,##0.00")
End If
If cashctr = 1 Then
    i = i + 1
    xlSheet.Cells(i, 5).VALUE = Format(SubCash(1), "##,##0.00")
    xlSheet.Cells(i, 6).VALUE = Format(SubCash(2), "##,##0.00")
    xlSheet.Cells(i, 7).VALUE = Format(SubCash(3), "##,#0.00")
    xlSheet.Cells(i, 8).VALUE = Format(SubCash(4), "##,#0.00")
    xlSheet.Cells(i, 9).VALUE = Format(SubCash(5), "##,##0.00")
    xlSheet.Cells(i, 10).VALUE = Format(SubCash(6), "##,##0.00")
    xlSheet.Cells(i, 11).VALUE = Format(SubCash(7), "##,##0.00")
    xlSheet.Cells(i, 12).VALUE = Format(SubCash(8), "##,##0.00")
    xlSheet.Cells(i, 13).VALUE = Format(SubCash(9), "##,##0.00")
    xlSheet.Cells(i, 14).VALUE = Format(SubCash(10), "##,##0.00")
    xlSheet.Cells(i, 15).VALUE = Format(SubCash(11), "##,##0.00")
    xlSheet.Cells(i, 16).VALUE = Format(SubCash(12), "##,##0.00")
End If
If holdctr = 1 Then
    i = i + 1
    xlSheet.Cells(i, 5).VALUE = Format(SubHold(1), "##,##0.00")
    xlSheet.Cells(i, 6).VALUE = Format(SubHold(2), "##,##0.00")
    xlSheet.Cells(i, 7).VALUE = Format(SubHold(3), "##,#0.00")
    xlSheet.Cells(i, 8).VALUE = Format(SubHold(4), "##,#0.00")
    xlSheet.Cells(i, 9).VALUE = Format(SubHold(5), "##,##0.00")
    xlSheet.Cells(i, 10).VALUE = Format(SubHold(6), "##,##0.00")
    xlSheet.Cells(i, 11).VALUE = Format(SubHold(7), "##,##0.00")
    xlSheet.Cells(i, 12).VALUE = Format(SubHold(8), "##,##0.00")
    xlSheet.Cells(i, 13).VALUE = Format(SubHold(9), "##,##0.00")
    xlSheet.Cells(i, 14).VALUE = Format(SubHold(10), "##,##0.00")
    xlSheet.Cells(i, 15).VALUE = Format(SubHold(11), "##,##0.00")
    xlSheet.Cells(i, 16).VALUE = Format(SubHold(12), "##,##0.00")
End If

'Print for Prepared by:, Checked by: and Approved by:
mRange = "B" & i + 10
xlSheet.Range(mRange) = "Prepared by:"
mRange = "F" & i + 10
xlSheet.Range(mRange) = "Checked by:"
mRange = "K" & i + 10
xlSheet.Range(mRange) = "Approved by:"

mRange = "B" & i + 13
xlSheet.Range(mRange) = "         lai"
mRange = "F" & i + 13
xlSheet.Range(mRange) = "EOP          /JQP"
mRange = "K" & i + 13
xlSheet.Range(mRange) = "      NLW"

mRange = "B" & i + 14
xlSheet.Range(mRange) = "____________"
mRange = "F" & i + 14
xlSheet.Range(mRange) = "____________"
mRange = "K" & i + 14
xlSheet.Range(mRange) = "____________"

mRange = "A5:AY" & i + 7
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 8          'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
End With

xlSheet.Range(mRange).Select
xlApp1.Selection.Columns.AutoFit
With xlApp1.Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .MergeCells = False
End With

'Format amount to 0.00.
mRange = "E7:AP" & i + 7
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
With xlApp1.Selection
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .MergeCells = False
End With

mRange = "AR7:AY" & i + 7
xlSheet.Range(mRange).Select
xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
With xlApp1.Selection
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .ShrinkToFit = False
    .MergeCells = False
End With


mRange = "A5:AY6"
xlSheet.Range(mRange).Select
xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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
xlApp1.Selection.Borders(xlInsideVertical).LineStyle = xlNone
xlApp1.Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

For j = Asc("A") To Asc("Y")
'    If Chr(j) = "H" Or Chr(j) = "I" Or Chr(j) = "M" Or Chr(j) = "N" Or Chr(j) = "O" Then
'        mRange = Chr(j) & "6:" & Chr(j) & "6"
'    Else
        mRange = Chr(j) & "5:" & Chr(j) & "6"
    'End If
    xlApp1.Range(mRange).Select
    xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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
    xlApp1.Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
Next

For j = Asc("A") To Asc("Y")
    If Chr(j) = "D" Or Chr(j) = "E" Or Chr(j) = "F" Or Chr(j) = "S" Or Chr(j) = "T" Or Chr(j) = "U" Or Chr(j) = "V" Then
        mRange = "A" & Chr(j) & "6:A" & Chr(j) & "6"
    Else
        mRange = "A" & Chr(j) & "5:A" & Chr(j) & "6"
    End If
    xlApp1.Range(mRange).Select
    xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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
    xlApp1.Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
Next

mRange = "AD5:AF5"
xlSheet.Range(mRange).Select
xlApp1.Selection.Merge
mRange = "AS5:AV5"
xlSheet.Range(mRange).Select
xlApp1.Selection.Merge

mRange = "A7:AY" & i + 7
xlSheet.Range(mRange).Select
xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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

'START
    With xlApp1
    .Range("A5:AY6").Select
    .Range("S5").Activate
    End With
    With xlApp1.Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
        
        '.Range("A5:AU6").Select
    'Range("AU6").Activate
   ' .Selection.Font.Bold = True
    End With
'END
mRange = "B7:B" & i + 7
xlSheet.Range(mRange).Select
With xlApp1.Selection
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .MergeCells = False
End With

xlSheet.Columns("E:E").ColumnWidth = 6#
xlSheet.Columns("F:F").ColumnWidth = 8.29
xlSheet.Columns("AO:AO").ColumnWidth = 7.57
xlSheet.Columns("AS:AS").ColumnWidth = 8.89
xlSheet.Columns("AT:AT").ColumnWidth = 7.86
xlSheet.Columns("AW:AW").ColumnWidth = 9.01
xlSheet.Columns("AY:AY").ColumnWidth = 9.01

xlSheet.Range("A2").Select

errR:

    Set xlApp1 = Nothing
    Set xlBook1 = Nothing
    Set xlSheet = Nothing
    Exit Sub
'    Range("C7").Select
'    Selection.Subtotal GroupBy:=3, Function:=xlCount, TotalList:=Array(7), _
'        Replace:=True, PageBreaks:=True, SummaryBelowData:=True
'    ActiveWindow.SelectedSheets.PrintPreview
    
Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
    GoTo errR
    
'xlSheet.Name = "Overtime"
End Sub


Private Sub PayrollBankCopy(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mGrandPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim sumString As String
Dim ctr As Long
Dim i As Integer, j As Integer
Dim cPos As Integer

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

'Set xlSheet = book1.Worksheets.Add
Set xlSheet = xlBook1.Worksheets(1)
'xlSheet.Name = "PayrollBankCopy"

cPos = 5
ctr = 7
i = 1

On Error GoTo Errh
' Print Header
xlSheet.Range("A2") = mCompanyName '& mSapi  '" (SAPI)"
With xlSheet.Range("A2")
    .Font.Name = "Arial"  'Font.Name
    .Font.Size = 14 'font.size
    .Font.Bold = True ' Font.Bold
    .Font.Italic = False 'font.Italic
    '.HorizontalAlignment = xlHAlignLeft 'left
    '.HorizontalAlignment = xlHAlignCenter 'center
    '.HorizontalAlignment = xlHAlignRight 'right
End With
xlSheet.Range("A3") = "Payroll Period: " & Format(rs!StartPayDate, "mm/dd/yyyy") & " to " & Format(rs!EndPayDate, "mm/dd/yyyy")
With xlSheet.Range("A3")
    .Font.Name = "Arial"  'Font.Name
    .Font.Size = 10 'font.size
    .Font.Bold = False ' Font.Bold
    .Font.Italic = False 'font.Italic
    '.HorizontalAlignment = xlHAlignLeft 'left
    '.HorizontalAlignment = xlHAlignCenter 'center
    '.HorizontalAlignment = xlHAlignRight 'right
End With
'With xlApp1.ActiveSheet.PageSetup
'    .Orientation = xlPortrait 'Landscape
'    .TopMargin = 5#  ' Application.InchesToPoints(CDbl(cboTopMargin))
'    .LeftMargin = 1 'Application.InchesToPoints(0.29)
'    '.LeftMargin = 3.75 'Application.InchesToPoints(CDbl(cboLeftMargin))
'    .RightMargin = 1#    'Application.InchesToPoints(CDbl(cboRightMargin))
'    .BottomMargin = 0.75 'Application.InchesToPoints(CDbl(cboBottomMargin))
'    .HeaderMargin = 0.5 'Application.InchesToPoints(CDbl(cboHeaderMargin))
'    .FooterMargin = 0.5 'Application.InchesToPoints(CDbl(cboFooterMargin))
'End With

xlSheet.Cells(cPos, 1).VALUE = "Account No."
xlSheet.Cells(cPos, 2).VALUE = "Amount"
xlSheet.Cells(cPos, 3).VALUE = "Employee Name"
mGrandPay = 0
Do While Not rs.EOF
    xlSheet.Cells(ctr, 1).VALUE = IIf(rs!AcctNo = "" Or IsNull(rs!AcctNo), "0000000000000000", rs!AcctNo)
    'mTotalPay = rs!BasicPay + rs!OTHours + rs!OTPay + rs!EcolaPay - (rs!AbsentDeduction + rs!TardinessDeduction)
    mTotalPay = rs!NetPay
    ' Total Deduction
    mTotalDeduction = rs!WTaxAmount + rs!SSSEmployee + rs!PHealthEmployee + rs!pagibigemployee + rs!TASAesla + rs!LMCDed + rs!ValeWithClient + rs!MiscFee + rs!CashAdvance
    'xlSheet.Cells(ctr, 2).VALUE = Format(mTotalPay - mTotalDeduction, "##,##0.00")
    xlSheet.Cells(ctr, 2).VALUE = Format(mTotalPay, "##,##0.00")
    xlSheet.Cells(ctr, 3).VALUE = rs!EmployeeName
    'mGrandPay = mGrandPay + (mTotalPay - mTotalDeduction)
    mGrandPay = mGrandPay + mTotalPay
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext
Loop

xlSheet.Cells(ctr + 1, 1).VALUE = "Total ==>"
mRange = "B" & i + 7
xlSheet.Range(mRange).Select
sumString = "=SUM(R[" & i * -1 & "]C:R[-2]C)"
xlApp1.ActiveCell.FormulaR1C1 = sumString '"=SUM(R[i-6]C:R[-2]C)"

mRange = "A5:A" & i + 7
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 10         'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
    .ColumnWidth = 20.4
End With
mRange = "B5:B" & i + 7
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 10         'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
    .ColumnWidth = 12
    .NumberFormat = "#,##0.00_);(#,##0.00)"
End With
mRange = "C5:C" & i + 7
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 10         'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
    .ColumnWidth = 25
End With
xlSheet.Range("A1").Select

errR:

    Set xlApp1 = Nothing
    Set xlBook1 = Nothing
    Set xlSheet = Nothing
    Exit Sub

Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
    GoTo errR
    
'xlSheet.Name = "Overtime"

End Sub

Private Sub EmployeePayrollList(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mGrandPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim sumString As String
Dim ctr As Long
Dim i As Integer, j As Integer
Dim cPos As Integer

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()

xlApp1.WindowState = xlMaximized
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

'Set xlSheet = book1.Worksheets.Add
Set xlSheet = xlBook1.Worksheets(1)

cPos = 5
ctr = 7
i = 1

On Error GoTo Errh
' Print Header
xlSheet.Range("A2") = mCompanyName '& mSapi  '" (SAPI)"
With xlSheet.Range("A2")
    .Font.Name = "Arial"  'Font.Name
    .Font.Size = 14 'font.size
    .Font.Bold = True ' Font.Bold
    .Font.Italic = False 'font.Italic
    '.HorizontalAlignment = xlHAlignLeft 'left
    '.HorizontalAlignment = xlHAlignCenter 'center
    '.HorizontalAlignment = xlHAlignRight 'right
End With
xlSheet.Range("A3") = "Payroll Period: " & Format(rs!StartPayDate, "mm/dd/yyyy") & " to " & Format(rs!EndPayDate, "mm/dd/yyyy")
With xlSheet.Range("A3")
    .Font.Name = "Arial"  'Font.Name
    .Font.Size = 10 'font.size
    .Font.Bold = False ' Font.Bold
    .Font.Italic = False 'font.Italic
    '.HorizontalAlignment = xlHAlignLeft 'left
    '.HorizontalAlignment = xlHAlignCenter 'center
    '.HorizontalAlignment = xlHAlignRight 'right
End With
'With xlApp1.ActiveSheet.PageSetup
'    .Orientation = xlPortrait 'Landscape
'    .TopMargin = 5#  ' Application.InchesToPoints(CDbl(cboTopMargin))
'    .LeftMargin = 1 'Application.InchesToPoints(0.29)
'    '.LeftMargin = 3.75 'Application.InchesToPoints(CDbl(cboLeftMargin))
'    .RightMargin = 1#    'Application.InchesToPoints(CDbl(cboRightMargin))
'    .BottomMargin = 0.75 'Application.InchesToPoints(CDbl(cboBottomMargin))
'    .HeaderMargin = 0.5 'Application.InchesToPoints(CDbl(cboHeaderMargin))
'    .FooterMargin = 0.5 'Application.InchesToPoints(CDbl(cboFooterMargin))
'End With

xlSheet.Cells(cPos, 1).VALUE = "Account No."
xlSheet.Cells(cPos, 2).VALUE = "Amount"
xlSheet.Cells(cPos, 3).VALUE = "Employee Name"
    
Do While Not rs.EOF
    xlSheet.Cells(ctr, 1).VALUE = IIf(rs!AcctNo = "", "0000000000000000", rs!AcctNo)
    mTotalPay = rs!BasicPay + rs!OTHours + rs!OTPay + rs!EcolaPay - (rs!AbsentDeduction + rs!TardinessDeduction)
    xlSheet.Cells(ctr, 2).VALUE = Format(mTotalPay, "##,##0.00")
    xlSheet.Cells(ctr, 3).VALUE = rs!EmployeeName
    mGrandPay = mGrandPay + mTotalPay
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext
Loop
xlSheet.Cells(ctr + 1, 1).VALUE = "Total ==>"
xlSheet.Cells(ctr + 1, 2).VALUE = Format(mGrandPay, "###,##0.00")

'mRange = "E" & i + 7 & ":R" & i + 7
'xlSheet.Range(mRange).Select
'xlApp1.Selection.HorizontalAlignment = xlCenter
'xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"

mRange = "A5:A" & i + 7
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 10         'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
    .ColumnWidth = 14.4
End With
mRange = "B5:B" & i + 7
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 10         'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
    .ColumnWidth = 12
    .NumberFormat = "#,##0.00_);(#,##0.00)"
End With
mRange = "C5:C" & i + 7
xlSheet.Range(mRange).Select
With xlApp1.Range(mRange)
    .Font.Name = "Arial"    'Font.Name
    .Font.Size = 10         'Font.size
    .Font.Bold = False      'Font.Bold
    .Font.Italic = False    'Font.Italic
    .ColumnWidth = 25
End With
xlSheet.Range("A1").Select

errR:

    Set xlApp1 = Nothing
    Set xlBook1 = Nothing
    Set xlSheet = Nothing
    Exit Sub
'    Range("C7").Select
'    Selection.Subtotal GroupBy:=3, Function:=xlCount, TotalList:=Array(7), _
'        Replace:=True, PageBreaks:=True, SummaryBelowData:=True
'    ActiveWindow.SelectedSheets.PrintPreview
    
Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
    GoTo errR
    
'xlSheet.Name = "Overtime"

End Sub

'Sub ExcelPageSetup()
'
'    With ActiveSheet.PageSetup
'        .PrintTitleRows = "$1:$4"
'        .PrintTitleColumns = ""
'    End With
'    ActiveSheet.PageSetup.PrintArea = ""
'    With ActiveSheet.PageSetup
'        .LeftHeader = ""
'        .CenterHeader = "Run Date / Time: &date (&time)"
'        .RightHeader = ""
'        .LeftFooter = ""
'        .CenterFooter = ""
'        .RightFooter = "Page: &page / &pages"
'        .LeftMargin = Application.InchesToPoints(1.04166666666667E-02)
'        .RightMargin = Application.InchesToPoints(1.04166666666667E-02)
'        .TopMargin = Application.InchesToPoints(1.04166666666667E-02)
'        .BottomMargin = Application.InchesToPoints(1.04166666666667E-02)
'        .HeaderMargin = Application.InchesToPoints(6.94444444444444E-03)
'        .FooterMargin = Application.InchesToPoints(6.94444444444444E-03)
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        .PrintQuality = Array(120, 144)
'        .CenterHorizontally = False
'        .CenterVertically = False
'        .Orientation = xlLandscape
'        .Draft = False
'        .PaperSize = xlPaperLetter
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = False
'        .Zoom = 100
'    End With
'End Sub
'
Private Sub optReport_Click(Index As Integer)

'MsgBox fraButton.Top & "  " & fraButton.Height
    If Index = 1 Then ' statement of account
        TxtAmount1.Text = Format(BillAmount, "#,##0.00")
         FraStatement.Visible = True
         Call setCutoff
         FraDebit.Visible = False
         Me.Height = 4935
         fraButton.Top = 4045 ' 4095
         Frame1.Height = 3930
'    ElseIf Index = 6 Then
'        TxtAmount.Text = Format(BillAmount, "#,##0.00")
'        Call showBank
'        cboBank.ListIndex = 0
'        Call setCutoff
'        FraStatement.Visible = False
'        FraDebit.Visible = True
'        Me.Height = 4935
'        fraButton.Top = 4045 ' 4095
'        Frame1.Height = 3930
    Else
       ' cmbEmployee.Visible = True
        FraStatement.Visible = False
        FraDebit.Visible = False
        Me.Height = 3330
        fraButton.Top = 2300  ' 2350 ' 2430
        Frame1.Height = 2130
    End If
        
End Sub

Private Sub setCutoff()
Me.txtCutoffDate(0).Text = frmSelectCutoff.txtCutoffDate(0).Text
Me.txtCutoffDate(1).Text = frmSelectCutoff.txtCutoffDate(1).Text
Me.txtCutOffDate1(0).Text = frmSelectCutoff.txtCutoffDate(0).Text
Me.txtCutOffDate1(1).Text = frmSelectCutoff.txtCutoffDate(1).Text
End Sub

Private Sub TxtAmount_GotFocus()

  TxtAmount.BackColor = &HFFFFC0
  Set oCommonFunction.ValidDouble = TxtAmount
  Call txtHLigths(TxtAmount)
  
End Sub

Private Sub TxtAmount_KeyPress(KeyAscii As Integer)
Dim strValid As String
strValid = "0123456789." & Chr(13) & Chr(8)
If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
If KeyAscii = 13 Then cboBank.SetFocus
End Sub

Private Sub TxtAmount_LostFocus()
TxtAmount.Text = Format(TxtAmount, "#,##0.00")
TxtAmount.BackColor = &HFFFFFF
End Sub

Private Sub TxtAmount1_GotFocus()
TxtAmount1.BackColor = &HFFFFC0
Call txtHLigths(TxtAmount1)
End Sub

Private Sub TxtAmount1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtAmount1_LostFocus()
    TxtAmount1.BackColor = &HFFFFFF
    Set oCommonFunction.ValidDouble = TxtAmount1
End Sub

Private Sub txtCreditDate_GotFocus()
  SendKeys "{HOME}+{END}"
  txtCreditDate.BackColor = &HFFFFC0
  Call mskHLigths(txtCreditDate)
End Sub

Private Sub txtCreditDate_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub txtCreditDate_LostFocus()
txtCreditDate.BackColor = &HFFFFFF
End Sub

Private Sub txtCreditDate_Validate(Cancel As Boolean)
If txtCreditDate <> "__/__/____" And IsDate(txtCreditDate) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: txtCreditDate.SetFocus
End Sub

Private Sub txtCutoffDate_GotFocus(Index As Integer)

  txtCutoffDate(Index).BackColor = &HFFFFC0
  Call mskHLigths(txtCutoffDate(Index))
End Sub

Private Sub txtCutoffDate_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCutOffDate_LostFocus(Index As Integer)
 txtCutoffDate(Index).BackColor = vbWhite
End Sub

Private Sub txtCutOffDate_Validate(Index As Integer, Cancel As Boolean)
If txtCutoffDate(Index) <> "__/__/____" And IsDate(txtCutoffDate(Index)) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: txtCutoffDate(Index).SetFocus

End Sub

Private Sub txtCutOffDate1_GotFocus(Index As Integer)
  SendKeys "{HOME}+{END}"
  txtCutOffDate1(Index).BackColor = &HFFFFC0
  Call mskHLigths(txtCutOffDate1(Index))
End Sub

Private Sub txtCutOffDate1_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCutOffDate1_LostFocus(Index As Integer)
   txtCutOffDate1(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtCutOffDate1_Validate(Index As Integer, Cancel As Boolean)
If txtCutOffDate1(Index) <> "__/__/____" And IsDate(txtCutOffDate1(Index)) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: txtCutOffDate1(Index).SetFocus

End Sub

Private Sub TxtDMNum_GotFocus()
  SendKeys "{HOME}+{END}"
  TxtDMNum.BackColor = &HFFFFC0
  Call txtHLigths(TxtDMNum)
End Sub

Private Sub TxtDMNum_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtDMNum_LostFocus()
TxtDMNum.BackColor = &HFFFFFF
End Sub

Private Sub TxtSANum_GotFocus()
  SendKeys "{HOME}+{END}"
  TxtSANum.BackColor = &HFFFFC0
  Call txtHLigths(TxtSANum)
End Sub

Private Sub TxtSANum_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtSANum_LostFocus()
TxtSANum.BackColor = &HFFFFFF
End Sub

Private Sub txtSOA_GotFocus()
    SendKeys "{HOME}+{END}"
    txtSOA.BackColor = &HFFFFC0
    Call txtHLigths(txtSOA)
End Sub

Private Sub txtSOA_LostFocus()
    txtSOA.BackColor = &HFFFFFF
End Sub

Private Sub txtSOA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub mskDueDate_GotFocus()
    SendKeys "{HOME}+{END}"
    mskDueDate.BackColor = &HFFFFC0
Call mskHLigths(mskDueDate)
End Sub

Private Sub mskDueDate_LostFocus()
    mskDueDate.BackColor = &HFFFFFF
End Sub

Private Sub mskDueDate_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub PayrollSummaryRegister2(ByRef rsPayroll As ADODB.Recordset)
Dim Status As String
Dim cAppl As CRAXDRT.Application
Dim cReport As CRAXDRT.Report
Dim rs As New ADODB.Recordset
Dim str As String

Unload Me

On Error GoTo errHandler
Me.MousePointer = vbHourglass

    Set cAppl = New CRAXDRT.Application

    Set cReport = cAppl.OpenReport(App.Path & "\Reports\Payroll Register.rpt")

    cReport.DiscardSavedData
    cReport.Database.SetDataSource rsPayroll

    frmReportViewer.CRViewer91.ReportSource = cReport
    frmReportViewer.CRViewer91.ViewReport

    frmReportViewer.CRViewer91.Zoom 75

    Do While frmReportViewer.CRViewer91.IsBusy
        DoEvents
    Loop

    frmReportViewer.Show

Set cAppl = Nothing
Set cReport = Nothing

Me.MousePointer = vbDefault

Exit Sub
errHandler:
    
    Me.MousePointer = vbDefault
    
    MsgBox "Error : " & Err.Number & ";" & Err.Description, vbExclamation + vbOKOnly, "Error"
    Err.Clear
    
Set cAppl = Nothing
Set cReport = Nothing

End Sub

Private Sub ConnectToSQL()

Set CON = New ADODB.Connection
CON.CursorLocation = adUseClient
CON.Provider = "SQLOLEDB"
CON.Open "File name=" & App.Path & "\ReportConnection.udl"
'CON.Open ReportConnection

ReportConnection = CON.ConnectionString

Exit Sub

errcon:

MsgBox "Unable to to establish the connection! Please check connection settings.", vbCritical, PayrollAlert
CON.Close

Err.Clear

'ShellExecute hWnd, "open", App.Path & "\connection.udl", vbNullString, vbNullString, conSwNormal


End Sub

Private Sub PayrollSummaryRegister(ByRef rsPayroll As ADODB.Recordset, _
    Optional ReportName As String = "Payroll Register.rpt")

Dim rsReport As New ADODB.Recordset
Dim cApp As CRAXDRT.Application
Dim cAppl As CRAXDRT.Application
Dim cReport As CRAXDRT.Report
Dim StartDate As String
Dim EndDate As String
Dim strPay As String

On Error GoTo errHandler
    
StartDate = frmSelectCutoff.txtCutoffDate(0).Text
EndDate = frmSelectCutoff.txtCutoffDate(1).Text
'Update the records.
strPay = "EXEC GetPayrollEntries 1, 1, " & mClientCode & ", '" & StartDate & "', '" & EndDate & "', '" & mUserID & "'"
Call Action_Query(strPay, cnnstr)

ConnectToSQL

strPay = "SELECT * FROM tmpPayrollTemplate ORDER BY EmployeeName"
'Set rsReport = ReturnRst(strPay, cnnstr, mstatic)
Set rsReport = ReturnRst(strPay, ReportConnection, mstatic)

If rsReport.RecordCount = 0 Then
    MsgBox "No record/s selected! Cannot proceed.", vbExclamation + vbOKOnly, PayrollAlert
    rsReport.Close
    Set rsReport = Nothing
    Exit Sub
End If

Unload Me
Me.MousePointer = vbHourglass

    Set cAppl = New CRAXDRT.Application

    'Set cReport = cAppl.OpenReport(App.Path & "\Reports\Payroll Register.rpt")
    Set cReport = cAppl.OpenReport(App.Path & "\Reports\" & ReportName)

    cReport.DiscardSavedData
    cReport.Database.SetDataSource rsReport

    frmReportViewer.CRViewer91.ReportSource = cReport
    frmReportViewer.CRViewer91.ViewReport
    
'    frmReportViewer.CRViewer91.Height = Me.Height - 150
'    frmReportViewer.CRViewer91.Width = Me.Width - 150
    
    frmReportViewer.CRViewer91.Zoom 90

    Do While frmReportViewer.CRViewer91.IsBusy
        DoEvents
    Loop

    frmReportViewer.Show

Set cAppl = Nothing
Set cReport = Nothing
rsReport.Close
Set rsReport = Nothing
CON.Close

Me.MousePointer = vbDefault

Exit Sub
errHandler:
    
    Me.MousePointer = vbDefault
    
    MsgBox "Error : " & Err.Number & ";" & Err.Description, vbExclamation + vbOKOnly, "Error"
    Err.Clear
    
Set cAppl = Nothing
Set cReport = Nothing
    
End Sub

