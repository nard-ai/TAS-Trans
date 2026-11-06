VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCompany 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Information"
   ClientHeight    =   7890
   ClientLeft      =   3165
   ClientTop       =   2910
   ClientWidth     =   10365
   Icon            =   "frmCompany.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10365
   Begin VB.CommandButton cmdlast 
      Height          =   255
      Left            =   9840
      Picture         =   "frmCompany.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Last record"
      Top             =   495
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Height          =   255
      Left            =   9600
      Picture         =   "frmCompany.frx":0A14
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Next record"
      Top             =   495
      Width           =   255
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   255
      Left            =   9360
      Picture         =   "frmCompany.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Previous record"
      Top             =   495
      Width           =   255
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   255
      Left            =   9120
      Picture         =   "frmCompany.frx":0CA8
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Top record"
      Top             =   495
      Width           =   255
   End
   Begin VB.Frame chkPagibigFlagchkPagibigFlag 
      BackColor       =   &H80000005&
      Height          =   7125
      Left            =   120
      TabIndex        =   27
      Top             =   600
      Width           =   10005
      Begin VB.CheckBox chkTaxFlag 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7080
         TabIndex        =   64
         Top             =   6600
         Width           =   855
      End
      Begin VB.CheckBox chkSSSFlag 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7080
         TabIndex        =   63
         Top             =   6240
         Width           =   1935
      End
      Begin VB.CheckBox chkPhilHealthFlag 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7080
         TabIndex        =   62
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CheckBox chkPagibigFlag 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7080
         TabIndex        =   61
         Top             =   5520
         Width           =   1935
      End
      Begin VB.CheckBox chkEcolaFlag 
         Caption         =   "Check1"
         Height          =   255
         Left            =   4800
         TabIndex        =   60
         Top             =   6240
         Width           =   1935
      End
      Begin VB.CheckBox chkEditTaxFlag 
         Caption         =   "Check1"
         Height          =   255
         Left            =   4800
         TabIndex        =   59
         Top             =   6600
         Width           =   1935
      End
      Begin VB.CheckBox chkAgencyStatus 
         Caption         =   "Check1"
         Height          =   255
         Left            =   4800
         TabIndex        =   58
         Top             =   5880
         Width           =   1935
      End
      Begin VB.CheckBox chk13MonthFlag 
         Caption         =   "Check1"
         Height          =   255
         Left            =   4800
         TabIndex        =   57
         Top             =   5520
         Width           =   1935
      End
      Begin VB.CommandButton cmdSelectCompany 
         Caption         =   "<< Select Company >>"
         Height          =   375
         Left            =   7920
         TabIndex        =   0
         Top             =   6600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TxtCutOffDates 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Top             =   6600
         Width           =   2835
      End
      Begin VB.TextBox TxtAgencyCutOff 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Top             =   6240
         Width           =   2835
      End
      Begin VB.TextBox TxtWorkHrsPerDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   5880
         Width           =   2835
      End
      Begin VB.TextBox TxtEcolaRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Top             =   5520
         Width           =   2835
      End
      Begin VB.TextBox TxtAgencyTaxNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6840
         TabIndex        =   18
         Top             =   4920
         Width           =   2835
      End
      Begin VB.TextBox TxtPhilHealthNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   4920
         Width           =   2835
      End
      Begin VB.TextBox TxtPagibigNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6840
         TabIndex        =   17
         Top             =   4560
         Width           =   2835
      End
      Begin VB.TextBox TxtSSSNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   4590
         Width           =   2835
      End
      Begin VB.TextBox TxtBankAccountNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6840
         TabIndex        =   16
         Top             =   4200
         Width           =   2835
      End
      Begin VB.TextBox TxtBankName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   4200
         Width           =   2835
      End
      Begin VB.TextBox TxtPermitNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   3840
         Width           =   2835
      End
      Begin VB.TextBox TxtAssitantPosition 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6840
         TabIndex        =   11
         Top             =   3120
         Width           =   2835
      End
      Begin VB.TextBox TxtAssistant 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   3120
         Width           =   3075
      End
      Begin VB.TextBox TxtContactPosition 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6840
         TabIndex        =   10
         Top             =   2760
         Width           =   2835
      End
      Begin VB.TextBox TxtContactPerson 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   7
         Top             =   2760
         Width           =   3075
      End
      Begin VB.TextBox TxtOwner 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   2400
         Width           =   3075
      End
      Begin VB.TextBox TxtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1680
         Width           =   8115
      End
      Begin VB.TextBox TxtTelefax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6840
         TabIndex        =   9
         Top             =   2040
         Width           =   2835
      End
      Begin VB.TextBox TxtShort 
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   3195
      End
      Begin VB.TextBox TxtTelNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   2040
         Width           =   3075
      End
      Begin VB.TextBox txtAgencyName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   5115
      End
      Begin VB.TextBox txtAgencyCode 
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   705
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   360
         X2              =   9840
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   360
         X2              =   9840
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Image ImgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   6840
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cut Off Dates"
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
         Height          =   270
         Left            =   240
         TabIndex        =   24
         Top             =   6720
         Width           =   1140
      End
      Begin VB.Label label444 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cut Off"
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
         Height          =   270
         Left            =   720
         TabIndex        =   25
         Top             =   6360
         Width           =   660
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Work Hrs per day"
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
         Height          =   270
         Left            =   120
         TabIndex        =   26
         Top             =   6000
         Width           =   1500
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   600
         TabIndex        =   29
         Top             =   5640
         Width           =   900
      End
      Begin VB.Label TAX 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "TIN No."
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
         Height          =   270
         Left            =   6000
         TabIndex        =   30
         Top             =   4920
         Width           =   780
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "PhilHealthNo"
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
         Height          =   270
         Left            =   360
         TabIndex        =   31
         Top             =   4920
         Width           =   1020
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pag ibig"
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
         Height          =   270
         Left            =   6000
         TabIndex        =   56
         Top             =   4560
         Width           =   900
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "SSS"
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
         Height          =   270
         Left            =   1080
         TabIndex        =   55
         Top             =   4560
         Width           =   420
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account No."
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
         Height          =   270
         Left            =   5355
         TabIndex        =   54
         Top             =   4275
         Width           =   1500
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
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
         Height          =   270
         Left            =   600
         TabIndex        =   53
         Top             =   4200
         Width           =   900
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Permit No."
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
         Height          =   270
         Left            =   720
         TabIndex        =   52
         Top             =   3840
         Width           =   900
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6075
         TabIndex        =   51
         Top             =   3150
         Width           =   780
      End
      Begin VB.Label Assistant 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Assistant"
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
         Height          =   270
         Left            =   720
         TabIndex        =   50
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6075
         TabIndex        =   49
         Top             =   2790
         Width           =   825
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   48
         Top             =   2760
         Width           =   1140
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Owner"
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
         Height          =   270
         Left            =   840
         TabIndex        =   47
         Top             =   2400
         Width           =   660
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   46
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Short"
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
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   1140
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tel no."
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
         Height          =   255
         Left            =   840
         TabIndex        =   44
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telefax"
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
         Height          =   270
         Left            =   6075
         TabIndex        =   43
         Top             =   2115
         Width           =   780
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Height          =   270
         Left            =   240
         TabIndex        =   32
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Code"
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
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   1185
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton cmdDelete 
         Height          =   405
         Left            =   1410
         Picture         =   "frmCompany.frx":0DF2
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   495
         Picture         =   "frmCompany.frx":0EF4
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   945
         Picture         =   "frmCompany.frx":103E
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   45
         Picture         =   "frmCompany.frx":1570
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "New"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1845
         Picture         =   "frmCompany.frx":1AA2
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog cdgDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FSO As New FileSystemObject
Dim rst As New ADODB.Recordset
 
Private Sub cmdSelectCompany_Click()

If ReportedFlag = 1 Then ' 1 is for Philhealth
    frmPHealthReport.TxtPhilHealthNo.Text = TxtPhilHealthNo.Text
    frmPHealthReport.txtEmpTIN.Text = TxtAgencyTaxNumber

    frmPHealthReport.txtEmployerName.Text = txtAgencyName.Text
    frmPHealthReport.txtMailingAddress.Text = TxtAddress.Text
    frmPHealthReport.txtTelephoneNo.Text = TxtTelNo.Text
    frmPHealthReport.txtReceivedBy.Text = TxtAssistant.Text
    frmPHealthReport.txtPrintedName.Text = TxtContactPerson.Text
    frmPHealthReport.txtDesignation.Text = TxtContactPosition.Text
ElseIf ReportedFlag = 2 Then ' 2 is for sss
    frmSSSReport.TxtPhilHealthNo.Text = TxtSSSNo.Text
    frmSSSReport.txtEmpTIN.Text = TxtAgencyTaxNumber

    frmSSSReport.txtEmployerName.Text = txtAgencyName.Text
    frmSSSReport.txtMailingAddress.Text = TxtAddress.Text
    frmSSSReport.txtTelephoneNo.Text = TxtTelNo.Text
    frmSSSReport.txtReceivedBy.Text = TxtAssistant.Text
    frmSSSReport.txtPrintedName.Text = TxtContactPerson.Text
    frmSSSReport.txtDesignation.Text = TxtContactPosition.Text
Else    ' if 3 for pagibig
    frmHDMFReport.TxtPhilHealthNo.Text = TxtPhilHealthNo.Text
    frmHDMFReport.txtEmpTIN.Text = TxtAgencyTaxNumber
    frmHDMFReport.TxtSSSNo.Text = TxtSSSNo.Text

    frmHDMFReport.txtEmployerName.Text = txtAgencyName.Text
    frmHDMFReport.txtMailingAddress.Text = TxtAddress.Text
    frmHDMFReport.txtTelephoneNo.Text = TxtTelNo.Text
    frmHDMFReport.txtReceivedBy.Text = TxtAssistant.Text
    frmHDMFReport.txtPrintedName.Text = TxtContactPerson.Text
    frmHDMFReport.txtDesignation.Text = TxtContactPosition.Text
End If
ReportedFlag = 0
Unload Me
End Sub

Private Sub Form_Load()
TextBox (False)

Call EnableButton(True, True, False, True, False)
If ReportedFlag = 1 Or ReportedFlag = 2 Or ReportedFlag = 3 Then
    Set rst = ReturnRst("Select * FROM CompanyTable WHERE AgencyCode = " & CompanyCode & " Order by Agencycode ASC", CnnstrHR, mstatic)
Else
    Set rst = ReturnRst("Select * FROM CompanyTable Order by Agencycode ASC", CnnstrHR, mstatic)
End If
  With rst
    If .EOF = False Then
        Showdata
        Call EnableNavigator(False, True, True, True)
    Else
       clear_textbox
    End If
  End With
    If SelectCompanyFlag = False Then
        cmdSelectCompany.Visible = False
    Else
        cmdSelectCompany.Visible = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Call EnableButton(True, True, False, True, False)
    Call EnableNavigator(True, True, True, True)
    TextBox (False)
    Showdata
End Sub

Private Sub cmdDelete_Click()
    'Validate first
     Dim rsDelete As New ADODB.Recordset
     
    If DeleteRights = False Then
        MsgBox "You do not have rights to Delete Record!", vbInformation, PayrollAlert
        Exit Sub
    End If
     
     Set rsDelete = ReturnRst("Select * from HRemployees where agencycode=" & Me.txtAgencyCode.Text, cnnstr, mstatic)
     If rsDelete.EOF = False Then
         MsgBox ("You cannot delete this agency, hremployees has records!")
         Exit Sub
     End If

     If CompanyShort = UCase(RTrim(Me.TxtShort.Text)) Then
        If MsgBox("This agency you want to delete is the ACTIVE agency." & vbCrLf & "" & vbCrLf & " The system will be closed once you click YES." & vbrlf & "" & vbCrLf & "" & vbCrLf & "Are you sure you want to remove this agency?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
        End If
     Else
        If MsgBox("Are you sure you want to remove this agency?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
        End If
     End If
     
     Call Action_Query(" DELETE FROM companytable where agencycode= " & txtAgencyCode.Text & "", CnnstrHR)
     If CompanyShort = UCase(RTrim(Me.TxtShort.Text)) Then
        AddFlag = False
        addflagship = False
        If frmCompany.Visible = True Then Unload frmCompany
        frmLogin.Show
        mdiMain.Hide
     Else
        Call EnableButton(True, True, False, True, False)
        Call EnableNavigator(True, True, True, True)
        rst.Requery
        Call Form_Load
        Showdata
     End If
       

End Sub

Private Sub cmdEdit_Click()
    
    If EditRights = False Then
        MsgBox "You do not have rights to edit this Record!", vbInformation, PayrollAlert
        Exit Sub
    End If
    
    Call EnableButton(False, False, True, False, True)
    TextBox (True)
    txtAgencyCode.Enabled = False
    txtAgencyName.SetFocus
End Sub

Private Sub cmdfirst_Click()
   rst.MoveFirst
   If rst.BOF = True Then
       MsgBox "record not found"
       Call EnableNavigator(False, False, False, False)
   Else
       Showdata
       Call EnableNavigator(False, False, True, True)
   End If
End Sub

Private Sub cmdlast_Click()
   rst.MoveLast
   If rst.EOF = True Then
       MsgBox "Record not Found."
       Call EnableNavigator(False, False, False, False)
   Else
       Showdata
       Call EnableNavigator(True, True, False, False)
   End If
End Sub

Private Sub cmdNext_Click()
   rst.MoveNext
   If rst.EOF = True Then
      MsgBox "End of File."
      Call EnableNavigator(True, True, False, False)
      rst.MoveLast
   Else
      Showdata
      Call EnableNavigator(True, True, True, True)
   End If
End Sub

Private Sub cmdPrev_Click()
   rst.MovePrevious
   If rst.BOF = True Then
       MsgBox "Beginning of File."
       Call EnableNavigator(False, False, True, True)
       rst.MoveFirst
   Else
       Showdata
       Call EnableNavigator(True, True, True, True)
   End If
End Sub

Private Sub clear_textbox()
    Me.TxtAddress.Text = ""
    Me.txtAgencyCode.Text = ""
    Me.TxtAgencyCutOff.Text = ""
    Me.txtAgencyName.Text = ""
    Me.TxtAgencyTaxNumber.Text = ""
    Me.TxtAssistant.Text = ""
    Me.TxtAssitantPosition.Text = ""
    Me.TxtBankAccountNo.Text = ""
    Me.TxtBankName.Text = ""
    Me.TxtContactPerson.Text = ""
    Me.TxtContactPosition.Text = ""
    Me.TxtCutOffDates.Text = ""
    Me.TxtEcolaRate.Text = ""
    Me.TxtOwner.Text = ""
    Me.TxtPagibigNo.Text = ""
    Me.TxtPermitNo.Text = ""
    Me.TxtPhilHealthNo.Text = ""
    Me.TxtShort.Text = ""
    Me.TxtSSSNo.Text = ""
    Me.TxtTelefax.Text = ""
    Me.TxtTelNo.Text = ""
    Me.TxtWorkHrsPerDay.Text = ""
    Me.chk13MonthFlag.VALUE = 0
    Me.chkAgencyStatus.VALUE = 0
    Me.chkEcolaFlag.VALUE = 0
    Me.chkEditTaxFlag.VALUE = 0
    Me.chkPagibigFlag.VALUE = 0
    Me.chkPhilHealthFlag.VALUE = 0
    Me.chkSSSflag.VALUE = 0
    Me.chkTaxFlag.VALUE = 0

End Sub

Private Sub cmdAdd_Click()
   
    TextBox (True)
    clear_textbox
    txtAgencyCode.Text = "<auto number>"
    txtAgencyCode.Enabled = False
    chkAgencyStatus.VALUE = True
     
    TxtShort.SetFocus
    Call EnableButton(False, False, True, False, True)
    Call EnableNavigator(False, False, False, False)
    
End Sub

Private Function isValidate() As String
     isValidate = False
     If Trim(txtAgencyName.Text) = "" Then
            MsgBox "Enter name.", vbInformation
            txtAgencyName.SetFocus
            Exit Function
     ElseIf Trim(TxtShort.Text) = "" Then
           MsgBox "Enter short name.", vbInformation
           TxtShort.SetFocus
                Exit Function
     ElseIf Trim(Me.cdgDialog.Filename) = "" Then
           MsgBox "Please select logo for this agency.", vbInformation
           ImgLogo_Click
            Exit Function
     End If
     
     If Trim(Me.TxtEcolaRate.Text) = "" Then
        TxtEcolaRate.Text = 0
     Else
        If IsNumeric(Trim(Me.TxtEcolaRate.Text)) = False Then
                MsgBox ("Ecola Rate should be numeric!")
                TxtEcolaRate.SetFocus
                Exit Function
        End If
     
     End If
     
     If Trim(Me.TxtWorkHrsPerDay.Text) = "" Then
        TxtWorkHrsPerDay.Text = 0
     Else
        If IsNumeric(Trim(Me.TxtWorkHrsPerDay.Text)) = False Then
                MsgBox ("Work hours per day should be numeric!")
                TxtWorkHrsPerDay.SetFocus
                Exit Function
        End If
     End If
 
    Dim rsValidate As New ADODB.Recordset
    Set rsValidate = ReturnRst("Select * from companytable where agencyshort ='" & Me.TxtShort.Text & "'", cnnstr, mstatic)
    If rsValidate.EOF = False And txtAgencyCode.Text = "<auto number>" Then
       MsgBox "Short name already exists.", vbInformation
       TxtShort.SetFocus
       isValidate = False
       Exit Function
    End If
   isValidate = True
End Function

Private Sub cmdSave_Click()
 
 If isValidate = False Then Exit Sub
    If CompanyShort = UCase(RTrim(Me.TxtShort.Text)) And Me.chkAgencyStatus.VALUE = False Then
        If MsgBox("This agency you want to DISABLE is the ACTIVE agency." & vbCrLf & "" & vbCrLf & " The system will be closed once you click YES." & vbrlf & "" & vbCrLf & "" & vbCrLf & "Are you sure you want to remove this agency?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
        End If
    End If
    
    Dim sqlUpdate As String
    
    If IsNumeric(txtAgencyCode.Text) = True Then
        sqlUpdate = "Update CompanyTable set " & _
                  " AgencyName ='" & Me.txtAgencyName.Text & "'," & _
                  " AgencyAddress ='" & Me.TxtAddress.Text & "'," & _
                  " AgencyCutOff='" & Me.TxtAgencyCutOff.Text & "'," & _
                  " AgencyTaxNumber='" & Me.TxtAgencyTaxNumber.Text & "',AgencyAssistant='" & Me.TxtAssistant.Text & "'," & _
                  " AgencyAssitantPosition='" & Me.TxtAssitantPosition.Text & "',AgencyBankAccountNo='" & Me.TxtBankAccountNo.Text & "'," & _
                  " AgencyBankName='" & Me.TxtBankName.Text & "',AgencyContactPerson='" & Me.TxtContactPerson.Text & "'," & _
                  " AgencyContactPosition='" & Me.TxtContactPosition.Text & "',AgencyCutoffDates='" & IIf(RTrim(Me.TxtCutOffDates.Text) = "", 0, Me.TxtCutOffDates.Text) & "'," & _
                  " AgencyEcolaRate=" & Me.TxtEcolaRate.Text & ",AgencyOwner='" & Me.TxtOwner.Text & "'," & _
                  " AgencyPagibigNo='" & Me.TxtPagibigNo.Text & "',AgencyPermitNo='" & Me.TxtPermitNo.Text & "'," & _
                  " AgencyPhilHealthNo='" & Me.TxtPhilHealthNo.Text & "',AgencyShort='" & UCase(RTrim(Me.TxtShort.Text)) & "'," & _
                  " AgencySSSNo='" & Me.TxtSSSNo.Text & "',AgencyTelefax= '" & Me.TxtTelefax.Text & "'," & _
                  " AgencyTelNo='" & Me.TxtTelNo.Text & "',AgencyWorkHrsperDay= " & Me.TxtWorkHrsPerDay.Text & "," & _
                  " AgencyWith13MonthFlag= " & IIf(Me.chk13MonthFlag.VALUE = True, 1, 0) & ",AgencyStatus=" & IIf(Me.chkAgencyStatus.VALUE = True, 1, 0) & "," & _
                  " AgencyEcolaFlag= " & IIf(Me.chkEcolaFlag.VALUE = True, 1, 0) & "," & _
                  " AgencyEditTaxFlag= " & IIf(Me.chkEditTaxFlag.VALUE = True, 1, 0) & ",AgencyPagibigFlag=" & IIf(Me.chkPagibigFlag.VALUE = True, 1, 0) & "," & _
                  " AgencyPhilHealthFlag =" & IIf(Me.chkPhilHealthFlag.VALUE = True, 1, 0) & ",AgencySSSFlag=" & IIf(Me.chkSSSflag.VALUE = True, 1, 0) & "," & _
                  " AgencyTaxFlag=" & IIf(Me.chkTaxFlag.VALUE = True, 1, 0) & "," & _
                  " logopath='" & Me.cdgDialog.Filename & "'" & _
                  " where agencycode=" & Me.txtAgencyCode.Text & ""
                Call Action_Query(sqlUpdate, CnnstrHR)
            MsgBox "Company Information was successfully updated.", vbInformation, PayrollAlert
'        Else
'            MsgBox "Cannot Update Company table.", vbInformation
'        End If
           
       If CompanyShort = UCase(RTrim(Me.TxtShort.Text)) And Me.chkAgencyStatus.VALUE = False Then
         AddFlag = False
            addflagship = False
            If frmCompany.Visible = True Then Unload frmCompany
            frmLogin.Show
            mdiMain.Hide
       Else
            rst.Requery
            rst.Find "agencycode = " & txtAgencyCode.Text & ""
            Call EnableButton(True, True, False, True, False)
            TextBox (False)
       End If
    Else
         sqlUpdate = "Insert into companytable (AgencyAddress,AgencyCutOff,AgencyName,AgencyTaxNumber,AgencyAssistant," & _
                     "AgencyAssitantPosition,AgencyBankAccountNo,AgencyBankName,AgencyContactPerson,AgencyContactPosition," & _
                     " AgencyCutoffDates,AgencyEcolaRate,AgencyOwner,AgencyPagibigNo,AgencyPermitNo,AgencyPhilHealthNo,AgencyShort," & _
                     "AgencySSSNo,AgencyTelefax,AgencyTelNo,AgencyWorkHrsperDay,AgencyWith13MonthFlag,AgencyStatus,AgencyEcolaFlag," & _
                     "AgencyEditTaxFlag,AgencyPagibigFlag,AgencyPhilHealthFlag,AgencySSSFlag,AgencyTaxFlag,logopath) values (" & _
                     "'" & Me.TxtAddress.Text & "','" & Me.TxtAgencyCutOff.Text & "','" & Me.txtAgencyName.Text & "'," & _
                      "'" & Me.TxtAgencyTaxNumber.Text & "','" & Me.TxtAssistant.Text & "','" & Me.TxtAssitantPosition.Text & "'," & _
                      "'" & Me.TxtBankAccountNo.Text & "','" & Me.TxtBankName.Text & "','" & Me.TxtContactPerson.Text & "'," & _
                      "'" & Me.TxtContactPosition.Text & "','" & IIf(RTrim(Me.TxtCutOffDates.Text) = "", 0, Me.TxtCutOffDates.Text) & "'," & IIf(RTrim(Me.TxtEcolaRate.Text) = "", 0, Me.TxtEcolaRate.Text) & "," & _
                     "'" & Me.TxtOwner.Text & "','" & Me.TxtPagibigNo.Text & "','" & Me.TxtPermitNo.Text & "','" & Me.TxtPhilHealthNo.Text & "'," & _
                     "'" & UCase(RTrim(Me.TxtShort.Text)) & "','" & Me.TxtSSSNo.Text & "','" & Me.TxtTelefax.Text & "','" & Me.TxtTelNo.Text & "'," & _
                     "" & IIf(RTrim(Me.TxtWorkHrsPerDay.Text) = "", 0, Me.TxtWorkHrsPerDay.Text) & "," & Me.chk13MonthFlag.VALUE & "," & Me.chkAgencyStatus.VALUE & "," & _
                     "" & Me.chkEcolaFlag.VALUE & "," & Me.chkEditTaxFlag.VALUE & "," & Me.chkPagibigFlag.VALUE & "," & _
                     "" & Me.chkPhilHealthFlag.VALUE & "," & Me.chkSSSflag.VALUE & "," & Me.chkTaxFlag.VALUE & ",'" & Me.cdgDialog.Filename & "')"

       Call Action_Query(sqlUpdate, CnnstrHR)
        
       Set FS0 = New FileSystemObject
       
       Dim src, des, fld As String
       fld = App.Path & "\" & UCase(RTrim(Me.TxtShort.Text)) & "\"
       src = App.Path & "\NEW DB\NEW.MDB"
       des = App.Path & "\" & UCase(RTrim(Me.TxtShort.Text)) & "\" & UCase(RTrim(Me.TxtShort.Text)) & ".MDB"
       
       'check for NEW.MDB
       If FSO.FileExists(src) = False Then
        MsgBox ("Cannot create the [" & des & "], because [" & src & "] is not found! ")
       Else
            If FSO.FolderExists(fld) = False Then
                FSO.CreateFolder (fld)
            End If
            
            'COPY AS NEW MDB
            FSO.CopyFile src, des, True
            SaveSetting "DataSearch", Me.TxtShort.Text, "Path2Version2", des
       End If
       
    
       MsgBox "saved", vbInformation
       Call EnableNavigator(True, True, False, False)
       rst.Requery
       Call Form_Load
       rst.MoveLast
    End If
End Sub

Private Sub EnableNavigator(Optional ByVal Blcmdfirst As Boolean, Optional ByVal BlcmdPrev As Boolean, Optional ByVal BlcmdNext As Boolean, Optional ByVal Blcmdlast As Boolean)
    cmdfirst.Enabled = Blcmdfirst
    cmdPrev.Enabled = BlcmdPrev
    cmdNext.Enabled = BlcmdNext
    cmdlast.Enabled = Blcmdlast
End Sub

Private Sub EnableButton(Optional ByVal Add As Boolean, Optional ByVal Edit As Boolean, Optional ByVal save As Boolean, Optional ByVal Delete As Boolean, Optional ByVal Cancel As Boolean)
    cmdadd.Enabled = Add
    cmdedit.Enabled = Edit
    cmdsave.Enabled = save
    cmdDelete.Enabled = Delete
    cmdcancel.Enabled = Cancel
End Sub

Private Sub Showdata()
 With rst
Me.TxtAddress.Text = IIf(IsNull(.Fields!AgencyAddress), "", .Fields!AgencyAddress)
Me.txtAgencyCode.Text = IIf(IsNull(.Fields!agencycode), "", .Fields!agencycode)
Me.TxtAgencyCutOff.Text = IIf(IsNull(.Fields!AgencyCutOff), "", .Fields!AgencyCutOff)
Me.txtAgencyName.Text = IIf(IsNull(.Fields!AgencyName), "", .Fields!AgencyName)
Me.TxtAgencyTaxNumber.Text = IIf(IsNull(.Fields!AgencyTaxNumber), "", .Fields!AgencyTaxNumber)
Me.TxtAssistant.Text = IIf(IsNull(.Fields!AgencyAssistant), "", .Fields!AgencyAssistant)
Me.TxtAssitantPosition.Text = IIf(IsNull(.Fields!AgencyAssitantPosition), "", .Fields!AgencyAssitantPosition)
Me.TxtBankAccountNo.Text = IIf(IsNull(.Fields!AgencyBankAccountNo), "", .Fields!AgencyBankAccountNo)
Me.TxtBankName.Text = IIf(IsNull(.Fields!AgencyBankName), "", .Fields!AgencyBankName)
Me.TxtContactPerson.Text = IIf(IsNull(.Fields!AgencyContactPerson), "", .Fields!AgencyContactPerson)
Me.TxtContactPosition.Text = IIf(IsNull(.Fields!AgencyContactPosition), "", .Fields!AgencyContactPosition)
Me.TxtCutOffDates.Text = IIf(IsNull(.Fields!AgencyCutoffDates) = True, "", .Fields!AgencyCutoffDates)
Me.TxtEcolaRate.Text = IIf(IsNull(.Fields!AgencyEcolaRate), 0#, .Fields!AgencyEcolaRate)
Me.TxtOwner.Text = IIf(IsNull(.Fields!AgencyOwner), "", .Fields!AgencyOwner)
Me.TxtPagibigNo.Text = IIf(IsNull(.Fields!AgencyPagibigNo), "", .Fields!AgencyPagibigNo)
Me.TxtPermitNo.Text = IIf(IsNull(.Fields!AgencyPermitNo), "", .Fields!AgencyPermitNo)
Me.TxtPhilHealthNo.Text = IIf(IsNull(.Fields!AgencyPhilHealthNo), "", .Fields!AgencyPhilHealthNo)
Me.TxtShort.Text = IIf(IsNull(UCase(RTrim(.Fields!AgencyShort))), "", UCase(RTrim(.Fields!AgencyShort)))
Me.TxtSSSNo.Text = IIf(IsNull(.Fields!AgencySSSNo), "", .Fields!AgencySSSNo)
Me.TxtTelefax.Text = IIf(IsNull(.Fields!AgencyTelefax), "", .Fields!AgencyTelefax)
Me.TxtTelNo.Text = IIf(IsNull(.Fields!AgencyTelNo), "", .Fields!AgencyTelNo)
Me.TxtWorkHrsPerDay.Text = .Fields!AgencyWorkHrsperDay
Me.chk13MonthFlag.VALUE = .Fields!AgencyWith13MonthFlag
Me.chkAgencyStatus.VALUE = .Fields!AgencyStatus
Me.chkEcolaFlag.VALUE = .Fields!AgencyEcolaFlag
Me.chkEditTaxFlag.VALUE = .Fields!AgencyEditTaxFlag
Me.chkPagibigFlag.VALUE = .Fields!AgencyPagibigFlag
Me.chkPhilHealthFlag.VALUE = .Fields!AgencyPhilHealthFlag
Me.chkSSSflag.VALUE = .Fields!AgencySSSFlag
Me.chkTaxFlag.VALUE = .Fields!AgencyTaxFlag
On Error GoTo SkipThis:
'Me.ImgLogo.Picture = LoadPicture(.Fields!logoPath)
SkipThis:
cdgDialog.Filename = .Fields!logoPath
'
'  AgencyName
'  AgencyShort
'  AgencyAddress
'  AgencyTelNo
'  AgencyTelefax
'   AgencyOwner
'   AgencyContactPerson
'   AgencyContactPosition
'   AgencyAssistant
'   AgencyAssitantPosition
'   AgencyPermitNo
'   AgencyBankName
'   AgencyBankAccountNo
'   AgencySSSNo
'   AgencyPagibigNo
'   AgencyPhilHealthNo
'   AgencyTaxNumber
'   AgencyStatus
'   AgencySSSFlag
'   AgencyPagibigFlag
'   AgencyPhilHealthFlag
'   AgencyTaxFlag
'   AgencyEditTaxFlag
'   AgencyEcolaFlag
'   AgencyEcolaRate
'   AgencyWorkHrsperDay
'   AgencyCutoff
'   AgencyCutoffDates
'   AgencyWith13MonthFlag
'   soaControlNo
'   LogoPath
'   AgencyLogo
'   AgencyFlag1
'   AgencyFlag2
'   CompanyTable
'    AgencyCode1
'    AgencyCode2
'    Createdby
'    CreatedDate
'    Updatedby
'    DateUpdated

 End With
End Sub

Private Function TextBox(ByVal txtval As Boolean)
Me.TxtAddress.Enabled = txtval
Me.txtAgencyCode.Enabled = txtval
Me.TxtAgencyCutOff.Enabled = txtval
Me.txtAgencyName.Enabled = txtval
Me.TxtAgencyTaxNumber.Enabled = txtval
Me.TxtAssistant.Enabled = txtval
Me.TxtAssitantPosition.Enabled = txtval
Me.TxtBankAccountNo.Enabled = txtval
Me.TxtBankName.Enabled = txtval
Me.TxtContactPerson.Enabled = txtval
Me.TxtContactPosition.Enabled = txtval
Me.TxtCutOffDates.Enabled = txtval
Me.TxtEcolaRate.Enabled = txtval
Me.TxtOwner.Enabled = txtval
Me.TxtPagibigNo.Enabled = txtval
Me.TxtPermitNo.Enabled = txtval
Me.TxtPhilHealthNo.Enabled = txtval
Me.TxtShort.Enabled = txtval
Me.TxtSSSNo.Enabled = txtval
Me.TxtTelefax.Enabled = txtval
Me.TxtTelNo.Enabled = txtval
Me.TxtWorkHrsPerDay.Enabled = txtval
Me.chk13MonthFlag.Enabled = txtval
Me.chkAgencyStatus.Enabled = txtval
Me.chkEcolaFlag.Enabled = txtval
Me.chkEditTaxFlag.Enabled = txtval
Me.chkPagibigFlag.Enabled = txtval
Me.chkPhilHealthFlag.Enabled = txtval
Me.chkSSSflag.Enabled = txtval
Me.chkTaxFlag.Enabled = txtval
End Function

Private Sub ImgLogo_Click()
If Me.txtAgencyName.Enabled = True Then
    cdgDialog.Flags = cdlOFNFileMustExist
    cdgDialog.Filter = "All pictures (*.jpg,*.jpeg,*.bmp,*.gif,*.tif)|" & _
   "*.jpf;*.jpeg;*.bmp;*.gif;*.tif|"
    cdgDialog.FilterIndex = 2
    'cdgDialog.Filename = txtbrowse
    cdgDialog.DialogTitle = "Select Image"
    cdgDialog.ShowOpen
    'Me.ImgLogo.Picture = LoadPicture(cdgDialog.Filename)
End If
End Sub

Private Sub TxtAddress_GotFocus()
TxtAddress.BackColor = &HFFFFC0
End Sub

Private Sub TxtAddress_LostFocus()
TxtAddress.BackColor = vbWhite
End Sub

Private Sub txtAgencyCode_GotFocus()
txtAgencyCode.BackColor = &HFFFFC0
End Sub

Private Sub txtAgencyName_GotFocus()
txtAgencyName.BackColor = &HFFFFC0
End Sub

Private Sub txtAgencyName_LostFocus()
txtAgencyName.BackColor = vbWhite
End Sub

Private Sub TxtAssistant_GotFocus()
TxtAssistant.BackColor = &HFFFFC0
End Sub

Private Sub TxtAssistant_LostFocus()
TxtAssistant.BackColor = vbWhite
End Sub

Private Sub TxtAssitantPosition_GotFocus()
TxtAssitantPosition.BackColor = &HFFFFC0
End Sub

Private Sub TxtAssitantPosition_LostFocus()
TxtAssitantPosition.BackColor = vbWhite
End Sub

Private Sub TxtContactPerson_GotFocus()
  TxtContactPerson.BackColor = &HFFFFC0
End Sub

Private Sub TxtContactPerson_LostFocus()
TxtContactPerson.BackColor = vbWhite
End Sub

Private Sub txtName_GotFocus()
    txtName.BackColor = &HFFFFC0
End Sub

Private Sub txtName_LostFocus()
    txtName.BackColor = vbWhite
End Sub

Private Sub TxtPosition_GotFocus()
TxtPosition.BackColor = &HFFFFC0
End Sub
Private Sub TxtPosition_LostFocus()
TxtPosition.BackColor = vbWhite
End Sub

Private Sub TxtContactPosition_GotFocus()
TxtContactPosition.BackColor = &HFFFFC0
End Sub

Private Sub TxtContactPosition_LostFocus()
TxtContactPosition.BackColor = vbWhite
End Sub

Private Sub TxtOwner_GotFocus()
TxtOwner.BackColor = &HFFFFC0
End Sub

Private Sub TxtOwner_LostFocus()
TxtOwner.BackColor = vbWhite
End Sub

Private Sub TxtShort_GotFocus()
TxtShort.BackColor = &HFFFFC0
End Sub

Private Sub TxtShort_LostFocus()
TxtShort.BackColor = vbWhite
End Sub

Private Sub TxtTelefax_GotFocus()
TxtTelefax.BackColor = &HFFFFC0
End Sub

Private Sub TxtTelefax_LostFocus()
TxtTelefax.BackColor = vbWhite
End Sub

Private Sub TxtTelno_GotFocus()
TxtTelNo.BackColor = &HFFFFC0
End Sub

Private Sub TxtTelno_LostFocus()
TxtTelNo.BackColor = vbWhite
End Sub
