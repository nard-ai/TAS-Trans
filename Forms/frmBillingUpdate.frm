VERSION 5.00
Begin VB.Form frmBillingUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Record"
   ClientHeight    =   5160
   ClientLeft      =   3540
   ClientTop       =   2925
   ClientWidth     =   8550
   Icon            =   "frmBillingUpdate.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8550
   Begin VB.TextBox txtClientCode 
      Height          =   285
      Left            =   6240
      TabIndex        =   62
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Client Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Width           =   8295
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   28
         Left            =   1320
         TabIndex        =   63
         Top             =   1560
         Width           =   3600
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update Client Settings"
         Height          =   375
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Update Client Settings"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   27
         Left            =   7500
         TabIndex        =   29
         Text            =   "0.10"
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   26
         Left            =   7500
         TabIndex        =   30
         Text            =   "15.00"
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   25
         Left            =   7500
         TabIndex        =   31
         Text            =   "0.12"
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   24
         Left            =   5580
         TabIndex        =   26
         Text            =   "1.95"
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   23
         Left            =   5580
         TabIndex        =   27
         Text            =   "2.60"
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   22
         Left            =   5580
         TabIndex        =   28
         Text            =   "3.38"
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   21
         Left            =   3420
         TabIndex        =   23
         Text            =   "1.0"
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   20
         Left            =   3420
         TabIndex        =   24
         Text            =   "1.60"
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   19
         Left            =   3420
         TabIndex        =   25
         Text            =   "1.69"
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   18
         Left            =   1380
         TabIndex        =   20
         Text            =   "1.25"
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   17
         Left            =   1380
         TabIndex        =   21
         Text            =   "0.30"
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   16
         Left            =   1380
         TabIndex        =   22
         Text            =   "0.50"
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Billing Template:"
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
         Index           =   29
         Left            =   120
         TabIndex        =   64
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Night Diff Rate:"
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
         Index           =   28
         Left            =   6240
         TabIndex        =   61
         Top             =   390
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   27
         Left            =   6240
         TabIndex        =   60
         Top             =   750
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "VAT Rate:"
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
         Index           =   26
         Left            =   6240
         TabIndex        =   59
         Top             =   1110
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Excess Day Off Spl:"
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
         Index           =   25
         Left            =   4080
         TabIndex        =   58
         Top             =   390
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Excess Legal Hol:"
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
         Index           =   24
         Left            =   4080
         TabIndex        =   57
         Top             =   750
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Excess Day Off Lgl:"
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
         Index           =   23
         Left            =   4080
         TabIndex        =   56
         Top             =   1110
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Legal w/ Premium:"
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
         Index           =   22
         Left            =   2040
         TabIndex        =   55
         Top             =   390
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Off Legal:"
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
         Index           =   21
         Left            =   2040
         TabIndex        =   54
         Top             =   750
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Excess Day Off:"
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
         Index           =   20
         Left            =   2040
         TabIndex        =   53
         Top             =   1110
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reg. OT Rate:"
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
         Index           =   19
         Left            =   120
         TabIndex        =   52
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Off OT Rate:"
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
         Index           =   18
         Left            =   120
         TabIndex        =   51
         Top             =   750
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Off Special:"
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
         Index           =   17
         Left            =   120
         TabIndex        =   50
         Top             =   1110
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H000000FF&
      Height          =   2805
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   8325
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Save Changes"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cancel Changes"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   15
         Left            =   4320
         TabIndex        =   12
         Top             =   1800
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   14
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   13
         Left            =   6900
         TabIndex        =   16
         Top             =   1800
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   12
         Left            =   6900
         TabIndex        =   15
         Top             =   1440
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   11
         Left            =   6900
         TabIndex        =   14
         Top             =   1080
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   10
         Left            =   4320
         TabIndex        =   10
         Top             =   1080
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   9
         Left            =   4320
         TabIndex        =   11
         Top             =   1440
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   8
         Left            =   4320
         TabIndex        =   9
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   7
         Left            =   7320
         TabIndex        =   4
         Top             =   240
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   6
         Left            =   6900
         TabIndex        =   13
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   5
         Left            =   5280
         TabIndex        =   3
         Top             =   240
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   4
         Left            =   1620
         TabIndex        =   8
         Top             =   1800
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   3
         Left            =   1620
         TabIndex        =   7
         Top             =   1440
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   2
         Left            =   1620
         TabIndex        =   6
         Top             =   1080
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   1
         Left            =   1620
         TabIndex        =   5
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox txtEditFld 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   1380
         TabIndex        =   1
         Top             =   240
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8160
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Excess Day Off Lgl:"
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
         Index           =   16
         Left            =   2760
         TabIndex        =   48
         Top             =   1830
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   2280
         TabIndex        =   47
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Incentive:"
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
         Index           =   13
         Left            =   5520
         TabIndex        =   45
         Top             =   1830
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Adjust:"
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
         Index           =   12
         Left            =   5520
         TabIndex        =   44
         Top             =   1470
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Night Diff OT:"
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
         Index           =   11
         Left            =   5520
         TabIndex        =   43
         Top             =   1110
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Excess Day Off Spl:"
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
         Index           =   10
         Left            =   2760
         TabIndex        =   42
         Top             =   1110
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Excess Legal Hol:"
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
         Index           =   9
         Left            =   2760
         TabIndex        =   41
         Top             =   1470
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Excess Day Off:"
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
         Index           =   8
         Left            =   2760
         TabIndex        =   40
         Top             =   750
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reg. OT Hours:"
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
         Index           =   7
         Left            =   6120
         TabIndex        =   39
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tardi Hours:"
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
         Index           =   6
         Left            =   5520
         TabIndex        =   38
         Top             =   750
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Days Worked:"
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
         Index           =   5
         Left            =   4080
         TabIndex        =   37
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Off Legal:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   36
         Top             =   1830
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Legal w/ Premium:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   1470
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Off Spl:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   1110
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Day Off OT:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   750
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Days Present:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   270
         Width           =   1065
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   46
      Top             =   2040
      Visible         =   0   'False
      Width           =   4785
   End
End
Attribute VB_Name = "frmBillingUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    frmBilling.chkSuccess.VALUE = 0
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim itmFound As ListItem   ' FoundItem variable.
Dim mEmpNo As String
Dim mSOANo As String
Dim mClientCode As Integer

mEmpNo = frmBilling.txtEmpNo
mSOANo = frmBilling.txtSOANo
mClientCode = frmBilling.cmbClient.ItemData(frmBilling.cmbClient.ListIndex)

'Set itmFound = frmBilling.ListView1(0). _
'    FindItem(mEmpno, lvwText, , 1)
    'FindItem(mEmpno, lvwText, , lvwPartial)
   
   ' If no ListItem is found, then inform user and exit. If a
   ' ListItem is found, scroll the control using the EnsureVisible
   ' method, and select the ListItem.
'   If itmFound Is Nothing Then  ' If no match, inform user and exit.
'   Else
       'itmFound.EnsureVisible ' Scroll ListView to show found ListItem.
       'itmFound.SELECTED = True   ' Select the ListItem.
       'Return focus to the control to see selection.
       'frmListView1(0).SetFocus
       
       If UpdateBilledTransTable(mEmpNo, mSOANo, mClientCode) = True Then
            With frmBilling.ListView1(0)
                .SetFocus
                .SelectedItem.ListSubItems(5).Text = Format(txtEditFld(0), "0.00")                      ' days present
                
                .SelectedItem.ListSubItems(7).Text = Format(txtEditFld(1), "0.00")                      ' rest day hours
                .SelectedItem.ListSubItems(9).Text = Format(txtEditFld(2), "0.00")                      ' rest day spl holiday
                .SelectedItem.ListSubItems(11).Text = Format(txtEditFld(3), "0.00")                     ' legal with premium
                .SelectedItem.ListSubItems(13).Text = Format(txtEditFld(4), "0.00")                     ' rest day legal holiday
                
                .SelectedItem.ListSubItems(15).Text = Format(txtEditFld(5), "0.00")                     ' days worked
                
                .SelectedItem.ListSubItems(21).Text = Format(txtEditFld(8), "0.00")     ' excess day off
                .SelectedItem.ListSubItems(22).Text = Format(txtEditFld(10), "0.00")    ' excess day off spl
                .SelectedItem.ListSubItems(24).Text = Format(txtEditFld(9), "0.00")     ' excess Legal with premium
                .SelectedItem.ListSubItems(26).Text = Format(txtEditFld(15), "0.00")    ' Excess legal holiday
                
                .SelectedItem.ListSubItems(28).Text = Format(txtEditFld(11), "0.00")    ' Night Diff Ot
                
                .SelectedItem.ListSubItems(17).Text = Format(txtEditFld(6), "0.00")     ' Tardi Hours
                .SelectedItem.ListSubItems(19).Text = Format(txtEditFld(7), "0.00")     ' Ot Hours
                
                .SelectedItem.ListSubItems(29).Text = Format(txtEditFld(14), "0.00")    ' Daily Rate
                
                .SelectedItem.ListSubItems(44).Text = Format(txtEditFld(12), "0.00")    ' Salary Adjust
                .SelectedItem.ListSubItems(55).Text = Format(txtEditFld(13), "0.00")    ' Incentive
                
            End With
            frmBilling.chkSuccess.VALUE = 1
            MsgBox "Changes successfully saved to " & frmBilling.ListView1(0).SelectedItem.ListSubItems(2).Text, vbInformation, PayrollAlert
       Else
            frmBilling.chkSuccess.VALUE = 0
            MsgBox "Cannot save changes, please repeat editing.", vbInformation, PayrollAlert
       End If
'   End If
    Unload Me

End Sub

Private Function UpdateBilledTransTable(ByVal mEmpNo As String, ByVal mSOANo As String, _
    ByVal mClientCode As Integer) As Boolean

StrQuery = "UPDATE BilledTrans SET DaysPresent2 = " & ChkVal(txtEditFld(0)) & ", TaxRefundHours2 = " & ChkVal(txtEditFld(1)) & ", DailyRate2 = " & ChkVal(txtEditFld(14)) & _
        ", NoRegHoliday2 = " & ChkVal(txtEditFld(2)) & ", SundayOT2 = " & ChkVal(txtEditFld(3)) & ", NoSplHoliday2 = " & ChkVal(txtEditFld(4)) & _
        ", HoursTardiness2 = " & ChkVal(txtEditFld(6)) & ", OTHours2 = " & ChkVal(txtEditFld(7)) & ", Excesssplholothours2 = " & ChkVal(txtEditFld(8)) & _
        ", LglHolidayOTHours2 = " & ChkVal(txtEditFld(9)) & ", SplHolidayPay2 = " & ChkVal(txtEditFld(10)) & ", ExcessOTSunLegalHrs2 = " & ChkVal(txtEditFld(15)) & ", NightDiffOT2 = " & ChkVal(txtEditFld(11)) & _
        ", OtherIncome2 = " & ChkVal(txtEditFld(12)) & ", Allowance2 = " & ChkVal(txtEditFld(13)) & _
        " WHERE EmpNo = '" & mEmpNo & "' and SOANo = '" & mSOANo & "' and Clientcode = " & mClientCode
If Action_Query(StrQuery, cnnstr) = True Then
    UpdateBilledTransTable = True
Else
    UpdateBilledTransTable = False
End If

End Function

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

Private Sub cmdUpdate_Click()
Dim rsClientUpdate As ADODB.Recordset
Dim ans As Variant
Dim ExcelTemplate As String
'Dim FSO As New FileSystemObject

ExcelTemplate = App.Path & "\Billing Templates\" & Replace(txtEditFld(28), "'", "''")
'Set FS0 = New FileSystemObject
If IsFileExists(ExcelTemplate) = False Then
    MsgBox "Billing template " & ExcelTemplate & " does not exists!" & Chr(13) & _
        "Please check it first, cannot update.", vbInformation, PayrollAlert
    GoTo ExitRtn1
End If

'Save current client code.
txtClientCode = frmBilling.cmbClient.ItemData(frmBilling.cmbClient.ListIndex)

ans = MsgBox("Do you want to update the Client default Rates?", vbYesNo, PayrollAlert)
If ans = vbNo Then
    GoTo ExitRtn1
End If

'        strQuery = "UPDATE ClientTable SET OverTimeRate = " & ChkVal(txtEditFld(18)) & ", RestDayRate = " & ChkVal(txtEditFld(17)) & _
'        ", SpecialHolidayRate = " & ChkVal(txtEditFld(16)) & ", OTwoPremiumRate = " & ChkVal(txtEditFld(21)) & ", LegalHolidayOTRate = " & ChkVal(txtEditFld(20)) & _
'        ", ExcessOTSplSunRate = " & ChkVal(txtEditFld(19)) & ", LegalHolidayRate = " & ChkVal(txtEditFld(24)) & ", ExcessOTLegalRate = " & ChkVal(txtEditFld(23)) & _
'        ", ExcessOTSunLglRate = " & ChkVal(txtEditFld(22)) & ", NDiffRate = " & ChkVal(txtEditFld(27)) & ", ClientAgencyRate = " & ChkVal(txtEditFld(26)) & _
'        ", VatPercentRate = " & ChkVal(txtEditFld(25)) & ", BillingTemplate = '" & Replace(txtEditFld(28), "'", "''") & "' WHERE ClientCode = " & Val(frmBilling.txtClientCode.Text)
        
        StrQuery = "UPDATE ClientTable SET ClientAgencyRate = " & ChkVal(txtEditFld(26)) & _
        ", VatPercentRate = " & ChkVal(txtEditFld(25)) & ", BillingTemplate = '" & Replace(txtEditFld(28), "'", "''") & "' WHERE ClientCode = " & Val(frmBilling.txtClientCode.Text)


'        RestDayWorkedOTRate = Format(!RestDayRate, "0.00")  ' 0.3
'        SpecialWorkedDayOffRate = Format(!SpecialHolidayRate, "0.00")  '0.5
'        LegalHolidayWithPremiumRate = Format(!OTwoPremiumRate, "0.00") '1#
'        ExcessDayoffSplOTRate = Format(!LegalHolidayRate, "0.00") '1.95
'        LegalHolidayWorkedDayoffRate = Format(!LegalHolidayOTRate, "0.00") '1.6
'        WorkHrsPerDay = !WorkHrsPerDay '8
'
'        OverTimeRate = Format(!OverTimeRate, "0.00") '1.25
'
'        ExcessSplOTRate = Format(!ExcessOTSplSunRate, "0.00") '1.69
'        ExcessDayoffSplOTRate = Format(!LegalHolidayRate, "0.00") '1.95
'        ExcessLglHolOTRate = Format(!ExcessOTLegalRate, "0.00")  '2.6
'        ExcessSunLglOTRate = Format(!ExcessOTSunLglRate, "0.00")   '3.38
'        NightDiffRate = Format(!NDiffRate, "0.00")  '0.1
'
'        AgencyRate = Format(!ClientAgencyRate, "0.00") ' 15
'        VatPercentRate = Format(!VatPercentRate, "0.00")  '0.12

If Action_Query(StrQuery, cnnstr) = True Then
    frmBilling.chkSuccess.VALUE = 0
    StrQuery = "SELECT a.* FROM clientTable a WHERE a.clientcode = " & Val(frmBilling.txtClientCode.Text)
    Set rsClientUpdate = ReturnRst(StrQuery, cnnstr, mstatic)
    With rsClientUpdate
        RestDayWorkedOTRate = Format(!RestDayRate, "0.00")  ' 0.3
        SpecialWorkedDayOffRate = Format(!SpecialHolidayRate, "0.00")  '0.5
        LegalHolidayWithPremiumRate = Format(!OTwoPremiumRate, "0.00") '1#
        LegalHolidayWorkedDayoffRate = Format(!legalholidayotrate, "0.00") '1.6
        WorkHrsPerDay = !WorkHrsPerDay '8
        
        OverTimeRate = Format(!OverTimeRate, "0.00") '1.25
        ExcessSplOTRate = Format(!ExcessOTSplSunRate, "0.00") '1.69
        ExcessDayoffSplOTRate = Format(!LegalHolidayRate, "0.00") '1.95
        ExcessLglHolOTRate = Format(!ExcessOTLegalRate, "0.00")  '2.6
        ExcessSunLglOTRate = Format(!ExcessOTSunLglRate, "0.00")   '3.38
        NightDiffRate = Format(!NDiffRate, "0.00")  '0.1
        
        AgencyRate = Format(!ClientAgencyRate, "0.00") ' 15
        VatPercentRate = Format(!VatPercentRate, "0.00")  '0.12
        BillingTemplate = IIf(IsNull(!BillingTemplate) Or !BillingTemplate = "", "Other Clients SOA Template.xls", !BillingTemplate)
    
    End With
    MsgBox "Client settings was successfully updated.", vbInformation, PayrollAlert
Else
    frmBilling.chkSuccess.VALUE = 0
    MsgBox "Failed to update Client settings.", vbInformation, PayrollAlert
End If

ExitRtn1:
    Unload Me
    
End Sub

Private Sub Form_Load()

    Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2

Frame2.Caption = frmBilling.cmbClient.Text

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub txtEditFld_GotFocus(Index As Integer)
txtEditFld(Index).BackColor = &HFFFFC0
'Set oCommonFunction.ValidAlpha = txtEditFld(Index)
Call txtHLigths(txtEditFld(Index))

End Sub

Private Sub txtEditFld_LostFocus(Index As Integer)
txtEditFld(Index).BackColor = vbWhite
End Sub
