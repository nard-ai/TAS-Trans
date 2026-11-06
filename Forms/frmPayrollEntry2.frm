VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPayrollEntry2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Payroll Transaction Entries"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPagibigLoan 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   9270
      TabIndex        =   108
      Text            =   "txtPagibigLoan"
      Top             =   5310
      Width           =   765
   End
   Begin VB.TextBox txtSSSLoan 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   6840
      TabIndex        =   107
      Text            =   "txtSSSLoan"
      Top             =   5310
      Width           =   765
   End
   Begin VB.TextBox txtSalAdjustment 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   58
      Text            =   "txtAdjustment"
      Top             =   6390
      Width           =   1215
   End
   Begin VB.TextBox txtPhilHealthYee 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   6840
      TabIndex        =   102
      Text            =   "txtPhilHealth"
      Top             =   5040
      Width           =   765
   End
   Begin VB.TextBox txtPHealthYer 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   9270
      TabIndex        =   101
      Text            =   "txtPHealthYer"
      Top             =   5040
      Width           =   765
   End
   Begin VB.TextBox txtTAX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   6930
      TabIndex        =   100
      Text            =   "txtTax"
      Top             =   5580
      Width           =   765
   End
   Begin VB.TextBox Text9 
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
      Left            =   5310
      TabIndex        =   99
      Text            =   "txtPlacementFee"
      Top             =   6570
      Width           =   765
   End
   Begin VB.TextBox txtEmpyeeSSS 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   6840
      TabIndex        =   93
      Text            =   "txtEmployeeSSS"
      Top             =   4500
      Width           =   765
   End
   Begin VB.TextBox txtEmpyerSSS 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   9270
      TabIndex        =   92
      Text            =   "txtEmpyerSSS"
      Top             =   4500
      Width           =   765
   End
   Begin VB.TextBox txtEmpyeePagibig 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   6840
      TabIndex        =   91
      Text            =   "txtEmpyeePagibig"
      Top             =   4770
      Width           =   765
   End
   Begin VB.TextBox txtEmpyerPagibig 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   9270
      TabIndex        =   90
      Text            =   "txtEmpyerPagibig"
      Top             =   4770
      Width           =   765
   End
   Begin VB.TextBox txtECOLA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   88
      Text            =   "txtEcola"
      Top             =   6660
      Width           =   855
   End
   Begin VB.TextBox txtOverPay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   9270
      TabIndex        =   86
      Text            =   "txtOverPay"
      Top             =   3960
      Width           =   765
   End
   Begin VB.TextBox txtMiscFee 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   6930
      TabIndex        =   84
      Text            =   "txtMiscFee"
      Top             =   3960
      Width           =   765
   End
   Begin VB.TextBox txtPlacementFee 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   9270
      TabIndex        =   82
      Text            =   "txtPlacementFee"
      Top             =   3690
      Width           =   765
   End
   Begin VB.TextBox txtVale 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   6930
      TabIndex        =   80
      Text            =   "txtVale"
      Top             =   3690
      Width           =   765
   End
   Begin VB.TextBox txtBDOCard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   9270
      TabIndex        =   78
      Text            =   "txtBDOCard"
      Top             =   3420
      Width           =   765
   End
   Begin VB.TextBox txtDeductions 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
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
      Left            =   8820
      TabIndex        =   76
      Text            =   "txtDeductions"
      Top             =   2820
      Width           =   1125
   End
   Begin VB.TextBox txtDeductions 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
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
      Left            =   8820
      TabIndex        =   75
      Text            =   "txtDeductions"
      Top             =   2460
      Width           =   1125
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   62
      Text            =   "txtSplHolidayOT"
      Top             =   4950
      Width           =   1215
   End
   Begin VB.TextBox txtLglHolidayOT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   61
      Text            =   "txtLglHolidayOT"
      Top             =   5310
      Width           =   1215
   End
   Begin VB.TextBox txtSickLeave 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   60
      Text            =   "txtSickLeave"
      Top             =   5670
      Width           =   1215
   End
   Begin VB.TextBox txtVacLeave 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   59
      Text            =   "txtVacLeave"
      Top             =   6030
      Width           =   1215
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   11
      Left            =   3840
      TabIndex        =   57
      Text            =   "txtAmount"
      Top             =   6390
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   10
      Left            =   3840
      TabIndex        =   55
      Text            =   "txtAmount"
      Top             =   6030
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   9
      Left            =   3840
      TabIndex        =   53
      Text            =   "txtAmount"
      Top             =   5670
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   8
      Left            =   3840
      TabIndex        =   52
      Text            =   "txtAmount"
      Top             =   5310
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   7
      Left            =   3840
      TabIndex        =   51
      Text            =   "txtAmount"
      Top             =   4950
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   6
      Left            =   3840
      TabIndex        =   49
      Text            =   "txtAmount"
      Top             =   4590
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   5
      Left            =   3840
      TabIndex        =   47
      Text            =   "txtAmount"
      Top             =   4230
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   4
      Left            =   3840
      TabIndex        =   45
      Text            =   "txtAmount"
      Top             =   3870
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   3
      Left            =   3840
      TabIndex        =   43
      Text            =   "txtAmount"
      Top             =   3510
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   2
      Left            =   3840
      TabIndex        =   41
      Text            =   "txtAmount"
      Top             =   3150
      Width           =   1125
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   1
      Left            =   3840
      TabIndex        =   39
      Text            =   "txtAmount"
      Top             =   2790
      Width           =   1125
   End
   Begin VB.TextBox txtCashBond 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   6930
      TabIndex        =   38
      Text            =   "txtCashBond"
      Top             =   3420
      Width           =   765
   End
   Begin VB.TextBox txtHoursLate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   36
      Text            =   "txtHoursLate"
      Top             =   2790
      Width           =   1215
   End
   Begin VB.TextBox txtDaysAbsent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   34
      Text            =   "txtAbsent"
      Top             =   2430
      Width           =   1215
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Index           =   0
      Left            =   3840
      TabIndex        =   31
      Text            =   "txtAmount"
      Top             =   2460
      Width           =   1125
   End
   Begin VB.TextBox txtLglHoliday 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   30
      Text            =   "txtLglHoliday"
      Top             =   4590
      Width           =   1215
   End
   Begin VB.TextBox txtSplHoliday 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   28
      Text            =   "txtSplHoliday"
      Top             =   4230
      Width           =   1215
   End
   Begin VB.TextBox txtSundayOT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   26
      Text            =   "txtSundayOT"
      Top             =   3870
      Width           =   1215
   End
   Begin VB.TextBox txtRestDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   24
      Text            =   "txtRestDay"
      Top             =   3510
      Width           =   1215
   End
   Begin VB.TextBox txtNightDiff 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   23
      Text            =   "txtNightDiff"
      Top             =   3150
      Width           =   1215
   End
   Begin VB.TextBox txtRegOT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   22
      Text            =   "txtRegularOt"
      Top             =   2790
      Width           =   1215
   End
   Begin VB.TextBox txtDaysWork 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   19
      Text            =   "txtDaysWorked"
      Top             =   2430
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   7230
      Width           =   10095
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.CommandButton Command1 
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
         TabIndex        =   13
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
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
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
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
         TabIndex        =   11
         Top             =   1920
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvwEmployeeList 
         Height          =   1515
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2672
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Movement No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Movement Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Movement Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Movement Remarks"
            Object.Width           =   3175
         EndProperty
      End
      Begin MSComctlLib.ListView lvwProcessEmp 
         Height          =   1515
         Left            =   7920
         TabIndex        =   3
         Top             =   390
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2672
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Movement No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Movement Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Movement Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Movement Remarks"
            Object.Width           =   3175
         EndProperty
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
         Height          =   375
         Left            =   180
         TabIndex        =   74
         Top             =   150
         Width           =   2295
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
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   1440
         Width           =   5295
      End
      Begin VB.Label lblEmpName 
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
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   1200
         Width           =   5295
      End
      Begin VB.Label lblEmpNo 
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
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Processed Employees"
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
         Left            =   7920
         TabIndex        =   8
         Top             =   150
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Employees:"
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
         Left            =   7920
         TabIndex        =   7
         Top             =   2010
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Employees:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   2010
         Width           =   1815
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
         Left            =   2340
         TabIndex        =   5
         Top             =   270
         Width           =   5385
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
         Left            =   2400
         TabIndex        =   4
         Top             =   600
         Width           =   5295
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
         TabIndex        =   2
         Top             =   2250
         Width           =   2175
      End
   End
   Begin VB.Label Label10 
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
      Left            =   7110
      TabIndex        =   111
      Top             =   6300
      Width           =   915
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
      TabIndex        =   110
      Top             =   5310
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
      TabIndex        =   109
      Top             =   5310
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
      TabIndex        =   106
      Top             =   5040
      Width           =   1635
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
      Left            =   7110
      TabIndex        =   105
      Top             =   6030
      Width           =   915
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
      TabIndex        =   104
      Top             =   5040
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
      TabIndex        =   103
      Top             =   5580
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
      TabIndex        =   98
      Top             =   4500
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
      TabIndex        =   97
      Top             =   4230
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
      TabIndex        =   96
      Top             =   4500
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
      TabIndex        =   95
      Top             =   4770
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
      TabIndex        =   94
      Top             =   4770
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
      TabIndex        =   89
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
      TabIndex        =   87
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblEntries 
      BackStyle       =   0  'Transparent
      Caption         =   "Miscellaneous Fee:"
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
      TabIndex        =   85
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
      TabIndex        =   83
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
      TabIndex        =   81
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
      TabIndex        =   79
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
      TabIndex        =   77
      Top             =   3150
      Width           =   1815
   End
   Begin VB.Label lblEntries 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Height          =   255
      Index           =   30
      Left            =   3870
      TabIndex        =   73
      Top             =   6660
      Width           =   1035
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
      TabIndex        =   72
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
      TabIndex        =   71
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
      TabIndex        =   70
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
      TabIndex        =   69
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
      TabIndex        =   68
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
      Index           =   28
      Left            =   3360
      TabIndex        =   67
      Top             =   4980
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
      Index           =   27
      Left            =   3360
      TabIndex        =   66
      Top             =   5340
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
      Index           =   26
      Left            =   3360
      TabIndex        =   65
      Top             =   5700
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
      Index           =   25
      Left            =   3360
      TabIndex        =   64
      Top             =   6060
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
      Index           =   24
      Left            =   3360
      TabIndex        =   63
      Top             =   6420
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
      Index           =   22
      Left            =   8310
      TabIndex        =   56
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
      TabIndex        =   54
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
      Height          =   255
      Index           =   18
      Left            =   3360
      TabIndex        =   50
      Top             =   4620
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
      Index           =   17
      Left            =   3360
      TabIndex        =   48
      Top             =   4260
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
      Index           =   16
      Left            =   3360
      TabIndex        =   46
      Top             =   3900
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
      Index           =   15
      Left            =   3360
      TabIndex        =   44
      Top             =   3540
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
      Index           =   14
      Left            =   3360
      TabIndex        =   42
      Top             =   3180
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
      Index           =   13
      Left            =   3360
      TabIndex        =   40
      Top             =   2820
      Width           =   495
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
      TabIndex        =   37
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
      TabIndex        =   35
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
      TabIndex        =   33
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
      Height          =   255
      Index           =   9
      Left            =   3360
      TabIndex        =   32
      Top             =   2460
      Width           =   495
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
      X1              =   5070
      X2              =   5070
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
      TabIndex        =   29
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
      TabIndex        =   27
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
      TabIndex        =   25
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   18
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
      TabIndex        =   15
      Top             =   2460
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
End
Attribute VB_Name = "frmPayrollEntry2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
