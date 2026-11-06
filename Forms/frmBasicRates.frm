VERSION 5.00
Begin VB.Form frmBasicRates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OT and ND Basic Rates Computation"
   ClientHeight    =   7110
   ClientLeft      =   3390
   ClientTop       =   3195
   ClientWidth     =   9570
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmBasicRates.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9570
   Begin VB.CommandButton cmdSaveReg 
      Caption         =   "Save Regular"
      Height          =   375
      Left            =   480
      TabIndex        =   116
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveOT 
      Caption         =   "Save Excess OT"
      Height          =   375
      Left            =   2160
      TabIndex        =   115
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   0
      Left            =   4200
      TabIndex        =   87
      Text            =   "0.00"
      Top             =   120
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   7
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "0.00"
      Top             =   1320
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   8
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "0.00"
      Top             =   1680
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   9
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "0.00"
      Top             =   2040
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   10
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "0.00"
      Top             =   2400
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   11
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "0.00"
      Top             =   2760
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   12
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "0.00"
      Top             =   3120
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   19
      Left            =   8160
      TabIndex        =   57
      Text            =   "0.00"
      Top             =   1320
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   20
      Left            =   8160
      TabIndex        =   58
      Text            =   "0.00"
      Top             =   1680
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   21
      Left            =   8160
      TabIndex        =   59
      Text            =   "0.00"
      Top             =   2040
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   22
      Left            =   8160
      TabIndex        =   60
      Text            =   "0.00"
      Top             =   2400
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   23
      Left            =   8160
      TabIndex        =   61
      Text            =   "0.00"
      Top             =   2760
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   24
      Left            =   8160
      TabIndex        =   62
      Text            =   "0.00"
      Top             =   3120
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   25
      Left            =   3840
      TabIndex        =   39
      Text            =   "1.00"
      Top             =   1320
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   26
      Left            =   3840
      TabIndex        =   40
      Text            =   "1.00"
      Top             =   1680
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   27
      Left            =   3840
      TabIndex        =   41
      Text            =   "1.00"
      Top             =   2040
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   28
      Left            =   3840
      TabIndex        =   42
      Text            =   "1.00"
      Top             =   2400
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   29
      Left            =   3840
      TabIndex        =   43
      Text            =   "1.00"
      Top             =   2760
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   30
      Left            =   3840
      TabIndex        =   44
      Text            =   "1.00"
      Top             =   3120
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   31
      Left            =   7200
      TabIndex        =   51
      Text            =   "0.00"
      Top             =   1320
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   32
      Left            =   7200
      TabIndex        =   52
      Text            =   "0.00"
      Top             =   1680
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   33
      Left            =   7200
      TabIndex        =   53
      Text            =   "0.00"
      Top             =   2040
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   34
      Left            =   7200
      TabIndex        =   54
      Text            =   "0.00"
      Top             =   2400
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   35
      Left            =   7200
      TabIndex        =   55
      Text            =   "0.00"
      Top             =   2760
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   36
      Left            =   7200
      TabIndex        =   56
      Text            =   "0.00"
      Top             =   3120
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   41
      Left            =   8160
      TabIndex        =   81
      Text            =   "0.00"
      Top             =   4320
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   42
      Left            =   8160
      TabIndex        =   82
      Text            =   "0.00"
      Top             =   4680
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   43
      Left            =   8160
      TabIndex        =   83
      Text            =   "0.00"
      Top             =   5040
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   44
      Left            =   8160
      TabIndex        =   84
      Text            =   "0.00"
      Top             =   5400
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   45
      Left            =   8160
      TabIndex        =   85
      Text            =   "0.00"
      Top             =   5760
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   46
      Left            =   8160
      TabIndex        =   86
      Text            =   "0.00"
      Top             =   6120
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   53
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   69
      Text            =   "0.00"
      Top             =   4320
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   54
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   70
      Text            =   "0.00"
      Top             =   4680
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   55
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   71
      Text            =   "0.00"
      Top             =   5040
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   56
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   72
      Text            =   "0.00"
      Top             =   5400
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   57
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   73
      Text            =   "0.00"
      Top             =   5760
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   58
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   74
      Text            =   "0.00"
      Top             =   6120
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   66
      Left            =   3840
      TabIndex        =   63
      Text            =   "0.00"
      Top             =   4320
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   67
      Left            =   3840
      TabIndex        =   64
      Text            =   "0.00"
      Top             =   4680
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   68
      Left            =   3840
      TabIndex        =   65
      Text            =   "0.00"
      Top             =   5040
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   69
      Left            =   3840
      TabIndex        =   66
      Text            =   "0.00"
      Top             =   5400
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   70
      Left            =   3840
      TabIndex        =   67
      Text            =   "0.00"
      Top             =   5760
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   71
      Left            =   3840
      TabIndex        =   68
      Text            =   "0.00"
      Top             =   6120
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   72
      Left            =   7200
      TabIndex        =   75
      Text            =   "0.00"
      Top             =   4320
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   73
      Left            =   7200
      TabIndex        =   76
      Text            =   "0.00"
      Top             =   4680
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   74
      Left            =   7200
      TabIndex        =   77
      Text            =   "0.00"
      Top             =   5040
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   75
      Left            =   7200
      TabIndex        =   78
      Text            =   "0.00"
      Top             =   5400
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   76
      Left            =   7200
      TabIndex        =   79
      Text            =   "0.00"
      Top             =   5760
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   77
      Left            =   7200
      TabIndex        =   80
      Text            =   "0.00"
      Top             =   6120
      Width           =   825
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   38
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "0"
      Top             =   1320
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "0"
      Top             =   1680
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "0"
      Top             =   2040
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "0"
      Top             =   2400
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "0"
      Top             =   2760
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "0"
      Top             =   3120
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "0"
      Top             =   1320
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "0"
      Top             =   1680
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "0"
      Top             =   2040
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "0"
      Top             =   2400
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "0"
      Top             =   2760
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   11
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "0"
      Top             =   3120
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   12
      Left            =   2880
      TabIndex        =   25
      Text            =   "0"
      Top             =   4320
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   13
      Left            =   2880
      TabIndex        =   24
      Text            =   "0"
      Top             =   4680
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   14
      Left            =   2880
      TabIndex        =   23
      Text            =   "0"
      Top             =   5040
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   15
      Left            =   2880
      TabIndex        =   22
      Text            =   "0"
      Top             =   5400
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   16
      Left            =   2880
      TabIndex        =   21
      Text            =   "0"
      Top             =   5760
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   17
      Left            =   2880
      TabIndex        =   20
      Text            =   "0"
      Top             =   6120
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   18
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0"
      Top             =   4320
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   19
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   4680
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   20
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0"
      Top             =   5040
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   21
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   5400
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   22
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0"
      Top             =   5760
      Width           =   705
   End
   Begin VB.TextBox mRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   23
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "0"
      Top             =   6120
      Width           =   705
   End
   Begin VB.TextBox TxtDRate 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2040
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   120
      Width           =   1305
   End
   Begin VB.CommandButton cmdSaveValue 
      Caption         =   "Save ND"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox txtOTHr 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7200
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   6600
      Width           =   825
   End
   Begin VB.TextBox txtOT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3840
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   6600
      Width           =   825
   End
   Begin VB.TextBox txtOTRate 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4800
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   6600
      Width           =   1185
   End
   Begin VB.TextBox txtNDRate 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8160
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   6600
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   81
      Left            =   8040
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   9240
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   80
      Left            =   7080
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   9240
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   79
      Left            =   4680
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   9240
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   78
      Left            =   3720
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   9240
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   40
      Left            =   8160
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   39
      Left            =   7200
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   7680
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   38
      Left            =   4800
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtRates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   37
      Left            =   3840
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   7680
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblclient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Day"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   32
      Left            =   120
      TabIndex        =   114
      Top             =   1440
      Width           =   2685
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   240
      X2              =   9480
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblclient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Work on:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   3
      Left            =   240
      TabIndex        =   113
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label lblclient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rest day Duty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   4
      Left            =   120
      TabIndex        =   112
      Top             =   1800
      Width           =   2685
   End
   Begin VB.Label lblclient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Special Holiday"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   5
      Left            =   120
      TabIndex        =   111
      Top             =   2160
      Width           =   2685
   End
   Begin VB.Label lblclient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Special Holiday on Restday"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   6
      Left            =   120
      TabIndex        =   110
      Top             =   2520
      Width           =   2685
   End
   Begin VB.Label lblclient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Legal Holiday"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   7
      Left            =   120
      TabIndex        =   109
      Top             =   2880
      Width           =   2685
   End
   Begin VB.Label lblclient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Legal Holiday on Restday"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   8
      Left            =   120
      TabIndex        =   108
      Top             =   3240
      Width           =   2685
   End
   Begin VB.Label lblclient 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Percentage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   630
      Index           =   9
      Left            =   2640
      TabIndex        =   107
      Top             =   720
      Width           =   1380
   End
   Begin VB.Label lblclient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate/Hour"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   10
      Left            =   4920
      TabIndex        =   106
      Top             =   960
      Width           =   990
   End
   Begin VB.Label lblclient 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ND Percentage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   630
      Index           =   11
      Left            =   6120
      TabIndex        =   105
      Top             =   720
      Width           =   1140
   End
   Begin VB.Label lblclient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ND Rate/Hour"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   12
      Left            =   8040
      TabIndex        =   104
      Top             =   960
      Width           =   1305
   End
   Begin VB.Label lblclient 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   14
      Left            =   3960
      TabIndex        =   103
      Top             =   960
      Width           =   660
   End
   Begin VB.Label lblclient 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   15
      Left            =   7320
      TabIndex        =   102
      Top             =   960
      Width           =   660
   End
   Begin VB.Label lblclient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OT ND Rate/Hr"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   17
      Left            =   8040
      TabIndex        =   101
      Top             =   3960
      Width           =   1365
   End
   Begin VB.Label lblclient 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OT ND Percentage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   630
      Index           =   18
      Left            =   6120
      TabIndex        =   100
      Top             =   3720
      Width           =   1140
   End
   Begin VB.Label lblclient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OT Rate/Hr"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   19
      Left            =   4920
      TabIndex        =   99
      Top             =   3960
      Width           =   1050
   End
   Begin VB.Label lblclient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Legal Holiday on Restday"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   21
      Left            =   120
      TabIndex        =   98
      Top             =   6240
      Width           =   2685
   End
   Begin VB.Label lblclient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Legal Holiday"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   22
      Left            =   120
      TabIndex        =   97
      Top             =   5880
      Width           =   2685
   End
   Begin VB.Label lblclient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Special Holiday on Restday"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   23
      Left            =   120
      TabIndex        =   96
      Top             =   5520
      Width           =   2685
   End
   Begin VB.Label lblclient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Special Holiday"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   24
      Left            =   120
      TabIndex        =   95
      Top             =   5160
      Width           =   2685
   End
   Begin VB.Label lblclient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rest day Duty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   25
      Left            =   120
      TabIndex        =   94
      Top             =   4800
      Width           =   2685
   End
   Begin VB.Label lblclient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "During Overtime"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   26
      Left            =   240
      TabIndex        =   93
      Top             =   3840
      Width           =   1860
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   120
      X2              =   9480
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lblclient 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Day"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   30
      Left            =   120
      TabIndex        =   92
      Top             =   4440
      Width           =   2685
   End
   Begin VB.Label lblclient 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   31
      Left            =   3960
      TabIndex        =   91
      Top             =   3960
      Width           =   660
   End
   Begin VB.Label lblclient 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   33
      Left            =   7320
      TabIndex        =   90
      Top             =   3960
      Width           =   660
   End
   Begin VB.Label lblclient 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OT Percentage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   20
      Left            =   2760
      TabIndex        =   89
      Top             =   3720
      Width           =   1140
   End
   Begin VB.Label lblclient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate per Day = "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   88
      Top             =   240
      Width           =   1755
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2760
      X2              =   6000
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   6120
      X2              =   9360
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   6120
      X2              =   9360
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2760
      X2              =   6000
      Y1              =   4200
      Y2              =   4200
   End
End
Attribute VB_Name = "frmBasicRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsBasicRates As ADODB.Recordset

Private Sub cmdReturn_Click()
Unload Me
End Sub

Private Sub cmdSaveOT_Click()
'frmPayrollEntry.txtNightDiff = txtOTHr
'frmPayrollEntry.TxtAmount(2) = txtNDRate
'
'frmPayrollEntry.txtRegOT = txtRates(66)
'frmPayrollEntry.TxtAmount(1) = txtRates(53)
'
'frmPayrollEntry.txtSundayOT = txtRates(26)
'frmPayrollEntry.TxtAmount(4) = txtRates(8)
'
'frmPayrollEntry.txtSplHoliday = txtRates(29)
'frmPayrollEntry.TxtAmount(5) = txtRates(11)
'
'frmPayrollEntry.txtLglHoliday = txtRates(27)
'frmPayrollEntry.TxtAmount(6) = txtRates(9)
'
'frmPayrollEntry.txtRDonSPL = txtRates(28)
'frmPayrollEntry.TxtAmount(19) = txtRates(10)
'
'frmPayrollEntry.txtRDonLGL = txtRates(30)
'frmPayrollEntry.TxtAmount(20) = txtRates(12)

' For Excess OT
frmPayrollEntry.txtExcessOTRegDay = txtOT
frmPayrollEntry.TxtAmount(12) = txtOTRate

MsgBox "Excess OT Entry was updated!", vbInformation

End Sub

Private Sub cmdSaveReg_Click()
'frmPayrollEntry.txtNightDiff = txtOTHr
'frmPayrollEntry.TxtAmount(2) = txtNDRate

frmPayrollEntry.txtRegOT = txtRates(66)
frmPayrollEntry.TxtAmount(1) = txtRates(53)

frmPayrollEntry.txtSundayOT = txtRates(26)
frmPayrollEntry.TxtAmount(4) = txtRates(8)

frmPayrollEntry.txtSplHoliday = txtRates(29)
frmPayrollEntry.TxtAmount(5) = txtRates(11)

frmPayrollEntry.txtLglHoliday = txtRates(27)
frmPayrollEntry.TxtAmount(6) = txtRates(9)

frmPayrollEntry.txtRDonSPL = txtRates(28)
frmPayrollEntry.TxtAmount(19) = txtRates(10)

frmPayrollEntry.txtRDonLGL = txtRates(30)
frmPayrollEntry.TxtAmount(20) = txtRates(12)
'' For Excess OT
'frmPayrollEntry.txtExcessOTRegDay = txtOT
'frmPayrollEntry.TxtAmount(12) = txtOTRate

MsgBox "Regular Entry was updated!", vbInformation

End Sub

Private Sub cmdSaveValue_Click()
frmPayrollEntry.txtNightDiff = txtOTHr
frmPayrollEntry.TxtAmount(2) = txtNDRate

'frmPayrollEntry.txtRegOT = txtRates(66)
'frmPayrollEntry.TxtAmount(1) = txtRates(53)
'
'frmPayrollEntry.txtSundayOT = txtRates(26)
'frmPayrollEntry.TxtAmount(4) = txtRates(8)
'
'frmPayrollEntry.txtSplHoliday = txtRates(29)
'frmPayrollEntry.TxtAmount(5) = txtRates(11)
'
'frmPayrollEntry.txtLglHoliday = txtRates(27)
'frmPayrollEntry.TxtAmount(6) = txtRates(9)
'
'frmPayrollEntry.txtRDonSPL = txtRates(28)
'frmPayrollEntry.TxtAmount(19) = txtRates(10)
'
'frmPayrollEntry.txtRDonLGL = txtRates(30)
'frmPayrollEntry.TxtAmount(20) = txtRates(12)
'' For Excess OT
'frmPayrollEntry.txtExcessOTRegDay = txtOT
'frmPayrollEntry.TxtAmount(12) = txtOTRate

MsgBox "ND Entry was updated!", vbInformation

'frmPayrollEntry.txtRegOT = txtRates(66)
'frmPayrollEntry.txtAmount(1) = txtRates(53)
'
'frmPayrollEntry.txtSundayOT = txtRates(67)
'frmPayrollEntry.txtAmount(4) = txtRates(54)
'
'frmPayrollEntry.txtSplHoliday = txtRates(68)
'frmPayrollEntry.txtAmount(5) = txtRates(55)
'
'frmPayrollEntry.txtLglHoliday = txtRates(70)
'frmPayrollEntry.txtAmount(6) = txtRates(57)
'
'frmPayrollEntry.txtRDonSPL = txtRates(69)
'frmPayrollEntry.txtAmount(19) = txtRates(56)
'
'frmPayrollEntry.txtRDonLGL = txtRates(71)
'frmPayrollEntry.txtAmount(20) = txtRates(58)

End Sub

Private Sub Form_Activate()
Dim i As Integer

For i = 7 To 12 ' compute the rates
    txtRates(i) = Format(mRate(i - 7) * (ChkVal(TxtDRate) / 8) * ChkVal(txtRates(18 + i)), "##0.00")
    txtRates(46 + i) = Format(mRate(5 + i) * (ChkVal(TxtDRate) / 8) * ChkVal(txtRates(59 + i)), "##0.00")
Next
If frmBasicRates.Visible = True Then
    txtRates(31).SetFocus
End If
End Sub

Private Sub Form_Load()
Call GetBasicRates
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsBasicRates = Nothing
frmPayrollEntry.txtHoursLate.SetFocus
End Sub

Private Sub GetBasicRates()
    
StrQuery = "Select * From BasicRateTable"
Set rsBasicRates = ReturnRst(StrQuery, cnnstr, mstatic)

If rsBasicRates.RecordCount <> 0 Then
    rsBasicRates.MoveFirst
    ' Display the basic rates.
    Call DisplayRates(rsBasicRates)
Else
    MsgBox "No records found in Basic Rates Table!", vbInformation
End If
End Sub

'Private Sub DisplayRates(ByRef rsBasic As Recordset)
'
'With rsBasic
'    txtRates(1) = !RegRate & "%"
'    txtRates(2) = !RestRate & "%"
'    txtRates(3) = !SplRate & "%"
'    txtRates(4) = !SplRest & "%"
'    txtRates(5) = !LglRate & "%"
'    txtRates(6) = !LglRest & "%"
'    txtRates(13) = !RegRateND & "%"
'    txtRates(14) = !RestRateND & "%"
'    txtRates(15) = !SplRateND & "%"
'    txtRates(16) = !SplRestND & "%"
'    txtRates(17) = !LglRateND & "%"
'    txtRates(18) = !LglRestND & "%"
'    txtRates(59) = !OTReg & "%"
'    txtRates(60) = !OTRest & "%"
'    txtRates(61) = !OTSpl & "%"
'    txtRates(62) = !OTSplRest & "%"
'    txtRates(63) = !OTLgl & "%"
'    txtRates(65) = !OTLglRest & "%"
'    txtRates(52) = !OTRegND & "%"
'    txtRates(51) = !OTRestND & "%"
'    txtRates(50) = !OTSplND & "%"
'    txtRates(49) = !OTSplRestND & "%"
'    txtRates(48) = !OTLglND & "%"
'    txtRates(47) = !OTLglRestND & "%"
'End With
'
'End Sub

Private Sub DisplayRates(ByRef rsBasic As Recordset, Optional recMode As Integer = 0)

Dim i As Integer
Dim mHolder As String, mUpdate As String

mHolder = ""
'recMode = 1 ' Update mode

With rsBasic
    .MoveFirst

    For i = 0 To 23
        If recMode = 1 Then ' for Update
            mHolder = mHolder & .Fields(i).Name & " = " & ChkVal(txtRates(19 + i)) & ", "
        Else ' just display the value
            mRate(i) = Format(.Fields(i + 1).VALUE, "#0.00")
            'txtRates(19 + i) = Format(.Fields(i).VALUE, "#0.00")
        End If
    Next
    If mHolder <> "" Then   ' there is an insert or update mode
        If mID(mHolder, Len(mHolder) - 1, 1) = "," Then
            mHolder = mID(mHolder, 1, Len(Trim(mHolder)) - 1)
        End If
        If recMode = 1 Then ' for update
            mUpdate = "UPDATE BasicRateTable SET " & mHolder
            If Action_Query(mUpdate, cnnstr) = False Then
                MsgBox "Client table was not successfully updated. Please repeat update process.", vbInformation, PayrollAlert
            Else
                Call GetBasicRates
            End If
        End If
    End If
End With
End Sub

Private Sub txtRates_GotFocus(Index As Integer)
    SendKeys "{HOME}+{END}"
    txtRates(Index).BackColor = &HFFFFC0
End Sub

Private Sub txtRates_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strValid As String

    strValid = "0123456789." & Chr(13) & Chr(8)
    If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
    
Private Sub txtRates_LostFocus(Index As Integer)
    'TxtAmount(23 + Index).BackColor = &HFFFFFF
    txtRates(Index).BackColor = &HFFFFFF
    Select Case Index
        Case 25, 26, 27, 28, 29, 30
            txtRates(Index - 18) = Format(mRate(Index - 25) * (ChkVal(TxtDRate) / 8) * ChkVal(txtRates(Index)), "##0.00")
        Case 66, 67, 68, 69, 70, 71
            txtRates(Index - 13) = Format(mRate(Index - 54) * (ChkVal(TxtDRate) / 8) * ChkVal(txtRates(Index)), "##0.00")
            Call AccumulateOT(Index)
        Case 31, 32, 33, 34, 35, 36
            txtRates(Index - 12).Text = Format(NDAmount(txtRates(Index).Text, Index), "###,##0.00")
            'txtRates(Index - 12).Text = Format(txtRates(Index - 12).Text, "##0.00")
        Case 72, 73, 74, 75, 76, 77
            txtRates(Index - 31).Text = Format(NDAmount(txtRates(Index).Text, Index), "###,##0.00")
            'txtRates(Index - 31).Text = Format(txtRates(Index - 31).Text, "##,##0.00")
        'Case Else
    End Select
    Call Accumulate(Index)
    txtRates(Index).Text = Format(txtRates(Index).Text, "##,##0.00")
    'TxtAmount(23 + Index).Text = Format(TxtAmount(23 + Index).Text, "##,##0.00")
End Sub
    
Private Sub Accumulate(ByVal mIndex As Integer)
Dim i As Integer

txtOTHr = 0: txtNDRate = 0
For i = 31 To 77
    If i = 37 Then
        i = i + 35
    End If
    txtOTHr = ChkVal(txtOTHr) + ChkVal(txtRates(i))
Next
For i = 19 To 46
    If i = 25 Then
        i = i + 16
    End If
    txtNDRate = ChkVal(txtNDRate) + ChkVal(txtRates(i))
Next

End Sub
    
Private Sub AccumulateOT(ByVal mIndex As Integer)
Dim i As Integer

txtOT = 0: txtOTRate = 0
For i = 67 To 71
    txtOT = ChkVal(txtOT) + ChkVal(txtRates(i))
Next
For i = 54 To 58
    txtOTRate = ChkVal(txtOTRate) + ChkVal(txtRates(i))
Next

End Sub
    
    
Private Function NDAmount(ByVal mAmt As String, ByVal IndexNo As Integer) As Double
Dim strAmt As String
Dim i As Integer
Dim DeciSw As Boolean
Dim HourlyRate As Double

DeciSw = False

    If IsNull(mAmt) Or mAmt = "" Or mAmt = "0" Or mAmt = "." Then
        mAmt = "0.00"
        NDAmount = 0#
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
        mAmt = Format(strAmt, "0.00")
        
        'Recompute for Daily Rate, Updated Oct. 18, 2014. Bernard
        If TxtDRate > 0 Then
            HourlyRate = TxtDRate / 8
        End If
        NDAmount = mAmt * HourlyRate
    End If
    NDAmount = NewComputedAmount(NDAmount, IndexNo, mAmt)

End Function

Private Function NewComputedAmount(ByVal CAmt As Double, _
    ByVal mKey As Integer, ByVal mNoofHrs As Double) As Double

With rsBasicRates
    Select Case mKey
'        Case 1 ' Get the Regular OT rate
'            NewComputedAmount = !OTReg * CAmt
'        Case 4 ' Get the RestDay Rate
'            NewComputedAmount = !RestRate * CAmt
'        Case 18 ' Restday OT
'            NewComputedAmount = !OTRest * CAmt
'        Case 5 ' Get the SPL holiday rate
'            NewComputedAmount = !SplRate * CAmt
'        Case 12 ' Get the SPL Holiday OT rate
'            NewComputedAmount = !OTSpl * CAmt
'
'        Case 6 ' Get the Legal holiday rate
'            NewComputedAmount = !LglRate * CAmt
'        Case 13 ' get the Legal Holiday OT rate
'            NewComputedAmount = !OTLgl * CAmt
'        Case 19 ' get the RD on SPL Rate
'            NewComputedAmount = !SplRest * CAmt
'        Case 14 ' get the RD on SPL OT Rate
'            NewComputedAmount = !OTSplRest * CAmt
'        Case 20 ' get the RD on SPL Rate
'            NewComputedAmount = !LglRest * CAmt
'        Case 22 ' get the RD on SPL OT Rate
'            NewComputedAmount = !OTLglRest * CAmt

        ' For ND
        Case 31 ' get the Regular ND Rate
            NewComputedAmount = !RegRateND * CAmt
        Case 72 ' get the Regular ND OT Rate
            NewComputedAmount = GetNDRate(66, txtRates(53)) * mNoofHrs
        Case 32 ' get the Restday ND Rate
            NewComputedAmount = GetNDRate(26, txtRates(8)) * mNoofHrs
        Case 73 ' get the Restday ND OT Rate
            NewComputedAmount = GetNDRate(67, txtRates(54)) * mNoofHrs
        Case 33 ' get the SPL ND Rate
            NewComputedAmount = GetNDRate(27, txtRates(9)) * mNoofHrs
        Case 74 ' get the SPL ND OT Rate
            NewComputedAmount = GetNDRate(68, txtRates(55)) * mNoofHrs
        Case 34 ' get the RD on SPL ND Rate
            NewComputedAmount = GetNDRate(28, txtRates(10)) * mNoofHrs
        Case 75 ' get the RD on SPL ND OT Rate
            NewComputedAmount = GetNDRate(69, txtRates(56)) * mNoofHrs
        Case 35 ' get the Lgl ND Rate
            NewComputedAmount = GetNDRate(29, txtRates(11)) * mNoofHrs
        Case 76 ' get the Lgl ND OT Rate
            NewComputedAmount = GetNDRate(70, txtRates(57)) * mNoofHrs
        Case 36 ' get the RD on LGL ND Rate
            NewComputedAmount = GetNDRate(30, txtRates(12)) * mNoofHrs
        Case 77 ' get the RD on LGL ND OT Rate
            NewComputedAmount = GetNDRate(71, txtRates(58)) * mNoofHrs
    End Select
    
End With

End Function

Private Function GetNDRate(ByVal IndexNo As Integer, ByVal RegAmt As TextBox) As Double

On Error GoTo ExitND

'    If IndexNo = 6 Then ' make it double for Legal
'        GetNDRate = ((txtRates(IndexNo) * 2) / ChkVal(NoofHrs)) * 0.1
'    ElseIf IndexNo = 5 Then ' add 100% to total amount
'        GetNDRate = ((txtRates(IndexNo) + TxtDRate) / ChkVal(NoofHrs)) * 0.1
'    Else
        GetNDRate = (ChkVal(RegAmt) / txtRates(IndexNo)) * 0.1
'    End If

    Exit Function

ExitND:

    GetNDRate = 0#

End Function


'    txtRates(21) = !RestRate
'    txtRates(22) = !SplRate
'    txtRates(23) = !SplRest
'    txtRates(24) = !LglRate
'    txtRates(25) = !LglRest
'    txtRates(26) = !RegRateND
'    txtRates(27) = !RestRateND
'    txtRates(28) = !SplRateND
'    txtRates(29) = !SplRestND
'    txtRates(30) = !LglRateND
'    txtRates(31) = !LglRestND
'    txtRates(32) = !OTReg
'    txtRates(33) = !OTRest
'    txtRates(34) = !OTSpl
'    txtRates(35) = !OTSplRest
'    txtRates(36) = !OTLgl
'    txtRates(37) = !OTLglRest
'    txtRates(38) = !OTRegND
'    txtRates(39) = !OTRestND
'    txtRates(40) = !OTSplND
'    txtRates(41) = !OTSplRestND
'    txtRates(42) = !OTLglND
'    txtRates(43) = !OTLglRestND

'End Function

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

