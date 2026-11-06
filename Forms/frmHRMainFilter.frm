VERSION 5.00
Begin VB.Form frmHRMainFilter 
   Caption         =   "Filter"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbField 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtFilterIndex 
      Height          =   315
      Left            =   840
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2295
   End
   Begin VB.ComboBox cmbOper 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtFilterIndex2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   960
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5520
      Width           =   2295
   End
   Begin VB.ComboBox cmbField2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4560
      Width           =   2295
   End
   Begin VB.ComboBox cmbOper2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5040
      Width           =   975
   End
   Begin VB.OptionButton optAnd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "And"
      Enabled         =   0   'False
      Height          =   255
      Left            =   960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4080
      Width           =   735
   End
   Begin VB.OptionButton optOr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Or"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4080
      Width           =   735
   End
   Begin VB.CheckBox chkConjunction 
      BackColor       =   &H00FFFFFF&
      Caption         =   "More Filter"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ComboBox cmbDept 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Picture         =   "frmHRMainFilter.frx":0000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Picture         =   "frmHRMainFilter.frx":0562
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.OptionButton optHRAll 
      BackColor       =   &H80000009&
      Caption         =   "All"
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
      Height          =   375
      Left            =   360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton optHRFiltered 
      BackColor       =   &H80000009&
      Caption         =   "Filtered"
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
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Field"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Field"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4560
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   3720
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   3720
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "frmHRMainFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkConjunction_Click()
    If chkConjunction.Value = 1 Then
        optAnd.Enabled = True
        optOr.Enabled = True
'        cmbField2.Enabled = False
'        cmbOper2.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    optHRAll.SetFocus
End Sub

Private Sub optHRAll_Click()
    If optHRAll.Value = True Then
        cmbDept.Enabled = False
        cmbField.Enabled = False
        cmbOper.Enabled = False
        txtFilterIndex.Enabled = False
        chkConjunction.Enabled = False
        optAnd.Enabled = False
        optOr.Enabled = False
        cmbField2.Enabled = False
        cmbOper2.Enabled = False
    End If
End Sub

Private Sub optHRFiltered_Click()
 If optHRFiltered.Value = True Then
         cmbDept.Enabled = True
         cmbField.Enabled = True
         cmbOper.Enabled = True
        txtFilterIndex.Enabled = True
        chkConjunction.Enabled = True
'        optAnd.Enabled = False
'        optOr.Enabled = False
'        cmbField2.Enabled = False
'        cmbOper2.Enabled = False
    End If
End Sub
