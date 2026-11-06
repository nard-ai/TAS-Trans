VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHRrecontract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recontract"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5700
   Begin VB.CommandButton CmdOk 
      Caption         =   "RECONTRACT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   3240
      TabIndex        =   10
      Top             =   3495
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CANCEL"
      Height          =   325
      Left            =   1560
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame Frame15 
      Height          =   690
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   4935
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
         TabIndex        =   5
         Top             =   270
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   12582912
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
         TabIndex        =   6
         Top             =   270
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   12582912
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
      Begin VB.Label lblClient 
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
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label lblClient 
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
         TabIndex        =   7
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Re-Contract to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   4935
      Begin VB.ComboBox cmbClient 
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
         TabIndex        =   2
         Top             =   585
         Width           =   2685
      End
      Begin VB.ComboBox cmbAgency 
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
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2685
      End
      Begin VB.Label lblClientName2 
         AutoSize        =   -1  'True
         Caption         =   "Client name"
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
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Agency name"
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Label lblClientName 
      Caption         =   "Juan Dela Cruz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label lblAgency 
      Caption         =   "Juan Dela Cruz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblName 
      Caption         =   "Juan Dela Cruz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   240
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5640
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "Client Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Agency Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmHRrecontract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstclient As ADODB.Recordset
Dim rstAgency As ADODB.Recordset

Private Sub cmbAgency_Click()
cmbClient.Clear
cmbClient.ListIndex = -1
    
If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
    StrQuery = "select * from clienttable where CLIENTCODE2 = 'FMS' and agencycode = " & cmbAgency.ListIndex + 1 & " order by clientname"
Else
    StrQuery = "select * from clienttable where agencycode = " & cmbAgency.ListIndex + 1 & " order by clientname"
End If
Set rstclient = ReturnRst(StrQuery, cnnstr, mstatic)
'Set rstclient = cnnHR.Execute("select * from clienttable where agencycode=" & cmbAgency.ListIndex + 1 & " order by clientname")
    
Do While Not rstclient.EOF
    DoEvents
    cmbClient.AddItem rstclient!ClientName
    cmbClient.ItemData(cmbClient.NewIndex) = rstclient!ClientCode
    rstclient.MoveNext
Loop

cmbClient.Enabled = True
lblClientName2.Enabled = True
rstclient.Close
    
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.lblName.Caption = frmHRFinalEmployees.txtsurname.Text & "," & frmHRFinalEmployees.txtfirstname.Text & frmHRFinalEmployees.txtmiddlename.Text
Me.lblAgency = CompanyShort
Me.lblClientName.Caption = frmHRFinalEmployees.cmbCompany.Text


    Set rstAgency = cnnHR.Execute("select agencyshort,agencycode from companytable order by agencycode")
    Do While Not rstAgency.EOF
        DoEvents
        cmbAgency.AddItem rstAgency!AgencyShort
        cmbAgency.ItemData(cmbAgency.NewIndex) = rstAgency!agencycode
        rstAgency.MoveNext
    Loop
    
    rstAgency.Close
    
    
   
    
End Sub

