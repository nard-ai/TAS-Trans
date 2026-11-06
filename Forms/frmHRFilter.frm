VERSION 5.00
Begin VB.Form frmHRFilter 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HR - Filter Employees"
   ClientHeight    =   2130
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4830
   Icon            =   "frmHRFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4830
   Begin VB.OptionButton optHRFiltered 
      BackColor       =   &H80000004&
      Caption         =   "Filter"
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
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   135
      Width           =   1095
   End
   Begin VB.OptionButton optHRAll 
      BackColor       =   &H80000004&
      Caption         =   "All"
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
      Height          =   255
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2790
      Picture         =   "frmHRFilter.frx":058A
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1710
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3795
      Picture         =   "frmHRFilter.frx":0AEC
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1665
      Width           =   855
   End
   Begin VB.ComboBox cmbClient 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Height          =   330
      ItemData        =   "frmHRFilter.frx":104E
      Left            =   765
      List            =   "frmHRFilter.frx":1050
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1035
      Width           =   3450
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   0
      X2              =   5400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "By Company"
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
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmHRFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim rstclient As ADODB.Recordset
Public Filter_var As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If optHRFiltered.VALUE = True Then
        If cmbClient.Text = "" Then
            MsgBox "Select client to filter.", vbInformation
            cmbClient.SetFocus
            Exit Sub
        Else
            HRFilter
            Unload Me
            frmHRMainEmployee.Show
            frmHRMainEmployee.SetFocus
        End If
    End If
    If optHRAll.VALUE = True Then
        Me.Filter_var = ""
        Me.Hide
        frmHRMainEmployee.lblFiltEmp.Visible = False
        frmHRMainEmployee.Show
        frmHRMainEmployee.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\pics\bg.gif")
    Top = 0
    Left = 0
    
'    If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'        StrQuery = "Select * from clienttable where CLIENTCODE2 = 'FMS'"
'    Else
'        StrQuery = "Select * from clienttable"
'    End If
    StrQuery = GetClientRecord(0)
    Set rstclient = ReturnRst(StrQuery, cnnstr, mstatic)
    
    'Set rstclient = cnn.Execute("Select * from clienttable")
    Do While Not rstclient.EOF
        DoEvents
        cmbClient.AddItem rstclient!ClientName
        cmbClient.ItemData(cmbClient.NewIndex) = rstclient!ClientCode
        rstclient.MoveNext
        Loop
        rstclient.Close
End Sub


Private Sub optHRAll_Click()
    If optHRAll.VALUE = True Then
        cmbClient.Enabled = False
    Else
        cmbClient.Enabled = True
    End If
End Sub

Private Sub HRFilter()
    Dim str As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    If Me.optHRAll.VALUE Then
        Me.Filter_var = ""
    Else
        If Me.cmbClient.Text = "*All" Then
        Else
        Filter_var = "clientcode=" & Me.cmbClient.ItemData(Me.cmbClient.ListIndex)
        End If
    End If
    'optAll = False
End Sub

Private Sub optHRFiltered_Click()
    If optHRFiltered.VALUE = True Then
        cmbClient.Enabled = True
     End If
End Sub

