VERSION 5.00
Begin VB.Form frmHRSrchNmNo 
   BackColor       =   &H80000009&
   Caption         =   "Search"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3735
   Begin VB.ComboBox cmbempNoNm 
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
      Height          =   345
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "cmbempNoNm"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Search by"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton OptEmpNo 
         BackColor       =   &H80000009&
         Caption         =   "Employee No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton OptName 
         BackColor       =   &H80000009&
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmHRSrchNmNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsEmp As ADODB.Recordset
Dim varEmpNo, varEmpNm As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    If OptEmpNo.Value = True Then
        mempnum = cmbempNoNm.Text
    
        Set rsEmp = cnDCPI.Execute("Select  employeenum,firstname, surname From HREmployees where employeenum='" & mempnum & "'")
            If Not rsEmp.EOF Then
                varEmpNm = rsEmp.Fields(2) & ", " & rsEmp.Fields(1)
            Else
                varEmpNm = ""
               ' cmbempNoNm.ListIndex = -1
            End If
        rsEmp.Close
        
        mfullnm = varEmpNm
    
    ElseIf OptName.Value = True Then
        mfullnm = cmbempNoNm.Text
    
        If cmbempNoNm.Text = "" Or Me.cmbempNoNm.ListIndex = -1 Then
            varEmpNo = ""
        Else
            varEmpNo = Format(Me.cmbempNoNm.ItemData(Me.cmbempNoNm.ListIndex), "000000")
        End If
        mempnum = varEmpNo
    End If
        
    
        varEmpNo = ""
        varEmpNm = ""
    
    Unload Me
End Sub

Private Sub Form_Load()
   Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    
    Set rsEmp = cnDCPI.Execute("Select  employeenum,firstname, surname From HREmployees")
    
    Do While Not rsEmp.EOF
        cmbempNoNm.AddItem rsEmp!employeenum
       
        rsEmp.MoveNext
    Loop
    rsEmp.Close

    
    OptEmpNo.Value = True
    OptName.Value = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsEmp = Nothing
End Sub

Private Sub OptEmpNo_Click()
    If OptEmpNo.Value = True Then
        OptName.Value = False
    End If
    cmbempNoNm.Clear
    rsEmp.Open
    Do While Not rsEmp.EOF
        cmbempNoNm.AddItem rsEmp!employeenum
        rsEmp.MoveNext
    Loop
    rsEmp.Close
    
End Sub

Private Sub OptName_Click()
    If OptName.Value = True Then
        OptEmpNo.Value = False
    End If
    cmbempNoNm.Clear
    rsEmp.Open
    Do While Not rsEmp.EOF
        cmbempNoNm.AddItem rsEmp!surname & ", " & rsEmp!FirstName
        cmbempNoNm.ItemData(cmbempNoNm.NewIndex) = rsEmp!employeenum
        rsEmp.MoveNext
    Loop
    
    rsEmp.Close
End Sub
