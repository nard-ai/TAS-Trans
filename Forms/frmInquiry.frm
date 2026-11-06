VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInquiry 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EMPLOYEE INQUIRY"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "frmInquiry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmInquiry.frx":08CA
   ScaleHeight     =   4860
   ScaleWidth      =   8070
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3375
      TabIndex        =   10
      Top             =   225
      Width           =   4515
   End
   Begin VB.ComboBox cboSelect 
      Height          =   315
      ItemData        =   "frmInquiry.frx":16020C
      Left            =   1200
      List            =   "frmInquiry.frx":160216
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   6
      Top             =   4320
      Width           =   1305
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   5
      Top             =   4320
      Width           =   1305
   End
   Begin VB.TextBox txtEmpID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   6840
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker dtRanges 
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   705
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   47644675
      CurrentDate     =   38148
   End
   Begin MSComCtl2.DTPicker dtRanges 
      Height          =   315
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   705
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   47644675
      CurrentDate     =   38163
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5106
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH FOR"
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
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TO :"
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
      Left            =   2760
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FROM :"
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
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frmInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mycls As clsEmpRS
Dim mSelect As Boolean
Private WithEvents rsList As ADODB.Recordset
Attribute rsList.VB_VarHelpID = -1

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cboSelect_Click()
    Me.txtEmpID = ""
    Me.txtSearch = ""
    If cboSelect.ListIndex = 1 Then
      Call loadlv(2, 2800, 700)
    Else
      Call loadlv(1, 700, 2800)
    End If
End Sub

Private Sub cmdBtn_Click(Index As Integer)
  ' Status.Show
    Select Case Index
        
        Case 0
            'If txtEmpID = "" Then Unload Status: MsgBox "Search data is required.", vbInformation: Exit Sub
            If Me.txtSearch <> Me.ListView1.SelectedItem.Text Then MsgBox "Employee " & cboSelect.Text & " [" & Me.txtSearch.Text & "] not found!", vbInformation: GoTo exitFunc
            Set mycls = New clsEmpRS
            Call mycls.Disply_RS_Employees(mEmpInquiry, dtRanges(0).Value, dtRanges(1).Value, , , " time01.empn=" & Val(txtEmpID.Text))
            Set mycls = Nothing
            Dim rsInquiry As New ADODB.Recordset
            Set rsInquiry = ReturnRst(SQLstring, cnnstr_mdb, mstatic)
            If rsInquiry.EOF Then Call MsgNoRecords(frmInquiry): Exit Sub
            Set rsInquiry = Nothing

            Set rsList = ReturnRst("Select EMPNAME,empid,empline,empfacility,empdept,emphired from Employee_List order by empname", cnnstr_mdb, mstatic)
               
            With crystlInqry
                .Reset
                .ReportFileName = "C:\Projects\HRProjects\Reports\Employee_Inquiry.rpt"
                .SelectionFormula = "{Emp_Inquiry.EmpID} = " & txtEmpID.Text & ""
                .Formulas(1) = "field_from = '" & (Format(dtRanges(0).Value, "MM/dd/yyyy")) & "'"
                .Formulas(2) = "field_To = '" & (Format(dtRanges(1).Value, "MM/dd/yyyy")) & "'"
                .WindowShowExportBtn = True
                .WindowShowRefreshBtn = True
                .WindowState = crptMaximized
                .WindowTitle = "HR Employee Inquiry"
                .WindowShowPrintBtn = True
                .WindowShowPrintSetupBtn = True
                .Action = 1
            End With
            
            Dim X As Integer
            X = 1
            While X <= 5
            ListView1.SelectedItem.ForeColor = vbRed
            ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(X).ForeColor = vbRed
            X = X + 1
            Wend
        Case 1
            Unload Me
            
    End Select
exitFunc:
    ' Unload Status
End Sub
Private Function disply_Name(ByVal mSelect As Integer) As String
    Dim rsname As New ADODB.Recordset
    'Set rsname = ReturnRst("Select EmpName from Employee_List Where EmpID like '" & Trim(mEmp_ID) & "'", cnnstr_mdb, mstatic)
    Set rsname = ReturnRst("Select EmpName from Employee_List", cnnstr_mdb, mstatic)
    If Not rsname.EOF Then
        disply_Name = IIf(IsNull(rsname.Fields(0)), "", rsname.Fields(0))
    Else
         PopulateList ListView1, rsname
    End If
    rsname.Close
    Set rsname = Nothing
End Function

Private Sub Form_Load()
    Call Setdate(frmInquiry)
    Call loadlv(1, 700, 2800)
    cboSelect.ListIndex = 0
End Sub
Private Sub loadlv(ByVal mSelect As Integer, ByVal column1 As Integer, ByVal column2 As Integer)
    Dim rsList As New ADODB.Recordset
    ListView1.ListItems.Clear
    ListView1.Sorted = False
    If mSelect = 1 Then
        Set rsList = ReturnRst("Select EmpID as [ID] ,EmpName as [NAME],empline as [PROJECT],empfacility as [FACILITY],empdept as [DEPT],emphired   as [DATE HIRED] from Employee_List order by [EMPID]", cnnstr_mdb, mstatic)
    Else
        Set rsList = ReturnRst("Select EmpName as [NAME],EmpID as [ID] ,empline as [PROJECT],empfacility as [FACILITY],empdept as [DEPT],emphired   as [DATE HIRED] from Employee_List order by [EMPNAME]", cnnstr_mdb, mstatic)
    End If
    PopulateList ListView1, rsList
    ListView1.ColumnHeaders(1).Width = column1
    ListView1.ColumnHeaders(2).Width = column2
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.Sorted = True
    If ColumnHeader.Index = 1 Then
        Call loadlv(1, 700, 2800)
    Else
        ListView1.SortKey = ColumnHeader.Index - 1
    End If
    ListView1.Refresh
End Sub

Private Sub ListView1_DblClick()
    Call cmdBtn_Click(0)
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtSearch = ListView1.SelectedItem.Text
End Sub

Private Sub txtsearch_Change()
    Autocomplete ListView1, (Trim$(Me.txtSearch)), Me.txtSearch
    If Me.cboSelect.ListIndex = 0 Then
        Me.txtEmpID = txtSearch
    Else
        txtEmpID = ListView1.SelectedItem.ListSubItems(1).Text
    End If
End Sub

Private Sub txtsearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Execute = True
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
    Execute = False
    End If
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then cmdBtn_Click (0)
End Sub

