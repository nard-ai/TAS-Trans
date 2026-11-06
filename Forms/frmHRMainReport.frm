VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMainReport 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Provincial Address Report"
   ClientHeight    =   8310
   ClientLeft      =   150
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnViewPrint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   9720
      Picture         =   "frmHRMainReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print"
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton btnViewClose 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   10320
      Picture         =   "frmHRMainReport.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   7320
      Width           =   615
   End
   Begin MSComctlLib.ListView HRListView 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11668
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Province"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmHRMainReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim list As ListItem

Private Sub btnViewClose_Click(Index As Integer)
    Unload Me
End Sub

Private Sub btnViewPrint_Click(Index As Integer)
    'frmModuleUnderConstruction.Show vbModal
    If prov = True Then
    'CrystalReport1.ReportFileName = "\\dcpi\development\humanresources\hrReports\hrprovaddress.rpt"
    End If
End Sub

Private Sub btnViewPrintPreview_Click(Index As Integer)
    frmModuleUnderConstruction.Show vbModal
End Sub

Private Sub Form_Load()
    
    Set rs = New ADODB.Recordset
    
    'Provincial Address
    If prov = True Then
        Set rs = cnDCPI.Execute("Select * From HREmployees")
        HRListView.ListItems.Clear
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!province), "", rs!province)
        rs.MoveNext
        Loop
        rs.Close
    'City
    ElseIf city = True Then
        frmHRMainReport.Caption = "City Address Report"
        Set rs = cnDCPI.Execute("Select * From HREmployees")
        HRListView.ListItems.Clear
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            HRListView.ColumnHeaders.Item(3) = "City"
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!city), "", rs!city)
        rs.MoveNext
        Loop
        rs.Close
        
    'Married
    ElseIf married = True Then
        frmHRMainReport.Caption = "Married Employees Report"
        Set rs = cnDCPI.Execute("Select * From HRRptMarried")
        HRListView.ListItems.Clear
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            HRListView.ColumnHeaders.Item(3) = "Civil Status"
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!civilstatusdesc), "", rs!civilstatusdesc)
        rs.MoveNext
        Loop
        rs.Close
        
    'Single
    ElseIf sngle = True Then
        frmHRMainReport.Caption = "Single Employees Report"
        Set rs = cnDCPI.Execute("Select * From HRRptSingle")
        HRListView.ListItems.Clear
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            HRListView.ColumnHeaders.Item(3) = "Civil Status"
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!civilstatusdesc), "", rs!civilstatusdesc)
        rs.MoveNext
        Loop
        rs.Close
    
    'SSS
    ElseIf sss = True Then
        frmHRMainReport.Caption = "Employees with SSS Report"
        Set rs = cnDCPI.Execute("Select * From HREmployees")
        HRListView.ListItems.Clear
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            HRListView.ColumnHeaders.Item(3) = "SSS No"
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!sss), "", rs!sss)
        rs.MoveNext
        Loop
        rs.Close
        
    'TIN
    ElseIf tin = True Then
        frmHRMainReport.Caption = "Employees with TIN Report"
        Set rs = cnDCPI.Execute("Select * From HREmployees")
        HRListView.ListItems.Clear
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            HRListView.ColumnHeaders.Item(3) = "TIN No"
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!taxid), "", rs!taxid)
        rs.MoveNext
        Loop
        rs.Close
    
    'Pag-Ibig
    ElseIf pagibig = True Then
        frmHRMainReport.Caption = "Employees with Pag-Ibig Report"
        Set rs = cnDCPI.Execute("Select * From HREmployees")
        HRListView.ListItems.Clear
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            HRListView.ColumnHeaders.Item(3) = "Pag-Ibig No"
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!pagibig), "", rs!pagibig)
        rs.MoveNext
        Loop
        rs.Close
    
    'PhilHealth
    ElseIf philhealth = True Then
        frmHRMainReport.Caption = "Employees with Phil Health Report"
        Set rs = cnDCPI.Execute("Select * From HREmployees")
        HRListView.ListItems.Clear
        HRListView.ColumnHeaders.Item(3) = "PhilHealth No"
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            HRListView.ColumnHeaders.Item(3) = "PhilHealth No"
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!phcode), "", rs!phcode)
        rs.MoveNext
        Loop
        rs.Close
    
    'Employee No
    ElseIf empnum = True Then
        frmHRMainReport.Caption = "Employees with Employee No Report"
        Set rs = cnDCPI.Execute("Select * From HRRptEmpNo")
        HRListView.ListItems.Clear
        HRListView.ColumnHeaders.Item(3) = "Department"
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!Department), "", rs!Department)
        rs.MoveNext
        Loop
        rs.Close
    
    'Chrono ID
    ElseIf chrono = True Then
        frmHRMainReport.Caption = "Employees with Chrono ID Report"
        Set rs = cnDCPI.Execute("Select * From HRRptChrono")
        HRListView.ListItems.Clear
        HRListView.ColumnHeaders.Item(3) = "Chrono ID"
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!chronoid), "", rs!chronoid)
        rs.MoveNext
        Loop
        rs.Close
    
    'Company ID
    ElseIf coid = True Then
        frmHRMainReport.Caption = "Employees with Company ID Report"
        Set rs = cnDCPI.Execute("Select * From HREmployees")
        HRListView.ListItems.Clear
        HRListView.ColumnHeaders.Item(3) = "Company ID"
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!jobcode), "", rs!jobcode)
        rs.MoveNext
        Loop
        rs.Close
        
    'Male
    ElseIf male = True Then
        frmHRMainReport.Caption = "Male Employees  Report"
        Set rs = cnDCPI.Execute("Select * From HrRptGenderMale")
        HRListView.ListItems.Clear
        HRListView.ColumnHeaders.Item(3) = "Gender"
        HRListView.ColumnHeaders.add (4), "", "Position", 1440, 0, 0
        HRListView.ColumnHeaders.add (5), "", "Employment Status", 1440, 0, 0
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!gender), "", rs!gender)
            list.SubItems(3) = IIf(IsNull(rs!Position), "", rs!Position)
            list.SubItems(4) = IIf(IsNull(rs!employmentstatusdesc), "", rs!employmentstatusdesc)
        rs.MoveNext
        Loop
        rs.Close
    
      'Female
    ElseIf female = True Then
        frmHRMainReport.Caption = "Female Employees  Report"
        Set rs = cnDCPI.Execute("Select * From HRRptGenderFemale")
        HRListView.ListItems.Clear
        HRListView.ColumnHeaders.Item(3) = "Gender"
        HRListView.ColumnHeaders.add (4), "", "Position", 1440, 0, 0
        HRListView.ColumnHeaders.add (5), "", "Employment Status", 1440, 0, 0
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!gender), "", rs!gender)
            list.SubItems(3) = IIf(IsNull(rs!Position), "", rs!Position)
            list.SubItems(4) = IIf(IsNull(rs!employmentstatusdesc), "", rs!employmentstatusdesc)
        rs.MoveNext
        Loop
        rs.Close
    
    'Per Position Per Emp Status
    ElseIf postat = True Then
        frmHRMainReport.Caption = "Employees per position per emploment status  Report"
        Set rs = cnDCPI.Execute("Select * From HRRptEmpPerPos order by Position, EmploymentStatusDesc")
        HRListView.ListItems.Clear
        HRListView.Sorted = False
        HRListView.ColumnHeaders.Item(3) = "Position"
        HRListView.ColumnHeaders.add (4), "", "Employment Status", 1440, 0, 0
        HRListView.ColumnHeaders.add (5), "", "Department", 1440, 0, 0
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!Position), "", rs!Position)
            list.SubItems(3) = IIf(IsNull(rs!employmentstatusdesc), "", rs!employmentstatusdesc)
            list.SubItems(4) = IIf(IsNull(rs!Department), "", rs!Department)
        rs.MoveNext
        Loop
        rs.Close
     
    'Regular
    ElseIf regular = True Then
        frmHRMainReport.Caption = "Regular Employees  Report"
        Set rs = cnDCPI.Execute("Select * From HRRptRegular ")
        HRListView.ListItems.Clear
        HRListView.ColumnHeaders.Item(3) = "Employment Status"
        HRListView.ColumnHeaders.add (4), "", "Department", 1440, 0, 0
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!employmentstatusdesc), "", rs!employmentstatusdesc)
            list.SubItems(3) = IIf(IsNull(rs!Department), "", rs!Department)
        rs.MoveNext
        Loop
        rs.Close
        
    'Contractual
    ElseIf contr = True Then
        frmHRMainReport.Caption = "Contractual Employees  Report"
        Set rs = cnDCPI.Execute("Select * From HRRptContructual order by Department")
        HRListView.ListItems.Clear
        HRListView.Sorted = False
        HRListView.ColumnHeaders.Item(3) = "Employment Status"
        HRListView.ColumnHeaders.add (4), "", "Department", 1440, 0, 0
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!employmentstatusdesc), "", rs!employmentstatusdesc)
            list.SubItems(3) = IIf(IsNull(rs!Department), "", rs!Department)
        rs.MoveNext
        Loop
        rs.Close
                
    'Probationary
    ElseIf proby = True Then
        frmHRMainReport.Caption = "Probationary Employees  Report"
        Set rs = cnDCPI.Execute("Select * From HRRptProbationary order by Department")
        HRListView.ListItems.Clear
        HRListView.Sorted = False
        HRListView.ColumnHeaders.Item(3) = "Employment Status"
        HRListView.ColumnHeaders.add (4), "", "Department", 1440, 0, 0
         Do While Not rs.EOF
            Set list = HRListView.ListItems.add(, , rs!employeenum)
            list.SubItems(1) = IIf(IsNull(rs!surname), "", rs!surname) & ", " & IIf(IsNull(rs!FirstName), "", rs!FirstName)
            list.SubItems(2) = IIf(IsNull(rs!employmentstatusdesc), "", rs!employmentstatusdesc)
            list.SubItems(3) = IIf(IsNull(rs!Department), "", rs!Department)
        rs.MoveNext
        Loop
        rs.Close
    
        
End If

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    prov = False
    city = False
    married = False
    sngle = False
    sss = False
    tin = False
    pagibig = False
    philhealth = False
    emp = False
    regular = False
    male = False
    female = False
    postat = False
    contr = False
End Sub
