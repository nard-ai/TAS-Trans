VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHrEOC 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of EOC"
   ClientHeight    =   4305
   ClientLeft      =   4155
   ClientTop       =   3510
   ClientWidth     =   7020
   Icon            =   "frmHrEOC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   7020
   Begin VB.ComboBox cmbClient 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Height          =   315
      ItemData        =   "frmHrEOC.frx":058A
      Left            =   135
      List            =   "frmHrEOC.frx":058C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3870
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   6990
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   285
         Left            =   3060
         TabIndex        =   9
         Top             =   3240
         Width           =   780
      End
      Begin MSComctlLib.ListView lvwdisplay 
         Height          =   2970
         Left            =   135
         TabIndex        =   1
         ToolTipText     =   "double click the mouse to show record of individual"
         Top             =   225
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5239
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SEX"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ADDRESS"
            Object.Width           =   4586
         EndProperty
      End
      Begin VB.Label lblRecount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "count"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   2
         Top             =   3285
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   847
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.Frame Frame2 
         Height          =   555
         Left            =   45
         TabIndex        =   4
         Top             =   -45
         Width           =   6900
         Begin VB.ComboBox cmbmonth 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Height          =   315
            ItemData        =   "frmHrEOC.frx":058E
            Left            =   4590
            List            =   "frmHrEOC.frx":0590
            Style           =   2  'Dropdown List
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   135
            Width           =   1245
         End
         Begin VB.ComboBox cmbYear 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Height          =   315
            ItemData        =   "frmHrEOC.frx":0592
            Left            =   5895
            List            =   "frmHrEOC.frx":0594
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   135
            Width           =   930
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  'Right Justify
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
            Left            =   630
            TabIndex        =   5
            Top             =   180
            Width           =   3840
         End
      End
   End
End
Attribute VB_Name = "frmHrEOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsClient As New ADODB.Recordset
Dim SelectString As String
Dim rs As ADODB.Recordset
Dim listmove As ListItem

Private Function Filter_var() As String

If Me.cmbClient.Text = "--ALL COMPANIES--" Then
    Filter_var = " month([END CONTRACT]) = " & cmbMonth.ListIndex + 1 & " and year([END CONTRACT]) = " & cmbYear.Text & " and employmentstatus = 'C' "
Else
    Filter_var = " clienttable.clientcode = " & Me.cmbClient.ItemData(Me.cmbClient.ListIndex) & " AND month([END CONTRACT]) = " & cmbMonth.ListIndex + 1 & " and year(ENDCONTRACT) = " & cmbYear.Text & " and employmentstatus = 'C' "
End If
       
End Function

Private Sub cmbClient_Click()
Loadlist (Filter_var)
       
End Sub

Private Sub cmbmonth_Click()
'Loadlist (Filter_var)
End Sub

Private Sub cmbYear_Click()
'Loadlist (Filter_var)
End Sub

Private Sub cmdRefresh_Click()
Loadlist (Filter_var)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

    Me.Top = (mdiMain.Height - Me.Height) / 2
    Me.Left = (mdiMain.Width - Me.Width)
 
    lblTitle.Caption = " List of End of Contract for the Month of "
             
'    If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'        StrQuery = "Select * from ClientTable where CLIENTCODE2 = 'FMS' AND ClientStatus = 1"
'    Else
'        StrQuery = "Select * from ClientTable where ClientStatus = 1"
'    End If
    StrQuery = GetClientRecord(0, True)
    Set rstclient = ReturnRst(StrQuery, cnnstr, mstatic)
             
    'Set rstclient = cnn.Execute("Select * from ClientTable where ClientStatus = 1")
        Do While Not rstclient.EOF
            cmbClient.AddItem rstclient!ClientName
            cmbClient.ItemData(cmbClient.NewIndex) = rstclient!ClientCode
            rstclient.MoveNext
        Loop
        rstclient.Close
  
    cmbClient.AddItem "--ALL COMPANIES--", 0
   

    y = 1
     Do While y <= 10
       DoEvents
        'cmbYear.AddItem 2003 + y
        cmbYear.AddItem Format(Now, "yyyy") - 3 + y
        cmbYear.ItemData(cmbYear.NewIndex) = y
       y = y + 1

        Loop
    
    cmbMonth.AddItem "January"
    cmbMonth.AddItem "February"
    cmbMonth.AddItem "March"
       cmbMonth.AddItem "April"
       cmbMonth.AddItem "May"
       cmbMonth.AddItem "June"
       cmbMonth.AddItem "July"
       cmbMonth.AddItem "August"
       cmbMonth.AddItem "September"
       cmbMonth.AddItem "October"
       cmbMonth.AddItem "November"
       cmbMonth.AddItem "December"
    cmbYear.ListIndex = 7
    'cmbYear.ListIndex = Format(Date, "yyyy")
    cmbMonth.ListIndex = Format(Date, "m") - 1
   cmbClient.ListIndex = 0
     '    Loadlist (" month(ENDCONTRACT)=" & Format(Date, "m"))
   '  Call cmbClient_Click
End Sub

Public Sub Loadlist(whereCon As String)
On Error Resume Next

    Me.lvwdisplay.ListItems.Clear
    
    Dim rsEmp1 As ADODB.Recordset
       Set rsEmp1 = New ADODB.Recordset
       With lvwdisplay
       
'       StrQuery = "SELECT [surname] +', ' & + [firstname] + ' ' & IIf(IsNull([middlename],1),'', Left([middlename],1)+'.') AS [Name], HREmployees.EmployeeID, HREmployees.EmpNo as [Control no],ClientTable.ClientName,HREmployees.StartContract as [Start], HREmployees.EndContract as [End],  ClientTable.ClientCode, HREmployees.EmploymentStatus" & _
'                " FROM HREmployees LEFT JOIN ClientTable ON HREmployees.ClientCode = ClientTable.ClientCode WHERE " & whereCon & "  order by endcontract ,surname ,firstname "
        StrQuery = "SELECT * FROM vwEmployeeContract WHERE " & whereCon
        Set rsEmp1 = ReturnRst(StrQuery, cnnstr, mstatic)
       'rsEmp1.Open strQuery, cnnstr, adOpenDynamic, adLockOptimistic

If rsEmp1.EOF = False Then
       PopulateList Me.lvwdisplay, rsEmp1
   '    lvwdisplay.ColumnHeaders(1).Text = "  " & .ListItems.Count & " RECORD/S "
        lvwdisplay.ColumnHeaders(1).Width = 2000
        lvwdisplay.ColumnHeaders(2).Width = 0
        lvwdisplay.ColumnHeaders(3).Width = 1000
        lvwdisplay.ColumnHeaders(4).Width = 1500
        lvwdisplay.ColumnHeaders(5).Width = 980
        lvwdisplay.ColumnHeaders(6).Width = 980
        lvwdisplay.ColumnHeaders(7).Width = 0
        lvwdisplay.ColumnHeaders(8).Width = 0
        lblRecount = .ListItems.Count & " record/s "
        
        X = 1
        Do While .ListItems.Count >= X
             If Val(Format(.ListItems.Item(X).SubItems(5), "d")) <= Val(Format(Date, "d")) Then
            lvwdisplay.ListItems.Item(X).ForeColor = vbRed
            lvwdisplay.ListItems(X).ListSubItems(2).ForeColor = vbRed
            lvwdisplay.ListItems(X).ListSubItems(3).ForeColor = vbRed
            lvwdisplay.ListItems(X).ListSubItems(4).ForeColor = vbRed
            lvwdisplay.ListItems(X).ListSubItems(5).ForeColor = vbRed
    
        End If
        
        X = X + 1
        Loop
        
'        strQuery = "select * from hremployees where employmentstatus <> 'Z' and endcontract between #" & Val(Format(Date, "m")) & "/1/" & Val(Format(Date, "yyyy")) & "# and #" & Format(Date, "mm/dd/yyyy") & "#"
'
'        Dim rsHReoc As New adodb.Recordset
'           Set rsHReoc = ReturnRst(strQuery, Cnnstr, mstatic)
'        If rsHReoc.EOF = False Then
'
'         strQuery = "UPDATE hremployees SET employmentstatus = 'C' WHERE  employmentstatus <> 'Z' and endcontract between #" & Val(Format(Date, "m")) & "/1/" & Val(Format(Date, "yyyy")) & "# and #" & Format(Date, "mm/dd/yyyy") & "#"
'
'         Call Action_Query(strQuery, Cnnstr)
'        End If
Else
    Me.lblRecount.Caption = "0 record"
    Set rsEmp1 = Nothing
End If
        End With
        
End Sub


Private Sub lvwdisplay_DblClick()
    strrow = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).SubItems(1)
      If lvwdisplay.ListItems.Count = 0 Then
        MsgBox "List doesn't have any record from the table selected !", vbOKOnly + vbInformation, "Select..."
    Exit Sub
    Else
    End If
    Call Search(strrow)
    Me.SetFocus
End Sub

Private Sub Search(ByVal strSearch As Integer)
Dim res, bkmrk

On Error Resume Next

With frmHRFinalEmployees
   Set .rsHREmp = New ADODB.Recordset
   StrQuery = "SELECT * FROM HREmployees WHERE agencycode= " & CompanyCode & " order by  surname"
   Set .rsHREmp = ReturnRst(StrQuery, cnnstr, mstatic)
   '.Open "Select * from HREmployees where agencycode= " & CompanyCode & " order by  surname", cnnHR, adOpenKeyset, adLockOptimistic

    bkmrk = .rsHREmp.Bookmark
    .rsHREmp.MoveFirst
    .rsHREmp.Find "EmployeeID=" & Trim(strSearch) & ""
    If Not .rsHREmp.EOF Then
        .Showdata .rsHREmp
    Else
        .rsHREmp.Bookmark = bkmrk
    End If
End With
    
End Sub


Private Sub lvwdisplay_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtSearch = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).Text
End Sub

Private Sub lvwdisplay_KeyDown(KeyCode As Integer, Shift As Integer)

'Call lvwdisplay_DblClick

End Sub

Private Sub lvwdisplay_KeyUp(KeyCode As Integer, Shift As Integer)
'Call lvwdisplay_DblClick

End Sub



