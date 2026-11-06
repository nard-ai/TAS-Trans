VERSION 5.00
Begin VB.Form FrmReportHired 
   Caption         =   "Report Manager Maintenance"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   Icon            =   "frmReportCustomized.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Sort Options"
      Height          =   2235
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox cboSort 
         Height          =   315
         Index           =   3
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1740
         Width           =   1755
      End
      Begin VB.ComboBox cboSort 
         Height          =   315
         Index           =   2
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1320
         Width           =   1755
      End
      Begin VB.ComboBox cboSort 
         Height          =   315
         Index           =   1
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   900
         Width           =   1755
      End
      Begin VB.ComboBox cboSort 
         Height          =   315
         Index           =   0
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1755
      End
      Begin VB.CheckBox chkSort 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   3420
         TabIndex        =   6
         Top             =   1770
         Width           =   1275
      End
      Begin VB.CheckBox chkSort 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   3420
         TabIndex        =   5
         Top             =   1350
         Width           =   1275
      End
      Begin VB.CheckBox chkSort 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   3420
         TabIndex        =   4
         Top             =   930
         Width           =   1275
      End
      Begin VB.CheckBox chkSort 
         Caption         =   "Descending"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   3420
         TabIndex        =   3
         Top             =   510
         Width           =   1275
      End
      Begin VB.Label lblSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F&ourth Sort Field:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label lblSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Third Sort Field:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   1380
         Width           =   1110
      End
      Begin VB.Label lblSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Second Sort Field:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&First Sort Field:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   540
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Sho&w Report"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmReportHired"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public rsItems As ADODB.Recordset

Private Sub cboSort_Click(Index As Integer)
  chkSort(Index).Enabled = cboSort(Index).ListIndex > 0
End Sub

Private Sub cmdShow_Click()

  With rptReportHired
    ' Bind the Recordset to the Report
    Set .DataSource = rsItems
    .DataMember = rsItems.DataMember
    .Sections("Section2").Controls("lblAddress").Caption = "Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
        
    ' Update the lblRecCount control on the report (Page Footer Section) to show RecordCount
'    .Sections("Section5").Controls("lblRecCount").Caption = CreateRecords.RecordCount & " items listed"
    DoEvents
    .Show
  End With
End Sub

Private Sub Form_Load()

  Dim i As Integer
  
  ' Get a recordset to use with the report
  Set rsItems = CreateRecords
  
  ' Populate Comboboxes
  For i = 0 To cboSort.Count - 1
    With cboSort(i)
      .AddItem "(None)"
      .AddItem "Employeenum"
      .AddItem "Empname"
      .AddItem "Position"
      .AddItem "ClientName"
      .ListIndex = 0
    End With
  Next
End Sub

Public Function CreateRecords() As ADODB.Recordset
  ' Creates a disconnected recordset on the fly
  Dim rsTemp As New ADODB.Recordset, i As Integer
  
 Set CreateRecords = ReturnRst("Select * from qryHRhiredREPORT", Cnnstr, mstatic)

       
'  With rsTemp
'    ' Create Recordset Fields (Columns) here
'    .Fields.Append "ID", adVarChar, 12
'    .Fields.Append "Description", adVarChar, 30
'    .Fields.Append "Unit Price", adCurrency
'    .Fields.Append "Units In Stock", adInteger
'    .Fields.Append "Sub Total", adCurrency
'    .CursorType = adOpenKeyset
'    .LockType = adLockOptimistic
'    .Open
'
'    For i = 0 To 119
'      ' Add a new item record
'      .AddNew
'      ![id] = "PID" & Format(i, "00000")
'      ![Description] = "Sample Stock Item " & i
'      Randomize Timer: ![Unit Price] = 101 * Rnd ' Add a random price
'      Randomize Timer: ![Units In Stock] = Int(2000 * Rnd) ' Add a random number of items
'      ![Sub Total] = ![Unit Price] * ![Units In Stock]
'      .Update
'    Next
'  End With
'Set CreateRecords = rsTemp
End Function

Private Sub SortRecords()
  Dim sTemp As String, i As Integer
  For i = 0 To cboSort.Count - 1
    sTemp = sTemp & IIf(cboSort(i).ListIndex > 0, "[" & cboSort(i) & "]" & IIf(chkSort(i).Value = vbChecked, " DESC,", ","), "")
  Next
 ' sTemp = "[empname],"
  rsItems.Sort = sTemp
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsItems.Close
  Set rsItems = Nothing
End Sub
