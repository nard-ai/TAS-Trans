VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmHREmpName 
   Caption         =   "List of Employees"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5535
   Begin MSComctlLib.ListView lvwEmp 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7011
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Last Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "First Name"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmHREmpName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsEmp As ADODB.Recordset
Dim list As ListItem
Dim msurnm, mfirnm As String

Private Sub Form_Load()

    Set rsEmp = cnDCPI.Execute("Select * From HREmployees")
    Do While Not rsEmp.EOF
        Set list = lvwEmp.ListItems.add(, , rsEmp!employeenum)
        list.SubItems(1) = IIf(IsNull(rsEmp!surname), "", rsEmp!surname)
        list.SubItems(2) = IIf(IsNull(rsEmp!FirstName), "", rsEmp!FirstName)
        rsEmp.MoveNext
    Loop
    rsEmp.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsEmp = Nothing
End Sub


Private Sub lvwEmp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvwEmp.Sorted = True
lvwEmp.SortKey = ColumnHeader.Index - 1
lvwEmp.Refresh
End Sub

Private Sub lvwEmp_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mempnum = Item.Text
    msurnm = Item.SubItems(1)
    mfirnm = Item.SubItems(2)
    mfullnm = msurnm & ", " & mfirnm
    
    Unload Me
End Sub
