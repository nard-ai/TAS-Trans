VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPhilHealth 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PhilHealth Table Maintenance"
   ClientHeight    =   4485
   ClientLeft      =   825
   ClientTop       =   5010
   ClientWidth     =   7995
   Icon            =   "frmPhilHealth.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   7995
   Begin TabDlg.SSTab SSTab1 
      Height          =   1275
      Index           =   2
      Left            =   450
      TabIndex        =   4
      Top             =   1935
      Visible         =   0   'False
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   2249
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   -2147483639
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "      "
      TabPicture(0)   =   "frmPhilHealth.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblExit"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.TextBox Text 
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
         Index           =   4
         Left            =   5715
         TabIndex        =   21
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox Text 
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
         Index           =   3
         Left            =   4320
         TabIndex        =   8
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox Text 
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
         Index           =   2
         Left            =   2970
         TabIndex        =   7
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox Text 
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
         Index           =   1
         Left            =   1620
         TabIndex        =   6
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox Text 
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
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "PH EE"
         Height          =   240
         Index           =   1
         Left            =   5715
         TabIndex        =   22
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label lblExit 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6840
         MouseIcon       =   "frmPhilHealth.frx":08E6
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   45
         Width           =   195
      End
      Begin VB.Label Label5 
         Caption         =   "PH ER"
         Height          =   240
         Index           =   0
         Left            =   4365
         TabIndex        =   12
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label4 
         Caption         =   "Salary"
         Height          =   240
         Left            =   2970
         TabIndex        =   11
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Max wage"
         Height          =   240
         Left            =   1620
         TabIndex        =   10
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Min wage"
         Height          =   240
         Left            =   270
         TabIndex        =   9
         Top             =   270
         Width           =   1320
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2940
      Left            =   315
      TabIndex        =   3
      Top             =   1170
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   5186
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3555
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   765
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   6271
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   -2147483638
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "PHILHEALTH"
      TabPicture(0)   =   "frmPhilHealth.frx":0BF0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7425
      Top             =   225
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPhilHealth.frx":0C0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   600
         Picture         =   "frmPhilHealth.frx":0C6A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1800
         Picture         =   "frmPhilHealth.frx":0DB4
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1215
         Picture         =   "frmPhilHealth.frx":12E6
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmPhilHealth.frx":1818
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "New"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   3015
         Picture         =   "frmPhilHealth.frx":1D4A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   2385
         Picture         =   "frmPhilHealth.frx":1E94
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   2
         X1              =   0
         X2              =   800
         Y1              =   1500
         Y2              =   1200
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PURPOSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(1)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   210
   End
End
Attribute VB_Name = "frmPhilHealth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public i As Integer
'
Dim blSave As Boolean

Private Sub cmdRefresh_Click()
   Call loadlv
   Call ValidateButton(Me, True, True, False, False, False, True)
    ListView1.Enabled = True
    ListView1.SelectedItem.SELECTED = True

End Sub

Private Sub Form_Load()
    Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
   Call loadlv
   Call ValidateButton(Me, True, False, False, False, False, False)

End Sub

Private Sub cmdAdd_Click()
Call ValidateButton(Me, False, False, True, True, False, False)
    ListView1.Enabled = False
        
    blSave = True
    SSTab1(2).Visible = True
    Text(0).Enabled = True
    Text(1).Enabled = True
    
    Call clearText
    Text(1).Text = "999999.99"
    Text(1).Enabled = False
End Sub

Private Sub save()
    If blSave = True Then
          If CheckText = True Then
           sqlString = "Update philhealthtable  set maxwage=" & Text(0).Text - 0.01 & " WHERE ID= " & getMaxWAGEid("philhealthtable") & ""
           Call Action_Query(sqlString, cnnStr)
      
            sqlString = "Insert into philhealthtable (MINWAGE,MAXWAGE ,SALARY, PH_ER,PH_EE,USERID) values ('" & Text(0).Text & "','" & Text(1).Text & "','" & Text(2).Text & "','" & Text(3).Text & "','" & Text(4).Text & "'," & operatorID & ")"
      Else
        Exit Sub
        
         End If
    Else
        sqlString = "Update philhealthtable  set minwage=" & IIf(Text(0).Text = "", 0, Text(0).Text) & ",maxwage=" & IIf(Text(1).Text = "", 0, Text(1).Text) & ",salary=" & IIf(Text(2).Text = "", 0, Text(2).Text) & ",ph_er=" & IIf(Text(3).Text = "", 0, Text(3).Text) & ",ph_ee=" & IIf(Text(4).Text = "", 0, Text(4).Text) & ",USERID=" & operatorID & " where id =" & ListView1.SelectedItem.Text & ""
    End If
    Call Action_Query(sqlString, cnnStr)
    SSTab1(2).Visible = False
    Call loadlv
End Sub

Function CheckText() As Boolean
If Val(Text(0)) > Val(Text(1)) Then
MsgBox "Minimum wage must be less than Maximum wage"
CheckText = False
Else
CheckText = True

End If


End Function
Private Sub cmdCancel_Click()
    SSTab1(2).Visible = False
    Call ValidateButton(Me, True, True, False, False, False, True)
    ListView1.Enabled = True
    ListView1.SelectedItem.SELECTED = True
   
End Sub

Private Sub cmdDelete_Click()
    Dim strID As Integer
    If ListView1.SelectedItem.Index < ListView1.ListItems.Count Then
    MsgBox "Cannot remove selected item. Only the last row is allow to delete.", vbExclamation
    Exit Sub
    End If
    
    If MsgBox("Are you sure you want to remove this item?", vbQuestion + vbYesNo) = vbYes Then
        strID = ListView1.ListItems(ListView1.SelectedItem.Index).Text
        Call Action_Query("Delete from philhealthtable where id=" & strID & "", cnnStr)
         sqlString = "Update philhealthtable  set maxwage=999999 WHERE ID= " & getMaxWAGEid("philhealthtable") & ""
         Call Action_Query(sqlString, cnnStr)
      
        Call loadlv
    End If
End Sub

Private Sub cmdEdit_Click()
Call ValidateButton(Me, False, False, True, True, False, False)
    ListView1.Enabled = False
    
    
    blSave = False
    Text(0).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text
    Text(1).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text
    Text(2).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(3).Text
    Text(3).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(4).Text
    Text(4).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(5).Text
    
    SSTab1(2).Visible = True
    Text(0).Enabled = False
    Text(1).Enabled = False
End Sub

Private Sub cmdSave_Click()
Dim X As Integer
For X = 0 To 4
If Text(X).Text = "" Then MsgBox "Fill up the blank box.", vbInformation: Exit Sub
Next

Call ValidateButton(Me, True, True, False, False, False, True)
    
    Call save
End Sub
Private Sub lblExit_Click()
    SSTab1(2).Visible = False
        Call cmdCancel_Click
End Sub


Private Sub loadlv()
    Dim rsList As New ADODB.Recordset
    ListView1.ListItems.Clear
        
            Set rsList = ReturnRst("Select ID, MINWAGE as [MIN WAGE],MAXWAGE AS [MAX WAGE],SALARY, PH_ER,PH_EE from philhealthtable order by ID ", cnnStr, mstatic)
    
        If rsList.EOF = False Then
            PopulateList ListView1, rsList
        End If
        ListView1.ColumnHeaders(1).Width = 0
        ListView1.ColumnHeaders(2).Width = 1400
        ListView1.ColumnHeaders(3).Width = 1400
        ListView1.ColumnHeaders(4).Width = 1400
        ListView1.ColumnHeaders(5).Width = 1400
        ListView1.ColumnHeaders(6).Width = 1400
        
        ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
        ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
        ListView1.ColumnHeaders(3).Alignment = lvwColumnRight
        ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
        
        ListView1.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub clearText()
    Dim X As Integer
    For X = 0 To 4
    Text(X).Text = ""
    Next X
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
cmdDelete.Enabled = True
cmdedit.Enabled = True
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strValid As String
strValid = "0123456789." & Chr(13) & Chr(8)
If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0

If Index <= 3 Then
    If KeyAscii = 13 Then
        If Index = 3 Then
            Text(0).SetFocus
        Else
            Text(Index + 1).SetFocus
        End If
    End If
End If
End Sub
