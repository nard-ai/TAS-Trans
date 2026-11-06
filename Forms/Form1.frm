VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1320
      Top             =   2280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   780
      Left            =   495
      TabIndex        =   0
      Top             =   1080
      Width           =   1140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
Dim StartDate
Dim EndDate
Dim mFortheMonth As Integer
Dim mTotalEmp As Integer

StartDate = Format(frmSelectCutoff.txtCutOffDate(0).Text, "yyyymmdd")
EndDate = Format(frmSelectCutoff.txtCutOffDate(1).Text, "yyyymmdd")
mFortheMonth = Format(frmSelectCutoff.txtCutOffDate(0).Text, "m")
mTotalEmp = rsPayrollTrans.RecordCount

strQuery = "select sum(grosspay) as Gross from payrolltrans WHERE " & _
        "AgencyCode = " & mAgencyCode & " and ClientCode = " & mclientcode & " and StartCutOffText = '" & StartDate & "' and EndCutOffText = '" & EndDate & "'"
    
    
    
    
    Dim Cls As clsCommon
    Set Cls = New clsCommon
    
    Set rsItems = ReturnRst("Select * from clienttable " & whereCon & Sortby, Cnnstr, mstatic)
    If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
    With rpStatement
         Set .DataSource = rsItems
             .DataMember = rsItems.DataMember
             .Sections("Section1").Controls("txtClientAddress").DataField = "ClientAddress"
              Dim billing As String
              billing = "      This  is to bill you the amount in PESOS: "
              
              billing = billing & UCase(Cls.Word(11427.68)) 'amount
              
              billing = billing & " ting for the period "
              billing = billing & " April 27 - May 10 2005 " 'cut off
              billing = billing & "."
             .Sections("Section1").Controls("lblBillStatement").Caption = billing
             .Sections("Section1").Controls("lblAgencyAssistant").Caption = GetAssistant(0)
             .Sections("Section1").Controls("lblAgencyAssistantPosition").Caption = GetAssistant(1)
              .Sections("Section1").Controls("lblGenManager").Caption = GetAssistant(2)
             .Sections("Section1").Controls("lblGenManCaption").Caption = GetAssistant(3)
           .Sections("Section1").Controls("lblAccounting").Caption = GetAccounting(0, 1)
           .Sections("Section1").Controls("lblAccountingCaption").Caption = GetAccounting(1, 1)
             .Show
    End With
End Sub

Public Function GetAssistant(ByVal WhatField As Integer)
Dim rsAssistant As New ADODB.Recordset
Set rsAssistant = ReturnRst("Select AgencyAssistant,AgencyAssitantPosition,AgencyContactPerson,AgencyContactPosition from companytable " & whereCon & Sortby, Cnnstr, mstatic)
If rsAssistant.EOF = False Then
    If WhatField = 0 Then GetAssistant = rsAssistant.Fields!AgencyAssistant
    If WhatField = 1 Then GetAssistant = rsAssistant.Fields!AgencyAssitantPosition
    If WhatField = 2 Then GetAssistant = rsAssistant.Fields!AgencyContactPerson
    If WhatField = 3 Then GetAssistant = rsAssistant.Fields!AgencyContactPosition
End If
End Function

Public Function GetAccounting(ByVal WhatField As Integer, ByVal whereCon As Integer) As String
Dim rsAssistant As New ADODB.Recordset
Set rsAssistant = ReturnRst("Select ClientContactPerson,ClientContactPosition from clienttable where clientcode= " & whereCon & Sortby, Cnnstr, mstatic)
If rsAssistant.EOF = False Then
    If WhatField = 0 Then GetAccounting = IIf(IsNull(rsAssistant.Fields!ClientContactPerson) = True, "<Not declared>", rsAssistant.Fields!ClientContactPerson)
    If WhatField = 1 Then GetAccounting = IIf(IsNull(rsAssistant.Fields!ClientContactPosition) = True, "<Not declared>", rsAssistant.Fields!ClientContactPosition)
End If
End Function

