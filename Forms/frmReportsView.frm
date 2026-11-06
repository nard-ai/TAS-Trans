VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReportsView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTS "
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   Icon            =   "frmReportsView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4305
   Begin VB.Frame Frame1 
      Caption         =   "Date Range"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   30
      TabIndex        =   18
      Top             =   1200
      Width           =   4215
      Begin MSComCtl2.DTPicker dtRanges 
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   19
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   60489731
         CurrentDate     =   38148
      End
      Begin MSComCtl2.DTPicker dtRanges 
         Height          =   315
         Index           =   1
         Left            =   2610
         TabIndex        =   20
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   60489731
         CurrentDate     =   38163
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2070
         TabIndex        =   22
         Top             =   330
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -60
         TabIndex        =   21
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.ListBox lstVw 
      Height          =   1035
      ItemData        =   "frmReportsView.frx":0442
      Left            =   30
      List            =   "frmReportsView.frx":0452
      TabIndex        =   17
      Top             =   0
      Width           =   4215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   435
      Left            =   3210
      Picture         =   "frmReportsView.frx":0499
      TabIndex        =   16
      Top             =   3240
      Width           =   1035
   End
   Begin VB.CheckBox chkFresh 
      Caption         =   "REFRESH DATA"
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
      Left            =   210
      TabIndex        =   15
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Preview"
      Enabled         =   0   'False
      Height          =   435
      Left            =   2160
      Picture         =   "frmReportsView.frx":08DB
      TabIndex        =   14
      Top             =   3240
      Width           =   1065
   End
   Begin VB.OptionButton optView 
      BackColor       =   &H8000000B&
      Caption         =   "Absences"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   5730
      TabIndex        =   9
      Top             =   3780
      Width           =   2295
   End
   Begin VB.OptionButton optView 
      BackColor       =   &H8000000B&
      Caption         =   "Under-time"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   5730
      TabIndex        =   8
      Top             =   3540
      Width           =   2295
   End
   Begin VB.OptionButton optView 
      BackColor       =   &H8000000B&
      Caption         =   "Tardiness"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   5730
      TabIndex        =   7
      Top             =   3300
      Width           =   2295
   End
   Begin VB.OptionButton optView 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      Caption         =   "Multiple Swipe"
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   5670
      TabIndex        =   6
      Top             =   3180
      Width           =   1995
   End
   Begin VB.OptionButton optView 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      Caption         =   "Invalid Shift"
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   5670
      TabIndex        =   5
      Top             =   2880
      Width           =   1995
   End
   Begin VB.OptionButton optView 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      Caption         =   "No Time-In or Time-Out"
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   5670
      TabIndex        =   4
      Top             =   2595
      Width           =   1995
   End
   Begin VB.Frame fraSelect 
      Height          =   1065
      Left            =   60
      TabIndex        =   10
      Top             =   2070
      Width           =   4185
      Begin VB.CheckBox chkOption 
         Caption         =   "Filter By"
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
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   2415
      End
      Begin VB.ComboBox cboList 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   0
         Top             =   630
         Width           =   1275
      End
      Begin VB.ComboBox cboList 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   1
         Top             =   630
         Width           =   1275
      End
      Begin VB.ComboBox cboList 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2700
         TabIndex        =   2
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Department"
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
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Project"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Facility"
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
         Index           =   2
         Left            =   2400
         TabIndex        =   11
         Top             =   420
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmReportsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsload As ADODB.Recordset
Private mdept As String
Private mLine As String

Private Property Get gdept() As String
    gdept = mdept
End Property

Private Property Get gLine() As String
    gLine = mLine
End Property

Private Function mystr_filter_load() As String
    Dim mysql As String
    If cboList(0).Text <> "" And cboList(1).Text = "" And cboList(2).Text = "" Then
        mysql = "EmpDept = '" & Replace(Trim(cboList(0).Text), "'", "''") & "'"
    ElseIf cboList(0).Text <> "" And cboList(1).Text <> "" And cboList(2).Text = "" Then
        mysql = "EmpDept = '" & Replace(Trim(cboList(0).Text), "'", "''") & "' And EmpLine = '" & Trim(cboList(1).Text) & "'"
    ElseIf cboList(0).Text <> "" And cboList(1).Text <> "" And cboList(2).Text <> "" Then
        mysql = "EmpDept = '" & Replace(Trim(cboList(0).Text), "'", "''") & "' " & _
                "And EmpLine = '" & Trim(cboList(1).Text) & "' " & _
                "And EmpFacility = '" & Trim(cboList(2).Text) & "'"
    ElseIf cboList(0).Text = "" And cboList(1).Text <> "" And cboList(2).Text = "" Then
        mysql = "EmpLine = '" & Trim(cboList(1).Text) & "'"
    ElseIf cboList(0).Text = "" And cboList(1).Text <> "" And cboList(2).Text <> "" Then
        mysql = "EmpLine = '" & Trim(cboList(1).Text) & "' And EmpFacility = '" & Trim(cboList(2).Text) & "'"
    ElseIf cboList(0).Text <> "" And cboList(1).Text = "" And cboList(2).Text <> "" Then
        mysql = "EmpDept = '" & Replace(Trim(cboList(0).Text), "'", "''") & "' And EmpFacility = '" & Trim(cboList(2).Text) & "'"
    ElseIf cboList(0).Text = "" And cboList(1).Text = "" And cboList(2).Text <> "" Then
        mysql = "EmpFacility = '" & Trim(cboList(2).Text) & "'"
    Else
        Exit Function
    End If
    mystr_filter_load = mysql
End Function

Private Function load_Transmission(ByVal Intload As myIntOpt)
    Dim mstr As String
    Dim mycls As New clsEmpRS
    
    Dim rsTrans As New ADODB.Recordset
    Set rsTrans = ReturnRst("Select Trans_ID from Emp_Transmission " & _
        "Where Trans_From like '" & str(dtRanges(0)) & "' " & _
        "And Trans_To like '" & str(dtRanges(1)) & "' " & _
        "And Trans_Mark like '" & Trim(Intload) & "'", cnnstr_mdb, mread)
    
    If rsTrans.EOF Then
        mstr = mAuto_ID(mTrans)
    Else
        If chkFresh.Value = 0 Then
            Call Load_Data(Intload, Trim(rsTrans(0)))
            Exit Function
        Else
            Call Action_Query("Delete from Emp_Transmission Where Trans_ID like '" & Trim(rsTrans(0)) & "'", cnnstr_mdb)
            mstr = rsTrans(0)
        End If
    End If
    
    Call Action_Query("Insert Into Emp_Transmission(Trans_ID,Trans_From,Trans_To,Trans_Mark) " & _
            "values ('" + Trim(mstr) + "','" + str(dtRanges(0)) + _
            "','" + str(dtRanges(1)) + "','" & Trim(Intload) & "')", cnnstr_mdb)
    
    Select Case Intload
        Case myIntOpt.mInOut
            Call Action_Query("Delete from Emp_InOut Where Empflg like '" & Trim(mstr) & "'", cnnstr_mdb)
            Set rsload = New ADODB.Recordset
            Set rsload = ReturnRst("Select empn, date, count(date) as dcnt, " & _
                     "max(mode) as mxcode, max(time) as mxtime from download having !Empty(empn) and dcnt = 1 " & _
                     "And dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                     "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' Group by empn,date " & _
                     "Order by date,empn", cnnstr_time, mstatic)
          '  If rsload.EOF = True Then GoTo msgResult
            
            Do Until rsload.EOF
                Call Action_Query("Insert Into Emp_InOut(EmpID,EmpDate,Empdcnt,EmpMode,EmpTime,Empflg) " & _
                    "values ('" + Trim(rsload(0)) + "','" + str(rsload(1)) + "','" + Trim(rsload(2)) + _
                    "','" + Trim(rsload(3)) + "','" + Trim(rsload(4)) + "','" & Trim(mstr) & "')", cnnstr_mdb)
                'If pbar.Value = 100 Then pbar.Value = 0 Else pbar.Value = pbar.Value + 10
                rsload.MoveNext
            Loop
            'pbar.Value = 0
            Set rsload = Nothing
            
        Case myIntOpt.mInValidShft
            Dim rs1st As New ADODB.Recordset
            Dim rs2nd As New ADODB.Recordset
            Dim rs3rd As New ADODB.Recordset
            Dim rs4th As New ADODB.Recordset
            Dim rs5th As New ADODB.Recordset
            
            Call Action_Query("Delete from Emp_InvalidShift", cnnstr_mdb)
          ' ORIGINAL' Set rs1st = ReturnRst("Select empn,date,in,out,shiftcode from time01 " & _
                        "where dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        "And shiftcode in ('1','M1') and not (In) between 5 and 14 " & _
                        "And In <> 0 And Out <> 0", cnnstr_time, mstatIC)
            '
            ' Set rs1st = ReturnRst("Select empn,date,in,out,shiftcode from time01 " & _
                        "where dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        "And shiftcode in ('1','M1') and not (In) between 5 and 14" & _
                        "And In <> 0 And Out <> 0 and date not in (Select date from holiday where  dtos(date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' and '" & Format(dtRanges(1).Value, "yyyyMMdd") & "')", cnnstr_time, mstatic)
                        
            Set rs1st = ReturnRst("Select empn,date,in,out,shiftcode from time01 " & _
                        "where (dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        ") And shiftcode in ('1','M1') and " & _
                        " (   ((In) between 12 and 16) and ((out) between 16.1 and 23.59)     or (( (In) <7) and ((out)between 3 and 10.3))           )" & _
                        "And In <> 0 And Out <> 0 and date not in (Select date from holiday where  dtos(date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' and '" & Format(dtRanges(1).Value, "yyyyMMdd") & "')", cnnstr_time, mstatic)
                        
                        
            Do Until rs1st.EOF
                Call Action_Query("Insert into Emp_InvalidShift(EmpID,Emp_Date,Emp_In,Emp_Out,Emp_ShiftCode,Empflg) " & _
                        "values ('" + Trim(rs1st(0)) + "','" + str(rs1st(1)) + "','" + Trim(rs1st(2)) + _
                        "','" + Trim(rs1st(3)) + "','" + Trim(rs1st(4)) + "','" + mstr & "')", cnnstr_mdb)
                        
                        
                        'If pbar.Value = 100 Then pbar.Value = 0 Else pbar.Value = pbar.Value + 10
                rs1st.MoveNext
            Loop
            'pbar.Value = 100
            Set rs1st = Nothing
            
          ' ORIGINAL' Set rs2nd = ReturnRst("Select empn,date,in,out,shiftcode from time01 " & _
                        "where dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        "And shiftcode in ('2','M2') and not (In) between 12 and 22 " & _
                        "And In <> 0 And Out <> 0", cnnstr_time, mread)
                        
     '    Set rs2nd = ReturnRst("Select empn,date,in,out,shiftcode from time01 " & _
                        "where dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        "And shiftcode in ('2','M2') and not (In) between 12 and 22  " & _
                        "And In <> 0 And Out <> 0 and date not in (Select date from holiday where  dtos(date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' and '" & Format(dtRanges(1).Value, "yyyyMMdd") & "')", cnnstr_time, mread)
          Set rs2nd = ReturnRst("Select empn,date,in,out,shiftcode from time01 " & _
                        "where (dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        ") And shiftcode in ('2','M2') and " & _
                        "(          ((In) between 5 and 8) and ((out) between 8.1 and 17)  or (( (In) <15) and ((out)between 6 and 15))       ) " & _
                         " And In <> 0 And Out <> 0 and date not in (Select date from holiday where  dtos(date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' and '" & Format(dtRanges(1).Value, "yyyyMMdd") & "')", cnnstr_time, mread)
                                                       
                    
                   
                   '  If rs2nd.EOF = True Then GoTo msgResult
            Do Until rs2nd.EOF
                Call Action_Query("Insert into Emp_InvalidShift(EmpID,Emp_Date,Emp_In,Emp_Out,Emp_ShiftCode,Empflg) " & _
                        "values ('" + Trim(rs2nd(0)) + "','" + str(rs2nd(1)) + "','" + Trim(rs2nd(2)) + _
                        "','" + Trim(rs2nd(3)) + "','" + Trim(rs2nd(4)) + "','" + Trim(mstr) & "')", cnnstr_mdb)
                        
                        'If pbar.Value = 100 Then pbar.Value = 0 Else pbar.Value = pbar.Value + 10
                rs2nd.MoveNext
            Loop
            'pbar.Value = 100
            Set rs2nd = Nothing
            
           'ORIGINAL' Set rs3rd = ReturnRst("Select empn,date,in,out,shiftcode from time01 " & _
                        "where dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        "And ShiftCode like '3' and not (in) between 22 and 24 " & _
                        "And not (in) between 1 and 6 and in <> 0 and Out <> 0 " & _
                        "and advance = 1", cnnstr_time, mstatic)
                        
                    'Set rs3rd = ReturnRst("Select empn,date,in,out,shiftcode from time01 " & _
                        "where dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        "And ShiftCode like '3' and not (in) between 22 and 24 " & _
                        "And not (in) between 1 and 6 and in <> 0 and Out <> 0 " & _
                        "and advance = 1 and date not in (Select date from holiday where  dtos(date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' and '" & Format(dtRanges(1).Value, "yyyyMMdd") & "')", cnnstr_time, mstatic)
                        
                     Set rs3rd = ReturnRst("Select empn,date,in,out,shiftcode from time01 " & _
                        "where (dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        ") And ShiftCode like '3' and " & _
                          "(     ((In)  between  1 and 8) and ((out) between 8 and 23)  or (((In) <15) and ((out)between 15 and 23.59))    ) " & _
                         " And In <> 0 And Out <> 0 and advance = 1 and date not in (Select date from holiday where  dtos(date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' and '" & Format(dtRanges(1).Value, "yyyyMMdd") & "')", cnnstr_time, mstatic)
                       
                       
           ' If rs3rd.EOF = True Then GoTo msgResult
            Do Until rs3rd.EOF
                Call Action_Query("Insert into Emp_InvalidShift(EmpID,Emp_Date,Emp_In,Emp_Out,Emp_ShiftCode,Empflg) " & _
                        "values ('" + Trim(rs3rd(0)) + "','" + str(rs3rd(1)) + "','" + Trim(rs3rd(2)) + _
                        "','" + Trim(rs3rd(3)) + "','" + Trim(rs3rd(4)) + "','" + Trim(mstr) & "')", cnnstr_mdb)
                        'If pbar.Value = 100 Then pbar.Value = 0 Else pbar.Value = pbar.Value + 10
                rs3rd.MoveNext
            Loop
'            pbar.Value = 100
            Set rs3rd = Nothing
            
            Set rs4th = ReturnRst("Select empn,date,in,out,shiftcode from time01 " & _
                        "where (dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        ") And shiftcode in ('4','S4') and not (In) between 7 and 15 " & _
                        "And In <> 0 And Out <> 0 And Advance = 1", cnnstr_time, mread)
            
            Do Until rs4th.EOF
                Call Action_Query("Insert into Emp_InvalidShift(EmpID,Emp_Date,Emp_In,Emp_Out,Emp_ShiftCode,Empflg) " & _
                        "values ('" + Trim(rs4th(0)) + "','" + str(rs4th(1)) + "','" + Trim(rs4th(2)) + _
                        "','" + Trim(rs4th(3)) + "','" + Trim(rs4th(4)) + "','" + Trim(mstr) & "')", cnnstr_mdb)
                        'If pbar.Value = 100 Then pbar.Value = 0 Else pbar.Value = pbar.Value + 10
                rs4th.MoveNext
            Loop
'            pbar.Value = 0
            Set rs4th = Nothing
            
            
             Set rs5th = ReturnRst("Select empn,date,in,out,shiftcode from time01 " & _
                        "where (dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        ") And shiftcode in ('R') and " & _
                        "(          ((In) between 5 and 8) and ((out) between 8.1 and 15 )  or (( (In) <8) and ((out)between 8.1 and 15 )) or ((In) between 12 and 16) and ((out) between 16.1 and 23.59)     or (( (In) <7) and ((out)between 3 and 10.3))       ) " & _
                         " And In <> 0 And Out <> 0 and date not in (Select date from holiday where  dtos(date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' and '" & Format(dtRanges(1).Value, "yyyyMMdd") & "')", cnnstr_time, mread)
                                                       
                    
                   
                    '    If rs5th.EOF = True Then GoTo msgResult
            Do Until rs5th.EOF
                Call Action_Query("Insert into Emp_InvalidShift(EmpID,Emp_Date,Emp_In,Emp_Out,Emp_ShiftCode,Empflg) " & _
                        "values ('" + Trim(rs5th(0)) + "','" + str(rs5th(1)) + "','" + Trim(rs5th(2)) + _
                        "','" + Trim(rs5th(3)) + "','" + Trim(rs5th(4)) + "','" + Trim(mstr) & "')", cnnstr_mdb)
                        'If pbar.Value = 100 Then pbar.Value = 0 Else pbar.Value = pbar.Value + 10
                rs5th.MoveNext
            Loop
            'pbar.Value = 100
            Set rs5th = Nothing
            
            
        Case myIntOpt.mXSwipe
            Call Action_Query("Delete from Emp_Swipe Where Empflg like '" & Trim(mstr) & "'", cnnstr_mdb)
            Set rsload = New ADODB.Recordset
            Set rsload = ReturnRst("Select empn,time,mode,count(mode) as mcnt,date from download " & _
                    "group by empn,date,mode,time having !Empty(empn) and mcnt <> 1 " & _
                    "And dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                    "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "'", cnnstr_time, mstatic)
       '   If rsload.EOF = True Then GoTo msgResult
            
            Do Until rsload.EOF
                Call Action_Query("Insert Into Emp_Swipe(EmpID,EmpTime,EmpMode,EmpCnt,EmpDate,Empflg) " & _
                    "values ('" + Trim(rsload(0)) + "','" + Trim(rsload(1)) + "','" + Trim(rsload(2)) + _
                    "','" + Trim(rsload(3)) + "','" + str(rsload(4)) + "','" & Trim(mstr) & "')", cnnstr_mdb)
                'If pbar.Value = 100 Then pbar.Value = 0 Else pbar.Value = pbar.Value + 10
                rsload.MoveNext
            Loop
'            pbar.Value = 0
            Set rsload = Nothing
            
        Case myIntOpt.mTardiness
            Call Disply_Print_Report
        Case myIntOpt.mUTime
            Set rsload = New ADODB.Recordset
            Call Action_Query("Delete from Emp_UnderTime Where Empflg like '" & Trim(mstr) & "'", cnnstr_mdb)
            Set rsload = ReturnRst("Select Empn,In,Out,Regular,Date from time01 " & _
                        "Where dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        "And Regular > 4 and Regular < 7.55 and advance = 1", cnnstr_time, mread)
            Do Until rsload.EOF
                Call Action_Query("Insert into Emp_UnderTime(EmpID,Emp_In,Emp_Out,Emp_Diff,Emp_Date,Empflg) " & _
                    "values ('" + Trim(rsload(0)) + "','" + Trim(rsload(1)) + "','" + Trim(rsload(2)) + _
                    "','" + Trim(rsload(3)) + "','" + str(rsload(4)) + "','" + Trim(mstr) + "')", cnnstr_mdb)
                    'If pbar.Value = 100 Then pbar.Value = 0 Else pbar.Value = pbar.Value + 10
                rsload.MoveNext
            Loop
            'pbar.Value = 1
            Set rsload = Nothing
        Case myIntOpt.mAbsences
            Dim mysql_str As String
            Dim rsHday As New ADODB.Recordset
            Set rsHday = ReturnRst("Select Holdate from Emp_Holiday Where Holdate between #" & str(dtRanges(0)) & _
                            "# And #" & str(dtRanges(1)) & "#", cnnstr_mdb, mstatic)
            Set rsload = New ADODB.Recordset
            
            If Not rsHday.EOF Then
                ReDim jn_str(1 To rsHday.RecordCount) As String
                Do Until rsHday.EOF
                    If rsHday.RecordCount <> 1 Then jn_str(rsHday.AbsolutePosition) = "'" & Format(rsHday.Fields(0), "yyyyMMdd") & _
                            "'": mysql_str = Join(jn_str, ",") Else mysql_str = "'" & Format(rsHday.Fields(0), "yyyyMMdd") & "'"
                    rsHday.MoveNext
                Loop
                Set rsHday = Nothing
                Set rsload = ReturnRst("Select distinct Empn,In,Out,Date from time01 " & _
                        "Where dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        "And in = 0 and out = 0 And !cdow(Date) = 'Sunday' " & _
                        "And !dtos(Date) in (" & mysql_str & ")", cnnstr_time, mread)
            Else
                Set rsload = ReturnRst("Select distinct Empn,In,Out,Date from time01 " & _
                        "Where dtos(Date) between '" & Format(dtRanges(0).Value, "yyyyMMdd") & "' " & _
                        "And '" & Format(dtRanges(1).Value, "yyyyMMdd") & "' " & _
                        "And in = 0 and out = 0 And !cdow(Date) = 'Sunday'", cnnstr_time, mread)
            End If
            
            Call Action_Query("Delete from Emp_Absences Where Empflg like '" & Trim(mstr) & "'", cnnstr_mdb)
            Do Until rsload.EOF
                Call Action_Query("Insert into Emp_Absences(EmpID,Emp_In,Emp_Out,EmpDate,Empflg) " & _
                        "values ('" + Trim(rsload(0)) + "','" + Trim(rsload(1)) + "','" + Trim(rsload(2)) + _
                        "','" + str(rsload(3)) + "','" + Trim(mstr) & "')", cnnstr_mdb)
                'If pbar.Value = 100 Then pbar.Value = 0 Else pbar.Value = pbar.Value + 10
                rsload.MoveNext
            Loop
            'pbar.Value = 0
            Set rsload = Nothing
    End Select
    Call Load_Data(Intload, mstr)
    Set rsTrans = Nothing
Exit Function
msgResult:
   Unload Status
   Call MsgNoRecords(frmReportsView)
End Function
    
Private Function m_tardiness(ByVal m_EmpID As String, ByVal m_late As String)
    Dim rs_trd As New ADODB.Recordset
    
    Set rs_trd = ReturnRst("Select EmpID from Emp_Tardiness_Sum " & _
                "Where EmpId = '" & Trim(m_EmpID) & "'", cnnstr_mdb, mstatic)
    If Not rs_trd.EOF Then
        
    Else
        
        'call action_query("Insert into Emp_Tardiness_Sum(EmpID,Emp_tLate,Empflg) " & _
                "values ('" + rs_trd.Fields(0) + "','" +
    End If
    
End Function

Private Function Load_Data(ByVal initload As myIntOpt, ByVal myflds As String)
    Dim mystr As String
    Select Case initload
        Case myIntOpt.mInOut
            'Call Load_Crystal_Report("C:\Projects\HRProjects\Reports\NoTime_IN_OUT.rpt")
            'frmPreview.Show
            Dim rs_InOut As New ADODB.Recordset
            Set rs_InOut = New ADODB.Recordset
            If chkOption.Value = 1 Then
                If mystr_filter_load <> "" Then mystr = mystr_filter_load Else MsgBox "Forgot to Select field...", vbCritical, "Select Data field": Exit Function
                Set rs_InOut = ReturnRst("SHAPE {SELECT DISTINCT Emp_InOut.EmpID, Employee_List.EmpName, Employee_List.EmpDept, " & _
                        "Employee_List.EmpLine, Employee_List.EmpFacility FROM Employee_List " & _
                        "INNER JOIN Emp_InOut ON Employee_List.EmpID = Emp_InOut.EmpID " & _
                        "WHERE " & mystr & " And Emp_InOut.Empflg like '" & myflds & "' ORDER BY Employee_List.EmpDept} AS cmdLine " & _
                        "APPEND ((SHAPE {SELECT DISTINCT Emp_InOut.EmpID, Employee_List.EmpName, Employee_List.EmpDept, Employee_List.EmpLine " & _
                        "FROM Employee_List INNER JOIN Emp_InOut ON Employee_List.EmpID = Emp_InOut.EmpID} AS cmdName " & _
                        "APPEND ({SELECT DISTINCT EmpID, EmpDate, EmpTime, EmpdCnt, EmpMode FROM Emp_InOut " & _
                        "WHERE Empflg like '" & myflds & "'} AS cmdTime " & _
                        "RELATE 'EmpID' TO 'EmpID') AS cmdTime) AS cmdName " & _
                        "RELATE 'EmpID' TO 'EmpID','EmpDept' TO 'EmpDept') AS cmdName", cnnstr_hmdb, mstatic)
            Else
                Set rs_InOut = ReturnRst("SHAPE {SELECT DISTINCT Emp_InOut.EmpID, Employee_List.EmpName, " & _
                        "Employee_List.EmpDept, Employee_List.EmpLine, Employee_List.EmpFacility " & _
                        "FROM Employee_List INNER JOIN Emp_InOut ON Employee_List.EmpID = Emp_InOut.EmpID " & _
                        "WHERE Emp_InOut.Empflg like '" & myflds & "' ORDER BY Employee_List.EmpDept} AS cmdLine " & _
                        "APPEND ((SHAPE {SELECT DISTINCT Emp_InOut.EmpID, Employee_List.EmpName, Employee_List.EmpDept, Employee_List.EmpLine " & _
                        "FROM Employee_List INNER JOIN Emp_InOut ON Employee_List.EmpID = Emp_InOut.EmpID} AS cmdName " & _
                        "APPEND ({SELECT DISTINCT EmpID, EmpDate, EmpTime, EmpdCnt, EmpMode FROM Emp_InOut " & _
                        "WHERE Empflg like '" & myflds & "'} AS cmdTime " & _
                        "RELATE 'EmpID' TO 'EmpID') AS cmdTime) AS cmdName " & _
                        "RELATE 'EmpID' TO 'EmpID','EmpDept' TO 'EmpDept') AS cmdName", cnnstr_hmdb, mstatic)
            End If
            Set RptInOut.DataSource = rs_InOut
            RptInOut.Sections("PageHeader").Controls("lblFrom").Caption = Format(dtRanges(0).Value, "MM/dd/yyyy")
            RptInOut.Sections("PageHeader").Controls("lblTo").Caption = Format(dtRanges(1).Value, "MM/dd/yyyy")
            Set rs_InOut = Nothing
            
        Case myIntOpt.mInValidShft
            Dim rsInvalid As New ADODB.Recordset
            If chkOption.Value = 1 Then
                If mystr_filter_load <> "" Then mystr = mystr_filter_load Else MsgBox "Forgot to Select Field...", vbCritical, "Select Data field": Exit Function
                Set rsInvalid = ReturnRst("SHAPE {SELECT DISTINCT Emp_InvalidShift.EmpID, Employee_List.EmpName ,Employee_List.EmpDept, " & _
                        "Employee_List.EmpLine, Employee_List.EmpFacility FROM Employee_List " & _
                        "INNER JOIN Emp_InvalidShift ON Employee_List.EmpID = Emp_InvalidShift.EmpID " & _
                        "WHERE " & mystr & " And Emp_InvalidShift.Empflg like '" & myflds & "' ORDER BY Employee_List.EmpDept} AS cmdLine " & _
                        "APPEND ((SHAPE {SELECT DISTINCT Emp_InvalidShift.EmpID, Employee_List.EmpName, Employee_List.EmpDept, Employee_List.EmpLine " & _
                        "FROM Employee_List INNER JOIN Emp_InvalidShift ON Employee_List.EmpID = Emp_InvalidShift.EmpID} AS cmdName " & _
                        "APPEND ({SELECT DISTINCT EmpID, Emp_Date, Emp_In, Emp_Out, Emp_ShiftCode FROM Emp_InvalidShift " & _
                        "WHERE Empflg like '" & myflds & "'} AS cmdTime " & _
                        "RELATE 'EmpID' TO 'EmpID') AS cmdTime) AS cmdName " & _
                        "RELATE 'EmpID' TO 'EmpID','EmpDept' TO 'EmpDept') AS cmdName", cnnstr_hmdb, mstatic)
            Else
                Set rsInvalid = ReturnRst("SHAPE {SELECT DISTINCT Emp_InvalidShift.EmpID, Employee_List.EmpName ,Employee_List.EmpDept, " & _
                        "Employee_List.EmpLine, Employee_List.EmpFacility FROM Employee_List " & _
                        "INNER JOIN Emp_InvalidShift ON Employee_List.EmpID = Emp_InvalidShift.EmpID " & _
                        "WHERE Emp_InvalidShift.Empflg like '" & myflds & "' ORDER BY Employee_List.EmpDept} AS cmdLine " & _
                        "APPEND ((SHAPE {SELECT DISTINCT Emp_InvalidShift.EmpID, Employee_List.EmpName, Employee_List.EmpDept, Employee_List.EmpLine " & _
                        "FROM Employee_List INNER JOIN Emp_InvalidShift ON Employee_List.EmpID = Emp_InvalidShift.EmpID} AS cmdName " & _
                        "APPEND ({SELECT DISTINCT EmpID, Emp_Date, Emp_In, Emp_Out, Emp_ShiftCode FROM Emp_InvalidShift " & _
                        "WHERE Empflg like '" & myflds & "'} AS cmdTime " & _
                        "RELATE 'EmpID' TO 'EmpID') AS cmdTime) AS cmdName " & _
                        "RELATE 'EmpID' TO 'EmpID','EmpDept' TO 'EmpDept') AS cmdName", cnnstr_hmdb, mstatic)
            End If
            Set RptInvalidShift.DataSource = rsInvalid
            RptInvalidShift.Sections("PageHeader").Controls("lblFrom").Caption = Format(dtRanges(0).Value, "MM/dd/yyyy")
            RptInvalidShift.Sections("PageHeader").Controls("lblTo").Caption = Format(dtRanges(1).Value, "MM/dd/yyyy")
            Set rsInvalid = Nothing
            
        Case myIntOpt.mXSwipe
            Dim rsSwipe As New ADODB.Recordset
            If chkOption.Value = 1 Then
                If mystr_filter_load <> "" Then mystr = mystr_filter_load Else MsgBox "Forgot to Select field...", vbCritical, "Select Data field": Exit Function
                Set rsSwipe = ReturnRst("SHAPE {SELECT DISTINCT Emp_Swipe.EmpID, Employee_List.EmpName ,Employee_List.EmpDept, " & _
                        "Employee_List.EmpLine, Employee_List.EmpFacility FROM Employee_List " & _
                        "INNER JOIN Emp_Swipe ON Employee_List.EmpID = Emp_Swipe.EmpID " & _
                        "WHERE " & mystr & " AND Emp_Swipe.Empflg like '" & myflds & "' ORDER BY Employee_List.EmpDept} AS cmdLine " & _
                        "APPEND ((SHAPE {SELECT DISTINCT Emp_Swipe.EmpID, Employee_List.EmpName, Employee_List.EmpDept, Employee_List.EmpLine " & _
                        "FROM Employee_List INNER JOIN Emp_Swipe ON Employee_List.EmpID = Emp_Swipe.EmpID} AS cmdName " & _
                        "APPEND ({SELECT DISTINCT EmpID, EmpDate, EmpTime, EmpCnt, EmpMode FROM Emp_Swipe " & _
                        "WHERE Empflg like '" & myflds & "'} AS cmdTime " & _
                        "RELATE 'EmpID' TO 'EmpID') AS cmdTime) AS cmdName " & _
                        "RELATE 'EmpID' TO 'EmpID','EmpDept' TO 'EmpDept') AS cmdName", cnnstr_hmdb, mstatic)
            Else
                Set rsSwipe = ReturnRst("SHAPE {SELECT DISTINCT Emp_Swipe.EmpID, Employee_List.EmpName ,Employee_List.EmpDept, " & _
                        "Employee_List.EmpLine, Employee_List.EmpFacility FROM Employee_List " & _
                        "INNER JOIN Emp_Swipe ON Employee_List.EmpID = Emp_Swipe.EmpID " & _
                        "WHERE Emp_Swipe.Empflg like '" & myflds & "' ORDER BY Employee_List.EmpDept} AS cmdLine " & _
                        "APPEND ((SHAPE {SELECT DISTINCT Emp_Swipe.EmpID, Employee_List.EmpName, Employee_List.EmpDept, Employee_List.EmpLine " & _
                        "FROM Employee_List INNER JOIN Emp_Swipe ON Employee_List.EmpID = Emp_Swipe.EmpID} AS cmdName " & _
                        "APPEND ({SELECT DISTINCT EmpID, EmpDate, EmpTime, EmpCnt, EmpMode FROM Emp_Swipe " & _
                        "WHERE Empflg like '" & myflds & "'} AS cmdTime " & _
                        "RELATE 'EmpID' TO 'EmpID') AS cmdTime) AS cmdName " & _
                        "RELATE 'EmpID' TO 'EmpID','EmpDept' TO 'EmpDept') AS cmdName", cnnstr_hmdb, mstatic)
            End If
            Set RptSwipe.DataSource = rsSwipe
            RptSwipe.Sections("PageHeader").Controls("lblFrom").Caption = Format(dtRanges(0).Value, "MM/dd/yyyy")
            RptSwipe.Sections("PageHeader").Controls("lblTo").Caption = Format(dtRanges(1).Value, "MM/dd/yyyy")
            Set rsSwipe = Nothing
    End Select
End Function
    
Private Function bchkEnb(ByVal mChck As Boolean)
    Label2(0).Enabled = mChck
    Label2(1).Enabled = mChck
    Label2(2).Enabled = mChck
    cboList(0).Enabled = mChck
    'cboList(1).Enabled = mChck
    cboList(2).Enabled = mChck
End Function

Public Function mEmpInfo(ByVal mEmpId As String) As String
    Dim rsname As New ADODB.Recordset
    Set rsname = ReturnRst("Select Lnam,Fnam,Mnam,Decd,Line from AmMast Where Empn like '" & Trim(mEmpId) & "'", cnnstr_main, mread)
    If Not rsname.EOF Then
        mEmpInfo = Trim(rsname(0)) & ", " & Trim(rsname(1)) & " " & Trim(rsname(2))
        mdept = IIf(IsNull(rsname.Fields(3)), "", rsname.Fields(3))
        mLine = IIf(IsNull(rsname.Fields(4)), "", rsname.Fields(4))
    End If
    Set rsname = Nothing
End Function

Private Sub cboList_Click(Index As Integer)
    Select Case Index
        Case 0, 1
            If cboList(0).Text = "PROD" Then
                Label2(1).Enabled = True
                cboList(1).Enabled = True
            Else
                Label2(1).Enabled = False
                cboList(1).Enabled = False
                cboList(1).Text = ""
            End If
        Case 2
    End Select
End Sub

Private Sub chkOption_Click()
    If chkOption.Value = 1 Then Call bchkEnb(True) Else Call bchkEnb(False)
    If chkOption.Value = 1 Then cboList(0).SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Me.MousePointer = vbHourglass
    If lstVw.ListIndex = 0 Then Call load_Transmission(mInOut)
    If lstVw.ListIndex = 1 Then Call load_Transmission(mInValidShft)
    If lstVw.ListIndex = 2 Then Call load_Transmission(mXSwipe)
    If lstVw.ListIndex = 3 Then Call load_Transmission(mTardiness)
    'If optView(4).Value = True Then Call load_Transmission(mUTime)
    'If optView(5).Value = True Then Call load_Transmission(mAbsences)
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Call bchkEnb(False)
    Call load_cbofill(cboList(0), mcDept)
    Call load_cbofill(cboList(1), mcLine)
    Call load_cbofill(cboList(2), mcFac)
    'dtRanges(0).Value = Date
    'dtRanges(1).Value = Date
     Call Setdate(frmReportsView)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'mdiMain.mnuReports.Enabled = True
End Sub

Private Function Disply_Print_Report()
    Dim mycls As New clsEmpRS
    If cboList(0).Text = "PROD" Then
        Call mycls.Disply_RS_Employees(mProd, dtRanges(0).Value, dtRanges(1).Value)
        Call mycls.Disply_RS_Employees(mProdRS)
        crystlOcx.Reset
        crystlOcx.ReportFileName = "C:\Projects\HRProjects\Reports\Production_Tardiness.rpt"
        crystlOcx.SelectionFormula = Disply_Filter_Data("QryProd_Tardiness")
        crystlOcx.WindowTitle = "DCPI Production Tardiness Report"
    Else
        Call mycls.Disply_RS_Employees(mStaff, dtRanges(0).Value, dtRanges(1).Value)
        Call mycls.Disply_RS_Employees(mStaffAttendance, dtRanges(0).Value, dtRanges(1).Value)
        crystlOcx.Reset
        crystlOcx.ReportFileName = "C:\Projects\HRProjects\Reports\Staff_Tardiness.rpt"
        crystlOcx.SelectionFormula = Disply_Filter_Data("QryProd_Tardiness")
        crystlOcx.WindowTitle = "DCPI Staff Tardiness Report"
    End If
    Set mycls = Nothing
    
    With crystlOcx
        .Formulas(1) = "dtfrom = '" & (Format(dtRanges(0).Value, "MM/dd/yyyy")) & "'"
        .Formulas(2) = "dtTo = '" & (Format(dtRanges(1).Value, "MM/dd/yyyy")) & "'"
        .WindowShowExportBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowGroupTree = True
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .Action = 1
    End With
End Function

Private Function Disply_Filter_Data(ByVal myQrytable As String) As String
    Dim mysql As String
    If cboList(0).Text <> "" And cboList(1).Text = "" And cboList(2).Text = "" Then
        mysql = "{" & myQrytable & ".EmpDept} = '" & Replace(Trim(cboList(0)), "'", "''") & "'"
    ElseIf cboList(0).Text <> "" And cboList(1).Text <> "" And cboList(2).Text = "" Then
         mysql = "{" & myQrytable & ".EmpDept} = '" & Replace(Trim(cboList(0)), "'", "''") & _
                    "' And {" & myQrytable & ".EmpLine} = '" & Trim(cboList(1)) & "'"
    ElseIf cboList(0).Text <> "" And cboList(1).Text <> "" And cboList(2).Text <> "" Then
        mysql = "{" & myQrytable & ".EmpDept} = '" & Replace(Trim(cboList(0)), "'", "''") & _
                    "' And {" & myQrytable & ".EmpLine} = '" & Trim(cboList(1)) & _
                    "' And {" & myQrytable & ".EmpFacility} = '" & Trim(cboList(2)) & "'"
    ElseIf cboList(0).Text = "" And cboList(1).Text <> "" And cboList(2).Text = "" Then
        mysql = "{" & myQrytable & ".EmpLine} = '" & Trim(cboList(1)) & "'"
    ElseIf cboList(0).Text = "" And cboList(1).Text <> "" And cboList(2).Text <> "" Then
        mysql = "{" & myQrytable & ".EmpLine} = '" & Trim(cboList(1)) & _
                "' And {" & myQrytable & ".EmpFacility} = '" & Trim(cboList(2)) & "'"
    ElseIf cboList(0).Text <> "" And cboList(1).Text = "" And cboList(2).Text <> "" Then
        mysql = "{" & myQrytable & ".EmpDept} = '" & Replace(Trim(cboList(0)), "'", "''") & _
                "' And {" & myQrytable & ".EmpFacility} = '" & Trim(cboList(2)) & "'"
    ElseIf cboList(0).Text = "" And cboList(1).Text = "" And cboList(2).Text <> "" Then
        mysql = "{" & myQrytable & ".EmpFacility} = '" & Trim(cboList(2)) & "'"
    Else
        Exit Function
    End If
    Disply_Filter_Data = mysql
End Function





Private Sub lstVw_Click()
cmdFind.Enabled = True
End Sub
