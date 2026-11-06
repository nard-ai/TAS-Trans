Attribute VB_Name = "modMain"
Option Explicit

Public cnnAction As New ADODB.Connection
Public CON As ADODB.Connection

Public rsClientTable As ADODB.Recordset
Public rsCompanyTable As ADODB.Recordset

Public rsPayrollMast As ADODB.Recordset    ' master payroll table
Public rsPayrollTrans As ADODB.Recordset   ' payroll transaction

Public empBasicSw As Boolean
Public mClientSw As Boolean

Public mClientCode As Integer
Public mCompanyName As String
Public PayrollAlert As String
Public IsLeapYear As Long
Public Display13thPay As Double

Public mUserID As String

Public oCommonFunction As New clsCommon

'Connection of shift, overtime, and time are in the same timekeep folder .......updated by Bernard 2/20/2005.....

Public Enum myIntMode
    mNoLate = 0
    mNoDays = 1
    mNoInOut = 2
End Enum

Public Enum myIntOpt
    mInOut = 0
    mInValidShft = 1
    mXSwipe = 2
    mTardiness = 3
    mUTime = 4
    mAbsences = 5
End Enum

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Public Sub Sendkeys(Text As Variant, Optional wait As Boolean = False) ' Updated to handle SendKeys error 1/24/2020 Bernard
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(Text), wait
   Set WshShell = Nothing
End Sub

Public Function Action_Query(ByVal strSQL As String, _
                            ByVal strConString As String) As Boolean
    
    Set cnnAction = New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim il As Long
    Dim NoofRetry As Integer
    
    On Error GoTo Errh:
    
    NoofRetry = 0
    
Retry:

    cnnAction.CommandTimeout = 60   ' updated 1/21/2011

    cnnAction.Open strConString
    cnnAction.BeginTrans
    With cmd
        .ActiveConnection = cnnAction
        .CommandText = strSQL
        .CommandType = adCmdText
        .Execute il, , adCmdText + adExecuteNoRecords
    End With
    
   'Call StatusBar
    cnnAction.CommitTrans
    Set cnnAction = Nothing
    Set cmd = Nothing
    Action_Query = True
    Exit Function

Errh:
    
    If cnnAction.State = 1 Then
        cnnAction.RollbackTrans
    End If
    Set cmd = Nothing
    cnnAction.Close
    Set cnnAction = Nothing
    Action_Query = False
    
    If Err.Number = -2147217873 And Err.Description = "Duplicate key was ignored." Then
    ElseIf InStr(1, Err.Description, "Timeout expired") <> 0 Then
        If NoofRetry < 5 Then
            NoofRetry = NoofRetry + 1
            GoTo Retry
        End If
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical, "Error in Action Query"
    End If
    Err.Clear

End Function

Public Function ReturnRst(ByVal strSQL As String, _
                        ByVal strConString As String, _
                        Optional rInt As mInt = 0) As ADODB.Recordset
    Set cnnAction = New ADODB.Connection
    Dim cmdcomm As ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim NoofRetry As Integer
    
    NoofRetry = 0

    On Error GoTo Errh:

Retry:

    cnnAction.CommandTimeout = 60   ' updated 1/21/2011
    cnnAction.Open strConString
    
    cnnAction.BeginTrans
    
    Set cmdcomm = New ADODB.Command
    cmdcomm.ActiveConnection = cnnAction
    cmdcomm.CommandType = adCmdText
    cmdcomm.CommandText = strSQL
    
    rs.CursorLocation = adUseClient    'Set rs.ActiveConnection = cnnAction
    Select Case rInt
        Case mInt.mread
            Set rs = cmdcomm.Execute
        Case mInt.mstatic
            rs.Open cmdcomm, , adOpenStatic, adLockReadOnly, adCmdText
        Case mInt.mKeySet
            rs.Open cmdcomm, , adOpenKeyset, adLockReadOnly, adCmdText
        Case mInt.mOptim
            rs.Open cmdcomm, , adOpenDynamic, adLockOptimistic, adCmdText
    End Select
    
    DoEvents
    Set ReturnRst = rs
    cnnAction.CommitTrans
    Set cnnAction = Nothing
    Set rs = Nothing
    Exit Function
    
Errh:

    If cnnAction.State = 1 Then
        cnnAction.RollbackTrans
    End If
    Set rs = Nothing
    Set cmdcomm = Nothing
    Set cnnAction = Nothing
    If Err.Number = 3704 Then
    ElseIf InStr(1, Err.Description, "Timeout expired") <> 0 Then
        If NoofRetry < 5 Then
            NoofRetry = NoofRetry + 1
            GoTo Retry
        End If
    Else
        MsgBox Err.Number & " " & Err.Description, vbCritical, "Error in ReturnRst"
    End If
    Err.Clear
End Function

Public Function txtHLigths(ByVal mtxt As TextBox)
    mtxt.SelStart = 0
    mtxt.SelLength = Len(mtxt.Text)
    'mtxt.SetFocus
End Function

Public Function mskHLigths(ByVal mtxt As MaskEdBox)
    mtxt.SelStart = 0
    mtxt.SelLength = Len(mtxt.Text)
  '  mtxt.SetFocus
End Function

Public Function IsFileExists(Optional Filename As String) As Boolean
On Error GoTo Err
Dim FILE
    FILE = Dir(Filename)
        If FILE <> "" Then
                IsFileExists = True
            Else
                IsFileExists = False
        End If
        Exit Function:
Err:
        IsFileExists = False
End Function

' Returns the version of Windows that the user is running
Public Function GetWindowsVersion() As String
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                GetWindowsVersion = "Win32s on Windows 3.1"
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = "Windows NT"

                Select Case osv.dwVerMajor
                    Case 3
                        GetWindowsVersion = "Windows NT 3.5"
                    Case 4
                        GetWindowsVersion = "Windows NT 4.0"
                    Case 5
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows 2000"
                            Case 1
                                GetWindowsVersion = "Windows XP"
                            Case 2
                                GetWindowsVersion = "Windows Server 2003"
                        End Select
                    Case 6
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows Vista/Server 2008"
                            Case 1
                                GetWindowsVersion = "Windows 7/Server 2008 R2"
                            Case 2
                                GetWindowsVersion = "Windows 8/Server 2012"
                            Case 3
                                GetWindowsVersion = "Windows 8.1/Server 2012 R2"
                        End Select
                End Select

            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case osv.dwVerMinor
                    Case 0
                        GetWindowsVersion = "Windows 95"
                    Case 90
                        GetWindowsVersion = "Windows Me"
                    Case Else
                        GetWindowsVersion = "Windows 98"
                End Select
        End Select
    Else
        ' Default to Windows 10.
        GetWindowsVersion = "Windows 10"
    End If
End Function
