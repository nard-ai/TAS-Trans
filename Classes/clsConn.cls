VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum mInt
    mread = 0
    mstatic = 1
    mOptim = 2
    mKeySet = 3
End Enum

Public Function Action_Query(ByVal strSQL As String, _
                            ByVal strConString As String, _
                            Optional sqlchk As Boolean = False) As Boolean
                            
    Set cnn = New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim il As Long
    
    On Error GoTo Errh:
    cnn.Open strConString
    cnn.BeginTrans
    If sqlchk Then MsgBox strSQL, vbExclamation, "SQL Statement" 'check the sql syntax////........."
    With cmd
        .ActiveConnection = cnn
        .CommandText = strSQL
        .CommandType = adCmdText
        .Execute il, , adCmdText + adExecuteNoRecords
    End With
    cnn.CommitTrans
    Set cnn = Nothing
    Set cmd = Nothing
    Action_Query = True
    Exit Function
Errh:
    cnn.RollbackTrans
    Set cmd = Nothing
    cnn.Close
    Set cnn = Nothing
    Action_Query = False
    MsgBox Err.Number & " " & Err.Description, vbCritical, "Error in Action Query"
    Err.Clear
End Function

Public Function ReturnRst(ByVal strSQL As String, _
                        ByVal strConString As String, _
                        Optional rInt As mInt = 0, _
                        Optional sqlchk As Boolean = False) As ADODB.Recordset
                        
    Set cnn = New ADODB.Connection
    Dim cmdcomm As ADODB.Command
    Dim rs As New ADODB.Recordset

    On Error GoTo Errh:
    cnn.Open strConString
    
    cnn.BeginTrans
    Set cmdcomm = New ADODB.Command
    cmdcomm.ActiveConnection = cnn
    cmdcomm.CommandType = adCmdText
    cmdcomm.CommandText = strSQL
    If sqlchk Then MsgBox strSQL, vbExclamation, "SQL Statement" 'check the sql syntax////........."
    
    rs.CursorLocation = adUseClient    'Set rs.ActiveConnection = cnn
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
    
    cnn.CommitTrans
    Set ReturnRst = rs
    Set cnn = Nothing
    Set cmdcomm = Nothing
    Set rs = Nothing
    Exit Function
    
Errh:
    cnn.RollbackTrans
    Set cmdcomm = Nothing
    Set rs = Nothing
    rs.Close
    Set cnn = Nothing
    cnn.Close
    MsgBox Err.Number & " " & Err.Description, vbCritical, "Error in ReturnRst"
    Err.Clear
End Function




