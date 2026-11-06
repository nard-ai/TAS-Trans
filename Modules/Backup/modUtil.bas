Attribute VB_Name = "XmodUtil"
Option Explicit

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     lParam As Any) As Long
     
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpDefault As String, _
     ByVal lpReturnedString As String, _
     ByVal lSize As Long, _
     ByVal lpFileName As String) As Long

Public Const CB_ERR = -1
Public Const CB_Findstring = &H14C
Public Const LB_FINDSTRING = &H18F
Public Const SUCCEED& = 1
Public Const LB_GETSELITEMS = &H191

'Public Function Code(ByVal s As String) As String
'    clsCipher.Text = s
'    clsCipher.Encrypt
'    Code = clsCipher.Text
'End Function

'Public Function Decode(ByVal s As String) As String
'    clsCipher.Text = s
'    clsCipher.Decrypt
'    Decode = clsCipher.Text
'End Function

Public Function GETINI(ByVal lpApplicationName As String, _
                       ByVal lpKeyName As String, _
                       ByVal lpDefault As String, _
                       ByVal lpFileName As String)

Dim lngRet         As Long          'return value
Dim strRetBuffer   As String * 150   'buffer
    
   lngRet = GetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, strRetBuffer, Len(strRetBuffer), lpFileName)
   GETINI = Left(strRetBuffer, InStr(strRetBuffer, Chr(0)) - 1)
End Function

Public Function IsDateValid(Obj As Object) As Boolean
   If Not IsDate(Obj.Text) Then
        MsgBox "Invalid Date!", vbExclamation, "MESSAGE"
        Obj.SetFocus
        IsDateValid = False
        Exit Function
    Else
        IsDateValid = True
    End If
End Function

Public Function IsObjectNull(Obj As Object, Title As String) As Boolean
    If Trim$(Obj.Text) = "" Then
        MsgBox Title & " should not be blank!", vbExclamation, "MESSAGE"
        Obj.SetFocus
        Exit Function
        IsObjectNull = False
    Else
        IsObjectNull = True
    End If
End Function

Sub FindMatch(cbo As ComboBox, KeyAscii As Integer)

Dim strBuffer      As String  'buffer
Dim lngRetVal      As Long    'return value
  
  If cbo.Text <> "" Then strBuffer = Left(cbo.Text, cbo.SelStart) & Chr(KeyAscii)
  lngRetVal = SendMessage((cbo.hWnd), CB_Findstring, -1, ByVal strBuffer)
  If lngRetVal <> CB_ERR Then
     cbo.ListIndex = lngRetVal
     cbo.Text = cbo.list(lngRetVal)
     cbo.SelStart = Len(strBuffer)
     cbo.SelLength = Len(cbo.Text)
     KeyAscii = 0
  End If
End Sub

Public Sub onADOerr()
Dim strErr As String

'    If cnPrms.Errors.Count > 0 Then
'        For Each dbErr In cnPrms.Errors
'            strErr = strErr & "Error #" & dbErr.Number & _
'                " " & dbErr.Description & vbCrLf & _
'                "NativeError: " & dbErr.NativeError & vbCrLf & _
'                "SQLState: " & dbErr.SQLState & vbCrLf & _
'                "Reported by: " & dbErr.Source & vbCrLf & _
'                "Help file: " & dbErr.HelpFile & vbCrLf & _
'                "Help Context ID: " & dbErr.HelpContext
'        Next
'        MsgBox strErr
'        cnPrms.Errors.Clear
'    End If
End Sub


Public Sub GetAllSelectedItem(LBox As ListBox, ItemIndex() As Long, iNumItems%)
Dim i%
    iNumItems = LBox.SelCount
    If iNumItems Then
        ReDim ItemIndex(iNumItems - 1)
        SendMessage LBox.hWnd, LB_GETSELITEMS, iNumItems, ItemIndex(0)
    End If
End Sub
