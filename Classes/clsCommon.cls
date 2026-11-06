VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public key As Integer
'Dim oFs As New FileSystemObject
'Dim oFile As TextStream
'Const mCellAlignment = flexAlignCenterCenter

Public WithEvents ValidCase As TextBox
Attribute ValidCase.VB_VarHelpID = -1
Public WithEvents ValidAlpha As TextBox
Attribute ValidAlpha.VB_VarHelpID = -1
Public WithEvents ValidInteger As TextBox
Attribute ValidInteger.VB_VarHelpID = -1
Public WithEvents ValidDouble As TextBox
Attribute ValidDouble.VB_VarHelpID = -1
Dim Index As Integer
Dim strValid As String

Private Sub ValidCase_KeyPress(KeyAscii As Integer)
    'strValid = "'" & Chr(13) & Chr(8)
    'If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub ValidAlpha_KeyPress(KeyAscii As Integer)
strValid = "'" & Chr(13) & Chr(8)
    If InStr(strValid, Chr(KeyAscii)) = 1 Then
    KeyAscii = 0
    Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub ValidInteger_KeyPress(KeyAscii As Integer) 'integer
    strValid = "0123456789" & Chr(13) & Chr(8)
    If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    'If KeyAscii = 13 Then SendKeys "{TAB}"     ' comment by Bernard 3/21/2010
End Sub

Private Sub ValidDouble_KeyPress(KeyAscii As Integer) 'no.
    strValid = "0123456789." & Chr(13) & Chr(8)
    If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Public Function UpdStk(ByRef TmpITCODE As String, mPlusMin As String, mVouNo As String, mQty As Double)
    Dim RsITstock As New Recordset
    Dim xtype As String
    ' TO UPDATE CURRENT STOCK
    xtype = mID(mVouNo, 1, 3)
    'Set RsITstock = oDb.RunQuery("itstock", , "ITCODE = ' " & TmpITCODE & "' AND purtype = '" & xtype & "'")
    Set RsITstock = oDb.RunQuery("ITSTOCK", , "ITCODE = '" & TmpITCODE & "' AND purtype = '" & xtype & "'")
    'MsgBox RsITstock.Fields("ITCODE")
    If RsITstock.EOF Or RsITstock.BOF Then
        If RsITstock.State = adStateOpen Then RsITstock.Close
        RsITstock.Open "ITSTOCK", oDb.db, adOpenStatic, adLockOptimistic
        RsITstock.AddNew
        RsITstock.Fields("ITCODE") = TmpITCODE
        RsITstock.Fields("purtype") = xtype
        If mPlusMin = "+" Then
            RsITstock.Fields("curstock") = Val(mQty)
        Else
            RsITstock.Fields("curstock") = -Val(mQty)
        End If
    Else
        If mPlusMin = "+" Then
            RsITstock.Fields("curstock") = Val(RsITstock.Fields("curstock")) + Val(mQty)
        Else
            RsITstock.Fields("curstock") = Val(RsITstock.Fields("curstock")) - Val(mQty)
        End If
    End If
    
    RsITstock.Update
    RsITstock.Close
End Function

Public Function Caps(ByVal KeyAscii As Integer) As Integer
    If KeyAscii >= 97 And KeyAscii <= 123 Then
        Caps = KeyAscii - 32
    Else
        Caps = KeyAscii
    End If
End Function

Public Function FillListbox(ByRef mList As ListBox, _
                    mRs As Recordset)
    mList.Clear
    While Not (mRs.EOF Or mRs.BOF)
        mList.AddItem Trim(mRs.Fields(1))
        mList.ItemData(mList.NewIndex) = mRs.Fields(0)
        mRs.MoveNext
    Wend
End Function

Public Function FillCombo(ByRef mCmb As ComboBox, _
    mRs As Recordset, Optional FieldNo As Integer = 1)
    
    mCmb.Clear
    If FieldNo = 999 Then       ' create a record signifies nothing
        mCmb.AddItem "Nothing"
        mCmb.ItemData(mCmb.NewIndex) = 0
    ElseIf FieldNo = 9 Then     ' add "All clients in the list
        mCmb.AddItem "--ALL CLIENTS--"
        FieldNo = 1
    ElseIf FieldNo = 11 Then  ' add ALL
        mCmb.AddItem "--ALL--"
        FieldNo = 1
    ElseIf FieldNo = 9999 Then ' do not add any
        FieldNo = 1
    End If
    While Not (mRs.EOF Or mRs.BOF)
        mCmb.AddItem Trim(Replace(mRs.Fields(FieldNo), "'", "''"))
        mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
        mRs.MoveNext
    Wend

End Function

Public Function FillListViv(ByRef mLsviv As ListView, _
                        mRs As Recordset, NoOfCol As Integer)
Dim litem As ListItem
Dim i As Integer

    mLsviv.View = lvwReport
    mLsviv.ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
        Set litem = mLsviv.ListItems.Add(, , mRs(0))
            For i = 1 To NoOfCol
                If IsDate(mRs(i)) = True Then
                    litem.SubItems(i) = Format(mRs(i), "dd/mm/yyyy")
                ElseIf IsNumeric(mRs(i)) = True Then
                    litem.SubItems(i) = Format(mRs(i), "##,##,##,##0.00")
                Else
                    litem.SubItems(i) = mRs(i)
                End If
            Next
        mRs.MoveNext
    Wend
        
End Function

' 0 for empty date not allowed
' 1 for empty date allowed

Public Function isNumberonly(mKeyAscii As Integer, mStatus As Integer) As Boolean
    If (mKeyAscii >= 48 And mKeyAscii <= 57) Or mKeyAscii = 46 Or mKeyAscii = 8 Or mKeyAscii = 13 Then
        isNumberonly = True
    Else
        isNumberonly = False
    End If
End Function

Public Function Word(ByVal t As Double) As String
    n = Int(t)
    R = t - n
    Const Thousand = 1000&
    Const Hundred = 100&
    Const Lakh = Thousand * Hundred
    Const Rupees = 1000000
    Const Caror = Lakh * Hundred
    Dim Buf As String: 'Buf = "Rupees "
    If (n = 0) Then Word = "Zero": Exit Function
    If (n < 0) Then Buf = "Negative ": n = Abs(n)
    
    If (n >= Caror) Then
        Buf = Buf & WordDigitGroup(n \ Caror) & " Caror"
        n = n Mod Caror
        If (n) Then Buf = Buf & " "
    End If
    
    If (n >= Rupees) Then
        Buf = Buf & WordDigitGroup(n \ Rupees) & " Million"
        n = n Mod Rupees
        If (n) Then Buf = Buf & " "
    End If
    
    If (n >= Lakh) Then
        Buf = Buf & WordDigitGroup(n \ Lakh) & " Hundred"
        n = n Mod Lakh
        If (n) Then Buf = Buf & " "
    End If
    
    If (n >= Thousand) Then
        Buf = Buf & WordDigitGroup(n \ Thousand) & " Thousand"
        n = n Mod Thousand
        If (n) Then Buf = Buf & " "
    End If
    
    If (n >= Hundred) Then
        Buf = Buf & WordDigitGroup(n \ Hundred) & " Hundred"
        n = n Mod Hundred
        If (n) Then Buf = Buf & " "
    End If
    
    If (n > 0) Then
        Buf = Buf & WordDigitGroup(n)
    End If
    If Not (R = 0) Then
        'Buf = Buf & " & Paise " & WordDigitGroup(r * 100)
        Buf = Buf & " & " & mID(Round(R, 2), 3, 2) & "/100 " & "(Php " & Format(t, "###,##0.00") & ")"
    End If
    Word = Buf '& " Only"
End Function

Private Function WordDigitGroup(ByVal n As Integer) As String
    Const Hundred = " Hundred"
    Const One = "One"
    Const Two = "Two"
    Const Three = "Three"
    Const Four = "Four"
    Const Five = "Five"
    Const Six = "Six"
    Const Seven = "Seven"
    Const Eight = "Eight"
    Const Nine = "Nine"
    Dim Buf As String: Buf = ""
    Dim Flag As Integer: Flag = False
    Select Case (n \ 100)
        Case 0: Buf = "": Flag = False
        Case 1: Buf = One & Hundred: Flag = True
        Case 2: Buf = Two & Hundred: Flag = True
        Case 3: Buf = Three & Hundred: Flag = True
        Case 4: Buf = Four & Hundred: Flag = True
        Case 5: Buf = Five & Hundred: Flag = True
        Case 6: Buf = Six & Hundred: Flag = True
        Case 7: Buf = Seven & Hundred: Flag = True
        Case 8: Buf = Eight & Hundred: Flag = True
        Case 9: Buf = Nine & Hundred: Flag = True
    End Select
    If (Flag) Then n = n Mod 100
    If (n) Then
        If (Flag) Then Buf = Buf & " "
    Else
        WordDigitGroup = Buf
        Exit Function
    End If
    Select Case (n \ 10)
        Case 0, 1: Buf = "": Flag = False
        Case 2: Buf = Buf & "Twenty": Flag = True
        Case 3: Buf = Buf & "Thirty": Flag = True
        Case 4: Buf = Buf & "Fourty": Flag = True
        Case 5: Buf = Buf & "Fifty": Flag = True
        Case 6: Buf = Buf & "Sixty": Flag = True
        Case 7: Buf = Buf & "Seventy": Flag = True
        Case 8: Buf = Buf & "Eighty": Flag = True
        Case 9: Buf = Buf & "Ninety": Flag = True
    End Select
    If (Flag) Then n = n Mod 10
    If (n) Then
        If (Flag) Then Buf = Buf & " "
    Else
        WordDigitGroup = Buf
        Exit Function
    End If
    
    Select Case (n)
        Case 0: 'do nothing
        Case 1: Buf = Buf & One
        Case 2: Buf = Buf & Two
        Case 3: Buf = Buf & Three
        Case 4: Buf = Buf & Four
        Case 5: Buf = Buf & Five
        Case 6: Buf = Buf & Six
        Case 7: Buf = Buf & Seven
        Case 8: Buf = Buf & Eight
        Case 9: Buf = Buf & Nine
        Case 10: Buf = Buf & "Ten"
        Case 11: Buf = Buf & "Eleven"
        Case 12: Buf = Buf & "Twelve"
        Case 13: Buf = Buf & "Thirteen"
        Case 14: Buf = Buf & "Fourteen"
        Case 15: Buf = Buf & "Fifteen"
        Case 16: Buf = Buf & "Sixteen"
        Case 17: Buf = Buf & "Seventeen"
        Case 18: Buf = Buf & "Eighteen"
        Case 19: Buf = Buf & "Nineteen"
    End Select
    WordDigitGroup = Buf
End Function

Function ProductDesc(code As Integer) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("ProdMast", "ProdName", "ProductID = " & code)
    If Not (rs.EOF Or rs.BOF) Then
        ProductDesc = rs.Fields("ProdName")
    Else
        MsgBox "There is no Product " & code
        ProductDesc = ""
    End If
End Function

Function GetBrCode(mName As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("BRMast", "BrCode", "BrName = '" & mName & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetBrCode = rs.Fields("BrCode")
    Else
        MsgBox "There is no Designation" & mName
        GetBrCode = ""
    End If
    rs.Close
End Function

Function GetBrName(mCode As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("BrMast", "BrName", "BrCode = '" & mCode & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetBrName = rs.Fields("BrName")
    Else
        MsgBox "There is no Designation" & mCode
        GetBrName = ""
    End If
    rs.Close
End Function

Function GetPubGCode(mName As String) As Double
    Dim rs As Recordset
    Set rs = oDb.RunQuery("PUBGRP", "GRPID", "PGNAME = '" & mName & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetPubGCode = rs.Fields("GRPID")
    Else
        MsgBox "There is no Group" & mName
        GetPubGCode = ""
    End If
    rs.Close
End Function

Function GetPubGName(mCode As Double) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("PUBGRP", "PGName", "GRPID = " & mCode)
    If Not (rs.EOF Or rs.BOF) Then
        GetPubGName = rs.Fields("PGName")
    Else
        MsgBox "There is no GROUP NAME " & mCode
        GetPubGName = ""
    End If
    rs.Close
End Function

Function GetPubCode(mName As String) As Double
    Dim rs As Recordset
    Set rs = oDb.RunQuery("PUBMAST", "PUBID", "PNAME = '" & mName & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetPubCode = rs.Fields("PUBID")
    Else
        MsgBox "There is no Group" & mName
        GetPubCode = ""
    End If
    rs.Close
End Function

Function GetPubName(mCode As Double) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("PUBMAST", "PName", "PUBID = " & mCode)
    If Not (rs.EOF Or rs.BOF) Then
        GetPubName = rs.Fields("PName")
    Else
        MsgBox "There is no GROUP NAME " & mCode
        GetPubName = ""
    End If
    rs.Close
End Function

Function GetEdition(mCode As Double) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("PUBMAST", "EDITION", "PUBID = " & mCode)
    If Not (rs.EOF Or rs.BOF) Then
        GetEdition = rs.Fields("EDITION")
    Else
        MsgBox "There is no Publication not exits " & mCode
        GetEdition = ""
    End If
    rs.Close
End Function

Function GetStateCode(mName As String) As Double
    Dim rs As Recordset
    Set rs = oDb.RunQuery("STATMAST", "STATE", "NAME = '" & mName & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetStateCode = rs.Fields("STATE")
    Else
        MsgBox "There is no Group" & mName
        GetStateCode = ""
    End If
    rs.Close
End Function

Function GetStateName(mCode As Double) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("STATMAST", "Name", "STATE = " & mCode)
    If Not (rs.EOF Or rs.BOF) Then
        GetStateName = rs.Fields("Name")
    Else
        MsgBox "There is no GROUP NAME " & mCode
        GetStateName = ""
    End If
    rs.Close
End Function

Function GetDesigId(mDsgDesc As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("Desgmst", "DsgCode", "DsgDesc = '" & mDsgDesc & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetDesigId = rs.Fields("DsgCode")
    Else
        MsgBox "There is no Designation" & mDsgDesc
        GetDesigId = ""
    End If
    rs.Close
End Function

Function GetDesigNm(mDsgCode As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("Desgmst", "DsgDesc", "DsgCode = '" & mDsgCode & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetDesigNm = rs.Fields("DsgDesc")
    Else
        MsgBox "There is no Designation" & mDsgCode
        GetDesigNm = ""
    End If
    rs.Close
End Function

Function GetItemCode(ItDesc As String) As Double
    Dim rs As Recordset
    Set rs = oDb.RunQuery("itemmst", "ITCODE", "itname = '" & ItDesc & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetItemCode = rs.Fields("ITCODE")
    Else
        MsgBox "There is no Item " & ItDesc
        GetItemCode = ""
    End If
    rs.Close
End Function

Function GetItemName(mCode As Double) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("ITEMMST", "ITNAME", "ITCODE = " & mCode)
    If Not (rs.EOF Or rs.BOF) Then
        GetItemName = rs.Fields("ITName")
    Else
        MsgBox "There is no GROUP NAME " & mCode
        GetItemName = ""
    End If
    rs.Close
End Function

Function GetCompanyName(mCompId As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("COMPMST", "cname", "compid = '" & mCompId & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetCompanyName = rs.Fields("cName")
    Else
        MsgBox "There is no Company " & mCompId
        GetCompanyName = ""
    End If
    rs.Close
End Function

Function GetCompanyId(mCName As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("COMPMST", "compid", "cname = '" & mCName & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetCompanyId = rs.Fields("compid")
    Else
        MsgBox "There is no Company " & mCName
        GetCompanyId = ""
    End If
    rs.Close
End Function

Sub enable(mFrm As Form, cond As Boolean)
    mFrm.cmdExit.Enabled = cond
    mFrm.cmdadd.Enabled = cond
    If mFrm.cmdOption1.Caption = "&Modify" Then
        mFrm.cmdOption1.Caption = "&Save"
    ElseIf mFrm.cmdOption1.Caption = "&Save" Then
        mFrm.cmdOption1.Caption = "&Modify"
    End If
    If mFrm.cmdOption2.Caption = "&Delete" Then
        mFrm.cmdOption2.Caption = "&Cancel"
    ElseIf mFrm.cmdOption2.Caption = "&Cancel" Then
        mFrm.cmdOption2.Caption = "&Delete"
    End If
End Sub

Public Function LowerCaps(str As String) As String
Dim newWord As String, newStr As String
Dim tempStr As String
Dim WDelimiter As Integer
    If Len(Trim(str)) = 0 Then
        LowerCaps = str
        Exit Function
    End If
    tempStr = LCase(Trim(str))
    WDelimiter = InStr(tempStr, " ")
    While WDelimiter > 0
        newWord = Left(tempStr, WDelimiter)
        tempStr = Right(tempStr, Len(tempStr) - WDelimiter)
        newStr = newStr & UCase(Left(newWord, 1)) & mID(newWord, 2, Len(newWord) - 1)
        WDelimiter = InStr(tempStr, " ")
    Wend
    newWord = tempStr
    newStr = newStr & UCase(Left(newWord, 1)) & mID(newWord, 2, Len(newWord) - 1)
    LowerCaps = newStr
End Function

Public Sub Upper(ByRef KeyAscii As Integer)
    KeyAscii = Asc(StrConv(Chr(KeyAscii), vbUpperCase))
End Sub

Function ValidDate(mDate As String, mPara As Integer) As String
    Dim mError As Boolean
    Dim i, Spos, Epos As Integer
    mError = False
    
    Select Case mPara
    Case 1
        If IsDate(mDate) Then
            ValidDate = Format(mDate, "dd-MM-yyyy")
        Else
            ValidDate = Format("01-01-1900", "dd-MM-yyyy")
        End If
    Case 2
        If IsDate(mDate) = True And (mDate <> "01-01-1900" And mDate <> "01/01/1900" Or mDate <> "01-01-1900") Then
            ValidDate = Format(mDate, "dd-MM-yyyy")
        Else
            ValidDate = ""
        End If
    Case 3
        If mID(Trim(mDate), 2, 1) = "." Then
            mDate = Format(mID(mDate, 1, 1), "00") + mID(mDate, 2, 9)
        End If
        If mID(Trim(mDate), 5, 1) = "." Then
            mDate = mID(mDate, 1, 3) + Format(mID(mDate, 4, 1), "00") + mID(mDate, 5, 5)
        End If
        
        If Len(mID(Trim(mDate), 7, 4)) = 2 And (Val(mID(Trim(mDate), 7, 4)) > 1 And Val(mID(Trim(mDate), 7, 4)) < 30) Then
            mDate = mID(mDate, 1, 6) + "20" + Format(mID(mDate, 7, 2), "00")
        End If
        
        If Len(mID(Trim(mDate), 7, 4)) = 2 And (Val(mID(Trim(mDate), 7, 4)) > 50 And Val(mID(Trim(mDate), 7, 4)) <= 99) Then
            mDate = mID(mDate, 1, 6) + "19" + Format(mID(mDate, 7, 2), "00")
        End If
        
        For i = 1 To Len(Trim(mDate))
            If (i <> 3 And i <> 6) And mID(Trim(mDate), i, 1) = "." Then
                mError = True
            End If
        Next
        
        If mID(Trim(mDate), 3, 1) <> "." And mID(Trim(mDate), 3, 1) <> "-" And mID(Trim(mDate), 3, 1) <> "/" And mID(Trim(mDate), 3, 1) <> "\" Then
            mError = True
        End If
        
        If (Val(mID(Trim(mDate), 1, 2)) = 0 Or Val(mID(Trim(mDate), 1, 2)) > 32) Or (Val(mID(Trim(mDate), 4, 2)) = 0 Or Val(mID(Trim(mDate), 4, 2)) > 12) Then
            mError = True
        End If
        
        If Val(mID(Trim(mDate), 7, 4)) < 1900 Then
            mError = True
        End If
        
        If mError = True Then
            If mDate <> "" Then
                MsgBox "Invalid Date"
            End If
            ValidDate = ""
            Exit Function
        Else
            'MsgBox "good" & mDate
            'ValidDate = Format(mDate, "dd-MM-yyyy")
            ValidDate = Format(mID(Trim(mDate), 1, 2), "00") & "-" & Format(mID(Trim(mDate), 4, 2), "00") & "-" & Format(mID(Trim(mDate), 7, 4), "0000")

        End If
    
    Case 4
        If mDate = "01-01-1900" Then
            ValidDate = " "
        Else
            ValidDate = Format(mDate, "dd-MM-yyyy")
        End If
    Case 5
        If mID(Trim(mDate), 2, 1) = "." Then
            mDate = Format(mID(mDate, 1, 1), "00") + mID(mDate, 2, 9)
        End If
        If mID(Trim(mDate), 5, 1) = "." Then
            mDate = mID(mDate, 1, 3) + Format(mID(mDate, 4, 1), "00") + mID(mDate, 5, 5)
        End If
        
        If Len(mID(Trim(mDate), 7, 4)) = 2 And (Val(mID(Trim(mDate), 7, 4)) > 1 And Val(mID(Trim(mDate), 7, 4)) < 30) Then
            mDate = mID(mDate, 1, 6) + "20" + Format(mID(mDate, 7, 2), "00")
        End If
        
        If Len(mID(Trim(mDate), 7, 4)) = 2 And (Val(mID(Trim(mDate), 7, 4)) > 50 And Val(mID(Trim(mDate), 7, 4)) <= 99) Then
            mDate = mID(mDate, 1, 6) + "19" + Format(mID(mDate, 7, 2), "00")
        End If
        
        For i = 1 To Len(Trim(mDate))
            If (i <> 3 And i <> 6) And mID(Trim(mDate), i, 1) = "." Then
                mError = True
            End If
        Next
        
        If mID(Trim(mDate), 3, 1) <> "." And mID(Trim(mDate), 3, 1) <> "-" And mID(Trim(mDate), 3, 1) <> "/" And mID(Trim(mDate), 3, 1) <> "\" Then
            mError = True
        End If
        
        If (Val(mID(Trim(mDate), 1, 2)) = 0 Or Val(mID(Trim(mDate), 1, 2)) > 32) Or (Val(mID(Trim(mDate), 4, 2)) = 0 Or Val(mID(Trim(mDate), 4, 2)) > 12) Then
            mError = True
        End If
        
        If Val(mID(Trim(mDate), 7, 4)) < 1900 Then
            mError = True
        End If
        
        If mError = True Then
            If mDate <> "" Then
                MsgBox "Invalid Date"
            End If
            ValidDate = ""
            Exit Function
        Else
            'MsgBox "good" & mDate
            'ValidDate = Format(mDate, "dd-MM-yyyy")
            ValidDate = Format(mID(Trim(mDate), 4, 2), "00") & "-" & Format(mID(Trim(mDate), 1, 2), "00") & "-" & Format(mID(Trim(mDate), 7, 4), "0000")

        End If
    
    End Select
End Function

Function ValidText(mText As String) As String
    If Len(Trim(mText)) = 0 Or IsNull(mText) = True Then
        ValidText = " "
    Else
        ValidText = Trim(mText)
    End If
    
End Function

Function ValidValue(mText As String) As Variant
    If Len(Trim(mText)) = 0 Then
        ValidValue = 0
    Else
        ValidValue = Val(mText)
    End If
    
End Function

Function GetAcCode(AcDesc As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("ACMAST", "AcCode", "ACNAME = '" & AcDesc & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetAcCode = rs.Fields("AcCode")
    Else
        MsgBox "There is no Account" & AcDesc
        GetAcCode = ""
    End If
    rs.Close
End Function

Function GetMemNobyAcCode(mAcCode As Integer) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("ACMAST", "MEMNO", "ACCODE = " & mAcCode)
    If Not (rs.EOF Or rs.BOF) Then
        GetMemNobyAcCode = rs.Fields("MEMNO")
    Else
        MsgBox "There is no Account" & mAcCode
        GetMemNobyAcCode = ""
    End If
    rs.Close
End Function

Function GetEmpIdbyAcCode(mAcCode As Integer) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("ACMAST", "EMPID", "ACCODE = " & mAcCode)
    If Not (rs.EOF Or rs.BOF) Then
        GetEmpIdbyAcCode = rs.Fields("EMPID")
    Else
        MsgBox "There is no Account" & mAcCode
        GetEmpIdbyAcCode = ""
    End If
    rs.Close
End Function

Function GetMemName(mMemNo As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("ACMAST", "ACNAME", "MEMNO = '" & mMemNo & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetMemName = rs.Fields("ACName")
    Else
        MsgBox "There is no Account" & mMemNo
        GetMemName = ""
    End If
    rs.Close
End Function

Function GetMemNobyEmpId(mEmpId As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("ACMAST", "MEMNO", "EMPID = '" & mEmpId & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetMemNobyEmpId = rs.Fields("MEMNO")
    Else
        MsgBox "There is no Account" & mEmpId
        GetMemNobyEmpId = ""
    End If
    rs.Close
End Function

Function GetAcName(mAcCode As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("ACMAST", "ACNAME", "AcCode = " & mAcCode)
    If Not (rs.EOF Or rs.BOF) Then
        GetAcName = rs.Fields("ACName")
    Else
        MsgBox "There is no Account" & mAcCode
        GetAcName = ""
    End If
    rs.Close
End Function

Function GetAcCodeByEmpid(mEmpId As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("ACMAST", "AcCode", "EMPID = '" & mEmpId & "'")
    
    If Not (rs.EOF Or rs.BOF) Then
        GetAcCodeByEmpid = rs.Fields("AcCode")
    Else
        'MsgBox "There is no Account" & AcDesc
        GetAcCodeByEmpid = 0
    End If
    rs.Close
End Function

Function Getclientcode(DeptDesc As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("DeptMst", "clientcode", "Dept = '" & DeptDesc & "'")
    If Not (rs.EOF Or rs.BOF) Then
        Getclientcode = rs.Fields("clientcode")
    Else
        MsgBox "There is no Account" & DeptDesc
        Getclientcode = ""
    End If
    rs.Close
End Function

Function GetDeptName(mClientCode As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("DEPTMST", "DEPT", "clientcode = '" & mClientCode & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetDeptName = rs.Fields("Dept")
    Else
        MsgBox "There is no Account" & mClientCode
        GetDeptName = ""
    End If
    rs.Close
End Function

Function GetDedCode(Dname As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("DedMast", "DedCode", "Dname = '" & Dname & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetDedCode = rs.Fields("DedCode")
    Else
        MsgBox "There is no Account" & Dname
        GetDedCode = ""
    End If
    rs.Close
End Function

Function GetDedName(mDedCode As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("DEDMAST", "DNAME", "DEDCode = '" & mDedCode & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetDedName = rs.Fields("DNAME")
    Else
        MsgBox "There is no Account" & mDedCode
        GetDedName = ""
    End If
    rs.Close
End Function

'Sub PrintTitle(ByRef flex As VSFlexGrid, mText As String, mRow As Integer, mCol As Integer)
'    With flex
'        .TextMatrix(mRow, mCol) = mText
'        .Select mRow, mCol
'        .CellBackColor = vbBlack
'        .CellForeColor = vbBlue
'        '.CellFontName = mFontName
'        '.CellFontSize = mFontSize
'        .CellAlignment = mCellAlignment
'    End With
'End Sub

Function RoDet(mRoNo As Double) As Recordset
    'Dim rs As Recordset
    Set RoDet = oDb.RunQuery("ROMAST", , "RONO = " & mRoNo)
End Function

Function GetProfTax(mProfApOn As Double, mCompId As String) As Double
    Dim rs As Recordset
    Dim mstateid%
    mProfApOn = mProfApOn * 12
    If mCompId <> "04" Then
        mstateid = "0"
        'Set rs = oDb.RunQuery("PROFTAX", , mProfApOn & " >= STAMT And " & mProfApOn & " <= ENDAMT AND STATEID = '" & mstateid & "'")
        Set rs = oDb.RunQuery("PROFTAX", , "STATEID = '" & mstateid & "' AND " & mProfApOn & " >= STAMT And " & mProfApOn & " <= ENDAMT")
    Else
        mstateid = "1"
        'Set rs = oDb.RunQuery("PROFTAX", , mProfApOn & " >= STAMT And " & mProfApOn & " <= ENDAMT AND STATEID = '" & mstateid & "'")
        Set rs = oDb.RunQuery("PROFTAX", , "STATEID = '" & mstateid & "' AND " & mProfApOn & " >= STAMT And " & mProfApOn & " <= ENDAMT")
    End If
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        If mProfApOn >= rs.Fields("STAMT") And mProfApOn <= rs.Fields("ENDAMT") Then
            GetProfTax = rs.Fields("PRFAMT")
        Else
            GetProfTax = 0
        End If
    Else
        GetProfTax = 0
    End If
    rs.Close
End Function

Function GetDataBaseName(mCompNm As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("COMPMST", , "CNAME = '" & mCompNm & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetDataBaseName = rs.Fields("DATABASENM")
    Else
        MsgBox "There is no Account" & mCompNm
        GetDataBaseName = ""
    End If
    rs.Close
End Function

Function GetEmpName(mEmpId As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("EMPMAST", "EMPNAME", "EMPID = '" & mEmpId & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetEmpName = rs.Fields("EMPNAME")
    Else
        MsgBox "There is no Employee " & mEmpId
        GetEmpName = ""
    End If
    rs.Close
End Function

Function GetEmpId(mEmpName As String) As String
    Dim rs As Recordset
    Set rs = oDb.RunQuery("EMPMAST", "EMPID", "EMPNAME = '" & mEmpName & "'")
    If Not (rs.EOF Or rs.BOF) Then
        GetEmpId = rs.Fields("EMPID")
    Else
        MsgBox "There is no Company " & mEmpName
        GetEmpId = ""
    End If
    rs.Close
End Function

Function GetOpLeaves(mEmpId As String, mOldDate As Date) As Double
    Dim rs As Recordset
    Set rs = oDb.RunQuery("SALTRAN", "BALLV", "EMPID = '" & mEmpId & "' AND  SALDT = #" & mOldDate & "#")
    If Not (rs.EOF Or rs.BOF) Then
        GetOpLeaves = rs.Fields("BALLV")
    Else
        GetOpLeaves = 0
    End If
    rs.Close
End Function

Function GetDays(mDate As Date) As Integer
    Select Case mID(mDate, 4, 2)
        Case "01"
            GetDays = 31
        Case "02"
            GetDays = 28
        Case "03"
            GetDays = 31
        Case "04"
            GetDays = 30
        Case "05"
            GetDays = 31
        Case "06"
            GetDays = 30
        Case "07"
            GetDays = 31
        Case "08"
            GetDays = 31
        Case "09"
            GetDays = 30
        Case "10"
            GetDays = 31
        Case "11"
            GetDays = 30
        Case "12"
            GetDays = 31
    End Select
End Function

Public Sub Flex_Renum(ByVal Col As Long, Order As Integer, mForm As Form)
    Dim i%
    For i = 1 To mForm.FgIns.Rows - 1
        mForm.FgIns.TextMatrix(i, 0) = i
    Next
End Sub

Public Sub PrintReport(id As Long, noCopy As Integer)
    Dim mTotamt  As Double
    Dim Repopath As String
    Dim comd As String
    Dim mdedname As String
    Dim mLen%, lCtr%, i%
    
    Repopath = "c:\test.txt"
    Set oFile = oFs.CreateTextFile(Repopath, True)
    
    Dim mRs As ADODB.Recordset
    Dim mRs2 As New ADODB.Recordset
    Set mRs = oDb.RunQuery("SALARYSLP", "EMPID", , "EMPID", "EMPID")
    If mRs.EOF() = False Then
        With oFile
            While Not mRs.EOF = True
            
                If mRs2.State = adStateOpen Then mRs2.Close
                Set mRs2 = oDb.RunQuery("SALARYSLP", , "EMPID = '" & mRs.Fields("EMPID") & "'", "EMPID")
                .WriteBlankLines (1)
                .WriteLine Chr(14) & Space(20) & "DECCAN CREDIT CO.OPERATIVE LTD."
                .WriteLine Space(7) & "20-22, ANOOP NAGAR, A.B.ROAD,INDORE - 452 008. Ph.:(0731) 2559095"
                .WriteBlankLines (1)
                
                .Write "Receipt No.  :" & mRs2.Fields("VOUNO")
                .Write Space(17) & "CASH/BANK RECEIPT"
                .Write Space(15)
                .WriteLine "Date :" & Format(mRs2.Fields("VOUDT"), "dd-MM-yyyy")
                .WriteBlankLines (1)
                .WriteLine "Membership No.:" & mRs2.Fields("MEMNO")
                .WriteLine "Received with thanks from :" & oCm.GetEmpName(mRs2.Fields("EMPID"))
                mTotamt = 0
                While Not mRs2.EOF = True
                    mTotamt = mTotamt + mRs2.Fields("AMOUNT")
                    mRs2.MoveNext
                Wend
                
                .WriteLine "Rs.:" & Format(mTotamt, "#####0.00") + " In words " & Trim(oCm.Word(mTotamt))
                .WriteLine "Cash/Cheque No. ..................... Date ............... Bank.............."
                .WriteBlankLines (1)
                
                .WriteLine "--------------------------------------------------------------------------------"
                .WriteLine "  Particulars                                                  Amount           "
                .WriteLine "--------------------------------------------------------------------------------"
                mdedname = ""
                mLen = 0
                lCtr = 0
                mRs2.MoveFirst
                While Not mRs2.EOF = True
                    mdedname = Trim(mRs2.Fields("LEDNARR"))
                    mLen = Len(mdedname)
                    lCtr = lCtr + 1
                    .Write Format(lCtr, "##0") & Space(2) & mdedname & Space(70 - (mLen + 5))
                    .WriteLine Format(mRs2.Fields("AMOUNT"), "#####0.00")
                    mRs2.MoveNext
                Wend
                i = 1
                lCtr = lCtr + 13
                If lCtr < 20 Then
                    For i = lCtr To 20
                        .WriteLine ""
                    Next
                End If
                
                .WriteLine "--------------------------------------------------------------------------------"
                .Write Space(10) & "TOTAL :"
                .Write Space(50)
                .WriteLine Format(mTotamt, "#####0.00")
                .WriteLine "--------------------------------------------------------------------------------"
                
                .WriteLine "NOTE :"
                .WriteLine "     1)Cheques/DDs received are subject to realisation."
                .WriteLine "     2)Subject to indore Jurisdiction."
                .WriteLine Space(45) & "For Deccan Credit Co-Operative Ltd."
                .WriteBlankLines (3)
                .WriteLine Space(60) & "Authorised Signatory"
                mRs.MoveNext
            Wend
            mRs.Close
            mRs2.Close
            .Close
        End With
        comd = "COMMAND.COM /c edit " & Repopath
        Shell comd
    End If
End Sub

Public Function EsicRound(mAmount As Double) As Double
Dim mPre, mPost, mPost1, mAddamt As Double
' ESIC Round
    'MsgBox Len(Format(Val(mAmount), "#.00"))
    mAmount = Val(Format(Val(mAmount), "#.00"))
    mPost = Val(mID(Format(Val(mAmount), "#.00"), Len(Format(Val(mAmount), "#.00")) - 1, 2))
    mPre = Val(mID(Format(Val(mAmount), "#.00"), 1, Len(Format(Val(mAmount), "#.00")) - 3))
    
    If mPost > 0 And mPost <= 5 Then
        mAddamt = 0.05
    ElseIf mPost >= 6 And mPost <= 10 Then
        mAddamt = 0.1
    
    ElseIf mPost > 10 And mPost <= 15 Then
        mAddamt = 0.15
    ElseIf mPost >= 16 And mPost <= 20 Then
        mAddamt = 0.2
    
    ElseIf mPost > 20 And mPost <= 25 Then
        mAddamt = 0.25
    ElseIf mPost >= 26 And mPost <= 30 Then
        mAddamt = 0.3
    
    ElseIf mPost > 30 And mPost <= 35 Then
        mAddamt = 0.35
    ElseIf mPost >= 36 And mPost <= 40 Then
        mAddamt = 0.4
    
    ElseIf mPost > 40 And mPost <= 45 Then
        mAddamt = 0.45
    ElseIf mPost >= 46 And mPost <= 50 Then
        mAddamt = 0.5
    
    ElseIf mPost > 50 And mPost <= 55 Then
        mAddamt = 0.55
    ElseIf mPost >= 56 And mPost <= 60 Then
        mAddamt = 0.6
    
    ElseIf mPost > 60 And mPost <= 65 Then
        mAddamt = 0.65
    ElseIf mPost >= 66 And mPost <= 70 Then
        mAddamt = 0.7
    
    ElseIf mPost > 70 And mPost <= 75 Then
        mAddamt = 0.75
    ElseIf mPost >= 66 And mPost <= 80 Then
        mAddamt = 0.8
    
    ElseIf mPost > 80 And mPost <= 85 Then
        mAddamt = 0.85
    ElseIf mPost >= 86 And mPost <= 90 Then
        mAddamt = 0.9
    
    ElseIf mPost > 90 And mPost <= 95 Then
        mAddamt = 0.95
    ElseIf mPost >= 96 And mPost <= 99 Then
        mAddamt = 1
    End If
    
    'MsgBox mPost & " ** " & mPost1
    'MsgBox mPre & " - " & mPost & " - " & mAddamt & " R---> " & mPre + mAddamt
    EsicRound = Val(mPre + mAddamt)
End Function

Public Function PFRound(mAmount As Double) As Double
Dim mPre, mPost, mAddamt As Double
' For PF Round Amount
    'MsgBox Len(Format(Val(mAmount), "#.00"))
    mAmount = Val(Format(Val(mAmount), "#.00"))
    mPost = Val(mID(Format(Val(mAmount), "#.00"), Len(Format(Val(mAmount), "#.00")) - 1, 2))
    mPre = Val(mID(Format(Val(mAmount), "#.00"), 1, Len(Format(Val(mAmount), "#.00")) - 3))
    If mPost > 0 And mPost <= 50 Then
        mAddamt = 0.5
    ElseIf mPost >= 51 And mPost <= 99 Then
        mAddamt = 1
    End If
    'MsgBox mPost
    'MsgBox mPre & " - " & mPost & " - " & mAddamt & " R---> " & mPre + mAddamt
    PFRound = Val(mPre + mAddamt)
End Function


'::--------------------------------------------------------------::
'::         Used to Check if the year is a leap year             ::
'::--------------------------------------------------------------::
Function IsLeapYear(ByVal sYear As String) As Boolean
   If IsDate("02/29/" & sYear) Then
      IsLeapYear = True
   Else
      IsLeapYear = False
   End If
End Function

'::--------------------------------------------------------------::
'::               Used To Proper Case a String                   ::
'::--------------------------------------------------------------::
Public Function ProperCase(ByVal vData As String) As String
   Dim i As Integer   'Used as a counter

   vData = Trim$(vData)   'Trim The String

   If Len(vData) < 1 Then   'Check the length
      ProperCase = vData
      Exit Function
   End If

   vData = UCase$(Left$(vData, 1)) & LCase$(Right$(vData, Len(vData) - 1))

   For i = 2 To Len(vData)
      If (mID$(vData, i, 1) = " ") Or (mID$(vData, i, 1) = "-") And (i + 1 <= Len(vData)) Then
         Mid$(vData, i + 1, 1) = UCase$(mID$(vData, i + 1, 1))
      End If
   Next i

   ProperCase = vData
End Function
'::--------------------------------------------------------------::

