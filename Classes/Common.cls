VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public key As Integer
'Dim oFs As New FileSystemObject
'Dim oFile As TextStream
'Const mCellAlignment = flexAlignCenterCenter

Public Function RunQuery(ByVal TableName As String, _
                    Optional ByVal ColumnName As String, _
                    Optional ByVal where As String, _
                    Optional ByVal Order As String, _
                    Optional ByVal Groupby As String) As Recordset
    Dim sql As String
    If ColumnName <> "" Then
        sql = "Select " & ColumnName & " from "
    Else
        sql = "Select * from "
    End If
    sql = sql & TableName
    If where <> "" Then
        sql = sql & " where " & where
    End If
    
    If Groupby <> "" Then
        sql = sql & " GROUP BY " & Groupby
    End If
    
    If Order <> "" Then
        sql = sql & " Order by " & Order
    End If
    Set RunQuery = Db.Execute(sql)
End Function

Public Function UpdStk(ByRef TmpITCODE As String, mPlusMin As String, mVouNo As String, mQty As Double)
    Dim RsITstock As New Recordset
    Dim xtype As String
    ' TO UPDATE CURRENT STOCK
    xtype = Mid(mVouNo, 1, 3)
    'Set RsITstock = oDb.RunQuery("itstock", , "ITCODE = ' " & TmpITCODE & "' AND purtype = '" & xtype & "'")
    Set RsITstock = oDb.RunQuery("ITSTOCK", , "ITCODE = '" & TmpITCODE & "' AND purtype = '" & xtype & "'")
    'MsgBox RsITstock.Fields("ITCODE")
    If RsITstock.EOF Or RsITstock.BOF Then
        If RsITstock.State = adStateOpen Then RsITstock.Close
        RsITstock.Open "ITSTOCK", oDb.Db, adOpenStatic, adLockOptimistic
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
                        mRs As Recordset)
        mCmb.Clear
        While Not (mRs.EOF Or mRs.BOF)
            mCmb.AddItem Trim(mRs.Fields(1))
            mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
            mRs.MoveNext
        Wend
        'mCmb.AddItem ""
        'mCmb.ItemData(mCmb.NewIndex) = 0
End Function

Public Function FillListViv(ByRef mLsviv As ListView, _
                        mRs As Recordset, NoOfCol As Integer)
Dim litem As ListItem
    mLsviv.View = lvwReport
    mLsviv.ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
        Set litem = mLsviv.ListItems.add(, , mRs(0))
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

Function Word(ByVal T As Double) As String
    n = Int(T)
    r = T - n
    Const Thousand = 1000&
    Const Hundred = 100&
    Const Lakh = Thousand * Hundred
    Const Caror = Lakh * Hundred
    Dim Buf As String: Buf = "Rupees "
    If (n = 0) Then Word = "Zero": Exit Function
    If (n < 0) Then Buf = "Negative ": n = Abs(n)
    If (n >= Caror) Then
        Buf = Buf & WordDigitGroup(n \ Caror) & " Caror"
        n = n Mod Caror
        If (n) Then Buf = Buf & " "
    End If
    
    If (n >= Lakh) Then
        Buf = Buf & WordDigitGroup(n \ Lakh) & " Lakh"
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
    If Not (r = 0) Then
        Buf = Buf & " & Paise " & WordDigitGroup(r * 100)
    End If
    Word = Buf & " Only"
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
        Case 4: Buf = Buf & "Forty": Flag = True
        Case 5: Buf = Buf & "Fifty": Flag = True
        Case 6: Buf = Buf & "Sixty": Flag = True
        Case 7: Buf = Buf & "Seventy": Flag = True
        Case 8: Buf = Buf & "Eighty": Flag = True
        Case 9: Buf = Buf & "Ninty": Flag = True
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
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("ProdMast", "ProdName", "ProductID = " & code)
    If Not (Rs.EOF Or Rs.BOF) Then
        ProductDesc = Rs.Fields("ProdName")
    Else
        MsgBox "There is no Product " & code
        ProductDesc = ""
    End If
End Function

Function GetBrCode(mName As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("BRMast", "BrCode", "BrName = '" & mName & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetBrCode = Rs.Fields("BrCode")
    Else
        MsgBox "There is no Designation" & mName
        GetBrCode = ""
    End If
    Rs.Close
End Function

Function GetBrName(mCode As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("BrMast", "BrName", "BrCode = '" & mCode & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetBrName = Rs.Fields("BrName")
    Else
        MsgBox "There is no Designation" & mCode
        GetBrName = ""
    End If
    Rs.Close
End Function

Function GetPubGCode(mName As String) As Double
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("PUBGRP", "GRPID", "PGNAME = '" & mName & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetPubGCode = Rs.Fields("GRPID")
    Else
        MsgBox "There is no Group" & mName
        GetPubGCode = ""
    End If
    Rs.Close
End Function

Function GetPubGName(mCode As Double) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("PUBGRP", "PGName", "GRPID = " & mCode)
    If Not (Rs.EOF Or Rs.BOF) Then
        GetPubGName = Rs.Fields("PGName")
    Else
        MsgBox "There is no GROUP NAME " & mCode
        GetPubGName = ""
    End If
    Rs.Close
End Function

Function GetPubCode(mName As String) As Double
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("PUBMAST", "PUBID", "PNAME = '" & mName & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetPubCode = Rs.Fields("PUBID")
    Else
        MsgBox "There is no Group" & mName
        GetPubCode = ""
    End If
    Rs.Close
End Function

Function GetPubName(mCode As Double) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("PUBMAST", "PName", "PUBID = " & mCode)
    If Not (Rs.EOF Or Rs.BOF) Then
        GetPubName = Rs.Fields("PName")
    Else
        MsgBox "There is no GROUP NAME " & mCode
        GetPubName = ""
    End If
    Rs.Close
End Function

Function GetEdition(mCode As Double) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("PUBMAST", "EDITION", "PUBID = " & mCode)
    If Not (Rs.EOF Or Rs.BOF) Then
        GetEdition = Rs.Fields("EDITION")
    Else
        MsgBox "There is no Publication not exits " & mCode
        GetEdition = ""
    End If
    Rs.Close
End Function

Function GetStateCode(mName As String) As Double
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("STATMAST", "STATE", "NAME = '" & mName & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetStateCode = Rs.Fields("STATE")
    Else
        MsgBox "There is no Group" & mName
        GetStateCode = ""
    End If
    Rs.Close
End Function

Function GetStateName(mCode As Double) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("STATMAST", "Name", "STATE = " & mCode)
    If Not (Rs.EOF Or Rs.BOF) Then
        GetStateName = Rs.Fields("Name")
    Else
        MsgBox "There is no GROUP NAME " & mCode
        GetStateName = ""
    End If
    Rs.Close
End Function

Function GetDesigId(mDsgDesc As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("Desgmst", "DsgCode", "DsgDesc = '" & mDsgDesc & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetDesigId = Rs.Fields("DsgCode")
    Else
        MsgBox "There is no Designation" & mDsgDesc
        GetDesigId = ""
    End If
    Rs.Close
End Function

Function GetDesigNm(mDsgCode As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("Desgmst", "DsgDesc", "DsgCode = '" & mDsgCode & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetDesigNm = Rs.Fields("DsgDesc")
    Else
        MsgBox "There is no Designation" & mDsgCode
        GetDesigNm = ""
    End If
    Rs.Close
End Function

Function GetItemCode(ItDesc As String) As Double
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("itemmst", "ITCODE", "itname = '" & ItDesc & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetItemCode = Rs.Fields("ITCODE")
    Else
        MsgBox "There is no Item " & ItDesc
        GetItemCode = ""
    End If
    Rs.Close
End Function

Function GetItemName(mCode As Double) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("ITEMMST", "ITNAME", "ITCODE = " & mCode)
    If Not (Rs.EOF Or Rs.BOF) Then
        GetItemName = Rs.Fields("ITName")
    Else
        MsgBox "There is no GROUP NAME " & mCode
        GetItemName = ""
    End If
    Rs.Close
End Function

Function GetCompanyName(mCompId As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("COMPMST", "cname", "compid = '" & mCompId & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetCompanyName = Rs.Fields("cName")
    Else
        MsgBox "There is no Company " & mCompId
        GetCompanyName = ""
    End If
    Rs.Close
End Function

Function GetCompanyId(mCName As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("COMPMST", "compid", "cname = '" & mCName & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetCompanyId = Rs.Fields("compid")
    Else
        MsgBox "There is no Company " & mCName
        GetCompanyId = ""
    End If
    Rs.Close
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
        newStr = newStr & UCase(Left(newWord, 1)) & Mid(newWord, 2, Len(newWord) - 1)
        WDelimiter = InStr(tempStr, " ")
    Wend
    newWord = tempStr
    newStr = newStr & UCase(Left(newWord, 1)) & Mid(newWord, 2, Len(newWord) - 1)
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
        If Mid(Trim(mDate), 2, 1) = "." Then
            mDate = Format(Mid(mDate, 1, 1), "00") + Mid(mDate, 2, 9)
        End If
        If Mid(Trim(mDate), 5, 1) = "." Then
            mDate = Mid(mDate, 1, 3) + Format(Mid(mDate, 4, 1), "00") + Mid(mDate, 5, 5)
        End If
        
        If Len(Mid(Trim(mDate), 7, 4)) = 2 And (Val(Mid(Trim(mDate), 7, 4)) > 1 And Val(Mid(Trim(mDate), 7, 4)) < 30) Then
            mDate = Mid(mDate, 1, 6) + "20" + Format(Mid(mDate, 7, 2), "00")
        End If
        
        If Len(Mid(Trim(mDate), 7, 4)) = 2 And (Val(Mid(Trim(mDate), 7, 4)) > 50 And Val(Mid(Trim(mDate), 7, 4)) <= 99) Then
            mDate = Mid(mDate, 1, 6) + "19" + Format(Mid(mDate, 7, 2), "00")
        End If
        
        For i = 1 To Len(Trim(mDate))
            If (i <> 3 And i <> 6) And Mid(Trim(mDate), i, 1) = "." Then
                mError = True
            End If
        Next
        
        If Mid(Trim(mDate), 3, 1) <> "." And Mid(Trim(mDate), 3, 1) <> "-" And Mid(Trim(mDate), 3, 1) <> "/" And Mid(Trim(mDate), 3, 1) <> "\" Then
            mError = True
        End If
        
        If (Val(Mid(Trim(mDate), 1, 2)) = 0 Or Val(Mid(Trim(mDate), 1, 2)) > 32) Or (Val(Mid(Trim(mDate), 4, 2)) = 0 Or Val(Mid(Trim(mDate), 4, 2)) > 12) Then
            mError = True
        End If
        
        If Val(Mid(Trim(mDate), 7, 4)) < 1900 Then
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
            ValidDate = Format(Mid(Trim(mDate), 1, 2), "00") & "-" & Format(Mid(Trim(mDate), 4, 2), "00") & "-" & Format(Mid(Trim(mDate), 7, 4), "0000")

        End If
    
    Case 4
        If mDate = "01-01-1900" Then
            ValidDate = " "
        Else
            ValidDate = Format(mDate, "dd-MM-yyyy")
        End If
    Case 5
        If Mid(Trim(mDate), 2, 1) = "." Then
            mDate = Format(Mid(mDate, 1, 1), "00") + Mid(mDate, 2, 9)
        End If
        If Mid(Trim(mDate), 5, 1) = "." Then
            mDate = Mid(mDate, 1, 3) + Format(Mid(mDate, 4, 1), "00") + Mid(mDate, 5, 5)
        End If
        
        If Len(Mid(Trim(mDate), 7, 4)) = 2 And (Val(Mid(Trim(mDate), 7, 4)) > 1 And Val(Mid(Trim(mDate), 7, 4)) < 30) Then
            mDate = Mid(mDate, 1, 6) + "20" + Format(Mid(mDate, 7, 2), "00")
        End If
        
        If Len(Mid(Trim(mDate), 7, 4)) = 2 And (Val(Mid(Trim(mDate), 7, 4)) > 50 And Val(Mid(Trim(mDate), 7, 4)) <= 99) Then
            mDate = Mid(mDate, 1, 6) + "19" + Format(Mid(mDate, 7, 2), "00")
        End If
        
        For i = 1 To Len(Trim(mDate))
            If (i <> 3 And i <> 6) And Mid(Trim(mDate), i, 1) = "." Then
                mError = True
            End If
        Next
        
        If Mid(Trim(mDate), 3, 1) <> "." And Mid(Trim(mDate), 3, 1) <> "-" And Mid(Trim(mDate), 3, 1) <> "/" And Mid(Trim(mDate), 3, 1) <> "\" Then
            mError = True
        End If
        
        If (Val(Mid(Trim(mDate), 1, 2)) = 0 Or Val(Mid(Trim(mDate), 1, 2)) > 32) Or (Val(Mid(Trim(mDate), 4, 2)) = 0 Or Val(Mid(Trim(mDate), 4, 2)) > 12) Then
            mError = True
        End If
        
        If Val(Mid(Trim(mDate), 7, 4)) < 1900 Then
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
            ValidDate = Format(Mid(Trim(mDate), 4, 2), "00") & "-" & Format(Mid(Trim(mDate), 1, 2), "00") & "-" & Format(Mid(Trim(mDate), 7, 4), "0000")

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
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("ACMAST", "AcCode", "ACNAME = '" & AcDesc & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetAcCode = Rs.Fields("AcCode")
    Else
        MsgBox "There is no Account" & AcDesc
        GetAcCode = ""
    End If
    Rs.Close
End Function

Function GetMemNobyAcCode(mAcCode As Integer) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("ACMAST", "MEMNO", "ACCODE = " & mAcCode)
    If Not (Rs.EOF Or Rs.BOF) Then
        GetMemNobyAcCode = Rs.Fields("MEMNO")
    Else
        MsgBox "There is no Account" & mAcCode
        GetMemNobyAcCode = ""
    End If
    Rs.Close
End Function

Function GetEmpIdbyAcCode(mAcCode As Integer) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("ACMAST", "EMPID", "ACCODE = " & mAcCode)
    If Not (Rs.EOF Or Rs.BOF) Then
        GetEmpIdbyAcCode = Rs.Fields("EMPID")
    Else
        MsgBox "There is no Account" & mAcCode
        GetEmpIdbyAcCode = ""
    End If
    Rs.Close
End Function

Function GetMemName(mMemNo As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("ACMAST", "ACNAME", "MEMNO = '" & mMemNo & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetMemName = Rs.Fields("ACName")
    Else
        MsgBox "There is no Account" & mMemNo
        GetMemName = ""
    End If
    Rs.Close
End Function

Function GetMemNobyEmpId(mEmpId As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("ACMAST", "MEMNO", "EMPID = '" & mEmpId & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetMemNobyEmpId = Rs.Fields("MEMNO")
    Else
        MsgBox "There is no Account" & mEmpId
        GetMemNobyEmpId = ""
    End If
    Rs.Close
End Function

Function GetAcName(mAcCode As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("ACMAST", "ACNAME", "AcCode = " & mAcCode)
    If Not (Rs.EOF Or Rs.BOF) Then
        GetAcName = Rs.Fields("ACName")
    Else
        MsgBox "There is no Account" & mAcCode
        GetAcName = ""
    End If
    Rs.Close
End Function

Function GetAcCodeByEmpid(mEmpId As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("ACMAST", "AcCode", "EMPID = '" & mEmpId & "'")
    
    If Not (Rs.EOF Or Rs.BOF) Then
        GetAcCodeByEmpid = Rs.Fields("AcCode")
    Else
        'MsgBox "There is no Account" & AcDesc
        GetAcCodeByEmpid = 0
    End If
    Rs.Close
End Function

Function GetDeptCode(DeptDesc As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("DeptMst", "DeptCode", "Dept = '" & DeptDesc & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetDeptCode = Rs.Fields("DeptCode")
    Else
        MsgBox "There is no Account" & DeptDesc
        GetDeptCode = ""
    End If
    Rs.Close
End Function

Function GetDeptName(mDeptCode As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("DEPTMST", "DEPT", "DEptCode = '" & mDeptCode & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetDeptName = Rs.Fields("Dept")
    Else
        MsgBox "There is no Account" & mDeptCode
        GetDeptName = ""
    End If
    Rs.Close
End Function

Function GetDedCode(Dname As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("DedMast", "DedCode", "Dname = '" & Dname & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetDedCode = Rs.Fields("DedCode")
    Else
        MsgBox "There is no Account" & Dname
        GetDedCode = ""
    End If
    Rs.Close
End Function

Function GetDedName(mDedCode As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("DEDMAST", "DNAME", "DEDCode = '" & mDedCode & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetDedName = Rs.Fields("DNAME")
    Else
        MsgBox "There is no Account" & mDedCode
        GetDedName = ""
    End If
    Rs.Close
End Function
'
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
    Dim Rs As Recordset
    Dim mstateid%
    mProfApOn = mProfApOn * 12
    If mCompId <> "04" Then
        mstateid = "0"
        'Set rs = oDb.RunQuery("PROFTAX", , mProfApOn & " >= STAMT And " & mProfApOn & " <= ENDAMT AND STATEID = '" & mstateid & "'")
        Set Rs = oDb.RunQuery("PROFTAX", , "STATEID = '" & mstateid & "' AND " & mProfApOn & " >= STAMT And " & mProfApOn & " <= ENDAMT")
    Else
        mstateid = "1"
        'Set rs = oDb.RunQuery("PROFTAX", , mProfApOn & " >= STAMT And " & mProfApOn & " <= ENDAMT AND STATEID = '" & mstateid & "'")
        Set Rs = oDb.RunQuery("PROFTAX", , "STATEID = '" & mstateid & "' AND " & mProfApOn & " >= STAMT And " & mProfApOn & " <= ENDAMT")
    End If
    If Not (Rs.EOF Or Rs.BOF) Then
        Rs.MoveFirst
        If mProfApOn >= Rs.Fields("STAMT") And mProfApOn <= Rs.Fields("ENDAMT") Then
            GetProfTax = Rs.Fields("PRFAMT")
        Else
            GetProfTax = 0
        End If
    Else
        GetProfTax = 0
    End If
    Rs.Close
End Function

Function GetDataBaseName(mCompNm As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("COMPMST", , "CNAME = '" & mCompNm & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetDataBaseName = Rs.Fields("DATABASENM")
    Else
        MsgBox "There is no Account" & mCompNm
        GetDataBaseName = ""
    End If
    Rs.Close
End Function

Function GetEmpName(mEmpId As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("EMPMAST", "EMPNAME", "EMPID = '" & mEmpId & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetEmpName = Rs.Fields("EMPNAME")
    Else
        MsgBox "There is no Employee " & mEmpId
        GetEmpName = ""
    End If
    Rs.Close
End Function

Function GetEmpId(mEmpName As String) As String
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("EMPMAST", "EMPID", "EMPNAME = '" & mEmpName & "'")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetEmpId = Rs.Fields("EMPID")
    Else
        MsgBox "There is no Company " & mEmpName
        GetEmpId = ""
    End If
    Rs.Close
End Function

Function GetOpLeaves(mEmpId As String, mOldDate As Date) As Double
    Dim Rs As Recordset
    Set Rs = oDb.RunQuery("SALTRAN", "BALLV", "EMPID = '" & mEmpId & "' AND  SALDT = #" & mOldDate & "#")
    If Not (Rs.EOF Or Rs.BOF) Then
        GetOpLeaves = Rs.Fields("BALLV")
    Else
        GetOpLeaves = 0
    End If
    Rs.Close
End Function

Function GetDays(mDate As Date) As Integer
    Select Case Mid(mDate, 4, 2)
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
    Dim mLen%, Lctr%, i%
    
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
                Lctr = 0
                mRs2.MoveFirst
                While Not mRs2.EOF = True
                    mdedname = Trim(mRs2.Fields("LEDNARR"))
                    mLen = Len(mdedname)
                    Lctr = Lctr + 1
                    .Write Format(Lctr, "##0") & Space(2) & mdedname & Space(70 - (mLen + 5))
                    .WriteLine Format(mRs2.Fields("AMOUNT"), "#####0.00")
                    mRs2.MoveNext
                Wend
                i = 1
                Lctr = Lctr + 13
                If Lctr < 20 Then
                    For i = Lctr To 20
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
    mPost = Val(Mid(Format(Val(mAmount), "#.00"), Len(Format(Val(mAmount), "#.00")) - 1, 2))
    mPre = Val(Mid(Format(Val(mAmount), "#.00"), 1, Len(Format(Val(mAmount), "#.00")) - 3))
    
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
    mPost = Val(Mid(Format(Val(mAmount), "#.00"), Len(Format(Val(mAmount), "#.00")) - 1, 2))
    mPre = Val(Mid(Format(Val(mAmount), "#.00"), 1, Len(Format(Val(mAmount), "#.00")) - 3))
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
      If (Mid$(vData, i, 1) = " ") Or (Mid$(vData, i, 1) = "-") And (i + 1 <= Len(vData)) Then
         Mid$(vData, i + 1, 1) = UCase$(Mid$(vData, i + 1, 1))
      End If
   Next i

   ProperCase = vData
End Function
'::--------------------------------------------------------------::

