Attribute VB_Name = "basFmtAndValidate"
Option Explicit

Const strModuleName As String * 30 = "basFmtAndValidate"

Public ctl As Control

'Types of input for function ValidateInput
Private Enum InputType
    Date_Slash_Input = 1
    Date_Dash_Input = 2
    Numeric_Input = 3
    Text_Input = 4
    Currency_Input = 5
End Enum

#Const VB6 = True 'if you got VB6 or higher, change this!

Const NONE = 0
Const STRINGTYPE = 1
Const INTEGERTYPE = 2
Const LONGTYPE = 3
Const FLOATTYPE = 4
Const CHARPERCENT = 5

Public m_szValue As String
Public m_szValueLen As Long

Public Function ConvertCase(TextBoxIn As TextBox, _
                            Optional ConvertTo As Integer = gstrCapWhat) As String
   If ConvertTo = 1 Then
      TextBoxIn.Text = StrConv$(TextBoxIn.Text, vbUpperCase)
   ElseIf converto = 2 Then
      TextBoxIn.Text = StrConv$(TextBoxIn.Text, vbLowerCase)
   ElseIf converto = 3 Then
      TextBoxIn.Text = StrConv$(TextBoxIn.Text, vbProperCase)
   End If
'  set select because if you cap something it can move the entry point to the right
   TextBoxIn.SelStart = Len(TextBoxIn.Text)

End Function

Public Function CompactSpaces(ByVal Text As String) As String

    CompactSpaces = Trim$(Text)
    While InStr(CompactSpaces, String(2, " ")) > 0
        CompactSpaces = vba.replace(CompactSpaces, String(2, " "), " ")
    Wend

End Function

Public Function SplitString(StringIn As String) As String

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "SplitString"
  
  Dim SplitArray() As String
  
  ' need to redim the array here
  
  SplitArray = Split(strAnimals, ",")
  ' Write #1, arrAnimals(0), arrAnimals(1), arrAnimals(2), arrAnimals(3), arrAnimals(4)

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function RemoveQuotes(pstr As String) As String
' Purpose:  Remove Single Quotes from a string

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "RemoveQuotes"
  
  Dim intPos As Integer
  intPos = InStr(pstr, "'")
  
  While intPos <> 0
    pstr = vba.Left$(pstr, intPos - 1) & vba.Right$(pstr, Len(pstr) - intPos)
    intPos = InStr(pstr, "'")
  Wend
  
  RemoveQuotes = pstr

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function PadLoR(StrIn As String, bytLen As Byte, strFillChar As String, Optional blnAlignRight As Boolean) As String
  ' Returns:Filled and aligned string.
  ' Does not work with fixed length strings.
  Dim i As Byte
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "PadLoR"

  If Len(StrIn) < bytLen Then
    For i = 1 To (bytLen - Len(StrIn))
      If blnAlignRight Then
        StrIn = strFillChar & StrIn
      Else
        StrIn = StrIn & strFillChar
      End If
    Next i
  End If
  
  PadLoR = vba.Left$(StrIn, bytLen)

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Sub RemoveApostrophes()
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "RemoveApostrophes"
  
  Dim iCount As Integer
    
  For iCount = 0 To 1
    If Text1(iCount).Text <> "" Then
      mstTextString = ReplaceSubstr(Text1(iCount).Text, "'", "")
      Text1(iCount).Text = UCase(mstTextString)
    End If
  Next

ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Function ReplaceSubstr(str As String, ByVal substr As String, ByVal newsubstr As String)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ReplaceSubstr"
  
  Dim Pos As Integer
  Dim startpos As Integer
  Dim new_str As String

  startpos = 1
  Pos = InStr(str, substr)
  
  Do While Pos > 0
    new_str = new_str & _
      vba.Mid$(str, startpos, Pos - startpos) & _
      newsubstr
    startpos = Pos + Len(substr)
    Pos = InStr(startpos, str, substr)
  Loop
  
  new_str = new_str & vba.Mid$(str, startpos)
  
  ReplaceSubstr = new_str
    
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function TrimNum(FieldIn As String)
  ' Removes extended characters from credit card numbers, dates, etc...
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "TrimNum"
  
  TrimNum = vba.replace(FieldIn, " ", "")
  TrimNum = vba.replace(TrimNum, ",", "")
  TrimNum = vba.replace(TrimNum, "-", "")
  TrimNum = vba.replace(TrimNum, "/", "")
  TrimNum = vba.replace(TrimNum, "(", "")
  TrimNum = vba.replace(TrimNum, ")", "")

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function TrimNull(Item As String)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "TrimNull"

  Dim Pos As Integer
  
  Pos = InStr(Item, Chr$(0))
  
  If Pos Then
    TrimNull = vba.Left$(Item, Pos - 1)
  Else
    TrimNull = Item
  End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Function SelectAll(ctl As Control)
   On Error GoTo ErrRtn
   gstrChkPt = "On Error": gstrProcName = "SelectAll"
   
   If ctl <> "" Then
      ctl.SelStart = 0
      ctl.SelLength = Len(ctl)
  End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Property Get Value() As String
    Value = m_szValue
End Property

Public Property Let Value(ByVal vData As String)
    m_szValue = vData
    UpdateLen
End Property

Public Sub UpdateLen()
    m_szValueLen = Len(m_szValue)
End Sub

Public Property Get Length() As Long
    Length = m_szValueLen
End Property

Public Function CountSubstring(ByVal strFind As String) As Long


    #If VB6 Then
        CountSubstring = (Len(m_szValue) - Len(vba.replace(m_szValue, strFind, ""))) / Len(strFind)
    #Else
        CountSubstring = (Len(m_szValue) - Len(Interfacevba.replace(m_szValue, strFind, ""))) / Len(strFind)
    #End If
End Function

Public Function Compare(ByVal szString As String) As Integer
    Compare = CompareStrings(m_szValue, szString)
End Function

Public Function CompareNoCase(ByVal szString As String) As Integer
    CompareNoCase = CompareStrings(UCase(m_szValue), UCase(szString))
End Function

Public Function Equals(ByVal szString As String) As Boolean
    Equals = (szString = m_szValue)
End Function

Public Function EqualsNoCase(ByVal szString As String) As Boolean
    EqualsNoCase = (UCase(szString) = UCase(m_szValue))
End Function

Public Function GetAt(ByVal nWhere As Long) As String
    GetAt = IIf(nWhere > m_szValueLen, "", vba.Mid$(m_szValue, nWhere, 1))
End Function

Public Sub SetAt(ByVal nWhere As Long, ByVal sChar As String)
    sChar = vba.Left$(sChar, 1)
    If nWhere < 1 Then nWhere = 1
    If nWhere > m_szValueLen Then nWhere = m_szValueLen
    m_szValue = vba.Left$(m_szValue, nWhere - 1) & sChar & vba.Right$(m_szValue, (m_szValueLen - nWhere))
End Sub

Public Function IsEmpty() As Boolean
    IsEmpty = (m_szValueLen = 0)
End Function

Public Sub MakeEmpty()
    m_szValue = ""
End Sub

Public Function vba.mid$(ByVal nFirst As Long, Optional ByVal nCount As Long) As String
    Mid = vba.Mid$(m_szValue, nFirst, nCount)
End Function

Public Function vba.left$(ByVal nCount As Long) As String
    Left = vba.Left$(m_szValue, nCount)
End Function

Public Function vba.right$(ByVal nCount As Long) As String
    Right = vba.Right$(m_szValue, nCount)
End Function

Public Function SpanIncluding(ByVal szChaRSet As String) As String
    Dim szRet As String
    Dim i As Long
    If m_szValueLen = 0 Or Len(szChaRSet) = 0 Then Exit Function


    For i = 1 To m_szValueLen


        If InStr(szChaRSet, vba.Mid$(m_szValue, i, 1)) <> 0 Then
            szRet = szRet & vba.Mid$(m_szValue, i, 1)
        End If
    Next
    SpanIncluding = szRet
End Function

Public Function SpanExcluding(ByVal szChaRSet As String) As String
    Dim szRet As String
    Dim i As Long
    If m_szValueLen = 0 Or Len(szChaRSet) = 0 Then Exit Function
    
    For i = 1 To m_szValueLen
        If InStr(szChaRSet, vba.Mid$(m_szValue, i, 1)) = 0 Then
            szRet = szRet & vba.Mid$(m_szValue, i, 1)
        End If
    Next
    SpanExcluding = szRet
End Function

Function UCaseMask(KeyAscii As Integer) As Integer
   'Use in Event KeyPress to set all characters in uppercase
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Function

Public Function LowerCase(KeyAscii As Integer)
   Dim char As String * 1
   
   char = Chr(KeyAscii)
   KeyAscii = Asc(LCase(char))
End Function

Public Function UpperCase(KeyAscii As Integer)
   Dim char As String * 1
   char = Chr(KeyAscii)
   KeyAscii = Asc(UCase(char))
End Function

Public Sub MakeUpper()
    m_szValue = UCase(m_szValue)
End Sub

Public Sub MakeLower()
    m_szValue = LCase(m_szValue)
End Sub

Public Sub MakeReverse()
    Dim szTemp As String
    Dim i As Long

    For i = m_szValueLen To 1 Step -1
        szTemp = szTemp & vba.Mid$(m_szValue, i, 1)
    Next
    m_szValue = szTemp
End Sub

Public Sub replace(strFind As String, strReplace As String)

    #If VB6 Then
        m_szValue = vba.replace(m_szValue, strFind, strReplace)
    #Else
        m_szValue = Interfacevba.replace(m_szValue, strFind, strReplace)
    #End If
    UpdateLen
End Sub

Public Sub Remove(ByVal szChar As String)
    If Len(szChar) > 1 Then szChar = vba.Left$(szChar, 1)
    replace szChar, ""
End Sub

Public Sub Insert(ByVal nIndex As Long, ByVal szStr As String)
    Dim szLeft As String
    Dim szRight As String
    If nIndex > m_szValueLen + 1 Then nIndex = m_szValueLen + 1
    szLeft = IIf(nIndex > 1 And m_szValueLen > 0, szLeft = vba.Left$(nIndex - 1), "")
    szRight = vba.Right$(m_szValueLen - nIndex + 1)
    m_szValue = szLeft & szStr & szRight
    UpdateLen
End Sub

Public Sub Delete(ByVal nIndex As Long, Optional nCount As Long = 1)
    Dim sLeft As String, sRight As String
    Dim nLen As Integer
    If nCount < 1 Then nCount = 1
    nLen = m_szValueLen

    If nIndex >= 0 And nIndex <= nLen Then
        If nIndex > 1 And nLen > 0 Then
            sLeft = vba.Left$(nIndex - 1)
        Else
            sLeft = ""
        End If


        If (nIndex + nCount) <= nLen Then
            sRight = vba.Mid$(m_szValue, nIndex + nCount)
        Else
            sRight = ""
        End If
        m_szValue = sLeft & sRight
    End If
    UpdateLen
End Sub

Public Sub Trimvba.left$()
    LTrim m_szValue
    UpdateLen
End Sub

Public Sub Trimvba.right$()
    RTrim m_szValue
    UpdateLen
End Sub

Public Function Find(ByVal szSubstr As String, Optional ByVal nStart As Long = 1) As Long
    Find = InStr(nStart, m_szValue, szSubstr)
End Function

Public Function ReverseFind(ByVal szSubstr As String) As Long
    #If VB6 Then
        ReverseFind = InStrRev(m_szValue, szSubstr)
    #Else
        ReverseFind = InterfaceInStrRev(m_szValue, szSubstr)
    #End If
End Function

Public Function FindOneOf(ByVal szChaRSet As String) As Long
    Dim i As Long

    If Not m_szValueLen > 0 Or Not Len(szChaRSet) > 0 Then
        Exit Function
    End If
    Dim iPos As Long

    For i = 1 To m_szValueLen
        iPos = InStr(szChaRSet, vba.Mid$(m_szValue, i, 1))

        If iPos <> 0 Then
            FindOneOf = iPos
            Exit Function
        End If
    Next
End Function

Public Sub AllocSysString(ByRef szString As String)
    szString = Space$(Length)
End Sub

Public Sub SetSysString(ByRef szString As String)
    szString = m_szValue
End Sub

Public Function Split(ByRef vBuf() As Variant, szDelim As String) As Long
    Split = InterfaceSplit(vBuf, m_szValue, szDelim)
End Function
'Public Function GetParameter(ByVal szFo
'     rmat As String, ByVal nRef As Integer) A
'     s String
' Dim szTemp As String
' Dim nPos As Integer
' Dim szBuf()
' nPos = 1
' szTemp = m_szValue
'
' If Not m_szValue Like szFormat Then Ex
'     it Function
'
' If VBA.vba.left$(szFormat, 1) = "*" Then sz
'     Format = " " & szFormat
    '
    ' InterfaceSplit szBuf, szFormat, "*"
    '
    ' For Each thing In szBuf
    ' szTemp = Interfacevba.replace(szTemp, (thi
    '     ng), Chr(255) & Chr(1))
    ' Next
    '
    ' InterfaceSplit szBuf, szTemp, Chr(255)
    '     & Chr(1)
    '
    ' If nRef - 1 < LBound(szBuf) Or nRef
    '     - 1 > UBound(szBuf) Then Exit Functio
    '     n
    '
    ' GetParameter = szBuf(nRef - 1)
    '
    '
    'End Function
    '

Public Sub Sprintf(DefString As String, ParamArray TheVals() As Variant)
    Dim DefLen As Integer, DefIdx As Integer
    Dim CurIdx As Integer, WorkString As String
    Dim CurVal As Integer, MaxVal As Integer
    Dim CurFormat As String, ValCount As Integer
    Dim xIndex As Integer, FoundV As Boolean, vType As Integer
    Dim CurParm As String
    DefLen = Len(DefString)
    DefIdx = 1
    CurVal = 0
    MaxVal = UBound(TheVals) + 1
    ValCount = 0
    ' Check for equal number of 'flags' as v
    '     alues, raise an error if inequal


    Do
        CurIdx = InStr(DefIdx, DefString, "%")


        If CurIdx > 0 Then


            If vba.Mid$(DefString, CurIdx + 1, 1) <> "%" Then ' don't count %%, will be converted To % later
                ValCount = ValCount + 1
                DefIdx = CurIdx + 1
            Else
                DefIdx = CurIdx + 2
            End If
        Else
            Exit Do
        End If
    Loop
    
    If ValCount <> MaxVal Then Err.Raise 450, , "Mismatch of parameteRS For String " & DefString & ". Expected " & ValCount & " but received " & MaxVal & "."
    
    DefIdx = 1
    CurVal = 0
    ValCount = 0
    
    WorkString = ""
    


    Do
        CurIdx = InStr(DefIdx, DefString, "%")


        If CurIdx <> 0 Then
            ' First, get the variable identifier. Sc
            '     an from Defidx (the %) to EOL looking fo
            '     r the
            ' first occurance of s, d, l, or f
            FoundV = False
            vType = NONE
            xIndex = CurIdx + 1


            Do While FoundV = False


                If Not FoundV Then
                    CurParm = vba.Mid$(DefString, xIndex, 1)


                    Select Case vba.Mid$(DefString, xIndex, 1)
                        Case "%"
                        vType = CHARPERCENT
                        FoundV = True
                        CurIdx = CurIdx + 1
                        CurVal = xIndex + 2
                        Case "s"
                        vType = STRINGTYPE
                        FoundV = True
                        CurVal = xIndex + 1
                        Case "d"
                        vType = INTEGERTYPE
                        FoundV = True
                        CurVal = xIndex + 1
                        Case "l"


                        If vba.Mid$(DefString, xIndex + 1, 1) = "d" Then
                            vType = LONGTYPE
                            FoundV = True
                            CurVal = xIndex + 2
                        End If
                        Case "f"
                        vType = FLOATTYPE
                        FoundV = True
                        CurVal = xIndex + 1
                    End Select
            End If
            If Not FoundV Then xIndex = xIndex + 1
        Loop
        If Not FoundV Then Err.Raise 93, , "Invalid % format in " & DefString
        CurParm = vba.Mid$(DefString, CurIdx, CurVal - CurIdx) ' For debugging purposes


        If vType = CHARPERCENT Then
            WorkString = WorkString & vba.Mid$(DefString, DefIdx, CurIdx - DefIdx)
            CurVal = CurVal - 1
        Else
            CurFormat = BuildFormat(CurParm, vType)
            WorkString = WorkString & vba.Mid$(DefString, DefIdx, CurIdx - DefIdx) & Format$(TheVals(ValCount), CurFormat)
            ValCount = ValCount + 1
        End If
        DefIdx = CurVal
    Else
        WorkString = WorkString & vba.Right$(DefString, Len(DefString) - DefIdx + 1)
        Exit Do
    End If
Loop
m_szValue = TreatBackSlash(WorkString)
End Sub

'***************************************
'Utility Functions
'***************************************

Public Function CompareStrings(ByVal szString1 As String, ByVal szString2 As String) As Integer
    Dim nValue1 As Long
    Dim nValue2 As Long
    Dim i As Long

    If szString1 <> "" And Len(szString1) > 0 Then

        For i = 1 To Len(szString1)
            nValue1 = nValue1 + CLng(Asc(vba.Mid$(szString1, i, 1)))
        Next

        For i = 1 To Len(szString2)
            nValue2 = nValue2 + CLng(Asc(vba.Mid$(szString2, i, 1)))
        Next
    End If
    
    Select Case nValue1 - nValue2
        Case 0: CompareStrings = 0
        Case Is > 0: CompareStrings = 1
        Case Is < 0: CompareStrings = -1
    End Select
End Function

Public Function BuildFormat(Parm As String, DataType As Integer) As String
    Dim Prefix As String, TmpFmt As String
    If DataType = LONGTYPE Then Prefix = vba.Mid$(Parm, 2, Len(Parm) - 3) Else Prefix = vba.Mid$(Parm, 2, Len(Parm) - 2)


    Select Case InStr(Prefix, ".")
        Case 0, Len(Prefix)
        If vba.Left$(Prefix, 1) = "0" Then TmpFmt = String(CInt(Prefix), "0") Else TmpFmt = "#"
        Case 1
        If vba.Mid$(Prefix, 2, 1) = "0" Then TmpFmt = "#." & String(CInt(vba.Right$(Prefix, 2)), "0") Else TmpFmt = "#.#"
        Case Else
        If vba.Left$(Prefix, 1) = "0" Then TmpFmt = String(CInt(vba.Left$(Prefix, InStr(Prefix, "."))), "0") & "." Else TmpFmt = "#."
        If vba.Mid$(Prefix, InStr(Prefix, ".") + 1, 1) = "0" Then TmpFmt = TmpFmt & String(CInt(vba.Right$(Prefix, InStr(Prefix, ".") - 1)), "0") Else TmpFmt = TmpFmt & "#"
    End Select
BuildFormat = TreatBackSlash(TmpFmt)
End Function

Public Function TreatBackSlash(sLine As String) As String
    #If VB6 Then
        sLine = vba.replace(sLine, "\n", vbCrLf)
        sLine = vba.replace(sLine, "\r", vbCr)
        sLine = vba.replace(sLine, "\t", vbTab)
        sLine = vba.replace(sLine, "\b", vbBack)
        sLine = vba.replace(sLine, "\0", vbNullString)
        sLine = vba.replace(sLine, "\\", "\")
    #Else
        sLine = Interfacereplace(sLine, "\n", vbCrLf)
        sLine = Interfacereplace(sLine, "\r", vbCr)
        sLine = Interfacereplace(sLine, "\t", vbTab)
        sLine = Interfacereplace(sLine, "\b", vbBack)
        sLine = Interfacereplace(sLine, "\0", vbNullString)
        sLine = Interfacereplace(sLine, "\\", "\")
    #End If
    TreatBackSlash = sLine
End Function

Public Function USPhoneNumber(CtlIn As TextBox) As Boolean

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "USPhoneNumber"

  'assume success
  USPhoneNumber = True
  Dim MyCursorPlace As Long
  Dim MyLen As Long
  Dim MyPlace As Long
  Dim MyBuffer As String
  Dim MyText As String
  Dim MyProfile As String
  Dim MyChar As String * 1
  Dim MyProfilePlace As Long
  'this is the format we are looking for
  MyProfile = "(###) ###-####"
  MyPlace = 1
  MyProfilePlace = 1
  'if there are more characters than allowed
  'then remove them
  If Len(CtlIn.Text) > Len(MyProfile) Then
     CtlIn.Text = vba.Left$(CtlIn.Text, Len(MyProfile))
     CtlIn.SelStart = Len(CtlIn.Text)
    Beep
  End If
 'here the statemachine parser begins
  MyText = CtlIn.Text
  MyLen = Len(MyText)
  'store the cursor position so we can restore it later
  MyCursorPlace = CtlIn.SelStart
  'the parser takes the pattern as the transition map.
  'starting at the beginning of the map, it compares the
  'current character with the state of the parser
  Do While MyPlace <= MyLen
    MyChar = vba.Mid$(MyText, MyPlace, 1)
      Select Case vba.Mid$(MyProfile, MyProfilePlace, 1)
        'if the current state calls for a numeric input
        'then check for a numeric value
        Case "#"
          If IsNumeric(MyChar) Then
              'add the character to the buffer
            MyBuffer = MyBuffer & MyChar
            'move to the next character
            MyPlace = MyPlace + 1
            'move to the next valid parser state
            MyProfilePlace = MyProfilePlace + 1
            'make sure we are indicating
            'a valid transition state
            CtlIn.ForeColor = MyDefaultColor
          Else
            'the character does not match the
            'parser's state so indicate an invalid
            'state and exit the parser
            CtlIn.ForeColor = vbRed
            GoTo ExitUSPhoneNumber
          End If
        'the parser state requires a "-"
        'in this character position
        Case "-"
          If MyChar = "-" Then
            'If it Is here Then add the
            'character to the buffer
            MyBuffer = MyBuffer & MyChar
           'move to next character position
            MyPlace = MyPlace + 1
            'move to next parser state
             MyProfilePlace = MyProfilePlace + 1
            'indicate a valid transition state
            'to the user
            CtlIn.ForeColor = MyDefaultColor
          Else
            'the required character is not present
            'and in this case we insert it meeting
            'the requirements of the parser state
            MyBuffer = MyBuffer & "-"
            'we shift the parser to the next state
            'but stay with the current character to
            'see if it matches the next state
            'transition
            MyProfilePlace = MyProfilePlace + 1
            'we also have to move the input cursor
            'over one to account for the added
            'character
            MyCursorPlace = MyCursorPlace + 1
          End If
        Case "(", ")", " "
          If MyChar = vba.Mid$(MyProfile, MyProfilePlace, _
            1) Then
            'If it Is here Then add the
            'character to the buffer
            MyBuffer = MyBuffer & MyChar
            'move to next character position
            MyPlace = MyPlace + 1
           'move to next parser state
            MyProfilePlace = MyProfilePlace + 1
            'indicate a valid transition state
            'to the user
            CtlIn.ForeColor = MyDefaultColor
          Else
            'the required character is not present
            'and in this case we insert it meeting
            'the requirements of the parser state
            MyBuffer = MyBuffer _
            & vba.Mid$(MyProfile, MyProfilePlace, 1)
            'we shift the parser to the next state
            'but stay with the current character
            'to see if it matches the next state
            'transition
             MyProfilePlace = MyProfilePlace + 1
            'we also have to move the input cursor
            'over one to account for the added
            'character
            MyCursorPlace = MyCursorPlace + 1
          End If
      End Select
  Loop
  If Len(MyBuffer) = Len(MyProfile) Then
    USPhoneNumber = True
  Else
    USPhoneNumber = False
  End If
  CtlIn.Text = MyBuffer
  CtlIn.SelStart = MyCursorPlace

ProcExit:
  Exit Function

ErrRtn:
  USPhoneNumber = False
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit
End Function

Public Function FormatPhone(PhoneNbrIn As String)
  ' A better way to do this would be to make it strip out any non-numeric value
  ' internatl #'s ?
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FormatPhone"
  
  TempPhone = replace$(TempPhone, "-", "")
  TempPhone = replace$(TempPhone, "(", "")
  TempPhone = replace$(TempPhone, ")", "")
  TempPhone = replace$(TempPhone, " ", "")
  TempPhone = replace$(TempPhone, ".", "")
  TempPhone = replace$(TempPhone, "x", "")
  
  If Len(TempPhone) = 7 Then
    FormatPhone = Format(TempPhone, "!&&&-&&&&")
  ElseIf Len(TempPhone) = 10 Then
    FormatPhone = Format(TempPhone, "!&&&-&&&-&&&&")
  ElseIf Len(TempPhone) > 10 Then
    FormatPhone = Format(TempPhone, "!&&&-&&&-&&&&x&&&&&&")
  Else
    FormatPhone = ""
  End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
  
End Function

Public Function SocSecNbrm(CtlIn As TextBox) As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "SocSecNbr"
  
  'assume success
  SocSecNum = True
  Dim MyCursorPlace As Long
  Dim MyLen As Long
  Dim MyPlace As Long
  Dim MyBuffer As String
  Dim MyText As String
  Dim MyProfile As String
  Dim MyChar As String * 1
  Dim MyProfilePlace As Long
  'this is the format we are looking for
  MyProfile = "###-##-####"
  MyPlace = 1
  MyProfilePlace = 1
  'if there are more characters than allowed
  'then remove them
  If Len(CtlIn.Text) > Len(MyProfile) Then
    CtlIn.Text = vba.Left$(CtlIn.Text, Len(MyProfile))
    CtlIn.SelStart = Len(CtlIn.Text)
    Beep
  End If
  'here the statemachine parser begins
  MyText = CtlIn.Text
  MyLen = Len(MyText)
  MyCursorPlace = CtlIn.SelStart
  'the parser takes the pattern
  'as the transition map. starting
  'at the beginning of the map,
  'it compares the current character
  'with the state of the parser
  Do While MyPlace <= MyLen
    MyChar = vba.Mid$(MyText, MyPlace, 1)
        Select Case vba.Mid$(MyProfile, MyProfilePlace, 1)
            'if the current state calls for a
            'numeric input then check for a
            'numeriv value
            Case "#"
                If IsNumeric(MyChar) Then
                    'add the character to the
                    'buffer
                    MyBuffer = MyBuffer & MyChar
                    'move to the next character
                    MyPlace = MyPlace + 1
                    'move to the next valid parser state
                    MyProfilePlace = MyProfilePlace + 1
                    'make sure we are indicating
                    'a valid transition state
                    CtlIn.ForeColor = MyDefaultColor
                Else
                    'the character does not match
                    'the parser's state so indicate
                    'an invalid state and exit
                    'the parser
                    CtlIn.ForeColor = vbRed
                    GoTo ExitSocSecNum
                End If
            'the parser state requires a "-"
            'in this character position
            Case "-"
                If MyChar = "-" Then
                    'If it Is here Then add the
                    'character to the buffer
                    MyBuffer = MyBuffer & MyChar
                    'move to next character position
                    MyPlace = MyPlace + 1
                    'move to next parser state
                    MyProfilePlace = MyProfilePlace + 1
                    'indicate a valid transition state
                    'to the user
                    CtlIn.ForeColor = MyDefaultColor
                Else
                    'the required character is not
                    'present and in this case we
                    'insert it meeting the requirements
                    'of the parser state
                    MyBuffer = MyBuffer & "-"
                    'we shift the parser to the next state
                    'but stay with the current character
                    'to see if it matches the next state
                    'transition
                    MyProfilePlace = MyProfilePlace + 1
                    'we also have to move the input cursor
                    'over one to account for the added
                    'character
                    MyCursorPlace = MyCursorPlace + 1
                End If
        End Select
  Loop
  If Len(MyBuffer) = Len(MyProfile) Then
    SocSecNum = True
  Else
    SocSecNum = False
  End If
  CtlIn.Text = MyBuffer
  CtlIn.SelStart = MyCursorPlace
End Function

Public Function RemoveMultSpaces(WordIn As String)
  ' Replaces sequential spaces with single spaces.
  ' Spaces on either end are removed.
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "RemoveMultSpaces"
  
  Dim i, WordLength, Character, LastCharacter, NewWord
  
  WordLength = Len(WordIn)
  For i = 1 To WordLength
    Character = vba.Mid$(WordIn, i, 1)
    If LastCharacter = " " And Character = " " Then
    Else
      NewWord = NewWordIn & Character
      LastCharacter = Character
    End If
  Next i

  RemoveMultSpaces = Trim(NewWord)

ProcExit:
  Exit Function

ErrRtn:
  SocSecNum = False
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit
End Function


Public Function MaxLength(ctl As Control, KeyAscii As Integer, MaxLen As Integer) As Integer
  'Use in Event KeyPress to limit number of characters in a textbox
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Max_Length"
  
  Dim strKey As String * 1
  Dim intLoc As Integer
  
  strKey = LCase(Chr(KeyAscii))
  
  ' check for all of the printable characters
  intLoc = InStr("abcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()_+-=[]{}\|':;", strKey)
  If intLoc > 0 Or strKey = Chr(34) Then    ' chr(34) = double quote
    'If the max length is not zero
    If MaxLength <> 0 Then
      'If the maximum length has been reached
       If KeyAscii <> vbKeyBack And (Len(ctl.Text) >= MaxLen) Then
        'If no text is selected
         If ctl.SelLength = 0 Then
          'Don't accept more characters
           MaxLength = 0
           Beep
         Else
           MaxLength = KeyAscii
         End If
      End If
    End If
  End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function CVNull(varIn As Variant, defType As VbVarType, defValue As Variant) As Variant
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "CVNull"
  
  If IsNull(varIn) Then
    Select Case defType
      Case vbInteger
        CVNull = CInt(defValue)
      Case vbLong
        CVNull = CLng(defValue)
      Case vbDate
        CVNull = CDate(defValue)
      Case vbDecimal
        CVNull = CDec(defValue)
      Case vbDouble
        CVNull = CDbl(defValue)
      Case vbSingle
        CVNull = CSng(defValue)
      Case vbBoolean
        CVNull = CBool(defValue)
      Case vbString
        CVNull = CStr(defValue)
      Case Else
        CVNull = varIn
    End Select
  Else
    If defType = vbInteger And CStr(varIn) = "" Then
      CVNull = CInt(0)
    ElseIf defType = vbLong And CStr(varIn) = "" Then
      CVNull = CLng(0)
    ElseIf defType = vbDouble And CStr(varIn) = "" Then
      CVNull = CDbl(0)
    ElseIf defType = vbSingle And CStr(varIn) = "" Then
      CVNull = CSng(0)
    ElseIf defType = vbString And CStr(varIn) = "" Then
      CVNull = CStr(defValue)
    Else
      CVNull = varIn
    End If
  End If
    
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function EmptyStr(str As Variant, Optional ReturnValueIfEmpty As Variant = "") As Variant
  'Use to return as value if a string is empty or null
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "EmptyStr"
  
  If str <> "" Then
    If IsNumeric(ReturnValueIfEmpty) Then
      If IsNumeric(str) Then
        EmptyStr = str
      Else
        EmptyStr = 0
      End If
    Else
      EmptyStr = str
    End If
  Else
    EmptyStr = ReturnValueIfEmpty
  End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function PadStr(StrIn As String, intFillLen As Integer, Optional strFillChar As String = " ") As String
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "PadStr"
  
  Dim i As Integer
  Dim strPrefix As String
  
  If Len(StrIn) > intFillLen Then
    PadStr = StrIn
  Else
    For i = 1 To intFillLen
      strPrefix = strPrefix + strFillChar
    Next i
      
    PadStr = strPrefix & StrIn
  End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Sub TxtBoxEdit(KeyAscii As Integer)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "TxtBoxEdit"
  
  If KeyAscii = 13 Then ' enter key
    KeyAscii = 0
    SendKeys "{Tab}"
  ' accept only numbers and back space keys
  ElseIf InStr(("1234567890" & vbBack & ""), Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
  Else
    ' chang the lower case alphabetics to upper case
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
  End If
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Function IsStringAlpha(s As String) As Long
  ' Returns 0 if the string is alpha.
  ' otherwise returns the position of the first character
   ' that failed the test.
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "IsStringAlpha"
  
   Dim i As Long
   
   For i = 1 To Len(s)
      If IsCharAlpha(Asc(vba.Mid$(s, i, 1))) = 0 Then
         IsStringAlpha = i
         Exit Function
      End If
   Next i
   
   IsStringAlpha = 0
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function IsStringAlphaNumeric(s As String) As Long
  ' Returns 0 if the string is alphaNumeric
  ' otherwise returns the position of the first character
  ' that failed the test.
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "IsStringAlphaNumeric"
  
   Dim i As Long
   
   For i = 1 To Len(s)
      If IsCharAlphaNumeric(Asc(vba.Mid$(s, i, 1))) = 0 Then
         IsStringAlphaNumeric = i
         Exit Function
      End If
   Next i
   
   IsStringAlphaNumeric = 0
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function IsStringNumeric(s As String) As Long
  ' Returns 0 if the string is Numeric
  ' otherwise returns the position of the first character
  ' that failed the test.
   Dim i As Long
   Dim j As Byte
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "IsStringNumeric"
   
   For i = 1 To Len(s)
      j = Asc(vba.Mid$(s, i, 1))
      If IsCharAlphaNumeric(j) = 1 Then
         If IsCharAlpha(j) = 1 Then
            IsStringNumeric = i
            Exit Function
         End If
      Else
         IsStringNumeric = i
         Exit Function
      End If
   Next i
   
   IsStringNumeric = 0
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function DisallowChar(StrIn As String) As String
  ' Name: Disallowing Certain Characters in a Textbox
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "DisallowChar"
  
  If InStr(1, txtInput.Text, "|") > 0 Then
     txtInput.Text = vba.replace(txtInput.Text, "|", "") ' simply replace any | With nothing
  End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function AutoComplete(TextBoxIn As TextBox, DBIn As Database, _
                             TableIn As String, FieldIn As String) As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "AutoComplete"
  
  'TextBoxIn is the textbox that will do the autocomplete thing
  'FieldIn is the field from the table that has the information that will fill the textbox
  'TableIn is the table where you will search for the information to fill the textbox
  'DBIn is the database with the TableIn

  Dim OldLen As Integer
  Dim dsTemp As Recordset

  AutoComplete = False
  If Not TextBoxIn.Text = "" And IsDelOrBack = False Then

    OldLen = Len(TextBoxIn.Text)
    Set dsTemp = DBIn.OpenRecordset("Select * from " & TableIn & " where " & FieldIn & " like '" & TextBoxIn.Text & "*'", dbOpenDynaset)
    If Err = 3075 Then
      'here we got a bug!!
    End If
         
    If Not dsTemp.RecordCount = 0 Then
      TextBoxIn.Text = dsTemp(FieldIn)
      If TextBoxIn.SelText = "" Then
        TextBoxIn.SelStart = OldLen
      Else
        TextBoxIn.SelStart = InStr(TextBoxIn.Text, TextBoxIn.SelText)
      End If
      TextBoxIn.SelLength = Len(TextBoxIn.Text)
      AutoComplete = True
    End If
  End If

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

 
  ' Public Function ValidateInput(KeyAscii As Integer, _
  '                              Format As InputType, _
  '                              Optional UpperCase As Boolean) As Integer
  ' Parameters:  - KeyAscii - The ANSI Keycode pressed
  '              - Uppercase - Display text in uppercase if true
  ' Returnvalue: - KeyAscii or 0 if not allowed
  '
  ' Description: - The ValidateInput function checks to see if the KeyCode is allowed
  ' examples of how to call this sub
  '
  'public Sub txtInput_KeyPress(KeyAscii As Integer)
  '
  '   If (optInput(1).Value = True) Then                       'Date   dd/mm/yy
  '      KeyAscii = ValidateInput(KeyAscii, Date_Slash_Input)
  '   ElseIf (optInput(2).Value = True) Then                   'Date  dd-mm-yy
  '      KeyAscii = ValidateInput(KeyAscii, Date_Dash_Input)
  '   ElseIf (optInput(3).Value = True) Then            '      'Numeric Only
  '      KeyAscii = ValidateInput(KeyAscii, Numeric_Input)
  '   ElseIf (optInput(4).Value = True) Then            '      'Text Only A-z
  '      KeyAscii = ValidateInput(KeyAscii, Text_Input, chkUppercase.Value)
  '   ElseIf (optInput(5).Value = True) Then            '      'Currency $0,000.00
  '      KeyAscii = ValidateInput(KeyAscii, Currency_Input)
  '   End If
  '
  ' End Sub
'  on error goto ErrRtn
'  gstrchkpt = "On Error": gstrProcName = "ValidateInput"
'
'  If (Format = Date_Slash_Input) Then
'    'dd/mm/yy Keycodes - 48 = 0, 57 = 9, 8 = BackSpace, 47 = /, 0 = Cancel user input
'
'    If (KeyAscii > 57 Or KeyAscii < 48) Then
'        If (KeyAscii <> 8) Then
'            If (KeyAscii <> 47) Then
'                KeyAscii = 0
'            End If
'        End If
'    End If
'  ElseIf (Format = Date_Dash_Input) Then
'    'dd-mm-yy Keycodes - 48 = 0, 57 = 9, 8 = BackSpace, 45 = -, 0 = Cancel user input
'
'    If (KeyAscii > 57 Or KeyAscii < 48) Then
'      If (KeyAscii <> 8) Then
'         If (KeyAscii <> 45) Then
'            KeyAscii = 0
'         End If
'      End If
'    End If
'  ElseIf (Format = Numeric_Input) Then
'    '0-9 Keycodes - 48 = 0, 57 = 9, 8 = BackSpace, 0 = Cancel user input
'    If (KeyAscii < 48 Or KeyAscii > 57) Then
'      If (KeyAscii <> 8) Then
'        KeyAscii = 0
'      End If
'    End If
'  ElseIf (Format = Text_Input) Then
'    'A-Z a-z 65 = A, 122 = Z, 32 = Space, 8 = BackSpace
'
'    If (KeyAscii >= 65 And KeyAscii <= 122 Or (KeyAscii = 32 Or KeyAscii = 8)) Then
'      'Change to uppercase
'      If (UpperCase = True) Then KeyAscii = Asc(UCase$(Chr(KeyAscii)))
'    Else
'      KeyAscii = 0
'    End If
'  ElseIf (Format = Currency_Input) Then
'    '$0,000.00 Keycodes - 48 = 0, 57 = 9, 8 = BackSpace, 36 = $, 44 = ',', 46 = ., '0 = Cancel
'
'    If (KeyAscii > 57 Or KeyAscii < 48) Then
'        If (KeyAscii <> 8 And KeyAscii <> 36 And KeyAscii <> 44 And KeyAscii <> 46) Then
'            KeyAscii = 0
'        End If
'    End If
'  End If
'
'  'Return KeyAscii or 0 if value is not allowed
'  ValidateInput = KeyAscii
'
'ProcExit:
'  Exit Function
'
'ErrRtn:
'   call errmsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
'   resume procexit:
'End Function

Public Function Validate(Text1 As TextBox) As Boolean
  '  example call:
  '    Dim xVal As Boolean
  '    xVal = Validate(Text1(Index))
  '    If Not xVal Then DisplayErrTbX Text1(Index)

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Validate"
  
  Dim xHour As String, xMinute As String
  Dim xDay As String, xMonth As String, xYear As String
  Dim CtrlVal As String
  Dim xInput As String
  Dim xRepeat As Single
  Dim Crka As String
  Dim xNo As Single
  Dim MaxNo As Single, MinNo As Single
  Dim displayMsg As String
    
  CtrlVal = UCase$(Text1.Tag)
    
  If InStr(CtrlVal, "NOTEMPTY;") Then
     xInput = Text1.Text
     If xInput <> Empty Then
        Validate = True
     Else
        Validate = False
        GoSub displayMsg
        Exit Function
     End If
  End If
    
  If InStr(CtrlVal, "UCASE;") Then
     xInput = Text1.Text
     Text1.Text = UCase$(xInput)
     Validate = True
    
  ElseIf InStr(CtrlVal, "LCASE;") Then
     xInput = Text1.Text
     Text1.Text = LCase$(xInput)
     Validate = True
  End If
    
  If InStr(CtrlVal, "TIME;") Then
     xInput = Text1.Text
     'preglej èe je v xInputu katerikoli drugi znak kot : in èe je ga zamenjaj z :
     For xRepeat = 1 To Len(xInput)
        Crka = vba.Mid$(xInput, xRepeat, 1)
        If Not IsNumeric(Crka) Then
            vba.Mid$(xInput, xRepeat) = ":"
        End If
     Next
     Text1 = xInput
     If Len(xInput) <= 2 Then
        Text1 = Text1 + ":0"
        xInput = Text1.Text
     ElseIf Len(xInput) = 3 Then
        xHour = vba.Mid$(xInput, 1, 1)
        xMinute = vba.Mid$(xInput, 2)
        Text1 = xHour + ":" + xMinute
        xInput = Text1.Text
     ElseIf Len(xInput) = 4 Then
        xHour = vba.Mid$(xInput, 1, 2)
        xMinute = vba.Mid$(xInput, 3)
        Text1 = xHour + ":" + xMinute
        xInput = Text1.Text
     End If
     xHour = Format$(xInput, "hh")
     xMinute = Format$(xInput, "nn")
     If Not IsNumeric(xHour) Then
        xHour = "HH"
     End If
     If Not IsNumeric(xMinute) Then
        xMinute = "MM"
     End If
     Text1 = xHour + ":" + xMinute
     If Len(Text1) > 5 Then
        Text1 = vba.Left$(Text1, 5)
     End If
     If xHour = "HH" And xMinute = "MM" Then
        Validate = False
        GoSub displayMsg
        Exit Function
     Else
        Validate = True
     End If
    
  ElseIf InStr(CtrlVal, "DATE;") Then
     xInput = Text1.Text
     If UCase$(xInput) = "N" Then
        Text1 = Format$(Now, "dd-mm-yyyy")
     End If
     If UCase$(xInput) = "T" Then
        Text1 = Format$(Now + 1, "dd-mm-yyyy")
     End If
     If UCase$(xInput) = "A" Then
        Text1 = Format$(Now + 2, "dd-mm-yyyy")
     End If
     If UCase$(xInput) = "Y" Then
        Text1 = Format$(Now - 1, "dd-mm-yyyy")
     End If
        
     If vba.Left$(xInput, 1) = "+" And Len(xInput) > 1 Then
        xNo = Val(vba.Mid$(xInput, 2))
        Text1 = Format$(Now + xNo, "dd-mm-yyyy")
     End If
        
     If vba.Left$(xInput, 1) = "-" And Len(xInput) > 1 Then
        xNo = Val(vba.Mid$(xInput, 2))
        Text1 = Format$(Now - xNo, "dd-mm-yyyy")
     End If
       
     xInput = Text1.Text
        
     If Len(xInput) <= 2 Then
        Text1 = Text1 + "-" + Format$(month(Now), "00")
        xInput = Text1.Text
     ElseIf Len(xInput) = 3 And InStr(xInput, "-") = 0 Then
        xDay = vba.Mid$(xInput, 1, 1)
        xMonth = vba.Mid$(xInput, 2)
        Text1 = xDay + "-" + xMonth
        xInput = Text1.Text
        'preveri èe je ta datum sploh možen!
        If Not IsDate(xInput) Or year(xInput) < 1990 Then
           xInput = xDay + xMonth
           xDay = vba.Mid$(xInput, 1, 2)
           xMonth = vba.Mid$(xInput, 3)
           Text1 = xDay + "-" + xMonth
           xInput = Text1.Text
        End If
     ElseIf Len(xInput) = 4 And InStr(xInput, "-") = 0 Then
        xDay = vba.Mid$(xInput, 1, 2)
        xMonth = vba.Mid$(xInput, 3)
        Text1 = xDay + "-" + xMonth
        xInput = Text1.Text
     End If
     If IsDate(xInput) Then
        If DateDiff("yyyy", "01-01-1990", xInput) < 0 Then
           Text1 = "dd-mm-yyyy"
           Validate = False
           GoSub displayMsg
           Exit Function
        Else
           Text1 = Format$(xInput, "dd-mm-yyyy")
           Validate = True
        End If
     Else
        Text1 = "dd-mm-yyyy"
        Validate = False
        GoSub displayMsg
        Exit Function
     End If
  End If
  If InStr(CtrlVal, "NUMERIC;") Then
     xInput = Text1.Text
     If IsNumeric(xInput) Then
        Validate = True
     Else
        Validate = False
        GoSub displayMsg
        Exit Function
     End If
  End If
  If InStr(CtrlVal, "MAX=") Then
     xInput = Text1.Text
     MaxNo = Val(Mid#(CtrlVal, InStr(CtrlVal, "MAX=") + 4))
     If Val(xInput) <= MaxNo Then
        Validate = True
     Else
        Validate = False
        GoSub displayMsg
        Exit Function
     End If
  End If
  If InStr(CtrlVal, "MIN=") Then
     xInput = Text1.Text
     MinNo = Val(Mid#(CtrlVal, InStr(CtrlVal, "MIN=") + 4))
     If Val(xInput) >= MinNo Then
        Validate = True
     Else
        Validate = False
        GoSub displayMsg
        Exit Function
     End If
  End If
  Validate = True

   ' displayMsg:
   ' If InStr(CtrlVal, "DISPLAY=") Then
   '    displayMsg = vba.mid$(CtrlVal, InStr(CtrlVal, "DISPLAY=") + 8)
   '    displayMsg = vba.left$(displayMsg, InStr(displayMsg, ";") - 1)
   '    MsgBox displayMsg, vbCritical
   ' End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Sub DisplayErrTbX(TbX As TextBox)
  ' see Validate above for usage
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "DaisplayErrTbX"
  
  Dim OldColor As Single
  Dim xRepeat As Single
  ' OldColor = TbX.BackColor
  For xRepeat = 1 To 1
    TbX.BackColor = vbRed
    TbX.Refresh
    Delay 5
  '    TbX.BackColor = OldColor
  '    TbX.Refresh
    Delay 5
  Next
  Beep
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Sub Delay(Sec)
  ' used by DisplayErrTbx
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Delay"
  
  Dim x As Single
  x = Timer
  Do
  Loop Until Abs(x - Timer) > Sec
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Sub TextBoxChange(TextBoxIn As TextBox)

  ' The StrConv line of code capitalizes the first letter of each word in a textbox. However, when
  ' you do this the text tends to move from right to left. In order to fix this problem I added the
  ' second line of code which tells the cursor to stay on the right side of the text for the entire
  ' length of the text.
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "TextBoxChange"
      
  TextBoxIn = StrConv(TextBoxIn, vbProperCase)
  TextBoxIn.SelStart = Len(TextBoxIn.Text)

ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Sub ChkAscii1(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Form_KeyDown"
  
  Dim FKey As Integer
  Dim shiftkey As Integer
  
  shiftkey = Shift And 7
  Select Case shiftkey
     Case 1 ' or vbShiftMask   SHIFT key
        
     Case 2 ' or vbCtrlMask    CTRL key
        
     Case 4 ' or vbAltMask     ALT Key
        
     Case 3 '                  SHIFT and CTRL Keys
       
     Case 5 '                  SHIFT and ALT
     
     Case 6 '                  CTRL and ALT
       
     Case 7 '                  SHIFT, CTRL and ALT Keys
       
  End Select

  If KeyCode = vbKeyHome Then
  If KeyCode = vbKeyLeft Then
  If KeyCode = vbKeyRight Then
  If KeyCode = vbKeyUp Then
  If KeyCode = vbKeyDown Then
  If KeyCode = vbKeyDelete Then
  If KeyCode = vbKeyEnd Then
  If KeyCode = vbKeyPageUp Then
  If KeyCode = vbKeyPageDown Then
  If KeyCode = vbKeyEscape Then
  If KeyCode = vbKeyPrint Then
  If KeyCode = vbKeyPause Then
  If KeyCode = vbKeyBack Then
  If KeyCode = vbKeyInsert Then
  If KeyCode = vbKeyReturn Then
  If KeyCode = vbKeyNumlock Then
  If KeyCode = vbKeyTab Then
  If KeyCode = vbKeyCapital Then
  If KeyCode = vbKeySpace Then
  If KeyCode = vbKeyScrollLock Then
  If KeyCode = vbkeyLineFeed Then
  If KeyCode = 92 Then           ' = "Win Start"
  If KeyCode = 93 Then           ' = "Menu"
  If KeyCode = 124 Then          ' = "|"
  
  For FKey = 112 To 123
    If KeyCode = FKey Then Label2.Caption = "F" & FKey - 111
  Next FKey
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Sub CheckForDup(Text As String)
  ' Prevent from adding duplicates to your listbox

  ' example call
  '    If lst1.ListCount = 0 Then
  '        lst1.AddItem txt1.text
  '    Else
  '        CheckForDup(txt1.text)
  '    End If
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "CheckForDup"
  
  For x = 0 To lst1.ListCount - 1
    lst1.ListIndex = x
    If Text = lst1.Text Then
       Exit Sub
    ElseIf x = lst1.ListCount - 1 Then
      lst1.AddItem Text
    End If
  Next x

ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Function ValidateControls(frm As Form) As Boolean
  ' usage:
  '    If ValidateControls(Me) = False Then Exit Sub
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ValidateControls"
  
    Dim DispMsg As String
    
    For Each ctl In frm.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Text = Empty Then
                DispMsg = "Enter data for " & ctl.Tag
                ctl.SetFocus
                Exit For
            End If
        ElseIf TypeOf ctl Is ComboBox Then
            If ctl.ListIndex < 0 Then
                DispMsg = "Select from the dropdown of " & ctl.Tag
                ctl.SetFocus
                Exit For
            End If
        End If
        
     Next ctl
        If Not DispMsg = Empty Then
            MsgBox DispMsg, vbInformation, "Error"
            Validate = False
        Else
            Validate = True
        End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Sub ClearControls(fm As Form)
  ' usage:
  '    Call ClearControls(Me)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ClearControls"
  

 For Each ctl In fm.Controls
    If TypeOf ctl Is TextBox Then
            ctl.Text = Empty
    ElseIf TypeOf ctl Is ComboBox Then
            ctl.ListIndex = -1
    End If
 Next ctl
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyPageUp
            KeyCode = 0
            If BtnPrevious.Enabled = True Then Call BtnPrevious_Click
        Case vbKeyPageDown
            KeyCode = 0
            If BtnNext.Enabled = True Then Call BtnNext_Click
        Case vbKeyF2
            KeyCode = 0
            Call BtnNew_Click
        Case vbKeyF5
            KeyCode = 0
            Call BtnDelete_Click
    End Select
End Sub

