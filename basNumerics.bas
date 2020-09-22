Attribute VB_Name = "basNumerics"
Option Explicit

Const strModuleName As String * 30 = "basNumerics"

Public Function IsEven(ByVal ValueIn As Long) As Boolean
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "IsEven"
  
  IsEven = Not -(ValueIn And 1)
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function IsOdd(ByVal ValueIn As Long) As Boolean
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "IsOdd"
   
  IsOdd = -(ValueIn And 1)
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function IsPrime(ByVal n As Long) As Boolean
    ' Returns true if the number is a prime number.
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "IsPrime"
    
    Dim i As Long

    IsPrime = False
    
    If n <> 2 And (n And 1) = 0 Then Exit Function 'test if div 2
    If n <> 3 And n Mod 3 = 0 Then Exit Function 'test if div 3
    For i = 6 To Sqr(n) Step 6
        If n Mod (i - 1) = 0 Then Exit Function
        If n Mod (i + 1) = 0 Then Exit Function
    Next
    
    IsPrime = True
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function ReturnPercent(NbrIn1 As Integer, NbrIn2 As Integer) As Integer
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ReturnPercent"
  
  return_percent = Int((NbrIn1 * 100) / NbrIn2)
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function NbrsOnly(TextBoxIn As TextBox, KeyAscii As Integer, _
                         Optional DecimalOK As Boolean = True)
  ' restricts the values being entered to a text box to Decimals, Negatives, And a period
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "NbrsOnly"
  
  Dim key As String
  
  key = Chr$(KeyAscii) 'Convert To String

  Select Case key
    Case "0" To "9", "-", "."
      If (key = "-" And (InStr(TextBoxIn.Text, "-") Or TextBoxIn.SelStart <> 0)) Or _
         (key = "." And Not DecimalOK) Then
          KeyAscii = 0
      Else
         TextBoxIn.Text = TextBoxIn & key
'        TextBoxIn.Text = val(vba.left$(TextBoxIn.Text, TextBoxIn.SelStart) + key + _
'          vba.mid$(TextBoxIn.Text, TextBoxIn.SelStart + TextBoxIn.SelLength + 1))
      End If
    Case Chr$(8) 'Backspace
    Case Else
      KeyAscii = 0
  End Select
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Static Function RegionDecimalPoint() As String
  ' used in NumericEntry
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "RegionDecimal"
  
  If vba.Mid$(Format(1, "#.0000"), 2, 1) = "," Then
    RegionDecimalPoint = ","
  Else
    RegionDecimalPoint = "."
  End If
         
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function FormatNbr(NbrIn As String, _
                          Optional DecPlaces As Integer = 2, _
                          Optional DigitPlaces As Integer = 15, _
                          Optional ShowLeadingZero As Boolean = False, _
                          Optional ShowTrailingZero As Boolean = True, _
                          Optional NegSign As String = "L", _
                          Optional CurrencySign As String = "$")
  ' NegSign - "X" = negative not allowed, "L" = leading, "T" = trailing

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FormatNbr"
  
  Dim i As Long, j As Long, tString As String
  Dim LeftString As String, RightString As String
  Dim TextString As String
  Dim blnAddNegSign As Boolean
    
  If Len(NbrIn) > 0 Then
     If InStr(NbrIn, "-") And NegSign = "X" Then
       MsgBox "Negative Numbers Not Allowed In This Field.", vbOKOnly
       NbrIn = " "
       Exit Function
     End If
     
     TextString = NbrIn
        
     '/* Remove extra leading zeros
     If ShowLeadingZero Then
        If vba.Left$(TextString, 2) = "00" Then
           TextString = "0"
        End If
     ElseIf vba.Left$(TextString, 1) <> "." And Val(TextString) = 0 Then
        TextString = "0"
     End If
     
     If Val(TextString) > 0 And vba.Left$(TextString, 1) = "0" Then TextString = vba.Mid$(TextString, 2)
        
     '/* Make sure it is a number
     tString = vbNullString
     blnAddNegSign = False
     For i = 1 To Len(TextString)
       Select Case vba.Mid$(TextString, i, 1)
         Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            tString = tString & vba.Mid$(TextString, i, 1)
         Case "."
            If DecPlaces > 0 Then
                tString = tString & vba.Mid$(TextString, i, 1)
            End If
         Case "-"
            If NegSign <> "X" Then
              ' ' we'll add it back after formatting the rest of the number
              blnAddNegSign = True
            End If
         Case ","
           ' do nothing - we'll remove them later
         Case CurrencySign
           ' do nothing - add back later if AddDollarSign = true
         Case Else
           MsgBox "Only Numbers, ',', '-', '.', and Currency Sign Allowed In This Field", vbOKOnly
           NbrIn = " "
           Exit Function
       End Select
     Next i
       
     '/* Remove double .. */
     i = InStr(tString, ".")
     j = InStrRev(tString, ".")
     If j <> i Then
        tString = vba.Mid$(tString, 1, j - 1) & vba.Mid$(tString, j + 1)
     End If
      
     '/* Left side of decimal place */
     If i > 0 Then
        LeftString = vba.Left$(tString, i - 1)
     Else
        LeftString = tString
     End If
       
     LeftString = vba.Left$(LeftString, DigitPlaces)
        
     If ShowLeadingZero Then
        If LeftString = vbNullString Then
           LeftString = "0"
        ElseIf TextString = "-." Then
           LeftString = "-0"
        End If
     Else
        If LeftString = "0" Then LeftString = vbNullString
     End If
     
     If Len(LeftString) > 3 Then
        LeftString = vba.Mid$(LeftString, 1, Len(LeftString) - 3) + gstrCommaChar + _
                     vba.Mid$(LeftString, Len(LeftString) - 2)
     End If
        
     If Len(LeftString) > 7 Then
        LeftString = vba.Mid$(LeftString, 1, Len(LeftString) - 7) + gstrCommaChar + _
                     vba.Mid$(LeftString, Len(LeftString) - 6)
     End If
        
     If Len(LeftString) > 11 Then
        LeftString = vba.Mid$(LeftString, 1, Len(LeftString) - 11) + gstrCommaChar + _
                     vba.Mid$(LeftString, Len(LeftString) - 10)
     End If
        
     '/* Right side of decimal place */
     If i > 0 Then
        If DecPlaces > 0 Then
           RightString = vba.Mid$(tString, i, DecPlaces + 1)
           If ShowTrailingZero Then RightString = vba.Left$(RightString & String$(DecPlaces, "0"), DecPlaces + 1)
        Else
           RightString = vbNullString
        End If
     Else
        If vba.Right$(tString, 1) = gstrDecPtChar Then
           RightString = gstrDecPtChar
        Else
           RightString = vbNullString
        End If
     End If
        
     '/* Combine left and right */
     FormatNbr = LeftString & RightString
     
     If blnAddNegSign = True Then
       If NegSign = "L" Then
         FormatNbr = "-" + FormatNbr
       ElseIf NegSign = "T" Then
         FormatNbr = FormatNbr + "-"
       End If
     End If
     
     If CurrencySign <> "" Then FormatNbr = CurrencySign + FormatNbr
    
  End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function RoundNbr(strNbrIn As String, strNbrPlaces As String)

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "RoundNbr"
  
  RoundNbr = FormatNbr(strNbrIn, strNbrPlaces)
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function
End Function

Function RoundOff(NbrIn As Long, DecPlacesIn As Integer) As Long

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "RoundOff"
     
  RoundOff = (NbrIn * (10 ^ DecPlacesIn) + 0.5) / (10 ^ DecPlacesIn)
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Function IsInteger(lngNbrIn As Long) As Boolean
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "IsInteger"
  
  If lngNbrIn Mod 1 = 0 Then
    IsInteger = True
  Else
    IsInteger = False
  End If

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function Ordinal(Number As Integer) As String

  ' Accepts an integer, returns the ordinal number
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Ordinal"

  Dim strSuffix As String

  If Number > 21 Then
    Select Case vba.Right$(Trim(str(Number)), 1)
      Case 1
    strSuffix = "st"
      Case 2
    strSuffix = "nd"
      Case 3
    strSuffix = "rd"
      Case 0, 4 To 9
    strSuffix = "th"
    End Select

  Else
    Select Case Number
      Case 1
    strSuffix = "st"
      Case 2
    strSuffix = "nd"
      Case 3
    strSuffix = "rd"
      Case 4 To 20
    strSuffix = "th"
    End Select

  End If

  Ordinal = Trim(str(Number)) & strSuffix

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
        
End Function

Public Function NbrToText(NbrIn As String, _
                          Optional NbrDecIn As Integer = 2, _
                          Optional IsCurrencyIn As Boolean = True, _
                          Optional ShowZeroDolIn As Boolean = False, _
                          Optional NegativeSymbolIn As String = "Negative", _
                          Optional NegLeadTrailIn As String = "L", _
                          Optional SpellOutDecIn As Boolean = False) As String
  ' usage:
  ' note: the max nbr of dec places in this version is 6
  '   Text1.Text = basNumerics.NbrToText(Text1.Text, 2, True, , "Minus", "T")
  '   Text1.Text = basNumerics.NbrToText(Text1.Text, 0, False, , "-", "L")
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "NbrToText"
        
  Dim i As Long
  Dim lngDecPtLoc As Long
  Dim strDigitPart As String
  Dim strDecimalPart As String
  Dim strDecimalLit As String
  Dim strCharToProcess As String
    
  Dim blnHasDecPt As Boolean
  Dim blnNegative As Boolean
  Dim blnPercent As Boolean
  Dim blnDollars As Boolean
  
  Dim strDecLit As String
  Dim strMillionths As String
  Dim strThousandths As String
  Dim strMillionthsLit As String
  Dim strThousandthslit As String
  
  Dim strHundreds As String
  Dim strThousands As String
  Dim strMillions As String
  Dim strBillions As String
  Dim strTrillions As String
  Dim strHundredsLit As String
  Dim strThousandslit As String
  Dim strMillionsLit As String
  Dim strBillionslit As String
  Dim strTrillionsLit As String
  
  NbrToText = NbrIn
  NbrToText = vba.replace(NbrToText, ",", "")
  
  If InStr(1, NbrToText, "$") Then
     blnDollars = True
     NbrToText = vba.replace(NbrToText, "$", "")
  End If
    
  If InStr(1, NbrToText, "-") Then
     blnNegative = True
     NbrToText = vba.replace(NbrToText, "-", "")
  End If
    
  If InStr(1, NbrToText, "%") Then
     blnPercent = True
     NbrToText = vba.replace(NbrToText, "%", "")
  End If
    
  If InStr(1, NbrToText, gstrDecPtChar) Then
     lngDecPtLoc = InStr(1, NbrToText, gstrDecPtChar)
     If InStr(lngDecPtLoc + 1, NbrToText, gstrDecPtChar) Then
        NbrToText = "Multi Decimaled"
        Exit Function
     Else
        blnHasDecPt = True
     End If
  End If

  If Not (IsNumeric(NbrToText)) Then
     NbrToText = "Not Numeric"
     Exit Function
  End If
      
  If blnHasDecPt Then
    strDecimalPart = vba.Mid$(NbrToText, lngDecPtLoc + 1)
    If NbrDecIn > 0 Then
      strDecimalPart = vba.Left$(strDecimalPart & "000000000000000", NbrDecIn)
    Else
      strDecimalPart = ""
    End If
    strDigitPart = vba.Left$(NbrToText, lngDecPtLoc - 1)
  Else
    strDecimalPart = ""
    strDigitPart = NbrToText
  End If
  
  ' Create a fixed length string to work on
  strDigitPart = vba.Right$("000000000000000" & strDigitPart, 15)
  
  ' build 3-digit fields to send to the GetNbrLit function
  For i = Len(strDigitPart) To 1 Step -1
     strCharToProcess = vba.Mid$(strDigitPart, i, 1)
     If i > 12 Then
       strHundreds = strCharToProcess & strHundreds
     ElseIf i > 9 Then
       strThousands = strCharToProcess & strThousands
     ElseIf i > 6 Then
       strMillions = strCharToProcess & strMillions
     ElseIf i > 3 Then
       strBillions = strCharToProcess & strBillions
     Else
       strTrillions = strCharToProcess & strTrillions
     End If
  Next i
  
  Do While vba.Left$(strTrillions, 1) = "0"
     If Len(strTrillions) = 1 Then
        strTrillions = ""
     Else
        strTrillions = vba.Mid$(strTrillions, 2)
     End If
  Loop
  
  Do While vba.Left$(strBillions, 1) = "0"
     If Len(strBillions) = 1 Then
        strBillions = ""
     Else
        strBillions = vba.Mid$(strBillions, 2)
     End If
  Loop
  
  Do While vba.Left$(strMillions, 1) = "0"
     If Len(strMillions) = 1 Then
        strMillions = ""
     Else
        strMillions = vba.Mid$(strMillions, 2)
     End If
  Loop
  
  Do While vba.Left$(strThousands, 1) = "0"
     If Len(strThousands) = 1 Then
        strThousands = ""
     Else
        strThousands = vba.Mid$(strThousands, 2)
     End If
  Loop
  
  Do While vba.Left$(strHundreds, 1) = "0"
     strHundreds = vba.Mid$(strHundreds, 2)
     If Len(strHundreds) = 1 Then
        strHundreds = ""
     Else
        strHundreds = vba.Mid$(strHundreds, 2)
     End If
  Loop
  
  If Len(strTrillions) <> 0 Then strTrillionsLit = GetNbrLit(strTrillions) & " Trillion, "
  If Len(strBillions) <> 0 Then strBillionslit = GetNbrLit(strBillions) & " Billion, "
  If Len(strMillions) <> 0 Then strMillionsLit = GetNbrLit(strMillions) & " Million, "
  If Len(strThousands) <> 0 Then strThousandslit = GetNbrLit(strThousands) & " Thousand, "
  If Len(strHundreds) <> 0 Then strHundredsLit = GetNbrLit(strHundreds)
  
  If Len(strTrillions) + Len(strBillions) + Len(strMillions) + _
     Len(strThousands) + Len(strHundreds) > 0 Then
     NbrToText = Trim(Trim(strTrillionsLit) & " " & Trim(strBillionslit) & _
                 Trim(strMillionsLit) & " " & Trim(strThousandslit) & " " & Trim(strHundredsLit))
  Else
     NbrToText = ""
  End If
  
  If IsCurrencyIn = True Then
     If NbrToText = "" Then
        If ShowZeroDolIn = True Then
           NbrToText = "Zero Dollars"
        End If
     Else
        NbrToText = NbrToText & " Dollar" & IIf(NbrToText <> "One", "s", "")
     End If
  End If
  
  
  If NbrDecIn > 0 Then
     strDecimalPart = vba.Left$(strDecimalPart, NbrDecIn)
  
     If SpellOutDecIn = True Then
  
        ' build 3-digit fields to send to the GetNbrLit function
        For i = Len(strDigitPart) To 1 Step -1
           strCharToProcess = vba.Mid$(strDigitPart, i, 1)
           If i > 12 Then
             strHundreds = strCharToProcess & strHundreds
           ElseIf i > 9 Then
             strThousands = strCharToProcess & strThousands
           ElseIf i > 6 Then
             strMillions = strCharToProcess & strMillions
           ElseIf i > 3 Then
             strBillions = strCharToProcess & strBillions
           Else
             strTrillions = strCharToProcess & strTrillions
           End If
        Next i
  
        Do While vba.Left$(strTrillions, 1) = "0"
           If Len(strTrillions) = 1 Then
              strTrillions = ""
           Else
              strTrillions = vba.Mid$(strTrillions, 2)
           End If
        Loop
  
        Do While vba.Left$(strBillions, 1) = "0"
           If Len(strBillions) = 1 Then
              strBillions = ""
           Else
              strBillions = vba.Mid$(strBillions, 2)
           End If
        Loop
  
        Do While vba.Left$(strMillions, 1) = "0"
           If Len(strMillions) = 1 Then
              strMillions = ""
           Else
              strMillions = vba.Mid$(strMillions, 2)
           End If
        Loop
  
        Do While vba.Left$(strThousands, 1) = "0"
           If Len(strThousands) = 1 Then
              strThousands = ""
           Else
              strThousands = vba.Mid$(strThousands, 2)
           End If
        Loop
  
        Do While vba.Left$(strHundreds, 1) = "0"
           strHundreds = vba.Mid$(strHundreds, 2)
           If Len(strHundreds) = 1 Then
              strHundreds = ""
           Else
              strHundreds = vba.Mid$(strHundreds, 2)
           End If
        Loop
  
        If Len(strTrillions) <> 0 Then strTrillionsLit = GetNbrLit(strTrillions) & " Trillion, "
        If Len(strBillions) <> 0 Then strBillionslit = GetNbrLit(strBillions) & " Billion, "
        If Len(strMillions) <> 0 Then strMillionsLit = GetNbrLit(strMillions) & " Million, "
        If Len(strThousands) <> 0 Then strThousandslit = GetNbrLit(strThousands) & " Thousand, "
        If Len(strHundreds) <> 0 Then strHundredsLit = GetNbrLit(strHundreds)
  
        If Len(strTrillions) + Len(strBillions) + Len(strMillions) + _
           Len(strThousands) + Len(strHundreds) > 0 Then
           strDecimalPart = Trim(Trim(strTrillionsLit) & " " & Trim(strBillionslit) & _
               Trim(strMillionsLit) & " " & Trim(strThousandslit) & " " & Trim(strHundredsLit))
        Else
           strDecimalPart = ""
        End If
        
        Select Case Len(strDecimalPart)
           Case 1: strDecLit = "Tenths"
           Case 2: strDecLit = "Hundreths"
           Case 3: strDecLit = "Thousandths"
           Case 4: strDecLit = "Ten Thousandths"
           Case 5: strDecLit = "Hundred Thousandths"
           Case 6: strDecLit = "Millionths"
        End Select
        strDecimalPart = strDecimalPart & strDecLit
     End If ' if spelloutdecin
   
     If IsCurrencyIn = True Then
        If NbrToText <> "" Then
           strDecimalLit = " and "
        End If
        strDecimalLit = strDecimalLit + strDecimalPart + " Cent" + _
              IIf((strDecimalPart <> "") And (strDecimalPart <> "One"), "s", "")
     Else
        strDecimalLit = " Point " + strDecimalPart
     End If
     NbrToText = NbrToText & strDecimalLit
  End If
  
  If blnNegative = True Then
     If NegLeadTrailIn = "L" Then
        NbrToText = NegativeSymbolIn & " " & Trim(NbrToText)
     Else
        NbrToText = Trim(NbrToText) & " " & NegativeSymbolIn
     End If
  End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Private Function GetNbrLit(NbrIn As String) As String

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "GetNbrLit"
  
  Dim strHundredsChar As String * 1
  Dim strTensChar As String * 1
  Dim strOnesChar As String * 1
  Dim strHundreds As String
  Dim strTens As String
  Dim strNbr As String
  
  If Len(NbrIn) = 0 Then
    Exit Function
  End If
  
  GetNbrLit = vba.Right$("000" & NbrIn, 3)
    
  strHundredsChar = vba.Left$(GetNbrLit, 1)
  strTensChar = vba.Mid$(GetNbrLit, 2, 1)
  strOnesChar = vba.Right$(GetNbrLit, 1)
  
  Select Case strHundredsChar
    Case 1: strHundreds = "One"
    Case 2: strHundreds = "Two"
    Case 3: strHundreds = "Three"
    Case 4: strHundreds = "Four"
    Case 5: strHundreds = "Five"
    Case 6: strHundreds = "Six"
    Case 7: strHundreds = "Seven"
    Case 8: strHundreds = "Eight"
    Case 9: strHundreds = "Nine"
    Case Else: strHundreds = ""
  End Select
  
  If Len(strHundreds) <> 0 Then
    strHundreds = strHundreds + " Hundred "
  End If
  
  If strTensChar = "0" Then ' The value is 0-9
    Select Case strOnesChar
      Case 1: strTens = "One"
      Case 2: strTens = "Two"
      Case 3: strTens = "Three"
      Case 4: strTens = "Four"
      Case 5: strTens = "Five"
      Case 6: strTens = "Six"
      Case 7: strTens = "Seven"
      Case 8: strTens = "Eight"
      Case 9: strTens = "Nine"
      Case Else
    End Select
  ElseIf strTensChar = "1" Then ' The value is 10-19
    Select Case strOnesChar
      Case 0: strTens = "Ten"
      Case 1: strTens = "Eleven"
      Case 2: strTens = "Twelve"
      Case 3: strTens = "Thirteen"
      Case 4: strTens = "Fourteen"
      Case 5: strTens = "Fifteen"
      Case 6: strTens = "Sixteen"
      Case 7: strTens = "Seventeen"
      Case 8: strTens = "Eighteen"
      Case 9: strTens = "Nineteen"
      Case Else
    End Select
  Else ' The value is 20-99
    Select Case strTensChar
      Case 2: strTens = "Twenty"
      Case 3: strTens = "Thirty"
      Case 4: strTens = "Forty"
      Case 5: strTens = "Fifty"
      Case 6: strTens = "Sixty"
      Case 7: strTens = "Seventy"
      Case 8: strTens = "Eighty"
      Case 9: strTens = "Ninety"
      Case Else
    End Select
  
    Select Case strOnesChar
      Case 1: strTens = strTens & "-One"
      Case 2: strTens = strTens & "-Two"
      Case 3: strTens = strTens & "-Three"
      Case 4: strTens = strTens & "-Four"
      Case 5: strTens = strTens & "-Five"
      Case 6: strTens = strTens & "-Six"
      Case 7: strTens = strTens & "-Seven"
      Case 8: strTens = strTens & "-Eight"
      Case 9: strTens = strTens & "-Nine"
      Case Else
    End Select
  End If
  
  GetNbrLit = strHundreds & strTens
  
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function


