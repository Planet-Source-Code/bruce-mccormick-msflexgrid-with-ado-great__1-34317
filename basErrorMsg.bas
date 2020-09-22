Attribute VB_Name = "basErrorMsg"
Option Explicit

Private Const strModuleName As String * 30 = "basErrorMsg"
Private strCompName As String * 15
Private strTime As String * 6
Private MSErrorMsg As String
Private strLogFileName As String

' following used to obtain MS error msg
Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
  (ByVal lngFlags As Long, lpSource As Any, ByVal lngMessageId As Long, ByVal lngLanguageId As Long, _
   ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
  
Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Const FORMAT_MESSAGE_FROM_STRING = &H400
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Const LANG_USER_DEFAULT = &H400&

Public Function ErrMsg(Optional strModuleNameIn As String = "Unknown", _
                       Optional strProcNameIn As String = "Unknown", _
                       Optional strChkPtIn As String = "Unknown", _
                       Optional lngErrNbrIn As Long = 0, _
                       Optional strErrDescIn As String = "Unknown", _
                       Optional strErrSourceIn As String = "Unkown", _
                       Optional strTableIn As String = "", _
                       Optional strKeyIn As String = "", _
                       Optional strMiscErrorInfo1 As String = "", _
                       Optional strHelpFile As String = "", _
                       Optional strHelpContext As String = "")
  ' The strChkPtIn  is just a string variable is use to pin an error down within a proc.
  ' If its a big proc, I'll put "gstrChkpt = <a literal of the next command>"
  ' so I could limit the possibile area to 1/2 or 1/3 or 1/4 of a proc
                       
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ErrMsg"
  
  Dim strMsg As String
  Dim intMsgType As Integer
  Dim intResponse As Integer
  
  ' log error
  
  strCompName = libUtilities.ComputerName
  strTime = Format(time, "HHMMSS")
  
  strLogFileName = App.Path & "/ErrorLogs"
  Call basFiles.CreateFolder(strLogFileName)
     
  strLogFileName = strLogFileName & "/" & Format(Now, "yyyymmdd") & ".txt"
  Open strLogFileName For Append As #1
  
  Write #1, " ************************************************* " & _
       vbCrLf & " Date/Time:    " & str(Date) & " " & strTime & _
       vbCrLf & " Module:       " & strModuleNameIn & _
       vbCrLf & " Procedure:    " & strProcNameIn & _
       vbCrLf & " Check Point:  " & strChkPtIn & _
       vbCrLf & " Error Source: " & strErrSourceIn & _
       vbCrLf & " Error Number: " & str(lngErrNbrIn) & _
       vbCrLf & " Description:  " & strErrDescIn & _
       vbCrLf & " UserID:       " & gstrUserID & _
       vbCrLf & " Machine ID:   " & strCompName

  Close #1

  strMsg = "Please Write Down Or Do A Screen Print Of This Screen And " & _
           "Report This Error To The Person Responsible For Maintaining " & _
           "This System " & vbCrLf & vbCrLf

  If lngErrNbrIn > vbObjectError Then
    strMsg = strMsg & "A Program "
  Else
    strMsg = strMsg & "A Visual Basic"
  End If
  
  strMsg = strMsg & " Error Has Occured: " & _
           vbCrLf & " In Module:    " & strModuleNameIn & _
           vbCrLf & " Procedure:    " & strProcNameIn & _
           vbCrLf & " Check Point:  " & strChkPtIn & _
           vbCrLf & " Error Source: " & strErrSourceIn & _
           vbCrLf & " Error Number: " & str(lngErrNbrIn) & _
           vbCrLf & " Description:  " & strErrDescIn
           
  If strTableIn <> "" Then
     strMsg = strMsg & vbCrLf & " Table:        " & strTableIn
  End If
           
  If strKeyIn <> "" Then
     strMsg = strMsg & vbCrLf & " Key:          " & strKeyIn
  End If
           
  If strMiscErrorInfo1 <> "" Then
     strMsg = strMsg & vbCrLf & " Other Info:   " & strMiscErrorInfo1
  End If
           
  If strHelpFile <> "" Then
     strMsg = strMsg & vbCrLf & " Help File:    " & strHelpFile
  End If
           
  If strHelpContext <> "" Then
     strMsg = strMsg & vbCrLf & " Help Context: " & strHelpContext
  End If
  
  intMsgType = vbExclamation
                         
  If lngErrNbrIn > vbObjectError Then
     intResponse = MsgBox(strMsg, intMsgType, "Program Error")
  Else
     MSErrorMsg = GetErrorMsg(lngErrNbrIn)
     strMsg = strMsg & _
             vbCrLf & "MS Error Message: " & MSErrorMsg & _
             vbCrLf
     If strHelpFile <> "" And _
        strHelpContext <> "" Then
        strMsg = strMsg & vbCrLf & "Press Help button or F1 for the Visual Basic Help" & _
             " topic for this error."
        intResponse = MsgBox(strMsg, intMsgType & vbMsgBoxHelpButton, "Error", Err.HelpFile, Err.HelpContext)
     Else
        intResponse = MsgBox(strMsg, intMsgType, "Error")
     End If
  End If
  
' I'm still trying to decide what to do here ????????????????????
' Should I give the user a choice of what to do or just END ??????????
   
      ' Return Value      Meaning
      ' 0                  Resume
      ' 1                  Resume Next
      ' 2                  Unrecoverable error
      ' 3                  Unrecognized error
   
' could probably use a RaiseError condition here
'  Select Case intResponse
'     Case 1, 4      ' OK, Retry buttons.
'        ErrMsg = 0
'     Case 5         ' Ignore button.
'        ErrMsg = 1
'     Case 2, 3      ' Cancel, End buttons.
'        ErrMsg = 2
'     Case Else
'        ErrMsg = 3
'  End Select
  
ProcExit:
  gstrChkPt = "On Error": gstrProcName = ""
  Err.Clear   ' Clear Err object properties
  Exit Function

ErrRtn:
  MsgBox ("An error occured in the error handling routine (" & Err.Source & ")" & _
          vbCrLf & vbCrLf & str(Err.Number) & " - " & Err.Description)
   Resume ProcExit:
End Function

Private Function GetErrorMsg(ErrNbr As Long) As String
   On Error GoTo ErrRtn
   gstrChkPt = "On Error": gstrProcName = "GetErrorMsg"
  
   Static sMsgBuf As String * 257
   Dim lngLen As Long

   GetErrorMsg = "Message Not Found"
   lngLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                          FORMAT_MESSAGE_IGNORE_INSERTS Or _
                          FORMAT_MESSAGE_MAX_WIDTH_MASK, ByVal 0&, ByVal ErrNbr, _
                          LANG_USER_DEFAULT, ByVal sMsgBuf, 256&, 0&)
   If lngLen Then GetErrorMsg = vba.Left$(sMsgBuf, lngLen)
  
ProcExit:
   Exit Function

ErrRtn:
    Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
    Resume ProcExit:
End Function
  
  
  



