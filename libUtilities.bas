Attribute VB_Name = "libUtilities"
Option Explicit

Const strModuleName As String = "libUtilities"

Public Enum ComputerNameFormat
    cnfComputerNameNetBIOS = 0
    cnfComputerNameDnsHostname = 1
    cnfComputerNameDnsDomain = 2
    cnfComputerNameDnsFullyQualified = 3
    cnfComputerNamePhysicalNetbios = 4
    cnfComputerNamePhysicalDnsHostname = 5
    cnfComputerNamePhysicalDnsDomain = 6
    cnfComputerNamePhysicalDnsFullyQualified = 7
    cnfComputerNameMax = 8
End Enum

Private Const dhcMaxComputerName = 15
 
Private Declare Function GetComputerNameEx _
   Lib "kernel32" Alias "GetComputerNameExA" _
   (ByVal NameType As ComputerNameFormat, ByVal lpBuffer As String, _
   nSize As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Sub CloseObject(objToClose As Object)
   gstrChkPt = "On Error": gstrProcName = "CloseObject"
   On Error GoTo ErrRtn

   If IsObject(objToClose) And Not objToClose Is Nothing Then
      objToClose.Close
      Set objToClose = Nothing
   End If
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
  
End Sub

Public Function GetUserID() As String
   gstrChkPt = "On Error": gstrProcName = "GetUserID"
   On Error GoTo ErrRtn
  
   Dim szBuffer As String
   Dim lBuffSize As Long
   Dim RetVal As Boolean

   szBuffer = Space(255)
   lBuffSize = Len(szBuffer)
   RetVal = GetUserName(szBuffer, lBuffSize)

   If RetVal Then
     '* Strip the null character from the User name
     GetUserID = vba.Left$(szBuffer, InStr(szBuffer, vbNullChar) - 1)
   Else
     GetUserID = "UnknownUser"
   End If
  
ProcExit:
   Exit Function

ErrRtn:
    Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
    Resume ProcExit:
End Function

Public Property Get ComputerName(Optional NameFormat As ComputerNameFormat = cnfComputerNameNetBIOS) As String

   ' Set or retrieve the NetBIOS name of the computer.
   Dim strBuffer As String
   Dim lngLen As Long
   Dim sTemp As String
    
   sTemp = String(255, 0)
   'get the computername
   GetComputerName sTemp, 255
   ComputerName = sTemp
'
'   If IsWin2000 Then
'      If NameFormat <> cnfComputerNameNetBIOS Then
'         ' If a particular NameFormat is requested and the
'         ' OS is Windows 2000, then use the Extended
'         ' version of the API function.
'
'         ' To determine the required buffer size for the
'         ' particular value of NameFormat, pass vbNullString
'         ' for strBuffer. When the function returns, lngLen will
'         ' contain the length of the required buffer.
'         Call GetComputerNameEx(NameFormat, vbNullString, lngLen)
'         strBuffer = String$(lngLen + 1, vbNullChar)
'         If CBool(GetComputerNameEx( _
'            NameFormat, strBuffer, lngLen)) Then
'            ComputerName = vba.left$(strBuffer, lngLen)
'         End If
'      Else
'         ' Specified NameFormat is cnfComputerNameNetBios
'         ' in which case, use GetComputerName API
'         strBuffer = String$(dhcMaxComputerName + 1, vbNullChar)
'         lngLen = Len(strBuffer)
'         If CBool(GetComputerName(strBuffer, lngLen)) Then
'            ' If successful, return the buffer
'             ComputerName = vba.left$(strBuffer, lngLen)
'         End If
'      End If
'   Else
'       ' The OS is not Win2000
'       ' Only cnfComputerNameNetBios is valid for NameFormat
'       If NameFormat = cnfComputerNameNetBIOS Then
'          strBuffer = String$(dhcMaxComputerName + 1, vbNullChar)
'          lngLen = Len(strBuffer)
'          If CBool(GetComputerName(strBuffer, lngLen)) Then
'             ' If successful, return the buffer
'             ComputerName = vba.left$(strBuffer, lngLen)
'          End If
'       Else
'          If RaiseErrors Then
'             Call HandleErrors(ERR_INVALID_OS)
'          End If
'       End If
'   End If
End Property

'Public Sub ClearText()
''  clears text of all bound controls - textboxes and combobox
'   gstrchkpt = "On Error": gstrProcName = "ClearText"
'   on error goto ErrRtn
'   Dim i As Integer
'
'   For i = 1 To Me.Controls.count - 1
'      If (TypeOf Me.Controls(i) Is TextBox) Then
'          Me.Controls(i).Text = ""
'      ElseIf (TypeOf Me.Controls(i) Is ComboBox) Then
'          Me.Controls(i).Text = ""
'      End If
'   Next i
'   cboContactName.SetFocus
'
'ProcExit:
'   Exit Function
'
'ErrRtn:
'   call errmsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
'   resume procexit:
'End Sub

Public Function AddToMsg(ByRef strMsg As String, stRStringToAdd As String)
    'Adds a string to a variable for formatted display by
    'the MsgBox function.
   gstrChkPt = "On Error": gstrProcName = "AddToMsg"
   On Error GoTo ErrRtn
    
   If strMsg = "" Then
      strMsg = stRStringToAdd
   Else
      strMsg = strMsg & Chr$(10) & stRStringToAdd
   End If
  
ProcExit:
   Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Function SelectAllTxt(CtlIn As Control)
   gstrChkPt = "On Error": gstrProcName = "SelecAllText"
   On Error GoTo ErrRtn
  
   If CtlIn <> "" Then
      CtlIn.SelStart = 1
      CtlIn.SelLength = Len(ctl)
   End If
  
ProcExit:
   Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function TrimSpaces(strTextIn As String) As String
   gstrChkPt = "On Error": gstrProcName = "TrimSpaces"
   On Error GoTo ErrRtn
  
   Dim lngCntr As Long
   Dim stRSpaceCheck As String
   Dim strFullString As String

   For lngCntr = 1 To Len(strTextIn)
      stRSpaceCheck = vba.Mid$(strTextIn, lngCntr, 1)
      If stRSpaceCheck <> " " Then
         strFullString = strFullString & stRSpaceCheck
      End If
   Next lngCntr
   TrimSpaces = strFullString
  
ProcExit:
   Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function
