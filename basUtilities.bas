Attribute VB_Name = "basUtilities"
Option Explicit

Const strModuleName As String = "basUtilities"

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
  On Error GoTo Err
  gstrProcName = "CloseObject"

  If IsObject(objToClose) And Not objToClose Is Nothing Then
    objToClose.Close
    Set objToClose = Nothing
  End If
  
ProcExit:
  Exit Sub

Err:
  Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
  Resume ProcExit:
  
End Sub

Public Function GetUserID() As String
  On Error GoTo Err
  gstrProcName = "GetUserID"
  
  Dim szBuffer As String
  Dim lBuffSize As Long
  Dim RetVal As Boolean

  szBuffer = Space(255)
  lBuffSize = Len(szBuffer)
  RetVal = GetUserName(szBuffer, lBuffSize)

  If RetVal Then
    '* Strip the null character from the User name
    GetUserID = VBA.Left$(szBuffer, InStr(szBuffer, vbNullChar) - 1)
  Else
    GetUserID = "UnknownUser"
  End If
  
ProcExit:
  Exit Function

Err:
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
'    If IsWin2000 Then
'        If NameFormat <> cnfComputerNameNetBIOS Then
'            ' If a particular NameFormat is requested and the
'            ' OS is Windows 2000, then use the Extended
'            ' version of the API function.
'
'            ' To determine the required buffer size for the
'            ' particular value of NameFormat, pass vbNullString
'            ' for strBuffer. When the function returns, lngLen will
'            ' contain the length of the required buffer.
'            Call GetComputerNameEx(NameFormat, vbNullString, lngLen)
'            strBuffer = String$(lngLen + 1, vbNullChar)
'            If CBool(GetComputerNameEx( _
'             NameFormat, strBuffer, lngLen)) Then
'                ComputerName = Left$(strBuffer, lngLen)
'            End If
'        Else
'            ' Specified NameFormat is cnfComputerNameNetBios
'            ' in which case, use GetComputerName API
'            strBuffer = String$(dhcMaxComputerName + 1, vbNullChar)
'            lngLen = Len(strBuffer)
'            If CBool(GetComputerName(strBuffer, lngLen)) Then
'                ' If successful, return the buffer
'                ComputerName = Left$(strBuffer, lngLen)
'            End If
'        End If
'    Else
'        ' The OS is not Win2000
'        ' Only cnfComputerNameNetBios is valid for NameFormat
'        If NameFormat = cnfComputerNameNetBIOS Then
'            strBuffer = String$(dhcMaxComputerName + 1, vbNullChar)
'            lngLen = Len(strBuffer)
'            If CBool(GetComputerName(strBuffer, lngLen)) Then
'                ' If successful, return the buffer
'                ComputerName = Left$(strBuffer, lngLen)
'            End If
'        Else
'            If RaiseErrors Then
'                Call HandleErrors(ERR_INVALID_OS)
'            End If
'        End If
'    End If
End Property

'Public Sub ClearText()
''clears text of all bound controls - textboxes and combobox
'  On Error GoTo Err
'  gstrProcName = "ClearText"
'    Dim i As Integer
'
'    For i = 1 To Me.Controls.count - 1
'        If (TypeOf Me.Controls(i) Is TextBox) Then
'            Me.Controls(i).Text = ""
'        ElseIf (TypeOf Me.Controls(i) Is ComboBox) Then
'            Me.Controls(i).Text = ""
'        End If
'    Next i
'    cboContactName.SetFocus
'
'ProcExit:
'  Exit Function
'
'Err:
'  Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
'  Resume ProcExit:
'End Sub

Public Function AddToMsg(ByRef strMsg As String, stRStringToAdd As String)
    'Adds a string to a variable for formatted display by
    'the MsgBox function.
  On Error GoTo Err
  gstrProcName = "AddToMsg"
    
  If strMsg = "" Then
      strMsg = stRStringToAdd
  Else
      strMsg = strMsg & Chr$(10) & stRStringToAdd
  End If
  
ProcExit:
  Exit Function

Err:
  Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
  Resume ProcExit:
End Function


Function SelectAllTxt(CtlIn As Control)
  On Error GoTo Err
  gstrProcName = "SelecAllText"
  
  If CtlIn <> "" Then
    CtlIn.SelStart = 1
    CtlIn.SelLength = Len(Ctl)
  End If
  
ProcExit:
  Exit Function

Err:
  Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
  Resume ProcExit:
End Function

Public Function TrimSpaces(strTextIn As String) As String
  
  On Error GoTo Err
  gstrProcName = "TrimSpaces"
  
  Dim lngCntr As Long
  Dim stRSpaceCheck As String
  Dim strFullString As String

  For lngCntr = 1 To Len(strTextIn)
    stRSpaceCheck = Mid(strTextIn, lngCntr, 1)
    If stRSpaceCheck <> " " Then
      strFullString = strFullString & stRSpaceCheck
    End If
  Next lngCntr
  TrimSpaces = strFullString
  
ProcExit:
  Exit Function

Err:
  Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
  Resume ProcExit:
End Function

Sub IndexObjectX()

   Dim dbsNorthwind As Database
   Dim tdfEmployees As TableDef
   Dim idxNew As Index
   Dim idxLoop As Index
   Dim rstEmployees As Recordset

   Set dbsNorthwind = OpenDatabase("Northwind.mdb")
   Set tdfEmployees = dbsNorthwind!Employees

   With tdfEmployees
      ' Create new index, create and append Field
      ' objects to its Fields collection.
      Set idxNew = .CreateIndex("NewIndex")

      With idxNew
         .Fields.Append .CreateField("Country")
         .Fields.Append .CreateField("LastName")
         .Fields.Append .CreateField("FirstName")
      End With

      ' Add new Index object to the Indexes collection
      ' of the Employees table collection.
      .Indexes.Append idxNew
      .Indexes.Refresh

      Debug.Print .Indexes.count & " Indexes in " & _
         .Name & " TableDef"

      ' Enumerate Indexes collection of Employees
      ' table.
      For Each idxLoop In .Indexes
         Debug.Print "  " & idxLoop.Name
      Next idxLoop

      Set rstEmployees = _
         dbsNorthwind.OpenRecordset("Employees")

      ' Print report using old and new indexes.
      IndexOutput rstEmployees, "PrimaryKey"
      IndexOutput rstEmployees, idxNew.Name
      rstEmployees.Close

      ' Delete new Index because this is a
      ' demonstration.
      .Indexes.Delete idxNew.Name
   End With

   dbsNorthwind.Close

End Sub

Sub IndexOutput(rstTemp As Recordset, _
                strIndex As String)
   ' Report function for FieldX.

   With rstTemp
      ' Set the index.
      .Index = strIndex
      .MoveFirst
      Debug.Print "Recordset = " & .Name & _
         ", Index = " & .Index
      Debug.Print "  EmployeeID - Country - Name"

      ' Enumerate the recordset using the specified
      ' index.
      Do While Not .EOF
         Debug.Print "  " & !EmployeeID & " - " & _
            !Country & " - " & !LastName & ", " & !FirstName
         .MoveNext
      Loop

   End With

End Sub

Public Function TrimSpaces(Text As String) As String
    Dim Loop1 As Long, SpaceCheck As String
    Dim FullString As String


    For Loop1& = 1 To Len(Text$)
        SpaceCheck$ = Mid(Text$, Loop1&, 1)

        If SpaceCheck$ <> " " Then
            FullString$ = FullString$ & SpaceCheck$
        End If
    Next Loop1&
    TrimSpaces = FullString$
End Function




