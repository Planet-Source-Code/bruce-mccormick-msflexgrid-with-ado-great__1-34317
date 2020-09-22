Attribute VB_Name = "basADO"
Option Explicit

Private Const strModuleName As String * 30 = "basADO"

Private errLoop

Private blnNoMatch  As Boolean
Private fldFields() As Field
Private intCnt As Integer

Public Enum e_DBTypes
    dbt_Undefined = 0
    dbt_OracleMSDA = 1
    dbt_OracleODBC = 2
    dbt_SQLserver = 3
    dbt_MSAccessFile = 5
    dbt_MSAccess97File = 6
    dbt_MSAccess2KFile = 7
    dbt_DSNFile = 8
    dbt_dbase = 9
End Enum

Public Function ConnOpen(cnIn As ADODB.Connection, _
                         DBTypeIn As e_DBTypes, _
                         ByVal ServerOrFileIn As String, _
                         Optional ByVal DBPathIn As String = "", _
                         Optional CommandTypeIn As CommandTypeEnum = adCmdStoredProc, _
                         Optional CursorLocIn As CursorLocationEnum = adUseClient, _
                         Optional ByVal UserNameIn As String = "", _
                         Optional ByVal PasswordIn As String = "") As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ConnOpen"
   
  Set cnIn = New ADODB.Connection
  If DBPathIn = vbNullString Then DBPathIn = ServerOrFileIn
   
  With cnIn
     .CursorLocation = CursorLocIn '/* default = adUseClient(3)
     .Open BuildConnStr(DBTypeIn, ServerOrFileIn, cnIn, DBPathIn, UserNameIn, PasswordIn)
  End With
  
ProcExit:
  Exit Function

ErrRtn:
   On Error Resume Next

   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL: " & errLoop.SQLState & " Native: " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:
End Function

Public Function BuildConnStr(ByVal DBTypeIn As e_DBTypes, _
                             ByVal ServerOrFileIn As String, _
                             cnIn As ADODB.Connection, _
                             Optional ByVal DBNameIn As String, _
                             Optional ByVal UserNameIn As String, _
                             Optional ByVal PasswordIn As String) As String
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "BuildConnStr"
    
  Select Case DBTypeIn
    Case dbt_OracleMSDA
      BuildConnStr = "Provider=MSDAORA;Data Source=" & ServerOrFileIn & ";User ID=" & _
                     IIf(UserNameIn <> "", UserNameIn, "") & ";PasswordIn=" & _
                     IIf(PasswordIn <> "", PasswordIn, "") & ";" & _
                     IIf(DBNameIn <> "", "Initial Catalog=" & DBNameIn & ";", "")
    Case dbt_OracleODBC
      BuildConnStr = "DRIVER={Microsoft ODBC for Oracle};SERVER=" & ServerOrFileIn & _
                     ";UID=" & UserNameIn & ";PWD=" & PasswordIn & ";" & _
                     IIf(DBNameIn <> "", "Initial Catalog=" & DBNameIn & ";", "")
    Case dbt_SQLserver
      BuildConnStr = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & _
                     ServerOrFileIn & ";User ID=" & IIf(UserNameIn <> "", UserNameIn, "") & _
                     ";PasswordIn=" & IIf(PasswordIn <> "", PasswordIn, "") & ";" & _
                     IIf(DBNameIn <> "", "Initial Catalog=" & DBNameIn & ";", "")
      ' or "MSDASQL;Driver={SQL Server};SERVER=" & ServerName & ";user id=" & uname & _
      '    ";Password=" & pass & ";Database=" & db & ""
    
    Case dbt_DSNFile
      BuildConnStr = "Provider=MSDASQL;DSN=" & ServerOrFileIn & ";UID=" & _
                     IIf(UserNameIn <> "", UserNameIn, "") & ";PWD=" & _
                     IIf(PasswordIn <> "", PasswordIn & ";", "") & ";" & _
                     IIf(DBNameIn <> "", "Initial Catalog=" & DBNameIn & ";", "")
    Case dbt_MSAccess2KFile, dbt_MSAccess97File, dbt_MSAccessFile
      BuildConnStr = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & ServerOrFileIn & _
                     ";DefaultDir=" & basFiles.RetPathOnly(ServerOrFileIn) & ";PWD=" & _
                     IIf(PasswordIn <> "", PasswordIn & ";", ";")
    Case dbt_dbase
      BuildConnStr = "Provider=MSDASQL.1;Data Source=dBASE Files;" & IIf(DBNameIn <> "", _
                     "Initial Catalog=" & DBNameIn & ";", "")
                ' or "MSDASQL.1;Data Source=dBASE Files;Initial Catalog=" & db
  End Select
  
ProcExit:
  Exit Function

ErrRtn:
   On Error Resume Next

   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:
End Function
  
Public Sub ConnClose(cnIn As ADODB.Connection)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "basADO.ConnClose"

   If cnIn Is Nothing Then Exit Sub
   Set cnIn = Nothing
  
ProcExit:
   Exit Sub

ErrRtn:
   On Error Resume Next

   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:
End Sub

Public Function rsOpen(RSIn As ADODB.Recordset, _
                       cnIn As ADODB.Connection, _
                       Optional LockTypeIn As Integer = adLockPessimistic, _
                       Optional CursorLocIn As Integer = adUseClient, _
                       Optional CursorTypeIn As Integer = adOpenDynamic) As Boolean
   ' alternate lock options: adLockReadOnly, adLockOptimistic                                                         '
   
   On Error GoTo ErrRtn
   gstrChkPt = "On Error": gstrProcName = "rsOpen"

   Set RSIn = New ADODB.Recordset
    
   With RSIn
     .LockType = LockTypeIn
     .CursorLocation = CursorLocIn
     .CursorType = CursorTypeIn
     .ActiveConnection = cnIn
     .Source = gstrOpenStmt
     .Open
     If Not .BOF And Not .EOF Then ' to set record count
        .MoveLast
        .MoveFirst
     End If
   End With

ProcExit:
  Exit Function

ErrRtn:
   On Error Resume Next
   
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)

   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:
End Function

Public Sub rsClose(RSIn As ADODB.Recordset, cnIn As ADODB.Connection)
   On Error GoTo ErrRtn
   gstrChkPt = "On Error": gstrProcName = "rsClose"
   
   If RSIn Is Nothing Then Exit Sub
   gstrChkPt = "Set RSIn = Nothing"
   Set RSIn = Nothing

ProcExit:
   gstrChkPt = "Exit Sub"
   Exit Sub
   gstrChkPt = "After Exit Sub"

ErrRtn:
   On Error Resume Next
   
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)

   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:
End Sub

Public Sub rsDelete(RSIn As ADODB.Recordset, cnIn As ADODB.Connection, _
                    Optional EOFActionIn As String = "Last", _
                    Optional ShowMsgIn As Boolean = False)
   On Error GoTo ErrRtn
   gstrChkPt = "On Error": gstrProcName = "rsDelete"
         
   Call basADO.rsOpen(RSIn, cnIn)
   gstrChkPt = "If rsIn.RecordCount": gstrProcName = "rsDelete"
   If RSIn.RecordCount = 0 Then
      If ShowMsgIn = True Then
         MsgBox "There Is No Record To Delete."
      End If
      Exit Sub
   End If
   
   gstrChkPt = "rsIn.Delete"
   RSIn.Delete
   
   gstrChkPt = "MvNext": gstrProcName = "rsDelete"
   RSIn.MoveNext
   
   If RSIn.EOF And RSIn.BOF Then
      If ShowMsgIn = True Then
         MsgBox ("The File Is Now Empty. You Must Add Records To Continue.")
      End If
      Exit Sub
   End If
      
   If RSIn.EOF Then
      If EOFActionIn = "Last" Then
         If ShowMsgIn Then
            MsgBox ("You Are At The End Of The File. Staying On Last Record.")
         End If
         gstrChkPt = "MovePrevious": gstrProcName = "rsDelete"
         RSIn.MovePrevious
      ElseIf EOFActionIn = "First" Then
         If ShowMsgIn Then
            MsgBox ("You Are At The End Of The File. Going To First Record.")
         End If
         gstrChkPt = "MoveFirst": gstrProcName = "rsDelete"
         RSIn.MoveFirst
      End If
   End If
   
   gstrChkPt = "Close": gstrProcName = "rsDelete"
   Call basADO.rsClose(RSIn, cnIn)

ProcExit:
   gstrChkPt = "Exit Sub"
   Exit Sub

ErrRtn:
   On Error Resume Next
   
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)

   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:
End Sub

Public Function MvNext(RSIn As ADODB.Recordset, _
                       cnIn As ADODB.Connection, _
                       Optional EOFActionIn As String = "NoAction", _
                       Optional ShowMsgIn As Boolean = False)
   ' the EOFActionIn gives you flexability
   ' In some cases you just want to return the EOF marker to the calling proc
   ' In some cases you may want to go to the first or last rec instead
  
   On Error GoTo ErrRtn
   gstrChkPt = "On Error": gstrProcName = "MvNext"
                              
   If Not RSIn.EOF Then
      RSIn.MoveNext
   Else
      If RSIn.BOF Then
         MsgBox ("The File Is Now Empty. You Must Add Records To Continue.")
         Exit Function
      End If
      
      If EOFActionIn = "Last" Then
         If ShowMsgIn Then
            MsgBox ("You Are At The End Of The File. Staying On Last Record.")
         End If
         RSIn.MovePrevious
      ElseIf EOFActionIn = "First" Then
         If ShowMsgIn Then
            MsgBox ("You Are At The End Of The File. Going To First Record.")
         End If
         RSIn.MoveFirst
      End If
   End If
  
ProcExit:
   Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
  
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
  
  Resume ProcExit:

End Function

Public Function MvPrev(RSIn As ADODB.Recordset, _
                       cnIn As ADODB.Connection, _
                       Optional BOFActionIn As String = "NoAction")
   ' the BOFActionIn gives you flexability
   ' In some cases you just want to return the BOF marker to the calling proc
   ' In some cases you may want to go to the first or last rec instead
   
   On Error GoTo ErrRtn
   gstrChkPt = "On Error": gstrProcName = "MvPrev"
                               
   RSIn.Moveprev
   
   If rs.BOF Then
      If RSIn.EOF Then
         MsgBox ("The File Is Now Empty. You Must Add Records To Continue.")
         Exit Function
      End If
      
      If BOFActionIn = "Last" Then
         MsgBox ("You Are At The Beginning Of The File. Going To Last Record.")
         RSIn.MoveLast
      ElseIf BOFActionIn = "First" Then
         MsgBox ("You Are At The Beginning Of The File. Staying On The First Record.")
         RSIn.MoveFirst
      End If
   End If
  
ProcExit:
   Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
  
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
  
   Resume ProcExit:
End Function

Public Function ADODatabaseConnected(cnIn As ADODB.Connection, _
                                     DBPathIn As String, _
                                     DBNameIn As String) As Boolean
 'function is separated to trap multiple connection error
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "DatabaseConnected"
  
 ADODatabaseConnected = False

RetryConnect:

  cnIn.Open "PROVIDER = Microsoft.Jet.OLEDB.4.0; Data Source = " & DBPathIn & ";"
  cnEOFAction = vbEOFActionAddNew
  cn.Refresh
  
  If RSIn.EOF And _
     RSIn.BOF Then
     MsgBoxAns = MsgBox("The Database Does Not Contain A Beginning Dataset. " & _
                        "You Must Add Records Before Doing Anything Else.", _
                        vbInformation & vbOKOnly, _
                        "Empty Dataset")
  End If
  
  ADODatabaseConnected = True 'function is set to true if connection is successfull
    
ProcExit:
  Exit Function

ErrRtn:
   On Error Resume Next

   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:
End Function

Public Function Check4Records(cnIn As ADODB.Connection, _
                              SQLIn As String) As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Check4Records"
  
  Check4Records = False
  
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
  With rs
    .ActiveConnection = cnIn
    .LockType = adLockBatchOptimistic
    .CursorLocation = adUseClient
    .Open SQLIn
  
    If .EOF And .BOF Then
      rs.Close
      Set rs = Nothing
      Exit Function
    End If
  
  End With
  
  rs.Close
  Set rs = Nothing
  
  Check4Records = True
  
  Exit Function
  
ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
 
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
 
End Function

Private Sub Adodc1_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, _
                                    ByVal cRecords As Long, _
                                    adStatus As ADODB.EventStatusEnum, _
                                    ByVal pRecordset As ADODB.Recordset, _
                                    cnIn As ADODB.Connection)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Adodc1_WillChangeRecord"
  
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
         Case adRsnAddNew
         Case adRsnClose
         Case adRsnDelete
         Case adRsnFirstChange
         Case adRsnMove
         Case adRsnRequery
         Case adRsnResynch
         Case adRsnUndoAddNew
         Case adRsnUndoDelete
         Case adRsnUndoUpdate
         Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Sub

Public Property Let Fields(Index As Variant, NewValue As Variant)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Let Fields"
    
  fldFields(Index) = NewValue
  
ProcExit:
  Exit Property

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Property

Public Property Get Fields(Index As Variant) As Field
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Get Fields"
    
  Set Fields = fldFields(Index)
  
ProcExit:
  Exit Property

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Property

Public Property Let ProperADORecordset(RSIn As ADODB.Recordset, _
      Index As Variant, NewValue As Variant)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Let proper ADORecordset"
  
  RSIn(Index) = NewValue
  
ProcExit:
  Exit Property

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Property

Public Property Get ProperADORecordset(RSIn As ADODB.Recordset, _
   Index As Variant) As Variant
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Get ProperADORecordset"
  
  ProperADORecordset = RSIn(Index)
  
ProcExit:
  Exit Property

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
      
   Resume ProcExit:

End Property

Public Function RecordCount(RSIn As ADODB.Recordset, _
                            cnIn As ADODB.Connection) As Long
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "RecordCount"
  
  RSIn.MoveLast
  RSIn.MoveFirst
  RecordCount = RSIn.RecordCount
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
Resume ProcExit:

End Function

Public Sub FindFirst(RSIn As ADODB.Recordset, _
                     cnIn As ADODB.Connection, Filter As String)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindFirst"
  
  Dim rsClone As ADODB.Recordset

  blnNoMatch = True
  'If (InStr(Filter, "(") > 0) Or (InStr(Filter, ")") > 0) Or _
  '   (InStr(UCase(Filter), " AND ") > 0) Or (InStr(UCase(Filter), " OR ") > 0) Then
     Set rsClone = New ADODB.Recordset
     Set rsClone = RSIn.Clone
     rsClone.Filter = Filter
     If (rsClone.RecordCount > 0) Then
        rsClone.MoveFirst
        RSIn.BookMark = rsClone.BookMark
        blnNoMatch = False
     Else
        If (rs.RecordCount > 0) Then
            RSIn.MoveLast: RSIn.MoveNext
            blnNoMatch = True
        End If
     End If
  'Else
  '    rsIn.Find Filter
  '    blnNoMatch = rsIn.EOF
  'End If
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Sub

Public Sub FindLast(RSIn As ADODB.Recordset, _
                    cnIn As ADODB.Connection, FilterGroupEnum As String)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindLast"
  
Dim rsClone As ADODB.Recordset

blnNoMatch = True
'If (InStr(Filter, "(") > 0) Or (InStr(Filter, ")") > 0) Or _
'   (InStr(UCase(Filter), " AND ") > 0) Or (InStr(UCase(Filter), " OR ") > 0) Then
    Set rsClone = New ADODB.Recordset
    Set rsClone = RSIn.Clone
    rsClone.Filter = Filter
    If (rsClone.RecordCount > 0) Then
        rsClone.MoveLast
        RSIn.BookMark = rsClone.BookMark
        blnNoMatch = False
    Else
        If (rs.RecordCount > 0) Then
            RSIn.MoveLast: RSIn.MoveNext
            blnNoMatch = True
        End If
    End If
'Else
'    rsIn.Find Filter
'    rsIn.MoveLast
'    blnNoMatch = rsIn.BOF
'End If
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Sub

Public Sub FindNext(RSIn As ADODB.Recordset, _
                    cnIn As ADODB.Connection, Filter As String)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindNext"

  Dim rsClone As ADODB.Recordset

  blnNoMatch = True
  'If (InStr(Filter, "(") > 0) Or (InStr(Filter, ")") > 0) Or _
  '   (InStr(UCase(Filter), " AND ") > 0) Or (InStr(UCase(Filter), " OR ") > 0) Then
    Set rsClone = New ADODB.Recordset
    Set rsClone = RSIn.Clone
    rsClone.Filter = Filter
    rsClone.Sort = RSIn.Sort
    If (rsClone.RecordCount > 0) Then
        rsClone.BookMark = RSIn.BookMark
        rsClone.MoveNext
        If (Not rsClone.EOF) Then
            RSIn.BookMark = rsClone.BookMark
            blnNoMatch = False
        Else
            blnNoMatch = True
        End If
    Else
        If (rs.RecordCount > 0) Then
            RSIn.MoveLast: RSIn.MoveNext
            blnNoMatch = True
        End If
    End If
'Else
'    rsIn.Find Filter
'    blnNoMatch = rsIn.EOF
'End If
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Sub

Public Sub FindPrevious(RSIn As ADODB.Recordset, _
                        cnIn As ADODB.Connection, Filter As String)
   On Error GoTo ErrRtn
   gstrChkPt = "On Error": gstrProcName = "FindPrevious"
   Dim rsClone As ADODB.Recordset

   blnNoMatch = True
   'If (InStr(Filter, "(") > 0) Or (InStr(Filter, ")") > 0) Or _
   '   (InStr(UCase(Filter), " AND ") > 0) Or (InStr(UCase(Filter), " OR ") > 0) Then
      Set rsClone = New ADODB.Recordset
      Set rsClone = RSIn.Clone
      rsClone.Filter = Filter
      rsClone.Sort = RSIn.Sort
      If (rsClone.RecordCount > 0) Then
         rsClone.BookMark = RSIn.BookMark
         rsClone.MovePrevious
         If (Not rsClone.BOF) Then
            RSIn.BookMark = rsClone.BookMark
            blnNoMatch = False
         Else
            blnNoMatch = True
         End If
      Else
         If (rs.RecordCount > 0) Then
             RSIn.MoveFirst: RSIn.MovePrevious
             blnNoMatch = True
         End If
     End If
  'Else
  '    rsIn.Find Filter
  '    blnNoMatch = rsIn.BOF
'End If
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Sub

Public Property Get NoMatch(cnIn As ADODB.Connection) As Variant
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Get NoMatch"
  
  NoMatch = blnNoMatch
  
ProcExit:
  Exit Property

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Property

Public Property Get BookMark(RSIn As ADODB.Recordset, _
                             cnIn As ADODB.Connection) As Variant
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Get BookMark"
  
  BookMark = RSIn.BookMark
  
ProcExit:
  Exit Property

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:

End Property

Public Property Let BookMark(RSIn As ADODB.Recordset, _
                             cnIn As ADODB.Connection, ByVal vNewValue As Variant)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "MovePrevious"
  
  RSIn.BookMark = "Let BookMark"
  
ProcExit:
  Exit Property

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Property

Public Function FindLike(RSIn As ADODB.Recordset, _
                         cnIn As ADODB.Connection, _
                         FileNameIn As String, FieldNameIn As String, _
                         LikeStrIn1 As String, WildCardIn As Boolean, _
                         Optional LikeStrIn2 As String)
  ' we need likestrin2 in case the wildcard occurs in the middle of the string
  ' exp "Select * from customers where lastname like 'SMI_H'"
  ' for ADO the wildcards are "_" for single char and "%" for multiple char
  ' for DAO the wildcards are "?" for single char and "*" for multiple char
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindLike"
  
  Dim rsClone As ADODB.Recordset
  Dim Suffix As String * 1
  
  Set rsClone = New ADODB.Recordset
  Set rsClone = RSIn.Clone
  gstrOpenStmt = "Select * from '" & FileNameIn & "' where '" & FieldNameIn & _
                 "' like '" & LikeStrIn1 & WildCardIn
  If Not IsMissing(LikeStrIn2) Then gstrOpenStmt = gstrOpenStmt & likestr2
  gstrOpenStmt = gstrOpenStmt & "'"
  
  rsClone.Filter = gstrOpenStmt
     
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Function

Public Sub Find(RSIn As ADODB.Recordset, _
                cnIn As ADODB.Connection, _
                Criteria As String, Optional SkipRows As Long = 0, _
                Optional SearchDirection As SearchDirectionEnum = adSearchForward, _
                Optional Start As Variant)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Find"
Dim rsClone As ADODB.Recordset
Dim Cnt As Integer

blnNoMatch = True
If (InStr(Criteria, "(") > 0) Or (InStr(Criteria, ")") > 0) Or (InStr(UCase(Criteria), " AND ") > 0) Or (InStr(UCase(Criteria), " OR ") > 0) Then
    Set rsClone = New ADODB.Recordset
    Set rsClone = RSIn.Clone
    rsClone.Filter = Criteria
    If (rsClone.RecordCount > 0) Then
        If (Not IsMissing(Start)) Then
            rsClone.BookMark = Start
            If (SearchDirection = adSearchForward) Then
                For Cnt = 0 To SkipRows
                    If (Not rsClone.EOF) Then rsClone.MoveNext Else Exit For
                Next Cnt
                If (Not rsClone.EOF) Then
                    RSIn.BookMark = rsClone.BookMark
                    blnNoMatch = False
                End If
            Else
                For Cnt = 0 To SkipRows
                    If (Not rsClone.BOF) Then rsClone.MovePrevious Else Exit For
                Next Cnt
                If (Not rsClone.BOF) Then
                    RSIn.BookMark = rsClone.BookMark
                    blnNoMatch = False
                End If
            End If
        Else
            rsClone.MoveFirst
            RSIn.BookMark = rsClone.BookMark
            blnNoMatch = False
        End If
    Else
        If (rs.RecordCount > 0) Then
            RSIn.MoveLast: RSIn.MoveNext
            blnNoMatch = True
        End If
    End If
Else
    RSIn.Find Criteria
End If
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Sub

Public Function AbsolutePosition(RSIn As ADODB.Recordset, _
                                 cnIn As ADODB.Connection)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "AbsolutePosition"
    
  AbsolutePosition = RSIn.AbsolutePosition
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Function

Public Sub Delete(RSIn As ADODB.Recordset, _
                  cnIn As ADODB.Connection, _
                  Optional AffectRecords As AffectEnum = adAffectCurrent)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Delete"
  With RSIn
    .Delete AffectRecords
    .MoveNext
    If .EOF Then .MoveLast
  End With
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Sub

Public Function Index(RSIn As ADODB.Recordset, _
                      cnIn As ADODB.Connection) As Long
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Index"

  Index = RSIn.Index
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Function

Public Property Get Filter(RSIn As ADODB.Recordset, _
                           cnIn As ADODB.Connection) As Variant
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Get Filter"
    
  Filter = RSIn.Filter
  
ProcExit:
  Exit Property

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Property

Public Property Let Filter(RSIn As ADODB.Recordset, _
                           cnIn As ADODB.Connection, ByVal vNewValue As Variant)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Let Filter"
  
  RSIn.Filter = vNewValue
  
ProcExit:
  Exit Property

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Property

Public Function Clone(RSIn As ADODB.Recordset, _
                      cnIn As ADODB.Connection, _
                      Optional LockType As LockTypeEnum = adLockUnspecified) _
                      As ADODB.Recordset
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Clone"
    
  Set Clone = RSIn.Clone(LockType)
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Function

Public Function Clone2(RSIn As ADODB.Recordset, _
                       cnIn As ADODB.Connection, _
                       Optional ByVal LockType As ADODB.LockTypeEnum = adLockBatchOptimistic) _
                       As ADODB.Recordset
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Clone2"
    'RETURNS A CLONE (COPY OF AN EXISTING RECORDSET)
        
    Dim objNewRS As ADODB.Recordset
    Dim objField As Object
    Dim lngCnt As Long
    On Error GoTo LocalError
    
    Set objNewRS = New ADODB.Recordset
    objNewRS.CursorLocation = adUseClient
    objNewRS.LockType = LockType

    For Each objField In objRecordset.Fields
            objNewRS.Fields.Append objField.Name, objField.Type, objField.DefinedSize, objField.Attributes
    Next objField

    If Not objRecordset.RecordCount = 0 Then
            Set objNewRS.ActiveConnection = objRecordset.ActiveConnection
            objNewRS.Open
          
        objRecordset.MoveFirst
        While Not objRecordset.EOF
              objNewRS.AddNew
            For lngCnt = 0 To objRecordset.Fields.count - 1
                objNewRS.Fields(lngCnt).Value = objRecordset.Fields(lngCnt).Value
            Next lngCnt
            objRecordset.MoveNext
        Wend
    objNewRS.MoveFirst
    End If
    
    Set Clone = objNewRS
    Exit Function
ProcExit:
  Exit Function

ErrRtn:

   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Function

Public Function Execute(gSQLIn As String, cnIn As ADODB.Connection) As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Execute"
    'TO DIRECTLY EXECUTE AN INSERT, UPDATE, OR DELETE
    'SQL STATMENT. SET THE CONNECTION STRING PROPERTY
    'TO A VALID CONNECTION STRING FIRST
    
    On Error GoTo LocalError
    Dim cn As New ADODB.Connection
    With cn
        .ConnectionString = ConnectionString
        .CursorLocation = adUseServer
        .Open
        .BeginTrans
        .Execute SQL
        .CommitTrans
        .Close
    End With
    Set cn = Nothing
    Execute = True

ProcExit:
  Exit Function

ErrRtn:
    If cn.State = adStateOpen Then
        cn.RollBackTrans
        cn.Close
    End If
    Set cn = Nothing
    Execute = False
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Function

Public Function GetCount(TableName As String, _
                         cnIn As ADODB.Connection, _
                         Optional WhereClause As String = "") As Long
   On Error GoTo ErrRtn
   gstrChkPt = "On Error": gstrProcName = "GetCount"
    
    'RETURNS COUNT OF RECORDS WITHIN A TABLE, WITH OPTIONAL WHERE CLAUSE
        
   On Error GoTo LocalError
   Dim rs  As New ADODB.Recordset
   Dim SQL As String
   GetCount = 0
   If WhereClause <> "" Then
       SQL = "Select COUNT (*) FROM " & TableName & " WHERE " & WhereClause
   Else
       SQL = "Select COUNT (*) FROM " & TableName
   End If
   With rs
      .ActiveConnection = ConnectionString
      .CursorLocation = adUseClient
      .LockType = adLockReadOnly
      .CursorType = adOpenKeyset
      .Source = SQL
      .Open
      Set .ActiveConnection = Nothing
   End With
   GetCount = RSIn.Fields(0).Value
   Set rs = Nothing
  
ProcExit:
   Exit Function

ErrRtn:
   Set rs = Nothing
   GetCount = -1
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Function

Public Function PutRS(RSIn As ADODB.Recordset, cnIn As ADODB.Connection) As Boolean
  'USE THIS TO UPDATE A RECORDSET IN BATCH (TRANSACTIONAL) MODE
  'IF CHANGES TO THE RECORDSET'S WERE MADE PRIOR TO THIS CALL
  'THIS FUNCTION WILL COMMIT THEM TO THE UNDERYLING DATABASE
  
   On Error GoTo ErrRtn
   gstrChkPt = "On Error": gstrProcName = "PutRS"

   PutRS = False
   If EmptyRS(RSIn) Then
       Exit Function
   ElseIf RSIn.LockType = adLockReadOnly Then
       Exit Function
   Else
   
      Dim cn As New ADODB.Connection
        
      With cn
         .ConnectionString = ConnectionString
         .CursorLocation = adUseServer
         .Open
         BeginTrans
      End With
        
      With rs
         .ActiveConnection = cn
         .UpdateBatch
         cn.CommitTrans
         Set .ActiveConnection = Nothing
      End With
        
      cn.Close
      Set cn = Nothing
   End If
   
   PutRS = True
  
ProcExit:
  Exit Function

ErrRtn:
   If cn.State = adStateOpen Then
      cn.RollBackTrans
      cn.Close
   End If
   Set cn = Nothing
   PutRS = False
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:
End Function

Public Function EmptyRS(RSIn As ADODB.Recordset, cnIn As ADODB.Connection) As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "EmptyRS"

    EmptyRS = True
    If Not RSIn Is Nothing Then
        EmptyRS = ((RSIn.BOF = True) And (RSIn.EOF = True))
    End If

ProcExit:
  Exit Function

ErrRtn:
   EmptyRS = True
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Function

Private Sub Adodc1_MoveComplete(cnIn As ADODB.Connection, _
                                ByVal adReason As ADODB.EventReasonEnum, _
                                ByVal pError As ADODB.Error, _
                                adStatus As ADODB.EventStatusEnum, _
                                ByVal pRecordset As ADODB.Recordset)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Adodc1_MoveComplete"
   
   'This will display the current record position for this recordset and the records count
   Adodc1.Caption = "Record: " & CStr(Adodc1.Recordset.AbsolutePosition) & " of " & Adodc1.Recordset.RecordCount
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Sub

Public Function sqlBoolean(TrueFalse As Boolean, cnIn As ADODB.Connection) As Integer
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "salBoolean"
  
  'CONVERTS BIT RETURN VALUE FROM SQL SERVER
    
  'This is because SQL True = 1
  'VB True = -1
  sqlBoolean = TrueFalse
  If isSQL Then
     If TrueFalse = True Then sqlBoolean = TrueFalse * TrueFalse
  End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Function

Public Function sqlEncode(SQLIn As String, cnIn As ADODB.Connection) As String
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "sqlEncode"

    'IF A STRING VALUE IN AN SQL STATMENT HAS A ' CHARACTER,
    'USE THIS FUNCTION SO THE STRING CAN BE USED IN THE STATEMENT
     sqlEncode = vba.replace(gSQLInValue, "'", "''")
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:

End Function

Public Function ExecuteID(SQLIn As String, cnIn As ADODB.Connection) As Long
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ExecuteID"
'PURPOSE: RETURN VALUE OF IDENTITY COLUMN OF A NEWLY INSERTED RECORD

'SQL is a valid Insert statement.
'ConnetionString properyt has been set to a valid Connection String
'Tested on SQL7 as well as ACCESS 2000 using Jet4

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim AutoID As Long

With rs

 'Prepare the RecordSet
 .CursorLocation = adUseServer
 .CursorType = adOpenForwardOnly
 .LockType = adLockReadOnly
 .Source = "SELECT @@IDENTITY"
End With

With cn
 .ConnectionString = ConnectionString
 .CursorLocation = adUseServer
 .Open
 .BeginTrans
 .Execute SQL, , adCmdText + adExecuteNoRecords

    With rs
      .ActiveConnection = cn
      .Open , , , , adCmdText
      AutoID = rs(0).Value
      .Close
    End With
 .CommitTrans
 .Close
End With
Set rs = Nothing
Set cn = Nothing
 'If we get here ALL was Okay
 ExecuteID = AutoID
  
ProcExit:
  Exit Function

ErrRtn:
   ExecuteID = 0
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
      If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
Resume ProcExit:
End Function

Function Datashape(cnIn As ADODB.Connection, _
                   ByVal tblParent As String, _
                   ByVal tblChild As String, _
                   ByVal fldParent As String, _
                   ByVal fldChild As String, _
                   Optional ordParent As String = "", _
                   Optional ordChild As String = "", _
                   Optional WhereParent As String = "", _
                   Optional WhereChild As String = "", _
                   Optional ParentFields As String = "*", _
                   Optional ChildFields As String = "*", _
                   Optional MaxRecords As Long = 0) As ADODB.Recordset
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "DataShape"
    '=========================================================
    'This function will return a DisConnected SHAPEd RecordSet
    'Assumptions:
    '
    'tblParent      = Valid Table in the Database   - String \ Parent Table
    'tblChild       = Valid Table in the Database   - String / Child  Table
    '
    'fldParent      = Valid Field in Parent Table   - String \ relate this field
    'fldChild       = Valid Field in Child Table    - String / ..to this field
    '
    'ordParent      = Valid Field in Parent Table   - String \ ordering
    'ordChild       = Valid Field in Child Table    - String /
    '
    'WhereParent    = Valid SQL Where Clause        - Variant (Optional)
    'WhereChild     = Valid SQL Where Clause        - Variant (Optional)
    '
    'ParentFields   = Specific Fields to return     - String (pipe delimitered)
    'ChildFields    = Specific Fields to return     - String (pipe delimitered)
    'MaxRecords     = Specify Maximum Child Records - Long (0 = ALL)
    
    'NOTE: You may have to change connection string:  Normal Connection Strings
    'Begin with "Provider=". For the MsDataShape Provider, the connection string
    'begins with "Data Provider = "
    
    'EXAMPLE: THIS RETURNS A HYPOTHETICAL RECORDSET OF CUSTOMERS,
    'WHERE ONE OF THE MEMBERS IS A HYPOTHETICAL CHILD RECORDSET
    'OF THE CUSTOMERS' ORDERS
    
    'Dim sShapeConnectionString As String
    'Dim oCustRs As ADODB.Recordset
    'Dim oOrderRs As ADODB.Recordset
    'Dim oADO As New AdoUtils
    'Dim sSQL As String

    'sShapeConnectionString = "Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyBusiness.mdb"
    'sSQL = "SELECT * FROM CUSTOMERS"
    'With oTest
    '   .ConnectionString = sShapeConnectionString
    '   Set oCustRs = .Datashape("Customers", "Orders", "ID", "CustomerID")
    '   Set oOrderRs = ors.Fields(ors.Fields.Count - 1).Value
    'End With
    
    
    '=========================================================

    Dim cn        As ADODB.Connection
    Dim rs        As ADODB.Recordset
    Dim lSQL      As String
    Dim pSQL      As String
    Dim cSQL      As String
    Dim pWhere    As String
    Dim cWhere    As String
    Dim pOrder    As String
    Dim cOrder    As String

    'Define the SQL Statement
    lSQL = ""
    ParentFields = vba.replace(ParentFields, "|", ", ")
    ChildFields = vba.replace(ChildFields, "|", ", ")
    pWhere = WhereParent
    cWhere = WhereChild
    pOrder = ordParent
    cOrder = ordChild

    If WhereParent <> "" Then WhereParent = " WHERE " & WhereParent
    If WhereChild <> "" Then WhereChild = " WHERE " & WhereChild
    If pOrder <> "" Then pOrder = " ORDER By " & pOrder
    If cOrder <> "" Then cOrder = " ORDER By " & cOrder
    'Define Parent SQL Statement
    pSQL = ""
    If MaxRecords > 0 Then
        If isSQL Then
            pSQL = pSQL & "{SET ROWCOUNT " & MaxRecords & " SELECT [@PARENTFIELDS]"
        Else
            pSQL = pSQL & "{SELECT TOP " & MaxRecords & " [@PARENTFIELDS]"
        End If
    Else
        pSQL = pSQL & "{SELECT " & "[@PARENTFIELDS]"
    End If
    pSQL = pSQL & " FROM [@PARENT]"
    pSQL = pSQL & " [@WHEREPARENT]"
    pSQL = pSQL & " [@ORDPARENT]} "
    'Substitute for actual values
    pSQL = vba.replace(pSQL, "[@PARENTFIELDS]", ParentFields)
    pSQL = vba.replace(pSQL, "[@PARENT]", tblParent)
    pSQL = vba.replace(pSQL, "[@WHEREPARENT]", pWhere)
    pSQL = vba.replace(pSQL, "[@ORDPARENT]", pOrder)
    pSQL = Trim(pSQL)
    'Define Child SQL Statement
    cSQL = ""
    cSQL = cSQL & "{SELECT " & "[@CHILDFIELDS]"
    cSQL = cSQL & " FROM [@CHILD]"
    cSQL = cSQL & " [@WHERECHILD]"
    cSQL = cSQL & " [@ORDCHILD]} "
    'Substitute for actual values
    cSQL = vba.replace(cSQL, "[@CHILDFIELDS]", ChildFields)
    cSQL = vba.replace(cSQL, "[@CHILD]", tblChild)
    cSQL = vba.replace(cSQL, "[@WHERECHILD]", cWhere)
    cSQL = vba.replace(cSQL, "[@ORDCHILD]", cOrder)
    cSQL = Trim(cSQL)

    'Define Parent Properties
    lSQL = "SHAPE " & pSQL & vbCrLf
    'Define Child Properties
    lSQL = lSQL & "APPEND (" & cSQL & " RELATE " & fldParent & " TO " & fldChild & ") AS ChildItems"

    'Get the data
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = ConnectionString
        .CursorLocation = adUseServer
        .Provider = "MSDataShape"
        .Open
    End With

    Set rs = New ADODB.Recordset
    With rs
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = lSQL
        .ActiveConnection = cn
        .Open
    End With
    Set rs.ActiveConnection = Nothing
    cn.Close
    Set cn = Nothing
    Set Datashape = rs
    Set rs = Nothing
  
ProcExit:
  Exit Function

ErrRtn:
   If Not cn Is Nothing Then
      If cn.State = adStateOpen Then cn.Close
      Set cn = Nothing
   End If
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   
   If cnIn.Errors.count > 0 Then
      ' Enumerate Errors collection and display
      ' properties of each Error object.
      For Each errLoop In cnIn.Errors
         Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, errLoop.Number, _
                     errLoop.Description, errLoop.Source, , , _
                     " SQL " & errLoop.SQLState & " Native " & errLoop.NativeError, _
                     errLoop.HelpFile, errLoop.HelpContext)
      Next
   End If
   
   Resume ProcExit:
End Function

