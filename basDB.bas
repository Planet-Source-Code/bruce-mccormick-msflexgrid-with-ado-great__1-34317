Attribute VB_Name = "basDB"
Option Explicit

' this class needs work
' sample connection srings:
' adConn.Open "Provider=Microsoft.Jet.OLEDB.3.51;" & "Persist Security Info=False;" & "Data Source=c:\??"

Const strModuleName As String * 30 = "basDB"

'Wrap ADO CommandTypeEnum with our own to future proof
Public Enum QueryOptions
   QO_UseTable = ADODB.adCmdTable
   QO_UseText = ADODB.adCmdText
   QO_UseStoredProc = ADODB.adCmdStoredProc
   QO_UseUnknown = ADODB.adCmdUnknown
   QO_UseDefault = -1
End Enum

Private gstrOpenStmt As String
Private MsgBoxAns As Variant
Private DBPath As String
Private db As Connection

Dim cn As New ADODB.Connection
Dim ServerName As String
Dim Provider As String
Dim uname As String
Dim pass As String

Public Sub RSNavigate(rsIn As ADODB.Recordset, MsgIn As String)
   Dim strMessage As String
   Dim intCommand As Integer

   Do Until rsIn.EOF
      ' Display information about current record
      ' and get user input
      strMessage = MsgIn & vbCr & vbCr & _
         "Enter command:    " & vbCr & _
         "1 - Next          " & vbCr & _
         "2 - Previous      " & vbCr & _
         "3 - First         " & vbCr & _
         "4 - Last          " & vbCr & _
         "5 - set Bookmark  " & vbCr & _
         "6 - go to bookmark"
      intCommand = Val(InputBox(strMessage))

      ' Check user input
      Select Case intCommand
         Case 1
            rsIn.MoveNext
            If rsIn.EOF Then
               MsgBox "You Are At End Of The File." & _
                  vbCr & "Try again."
               rsIn.MoveLast
            End If
         Case 2
            rsIn.MovePrevious
            If rsIn.BOF Then
               MsgBox "You Are At The Beginning Of The File." & _
                  vbCr & "Try again."
               rsIn.MoveFirst
            End If
         Case 3
            rsIn.MoveFirst
         Case 4
            rsIn.MoveLast
         Case 5
            RSBookmark = rsIn.BookMark
         Case 6
            If IsEmpty(RSBookmark) Then
               MsgBox "No Bookmark set!"
            Else
               rsIn.BookMark = RSBookmark
            End If
         Case Else
            Exit Do
      End Select
   Loop
End Sub
Public Sub LogonServer(Provider As String)

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "LogonServer"
  
  If Provider = "SQL Server" Then
    cn.ConnectionString = ""
    cn.Provider = "MSDASQL;Driver={SQL Server};SERVER=" & ServerName & ";user id=" & uname & ";Password=" & pass & ";Database=" & db & ""
    cn.Open
    If Err.Number <> 0 Then
        MsgBox Err.Description
         Exit Sub
    End If
    GetDatabase (Provider)
  End If

  If Provider = "Ms Access 2000" Or Provider = "Ms Access 97" Then

    If Provider = "Ms Access 2000" Then
        cn.Provider = "Microsoft.Jet.Oledb.4.0.Provider"
    End If
    If Provider = "Ms Access 97" Then
        cn.Provider = "Microsoft.Jet.Oledb.3.51.Provider"
    End If
        cn.ConnectionString = db
        cn.Open
        If Err.Number <> 0 Then
            MsgBox Err.Description
             Exit Sub
        End If
        GetDatabase (Provider)
  End If

  If Provider = "Oracle" Then
    cn.ConnectionString = ""
    cn.Provider = "MSDAORA.1;Data Source=" & ServerName & ";user id=" & uname & ";Password=" & pass & ";Database=" & db & ""
    cn.Open
    If Err.Number <> 0 Then
        MsgBox Err.Description
         Exit Sub
   End If
    GetDatabase (Provider)
  End If

  If Provider = "Foxpro" Then
    cn.ConnectionString = ""
    cn.Provider = "MSDASQL.1;Data Source=dBASE Files;Initial Catalog=" & db
    cn.Open
        If Err.Number <> 0 Then
        MsgBox Err.Description
         Exit Sub
   End If
   GetDatabase (Provider)
  End If
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Sub GetDatabase(Provider As String)

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "GetDatabase"
  
  If Provider = "SQL Server" Then
    Set rs = cn.Execute("sp_databases")
    Unload Form2
    Do While Not rs.EOF
        Form3.List1.AddItem rs.Fields(0)
        rs.MoveNext
    Loop
    Form3.Show 1
'    rs.Close
  End If

  If Provider = "Oracle" Then
    Set rs = cn.Execute("select * from cat")
    Form1.List1.Clear
    Do While Not rs.EOF
        Form1.List1.AddItem rs("table_name")
        rs.MoveNext
    Loop
 'rs.close
    Form1.Show
    Form1.List1.Enabled = True
    Form1.Command1.Enabled = True
    Form1.Command2.Enabled = True
    Form1.Command3.Enabled = True
    Form1.Command4.Enabled = True
    Form1.Text1.Enabled = True
  End If

  If Provider = "Ms Access 2000" Or Provider = "Ms Access 97" Then
    Set rs = cn.OpenSchema(adSchemaTables)
    Form1.List1.Clear
    Do While Not rs.EOF
        If vba.Left$(rs("table_name"), 4) <> "MSys" Then
            Form1.List1.AddItem rs("table_name")
        End If
        rs.MoveNext
    Loop
    Form1.Show
'    rs.Close
    Form1.List1.Enabled = True
    Form1.Command1.Enabled = True
    Form1.Command2.Enabled = True
    Form1.Command3.Enabled = False
    Form1.Command4.Enabled = True
    Form1.Text1.Enabled = True
  End If

  If Provider = "Foxpro" Then
    Set rs = cn.OpenSchema(adSchemaTables)
    Form1.List1.Clear
    Do While Not rs.EOF
        Form1.List1.AddItem rs("table_name")
        rs.MoveNext
    Loop
    Form1.Show
    Form1.List1.Enabled = True
    Form1.Command1.Enabled = True
    Form1.Command2.Enabled = False
    Form1.Command3.Enabled = True
    Form1.Command4.Enabled = True
    Form1.Text1.Enabled = True
  End If
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Sub GetTables()

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "GetTables"

  cn.Close
  LogonServer (Provider)
  Form1.List1.Clear
  Dim rs As New Recordset
  rs.Open "Select * from sysobjects where xtype='U'", cn, adOpenForwardOnly, adLockOptimistic
  Do While Not rs.EOF
    Form1.List1.AddItem rs!Name
    rs.MoveNext
  Loop
  rs.Close
  Form1.List1.Enabled = True
  Form1.Command1.Enabled = True
  Form1.Command2.Enabled = True
  Form1.Command3.Enabled = True
  Form1.Command4.Enabled = True
  Form1.Text1.Enabled = True
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Sub Command3_Click()

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Command3_Click"

  If Provider = "SQL Server" Or Provider = "Oracle" Then
    GetTables
  End If
  If Provider = "Ms Access 2000" Or Provider = "Ms Access 97" Or Provider = "Foxpro" Then
    Form2.Show 1
  End If
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Sub RunQuery()

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "RunQuery"

  If Provider = "Oracle" Then
    Text1.Text = "Select * from " & List1.Text
  Else
    Text1.Text = "Select * from " & "[" & List1.Text & "]"
  End If

  If Text1.Text = "" Then
    MsgBox "Please Enter Query", vbCritical, "Help!"
     Exit Sub
  End If

  Dim rs As New Recordset
  rs.Open Text1.Text, cn, adOpenKeyset, adLockOptimistic, 1

  If Err.Number <> 0 Then
    MsgBox Err.Description
     Exit Sub
  End If

  Set grid.DataSource = rs
  Label5.Caption = rs.RecordCount
  Label6.Caption = rs.Fields.count

  rs.Close
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)

  If Msg = "True" Then
    If UCase(vba.Left$(Text1.Text, 6)) <> "SELECT" Then
        If Provider = "Oracle" Then
            rs.Open "select * from " & List1.Text, cn, adOpenForwardOnly, adLockOptimistic, 1
        Else
            rs.Open "select * from " & "[" & List1.Text & "]", cn, adOpenForwardOnly, adLockOptimistic, 1
        End If
        Set grid.DataSource = rs
        Msg = "False"
    End If
  End If
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Sub InputTablesToComboBox(ComboBoxIn As ComboBox)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "InputTablesToComboBox"
  Dim TablesCount As Long
  Dim TableName As String
  Dim i As Integer
    ComboBoxIn.Clear
    'broqt na kolonite
    TablesCount = dbDataBase.TableDefs.count
    'pupvite 6 ne sa za pokazvane (nekvi sturotii na access)
    For i = 0 To TablesCount - 1
        Set dbTableDef = dbDataBase.TableDefs(i)
        TableName = dbTableDef.Name
        'tova sa tablici na Access koito ne trqbva da se pipat
        If TableName <> "MSysAccessObjects" And TableName <> "MSysACEs" And TableName <> "MSysObjects" And TableName <> "MSysQueries" And TableName <> "MSysRelationships" Then
           ComboBoxIn.AddItem TableName
        End If
    Next i
     Exit Sub
        
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Sub InputTablesToListBox(List1 As ListBox)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "InputTablesToListBox"
Dim TablesCount As Long
Dim TableName As String
Dim i As Integer
    List1.Clear
    TablesCount = dbDataBase.TableDefs.count
    For i = 0 To TablesCount - 1
        Set dbTableDef = dbDataBase.TableDefs(i)
        TableName = dbTableDef.Name
        If Not InStr("MSys", TableName) Then
           List1.AddItem TableName
        End If
    Next i
    List1.Tag = "Tables"
     Exit Sub

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Sub InputQueriesToListBox(List1 As ListBox)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "InputQueriesToListBox"

Dim QueriesCount As Long
Dim i As Integer
    List1.Clear
    QueriesCount = dbDataBase.QueryDefs.count - 1
    For i = 0 To QueriesCount
        Set dbQueryDef = dbDataBase.QueryDefs(i)
        List1.AddItem dbQueryDef.Name
    Next i
    List1.Tag = "Queries"
     Exit Sub

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit: End Sub

Public Function ConvertToCSV(rs As ADODB.Recordset) As String
   ' usage for this and following proc:
   
   'Convert the recordset to a string
   'If RS.EOF = False And RS.BOF = False Then
   '   'Dermine formatting
   '   If StringType = "HTML" Then
   '      GetString = ConvertToHTML(RS)
   '   Else
   '      GetString = ConvertToCSV(RS)
   '   End If
   'Else
   '   GetString = ""
   'End If
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ConvertToCSV"

   'Check for closed or empty recordset
   If rs.EOF = True Or rs.BOF = True Then
      ConvertToCSV = ""
      Exit Function
   End If

   'Convert recordset to comma seperated values
   ConvertToCSV = rs.GetString(adClipString, -1, ",", vbCrLf, "(NULL)")


ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function ConvertToHTML(rs As ADODB.Recordset) As String
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ConvertToHTML"

   'Check for closed or empty recordset
   If rs.EOF = True Or rs.BOF = True Then
      ConvertToHTML = ""
      Exit Function
   End If

   'Convert recordset to HTML table syntax
   ConvertToHTML = rs.GetString(adClipString, -1, "</TD><TD>", "</TD></TR>" & vbCrLf & "<TR><TD>", "(NULL)")
   ConvertToHTML = "<TR><TD>" & vba.Left$(ConvertToHTML, Len(ConvertToHTML) - 8)


ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function ConvertToArray(ByVal sQuery As String, ByVal sConnect As String, eType As QueryOptions, ParamArray aParams() As Variant) As Variant()
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ConvertToArray"

   'Create the ADO objects
   Dim rs As ADODB.Recordset, Cmd As ADODB.Command
   Set rs = New ADODB.Recordset
   Set Cmd = New ADODB.Command

   'Use helper function to build parameters for command object
   CollectParams Cmd, aParams

   'Determine whether passed in, or hard coded connection
   If sConnect = vbNullString Then
      Cmd.ActiveConnection = GetConnectionString()
   Else
      Cmd.ActiveConnection = sConnect
   End If

   'Init the ADO objects & the query parameters (if any)
   Cmd.CommandText = sQuery
   Cmd.CommandType = eType

   'Execute the query for readonly
   rs.CursorLocation = adUseClient
   rs.CursorType = adOpenForwardOnly
   rs.LockType = adLockBatchOptimistic
   rs.Open Cmd

   'Convert the recordset to an array
   If rs.EOF = False And rs.BOF = False Then
      GetArray = rs.GetRows
   Else
      ReDim GetArray(-1 To -1, -1 To -1)
   End If

   'Clean up and exit
   Set rs = Nothing
   Set Cmd = Nothing
   Exit Function

ProcExit:
  Exit Function

ErrRtn:
   Set rs = Nothing
   Set Cmd = Nothing
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function FindTypeConstant(strType As String) As Byte
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindTypeConstant"
    Select Case strType
           Case "Boolean": FindTypeConstant = 1
           Case "Byte": FindTypeConstant = 2
           Case "Integer": FindTypeConstant = 3
           Case "Long": FindTypeConstant = 4
           Case "Currency": FindTypeConstant = 5
           Case "Single": FindTypeConstant = 6
           Case "Double": FindTypeConstant = 7
           Case "Date/Time": FindTypeConstant = 8
           Case "Text": FindTypeConstant = 10
           Case "Binary": FindTypeConstant = 9
           Case "Memo": FindTypeConstant = 12
     End Select

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit: End Function

Public Function NextSeqNbr(TableIn As String, KeyFieldIn As String) As Long
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "NextSeqNbr"
    Dim rsNbrGen As Recordset
    
    gstrOpenStmt = "select max('" & KeyFieldIn & "') from '" & TableIn & "'"
    Set rsNbrGen = cn.Execute(gstrOpenStmt)
    
    If IsNull(rs_NoGen(0)) = False Then
        NextSeqNbr = rs_NoGen(0) + 1
    Else
        NextSeqNbr = 1
    End If
        
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function DBFind(Provider As String) As String

  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "DBFind"

  Dim a As String
  Dim b As String

  If Provider = "Ms Access 2000" Or Provider = "Ms Access 97" Then
    Cmd.Filter = "*.mdb"
    Cmd.FileName = "*.mdb"
  End If
  If Provider = "Foxpro" Then
    Cmd.Filter = "*.dbf"
    Cmd.FileName = "*.dbf"
  End If

  Cmd.DialogTitle = "Select Database"
  Cmd.ShowOpen

  If Provider = "Foxpro" Then
    a = Cmd.FileTitle
    b = Cmd.FileName
    DBFind = vba.Mid$(b, 1, Len(b) - Len(a) - 1)
  Else
    DBFind = Cmd.FileName
  End If
  
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function FindDB(DBPathIn As String, Optional DBNameIn As String) As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindDB"

  With frmCommonDialog.CommonDialog1
    .DialogTitle = "Please Select The Database"
    .Filter = "All Files (*.*)|*.*|Access DB(*.mdb)|*.mdb"
    .InitDir = DBPathIn
    If Not IsMissing(DBNameIn) Then .FileName = DBNameIn
    .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist
    .ShowOpen
      
    If .FileName <> "" Then
      If .FileName = DBNameIn Then
        FindDB = True
      Else
        MsgBoxAns = MsgBox("Database Not Found. Returning To Menu.", vbCritical & vbOKOnly, Not db)
        FindDB = False
        Unload Me
      End If
    Else
      MsgBoxAns = MsgBox("Database Not Found. Returning To Menu.", vbCritical & vbOKOnly, Not db)
      FindDB = False
      Unload Me
    End If
  End With
    
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
  FindDB DBPathIn, DBNameIn
   Resume ProcExit:
End Function

Function NavButtons(Frs As CommandButton, _
                    Prvs As CommandButton, _
                    Nxt As CommandButton, _
                    Lst As CommandButton, _
                    db As Data)
  ' ManageButtons BtnFrs, BtnPrev, BtnNext, BtnLast, DBName
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ManageButtons"
  
  Dim Pos As Integer
  Dim RSRecCnt As Integer
           
  'Get position and number of records
  Pos = Se(db.Recordset.AbsolutePosition + 1, 1)
  RSRecCnt = Se(GetRecCount("Customer"), 1)
    
  'If there are many records
  If RSRecCnt > 1 Or Pos > 1 Then
    
    'If the current record is the last of the recordset
    If Pos >= RSRecCnt Then
      Frs.Enabled = True
      Prvs.Enabled = True
      Nxt.Enabled = False
      Lst.Enabled = False
    End If
    
    'If the current record is the firs of the recordset
    If Pos = 1 Then
      Nxt.Enabled = True
      Lst.Enabled = True
      Frs.Enabled = False
      Prvs.Enabled = False
    End If
        
    'If the current record is not the firs and not the last
    If Pos > 1 And Pos < RSRecCnt Then
      Frs.Enabled = True
      Prvs.Enabled = True
      Nxt.Enabled = True
      Lst.Enabled = True
    End If
  Else
    'If there no records
    Frs.Enabled = False
    Prvs.Enabled = False
    Nxt.Enabled = False
    Lst.Enabled = False
  End If
    
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Function DAOGetRecCount(TableIn As String, DBNameIn As String) As Integer
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "GetRecCnt"
    
  Dim dbs As Database
  Dim rs As Recordset
   
  Set dbs = OpenDatabase(App.Path + "\" + DBNameIn)
  Set rs = dbs.OpenRecordset("SELECT Count(*) AS Nbr FROM " + TableIn)
  
  If Not rs.EOF Then
    GetRecCount = rs!nbr
    rs.Close
  Else
    GetRecCount = 0
  End If
  
  dbs.Close
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Sub Form_KeyDown(KeycodeIn As Integer)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Form_KeyDown"
  
  Select Case KeycodeIn
    Case vbKeyPageUp, vbArrowUp, vbArrowLeft
      KeycodeIn = 0
      If BtnPrevious.Enabled = True Then Call BtnPrevious_Click
    Case vbKeyPageDown, vbKeyArrowDown, vbArrowRight
      KeycodeIn = 0
      If BtnNext.Enabled = True Then Call BtnNext_Click
    Case vbKeyF1
      KeycodeIn = 0
      Call BtnHelp_Click
    Case vbKeyF2
      KeycodeIn = 0
      Call BtnNew_Click
    Case vbKeyF5
      KeycodeIn = 0
      Call BtnDelete_Click
  End Select

ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Function DAODoesRecExist(TableIn As String, FieldIn As String, ValueIn As String) As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "DoesRecExist"
  
   ' call example"
     ' public Sub TBCodeTemp_Validate(Cancel As Boolean)
     '  Dim Code As String
     '  Code = TBCodeTemp
     '  If DoesRecExist("Customer", "Code", TBCodeTemp) And Code <> OldCode Then
     '      If All_Empty Then ' see fmtand validate for all_empty
     '          MainDB.Recordset.CancelUpdate
     '      End If
     '      Find_Item "Code", Code
     '  Else
     '      If Code <> OldCode Then
     '         TBCode = TBCodeTemp
     '         MainDB.Refresh
     '          Find_Item "Code", Code
     '      End If
     '  End If
     '  OldCode = TBCode
     ' End Sub


  Dim dbs As Database
  Dim rs As Recordset
    
  Set dbs = OpenDatabase(App.Path + gstrMainDB)
  Set rs = dbs.OpenRecordset("SELECT * FROM " + TableIn + " WHERE " + FieldIn + "=""" + ValueIn + """")
  If Not rs.EOF Then
    DoesRecExist = True
  Else
    DoesRecExist = False
  End If
  
  rs.Close
  rs = Nothing
  dbs.Close

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function FindRec() As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindRec"
    Dim strTemp
    
    strTemp = "'" & strSearch & "'"
    On Error Resume Next
    rsContacts.MoveFirst
    
    rsContacts.Find "ContactName = " & strTemp, 0, adSearchForward
    
    If rsContacts!ContactName = strSearch Then FindRec = True

ProcExit:
  Exit Function

ErrRtn:
  FindRec = False      'not found
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Sub FindRec2()
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindRec2"

  On Error Resume Next
  Dim x As Double
  x = InputBox("Type in the Employee ID to Edit", "Edit Employee ID", Val(txtID) - 1)

  Dim rs As New Recordset
  rs.Open "Select * from Master where ID=" & Val(x), CON, adOpenKeyset, adLockOptimistic
  If rs.RecordCount <> 0 Then
    txtID.Text = rs![ID]
    txtName.Text = rs![Name]
    txtAge.Text = rs![Age]
    txtSalary.Text = rs![Salary]
    rs.Close
  Else
    MsgBox "Invalid Employee ID", vbCritical, "Invalid ID"
  End If

ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Function FindRec3(Field As String, Code As String, Optional LogicalOperator As String = "=")
  'Find a record
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindRec3"
    
  If Not MainDB.Recordset.EOF And Not MainDB.Recordset.BOF Then
     MainDB.Recordset.FindFirs Field + LogicalOperator + " """ + Code + """"
     ManageButtons BtnFirs, BtnPrevious, BtnNext, BtnLast, MainDB
  End If

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function
        
Public Function rs_find(KeyIn As String, KeyInType As String, DBFieldIn As String)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "rs_find"
        Dim intKey As Integer
        Dim strKey As String
        Dim datKey As Date
        
        rs.MoveFirst
        
        If KeyInType = "str" Then
          rs.Find (DBFieldIn & " Like '" & KeyIn & "'")
        ElseIf KeyInType = "int" Then
          intKey = CInt(KeyIn)
          rs.Find (DBFieldIn & " Like '" & intKey & "'")
        ElseIf keytype = "dat" Then
          datKey = CDate(KeyIn)
          rs.Find (DBFieldIn & " Like '" & datKey & "'")
        End If
        
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function FindFirs(rsIn As Recordset, _
                          LogicalOperatorIn As String, _
                          TypeIn As String) As Boolean
  
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindFirs"
  
  Dim varBookMark As Variant
  
  If rsIn.EOF And rsIn.BOF Then
    FindFirs = False
    GoTo ProcExit:
  End If
  
  FindInit LogicalOperatorIn
  
  If rsIn.EOF Then
    FindFirs = False
    GoTo ProcExit:
  End If
  
  If rsIn.BOF Then varBookMark = rsIn.BookMark
    
  If TypeIn = "Numeric" Then
    rsIn.Find PKFieldName & "=" & rsempAS(0).Value
  Else   ' date and text use ' delimiter
    rsIn.Find PKFieldName & "='" & rsempAS(0).Value & "'"
  End If
    
  FindFirs = Not rsIn.EOF And Not rsIn.BOF
    
  If Not FindFirs And Not IsNull(BookMark) Then rsIn.BookMark = varBookMark

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function FindLast(rsIn As Recordset, LogicalOperatorIn As String) As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindLast"
  
  Dim BookMark As Variant
  
  If rsIn.EOF And rsIn.BOF Then Exit Function
  
  FindInit LogicalOperatorIn
  
  If rsempAS.EOF Then
    FindLast = False
  Else
    rsempAS.MoveLast
    
    If Not rsIn.EOF And Not rsIn.BOF Then BookMark = rsIn.BookMark
    
    If mType = asNumeric Then
      rsIn.Find PKFieldName & "=" & rsempAS(0).Value
    Else   ' date and text use ' delimiter
      rsIn.Find PKFieldName & "='" & rsempAS(0).Value & "'"
    End If
    
    FindFirs = Not rsIn.EOF And Not rsIn.BOF
    
    If Not FindFirs And Not IsNull(BookMark) Then rsIn.BookMark = BookMark
  End If

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function FindNext() As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindNext"
  
  Dim BookMark As Variant
  If rsIn.EOF And rsIn.BOF Then Exit Function
  If rsempAS Is Nothing Then Exit Function
  If rsempAS.State = 0 Then Exit Function
  If rsempAS.EOF Then Exit Function
  
  rsempAS.MoveNext
  
  If rsempAS.EOF Then Exit Function
  If Not rsIn.EOF And Not rsIn.BOF Then BookMark = rsIn.BookMark
  
  If mType = asNumeric Then
    rsIn.Find PKFieldName & "=" & rsempAS(0).Value
  Else   ' date and text use ' delimiter
    rsIn.Find PKFieldName & "='" & rsempAS(0).Value & "'"
  End If
  
  FindNext = Not rsIn.EOF And Not rsIn.BOF
  
  If Not FindNext And Not IsNull(BookMark) Then rsIn.BookMark = BookMark
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function FindPrev() As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FindPrev"
  
  Dim BookMark As Variant
  
  If rsIn.EOF And rsIn.BOF Then Exit Function
  If rsempAS Is Nothing Then Exit Function
  If rsempAS.State = 0 Then Exit Function
  If rsempAS.BOF Then Exit Function
  
  rsempAS.MovePrevious
  
  If rsempAS.BOF Then Exit Function
  If Not rsIn.EOF And Not rsIn.BOF Then BookMark = rsIn.BookMark
  
  If mType = asNumeric Then
    rsIn.Find PKFieldName & "=" & rsempAS(0).Value
  Else   ' date and text use ' delimiter
    rsIn.Find PKFieldName & "='" & rsempAS(0).Value & "'"
  End If
  
  FindPrevious = Not rsIn.EOF And Not rsIn.BOF
  
  If Not FindPrevious And Not IsNull(BookMark) Then rsIn.BookMark = BookMark

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function DBSearch(stRSrchFor As String, _
                         rs As Recordset, _
                         fldDBField As Field) As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "DBSearch"
  
  Dim foundFlag As Boolean
  foundFlag = False
    
  With rs
    If .RecordCount > 0 Then
      .MoveFirst
      For i = 1 To .RecordCount
        If fldDBField = stRSrchFor Then
          foundFlag = True
          i = .RecordCount
        End If
        
        If foundFlag = False Then
          .MoveNext
        End If
        
      Next i
        
      If foundFlag = True Then
         MsgBox ("Record has been located!")
      Else
         MsgBox ("No Match in Database!")
         .MoveFirst
      End If
    Else
      MsgBox ("There are no records to search!")
    End If
  End With
  Search = foundFlag

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function ConnectedToDB(db As Database, Path As String, _
                              Optional ReadOnly As Boolean, _
                              Optional ConnectString As String, _
                              Optional ByVal ConnectRegardless As Boolean) As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ConnectToDB"
  Dim strName As String
  Dim secAttempt As Boolean

  If Not ConnectRegardless Then
    If Valid(db) Then ConnectRegardless = db.Name = "" Else ConnectRegardless = True                              'If we get here we need to, you guessed it, connect regardless...
  End If
  
  If ConnectRegardless Then
    Set db = Workspaces(0).OpenDatabase(Path, False, False, ConnectString)
    ConnectedToDB = Valid(db)
  Else
    ConnectedToDB = True
  End If

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Sub List1_Click()

    On Error GoTo ErrRtn
    gstrChkPt = "On Error": gstrProcName = "List1_Click"

    StripOutApostrophes

    Select Case List1.ListIndex
    Case 0
      mstQuery = "SELECT * FROM ACR WHERE Name LIKE '%" & Text1(1).Text & "%'"
    Case 1
      mstQuery = "SELECT * FROM ACR WHERE Name LIKE '%" & Text1(1).Text & "%' AND Date LIKE '%" & DTPicker1.Value & "%'"
    Case 2
      mstQuery = "SELECT * FROM ACR WHERE Name LIKE '%" & Text1(1).Text & "%' AND Date LIKE '%" & DTPicker1.Value & "%' AND Call LIKE '%" & Text1(0).Text & "%'"
    End Select
        
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub
' Easy steps to use this ADO library:
' -----------------------------------
' 1. Add modGlobals.bas, modDatabaseUtilities.bas, modErrorHandling.bas
'    to your VB project;
' 2. Initialize the following global variables defined in modGlobals.bas as follows:
'    - hDBConn with your ADO connection
'    - vbDateFormat with the date format you use (something like "mm/dd/yyyy"...)

Public Function LoadCBO(cbo As Object, stRSQL As String, _
                        Optional PreserveSelectedValue As Boolean = False, _
                        Optional AllowNull As Boolean = True) As Boolean
  Dim OldValue As String
  
  On Error GoTo LoadComboDataErr
  
  LoadComboData = False
  
  If PreserveSelectedValue = True Then
    OldValue = cbo.Text
  End If
  
  cbo.Clear
  
  Dim recResult As ADODB.Recordset
  Set recResult = New ADODB.Recordset
  recResult.Open stRSQL, hDBConn
  
  If cbo.style = 2 And AllowNull = True Then
    cbo.AddItem ""
  End If
  
  If recResult.EOF = True Then
    If Not recResult Is Nothing Then
      If CBool(recResult.State And ADODB.adStateOpen) Then
        recResult.Close
      End If
      Set recResult = Nothing
    End If
    If PreserveSelectedValue = True Then
      cbo.Text = OldValue
    End If
    Exit Function
  End If
  
  While Not recResult.EOF
    cbo.AddItem CStr(recResult.Fields(0).Value)
    If Not recResult.EOF Then
      recResult.MoveNext
    End If
  Wend
  
  If PreserveSelectedValue = True Then
    cbo.Text = OldValue
  End If
  
  If Not recResult Is Nothing Then
    If CBool(recResult.State And ADODB.adStateOpen) Then
      recResult.Close
    End If
    Set recResult = Nothing
  End If
  
  LoadComboData = True

  Exit Function

LoadComboDataErrRtn:
  On Error Resume Next
  
  If Not recResult Is Nothing Then
    If CBool(recResult.State And ADODB.adStateOpen) Then
      recResult.Close
    End If
    Set recResult = Nothing
  End If
End Function

Public Function LoadListData(ListBox As Object, stRSQL As String) As Boolean
  On Error GoTo LoadListDataErr
  
  LoadListData = False
  ListBox.Clear
  
  Dim recResult As ADODB.Recordset
  Set recResult = New ADODB.Recordset
  recResult.Open stRSQL, hDBConn, adOpenStatic, adLockOptimistic, adCmdText
  
  If recResult.EOF = True Then
    recResult.Close
    Set recResult = Nothing
    Exit Function
  End If
  
  recResult.MoveFirst
  While Not recResult.EOF
    ListBox.AddItem CStr(recResult.Fields(0).Value)
    If Not recResult.EOF Then
      recResult.MoveNext
    End If
  Wend
  
  recResult.Close
  Set recResult = Nothing
  
  LoadListData = True

  Exit Function

LoadListDataErrRtn:
  On Error Resume Next
  
  If Not recResult Is Nothing Then
    If CBool(recResult.State And ADODB.adStateOpen) Then
      recResult.Close
    End If
    Set recResult = Nothing
  End If
End Function

Public Function ExecuteSQL(stRSQL As String, lRowsAffected As Long) As Boolean
  
  Dim recAffected As Long
  
  On Error GoTo ExecuteSQLErr
  
  ExecuteSQL = False
  
  BeginTrans
  hDBConn.Execute stRSQL, recAffected
  CommitTrans
  
  lRowsAffected = recAffected
  ExecuteSQL = True

  Exit Function

ExecuteSQLErrRtn:

  RollBackTrans
  
  GetAPPError
  GetADOErrors

End Function

Public Function LoadCtl(ctl As Object, stRSQL As String, strIDColumn As String, ParamArray strColumns() As Variant) As Boolean
  On Error GoTo LoadCtlErr
  
  LoadCtl = False
  
  ctl.Clear
  
  Dim recResult As ADODB.Recordset
  Set recResult = New ADODB.Recordset
  recResult.Open stRSQL, hDBConn, adOpenStatic, adLockOptimistic, adCmdText
  
  If recResult.EOF = True Then
    recResult.Close
    Set recResult = Nothing
    Exit Function
  End If
  
  Dim strData As String
  Dim i As Integer
  
  recResult.MoveFirst
  While Not recResult.EOF
    strData = ""
    For i = 0 To UBound(strColumns)
      strData = strData & " - " & CStr(recResult(CStr(strColumns(i))))
    Next i
    ctl.AddItem strData
    ctl.ItemData(ctl.NewIndex) = recResult(strIDColumn)
    If Not recResult.EOF Then
      recResult.MoveNext
    End If
  Wend
  
  recResult.Close
  Set recResult = Nothing
  
  LoadCtl = True

  Exit Function

LoadCtlErrRtn:
  GetAPPError
  GetADOErrors

  On Error Resume Next
  
  If Not recResult Is Nothing Then
    If CBool(recResult.State And ADODB.adStateOpen) Then
      recResult.Close
    End If
    Set recResult = Nothing
  End If
End Function

Public Function ExecuteValueReturn(strSELECT As String) As String
  On Error GoTo ExecuteValueReturnErr
  
  ExecuteValueReturn = ""
  
  Dim recResult As ADODB.Recordset
  Set recResult = New ADODB.Recordset
  
  With recResult
    .ActiveConnection = hDBConn
    .LockType = adLockBatchOptimistic
    .CursorLocation = adUseClient
    .Open strSELECT
  End With
  
  If Not recResult.EOF Then
    ExecuteValueReturn = CStr(CVNull(recResult.Fields(0).Value, vbString, ""))
  Else
    ExecuteValueReturn = ""
  End If
  
  If Not recResult Is Nothing Then
    If CBool(recResult.State And ADODB.adStateOpen) Then
      recResult.Close
    End If
    Set recResult = Nothing
  End If
  
  Exit Function

ExecuteValueReturnErrRtn:

  GetAPPError "SELECT: " + strSELECT
  GetADOErrors "SELECT: " + strSELECT
  
  If Not recResult Is Nothing Then
    If CBool(recResult.State And ADODB.adStateOpen) Then
      recResult.Close
    End If
    Set recResult = Nothing
  End If
  
End Function

Public Function PrepareString(str As String) As String
  Dim strTemp As String
  Dim i As Integer
  
  strTemp = ""

  For i = 1 To Len(str)
    If vba.Mid$(str, i, 1) = "'" Then
      strTemp = strTemp + "'"
    End If

    strTemp = strTemp + vba.Mid$(str, i, 1)
  Next i
  
  PrepareString = strTemp
End Function

Public Sub RollBackTrans()
  On Error GoTo ErrRtnorHandler
  
  If bInsideTransaction = True Then
    hDBConn.RollBackTrans
    bInsideTransaction = False
  End If
  
   Exit Sub
  
ErrorHandler:
  If Err.Number = -2147168242 Then
  End If
End Sub

Public Sub BeginTrans()
  On Error GoTo ErrRtnorHandler
  
  If bInsideTransaction = True Then
    hDBConn.RollBackTrans
  End If
  
  bInsideTransaction = True
  hDBConn.BeginTrans
  
   Exit Sub
  
ErrorHandler:
  If Err.Number = -2147168242 Then
  End If
End Sub

Public Sub CommitTrans()
  On Error GoTo ErrRtnorHandler
  
  If bInsideTransaction = True Then
    hDBConn.CommitTrans
    bInsideTransaction = False
  End If
  
   Exit Sub
  
ErrorHandler:
  If Err.Number = -2147168242 Then
  End If
End Sub

Public Sub CreateIndex()
   gstrChkPt = "On Error": gstrProcName = "CreateIndex"
   On Error GoTo ErrRtn

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

ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Sub IndexOutput(rsIn As Recordset, strIndex As String)
   ' Report function for FieldX.
   gstrChkPt = "On Error": gstrProcName = "IndexOutput"
   On Error GoTo ErrRtn

   With rsIn
      ' Set the index.
      .Index = strIndex
      .MoveFirst
   End With

ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub


