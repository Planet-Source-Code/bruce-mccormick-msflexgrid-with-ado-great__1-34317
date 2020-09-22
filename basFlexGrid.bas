Attribute VB_Name = "basFlexGrid"
Option Explicit

' Dim WithEvents adoPrimaryrs As Recordset  '  for DataGrid Sort

Const strModuleName As String * 40 = "basFlexGrid"

Public Function GridColors(GridIn As MSFlexGrid, RedIn As Integer, GreenIn As Integer, BlueIn As Integer)
  ' to activat the SUB:
  '(general: GridColors MSFlexGrid, Red, Green, Blue)
  ' GridColors Form1.MSFlexGrid, 192, 255, 192
  
   On Error GoTo ErrRtn
   gstrChkPt = "On Error": gstrProcName = "GridColors"
  
   Dim i As Integer
   Dim j As Integer
  
   ' I used FixedCols here because I didn't want to color and fixed cols - only the "interior" cols
   ' You could use for j = 0 and for i = 1
   For j = GridIn.FixedCols To GridIn.Cols - 1
      For i = GridIn.FixedRows To GridIn.Rows - 1
         
         GridIn.Row = i
         GridIn.Col = j
         
         If i Mod 2 <> 0 Then
            GridIn.CellBackColor = RGB(RedIn, GreenIn, BlueIn)
         Else
            GridIn.CellBackColor = GridIn.BackColor
         End If
      Next i
   Next j

ProcExit:
   Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Sub GridClear(GridIn As MSFlexGrid)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ClearGrid"
  
  Dim a As Integer
  Dim b As Integer
  With GridIn
    .Redraw = False
    For a = 1 To .Rows - 1
      For b = 1 To .Cols - 1
        .Row = a: .Col = b: .Text = " "
      Next b
    Next a
    .Redraw = True
  End With
  
ProcExit:
   Exit Sub      ''

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Public Function DragDropCol(GridIn As MSFlexGrid)
     'usage:
     'Public sub <grid>_DragDrop <grid-name>
     '   basGridFunctions.GridDragDrop <grid>
     'End Sub
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "DragDropCol"

  If GridIn.Tag = "" Then Exit Function
  GridIn.Redraw = False
  GridIn.ColPosition(Val(GridIn.Tag)) = GridIn.MouseCol
    
  GridIn.Col = 0
  GridIn.ColSel = GridIn.Cols - 1
  GridIn.Sort = 1 ' Generic Ascending
    
  GridIn.Redraw = True

ProcExit:
  Exit Function ''

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function MouseDown(GridIn As MSFlexGrid)

  'usage:
  'Public Sub <grid>_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '  basGridFunctions.GridMouseDown <grid>
  'End Sub
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "MouseDown"
  
  GridIn.Tag = ""
  If GridIn.MouseRow <> 0 Then Exit Function
  GridIn.Tag = str(GridIn.MouseCol)
  GridIn.Drag 1

ProcExit:
  Exit Function ''

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function AddItem(GridIn As MSFlexGrid)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "AddCItem"

  GridIn.AddItem ""

ProcExit:
  Exit Function ''

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function RemoveRow(GridIn As MSFlexGrid, RowIn As Integer)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "RemoveRow"

  Dim Answer As Integer
  Dim ColCounter As Integer
  
  Answer = MsgBox("Are You Sure You Wwant To Remove This Row?", _
                  vbYesNo + vbDefaultButton2, "Confirm remove...")
  Select Case Answer
    Case vbYes
      If GridIn.Rows = GridIn.FixedRows + 1 Then
        For ColCounter = 1 To GridIn.Cols - 1
          GridIn.Col = ColCounter
          GridIn.Text = ""
        Next
      Else
        GridIn.RemoveItem RowIn
      End If
    Case vbNo

  End Select

ProcExit:
  Exit Function ''

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function GridClick(GridIn As MSFlexGrid, TxtBoxIn As TextBox)
  ' position textbox inside flexgrid cells. column 0 is used as a marker
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "GridClick"

  TxtBoxIn.Visible = True
  TxtBoxIn.Height = GridIn.CellHeight - 10 'minus 10 so that grid lines
  TxtBoxIn.Width = GridIn.CellWidth - 10 '  will not be overwritten
  TxtBoxIn.Left = GridIn.CellLeft + GridIn.Left
  TxtBoxIn.Top = GridIn.CellTop + GridIn.Top
  TxtBoxIn.Text = GridIn.Text
  TxtBoxIn.SetFocus

ProcExit:
  Exit Function ''

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function GridTxtBoxChange(GridIn As MSFlexGrid, TxtBoxIn As TextBox)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "GridTxtBoxChange"

  GridIn.Text = TxtBoxIn.Text

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function GridSetup(GridIn As MSFlexGrid, _
                          TxtBoxIn As TextBox, _
                          NbrRows As Integer, _
                          NbrCols As Integer)
  
 ' need to standardize his module with col headings, rows, cols, etc.
 
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "GridSetup"

  Dim i As Integer
    
  TxtBoxIn.Visible = False
  
  With GridIn
    .Clear
    .Redraw = False
    .Appearance = flexFlat
    .Rows = 2
    .Cols = 3
    .FixedCols = 0
 
    ' To set the heading of the columns - need to make generic - get from DB
    .FormatString = "Employee No " & "|" & "Employee Name                " & "|" & "Salary        " & "|" & "Age " & "|" & "Address               " & "|" & "Designation           "
    ' or
    .TextMatrix(0, 0) = "Item Code"
    .TextMatrix(0, 1) = "Description"
    .TextMatrix(0, 2) = "Rate"
    .MergeCells = flexMergeRestrictColumns
    .AllowUserResizing = flexResizeBoth
    .Row = 0
    .ColAlignment(0) = 7
    For i = 0 To .Cols - 1
      .Col = i
      .CellFontSize = 14
      .CellAlignment = flexAlignLeftCenter

      .MergeCol(i) = True     ' Allow merge on Columns 0 thru 3
      .ColWidth(i) = 2000     ' Set column's width
    Next i

    .MergeCells = flexMergeRestrictColumns
  End With
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function ColSort(GridIn As MSFlexGrid, ByVal ColIndexIn As Integer)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ColSort"
  
  ' Description:Sorts the records in a datagrid form when clicking on the column header.
  '   Toggles between ascending/descending sort order.
  
  Dim strColName As String
  Static blnSortAsc As Boolean
  Static strPrevCol As Integer
  
  ' Did the user click again on the same column ? If so, check
  ' the previous state, in order to toggle between sorting ascending
  ' or descending. If this is the first time the user clicks on a column
  ' or if he/she clicks on another column, then sort ascending.

'  If ColIndexIn = strPrevCol Then
'    If blnSortAsc Then
'      adoPrimaryrs.Sort = strColName & " DESC"
'      blnSortAsc = False
'    Else
'      adoPrimaryrs.Sort = strColName
'      blnSortAsc = True
'    End If
'  Else
'    adoPrimaryrs.Sort = strColName
'    blnSortAsc = True
'  End If
    
'  strPrevCol = strColName

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Sub GridSort(GridIn As MSFlexGrid, SortColIn As Integer, ColType As String)
  ' need to modifythis for coltype
  Dim Ro As Integer
  Dim sortdate As Double

  'add a column to hold the sort key
  GridIn.Cols = GridIn.Cols + 1
  SortCol = GridIn.Cols - 1
  GridIn.ColWidth(SortCol) = 0  'invisible
  'calculate key values & populate grid
  For Ro = 1 To GridIn.Rows - 1
  ' need to be able to mark each col as to whether its a date, str, etc
  ' need to make the sort col generic
  
     sortdate = DateValue(GridIn.TextMatrix(Ro, 2))
     GridIn.TextMatrix(Ro, SortCol) = sortdate
  Next Ro
  'do the sort
  GridIn.Col = SortColIn  'set the key
  GridIn.Sort = flexSortNumericAscending
  
End Sub

Public Function GridSort2(GridIn As MSFlexGrid)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "GridSort"
    
  GridIn.Col = 0
  GridIn.ColSel = GridIn.Cols - 1
  GridIn.Sort = 1 ' Generic Ascending

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function FormResize(GridIn As MSFlexGrid)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "FormResize"
    
  GridIn.Move ScaleLeft, cmdClpbrd.Height, ScaleWidth, ScaleHeight - (cmdClpbrd.Top + cmdClpbrd.Height)
        
ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function CopyToClipBd(GridIn As MSFlexGrid)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "CopyToClipBd"

' If you just want a straightforward copy into the clipboard of the
' selected cells without VBCRLF (using VBCR instead), you can use :
'
' Clipboard.Clear
' Clipboard.SetText gridin.Clip
'
' To Directly copy all selected cells into the clipboard

  Dim i As Long
  Dim j As Long
  Dim maxj As Long
  Dim maxi As Long
  Dim strBuffer As String

  strBuffer = ""
  Clipboard.Clear

  maxj = GridIn.Rows
  maxi = GridIn.Cols
'
' Use Below Code for Standard gridin.OCX Grid
'    For j = firstrow To maxj
'        gridin.Row = j
'        For i = firstcol To maxi
'            gridin.Col = i
'            If i = maxi Then
'                strBuffer = strBuffer & gridin.Text & vbCrLf
'            Else
'                strBuffer = strBuffer & gridin.Text & Chr(9)
'            End If
'        Next
'    Next
'
' Use Below Code for MSFlexGrid
  For j = FirstRow To maxj
    For i = firstcol To maxi
      If i = maxi Then
        strBuffer = strBuffer & GridIn.TextMatrix(j, i) & vbCrLf
      Else
        strBuffer = strBuffer & GridIn.TextMatrix(j, i) & vbTab
      End If
    Next
  Next
    
  Clipboard.SetText strBuffer
        
  MsgBox "Data copied to clipboard !!!"

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function RemoveCol(GridIn As MSFlexGrid, ColNameIn As String)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "RemoveCol"
  
  If GridIn.Cols = 1 Then GoTo ProcExit:
  
  For a = 0 To GridIn.Cols - 1
    If GridIn.TextMatrix(0, a) = ColNameIn Then
        GridIn.TextMatrix(0, a) = ""
        'shift to left
        For b = a To GridIn.Cols - 2
            GridIn.TextMatrix(0, b) = GridIn.TextMatrix(0, b + 1)
        Next b
        GridIn.Cols = GridIn.Cols - 1
        Exit For
    End If
  Next a

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function AddCol(GridIn As MSFlexGrid, HeadingIn As String)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "AddCol"

  Dim there As Boolean
  there = False
  
  If GridIn.Cols = 0 Then GoTo SkipCalc
  
  For a = 0 To GridIn.Cols - 1
    If GridIn.TextMatrix(0, a) = HeadingIn Then
    there = True
    End If
  Next a
  
  If there = True Then GoTo ProcExit:

SkipCalc:
  GridIn.Cols = GridIn.Cols + 1
  GridIn.TextMatrix(0, GridIn.Cols - 1) = HeadingIn

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Function ClearCol(GridIn As MSFlexGrid, ByVal ColIn As Integer)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "ClearCol"
  
  For a = 0 To GridIn.Rows - 1
    GridIn.TextMatrix(a, ColIn) = ""
  Next a

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function YesNo(GridIn As MSFlexGrid)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "YesNo"
  
  For a = GridIn.FixedCols To GridIn.Cols - 1
    For b = GridIn.FixedRows To GridIn.Rows - 1
        If GridIn.TextMatrix(b, a) = "0" Then GridIn.TextMatrix(b, a) = "No"
        If GridIn.TextMatrix(b, a) = "1" Then GridIn.TextMatrix(b, a) = "Yes"
    Next b
  Next a '

ProcExit:
  Exit Function '

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function AutoSizeCells(GridIn As MSFlexGrid, FormIn As Object)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "AutoSizeCells"

  For a = 0 To GridIn.Cols - 1
    GridIn.ColWidth(a) = FormIn.TextWidth(GridIn.TextMatrix(0, a)) * 1.5
  Next a
  For a = 0 To GridIn.Cols - 1
    For b = 0 To GridIn.Rows - 1
        If GridIn.ColWidth(a) < FormIn.TextWidth(GridIn.TextMatrix(b, a)) * 1.5 Then
           GridIn.ColWidth(a) = FormIn.TextWidth(GridIn.TextMatrix(b, a)) * 1.5
           If Len(GridIn.TextMatrix(b, a)) = 1 Then
                GridIn.ColWidth(a) = FormIn.TextWidth(GridIn.TextMatrix(b, a)) * 5
           End If
        End If
    Next b
  Next a

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function AutoSizeCells_ExZeroLngth(GridIn As MSFlexGrid, FormIn As Form)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "AutoSizeCells_ExZeroLngth"

  For a = 0 To GridIn.Cols - 1
    If GridIn.ColWidth(a) = 0 Then Next a
    If GridIn.TextMatrix(0, a) = "" Then Next a
    GridIn.ColWidth(a) = FormIn.TextWidth(GridIn.TextMatrix(0, a)) * 1.5
  Next a
  For a = 0 To GridIn.Cols - 1
    For b = 0 To GridIn.Rows - 1
        If GridIn.ColWidth(a) = 0 Then Next b
        If Len(GridIn.TextMatrix(b, a)) = 1 Then
            GridIn.ColWidth(a) = FormIn.TextWidth(GridIn.TextMatrix(b, a)) * 5
        End If
        If GridIn.TextMatrix(b, a) = "" Then Next b
        If GridIn.ColWidth(a) < FormIn.TextWidth(GridIn.TextMatrix(b, a)) * 1.5 Then
            GridIn.ColWidth(a) = FormIn.TextWidth(GridIn.TextMatrix(b, a)) * 1.5
        End If
    Next b
  Next a

ProcExit:
  Exit Function '

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function AutoSizeCells_WeeklyDataEntry(GridIn As MSFlexGrid, FormIn As Form)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "AutoSizeCells_WeeklyDataEntry"

  For a = 0 To GridIn.Cols - 1
    If a = 0 Then Next a
    If GridIn.TextMatrix(0, a) = "" Then Next a
    If GridIn.ColWidth(2) = 0 Then Next a
    If GridIn.ColWidth(3) = 0 Then Next a
    If GridIn.ColWidth(4) = 0 Then Next a
    GridIn.ColWidth(a) = FormIn.TextWidth(GridIn.TextMatrix(0, a)) * 1.5
  Next a

  For a = 0 To GridIn.Cols - 1
    For b = 0 To GridIn.Rows - 1
        If GridIn.ColWidth(a) = 0 Then Next b
        If GridIn.ColWidth(a) < FormIn.TextWidth(GridIn.TextMatrix(b, a)) * 1.5 Then
           GridIn.ColWidth(a) = FormIn.TextWidth(GridIn.TextMatrix(b, a)) * 1.5
        End If
    Next b
  Next a '

ProcExit:
  Exit Function '

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function Fill_NoData(GridIn As MSFlexGrid)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "Fill_NoData"

  For a = GridIn.FixedCols To GridIn.Cols - 1
    For b = GridIn.FixedRows To GridIn.Rows - 1
        If GridIn.TextMatrix(b, a) = "" Then GridIn.TextMatrix(b, a) = "NoData"
    Next b
  Next a

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function RemoveRow_Engine(GridIn As MSFlexGrid, _
                                     ArrayIn() As String, _
                                     ByVal ArrayBeginIn As Integer, _
                                     ByVal ArrayEndIn As Integer, _
                                     ByVal LogicIn As String)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "RemoveRow_Engine"

  Dim rmv As Boolean
  rmv = True
  If LogicIn = "And" Then

  End If

  If LogicIn = "Or" Then
    For a = 1 To GridIn.Rows - 1
      For b = 0 To GridIn.Cols - 1
        For c = ArrayBeginIn To ArrayEndIn
          If ArrayIn(c) = GridIn.TextMatrix(a, b) Then
            rmv = False
            GoTo 1
          End If
        Next c
      Next b
1:
      If rmv = True Then
        RemoveRow_Data GridIn, a
      End If
      rmv = True
    Next a
  End If

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function RemoveRow_Data(GridIn As MSFlexGrid, ByVal RowNbrIn As Integer)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "RemoveRow_Data"

  For a = 0 To GridIn.Cols - 1
    GridIn.TextMatrix(RowNbrIn, a) = ""
  Next a

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function TrimEnds(GridIn As MSFlexGrid)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "TrimEnds"

  For a = 0 To GridIn.Cols - 1
    For b = 0 To GridIn.Rows - 1
        GridIn.TextMatrix(b, a) = Trim(GridIn.TextMatrix(b, a))
    Next b
  Next a

ProcExit:
  Exit Function ''

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function CapFirstLetter(GridIn As MSFlexGrid)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "CapFirstLetter"

  Dim FirstLet As String
  Dim str As String
  For a = 0 To GridIn.Cols - 1
    For b = 0 To GridIn.Rows - 1
      If GridIn.TextMatrix(b, a) = "" Then Next b
      FirstLet = vba.Left$(GridIn.TextMatrix(b, a), 1)
      str = vba.Right$(GridIn.TextMatrix(b, a), Len(GridIn.TextMatrix(b, a)) - 1)
      FirstLet = UCase(FirstLet)
      str = LCase(str)
      GridIn.TextMatrix(b, a) = FirstLet & str
    Next b
  Next a

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function TwoIdenticalRows(GridIn As MSFlexGrid, ByVal Col1In As Integer, ByVal Col2In As Integer) As Boolean
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "TwoIdenticalRows"

  TwoIdenticalRows = False

  For a = 0 To GridIn.Rows - 2
    For b = a + 1 To GridIn.Rows - 1
        If GridIn.TextMatrix(a, Col1In) = "" And GridIn.TextMatrix(a, Col2In) = "" Then GoTo SkipTest
        If GridIn.TextMatrix(a, Col1In) = GridIn.TextMatrix(b, Col1In) And _
           GridIn.TextMatrix(a, Col2In) = GridIn.TextMatrix(b, Col2In) Then
           TwoIdenticalRows = True
        End If
SkipTest:
    Next b
  Next a

ProcExit:
  Exit Function '

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function AutoSizeGridToForm(GridIn As MSFlexGrid, FormIn As Form)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "AutoSizeGridToForm"

  If FormIn.Height < 2000 Then GoTo ProcExit:
  GridIn.Left = 0
  GridIn.Top = 0
  GridIn.Width = FormIn.Width - 200
  GridIn.Height = FormIn.Height - 700

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function TrimGrdRows(GridIn As MSFlexGrid)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "TrimGrdRows"
  
  Dim Emptie As Boolean
  Dim lngth As Integer

  Emptie = True
  lngth = GridIn.Rows - 1

  For a = 0 To lngth
    For b = 0 To GridIn.Cols - 1
      If GridIn.TextMatrix(lngth - a, b) <> "" Then Emptie = False
    Next b
    If Emptie = True Then GridIn.Rows = GridIn.Rows - 1
    If Emptie = False Then Exit Function
    Emptie = True
  Next a

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function SetButtonToCol(GridIn As MSFlexGrid, CmdIn As CommandButton, ByVal ColIn As Integer)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "SetButtonToCol"

  CmdIn.Left = GridIn.Left + 35 + GridIn.CellLeft
  CmdIn.Top = GridIn.Top + 35
  CmdIn.Width = GridIn.CellWidth
  CmdIn.Height = GridIn.CellHeight
  CmdIn.Caption = GridIn.TextMatrix(0, ColIn)
  CmdIn.Visible = True

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

Public Function SetButtonToCell(GridIn As MSFlexGrid, CmdIn As CommandButton)
  On Error GoTo ErrRtn
  gstrChkPt = "On Error": gstrProcName = "SetButtonToCell"

  CmdIn.Left = GridIn.Left + 50 + GridIn.CellLeft
  CmdIn.Top = GridIn.Top + 50 + GridIn.CellTop
  CmdIn.Width = GridIn.CellWidth
  CmdIn.Height = GridIn.CellHeight
  CmdIn.Visible = True

ProcExit:
  Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Function

