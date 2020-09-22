VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPurchMaint 
   Caption         =   "Purchasing Maintenance"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   9000
      Width           =   1455
   End
   Begin VB.ComboBox cboVendNbr 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3720
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox cboLoc 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtDt 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Text            =   " "
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   5280
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   7965
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   14049
      _Version        =   393216
      BackColor       =   12632256
      FocusRect       =   2
      HighLight       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Loc"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmPurchMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const strModuleName As String * 30 = "frmPurchMaint"

Private Const strEditMask As String * 12 = "#,###,###.00"

Private rs As ADODB.Recordset

Private strCurrTable As String
Private strCurrKey As String
Private strRsetFld As String * 12

Private blnGoingRight As Boolean
Private blnAddRow As Boolean
Private blnRowChanged As Boolean
Private blnInFormLoad As Boolean
Private blnManualCall As Boolean 'used to tell the diff betw an actual mouse click and a 'call'
                                 ' to the grid_Click event
Private intMsgAns As Integer
Private intRowCurr As Integer
Private intRowWork As Integer
Private intColCurr As Integer
Private intColWork As Integer

Private dblRowTotal As Double
Private dblColTotal As Double
Private dblGrandTot As Double

Private i As Integer
Private j As Integer

Private Sub grid_Click()
   gstrChkPt = "On Error": gstrProcName = "grid_Click"
   On Error GoTo ErrRtn
  
   ' I only want to call ExitCell here if we got to grid_Click via an actual grid_Click, ie.,
   ' below I call grid_Click from the KyeCodeActions proc and I may have already accounted for
   ' the grid_Click in that code
   If Not blnManualCall Then
      Call ExitCell
      gstrChkPt = "After ExitCell in 'If not blnManualCall'": gstrProcName = "grid_Click"
   End If
  
   With grid
      ' don't allow click into rows around the edge -
      ' row 0 = hdings, rows -1 = totals, col 0 = status, cols - 1 = row totals
      If Not blnInFormLoad Then
         If .Row < .FixedRows Or _
            .Col < .FixedCols Or _
            .Col >= .Cols Or _
            .Row >= .Rows Then
            intMsgAns = MsgBox("You Can Not Click Into The Outside Row/Col", vbOKOnly)
            Exit Sub
         End If
      End If
  
      gstrChkPt = "If .col <> 1": gstrProcName = "grid_Click"
      If .Col <> 1 Then cboVendNbr.Visible = False
  
      If .Col = 1 Then
         gstrChkPt = "If .col = 1"
         Text1.Visible = False
         cboVendNbr.Width = .CellWidth - 10
         cboVendNbr.Top = .CellTop + .Top
         cboVendNbr.Left = .CellLeft + .Left
         cboVendNbr.Text = cboVendNbr.Text
         cboVendNbr.Visible = True
         cboVendNbr.SetFocus
         cboVendNbr.ZOrder
      ElseIf .Col = 3 And blnGoingRight = True Then
         gstrChkPt = "If Not IsDate(.Text)"
         If Not IsDate(.Text) Then
            frmGetDate.SelDate = Now - 1
         Else
            frmGetDate.SelDate = CDate(.Text)
         End If
         .Text = Format(frmGetDate.SelDate, "mm/dd/yyyy")
         Call SetUpTextBox
         
         frmGetDate.Go   ' Call
         .Text = Format(frmGetDate.SelDate, "mm/dd/yyyy")
         .Col = 4
         Call SetUpTextBox
         gstrChkPt = "After SetupTextBox in 'ElseIf .Col = 3'": gstrProcName = "grid_Click"
      Else
         gstrChkPt = "Else"
         Call SetUpTextBox
      End If
   End With
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

Private Sub SetUpTextBox()
   gstrChkPt = "On Error": gstrProcName = "SetUpTextBox"
   On Error GoTo ErrRtn

   With Text1
      .Width = grid.CellWidth - 10
      .Height = grid.CellHeight - 10
      .Top = grid.CellTop + grid.Top
      .Left = grid.CellLeft + grid.Left
      .ZOrder
      .Visible = True
      .Text = grid.Text
      .SetFocus
      .SelStart = 0
      .SelLength = Len(grid.Text)
   End With
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

Private Sub cboVendNbr_Change()
   gstrChkPt = "On Error": gstrProcName = "cboVendNbr_Change"
   On Error GoTo ErrRtn

   grid.TextMatrix(grid.Row, grid.Col) = cboVendNbr.Text

ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

Private Sub cboVendNbr_KeyDown(KeyCode As Integer, Shift As Integer)
   gstrChkPt = "On Error": gstrProcName = "cboVendNbr_KeyDown"
   On Error GoTo ErrRtn

   Call KeycodeActions(KeyCode, Shift)
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   gstrChkPt = "On Error": gstrProcName = "Text1_KeyDown"
   On Error GoTo ErrRtn

   Call KeycodeActions(KeyCode, Shift)
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

Private Sub KeycodeActions(KeyCode As Integer, Shift As Integer)
   gstrChkPt = "On Error": gstrProcName = "KeycodeActions"
   On Error GoTo ErrRtn

   Dim i As Integer
  ' if we're going to be leaving a cell, validate the data before leaving
'   Select Case KeyCode
'      Case vbKeyDelete, vbKeyEscape
'
'      Case Else
'         If IsCellValid = False Then
'            Exit Sub
'         End If
'   End Select
    
   With grid
      Select Case KeyCode
       ' Generally the way this works is:
       '    If we're leaving a cell then execute the ExitCell proc
       '    If we're leaving a row, execute the UpdateDB proc
       '    After moving to the next cell, run grid_click to move the text or combo box
       '       over the top of the grid
         Case vbKeyReturn
            ' if we're in the first col of the first row and there's nothing in
            ' that cell, then this is the first record, so mark it as "new"
            If .Col = 1 And .Row = 1 And _
               .TextMatrix(1, 1) = " " Then
               .TextMatrix(1, 0) = "N"
               Call ExitCell
               gstrChkPt = ".col = 2 in 'If .col = 1": gstrProcName = "KeyCodeActions"
               .Col = 2
               blnManualCall = True: Call grid_Click: blnManualCall = False
               ' I use - 2 here because the last col is the total col which they can't access
            ElseIf .Col = .Cols - 2 Then
               blnAddRow = True
               Call CheckForUpdate
            Else
               Call ExitCell
               gstrChkPt = ".col = .Col + 1": gstrProcName = "KeyCodeActions"
               .Col = .Col + 1
               blnGoingRight = True
               blnManualCall = True: Call grid_Click: blnManualCall = False
               blnGoingRight = False
            End If
         Case vbKeyInsert
            blnAddRow = True
            Call CheckForUpdate
            gstrChkPt = "After Update in KeyInsert": gstrProcName = "KeyCodeActions"
         Case vbKeyRight
            If .Col < .Cols - 2 Then ' < cols -2 because last col is total col
               Call ExitCell
               gstrChkPt = ".col = .col + 1 in if . col < .cols - 2": gstrProcName = "KeyCodeActions"
               .Col = .Col + 1
               blnGoingRight = True
               blnManualCall = True: Call grid_Click: blnManualCall = False
               blnGoingRight = False
            ElseIf .Col = .Col - 2 Then
               blnAddRow = True
               Call CheckForUpdate
               gstrChkPt = "After Update in KeyRight": gstrProcName = "KeyCodeActions"
            End If
         Case vbKeyLeft
            If .Col > 1 Then
               Call ExitCell
               gstrChkPt = ".col = .col - 1 in 'vbKeyleft": gstrProcName = "KeyCodeActions"
               .Col = .Col - 1
               blnManualCall = True: Call grid_Click: blnManualCall = False
            End If
         Case vbKeyUp
            If .Col = 1 Then Exit Sub
            If .Row > 1 Then
               blnAddRow = False
               Call CheckForUpdate
               gstrChkPt = ".row = .rwo - 1 in vbKeyUp": gstrProcName = "KeyCodeActions"
               .Row = .Row - 1
               blnManualCall = True: Call grid_Click: blnManualCall = False
            End If
         Case vbKeyDown
            If .Col = 1 Then Exit Sub
                 
            If .Row < .Rows - 2 Then
               blnAddRow = False
               Call CheckForUpdate
               gstrChkPt = "row = .row - 1 in vbkeydown": gstrProcName = "KeyCodeActions"
               grid.Row = .Row + 1
               blnManualCall = True: Call grid_Click: blnManualCall = False
            ElseIf .Row = .Rows - 2 Then
               blnAddRow = True
               Call CheckForUpdate
            End If
         Case vbKeyPageUp
            blnAddRow = False
            Call CheckForUpdate
            gstrChkPt = "If .row > 20 in vbKeyPageUp": gstrProcName = "KeyCodeActions"
            If .Row > 20 Then
               .Row = .Row - 20
            Else
               .Row = 1
            End If
            .Col = 1
            blnManualCall = True: Call grid_Click: blnManualCall = False
         Case vbKeyPageDown
            blnAddRow = False
            Call CheckForUpdate
            gstrChkPt = "if .row < .rows - 22 in vbKeyPageDown": gstrProcName = "KeyCodeActions"
            If .Row < .Rows - 22 Then
               .Row = .Row + 20
            Else
               .Row = .Rows - 2 ' remember that the last row (.rows -1) is saved for totals
            End If
            .Col = 1
            blnManualCall = True: Call grid_Click: blnManualCall = False
         Case vbKeyHome
            blnAddRow = False
            Call CheckForUpdate
            gstrChkPt = ".col = 1 in vbKeyHome": gstrProcName = "KeyCodeActions"
            .Col = 1
            .Row = 1
            blnManualCall = True: Call grid_Click: blnManualCall = False
         Case vbKeyEnd
            blnAddRow = False
            Call CheckForUpdate
            gstrChkPt = "vbKeyEnd": gstrProcName = "KeyCodeActions"
            .Col = 1
            .Row = .Rows - 2
            blnManualCall = True: Call grid_Click: blnManualCall = False
         Case vbKeyEscape
            cmdExit_Click   ' Call
         Case vbKeyDelete
            intMsgAns = MsgBox("Delete the record permanently?", vbYesNo + vbQuestion, "Delete ?")
            If intMsgAns = vbYes Then
               .TextMatrix(.Row, 0) = "D"
               Call UpdateDB
               gstrChkPt = "blnManualCall in vbkeyDelete": gstrProcName = "KeyCodeActions"
               blnManualCall = True: Call grid_Click: blnManualCall = False
               gstrChkPt = "After grid_click in vbkeyDelete": gstrProcName = "KeyCodeActions"
            End If
      End Select
   End With
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

Private Sub CheckForUpdate()
   On Error GoTo Err
   gstrChkPt = "On Error GOTo Err": gstrProcName = "CheckForUpdate"
   With grid
      If .TextMatrix(.Row, 0) = "A" Or _
         .TextMatrix(.Row, 0) = "M" Or _
         .TextMatrix(.Row, 0) = "D" Then
         Call ExitCell
         gstrChkPt = "UpdateDB": gstrProcName = "CheckForUpdate"
         Call UpdateDB
         gstrChkPt = "After UpdateDB": gstrProcName = "CheckForUpdate"
      End If
   End With
   
   If blnAddRow = True Then
      gstrChkPt = "Call AddRow": gstrProcName = "CheckForUpdate"
      Call AddRow
      gstrChkPt = "After Call AddRow": gstrProcName = "CheckForUpdate"
   End If

ProcExit:
  Exit Sub

Err:
  Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
  Resume ProcExit:

End Sub

Private Sub AddRow()
   gstrChkPt = "On Error": gstrProcName = "AddRow"
   On Error GoTo ErrRtn
   
   With grid
      .Row = .Rows - 2
      If .TextMatrix(.Row, 0) <> "N" Then
         ' add a new Row
         .Rows = .Rows + 1
         .Row = .Rows - 2
           
         ' move the totals down a row and clear the new cols to avoid nulls
         For i = 1 To .Cols - 1
           If i > 4 Then .TextMatrix(.Rows - 1, i) = .TextMatrix(.Row, i)
           .TextMatrix(.Row, i) = " "
         Next
         
         intRowWork = .Row
            
         Call basFlexGrid.GridColors(grid, 192, 255, 192)
         gstrChkPt = ".row = introwwork": gstrProcName = "AddRow"
            
         .Row = intRowWork
         .Col = 1
         .TextMatrix(.Row, 0) = "N"
          blnManualCall = True: Call grid_Click: blnManualCall = False
      ' make sure they've entered the necessary key info before allowing them to
      ' move to the next row
      ElseIf .TextMatrix(.Row, 1) = " " Or _
             .TextMatrix(.Row, 2) = " " Then
         .Col = 1
         blnManualCall = True: Call grid_Click: blnManualCall = False
      End If
   End With
   
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

Private Sub ClearRow()
   gstrChkPt = "On Error": gstrProcName = "ClearRow"
   On Error GoTo ErrRtn
  
   With grid
      For i = 1 To .Cols - 1
         .TextMatrix(.Row, i) = ""
      Next i
   End With
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
   Resume ProcExit:
End Sub

Private Sub ExitCell()
   gstrChkPt = "On Error": gstrProcName = "ExitCell"
   On Error GoTo ErrRtn
  
   If blnInFormLoad = True Then Exit Sub
   
   gstrChkPt = "With Grid"
   With grid
      ' if we're leaving the first col, update the grid from the cbo
      ' also plug text1.text with the vendor number so all of the rest
      ' of the code can fall through as if it was the text box, not the cbo
      If .Col = 1 Then
         If Len(cboVendNbr.Text) <> 0 Then
            .Text = cboVendNbr.Text
            Text1.Text = cboVendNbr.Text
         Else
            .Text = " "
            Text1.Text = " "
         End If
         cboVendNbr.Visible = False
      End If
      
      gstrChkPt = "If Text1"
      If Text1.Text <> .TextMatrix(.Row, .Col) Then
         ' Make sure that they've entered the key data before they add/change anything else
         If .Col > 2 And .Col < 15 And _
            (cboLoc.Text = " " Or txtDt = " " Or _
            .TextMatrix(.Row, 1) = " " Or .TextMatrix(.Row, 2) = " ") Then
            intMsgAns = MsgBox("You Must Enter A Loc, Trx Dt, Vendor & Invc Nbr In Order Add A Record", _
                               vbOKOnly & vbExclamation, "Missing Information")
            Text1.Text = " "

            If .TextMatrix(.Row, 1) = " " Then
               .Col = 1
            Else
               .Col = 2
            End If
            Call grid_Click
            
            gstrChkPt = "Exit Sub": gstrProcName = "ExitCell"
            Exit Sub
         End If

         If .Col > 4 Then
            RSet strRsetFld = Format(Text1.Text, strEditMask)
            .Text = strRsetFld
         Else
            .Text = Text1.Text
         End If

         Text1.Visible = False
         Text1.Text = " "
         
         gstrChkPt = "If (.textmatrix"
         ' if col 1 and 2 contain info and the record
         ' hasn't been looked up yet, look it up
         If (.TextMatrix(.Row, 1) <> " " And _
             .TextMatrix(.Row, 2) <> " ") And _
             .TextMatrix(.Row, 0) <> "A" And _
             .TextMatrix(.Row, 0) <> "M" And _
             .TextMatrix(.Row, 0) <> "D" Then
            strCurrTable = "Purchases"
            strCurrKey = cboLoc & " " & txtDt & " " & .TextMatrix(.Row, 1) & " " & _
               .TextMatrix(.Row, 2)
            gstrOpenStmt = "Select * from Purchases " & _
                "Where VendNbr = '" & .TextMatrix(.Row, 1) & "'" & _
                " and InvcNbr = '" & .TextMatrix(.Row, 2) & "'"
            
            Call basADO.rsOpen(rs, gMainConn)
            gstrChkPt = "If rs.recordCount - 0": gstrProcName = "ExitCell"
            
            If rs.RecordCount = 0 Then ' its a new record
               .TextMatrix(.Row, 0) = "A"
            Else
               ' see if the record is already in the grid
               ' if so, go to that line so we don't dup the record in the grid
               ' and hence screw up the totals with dups
               intRowWork = 0
               For i = 1 To .Rows - 2
                  If (.TextMatrix(i, 1) = .TextMatrix(.Row, 1) And _
                      .TextMatrix(i, 2) = .TextMatrix(.Row, 2)) And _
                      i <> .Row Then
                      intRowWork = i
                      Exit For
                  End If
               Next i
               
               If intRowWork <> 0 Then
                  Call ClearRow
                  .TextMatrix(.Row, 0) = "N"
                  .Row = intRowWork
                  .Col = 1
               Else
                  intMsgAns = MsgBox("Invoice " & .TextMatrix(.Row, 2) & " Already Exists For " & _
                               "Vendor" & .TextMatrix(.Row, 1) & ". Do You Want To Update It? ", _
                               vbYesNo)
                  If intMsgAns = vbNo Then
                     Call ClearRow
                     gstrChkPt = ".TextMatrix(.Row, 0) = N": gstrProcName = "ExitCell"
                     .TextMatrix(.Row, 0) = "N"
                     .Col = 1
                     Call grid_Click
                     gstrChkPt = "After Grid_Click": gstrProcName = "ExitCell"
                  
                     Call basADO.rsClose(rs, gMainConn)
                     gstrChkPt = "Exit Sub after rsClose": gstrProcName = "ExitCell"
                     Exit Sub
                  End If
               End If
               
               .TextMatrix(.Row, 0) = "M"
               .TextMatrix(.Row, 3) = rs!InvcDt
               .TextMatrix(.Row, 4) = rs!AcctNbr
               RSet strRsetFld = Format(rs!Supplies, strEditMask): .TextMatrix(.Row, 5) = strRsetFld
               RSet strRsetFld = Format(rs!Gifts, strEditMask): .TextMatrix(.Row, 6) = strRsetFld
               RSet strRsetFld = Format(rs!Bev, strEditMask): .TextMatrix(.Row, 7) = strRsetFld
               RSet strRsetFld = Format(rs!Bread, strEditMask): .TextMatrix(.Row, 8) = strRsetFld
               RSet strRsetFld = Format(rs!Dairy, strEditMask): .TextMatrix(.Row, 9) = strRsetFld
               RSet strRsetFld = Format(rs!Groceries, strEditMask): .TextMatrix(.Row, 10) = strRsetFld
               RSet strRsetFld = Format(rs!Meats, strEditMask): .TextMatrix(.Row, 11) = strRsetFld
               RSet strRsetFld = Format(rs!Produce, strEditMask): .TextMatrix(.Row, 12) = strRsetFld
               RSet strRsetFld = Format(rs!Seafood, strEditMask): .TextMatrix(.Row, 13) = strRsetFld
               RSet strRsetFld = Format(rs!Liquor, strEditMask): .TextMatrix(.Row, 14) = strRsetFld
               RSet strRsetFld = Format(rs!TotalCost, strEditMask): .TextMatrix(.Row, 15) = strRsetFld
               .Col = 2
               blnManualCall = True: Call grid_Click: blnManualCall = False
            End If
            Call basADO.rsClose(rs, gMainConn)
            gstrChkPt = "After rsClose br Recalc": gstrProcName = "ExitCell"
         End If
         Call RecalcTotals
         gstrChkPt = "After Recalc": gstrProcName = "ExitCell"
      End If       ' if Text1.Text <>
      intRowCurr = .Row
      intColCurr = .Col
   End With
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

Private Sub RecalcTotals()
   On Error GoTo ErrRtn
   gstrChkPt = "On Error": gstrProcName = "RecalcTotals"
   
   With grid
      ' recalc totals when we leave any of the dollar-amt cells
      dblRowTotal = 0
      dblColTotal = 0
      dblGrandTot = 0
      
      If .Col > 4 Then
         RSet strRsetFld = Format(.TextMatrix(.Row, .Col), strEditMask)
         .TextMatrix(.Row, .Col) = strRsetFld
      End If
      
      For i = 1 To .Rows - 2
         For j = 5 To 14
            dblRowTotal = dblRowTotal + CDbl(Val(vba.replace(.TextMatrix(i, j), ",", "")))
            If j = 14 Then
               RSet strRsetFld = Format(dblRowTotal, strEditMask)
               .TextMatrix(i, 15) = strRsetFld
               dblGrandTot = dblGrandTot + CDbl(Val(vba.replace(.TextMatrix(i, 15), ",", "")))
               dblRowTotal = 0
            End If
         Next j
      Next i
      
      For j = 5 To 14
         For i = 1 To .Rows - 2
            dblColTotal = dblColTotal + CDbl(Val(vba.replace(.TextMatrix(i, j), ",", "")))
            If i = .Rows - 2 Then
               RSet strRsetFld = Format(dblColTotal, strEditMask)
               .TextMatrix(.Rows - 1, j) = strRsetFld
               dblColTotal = 0
            End If
         Next i
      Next j
      
      RSet strRsetFld = Format(dblGrandTot, strEditMask)
      .TextMatrix(.Rows - 1, 15) = strRsetFld

   End With
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

Private Function IsCellValid() As Boolean
   ' This is an editing routine - not yet implemented
   gstrChkPt = "On Error": gstrProcName = "IsCellValid"
   On Error GoTo ErrRtn
  
   Dim blnTest As Boolean

   IsCellValid = True
   
   With grid
      Select Case .Col
         Case 0
         Case 1
         Case 2
         Case 3
         Case 4
         Case 5
         Case 6
         Case 7
         Case 8
         Case 9
         Case 10
         Case 11
         Case 12
         Case 13
         Case 14
         Case 15
         Case 16
         Case 17
         Case 18
      End Select
  End With
  
ProcExit:
   Exit Function

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Function

Private Sub UpdateDB()
   gstrChkPt = "On Error": gstrProcName = "UpdateDB"
   On Error GoTo ErrRtn
  
   Select Case grid.TextMatrix(grid.Row, 0)
      Case "A"
         gstrChkPt = "Add"
         gstrOpenStmt = "Purchases"
         Call basADO.rsOpen(rs, gMainConn)
         gstrChkPt = "rs.AddNew": gstrProcName = "UpdateDB"
         rs.AddNew
         rs!Loc = cboLoc.Text
         rs!Dt = CDate(txtDt)
         rs!VendNbr = grid.TextMatrix(grid.Row, 1)
         rs!InvcNbr = grid.TextMatrix(grid.Row, 2)
         
      Case "M"
         gstrChkPt = "Modify"
         gstrOpenStmt = "SELECT * FROM Purchases" & _
                        " Where VendNbr = '" & grid.TextMatrix(grid.Row, 1) & "' " & _
                        "   and InvcNbr = '" & grid.TextMatrix(grid.Row, 2) & "'"
         Call basADO.rsOpen(rs, gMainConn)
      
         gstrChkPt = "If cboLoc <> rs!Loc": gstrProcName = "UpdateDB"
      
         ' only update key values if they have changed
         If cboLoc.Text <> rs!Loc Then rs!Loc = cboLoc.Text
         If CDate(txtDt) <> rs!Dt Then rs!Dt = CDate(txtDt)
         If grid.TextMatrix(grid.Row, 1) <> rs!VendNbr Then rs!VendNbr = grid.TextMatrix(grid.Row, 1)
         If grid.TextMatrix(grid.Row, 2) <> rs!InvcNbr Then rs!InvcNbr = grid.TextMatrix(grid.Row, 2)
   End Select
   
   Select Case grid.TextMatrix(grid.Row, 0)
      Case "A", "M"
         rs!InvcDt = CDate(grid.TextMatrix(grid.Row, 3))
         rs!InvcMo = month(CDate(grid.TextMatrix(grid.Row, 3)))
         rs!InvcDa = day(CDate(grid.TextMatrix(grid.Row, 3)))
         rs!InvcYr = year(CDate(grid.TextMatrix(grid.Row, 3)))
         rs!AcctNbr = grid.TextMatrix(grid.Row, 4)
         rs!Supplies = Val(vba.replace(grid.TextMatrix(grid.Row, 5), ",", ""))
         rs!Gifts = Val(vba.replace(grid.TextMatrix(grid.Row, 6), ",", ""))
         rs!Bev = Val(vba.replace(grid.TextMatrix(grid.Row, 7), ",", ""))
         rs!Bread = Val(vba.replace(grid.TextMatrix(grid.Row, 8), ",", ""))
         rs!Dairy = Val(vba.replace(grid.TextMatrix(grid.Row, 9), ",", ""))
         rs!Groceries = Val(vba.replace(grid.TextMatrix(grid.Row, 10), ",", ""))
         rs!Meats = Val(vba.replace(grid.TextMatrix(grid.Row, 11), ",", ""))
         rs!Produce = Val(vba.replace(grid.TextMatrix(grid.Row, 12), ",", ""))
         rs!Seafood = Val(vba.replace(grid.TextMatrix(grid.Row, 13), ",", ""))
         rs!Liquor = Val(vba.replace(grid.TextMatrix(grid.Row, 14), ",", ""))
         rs!TotalCost = Val(vba.replace(grid.TextMatrix(grid.Row, 15), ",", ""))
         rs!EnteredDt = Now
         rs!EnteredBy = libUtilities.GetUserID

         rs.Update
         Call basADO.rsClose(rs, gMainConn)
         gstrChkPt = "grid.TextMatrix(frid.Row, 0) = S": gstrProcName = "UpdateDB"
         grid.TextMatrix(grid.Row, 0) = "S" 'save marker
      
      Case "D"
         gstrChkPt = "Delete"
         gstrOpenStmt = "SELECT * FROM Purchases" & _
                        " Where VendNbr = '" & grid.TextMatrix(grid.Row, 1) & "' " & _
                        "   and InvcNbr = '" & grid.TextMatrix(grid.Row, 2) & "'"
         Call basADO.rsDelete(rs, gMainConn)
         gstrChkPt = "If grid.Rows > 3": gstrProcName = "UpdateDB"
         If grid.Row = 1 And _
            grid.Rows = 3 Then
            Call ClearRow
            grid.TextMatrix(grid.Row, 0) = "N"
         Else
            grid.RemoveItem (grid.Row)
            grid.Row = grid.Rows - 2
         End If
         
         intRowCurr = grid.Row
         Call basFlexGrid.GridColors(grid, 192, 255, 192)
         Call RecalcTotals
         grid.Row = intRowCurr
         grid.Col = 1
   End Select

ProcExit:
   gstrChkPt = "UpdateDB Exit Sub": gstrProcName = "UpdateDB"
   Exit Sub

ErrRtn:
    Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
    Resume ProcExit:
End Sub
 
Private Sub txtDt_Click()
   gstrChkPt = "On Error": gstrProcName = "txtDt_Click"
   On Error GoTo ErrRtn
         
   If Not IsDate(txtDt.Text) Then
      frmGetDate.SelDate = Now - 1
   Else
      frmGetDate.SelDate = CDate(txtDt.Text)
   End If
   txtDt.Text = Format(frmGetDate.SelDate, "mm/dd/yyyy")
      
   frmGetDate.Go   ' Call
   txtDt.Text = Format(frmGetDate.SelDate, "mm/dd/yyyy")
   
   grid.SetFocus
   
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

Private Sub Form_Load()
   gstrChkPt = "On Error": gstrProcName = "Form_Load"
   On Error GoTo ErrRtn
  
   blnInFormLoad = True
  
   Dim i As Long
  
   grid.Appearance = flexFlat
   grid.Rows = 3 ' first for labels, 2nd for first row of data, 3rd for col totals
   grid.Cols = 16
   grid.FixedCols = 1
   grid.FixedRows = 1
   grid.RowHeightMin = Text1.Height

   RSet strRsetFld = Format("00", strEditMask)
      
   For i = 0 To grid.Cols - 1
      grid.Col = i
'      grid.CellFontSize = 10
      grid.ColAlignment(i) = flexAlignLeftCenter
    
      Select Case i
         Case 0
            grid.ColWidth(i) = 300
         Case 1
            grid.TextMatrix(0, i) = "Vendor"
            grid.ColWidth(i) = 1100
            grid.TextMatrix(1, i) = " "
         Case 2
            grid.TextMatrix(0, i) = "Invc Nbr"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
         Case 3
            grid.TextMatrix(0, i) = "Invc Dt"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
         Case 4
            grid.TextMatrix(0, i) = "Acct#"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
         Case 5
            grid.TextMatrix(0, i) = "Supplies"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
            grid.TextMatrix(2, i) = strRsetFld
         Case 6
            grid.TextMatrix(0, i) = "Gifts"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
            grid.TextMatrix(2, i) = strRsetFld
         Case 7
            grid.TextMatrix(0, i) = "Bev"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
            grid.TextMatrix(2, i) = strRsetFld
         Case 8
            grid.TextMatrix(0, i) = "Bread"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
            grid.TextMatrix(2, i) = strRsetFld
         Case 9
            grid.TextMatrix(0, i) = "Dairy"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
            grid.TextMatrix(2, i) = strRsetFld
         Case 10
            grid.TextMatrix(0, i) = "Groc"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
            grid.TextMatrix(2, i) = strRsetFld
         Case 11
            grid.TextMatrix(0, i) = "Meats"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
            grid.TextMatrix(2, i) = strRsetFld
         Case 12
            grid.TextMatrix(0, i) = "Prod"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
            grid.TextMatrix(2, i) = strRsetFld
         Case 13
            grid.TextMatrix(0, i) = "Seafood"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
            grid.TextMatrix(2, i) = strRsetFld
         Case 14
            grid.TextMatrix(0, i) = "Liquor"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
            grid.TextMatrix(2, i) = strRsetFld
         Case 15
            grid.TextMatrix(0, i) = "Total"
            grid.ColWidth(i) = 940
            grid.TextMatrix(1, i) = " "
            grid.TextMatrix(2, i) = strRsetFld
      End Select
   Next
              
   Call basFlexGrid.GridColors(grid, 192, 255, 192)
  
   gstrChkPt = "After GridColors": gstrProcName = "Form_Load"
   grid.Row = 1
   grid.Col = 1

   If grid.TextMatrix(1, 1) <> " " Then
      Text1.Text = grid.TextMatrix(1, 1)
   Else
      Text1.Text = " "
   End If
   
   Call basADO.ConnOpen(gMainConn, gstrMainDBType, gstrDBPath)
   gstrChkPt = "After ConnOpen": gstrProcName = "Form_Load"

   gstrOpenStmt = "SELECT loc FROM stores"

   Call basADO.rsOpen(rs, gMainConn)
   
   gstrChkPt = "If EmptyRS - Stores": gstrProcName = "Form_Load"
   If basADO.EmptyRS(rs, gMainConn) Then
      intMsgAns = MsgBox("There Are No Stores Defined. Please Do Stores/Maintenance " & _
                         "And Then Return To Here", vbOKOnly, "Stores Table Empty")
      cmdExit_Click   ' call
   End If
  
   Do Until rs.EOF
      cboLoc.AddItem rs.Fields(0)
      Call basADO.MvNext(rs, gMainConn)
      gstrChkPt = "After MvNext": gstrProcName = "Form_Load"
   Loop
  
   Call basADO.rsClose(rs, gMainConn)
  
   gstrChkPt = "gstrOpenStmt-Vendors": gstrProcName = "Form_Load"
   gstrOpenStmt = "SELECT VendCd FROM Vendors"
  
   Call basADO.rsOpen(rs, gMainConn)
  
   gstrChkPt = "If EmptyRS - Vendors": gstrProcName = "Form_Load"
   If basADO.EmptyRS(rs, gMainConn) Then
      intMsgAns = MsgBox("There Are No Vendors Defined. Please Do Vendors/Maintenance " & _
                         "And Then Return To Here", vbOKOnly, "Vendors Table Empty")
      cmdExit_Click   ' Call
   End If
  
   Do Until rs.EOF
     cboVendNbr.AddItem rs.Fields(0)
     Call basADO.MvNext(rs, gMainConn)
   Loop
   
   Call basADO.rsClose(rs, gMainConn)
   
   gstrChkPt = "If cboLoc.ListCount": gstrProcName = "Form_Load"
   If cboLoc.ListCount > 0 Then
      cboLoc.Text = cboLoc.List(0)
   End If
   
   If cboVendNbr.ListCount > 0 Then
      cboVendNbr.Text = cboVendNbr.List(0)
   End If
  
   cboVendNbr.Visible = True
   Text1.Visible = False
   txtDt = str(Date)
   Me.Show
  
   Call grid_Click
  
   gstrChkPt = "cboLoc.SetFocus": gstrProcName = "Form_Load"
   cboLoc.SetFocus
  
   blnInFormLoad = False

ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

Private Sub Form_Unload(Cancel As Integer)
   gstrChkPt = "On Error": gstrProcName = "Form_Unload"
   On Error GoTo ErrRtn
   
   blnRowChanged = False
   
   ' see if we have any rows that need to be saved before exiting
   With grid
      For i = 1 To .Rows - 2
         .Row = i
         If .TextMatrix(.Row, 0) = "A" Or _
            .TextMatrix(.Row, 0) = "M" Or _
            .TextMatrix(.Row, 0) = "D" Then
            blnRowChanged = True
            Exit For
         End If
      Next i
      
      If blnRowChanged = True Then
         intMsgAns = MsgBox("There Is Unsaved Data. Do You Wish To Save It Before Leaving?", _
                     vbYesNoCancel, "Save Data ?")
         If intMsgAns = vbCancel Then Exit Sub
         
         ' There should never be > 1 row that has not been updated, BUT . . .
         If intMsgAns = vbYes Then
            For i = 1 To .Rows - 2
               .Row = i
               If .TextMatrix(.Row, 0) = "A" Or _
                  .TextMatrix(.Row, 0) = "M" Or _
                  .TextMatrix(.Row, 0) = "D" Then
                  Call UpdateDB
               End If
            Next i
         End If
      End If
   End With
   
   Call basADO.rsClose(rs, gMainConn)
   
   gstrChkPt = "ConnClose": gstrProcName = "Form_Unload"
   Call basADO.ConnClose(gMainConn)
   gstrChkPt = "After ConnClose": gstrProcName = "Form_Unload"
   
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

Private Sub cmdExit_Click()
   gstrChkPt = "On Error": gstrProcName = "cmdExit_Click"
   On Error GoTo ErrRtn
   
   Unload Me
  
ProcExit:
   Exit Sub

ErrRtn:
   Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source, strCurrTable, strCurrKey)
   Resume ProcExit:
End Sub

