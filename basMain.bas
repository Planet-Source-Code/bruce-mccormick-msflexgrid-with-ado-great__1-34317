Attribute VB_Name = "basMain"
Option Explicit

Private Const strModuleName As String * 30 = "basMain"

Global Const gstrCompNameLong As String = "Dewey Cheatem and Howe, Inc."
Global Const gstrCompNameShort As String = "DCH"
Global Const gstrMainDB As String = "PII.mdb"
Global Const gstrMainDBType As String = "7"  ' see basADO Public Enum e_DBTypes
Global Const gstrAdminDB As String = "AdminDB.mdb"

' For gstrCapWhat: 0 = Don't Change What They Enter, 1 = ALL UPPER, 2 = all lower, 3 = First Letter
Global Const gstrCapWhat As Integer = 3  ' not yet implemented

Global gMainConn As New ADODB.Connection
Global gdbMain As DAO.Database

Global gintMsgBoxAns As Integer
Global gstrOpenStmt As String
Global gblnIsDirty As Boolean
Global gstrProcName As String * 40
Global gstrChkPt As String * 60
Global gstrDBPath As String
Global gstrUserID As String * 156
Global gstrMachineID As String * 30
Global gstrDecPtChar As String * 1
Global gstrCommaChar As String * 1
Global gstrDBTables() As String

Sub Main()
   gstrChkPt = "On Error": gstrProcName = "basMain"
   On Error GoTo ErrRtn
   
   If App.PrevInstance = True Then
      MsgBox "This Application Is Already Running. You Cannot Start it Again."
      End
   End If
   
 '  frmLogon.Show vbModal
 '  If UserID = 0 Then
 '       exit sub
 '  End If
   
   Screen.MousePointer = 11
   frmStartUp.Show
   frmStartUp.Refresh

   ChDrive App.Path
   ChDir App.Path
   
   App.Title = "(" & gstrCompNameShort & " Mgmt Rpts (v" & App.Major & _
               "." & App.Minor & App.Revision & ")"
   
   gstrChkPt = "gstrDBPpath ="
   gstrDBPath = App.Path & "\" & gstrMainDB
   
   If Not basFiles.FileExists(gstrDBPath) Then
      Call basDB.FindDB(App.Path, gstrMainDB)
   End If
   
   gstrChkPt = "gstrUserId"
   gstrUserID = libUtilities.GetUserID   ' Call
   gstrDecPtChar = basNumerics.RegionDecimalPoint  ' Call
   
   If gstrDecPtChar = "." Then
      gstrCommaChar = ","
   Else
      gstrCommaChar = "."
   End If
   
   gstrChkPt = "Load frmMainMenu"
   Load frmMainMenu
   frmMainMenu.Show
   gblnIsDirty = True

   gstrChkPt = "Unload frmStartUp"
   Unload frmStartUp
   Screen.MousePointer = 0
   
ProcExit:
   Exit Sub

ErrRtn:
    Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
    Resume ProcExit:
End Sub

Sub EndProgram()
   gstrChkPt = "On Error": gstrProcName = "EndProgram"
   On Error GoTo ErrRtn
   
   Private Const clngMsgbxExitApp As Long = vbDefaultButton2 + vbQuestion + vbYesNo
   
   Dim blnExitApp As Boolean
   Dim intCancel As Integer
   Dim OpenForm As VB.Form
    
   blnExitApp = (MsgBox("Save Changes and Exit?", clngMsgbxExitApp, "Exit Application") = vbYes)
   If blnExitApp = False Then Exit Sub
    
   If gblnIsDirty Then
     intCancel = CInt(blnExitApp = False)
     If Not intCancel Then
       For Each OpenForm In Forms
         Unload OpenForm
         OpenForm = Nothing
        Next
     End If
   End If
   
   Call basADO.rsClose(rs, gMainConn)
   Call basADO.ConnClose(gMainConn)
   
   Set gdbMain = Nothing

ProcExit:
    Exit Sub

ErrRtn:
    Call ErrMsg(strModuleName, gstrProcName, gstrChkPt, Err.Number, Err.Description, Err.Source)
    Resume ProcExit:
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call EndProgram
End Sub

