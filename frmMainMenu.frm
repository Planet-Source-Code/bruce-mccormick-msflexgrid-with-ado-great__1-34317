VERSION 5.00
Begin VB.Form frmMainMenu 
   Caption         =   "Pomeroy Investments"
   ClientHeight    =   60
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   60
   ScaleWidth      =   8430
   StartUpPosition =   1  'CenterOwner
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEmp 
      Caption         =   "&Employees"
      Begin VB.Menu mnuEmpMaint 
         Caption         =   "&Maintenance"
         Begin VB.Menu mnuEmpMaintEmp 
            Caption         =   "Employee Maintenance"
         End
         Begin VB.Menu mnuEmpMaintPayRate 
            Caption         =   "Pay Rates Maintenance"
         End
      End
      Begin VB.Menu mnuEmpRpts 
         Caption         =   "&Reports"
      End
   End
   Begin VB.Menu mnuLab 
      Caption         =   "&Labor"
      Begin VB.Menu mnuLabMaint 
         Caption         =   "&Maintenance"
      End
      Begin VB.Menu mnuLabRpts 
         Caption         =   "&Reports"
      End
   End
   Begin VB.Menu mnuPur 
      Caption         =   "&Purchases"
      Begin VB.Menu mnuPurMaint 
         Caption         =   "&Maintenance"
      End
      Begin VB.Menu mnuPurRpts 
         Caption         =   "&Reports"
      End
   End
   Begin VB.Menu mnuSales 
      Caption         =   "&Sales"
      Begin VB.Menu mnuSalesMaint 
         Caption         =   "&Maintenance"
      End
      Begin VB.Menu mnuSalesRpts 
         Caption         =   "&Reports"
      End
   End
   Begin VB.Menu mnuStores 
      Caption         =   "S&tores"
      Begin VB.Menu mnuStoresMaint 
         Caption         =   "&Maintenance"
      End
      Begin VB.Menu mnuStoresRpts 
         Caption         =   "&Reports"
      End
      Begin VB.Menu mnuStoresShifts 
         Caption         =   "&Shifts"
         Begin VB.Menu mnuStoresShiftsMaint 
            Caption         =   "&Maintenance"
         End
         Begin VB.Menu mnuStoresShiftsRpts 
            Caption         =   "&Reports"
         End
      End
   End
   Begin VB.Menu mnuTaxes 
      Caption         =   "Ta&xes"
      Begin VB.Menu mnuTaxesMaint 
         Caption         =   "&Maintenance"
      End
      Begin VB.Menu mnuTaxesRpts 
         Caption         =   "&Reports"
      End
   End
   Begin VB.Menu mnuVend 
      Caption         =   "&Vendors"
      Begin VB.Menu mnuVendMaint 
         Caption         =   "&Maintenance"
      End
      Begin VB.Menu mnuVendRpts 
         Caption         =   "&Reports"
      End
   End
   Begin VB.Menu mnuBud 
      Caption         =   "&Budgets"
      Begin VB.Menu mnuBudMaint 
         Caption         =   "&Maintenance"
         Begin VB.Menu mnuBudMaintSales 
            Caption         =   "&Sales Forecast"
         End
         Begin VB.Menu mnuLabBudMaint 
            Caption         =   "&Labor Budget Maintenance"
         End
         Begin VB.Menu mnuPurchBudMaint 
            Caption         =   "&Purchases Budget Maintenance"
         End
      End
      Begin VB.Menu mnuBudRpts 
         Caption         =   "&Reports"
      End
   End
   Begin VB.Menu mnuRpts 
      Caption         =   "&MgmtRpts"
      Begin VB.Menu mnuRptsDailyRpt 
         Caption         =   "&Daily Report"
      End
      Begin VB.Menu mnuRptsDtlSalesAnalysis 
         Caption         =   "Detail &Sales Analysis"
      End
      Begin VB.Menu mnuRptsDtlLaborAnalysis 
         Caption         =   "Detail &Labor Analysis"
      End
      Begin VB.Menu mnuRptsDtlPurchAnalysis 
         Caption         =   "Detail &Purchases Analysis"
      End
      Begin VB.Menu mnuRptsCompare 
         Caption         =   "&Comparative Report"
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuUtilAloha 
         Caption         =   "&Aloha Transfer"
      End
      Begin VB.Menu mnuUtilClearData 
         Caption         =   "&Clear Data From Tables"
      End
      Begin VB.Menu mnuUtilRefList 
         Caption         =   "Missing &Reference List"
      End
      Begin VB.Menu mnuUtilViewClipBd 
         Caption         =   "&View ClipBoard"
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuBudgetsAnnual_Click()
   frmAnnBudMaint.Show nonmodal
End Sub

Private Sub mnuFileExit_Click()
   Call basMain.EndProgram
   End
End Sub

Private Sub mnuEmpMaintEmp_Click()
   frmEmpMaint.Show nonmodal
End Sub

Private Sub mnuEmpMaintPayRate_Click()
   frmPayRateMaint.Show nonmodal
End Sub

Private Sub mnuLabBudMaint_Click()
   frmLabBudgetMaint.Show nonmodal
End Sub

Private Sub mnuLabMaint_Click()
   frmLaborMaint.Show nonmodal
End Sub

Private Sub mnuPurchBudMaint_Click()
   frmPurchBudgetMaint.Show nonmodal
End Sub

Private Sub mnuPurMaint_Click()
  frmPurchMaint.Show nonmodal
End Sub

Private Sub mnuSalesMaint_Click()
   frmSalesMaint.Show nonmodal
End Sub

Private Sub mnuStoresMaint_Click()
   frmStoresMaint.Show nonmodal
End Sub

Private Sub mnuStoresShiftsMaint_Click()
    frmShiftsMaint.Show nonmodal
End Sub

Private Sub mnuTaxesMaint_Click()
   frmTaxesMaint.Show nonmodal
End Sub

Private Sub mnuBudMaintSales_Click()
   frmSalesFcstMaint.Show nonmodal
End Sub

Private Sub mnuUtilAloha_Click()
   basPOSTransfer.TransferPOS
End Sub

Private Sub mnuUtilClearData_Click()
   frmClearDBTables.Show nonmodal
End Sub

Private Sub mnuUtilRefList_Click()
   Call basRefList.RefList
End Sub

Private Sub mnuUtilViewClipBd_Click()
   frmClipBdView.Show nonmodal
End Sub

Private Sub mnuVendMaint_Click()
   frmVendorsMaint.Show nonmodal
End Sub

Private Sub mnuRptsDailyRpt_Click()
   frmDailyRpt.Show nonmodal
End Sub

Private Sub mnuRptsDtlSalesAnalysis_Click()
   frmDtlSalesAnalysis.Show nonmodal
End Sub

Private Sub mnuRptsDtlLaborAnalysis_Click()
   frmDtlLaborAnalysis.Show nonmodal
End Sub

Private Sub mnuRptsDtlPurchAnalysis_Click()
   frmDtlPurchAnalysis.Show nonmodal
End Sub

Private Sub mnuRptsCompare_Click()
   frmComparative.Show nonmodal
End Sub

