VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SankeyEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'  Application-level event hook for the Sankey add-in.
'  Redraws a sheet's Sankey chart when its source data changes (auto-update)
'  and keeps the ribbon controls in sync with the active sheet's chart.
'==============================================================================
Option Explicit

Public WithEvents App As Application

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    SankeyAddin.AutoRefresh Sh, Target
End Sub

Private Sub App_SheetCalculate(ByVal Sh As Object)
    SankeyAddin.AutoRefresh Sh, Nothing        ' catches values driven by formulas
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    SankeyAddin.RibbonRefresh
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    SankeyAddin.RibbonRefresh
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    SankeyAddin.ForgetWorkbook Wb.Name         ' drop in-memory settings so a new file starts from defaults
End Sub
