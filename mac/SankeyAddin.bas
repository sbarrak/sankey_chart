Attribute VB_Name = "SankeyAddin"
'==============================================================================
'  SANKEY CHART BUILDER  v2.0
'  Draws professional Sankey (flow) diagrams from a 3-column table:
'       From | To | Value
'  - Chart = native, editable Excel shapes (grouped as "Sankey_Chart")
'  - Auto-updates when source values change (see SankeyEvents class)
'  - Ribbon tab "Sankey" (needs Install-Ribbon.ps1) or fallback toolbar
'  Per-chart settings are stored in hidden sheet-level names (Sankey_*), so they
'  travel with the workbook. Defaults for new charts live in the registry.
'==============================================================================
Option Explicit

Private Const APP_NAME As String = "Sankey Chart Builder"
Private Const APP_VERSION As String = "2.0"
Private Const CHART_NAME As String = "Sankey_Chart"
Private Const PLACEHOLDER_TAG As String = "Sankey empty chart area"
Private Const REG_APP As String = "SankeyAddin"
Private Const FONT_NAME As String = "Arial"
Private Const FONT_SIZE As Double = 10
Private Const NODE_W As Double = 14
Private Const NODE_GAP As Double = 12

' ---- globals shared with the ribbon / event class ------------------------------
Public gRibbon As IRibbonUI
Public gEvents As SankeyEvents
Private sigCache As Object
Private gBusy As Boolean
Private gErrNum As Long, gErrDesc As String
Private gPolling As Boolean, gPollNext As Date
Private gChartSel As Boolean                    ' Sankey chart selected = Sankey tab visible
Private lastGW As Double, lastGH As Double      ' actual size of the group as drawn
Private pend As Object                          ' settings chosen before a chart exists (per sheet, memory only)
Private gPurged As Boolean
Private cPal As String                          ' imported palette as "RRGGBB,RRGGBB,..."
Private cWb As Workbook                         ' workbook of the chart (for its theme colors)
Private gNodeCol As Object                      ' custom node colors: node name -> RGB

' ---- active settings for the chart being drawn ---------------------------------
Private cTitle As String, cSub As String, cUnit As String, cAuto As String
Private cTheme As Long, cLink As Long
Private cW As Double, cH As Double
Private cShowVal As Boolean
Private mL As Double, mR As Double, mT As Double, mB As Double

' ---- model ---------------------------------------------------------------------
Private nName() As String, nLayer() As Long, nIn() As Double, nOut() As Double
Private nVal() As Double, nX() As Double, nY() As Double, nH() As Double
Private nOrd() As Long, nPos() As Double, nCnt As Long
Private lSrc() As Long, lDst() As Long, lVal() As Double
Private lSy() As Double, lTy() As Double, lCnt As Long
Private maxLayer As Long, scaleF As Double
Private drawn() As Variant, drawnCnt As Long

'==============================================================================
'  PUBLIC COMMANDS
'==============================================================================
Public Sub CreateSankey()
    Dim ws As Worksheet, rng As Range
    On Error GoTo EH
    EnsureEvents
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Activate a worksheet that contains your data first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    Set ws = ActiveSheet
    Set rng = ResolveSource(ws, True)
    If Not rng Is Nothing Then
        If BuildSankey(ws, rng, True) Then
            RibbonRefresh
            Exit Sub
        End If
    End If
    ' no usable data: leave an empty chart area to style (or keep the chart that is already there)
    If HasChart(ws) And Not IsPlaceholder(ws) Then
        MsgBox "Select a range with three columns (From, To, Value) with positive numeric values, then click again.", vbExclamation, APP_NAME
        Exit Sub
    End If
    ShowPlaceholder ws
    RibbonRefresh
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Sankey failed (error " & Err.Number & "): " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Function IsPlaceholder(ws As Worksheet) As Boolean
    On Error Resume Next
    IsPlaceholder = (ws.Shapes(CHART_NAME).AlternativeText = PLACEHOLDER_TAG)
End Function

Private Sub ShowPlaceholder(ws As Worksheet)
    Dim s As Shape, x As Double, y As Double
    If IsPlaceholder(ws) Then
        ws.Shapes(CHART_NAME).Select
        Exit Sub
    End If
    LoadConfig ws
    x = ActiveWindow.VisibleRange.Left: y = ActiveWindow.VisibleRange.Top
    If TypeName(Selection) = "Range" Then x = Selection.Left: y = Selection.Top
    Set s = ws.Shapes.AddShape(msoShapeRectangle, x, y, cW, cH)
    s.Name = CHART_NAME
    s.AlternativeText = PLACEHOLDER_TAG
    s.Fill.ForeColor.RGB = RGB(248, 249, 251)
    s.Line.ForeColor.RGB = RGB(150, 160, 175)
    s.Line.DashStyle = msoLineDash
    s.Line.Weight = 1.25
    With s.TextFrame2
        .TextRange.Text = "Sankey chart area" & vbLf & "Select your table (From | To | Value) and click the Sankey icon again."
        .TextRange.Font.Size = 14
        .TextRange.Font.Fill.ForeColor.RGB = RGB(90, 100, 115)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With
    s.Select
End Sub

Public Sub DeleteSankey(Optional ws As Worksheet)
    Dim i As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    DeleteChartShape ws
    ClearPending ws
    On Error Resume Next
    For i = ws.Names.Count To 1 Step -1
        If ws.Names(i).Name Like "*!Sankey_*" Then ws.Names(i).Delete
    Next i
End Sub

Public Sub DeleteSankeyActive()
    If TypeName(ActiveSheet) = "Worksheet" Then DeleteSankey ActiveSheet
    RibbonRefresh
End Sub

Public Sub InsertSampleData()
    Dim wb As Workbook, ws As Worksheet, d As Variant, i As Long, nm As String, k As Long
    On Error GoTo EH
    EnsureEvents
    If Workbooks.Count = 0 Then Workbooks.Add
    Set wb = ActiveWorkbook
    nm = "Sankey Demo": k = 1
    Do While SheetExists(wb, nm)
        k = k + 1: nm = "Sankey Demo " & k
    Loop
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = nm
    d = Array("Salary", "Budget", 5000, "Freelance", "Budget", 1200, "Investments", "Budget", 800, _
              "Budget", "Housing", 2400, "Budget", "Food", 1100, "Budget", "Transport", 600, _
              "Budget", "Savings", 1800, "Budget", "Leisure", 1100, "Savings", "Emergency Fund", 600, _
              "Savings", "Retirement", 1200, "Housing", "Rent", 2000, "Housing", "Utilities", 400)
    ws.Range("A1:C1").Value = Array("From", "To", "Value")
    For i = 0 To 11
        ws.Cells(i + 2, 1).Value = d(i * 3)
        ws.Cells(i + 2, 2).Value = d(i * 3 + 1)
        ws.Cells(i + 2, 3).Value = d(i * 3 + 2)
    Next i
    With ws.Range("A1:C1")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(31, 78, 121)
    End With
    ws.Range("C2:C13").NumberFormat = "#,##0"
    ws.Columns("A:C").AutoFit
    ws.Range("A1:C13").Select
    BuildSankey ws, ws.Range("A1:C13"), False, "Monthly Budget Flow", "Where the money comes from and where it goes"
    RibbonRefresh
    Exit Sub
EH:
    MsgBox "Could not create the sample (error " & Err.Number & "): " & Err.Description, vbExclamation, APP_NAME
End Sub

Public Sub QuickTitle()                      ' InputBox alternative to the ribbon edit boxes
    Dim ws As Worksheet, t As String, s As String
    Set ws = CurSheet()
    t = InputBox("Chart title:", APP_NAME, CfgRead(ws, "Title"))
    s = InputBox("Subtitle (optional):", APP_NAME, CfgRead(ws, "Sub"))
    CfgWrite ws, "Title", t
    CfgWrite ws, "Sub", s
    RedrawActive
End Sub

Public Sub ShowHelp()
    MsgBox "HOW TO USE" & vbCrLf & vbCrLf & _
        "1. Put your data in three columns:  From | To | Value  (header row optional)." & vbCrLf & _
        "2. Select the table (or click one cell inside it) and click the Sankey icon on the Add-ins tab." & vbCrLf & _
        "3. The chart appears to the right of the data. Move, resize or recolor it freely." & vbCrLf & _
        "4. Click the chart: the Sankey tab appears with the design options." & vbCrLf & vbCrLf & _
        "AUTO-UPDATE" & vbCrLf & _
        "With Auto-update ticked, editing a value, a name, or adding a row directly below " & _
        "the table redraws the chart. Note: each redraw clears Excel's Undo history." & vbCrLf & vbCrLf & _
        "DESIGN" & vbCrLf & _
        "Title, subtitle, color theme, link style, size and unit are on the ribbon. " & _
        "Changing them restyles the chart on the active sheet." & vbCrLf & vbCrLf & _
        "COLORS" & vbCrLf & _
        "- Colors dropdown: built-in palettes, 'Workbook theme' (the theme applied in Page Layout > Colors, " & _
        "e.g. your company palette) or 'Custom palette'." & vbCrLf & _
        "- Import palette: read colors from selected cells (fill color or hex codes like #E4002B) or from a " & _
        "text/CSV/XML file." & vbCrLf & _
        "- Change node color: click the chart, click a node once more so it is selected alone, then press the button." & vbCrLf & _
        "  Reset node colors returns every node to the palette." & vbCrLf & vbCrLf & _
        "TIPS" & vbCrLf & _
        "- Flows must go one way (A > B > C). Circular flows are rejected." & vbCrLf & _
        "- Use 'Sample data' to see a working example." & vbCrLf & _
        "- Select multiple cells to point the chart at a different table.", _
        vbInformation, APP_NAME & " - Help"
End Sub

Public Sub ShowAbout()
    MsgBox APP_NAME & "  v" & APP_VERSION & vbCrLf & vbCrLf & _
        "Creates Sankey flow diagrams from a From / To / Value table, " & _
        "using native Excel shapes, with automatic updates when the data changes." & vbCrLf & vbCrLf & _
        "Excel for Mac (Microsoft 365).", vbInformation, "About " & APP_NAME
End Sub

'==============================================================================
'  AUTO-UPDATE (called by the SankeyEvents class)
'==============================================================================
Public Sub EnsureEvents()
    If gEvents Is Nothing Then
        Set gEvents = New SankeyEvents
        Set gEvents.App = Application
    End If
    StartPolling
    If Not gPurged Then                                    ' remove defaults left in the registry by older versions
        gPurged = True
        On Error Resume Next
        DeleteSetting REG_APP
    End If
End Sub

'--- resize watcher: Excel has no "shape resized" event, so a 1-second timer compares ----
'--- the chart's size with the size we last drew; a difference = user resized it.  ----
Private Sub StartPolling()
    If gPolling Then Exit Sub
    gPolling = True
    SchedulePoll
End Sub

Private Sub SchedulePoll()
    On Error Resume Next
    gPollNext = Now + TimeSerial(0, 0, 1)
    Application.OnTime gPollNext, "'" & ThisWorkbook.Name & "'!PollResize"
End Sub

Public Sub PollResize()
    If Not gPolling Then Exit Sub
    CheckSelection
    CheckResize
    SchedulePoll
End Sub

'--- Mac stand-in for a contextual tab: show the Sankey tab only while the chart is selected ----
Private Sub CheckSelection()
    Dim sel As Boolean
    sel = SankeySelected()
    If sel = gChartSel Then Exit Sub
    gChartSel = sel
    On Error Resume Next
    If gRibbon Is Nothing Then Exit Sub
    gRibbon.Invalidate                                             ' re-reads visibility and the control values
    If sel Then gRibbon.ActivateTab "tabSankey"
End Sub

Private Function SankeySelected() As Boolean
    Dim shp As Object
    On Error GoTo Done
    If Workbooks.Count = 0 Then Exit Function
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function
    If TypeName(Selection) = "Range" Then Exit Function
    Set shp = Selection.ShapeRange(1)
    If shp.Name = CHART_NAME Then SankeySelected = True: Exit Function
    If shp.ParentGroup.Name = CHART_NAME Then SankeySelected = True   ' a node or flow inside the chart
Done:
End Function

Public Sub Rib_GetVisible(control As IRibbonControl, ByRef visible)
    visible = gChartSel
End Sub

Private Sub CheckResize()
    Dim ws As Worksheet, s As Shape, found As Boolean
    Dim gw As Double, gh As Double, oldW As Double, oldH As Double, nw As Double, nh As Double
    On Error GoTo Quiet
    If gBusy Then Exit Sub
    If Workbooks.Count = 0 Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet
    If ws.ProtectContents Then Exit Sub
    For Each s In ws.Shapes
        If s.Name = CHART_NAME Then found = True: Exit For
    Next s
    If Not found Then Exit Sub
    gw = Val(NameStr(ws, "Sankey_GW", "0"))
    gh = Val(NameStr(ws, "Sankey_GH", "0"))
    If gw <= 0 Or gh <= 0 Then Exit Sub                            ' size unknown (older chart): nothing to compare
    If Abs(s.Width - gw) <= 1.5 And Abs(s.Height - gh) <= 1.5 Then Exit Sub

    oldW = Val(NameStr(ws, "Sankey_W", "960")): oldH = Val(NameStr(ws, "Sankey_H", "560"))
    nw = ClampD(oldW * s.Width / gw, 400, 3000)                    ' keep proportion to the drawn card
    nh = ClampD(oldH * s.Height / gh, 250, 2500)
    CfgWrite ws, "W", CStr(CLng(nw))
    CfgWrite ws, "H", CStr(CLng(nh))
    RedrawActive                                                   ' re-lay out at the new size (crisp text)
    RibbonRefresh                                                  ' show the new numbers in the ribbon boxes
Quiet:
End Sub

Public Sub AutoRefresh(ByVal Sh As Object, ByVal Target As Range)
    Dim ws As Worksheet, src As Range, area As Range, sig As String
    On Error GoTo Quiet
    If gBusy Then Exit Sub                                 ' ignore events caused by our own drawing
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    Set ws = Sh
    If Not HasChart(ws) Then Exit Sub
    If NameStr(ws, "Sankey_Auto", "1") <> "1" Then Exit Sub
    If ws.ProtectContents Then Exit Sub
    Set src = StoredSource(ws)
    If src Is Nothing Then Exit Sub
    If Not Target Is Nothing Then
        Set area = src.Resize(src.Rows.Count + 1)          ' include the row just below (appended data)
        If Intersect(Target, area) Is Nothing Then Exit Sub
    End If
    Set src = ExtendRange(src)
    If sigCache Is Nothing Then Set sigCache = New SkDict
    sig = Signature(src)
    If sigCache.Exists(SigKey(ws)) Then
        If sigCache(SigKey(ws)) = sig Then Exit Sub        ' nothing relevant changed
    End If
    BuildSankey ws, src, True
Quiet:
End Sub

Public Sub RibbonRefresh()
    On Error Resume Next
    If Not gRibbon Is Nothing Then gRibbon.Invalidate
End Sub

Private Function SigKey(ws As Worksheet) As String
    SigKey = ws.Parent.Name & "|" & ws.Name
End Function

Private Function Signature(rng As Range) As String
    Dim v As Variant, r As Long, c As Long, parts() As String, n As Long
    v = rng.Value
    ReDim parts(0 To UBound(v, 1) * UBound(v, 2))
    parts(0) = rng.Address
    For r = 1 To UBound(v, 1)
        For c = 1 To UBound(v, 2)
            n = n + 1
            If IsError(v(r, c)) Then parts(n) = "#err" Else parts(n) = CStr(v(r, c))
        Next c
    Next r
    Signature = Join(parts, Chr(1))
End Function

Private Function ExtendRange(src As Range) As Range
    Dim ws As Worksheet, lastRow As Long
    Set ws = src.Worksheet
    lastRow = src.Row + src.Rows.Count - 1
    Do While lastRow < ws.Rows.Count
        If Len(Trim$(CStr(ws.Cells(lastRow + 1, src.Column).Value))) = 0 And _
           Len(Trim$(CStr(ws.Cells(lastRow + 1, src.Column + 1).Value))) = 0 Then Exit Do
        lastRow = lastRow + 1
    Loop
    Set ExtendRange = ws.Range(src.Cells(1, 1), ws.Cells(lastRow, src.Column + 2))
End Function

'==============================================================================
'  SETTINGS  (sheet-level hidden names when a chart exists, registry otherwise)
'==============================================================================
Private Function CurSheet() As Worksheet
    On Error Resume Next
    If TypeName(ActiveSheet) = "Worksheet" Then Set CurSheet = ActiveSheet
End Function

Private Function DefaultOf(key As String) As String
    Select Case key
        Case "Theme": DefaultOf = "1"
        Case "Link": DefaultOf = "2"
        Case "W": DefaultOf = "960"
        Case "H": DefaultOf = "560"
        Case "ShowVal", "Auto": DefaultOf = "1"
        Case Else: DefaultOf = ""
    End Select
End Function

' Settings never leave the file: with a chart they live in that sheet's hidden names (saved with the
' workbook); before a chart exists they are held in memory for that sheet only. New files = defaults.
Private Function PendKey(ws As Worksheet, ByVal key As String) As String
    PendKey = SigKey(ws) & "|" & key
End Function

Private Function CfgRead(ws As Worksheet, key As String) As String
    Dim s As String
    If Not ws Is Nothing Then
        If HasChart(ws) Then
            s = NameStr(ws, "Sankey_" & key, Chr(1))
            If s <> Chr(1) Then CfgRead = s: Exit Function
        End If
        If Not pend Is Nothing Then
            If pend.Exists(PendKey(ws, key)) Then CfgRead = pend(PendKey(ws, key)): Exit Function
        End If
    End If
    CfgRead = DefaultOf(key)
End Function

Private Sub CfgWrite(ws As Worksheet, key As String, val As String)
    If ws Is Nothing Then Exit Sub
    If pend Is Nothing Then Set pend = New SkDict
    If HasChart(ws) Then
        SetNameStr ws, "Sankey_" & key, val
    Else
        pend(PendKey(ws, key)) = val
    End If
End Sub

Private Sub ClearPending(ws As Worksheet)
    Dim k As Variant, pre As String
    If pend Is Nothing Then Exit Sub
    pre = SigKey(ws) & "|"
    For Each k In pend.Keys
        If Left$(CStr(k), Len(pre)) = pre Then pend.Remove k
    Next k
End Sub

Public Sub ForgetWorkbook(ByVal wbName As String)          ' called when a workbook closes
    Dim k As Variant, pre As String
    pre = wbName & "|"
    On Error Resume Next
    If Not pend Is Nothing Then
        For Each k In pend.Keys
            If Left$(CStr(k), Len(pre)) = pre Then pend.Remove k
        Next k
    End If
    If Not sigCache Is Nothing Then
        For Each k In sigCache.Keys
            If Left$(CStr(k), Len(pre)) = pre Then sigCache.Remove k
        Next k
    End If
End Sub

Private Sub LoadConfig(ws As Worksheet)
    cTitle = CfgRead(ws, "Title")
    cSub = CfgRead(ws, "Sub")
    cUnit = CfgRead(ws, "Unit")
    cAuto = CfgRead(ws, "Auto")
    cPal = CfgRead(ws, "Pal")
    Set cWb = ws.Parent
    Set gNodeCol = LoadNodeColors(ws)
    cTheme = ClampL(Val(CfgRead(ws, "Theme")), 1, 7)
    cLink = ClampL(Val(CfgRead(ws, "Link")), 1, 3)
    cW = ClampD(Val(CfgRead(ws, "W")), 400, 3000)
    cH = ClampD(Val(CfgRead(ws, "H")), 250, 2500)
    cShowVal = (CfgRead(ws, "ShowVal") = "1")
End Sub

Private Sub PersistConfig(ws As Worksheet, src As Range)
    SetNameStr ws, "Sankey_Title", cTitle
    SetNameStr ws, "Sankey_Sub", cSub
    SetNameStr ws, "Sankey_Unit", cUnit
    SetNameStr ws, "Sankey_Auto", cAuto
    SetNameStr ws, "Sankey_Theme", CStr(cTheme)
    SetNameStr ws, "Sankey_Pal", cPal
    SetNameStr ws, "Sankey_Link", CStr(cLink)
    SetNameStr ws, "Sankey_W", CStr(cW)
    SetNameStr ws, "Sankey_H", CStr(cH)
    SetNameStr ws, "Sankey_GW", CStr(lastGW)
    SetNameStr ws, "Sankey_GH", CStr(lastGH)
    SetNameStr ws, "Sankey_ShowVal", IIf(cShowVal, "1", "0")
    ws.Names.Add Name:="Sankey_Src", RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!" & src.Address, Visible:=False
End Sub

Private Function NameStr(ws As Worksheet, nm As String, def As String) As String
    Dim s As String
    On Error GoTo Missing
    s = ws.Names(nm).RefersTo
    If Left$(s, 2) = "=""" And Right$(s, 1) = """" Then
        NameStr = Replace(Mid$(s, 3, Len(s) - 3), """""", """")
    Else
        NameStr = def
    End If
    Exit Function
Missing:
    NameStr = def
End Function

Private Sub SetNameStr(ws As Worksheet, nm As String, val As String)
    Dim v As String
    v = Replace(Left$(val, 240), """", """""")
    ws.Names.Add Name:=nm, RefersTo:="=""" & v & """", Visible:=False
End Sub

Private Function StoredSource(ws As Worksheet) As Range
    On Error Resume Next
    Set StoredSource = ws.Names("Sankey_Src").RefersToRange
End Function

Private Function HasChart(ws As Worksheet) As Boolean
    Dim s As Shape
    For Each s In ws.Shapes
        If s.Name = CHART_NAME Then HasChart = True: Exit Function
    Next s
End Function

Private Sub DeleteChartShape(ws As Worksheet)
    Dim s As Shape
    For Each s In ws.Shapes
        If s.Name = CHART_NAME Then s.Delete: Exit For
    Next s
End Sub

Private Function ClampL(ByVal v As Double, ByVal lo As Long, ByVal hi As Long) As Long
    If v < lo Then v = lo
    If v > hi Then v = hi
    ClampL = CLng(v)
End Function

Private Function ClampD(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Double
    If v < lo Then v = lo
    If v > hi Then v = hi
    ClampD = v
End Function

Private Function SheetExists(wb As Workbook, nm As String) As Boolean
    Dim s As Object
    For Each s In wb.Sheets
        If StrComp(s.Name, nm, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next s
End Function

'==============================================================================
'  SOURCE RANGE
'==============================================================================
Private Function ResolveSource(ws As Worksheet, Optional silent As Boolean) As Range
    Dim sel As Range, stored As Range, r As Range
    Set stored = StoredSource(ws)
    If TypeName(Selection) = "Range" Then
        Set sel = Selection.Areas(1)
        If sel.CountLarge > 1 Then
            Set r = sel                                       ' explicit selection = (new) source
        ElseIf Not stored Is Nothing And HasChart(ws) Then
            If Intersect(sel, stored) Is Nothing Then Set r = sel.CurrentRegion Else Set r = ExtendRange(stored)
        Else
            Set r = sel.CurrentRegion
        End If
    ElseIf Not stored Is Nothing And HasChart(ws) Then
        Set r = ExtendRange(stored)                           ' chart selected -> refresh it
    Else
        If Not silent Then MsgBox "Select a range with three columns: From, To, Value.", vbExclamation, APP_NAME
        Exit Function
    End If
    Set ResolveSource = NormalizeSource(ws, r, silent)
End Function

Private Function NormalizeSource(ws As Worksheet, r As Range, Optional silent As Boolean) As Range
    Set r = Intersect(r, ws.UsedRange)                        ' trims whole-column selections
    If r Is Nothing Then
        If Not silent Then MsgBox "The selection is empty.", vbExclamation, APP_NAME
        Exit Function
    End If
    If r.Columns.Count < 3 Then
        If Not silent Then MsgBox "Select a range with three columns: From, To, Value.", vbExclamation, APP_NAME
        Exit Function
    End If
    If r.Columns.Count > 3 Then Set r = r.Resize(, 3)         ' first three columns
    Set NormalizeSource = r
End Function

'==============================================================================
'  BUILD
'==============================================================================
Private Function BuildSankey(ws As Worksheet, rng As Range, silent As Boolean, _
                             Optional ByVal tTitle As String = "", _
                             Optional ByVal tSub As String = "") As Boolean
    ' re-entrancy guard: drawing/saving names can fire Change/Calculate events that would call us again
    If gBusy Then Exit Function
    gBusy = True
    BuildSankey = BuildSankeyCore(ws, rng, silent, tTitle, tSub)
    gBusy = False
End Function

Private Function BuildSankeyCore(ws As Worksheet, rng As Range, silent As Boolean, _
                                 ByVal tTitle As String, ByVal tSub As String) As Boolean
    Dim ox As Double, oy As Double, s As Shape
    On Error GoTo EH
    LoadConfig ws
    If Len(tTitle) > 0 Then cTitle = tTitle
    If Len(tSub) > 0 Then cSub = tSub

    LoadData rng
    If lCnt = 0 Then
        If Not silent Then MsgBox "No valid rows found. Each row needs a From, a To and a positive numeric Value.", vbExclamation, APP_NAME
        Exit Function
    End If
    If Not AssignLayers() Then
        If Not silent Then MsgBox "The data contains a cycle (e.g. A > B > A). Sankey charts need flows that go one way.", vbExclamation, APP_NAME
        Exit Function
    End If

    Application.ScreenUpdating = False
    OrderNodes
    PrepareMargins
    Layout

    ox = rng.Left + rng.Width + 30: oy = rng.Top              ' default position; keep it if chart exists
    For Each s In ws.Shapes
        If s.Name = CHART_NAME Then ox = s.Left: oy = s.Top: Exit For
    Next s
    DeleteChartShape ws
    Draw ws, ox, oy
    If sigCache Is Nothing Then Set sigCache = New SkDict
    sigCache(SigKey(ws)) = Signature(rng)                     ' before PersistConfig: names can fire events
    PersistConfig ws, rng

    Application.ScreenUpdating = True
    BuildSankeyCore = True
    Exit Function
EH:
    gErrNum = Err.Number
    gErrDesc = Err.Description
    CleanupDrawn ws
    Application.ScreenUpdating = True
    If Not silent Then
        MsgBox "Sankey failed (error " & gErrNum & "): " & gErrDesc, vbExclamation, APP_NAME
    End If
End Function

Private Sub RedrawActive()
    Dim ws As Worksheet, src As Range
    Set ws = CurSheet()
    If ws Is Nothing Then Exit Sub
    If Not HasChart(ws) Then Exit Sub
    If IsPlaceholder(ws) Then
        LoadConfig ws
        ws.Shapes(CHART_NAME).Width = cW: ws.Shapes(CHART_NAME).Height = cH
        FocusChart ws
        Exit Sub
    End If
    Set src = StoredSource(ws)
    If src Is Nothing Then Exit Sub
    If BuildSankey(ws, ExtendRange(src), True) Then FocusChart ws
End Sub

Private Sub FocusChart(ws As Worksheet)          ' keep the redrawn chart selected and the Sankey tab active
    On Error Resume Next
    If Not ws Is ActiveSheet Then Exit Sub
    ws.Shapes(CHART_NAME).Select
    If gRibbon Is Nothing Then Exit Sub
    If Not gChartSel Then gChartSel = True: gRibbon.Invalidate
    gRibbon.ActivateTab "tabSankey"
End Sub

Private Sub CleanupDrawn(ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = 1 To drawnCnt
        ws.Shapes(CStr(drawn(i))).Delete
    Next i
    drawnCnt = 0
End Sub

'--- read the table -------------------------------------------------------------
Private Sub LoadData(rng As Range)
    Dim v As Variant, r As Long, first As Long, rows As Long
    Dim d As Object, a As String, b As String, x As Double
    v = rng.Value
    nCnt = 0: lCnt = 0
    rows = UBound(v, 1)
    first = 1
    If IsError(v(1, 3)) Then
        first = 2
    ElseIf Not IsNumeric(v(1, 3)) Or IsEmpty(v(1, 3)) Then
        first = 2                                                   ' header row
    End If
    If first > rows Then Exit Sub

    ReDim nName(1 To 2 * rows): ReDim lSrc(1 To rows): ReDim lDst(1 To rows): ReDim lVal(1 To rows)
    Set d = New SkDict
    d.CompareMode = 1                                               ' "a" and "A" are the same node
    For r = first To rows
        If IsError(v(r, 1)) Or IsError(v(r, 2)) Or IsError(v(r, 3)) Then GoTo NextRow
        a = Trim$(CStr(v(r, 1))): b = Trim$(CStr(v(r, 2)))
        If Len(a) > 0 And Len(b) > 0 And a <> b And IsNumeric(v(r, 3)) And Not IsEmpty(v(r, 3)) Then
            x = CDbl(v(r, 3))
            If x > 0 And x < 1E+100 Then
                lCnt = lCnt + 1
                lSrc(lCnt) = NodeIndex(d, a)
                lDst(lCnt) = NodeIndex(d, b)
                lVal(lCnt) = x
            End If
        End If
NextRow:
    Next r
    If lCnt = 0 Then Exit Sub

    ReDim nLayer(1 To nCnt): ReDim nIn(1 To nCnt): ReDim nOut(1 To nCnt): ReDim nVal(1 To nCnt)
    ReDim nX(1 To nCnt): ReDim nY(1 To nCnt): ReDim nH(1 To nCnt): ReDim nOrd(1 To nCnt): ReDim nPos(1 To nCnt)
    ReDim lSy(1 To lCnt): ReDim lTy(1 To lCnt)
    For r = 1 To lCnt
        nOut(lSrc(r)) = nOut(lSrc(r)) + lVal(r)
        nIn(lDst(r)) = nIn(lDst(r)) + lVal(r)
    Next r
    For r = 1 To nCnt
        nVal(r) = IIf(nIn(r) > nOut(r), nIn(r), nOut(r))
    Next r
End Sub

Private Function NodeIndex(d As Object, nm As String) As Long
    If Not d.Exists(nm) Then
        nCnt = nCnt + 1
        nName(nCnt) = nm
        d.Add nm, nCnt
    End If
    NodeIndex = d(nm)
End Function

'--- columns (longest path layering) ------------------------------------------------
Private Function AssignLayers() As Boolean
    Dim pass As Long, i As Long, k As Long, changed As Boolean, m As Long
    For pass = 0 To nCnt
        changed = False
        For i = 1 To lCnt
            If nLayer(lDst(i)) < nLayer(lSrc(i)) + 1 Then
                nLayer(lDst(i)) = nLayer(lSrc(i)) + 1
                changed = True
            End If
        Next i
        If Not changed Then Exit For
    Next pass
    If changed Then Exit Function               ' cycle

    For i = 1 To nCnt                           ' pure sources: slide right next to what they feed
        If nIn(i) = 0 Then
            m = 1000000
            For k = 1 To lCnt
                If lSrc(k) = i And nLayer(lDst(k)) < m Then m = nLayer(lDst(k))
            Next k
            If m < 1000000 Then nLayer(i) = m - 1
        End If
    Next i
    maxLayer = 0
    For i = 1 To nCnt
        If nLayer(i) > maxLayer Then maxLayer = nLayer(i)
    Next i
    AssignLayers = True
End Function

'--- vertical order within columns (barycenter sweeps to cut crossings) -------------
Private Sub OrderNodes()
    Dim L As Long, i As Long, c As Long, sweep As Long
    For L = 0 To maxLayer
        c = 0
        For i = 1 To nCnt
            If nLayer(i) = L Then c = c + 1: nOrd(i) = c
        Next i
        For i = 1 To nCnt
            If nLayer(i) = L Then nPos(i) = (nOrd(i) - 0.5) / c
        Next i
    Next L
    For sweep = 1 To 6
        For L = 1 To maxLayer: SortLayer L, True: Next L
        For L = maxLayer - 1 To 0 Step -1: SortLayer L, False: Next L
    Next sweep
End Sub

Private Sub SortLayer(L As Long, useIncoming As Boolean)
    Dim idx() As Long, key() As Double, n As Long, i As Long, j As Long, k As Long
    Dim sw As Double, sv As Double, t As Long, tk As Double
    For i = 1 To nCnt
        If nLayer(i) = L Then n = n + 1
    Next i
    If n < 2 Then Exit Sub
    ReDim idx(1 To n): ReDim key(1 To n)
    n = 0
    For i = 1 To nCnt
        If nLayer(i) = L Then n = n + 1: idx(n) = i
    Next i
    For i = 1 To n
        sw = 0: sv = 0
        For k = 1 To lCnt
            If useIncoming Then
                If lDst(k) = idx(i) Then sw = sw + lVal(k) * nPos(lSrc(k)): sv = sv + lVal(k)
            Else
                If lSrc(k) = idx(i) Then sw = sw + lVal(k) * nPos(lDst(k)): sv = sv + lVal(k)
            End If
        Next k
        If sv > 0 Then key(i) = sw / sv Else key(i) = nPos(idx(i))   ' (not IIf: 0/0 would overflow)
    Next i
    For i = 2 To n                                   ' stable insertion sort
        t = idx(i): tk = key(i): j = i - 1
        Do While j >= 1
            If key(j) <= tk Then Exit Do
            idx(j + 1) = idx(j): key(j + 1) = key(j): j = j - 1
        Loop
        idx(j + 1) = t: key(j + 1) = tk
    Next i
    For i = 1 To n
        nOrd(idx(i)) = i
        nPos(idx(i)) = (i - 0.5) / n
    Next i
End Sub

'--- coordinates -------------------------------------------------------------------
Private Function FmtVal(ByVal v As Double) As String
    If Abs(v - Fix(v)) < 0.000001 Then
        FmtVal = Format$(v, "#,##0")
    Else
        FmtVal = Format$(v, "#,##0.0#")
    End If
    If Len(cUnit) > 0 Then FmtVal = FmtVal & " " & cUnit
End Function

Private Function LabelText(ByVal i As Long) As String
    LabelText = nName(i)
    If cShowVal Then LabelText = LabelText & "  " & FmtVal(nVal(i))
End Function

Private Sub PrepareMargins()
    Dim i As Long, w As Double, maxL As Double, maxR As Double
    For i = 1 To nCnt
        w = Len(LabelText(i)) * FONT_SIZE * 0.6 + 14
        If nLayer(i) = 0 Then
            If w > maxL Then maxL = w
        ElseIf nLayer(i) = maxLayer Then
            If w > maxR Then maxR = w
        End If
    Next i
    mL = ClampD(maxL, 40, cW * 0.3) + 16
    mR = ClampD(maxR, 40, cW * 0.3) + 16
    mT = 24
    If Len(cTitle) > 0 Then mT = mT + 28
    If Len(cSub) > 0 Then mT = mT + 20
    mB = 38
End Sub

Private Sub Layout()
    Dim L As Long, i As Long, o As Long, cnt As Long, sm As Double, s As Double
    Dim usableH As Double, usedH As Double, y As Double, stepX As Double

    usableH = cH - mT - mB
    scaleF = 1E+30
    For L = 0 To maxLayer
        cnt = 0: sm = 0
        For i = 1 To nCnt
            If nLayer(i) = L Then cnt = cnt + 1: sm = sm + nVal(i)
        Next i
        If sm > 0 Then
            s = (usableH - NODE_GAP * (cnt - 1)) / sm
            If s < scaleF Then scaleF = s
        End If
    Next L
    If scaleF <= 0 Or scaleF > 1E+29 Then scaleF = 0.01
    If scaleF < 0.000001 Then scaleF = 0.000001

    If maxLayer > 0 Then stepX = (cW - mL - mR - NODE_W) / maxLayer
    For L = 0 To maxLayer
        cnt = 0: sm = 0
        For i = 1 To nCnt
            If nLayer(i) = L Then cnt = cnt + 1: sm = sm + nVal(i)
        Next i
        usedH = sm * scaleF + NODE_GAP * (cnt - 1)
        y = mT + (usableH - usedH) / 2
        For o = 1 To cnt
            For i = 1 To nCnt
                If nLayer(i) = L And nOrd(i) = o Then
                    nX(i) = mL + L * stepX
                    nY(i) = y
                    nH(i) = nVal(i) * scaleF
                    y = y + nH(i) + NODE_GAP
                End If
            Next i
        Next o
    Next L

    For i = 1 To nCnt                               ' ribbon attachment points on each node
        AssignOffsets i, True
        AssignOffsets i, False
    Next i
End Sub

Private Sub AssignOffsets(node As Long, outgoing As Boolean)
    Dim ids() As Long, key() As Double, n As Long, k As Long, i As Long, j As Long
    Dim t As Long, tk As Double, cum As Double
    For k = 1 To lCnt
        If outgoing Then
            If lSrc(k) = node Then n = n + 1
        Else
            If lDst(k) = node Then n = n + 1
        End If
    Next k
    If n = 0 Then Exit Sub
    ReDim ids(1 To n): ReDim key(1 To n)
    n = 0
    For k = 1 To lCnt
        If outgoing Then
            If lSrc(k) = node Then n = n + 1: ids(n) = k: key(n) = nY(lDst(k))
        Else
            If lDst(k) = node Then n = n + 1: ids(n) = k: key(n) = nY(lSrc(k))
        End If
    Next k
    For i = 2 To n
        t = ids(i): tk = key(i): j = i - 1
        Do While j >= 1
            If key(j) <= tk Then Exit Do
            ids(j + 1) = ids(j): key(j + 1) = key(j): j = j - 1
        Loop
        ids(j + 1) = t: key(j + 1) = tk
    Next i
    cum = 0
    For i = 1 To n
        If outgoing Then lSy(ids(i)) = nY(node) + cum Else lTy(ids(i)) = nY(node) + cum
        cum = cum + lVal(ids(i)) * scaleF
    Next i
End Sub

'==============================================================================
'  DRAWING
'==============================================================================
Private Sub Reg(shp As Shape)
    drawnCnt = drawnCnt + 1
    drawn(drawnCnt) = shp.Name
End Sub

Private Function AddText(ws As Worksheet, ByVal txt As String, ByVal x As Double, ByVal y As Double, _
                         ByVal sz As Double, ByVal isBold As Boolean, ByVal clr As Long) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, 10, 10)
    With s.TextFrame2
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeShapeToFitText
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .TextRange.Text = txt
        .TextRange.Font.Name = FONT_NAME
        .TextRange.Font.Size = sz
        .TextRange.Font.Bold = IIf(isBold, msoTrue, msoFalse)
        .TextRange.Font.Fill.ForeColor.RGB = clr
    End With
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoFalse
    Set AddText = s
End Function

Private Sub Draw(ws As Worksheet, ox As Double, oy As Double)
    Dim shp As Shape, fb As FreeformBuilder, i As Long, total As Double
    Dim x0 As Double, x1 As Double, mid As Double, t As Double
    Dim pal As Variant, cA As Long, cB As Long, lbl As String

    pal = ThemeColors(cTheme)
    drawnCnt = 0
    ReDim drawn(1 To 6 + lCnt + 2 * nCnt)

    ' card background
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ox, oy, cW, cH)
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.ForeColor.RGB = RGB(222, 226, 232)
    shp.Line.Weight = 0.75
    shp.Shadow.Visible = msoFalse
    On Error Resume Next
    shp.Adjustments.Item(1) = 0.02
    On Error GoTo 0
    Reg shp

    ' header
    If Len(cTitle) > 0 Then
        Set shp = AddText(ws, cTitle, ox + 24, oy + 16, 16, True, RGB(31, 41, 55)): Reg shp
    End If
    If Len(cSub) > 0 Then
        Set shp = AddText(ws, cSub, ox + 24, oy + 16 + IIf(Len(cTitle) > 0, 28, 0), 10, False, RGB(107, 114, 128)): Reg shp
    End If

    ' flows
    For i = 1 To lCnt
        x0 = ox + nX(lSrc(i)) + NODE_W: x1 = ox + nX(lDst(i)): mid = (x0 + x1) / 2
        t = lVal(i) * scaleF
        cA = NodeColor(pal, lSrc(i)): cB = NodeColor(pal, lDst(i))
        Set fb = ws.Shapes.BuildFreeform(msoEditingCorner, x0, oy + lSy(i))
        fb.AddNodes msoSegmentCurve, msoEditingCorner, mid, oy + lSy(i), mid, oy + lTy(i), x1, oy + lTy(i)
        fb.AddNodes msoSegmentLine, msoEditingAuto, x1, oy + lTy(i) + t
        fb.AddNodes msoSegmentCurve, msoEditingCorner, mid, oy + lTy(i) + t, mid, oy + lSy(i) + t, x0, oy + lSy(i) + t
        fb.AddNodes msoSegmentLine, msoEditingAuto, x0, oy + lSy(i)
        Set shp = fb.ConvertToShape
        shp.Line.Visible = msoFalse
        Select Case cLink
            Case 1                                          ' solid, source colour
                shp.Fill.ForeColor.RGB = cA
                shp.Fill.Transparency = 0.55
            Case 3                                          ' neutral grey
                shp.Fill.ForeColor.RGB = RGB(190, 195, 202)
                shp.Fill.Transparency = 0.35
            Case Else                                       ' gradient source -> target
                shp.Fill.ForeColor.RGB = cA
                shp.Fill.Transparency = 0.5
                On Error Resume Next
                With shp.Fill
                    .TwoColorGradient msoGradientVertical, 1
                    .ForeColor.RGB = cA
                    .BackColor.RGB = cB
                    .GradientAngle = 0
                    .GradientStops(1).Transparency = 0.45
                    .GradientStops(2).Transparency = 0.45
                End With
                On Error GoTo 0
        End Select
        shp.AlternativeText = nName(lSrc(i)) & " " & ChrW(8594) & " " & nName(lDst(i)) & ": " & FmtVal(lVal(i)) & _
                              " (" & Format$(lVal(i) / nOut(lSrc(i)), "0.0%") & " of " & nName(lSrc(i)) & ")"
        SetTitle shp, "SankeyLink|" & nName(lSrc(i))
        Reg shp
    Next i

    ' nodes + labels
    For i = 1 To nCnt
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, ox + nX(i), oy + nY(i), NODE_W, nH(i))
        shp.Fill.ForeColor.RGB = NodeColor(pal, i)
        shp.Line.Visible = msoFalse
        shp.Shadow.Visible = msoFalse
        shp.AlternativeText = nName(i) & ": " & FmtVal(nVal(i))
        SetTitle shp, "SankeyNode|" & nName(i)
        Reg shp

        lbl = LabelText(i)
        Set shp = AddText(ws, lbl, 0, 0, FONT_SIZE, False, RGB(55, 65, 81))
        With shp.TextFrame2.TextRange
            .Characters(1, Len(nName(i))).Font.Bold = msoTrue
            If cShowVal Then .Characters(Len(nName(i)) + 1, Len(lbl) - Len(nName(i))).Font.Fill.ForeColor.RGB = RGB(107, 114, 128)
        End With
        If nLayer(i) = 0 Then
            shp.Left = ox + nX(i) - 6 - shp.Width
        Else
            If nLayer(i) < maxLayer Then                    ' middle columns sit on top of flows: soft backing
                shp.TextFrame2.MarginLeft = 3: shp.TextFrame2.MarginRight = 3
                shp.TextFrame2.MarginTop = 1: shp.TextFrame2.MarginBottom = 1
                shp.Fill.Visible = msoTrue
                shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
                shp.Fill.Transparency = 0.2
            End If
            shp.Left = ox + nX(i) + NODE_W + 4
        End If
        shp.Top = oy + nY(i) + nH(i) / 2 - shp.Height / 2
        SetTitle shp, "SankeyLabel|" & nName(i)
        Reg shp
    Next i

    ' footer
    For i = 1 To nCnt
        If nIn(i) = 0 Then total = total + nOut(i)
    Next i
    Set shp = AddText(ws, "Total flow: " & FmtVal(total) & "     |     " & nCnt & " nodes, " & lCnt & " flows", _
                      ox + 24, oy + cH - 24, 8, False, RGB(156, 163, 175))
    Reg shp

    ReDim Preserve drawn(1 To drawnCnt)
    Set shp = ws.Shapes.Range(drawn).Group
    shp.Name = CHART_NAME
    shp.AlternativeText = "Sankey diagram: " & nCnt & " nodes, " & lCnt & " flows"
    shp.Placement = xlFreeFloating
    lastGW = shp.Width: lastGH = shp.Height
    drawnCnt = 0
End Sub

Private Function ThemeColors(ByVal t As Long) As Variant
    Select Case t
        Case 2      ' Vibrant
            ThemeColors = Array(RGB(78, 121, 167), RGB(242, 142, 43), RGB(225, 87, 89), RGB(118, 183, 178), _
                                RGB(89, 161, 79), RGB(237, 201, 72), RGB(176, 122, 161), RGB(255, 157, 167), _
                                RGB(156, 117, 95), RGB(186, 176, 172))
        Case 3      ' Pastel
            ThemeColors = Array(RGB(140, 170, 190), RGB(240, 150, 160), RGB(150, 200, 235), RGB(245, 200, 150), _
                                RGB(150, 225, 175), RGB(205, 160, 215), RGB(235, 225, 140), RGB(175, 205, 140), _
                                RGB(230, 165, 165), RGB(190, 190, 180))
        Case 4      ' Earth
            ThemeColors = Array(RGB(140, 86, 50), RGB(191, 144, 0), RGB(112, 128, 60), RGB(160, 82, 45), _
                                RGB(94, 114, 121), RGB(205, 133, 63), RGB(85, 107, 47), RGB(139, 115, 85), _
                                RGB(178, 102, 70), RGB(120, 120, 100))
        Case 5      ' Ocean
            ThemeColors = Array(RGB(0, 95, 115), RGB(10, 147, 150), RGB(110, 190, 170), RGB(210, 190, 120), _
                                RGB(238, 155, 0), RGB(202, 103, 2), RGB(187, 62, 3), RGB(174, 32, 18), _
                                RGB(155, 34, 38), RGB(0, 60, 80))
        Case 6      ' colors of the workbook's own theme (Page Layout > Colors)
            ThemeColors = WorkbookPalette()
        Case 7      ' imported / custom palette
            ThemeColors = CustomPalette()
        Case Else   ' 1 = Corporate
            ThemeColors = Array(RGB(31, 78, 121), RGB(46, 117, 182), RGB(0, 150, 160), RGB(112, 173, 71), _
                                RGB(237, 125, 49), RGB(255, 192, 0), RGB(91, 155, 213), RGB(165, 165, 165), _
                                RGB(68, 84, 106), RGB(192, 80, 77))
    End Select
End Function

'==============================================================================
'  COLORS: palettes, imported palettes, per-node custom colors
'==============================================================================
Private Sub SetTitle(shp As Shape, ByVal t As String)       ' tags shapes so a click can be mapped to a node
    On Error Resume Next
    shp.Title = t
End Sub

Private Function NodeColor(pal As Variant, ByVal i As Long) As Long
    If Not gNodeCol Is Nothing Then
        If gNodeCol.Exists(nName(i)) Then
            NodeColor = CLng(gNodeCol(nName(i)))
            Exit Function
        End If
    End If
    NodeColor = pal((i - 1) Mod (UBound(pal) + 1))
End Function

Private Function WorkbookPalette() As Variant
    Dim arr(0 To 6) As Long, ids As Variant, i As Long
    On Error GoTo Fail
    ids = Array(msoThemeAccent1, msoThemeAccent2, msoThemeAccent3, msoThemeAccent4, _
                msoThemeAccent5, msoThemeAccent6, msoThemeDark2)
    For i = 0 To 6
        arr(i) = cWb.Theme.ThemeColorScheme.Colors(ids(i)).RGB
    Next i
    WorkbookPalette = arr
    Exit Function
Fail:
    WorkbookPalette = ThemeColors(1)
End Function

Private Function CustomPalette() As Variant
    Dim parts() As String, arr() As Long, i As Long, n As Long
    If Len(cPal) = 0 Then
        CustomPalette = ThemeColors(1)
        Exit Function
    End If
    parts = Split(cPal, ",")
    ReDim arr(0 To UBound(parts))
    For i = 0 To UBound(parts)
        If Len(parts(i)) = 6 Then
            arr(n) = HexToRgb(parts(i))
            n = n + 1
        End If
    Next i
    If n = 0 Then
        CustomPalette = ThemeColors(1)
        Exit Function
    End If
    ReDim Preserve arr(0 To n - 1)
    CustomPalette = arr
End Function

Private Function HexToRgb(ByVal h As String) As Long
    HexToRgb = RGB(CLng("&H" & Mid$(h, 1, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Mid$(h, 5, 2)))
End Function

Private Function RgbToHex(ByVal c As Long) As String
    RgbToHex = Right$("0" & Hex$(c And 255), 2) & Right$("0" & Hex$((c \ 256) And 255), 2) & _
               Right$("0" & Hex$((c \ 65536) And 255), 2)
End Function

Private Function IsLightHex(ByVal h As String) As Boolean
    Dim c As Long
    c = HexToRgb(h)
    IsLightHex = (0.299 * (c And 255) + 0.587 * ((c \ 256) And 255) + 0.114 * ((c \ 65536) And 255)) > 225
End Function

Private Sub AddHex(ByRef list As String, ByRef n As Long, ByVal h As String)
    h = UCase$(Trim$(h))
    If Not (h Like "[0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F]") Then Exit Sub
    If IsLightHex(h) Then Exit Sub                           ' near-white is useless as a node color
    If n >= 20 Then Exit Sub
    If InStr(1, "," & list & ",", "," & h & ",") > 0 Then Exit Sub
    If n > 0 Then list = list & ","
    list = list & h
    n = n + 1
End Sub

'--- custom node colors, stored in hidden names Sankey_NC1.. (keyed by node name) -------
Private Function EncName(ByVal s As String) As String
    s = Replace(s, "%", "%25")
    s = Replace(s, ";", "%3B")
    s = Replace(s, "=", "%3D")
    s = Replace(s, """", "%22")
    EncName = s
End Function

Private Function DecName(ByVal s As String) As String
    s = Replace(s, "%22", """")
    s = Replace(s, "%3D", "=")
    s = Replace(s, "%3B", ";")
    s = Replace(s, "%25", "%")
    DecName = s
End Function

Private Function LoadNodeColors(ws As Worksheet) As Object
    Dim d As Object, s As String, k As Long, part As String
    Dim items() As String, i As Long, p As Long
    Set d = New SkDict
    d.CompareMode = 1
    Do
        k = k + 1
        part = NameStr(ws, "Sankey_NC" & k, Chr(1))
        If part = Chr(1) Then Exit Do
        s = s & part
    Loop
    If Len(s) > 0 Then
        items = Split(s, ";")
        For i = LBound(items) To UBound(items)
            p = InStrRev(items(i), "=")
            If p > 1 And Len(items(i)) - p = 6 Then
                d(DecName(Left$(items(i), p - 1))) = HexToRgb(Mid$(items(i), p + 1))
            End If
        Next i
    End If
    Set LoadNodeColors = d
End Function

Private Sub ClearNodeColorNames(ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = ws.Names.Count To 1 Step -1
        If ws.Names(i).Name Like "*!Sankey_NC*" Then ws.Names(i).Delete
    Next i
End Sub

Private Sub SaveNodeColors(ws As Worksheet, d As Object)
    Dim k As Variant, s As String, i As Long, n As Long
    ClearNodeColorNames ws
    For Each k In d.Keys
        s = s & EncName(CStr(k)) & "=" & RgbToHex(CLng(d(k))) & ";"
    Next k
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    i = 1
    Do While i <= Len(s)
        n = n + 1
        SetNameStr ws, "Sankey_NC" & n, Mid$(s, i, 200)
        i = i + 200
    Loop
End Sub

'--- change the color of the selected node ----------------------------------------------
Private Function SelectedNodeName() As String
    Dim t As String, p As Long
    On Error GoTo Done
    If TypeName(Selection) = "Range" Then Exit Function
    t = Selection.ShapeRange(1).Title
    p = InStr(t, "|")
    If Left$(t, 6) = "Sankey" And p > 0 Then SelectedNodeName = Mid$(t, p + 1)
Done:
End Function

Private Function NodeFillNow(ws As Worksheet, ByVal nm As String) As Long
    Dim s As Shape, g As Shape
    NodeFillNow = RGB(46, 117, 182)
    On Error Resume Next
    Set g = ws.Shapes(CHART_NAME)
    For Each s In g.GroupItems
        If s.Title = "SankeyNode|" & nm Then NodeFillNow = s.Fill.ForeColor.RGB: Exit For
    Next s
End Function

Private Function PickColor(ByVal cur As Long, ByRef result As Long) As Boolean
    ' Excel for Mac has no color dialog for macros: ask for a hex code instead.
    Dim s As String
    s = InputBox("Enter the color as a hex code, e.g. #E4002B", APP_NAME, "#" & RgbToHex(cur))
    s = Trim$(Replace(s, "#", ""))
    If Len(s) = 0 Then Exit Function
    If Not IsHex6(s) Then
        MsgBox "That is not a valid hex color. Use 6 characters 0-9 / A-F, e.g. E4002B.", vbExclamation, APP_NAME
        Exit Function
    End If
    result = HexToRgb(s)
    PickColor = True
End Function

Private Function IsHex6(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) <> 6 Then Exit Function
    For i = 1 To 6
        If InStr("0123456789abcdefABCDEF", Mid$(s, i, 1)) = 0 Then Exit Function
    Next i
    IsHex6 = True
End Function

Public Sub ChangeNodeColor()
    Dim ws As Worksheet, nm As String, clr As Long, d As Object
    Set ws = CurSheet()
    If ws Is Nothing Then Exit Sub
    If Not HasChart(ws) Then
        MsgBox "There is no Sankey chart on this sheet.", vbExclamation, APP_NAME
        Exit Sub
    End If
    nm = SelectedNodeName()
    If Len(nm) = 0 Then
        MsgBox "Click the chart, then click a node bar (or its label) once more so it is selected on its own, " & _
               "and press this button again." & vbCrLf & vbCrLf & _
               "Tip: clicking a flow recolors the node it starts from.", vbInformation, APP_NAME
        Exit Sub
    End If
    If Not PickColor(NodeFillNow(ws, nm), clr) Then Exit Sub
    Set d = LoadNodeColors(ws)
    d(nm) = clr
    SaveNodeColors ws, d
    RedrawActive
    RibbonRefresh
End Sub

Public Sub ResetNodeColors()
    Dim ws As Worksheet
    Set ws = CurSheet()
    If ws Is Nothing Then Exit Sub
    If Not HasChart(ws) Then Exit Sub
    ClearNodeColorNames ws
    RedrawActive
End Sub

'--- import a palette --------------------------------------------------------------------
Private Function CellColorHex(c As Range) As String
    Dim t As String, p() As String, v As Variant
    On Error GoTo Done
    v = c.Value
    If Not IsError(v) Then
        t = Replace(Trim$(CStr(v)), "#", "")
        If t Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
            CellColorHex = UCase$(t): Exit Function
        End If
        p = Split(t, ",")
        If UBound(p) = 2 Then
            If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
                If Val(p(0)) >= 0 And Val(p(0)) <= 255 And Val(p(1)) >= 0 And Val(p(1)) <= 255 _
                   And Val(p(2)) >= 0 And Val(p(2)) <= 255 Then
                    CellColorHex = RgbToHex(RGB(Val(p(0)), Val(p(1)), Val(p(2)))): Exit Function
                End If
            End If
        End If
    End If
    If c.Interior.ColorIndex <> xlColorIndexNone Then CellColorHex = RgbToHex(c.Interior.Color)
Done:
End Function

Private Sub ApplyPalette(ws As Worksheet, ByVal hexes As String, ByVal n As Long)
    CfgWrite ws, "Pal", hexes
    CfgWrite ws, "Theme", "7"
    RedrawActive
    RibbonRefresh
    MsgBox n & " colors imported. They are used by the 'Custom palette' option in the Colors list.", vbInformation, APP_NAME
End Sub

Public Sub ImportPaletteFromCells()
    Dim ws As Worksheet, r As Range, c As Range, hexes As String, n As Long, k As Long
    Set ws = CurSheet()
    If ws Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select cells that are filled with your colors, or that contain hex codes (e.g. #E4002B).", vbInformation, APP_NAME
        Exit Sub
    End If
    Set r = Intersect(Selection.Areas(1), ws.UsedRange)
    If r Is Nothing Then
        MsgBox "No colored cells found in the selection.", vbInformation, APP_NAME
        Exit Sub
    End If
    For Each c In r.Cells
        k = k + 1
        If k > 2000 Then Exit For
        AddHex hexes, n, CellColorHex(c)
        If n >= 20 Then Exit For
    Next c
    If n = 0 Then
        MsgBox "No usable colors found. Use cells with a fill color, or cells containing hex codes like #E4002B " & _
               "(white and near-white are ignored).", vbInformation, APP_NAME
        Exit Sub
    End If
    ApplyPalette ws, hexes, n
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    Dim fn As Integer, s As String
    fn = FreeFile
    Open path For Binary Access Read As #fn
    If LOF(fn) > 0 Then
        If LOF(fn) > 1000000 Then s = Space$(1000000) Else s = Space$(LOF(fn))
        Get #fn, , s
    End If
    Close #fn
    ReadTextFile = Replace(s, vbNullChar, "")                ' also copes with UTF-16 files
End Function

Private Function IsAlnum(ByVal c As String) As Boolean
    IsAlnum = (c Like "[0-9A-Za-z_]")
End Function

Private Function ExtractHexColors(ByVal txt As String, ByRef n As Long) As String
    Dim list As String, i As Long, j As Long, L As Long, tok As String, p As Long, q As Long
    L = Len(txt)
    If InStr(1, txt, "srgbClr", vbTextCompare) > 0 Then          ' Office theme-colors XML: srgbClr val="RRGGBB"
        p = 1
        Do
            p = InStr(p, txt, "srgbClr", vbTextCompare)
            If p = 0 Then Exit Do
            q = InStr(p, txt, "val=""", vbTextCompare)
            If q = 0 Then Exit Do
            tok = Mid$(txt, q + 5, 6)
            If IsHex6(tok) Then AddHex list, n, tok
            p = q + 5
        Loop
    Else                                                         ' hex codes: 6 hex characters standing alone
        i = 1
        Do While i <= L
            If IsAlnum(Mid$(txt, i, 1)) Then
                j = i
                Do While j <= L
                    If Not IsAlnum(Mid$(txt, j, 1)) Then Exit Do
                    j = j + 1
                Loop
                tok = Mid$(txt, i, j - i)
                If Len(tok) = 6 Then
                    If IsHex6(tok) Then AddHex list, n, tok
                End If
                i = j
            Else
                i = i + 1
            End If
        Loop
    End If
    If n = 0 Then ParseRgbTriples txt, list, n                   ' GIMP .gpl or "R,G,B" lists
    ExtractHexColors = list
End Function

Private Function IsSep(ByVal c As String) As Boolean
    IsSep = (c = " " Or c = vbTab Or c = "," Or c = ";")
End Function

Private Sub ParseRgbTriples(ByVal txt As String, ByRef list As String, ByRef n As Long)
    Dim i As Long, L As Long, k As Long, v(1 To 3) As Long, ok As Boolean, c As String, num As String
    L = Len(txt): i = 1
    Do While i <= L
        If Mid$(txt, i, 1) Like "[0-9]" Then
            ok = True: k = 0
            Do While ok And k < 3
                num = ""
                Do While i <= L
                    c = Mid$(txt, i, 1)
                    If Not (c Like "[0-9]") Then Exit Do
                    num = num & c: i = i + 1
                Loop
                If Len(num) = 0 Or Len(num) > 3 Then
                    ok = False
                Else
                    k = k + 1: v(k) = CLng(num)
                    If k < 3 Then
                        If i > L Then
                            ok = False
                        ElseIf Not IsSep(Mid$(txt, i, 1)) Then
                            ok = False
                        Else
                            Do While i <= L
                                If Not IsSep(Mid$(txt, i, 1)) Then Exit Do
                                i = i + 1
                            Loop
                        End If
                    End If
                End If
            Loop
            If ok And k = 3 Then
                If v(1) <= 255 And v(2) <= 255 And v(3) <= 255 Then AddHex list, n, RgbToHex(RGB(v(1), v(2), v(3)))
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub

Public Sub ImportPaletteFromFile()
    Dim f As Variant, hexes As String, n As Long
    On Error GoTo EH
    f = Application.GetOpenFilename(, , "Import color palette")      ' no filter: Mac uses a different filter syntax
    If VarType(f) = vbBoolean Then Exit Sub
    hexes = ExtractHexColors(ReadTextFile(CStr(f)), n)
    If n = 0 Then
        MsgBox "No colors found. The file should contain hex codes (#E4002B) or RGB triples (228,0,43).", vbInformation, APP_NAME
        Exit Sub
    End If
    ApplyPalette CurSheet(), hexes, n
    Exit Sub
EH:
    MsgBox "Could not read the palette file (error " & Err.Number & "): " & Err.Description, vbExclamation, APP_NAME
End Sub

Public Sub ResetPalette()
    Dim ws As Worksheet
    Set ws = CurSheet()
    CfgWrite ws, "Pal", ""
    If CfgRead(ws, "Theme") = "7" Then CfgWrite ws, "Theme", "1"
    RedrawActive
    RibbonRefresh
End Sub

'==============================================================================
'  RIBBON CALLBACKS  (customUI14.xml)
'==============================================================================
Public Sub Rib_OnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
    EnsureEvents
    On Error Resume Next
    Application.CommandBars("Sankey").Delete                 ' ribbon is available: drop fallback toolbar
End Sub

Public Sub Rib_Button(control As IRibbonControl)
    Select Case control.ID
        Case "btnCreate": CreateSankey
        Case "btnSample": InsertSampleData
        Case "btnNodeColor": ChangeNodeColor
        Case "btnResetNodes": ResetNodeColors
        Case "btnPalCells": ImportPaletteFromCells
        Case "btnPalFile": ImportPaletteFromFile
        Case "btnPalReset": ResetPalette
        Case "btnHelp": ShowHelp
        Case "btnAbout": ShowAbout
    End Select
End Sub

Private Function KeyOf(id As String) As String
    Select Case id
        Case "edtTitle": KeyOf = "Title"
        Case "edtSub": KeyOf = "Sub"
        Case "edtUnit": KeyOf = "Unit"
        Case "edtW": KeyOf = "W"
        Case "edtH": KeyOf = "H"
        Case "ddTheme": KeyOf = "Theme"
        Case "ddLink": KeyOf = "Link"
        Case "chkVals": KeyOf = "ShowVal"
        Case "chkAuto": KeyOf = "Auto"
    End Select
End Function

Public Sub Rib_GetText(control As IRibbonControl, ByRef text)
    text = CfgRead(CurSheet(), KeyOf(control.ID))
End Sub

Public Sub Rib_SetText(control As IRibbonControl, text As String)
    Dim k As String
    k = KeyOf(control.ID)
    If k = "W" Or k = "H" Then
        If Not IsNumeric(text) Then
            MsgBox "Please enter a number (points), e.g. 960.", vbExclamation, APP_NAME
            RibbonRefresh
            Exit Sub
        End If
        text = CStr(CLng(Val(text)))
    End If
    CfgWrite CurSheet(), k, text
    RedrawActive
End Sub

Public Sub Rib_GetIndex(control As IRibbonControl, ByRef index)
    index = CLng(Val(CfgRead(CurSheet(), KeyOf(control.ID)))) - 1
End Sub

Public Sub Rib_SetIndex(control As IRibbonControl, id As String, index As Integer)
    CfgWrite CurSheet(), KeyOf(control.ID), CStr(index + 1)
    RedrawActive
End Sub

Public Sub Rib_GetPressed(control As IRibbonControl, ByRef pressed)
    pressed = (CfgRead(CurSheet(), KeyOf(control.ID)) = "1")
End Sub

Public Sub Rib_SetPressed(control As IRibbonControl, pressed As Boolean)
    Dim k As String
    k = KeyOf(control.ID)
    CfgWrite CurSheet(), k, IIf(pressed, "1", "0")
    If k <> "Auto" Then RedrawActive
End Sub

'==============================================================================
'  STARTUP / FALLBACK TOOLBAR  (used only when the ribbon tab is not installed)
'==============================================================================
Public Sub Auto_Open()
    EnsureEvents
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 3), "'" & ThisWorkbook.Name & "'!CheckRibbon"
End Sub

Public Sub CheckRibbon()
    If gRibbon Is Nothing Then BuildToolbar
End Sub

Private Sub BuildToolbar()
    ' Excel for Mac has no floating toolbars; the Sankey ribbon tab is the only UI
End Sub

Private Sub AddToolbarButton(cb As Object, cap As String, macro As String)
    Dim b As Object
    Set b = cb.Controls.Add(Type:=msoControlButton)
    b.Caption = cap
    b.Style = msoButtonCaption
    b.OnAction = macro
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    gPolling = False
    Application.OnTime gPollNext, "'" & ThisWorkbook.Name & "'!PollResize", , False   ' cancel resize watcher
    Application.CommandBars("Sankey").Delete
    Set gEvents = Nothing
End Sub
