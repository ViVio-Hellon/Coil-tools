Attribute VB_Name = "GW用"
Option Explicit
'==============================================================================
' modTeslaConfig  ― ラベル転記用の列定数追加
'==============================================================================
Public Const COL_COILH As Long = 43   ' 本数
Public Const COL_TA    As Long = 48   ' 高さ
Public Const COL_NW    As Long = 53   ' NW
Public Const COL_GW    As Long = 59   ' GW

'==============================================================================
' modTeslaWrite  ― ラベル転記共通Sub
'==============================================================================

'--------------------------------------------------------------
' 1シート分のコイル情報転記
'--------------------------------------------------------------
' ws       : 対象シート
' dataRow  : 基準行 (3 or 92)
' extraRow : 追加行 (92 or 0=なし)
' coilH1   : 本数1段目
' coilH2   : 本数2段目
' ta1,ta2  : 高さ
' nw1,nw2  : NW
' gw1,gw2  : GW
'--------------------------------------------------------------
Public Sub WriteCoilInfo(ByVal ws As Worksheet, _
                         ByVal dataRow As Long, _
                         ByVal extraRow As Long, _
                         ByVal CoilH1 As String, ByVal CoilH2 As String, _
                         ByVal TA1 As String, ByVal TA2 As String, _
                         ByVal NW1 As String, ByVal NW2 As String, _
                         ByVal GW1 As String, ByVal GW2 As String)

    ws.Activate

    '--- 1段目 (dataRow) ---
    ws.Cells(dataRow, COL_COILH) = CoilH1
    ws.Cells(dataRow, COL_TA) = TA1
    ws.Cells(dataRow, COL_NW) = NW1
    ws.Cells(dataRow, COL_GW) = GW1

    ws.Cells(dataRow, HDR_WT1_COL) = val(ws.Cells(dataRow, HDR_WT1_COL))
    ws.Cells(dataRow, HDR_WT1_COL).NumberFormatLocal = "0.0_ "
    ws.Cells(dataRow, COL_COILH).NumberFormatLocal = "0_ "
    ws.Cells(dataRow, COL_TA).NumberFormatLocal = "0.0_ "
    ws.Cells(dataRow, COL_NW).NumberFormatLocal = "0.0_ "
    ws.Cells(dataRow, COL_GW).NumberFormatLocal = "0.0_ "

    '--- 2段目 (dataRow+1) ---
    ws.Cells(dataRow + 1, COL_COILH) = CoilH2
    ws.Cells(dataRow + 1, COL_TA) = TA2
    ws.Cells(dataRow + 1, COL_NW) = NW2
    ws.Cells(dataRow + 1, COL_GW) = GW2

    ws.Cells(dataRow + 1, COL_COILH).NumberFormatLocal = "0_ "
    ws.Cells(dataRow + 1, COL_TA).NumberFormatLocal = "0.0_ "
    ws.Cells(dataRow + 1, COL_NW).NumberFormatLocal = "0.0_ "
    ws.Cells(dataRow + 1, COL_GW).NumberFormatLocal = "0.0_ "

    '--- extraRow がある場合は同じ値を複写 ---
    If extraRow > 0 Then
        ws.Cells(extraRow, COL_COILH) = CoilH1
        ws.Cells(extraRow, COL_TA) = TA1
        ws.Cells(extraRow, COL_NW) = NW1
        ws.Cells(extraRow, COL_GW) = GW1

        ws.Cells(extraRow, HDR_WT1_COL) = val(ws.Cells(extraRow, HDR_WT1_COL))
        ws.Cells(extraRow, HDR_WT1_COL).NumberFormatLocal = "0.0_ "
        ws.Cells(extraRow, COL_COILH).NumberFormatLocal = "0_ "
        ws.Cells(extraRow, COL_TA).NumberFormatLocal = "0.0_ "
        ws.Cells(extraRow, COL_NW).NumberFormatLocal = "0.0_ "
        ws.Cells(extraRow, COL_GW).NumberFormatLocal = "0.0_ "

        ws.Cells(extraRow + 1, COL_COILH) = CoilH2
        ws.Cells(extraRow + 1, COL_TA) = TA2
        ws.Cells(extraRow + 1, COL_NW) = NW2
        ws.Cells(extraRow + 1, COL_GW) = GW2

        ws.Cells(extraRow + 1, COL_COILH).NumberFormatLocal = "0_ "
        ws.Cells(extraRow + 1, COL_TA).NumberFormatLocal = "0.0_ "
        ws.Cells(extraRow + 1, COL_NW).NumberFormatLocal = "0.0_ "
        ws.Cells(extraRow + 1, COL_GW).NumberFormatLocal = "0.0_ "
    End If

End Sub
'

'==============================================================================
' ラベル転記OP  ― リファクタリング版 (OB単体モード)
'==============================================================================
Sub ラベル転記OP()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim obIdx As Long
    obIdx = DetectSelectedOB(テスラ指定サイズ)
    If obIdx = 0 Then GoTo Cleanup

    Dim cfg As Variant:     cfg = GetSizeConfig(obIdx)
    Dim shtName As String:  shtName = BuildSheetName(obIdx)
    Dim takeType As Long:   takeType = GetTakeType(obIdx)
    Dim dataRow As Long:    dataRow = CLng(cfg(3))
    Dim extraRow As Long:   extraRow = CLng(cfg(4))

    '--- 丈1→サフィックス1,2  丈2→サフィックス3,4 ---
    Dim s1 As String, s2 As String
    If takeType = 1 Then s1 = "1": s2 = "2" Else s1 = "3": s2 = "4"

    With テスラ指定サイズ
        Call WriteCoilInfo( _
            ThisWorkbook.Sheets(shtName), dataRow, extraRow, _
            .Controls("CoilH" & s1), .Controls("CoilH" & s2), _
            .Controls("TA" & s1).Caption, .Controls("TA" & s2).Caption, _
            .Controls("NW" & s1).Caption, .Controls("NW" & s2).Caption, _
            .Controls("GW" & s1).Caption, .Controls("GW" & s2).Caption)
    End With

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'==============================================================================
' ラベル転記CK  ― リファクタリング版 (CB丈1,2同時モード)
'==============================================================================
Sub ラベル転記CK()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ob1 As Long, ob2 As Long
    Dim cbIdx As Long, Found As Boolean

    For cbIdx = 1 To 6
        If テスラ指定サイズ.Controls("CheckBox" & cbIdx).Value Then
            GetCBtoPair cbIdx, ob1, ob2
            Found = True
            Exit For
        End If
    Next cbIdx

    If Not Found Then
        If テスラ指定サイズ.ck_08_53.Value Then
            ob1 = 13: ob2 = 14
            Found = True
        End If
    End If

    If Not Found Then GoTo Cleanup

    Dim cfg1 As Variant: cfg1 = GetSizeConfig(ob1)
    Dim cfg2 As Variant: cfg2 = GetSizeConfig(ob2)

    With テスラ指定サイズ
        '--- 丈1 (Take1 が表示中なら) ---
        If .Take1.Visible Then
            Call WriteCoilInfo( _
                ThisWorkbook.Sheets(BuildSheetName(ob1)), _
                CLng(cfg1(3)), CLng(cfg1(4)), _
                .CoilH1, .CoilH2, _
                .TA1.Caption, .TA2.Caption, _
                .NW1.Caption, .NW2.Caption, _
                .GW1.Caption, .GW2.Caption)
        End If

        '--- 丈2 (Take2 が表示中なら) ---
        If .Take2.Visible Then
            Call WriteCoilInfo( _
                ThisWorkbook.Sheets(BuildSheetName(ob2)), _
                CLng(cfg2(3)), CLng(cfg2(4)), _
                .CoilH3, .CoilH4, _
                .TA3.Caption, .TA4.Caption, _
                .NW3.Caption, .NW4.Caption, _
                .GW3.Caption, .GW4.Caption)
        End If
    End With

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub 羅列計算()
    Dim i As Long, Q As Long
    Dim col As Double, TA As Single, GAI As Double
    Dim TIPT As Double
    
    '--- DB から資材マスタ取得 ---
    Dim tableName As String: tableName = "資材重量"
    Dim filePath As String:  filePath = PATH_AIM_参照 & "梱包資材マスタ.accdb"
    Dim TempHiki As Variant: TempHiki = GetRecordsArr(filePath, tableName, "")
    If IsEmpty(TempHiki) Then MsgBox "資材重量なし", vbCritical: Exit Sub
    
    Dim Hikifields As Variant: Hikifields = GetFieldsArr(filePath, tableName, "")
    If IsEmpty(Hikifields) Then MsgBox "資材重量なし", vbCritical: Exit Sub
    
    Dim fieldIdx As MaterialFieldIndices
    fieldIdx = GetFieldIndices(Hikifields)
    
    '--- DB から各資材の単重・係数を取得 ---
    Dim uPIN As Variant:     uPIN = GetUnitMassAndCoefficient(TempHiki, "テスラピン", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    Dim uTIP As Variant:     uTIP = GetUnitMassAndCoefficient(TempHiki, "チップボール1000ф", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    Dim uSUT As Variant:     uSUT = GetUnitMassAndCoefficient(TempHiki, "ストレッチフィルム", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    Dim uESA As Variant:     uESA = GetUnitMassAndCoefficient(TempHiki, "エサフォーム", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    Dim uPOR As Variant:     uPOR = GetUnitMassAndCoefficient(TempHiki, "ポリシート", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    Dim uPET As Variant:     uPET = GetUnitMassAndCoefficient(TempHiki, "PETバンド", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    Dim uPAR As Variant:     uPAR = GetUnitMassAndCoefficient(TempHiki, "樹脂パレット", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    Dim uDRY As Variant:     uDRY = GetUnitMassAndCoefficient(TempHiki, "EX-DRY", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    
    '--- HB 用マスタ ---
    Dim uHB As Variant:      uHB = GetUnitMassAndCoefficient(TempHiki, "ハードボード", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    Dim uHB八角 As Variant:  uHB八角 = GetUnitMassAndCoefficient(TempHiki, "八角ハードボード", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    Dim uHB590 As Variant:   uHB590 = GetUnitMassAndCoefficient(TempHiki, "HB590*2000", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    Dim uHB530 As Variant:   uHB530 = GetUnitMassAndCoefficient(TempHiki, "HB530*2000", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    Dim uHB580 As Variant:   uHB580 = GetUnitMassAndCoefficient(TempHiki, "HB580*2000", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
    Dim uHB500 As Variant:   uHB500 = GetUnitMassAndCoefficient(TempHiki, "HB500*2000", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)

    '--- TIP サイズ別 ---
    With テスラ指定サイズ
        If .TIP950 Then
            uTIP = GetUnitMassAndCoefficient(TempHiki, "チップボール950ф", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
            GAI = val(1000 * WorksheetFunction.PI() / 10 ^ 3)
        ElseIf .TIP820 Then
            uTIP = GetUnitMassAndCoefficient(TempHiki, "チップボール820ф", fieldIdx.梱包資材名, fieldIdx.単位質量, fieldIdx.係数)
            GAI = val(900 * WorksheetFunction.PI() / 10 ^ 3)
        Else
            GAI = val(1100 * WorksheetFunction.PI() / 10 ^ 3)
        End If
    End With

    '--- col（巾）の決定 ---
    col = GetSelectedCoilWidth()

    '--- 風袋シートの作成 ---
    Dim ListName As String: ListName = "50まで"
    Dim ws As Worksheet, Flag As Boolean
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = ListName Then Flag = True
    Next ws
    Application.ScreenUpdating = False
    If Flag Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(ListName).Delete
        Application.DisplayAlerts = True
    End If
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(Sheets.count)
    ActiveSheet.name = ListName
    ThisWorkbook.Sheets(ListName).Visible = False

    '--- ラベル設定 ---
    With ThisWorkbook.Sheets(ListName)
        .Cells(1, 1) = "本数"
        .Cells(1, 2) = "ﾋﾟﾝ" & Chr(10) & "(kg)"
        .Cells(1, 3) = "ﾁｯﾌﾟﾎﾞｰﾙ" & Chr(10) & "(kg)"
        .Cells(1, 4) = "ｽﾄﾚｯﾁﾌｨﾙﾑ" & Chr(10) & "(kg)"
        .Cells(1, 5) = "ｴｻﾌｫｰﾑ" & Chr(10) & "(kg)"
        .Cells(1, 6) = "ﾎﾟﾘｼｰﾄ" & Chr(10) & "(kg)"
        .Cells(1, 7) = "ﾊｰﾄﾞﾎﾞｰﾄﾞ" & Chr(10) & "(kg)"
        .Cells(1, 8) = "樹脂ﾊﾟﾚｯﾄ" & Chr(10) & "(kg)"
        .Cells(1, 9) = "ﾍﾟｯﾄﾊﾞﾝﾄﾞ" & Chr(10) & "(kg)"
        .Cells(1, 10) = "NW" & Chr(10) & "(kg)"
        .Cells(1, 11) = "風袋" & Chr(10) & "(kg)"
        .Cells(1, 12) = "GW" & Chr(10) & "(kg)"
        .Cells(1, 13) = "ｺｲﾙ高さ" & Chr(10) & "(mm)"
        .Cells(1, 14) = "梱包高さ" & Chr(10) & "(mm)"
        .Cells(1, 15) = "ﾁｯﾌﾟﾎﾞｰﾙ高さ" & Chr(10) & "(mm)"
    End With

    '--- 1～50本ループ ---
    Dim DOU As Double, TATE As Double
    Dim AREA As Double, AREA2 As Double
    Dim T1 As Double, T2 As Double
    Dim TempDiam As Long

    '--- TIP サイズで直径決定 ---
    With テスラ指定サイズ
        If .TIP1000 Then TempDiam = 1100
        If .TIP950 Then TempDiam = 1000
        If .TIP820 Then TempDiam = 900
    End With

    '--- 重量取得 ---
    With テスラ指定サイズ
        If .lblWeight1 <> "" Then T1 = CDbl(Replace(.lblWeight1, "kg", ""))
        If .lblWeight2 <> "" Then T2 = CDbl(Replace(.lblWeight2, "kg", ""))
    End With

    For Q = 1 To 50
        '--- 高さ計算 ---
        TIPT = val((Q + 1) * 0.7 / 10 ^ 3)
        TA = val(Q * col / 10 ^ 3) + TIPT

        With ThisWorkbook.Sheets(ListName)
            .Cells(Q + 1, 1) = Q
            '--- ピン ---
            .Cells(Q + 1, 2) = val(4 * uPIN(0) * uPIN(1))
            '--- チップボール ---
            .Cells(Q + 1, 3) = val((Q + 1) * uTIP(0) * uTIP(1))
            '--- ストレッチフィルム ---
            .Cells(Q + 1, 4) = val(GAI * 5 * uSUT(0) * uSUT(1))
            '--- エサフォーム ---
            If TA > 0.82 Then
                .Cells(Q + 1, 5) = val(GAI * 2 * uESA(0) * uESA(1))
            Else
                .Cells(Q + 1, 5) = val(GAI * uESA(0) * uESA(1))
            End If
            '--- ポリシート ---
            AREA = val((TempDiam / 2) ^ 2 * WorksheetFunction.PI())
            AREA2 = val((TA + TA / 2) * GAI)
            .Cells(Q + 1, 6) = val((AREA * 2 + AREA2) / 10 ^ 6 * uPOR(0) * uPOR(1))
        End With

        '--- ハードボード (HB) ---
        Dim hbVal As Double
        hbVal = CalcHBForList(col, Q, TA, GAI, _
                              uHB, uHB八角, uHB590, uHB530)
        ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = hbVal

        With ThisWorkbook.Sheets(ListName)
            '--- 樹脂パレット ---
            .Cells(Q + 1, 8) = uPAR(0) * uPAR(1) * 2
            '--- PETバンド ---
            DOU = val((GAI + 0.2) * 4)
            TATE = val((((TA + 0.2) * 4) * 2) + ((TempDiam * 4) * 2) / 10 ^ 3)
            .Cells(Q + 1, 9) = (DOU + TATE) * uPET(0) * uPET(1)
            '--- NW ---
            With テスラ指定サイズ
                If .Take1.Visible And .Take2.Visible Then
                    ThisWorkbook.Sheets(ListName).Cells(Q + 1, 10) = Format(val(Q * T1), "0.0")
                ElseIf .Take1.Visible Then
                    ThisWorkbook.Sheets(ListName).Cells(Q + 1, 10) = Format(val(Q * T1), "0.0")
                ElseIf .Take2.Visible Then
                    ThisWorkbook.Sheets(ListName).Cells(Q + 1, 10) = Format(val(Q * T2), "0.0")
                End If
            End With
            '--- 風袋・GW ---
            Dim r As Long: r = Q + 1
            .Cells(r, 11) = Format(Application.WorksheetFunction.Round( _
                .Cells(r, 2) + .Cells(r, 3) + .Cells(r, 4) + .Cells(r, 5) + _
                .Cells(r, 6) + .Cells(r, 7) + .Cells(r, 8) + .Cells(r, 9), 0), "0.0")
            .Cells(r, 12) = Format(Application.WorksheetFunction.Round( _
                .Cells(r, 2) + .Cells(r, 3) + .Cells(r, 4) + .Cells(r, 5) + _
                .Cells(r, 6) + .Cells(r, 7) + .Cells(r, 8) + .Cells(r, 9) + .Cells(r, 10), 0), "0.0")
            '--- コイル高さ・梱包高さ・チップ高さ ---
            .Cells(r, 13) = Format(val(TA * 1000))
            .Cells(r, 14) = Format(Application.WorksheetFunction.RoundUp(val((TA * 1000) + 150 + 150), -1), "0.0")
            .Cells(r, 15) = Format(val(TIPT * 1000))
            
            .Range("D2:G53,I2:I53").NumberFormatLocal = "#,##0.000_ "
            .Range("H2:H53,J2:O53").NumberFormatLocal = "0.0_ "
        End With
    Next Q

    '--- 書式・罫線・印刷設定 ---
    Call FormatListSheet(ListName)

    Application.ScreenUpdating = True
End Sub

'==============================================================================
' 書式・罫線・印刷設定（羅列計算から分離）
'==============================================================================
Private Sub FormatListSheet(ByVal ListName As String)
    With ThisWorkbook.Sheets(ListName)
        .Activate
        .Cells.Select
        With Selection.Font
            .name = "メイリオ"
            .Size = 10.5
            .ThemeFont = xlThemeFontNone
        End With
        .Cells.EntireColumn.AutoFit
        
        '--- 罫線 ---
        With .Range("A1:O51").Borders
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .weight = xlThin
        End With
        
        '--- 印刷設定 ---
        Application.PrintCommunication = False
        .PageSetup.PrintArea = "A1:O51"
        With .PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            
            With テスラ指定サイズ
                If .lblSize1 <> "" And .lblSize2 <> "" Then
                    ActiveSheet.PageSetup.LeftFooter = "＊検査番号：" & .lblKensaNo & "＊ｺｲﾙｻｲｽﾞ：" & .lblSize1 & " " & .lblSize2
                ElseIf .lblSize1 <> "" Then
                    ActiveSheet.PageSetup.LeftFooter = "＊検査番号：" & .lblKensaNo & "＊ｺｲﾙｻｲｽﾞ：" & .lblSize1
                ElseIf .lblSize2 <> "" Then
                    ActiveSheet.PageSetup.LeftFooter = "＊検査番号：" & .lblKensaNo & "＊ｺｲﾙｻｲｽﾞ：" & .lblSize2
                End If
            End With
            
            .CenterFooter = ""
            .RightFooter = Format$(Date, "yyyy.mm.dd") & "印刷"
            .LeftMargin = Application.InchesToPoints(0)
            .RightMargin = Application.InchesToPoints(0)
            .TopMargin = Application.InchesToPoints(0)
            .BottomMargin = Application.InchesToPoints(0)
            .HeaderMargin = Application.InchesToPoints(0)
            .FooterMargin = Application.InchesToPoints(0)
            .PaperSize = xlPaperA4
            .Orientation = xlPortrait
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End With
        Application.PrintCommunication = True
        ActiveWindow.View = xlPageBreakPreview
    End With
End Sub

'==============================================================================
' modTeslaConfig ― 巾取得の共通関数
'==============================================================================
Public Function GetSelectedCoilWidth() As Double
    Dim obIdx As Long
    
    '--- OB モード ---
    obIdx = DetectSelectedOB(テスラ指定サイズ)
    If obIdx > 0 Then
        Dim cfg As Variant: cfg = GetSizeConfig(obIdx)
        GetSelectedCoilWidth = ExtractWidthFromName(CStr(cfg(0)))
        Exit Function
    End If
    
    '--- CB モード ---
    Dim cbIdx As Long
    Dim ob1 As Long, ob2 As Long
    For cbIdx = 1 To 6
        If テスラ指定サイズ.Controls("CheckBox" & cbIdx).Value Then
            GetCBtoPair cbIdx, ob1, ob2
            cfg = GetSizeConfig(ob1)
            GetSelectedCoilWidth = ExtractWidthFromName(CStr(cfg(0)))
            Exit Function
        End If
    Next cbIdx
    
    If テスラ指定サイズ.ck_08_53.Value Then
        GetSelectedCoilWidth = 53.5
        Exit Function
    End If
    
    GetSelectedCoilWidth = 0
End Function

'==============================================================================
' modTeslaWrite ― 羅列計算用 HB 計算
'==============================================================================
' 積高から最適な HB サイズを選定
'--------------------------------------------------------------
Public Function CalcHBForList(ByVal col As Double, _
                              ByVal coilCount As Long, _
                              ByVal TA As Single, _
                              ByVal GAI As Double, _
                              ByVal uHB As Variant, _
                              ByVal uHB八角 As Variant, _
                              ByVal uHB590 As Variant, _
                              ByVal uHB530 As Variant) As Double

    If col = 53.5 And coilCount = 11 Then
        ' HB590*2000 上下2枚 + 八角2枚
        CalcHBForList = 2 * uHB590(0) * uHB590(1) _
                      + 2 * uHB八角(0) * uHB八角(1)

    ElseIf col = 53.5 And coilCount = 10 Then
        ' HB530*2000 上下2枚 + 八角2枚
        CalcHBForList = 2 * uHB530(0) * uHB530(1) _
                      + 2 * uHB八角(0) * uHB八角(1)

    Else
        ' ハードボード計算 + 八角2枚
        CalcHBForList = (GAI + 0.7) * TA * uHB(0) * uHB(1) _
                      + 2 * uHB八角(0) * uHB八角(1)
    End If

End Function

'--------------------------------------------------------------
' シート名ベースから巾を抽出  "1.0mm×73mm" → 73
'--------------------------------------------------------------
Private Function ExtractWidthFromName(ByVal baseName As String) As Double
    Dim parts As Variant
    parts = Split(baseName, "×")
    If UBound(parts) >= 1 Then
        ExtractWidthFromName = val(Replace(parts(1), "mm", ""))
    End If
End Function

'==============================================================================
' クリア  ― リファクタリング版（0.8mm×53.5mm 対応済み）
'==============================================================================
Sub クリア()
    Dim i As Long

    '--- フォーム側クリア ---
    With テスラ指定サイズ
        For i = 1 To 4
            .Controls("NW" & i) = ""
            .Controls("GW" & i) = ""
            .Controls("TA" & i) = ""
        Next i
    End With

    '--- 全シートのコイル情報クリア ---
    Application.ScreenUpdating = False

    Dim obIdx As Long
    Dim cfg   As Variant
    Dim ws    As Worksheet
    Dim dRow  As Long, exRow As Long

    For obIdx = 1 To 14   ' OB1～14 = 全シート
        cfg = GetSizeConfig(obIdx)
        dRow = CLng(cfg(3))
        exRow = CLng(cfg(4))

        Set ws = ThisWorkbook.Sheets(BuildSheetName(obIdx))

        ' dataRow, dataRow+1 をクリア
        ClearCoilCells ws, dRow

        ' extraRow がある場合はそちらもクリア
        If exRow > 0 Then ClearCoilCells ws, exRow
    Next obIdx

    Application.ScreenUpdating = True
End Sub

'--------------------------------------------------------------
' 指定行 + 次行の CoilH/TA/NW/GW 列をクリア
'--------------------------------------------------------------
Private Sub ClearCoilCells(ByVal ws As Worksheet, ByVal r As Long)
    ws.Cells(r, COL_COILH) = ""
    ws.Cells(r, COL_TA) = ""
    ws.Cells(r, COL_NW) = ""
    ws.Cells(r, COL_GW) = ""
    ws.Cells(r + 1, COL_COILH) = ""
    ws.Cells(r + 1, COL_TA) = ""
    ws.Cells(r + 1, COL_NW) = ""
    ws.Cells(r + 1, COL_GW) = ""
End Sub

Sub 印刷設定()
    '******************************************罫線
  If テスラ指定サイズ.Take1.Visible = True Then
    If テスラ指定サイズ.CoilH1 <> "" Then
    Range("A2:C12").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    End If
    If テスラ指定サイズ.CoilH2 <> "" Then
    Range("A14:C24").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    End If
  End If
    '******************************************罫線
  If テスラ指定サイズ.Take2.Visible = True Then
    If テスラ指定サイズ.CoilH3 <> "" Then
    Range("A26:C36").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    End If
    If テスラ指定サイズ.CoilH4 <> "" Then
    Range("A38:C48").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .weight = xlThin
    End With
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    End If
  End If
    '***************************************************印刷設定
  If テスラ指定サイズ.Take1.Visible = True And テスラ指定サイズ.Take2.Visible = True Then
   If テスラ指定サイズ.CoilH1 <> "" And テスラ指定サイズ.CoilH2 <> "" And テスラ指定サイズ.CoilH3 <> "" And テスラ指定サイズ.CoilH4 <> "" Then
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$1:$C$48"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = "＊検査番号：" & テスラ指定サイズ.lblKensaNo & "＊" & "ｺｲﾙｻｲｽﾞ：" & テスラ指定サイズ.lblSize1 & " " & テスラ指定サイズ.lblSize2
        .CenterFooter = ""
        .LeftFooter = ""
        .RightFooter = Format$(Date, "yyyy.mm.dd") & "印刷"
        
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape '縦xlPortrait　横 xllandscape
        .FitToPagesWide = 1 '1ページで納める
        .FitToPagesTall = 1
        '.Zoom = 100
    End With
    Application.PrintCommunication = True
   End If
  End If
  If テスラ指定サイズ.Take1.Visible = True And テスラ指定サイズ.Take2.Visible = False Then
   If テスラ指定サイズ.CoilH1 <> "" Or テスラ指定サイズ.CoilH2 <> "" Then
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$1:$C$24"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = "＊検査番号：" & テスラ指定サイズ.lblKensaNo & "＊" & "ｺｲﾙｻｲｽﾞ：" & テスラ指定サイズ.lblSize1
        .CenterFooter = ""
        .LeftFooter = ""
        .RightFooter = Format$(Date, "yyyy.mm.dd") & "印刷"
        
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape '縦xlPortrait　横 xllandscape
        .FitToPagesWide = 1 '1ページで納める
        .FitToPagesTall = 1
        '.Zoom = 100
    End With
   End If
  End If
  If テスラ指定サイズ.Take1.Visible = False And テスラ指定サイズ.Take2.Visible = True Then
   If テスラ指定サイズ.CoilH3 <> "" Or テスラ指定サイズ.CoilH4 <> "" Then
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$25:$C$48"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = "＊検査番号：" & テスラ指定サイズ.lblKensaNo & "＊" & "ｺｲﾙｻｲｽﾞ：" & テスラ指定サイズ.lblSize2
        .CenterFooter = ""
        .LeftFooter = ""
        .RightFooter = Format$(Date, "yyyy.mm.dd") & "印刷"
        
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape '縦xlPortrait　横 xllandscape
        .FitToPagesWide = 1 '1ページで納める
        .FitToPagesTall = 1
        '.Zoom = 100
    End With
    Application.PrintCommunication = True
   End If
  End If
End Sub

'Sub 計算1() '丈1
''NW　GW　高さの計算
''サイズ変更があった場合巾修正を忘れると高さ計算が狂う為注意
'Dim col As Double 'コイル巾
'Dim TA As Single, TA2 As Single 'コイル高さ
'Dim KTA As Single, KTA2 As Single '梱包高さ
'Dim PI As Double, PI2 As Double 'ﾋﾟﾝ
'Dim TIP As Double, TIP2 As Double 'ﾁｯﾌﾟﾎﾞｰﾙ
'Dim SUT As Double, SUT2 As Double 'ｽﾄﾚｯﾁﾌｨﾙﾑ
'Dim ESA As Double, ESA2 As Double 'ｴｻﾌｫ-ﾑ
'Dim POR As Double, POR2 As Double 'ポリシート
'Dim HB As Double, HB2 As Double 'ハードボード
'Dim PET As Double, PET2 As Double 'バンド
'Dim PAR As Double, PAR2 As Double 'パレット
'Dim NW As Double, NW2 As Double 'NW
'Dim HU As Double, HU2 As Double '風袋
'Dim GW As Double, GW2 As Double 'GW
'Dim i As Long
'Dim n As String
'n = ThisWorkbook.Sheets("設定").Cells(Rows.count, 1).End(xlUp).Row
' '*************************風袋シートの作成
'  Dim ListName As String, SheetsName As String
'  Dim ws As Worksheet, Flag As Boolean
'
'    ListName = "風袋計算"
'    SheetsName = ListName
'
'    For Each ws In ThisWorkbook.Worksheets
'        If ws.name = SheetsName Then Flag = True
'    Next ws
' Application.ScreenUpdating = False
'    If Flag = True Then    '同じ名前のシートあれば消す
'        Application.DisplayAlerts = False
'        ThisWorkbook.Sheets(SheetsName).Delete
'        Application.DisplayAlerts = True
'    Else
'    End If
'      ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(Sheets.count) '新たにシートを追加する
'      ActiveSheet.name = ListName
'      ThisWorkbook.Sheets(ListName).Visible = False
''    AT = Val(.ATU)   '板厚
''    HA = Val(.HABA)   '板幅
''    TA = Val(.TAKE)  '板丈
''    MA = Val(.MAI)  '枚数
''    MA2 = Val(.MAI) '- Val(.TextBox45) '枚数
''    B1 = Val(.TABA)   '縦バンド
''    B2 = Val(.YOBA)   '横バンド
''    KT = Val(.KOT)   '梱包高さ
''    PW = Val(.PAW)   'パレット重量
''    ip = Val(.TextBox38)    'Aインプット重量
''    GW = Val(.TextBox39)    'GW
''
''   BandL = (TA + (AT * tumisuu)) * 2 * .TABA  '縦バンド
''   BandW = ((HA * itaretu) + KT) * 2 * .YOBA  '横バンド
''
''  .TextBox30 = Format$((BandW + BandL) / 10 ^ 3 * W1(6) * K1(6), "0.00")   'PET
''  .TextBox31 = Format$((((((HA * itaretu) + TA) * 2) * (AT * tumisuu)) + (((HA * itaretu) * TA) * 2)) / 10 ^ 6 * W1(7) * K1(7), "0.00") 'ポリシート
''  .TextBox33 = Format$((HA * itaretu) * TA * 2 / 10 ^ 6 * W1(9) * K1(9), "0.00")  'ハードボード
'
' '*************************係数読み込み
'  For i = 3 To n 'n=15
'      W1(i - 2) = ThisWorkbook.Sheets("設定").Cells(i, 2)
'      K1(i - 2) = ThisWorkbook.Sheets("設定").Cells(i, 3)
'  Next i
' '*************************コイル高さ　新サイズの為修正OptionButton5、6　2019.04.26
'    With テスラ指定サイズ
'       If .CheckBox1 = True Then
'         col = 40
'       ElseIf .CheckBox2 = True Then
'         col = 33
'       ElseIf .CheckBox3 = True Then
'         col = 73
'       ElseIf .CheckBox4 = True Then
'         col = 82.5
'       ElseIf .CheckBox5 = True Then
'         col = 53.5
'       ElseIf .CheckBox6 = True Then
'         col = 63
'       End If
'
'       If .OptionButton7 = True Or .OptionButton8 = True Then
'         col = 40
'       ElseIf .OptionButton9 = True Or .OptionButton10 = True Then
'         col = 33
'       ElseIf .OptionButton1 = True Or .OptionButton2 = True Then
'         col = 73
'       ElseIf .OptionButton3 = True Or .OptionButton4 = True Then
'         col = 82.5
'       ElseIf .OptionButton5 = True Or .OptionButton6 = True Then
'         col = 53.5
'       ElseIf .OptionButton11 = True Or .OptionButton12 = True Then
'         col = 63
'       End If
'      If .CoilH1 = "" And .CoilH2 = "" Then
'        .NW1 = "" '本数が空白の時NWGWは空白にする
'        .GW1 = ""
'        .TA1 = ""
'        .NW2 = "" '本数が空白の時NWGWは空白にする
'        .GW2 = ""
'        .TA2 = ""
'      Exit Sub
'      Else
'        If IsNumeric(.CoilH1) = True Then
'        Dim TIPT As Double 'ﾁｯﾌﾟﾎﾞｰﾙの高さ　1枚0.7mm
'        TIPT = Val((.CoilH1 + 1) * 0.7 / 10 ^ 3)
'        'TAはコイルとチップの高さのみとしないと別の計算に影響する
'        'HB上下の高さが入っていなかった1枚2.35mm　2019.04.26
'         TA = Val(.CoilH1 * col / 10 ^ 3) + TIPT   'ﾃｷｽﾄﾎﾞｯｸｽの数字×ﾁｪｯｸの入っているｺｲﾙ 'M単位
'         KTA = Val((TA * 1000) + 150 + 150) + (2.35 * 2) '樹脂パレット高さ+150*2
'         .TA1 = Format(Application.WorksheetFunction.RoundUp(KTA, -1), "0.0") ' & Chr(10) & (TA * 1000) + 150 + 150 '　1桁繰り上げ
'        Else
'        End If
'
'        If IsNumeric(.CoilH2) = True Then
'        Dim TIPT2 As Double 'ﾁｯﾌﾟﾎﾞｰﾙの高さ　1枚0.7mm
'        TIPT2 = Val((.CoilH2 + 1) * 0.7 / 10 ^ 3)
'        'HB上下の高さが入っていなかった1枚2.35mm　2019.04.26
'         TA2 = Val(.CoilH2 * col / 10 ^ 3) + TIPT2  'ﾃｷｽﾄﾎﾞｯｸｽの数字×ﾁｪｯｸの入っているｺｲﾙ 'M単位
'         KTA2 = Val((TA2 * 1000) + 150 + 150) + (2.35 * 2) '樹脂パレット高さ+150*2
'         .TA2 = Format(Application.WorksheetFunction.RoundUp(KTA2, -1), "0.0") ' & Chr(10) & (TA2 * 1000) + 150 + 150 '　1桁繰り上げ
'        Else
'        End If
'      End If
'    End With
' '*************************ﾋﾟﾝの本数は固定で4本
'    PI = Val(4 * W1(1) * K1(1))
'    PI2 = Val(4 * W1(1) * K1(1))
' '*************************ﾁｯﾌﾟﾎﾞｰﾙ数は積み本数+1
'  With テスラ指定サイズ
'   If .CoilH1 <> "" Then
'    TIP = Val((.CoilH1 + 1) * W1(2) * K1(2))
'   End If
'
'   If .CoilH2 <> "" Then
'    TIP2 = Val((.CoilH2 + 1) * W1(2) * K1(2))
'   End If
'  End With
' '*************************ｽﾄﾚｯﾁﾌｨﾙﾑ
'   'コイル直径を1100とする(パレットサイズと同等)
'    Dim GAI As Double, GAI2 As Double
'    GAI = Val(1100 * 3.141592 / 10 ^ 3) 'M単位
'    GAI2 = Val(1100 * 3.141592 / 10 ^ 3) 'M単位
'    SUT = Val(GAI * 5 * W1(3) * K1(3)) '5周は巻く
'    SUT2 = Val(GAI2 * 5 * W1(3) * K1(3)) '5周は巻く
' '*************************ｴｻﾌｫｰﾑ
'    If TA > 0.82 Then 'エサフォームの高さが550　1.5倍ほどの高さになった時2周とする
'     ESA = Val(GAI * 2 * W1(4) * K1(4)) '2周
'    Else
'     ESA = Val(GAI * W1(4) * K1(4)) '1周
'    End If
'
'    If TA2 > 0.82 Then 'エサフォームの高さが550　1.5倍ほどの高さになった時2周とする
'     ESA2 = Val(GAI2 * 2 * W1(4) * K1(4)) '2周
'    Else
'     ESA2 = Val(GAI2 * W1(4) * K1(4)) '1周
'    End If
' '*************************ﾎﾟﾘｼｰﾄ
'   'コイル直径を1100とする(パレットサイズと同等)
'    Dim AREA As Double, AREA1 As Double '円部分
'    Dim AREA2 As Double, AREA3 As Double '同部分
'    AREA = Val(550 ^ 2 * 3.141592)
'    AREA2 = Val((TA + TA / 2) * GAI) '高さﾌﾟﾗｽ分で包む　ﾌﾟﾗｽ分は高さの半分とする
'    POR = Val((AREA * 2 + AREA2) / 10 ^ 6 * W1(6) * K1(6))
'
'    AREA1 = Val(550 ^ 2 * 3.141592)
'    AREA3 = Val((TA2 + TA2 / 2) * GAI2) '高さﾌﾟﾗｽ分で包む　ﾌﾟﾗｽ分は高さの半分とする
'    POR2 = Val((AREA1 * 2 + AREA3) / 10 ^ 6 * W1(6) * K1(6))
' '*************************HB
'    Call HB計算1
' '*************************ペットバンド
' '胴巻き4本
'  Dim DOU As Double, DOU2 As Double
'   DOU = Val((GAI + 0.2) * 4) '締める時の余分を200mmとした
'   DOU2 = Val((GAI2 + 0.2) * 4) '締める時の余分を200mmとした
' '縦バンド4本
'  Dim TATE As Double, TATE2 As Double
'   TATE = Val((((TA + 0.2) * 4) * 2) + ((1100 * 4) * 2) / 10 ^ 3) '締める時の余分を200mmとした
'   PET = (DOU + TATE) * W1(5) * K1(5)
'   TATE2 = Val((((TA2 + 0.2) * 4) * 2) + ((1100 * 4) * 2) / 10 ^ 3) '締める時の余分を200mmとした
'   PET2 = (DOU2 + TATE2) * W1(5) * K1(5)
' '*************************パレット
'  '樹脂パレット固定
'  PAR = Val(38) '2台分
'  PAR2 = Val(38) '2台分
' '*************************丈1と丈2の重量からKgを抜く　変数とﾌｫｰﾑﾗﾍﾞﾙに注意
' Dim T1 As Double, T2 As Double
'  With テスラ指定サイズ
'  If .lblWeight1 <> "" Then
'   T1 = Replace(.lblWeight1, "kg", "")
'  End If
'  If .lblWeight2 <> "" Then
'   T2 = Replace(.lblWeight2, "kg", "")
'  End If
' '*************************NWの計算1
'   If .CoilH1 = "" Or .CoilH1 = " " Then '丈１　１つめ
'    .NW1 = "" '本数が空白の時NWGWTAは空白にする
'    .GW1 = ""
'    .TA1 = ""
'   Else
'     If .Take1.Visible = True Then 'フレーム
'        NW = Format((Val(.CoilH1 * T1)), "0.0")
'       .NW1 = Format(NW, "0.0")
'      If .NW1 = "0.0" Then
'       .NW1 = ""
'      End If
'     End If
'     '*************************風袋の計算1
'       HU = Format(Val((PI + TIP + SUT + ESA + POR + Val(ThisWorkbook.Sheets(ListName).Cells(7, 2)) + PAR + PET)), "0.0")
'       HU = Format(Application.WorksheetFunction.Round((HU), 0), "0.0") '四捨五入繰り上げ
'     '*************************合計する GW1
'     If .NW1 <> "" Then
'       GW = Format(Val(PI + TIP + SUT + ESA + POR + Val(ThisWorkbook.Sheets(ListName).Cells(7, 2)) + PAR + PET + NW), "0.0")
'      .GW1 = Format(Application.WorksheetFunction.Round(GW, 0), "0.0") '四捨五入繰り上げ
'     ElseIf .NW1 = "" Or .NW1 = " " Then
'      .GW1 = ""
'     End If
'   End If
' '*************************NWの計算2
'   If .CoilH2 = "" Or .CoilH2 = " " Then  '丈１　2つめ
'    .NW2 = "" '本数が空白の時NWGWTAは空白にする
'    .GW2 = ""
'    .TA2 = ""
'   Else
'     If .Take1.Visible = True Then 'フレーム
'        NW2 = Format((Val(.CoilH2 * T1)), "0.0")
'       .NW2 = Format(NW2, "0.0")
'      If .NW2 = "0.0" Then
'       .NW2 = ""
'      End If
'     End If
'      '*************************風袋の計算2
'       HU2 = Format(Val(PI2 + TIP2 + SUT2 + ESA2 + POR2 + Val(ThisWorkbook.Sheets(ListName).Cells(19, 2)) + PAR2 + PET2), "0.0") '四捨五入繰り上げ
'       HU2 = Format(Application.WorksheetFunction.Round((HU2), 0), "0.0") '四捨五入繰り上げ
'      '*************************合計する GW2
'     If .NW2 <> "" Then
'       GW2 = Format(Val(PI2 + TIP2 + SUT2 + ESA2 + POR2 + Val(ThisWorkbook.Sheets(ListName).Cells(19, 2)) + PAR2 + PET2 + NW2), "0.0")
'      .GW2 = Format(Application.WorksheetFunction.Round(GW2, 0), "0.0") '四捨五入繰り上げ
'     ElseIf .NW2 = "" Or .NW2 = " " Then
'      .GW2 = ""
'     End If
'   End If
'
'  End With
' '*************************表記フォームへの記載
'  With 表記用
'  'クリア
'   .PI = ""
'   .TIP = ""
'   .SUT = ""
'   .ESA = ""
'   .POR = ""
'   .PAR = ""
'   .PET = ""
'   .NW = ""
'   .HU = ""
'   .GW = ""
'
'   .PI = Format(PI, "0.000")
'   .TIP = Format(TIP, "0.00")
'   .SUT = Format(SUT, "0.000")
'   .ESA = Format(ESA, "0.000")
'   .POR = Format(POR, "0.000")
'   .PAR = Format(PAR, "0.0")
'   .PET = Format(PET, "0.000")
'   .NW = Format(NW, "0.0")
'   .HU = Format(Application.WorksheetFunction.Round(HU, 0), "0.0")
'   .GW = Format(Application.WorksheetFunction.Round(GW, 0), "0.0")
' '*************************検算用
'  'クリア
'   .PIsiki = ""
'   .TIPsiki = ""
'   .SUTsiki = ""
'   .ESAsiki = ""
'   .PORsiki = ""
'   .PARsiki = ""
'   .PETsiki = ""
'   .NWsiki = ""
'   .HUsiki = ""
'   .GWsiki = ""
'
'   .PIsiki = 4 & "*" & W1(1) & "*" & K1(1)
'
'   .TIPsiki = "(" & テスラ指定サイズ.CoilH1 & "+" & 1 & ")" & "*" & W1(2) & "*" & K1(2)
'
'   .SUTsiki = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & 5 & "*" & W1(3) & "*" & K1(3) & "[約5周計算]"
'
'    If TA > 0.82 Then
'     .ESAsiki = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & 2 & "*" & W1(4) & "*" & K1(4) & "[2周計算]"
'    Else
'     .ESAsiki = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & W1(4) & "*" & K1(4) & "[1周計算]"
'    End If
'
'   .PORsiki = "[" & "(" & 550 & "^ 2 * 3.141592" & "*" & 2 & ")" & "]" & "+" & "[" & "(" & "(" & TA & "+" & TA & "/ 2)" & "*" & 1100 & "*" & 3.141592 & "/ 10 ^ 3" & ")" & ")" & _
'   "/ 10 ^ 6" & "]" & "*" & W1(6) & "*" & K1(6) & "  " & "包み分は高さの半分とする"
'
'   .PARsiki = 19 & "*" & 2
'
'   .PETsiki = "[" & "(" & 1100 & "*" & 3.141592 & "/ 10 ^ 3" & "+" & 0.2 & ")" & "*" & 4 & ")" & "]" & "+" & "[" & "(" & "(" _
'   & TA & "+" & 0.2 & ")" & "*" & 4 & ")" & "*" & 2 & ")" & "+" & "(" & "(" _
'   & 1100 & "*" & 4 & ")" & "*" & 2 & ") / 10 ^ 3" & "]" & "*" & W1(5) & "*" & K1(5) & "  " & "締める時の余分を200mmとした"
'
'    If テスラ指定サイズ.Take1.Visible = True Then
'      .NWsiki = テスラ指定サイズ.CoilH1 & "*" & T1
'    End If
'
'   .HUsiki = Format(PI, "0.00") & "+" & Format(TIP, "0.00") & "+" & Format(SUT, "0.00") & "+" & Format(ESA, "0.00") & "+" _
'   & Format(POR, "0.00") & "+" & Format(Val(ThisWorkbook.Sheets(ListName).Cells(7, 2)), "0.00") & "+" & Format(PAR, "0.0") & "+" & Format(PET, "0.00")
'
'   .GWsiki = Format(PI, "0.00") & "+" & Format(TIP, "0.00") & "+" & Format(SUT, "0.00") & "+" & Format(ESA, "0.00") & "+" _
'   & Format(POR, "0.00") & "+" & Format(Val(ThisWorkbook.Sheets(ListName).Cells(7, 2)), "0.00") & "+" & Format(PAR, "0.0") & "+" & Format(PET, "0.00") & "+" & Format(NW, "0.0")
'
'  End With
' '*************************検算用をシートへ移す
' With ThisWorkbook.Sheets(ListName)
'  If テスラ指定サイズ.CoilH1 <> "" Then '************************************************1
'  .Cells(2, 3) = 4 & "*" & W1(1) & "*" & K1(1)
'
'  .Cells(3, 3) = "(" & テスラ指定サイズ.CoilH1 & "+" & 1 & ")" & "*" & W1(2) & "*" & K1(2)
'
'  .Cells(4, 3) = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & 5 & "*" & W1(3) & "*" & K1(3) & "[約5周計算]"
'
'    If TA > 0.82 Then
'     .Cells(5, 3) = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & 2 & "*" & W1(4) & "*" & K1(4) & "[2周計算]"
'    Else
'     .Cells(5, 3) = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & W1(4) & "*" & K1(4) & "[1周計算]"
'    End If
'
'  .Cells(6, 3) = "[" & "(" & 550 & "^ 2 * 3.141592" & "*" & 2 & ")" & "]" & "+" & "[" & "(" & "(" & TA & "+" & TA & "/ 2)" & "*" & 1100 & "*" & 3.141592 & "/ 10 ^ 3" & ")" & ")" & _
'   "/ 10 ^ 6" & "]" & "*" & W1(6) & "*" & K1(6) & "  " & "包み分は高さの半分とする"
'
'  .Cells(8, 3) = 19 & "*" & 2
'
'  .Cells(9, 3) = "[" & "(" & 1100 & "*" & 3.141592 & "/ 10 ^ 3" & "+" & 0.2 & ")" & "*" & 4 & ")" & "]" & "+" & "[" & "(" & "(" _
'   & TA & "+" & 0.2 & ")" & "*" & 4 & ")" & "*" & 2 & ")" & "+" & "(" & "(" _
'   & 1100 & "*" & 4 & ")" & "*" & 2 & ") / 10 ^ 3" & "]" & "*" & W1(5) & "*" & K1(5) & "  " & "締める時の余分を200mmとした"
'
'    If テスラ指定サイズ.Take1.Visible = True Then
'      .Cells(10, 3) = テスラ指定サイズ.CoilH1 & "*" & T1
'    End If
'
'  .Cells(11, 3) = Format(PI, "0.000") & "+" & Format(TIP, "0.00") & "+" & Format(SUT, "0.000") & "+" & Format(ESA, "0.000") & "+" _
'   & Format(POR, "0.000") & "+" & Format(Val(ThisWorkbook.Sheets(ListName).Cells(7, 2)), "0.000") & "+" & Format(PAR, "0.0") & "+" & Format(PET, "0.000")
'
'  .Cells(12, 3) = Format(PI, "0.000") & "+" & Format(TIP, "0.00") & "+" & Format(SUT, "0.000") & "+" & Format(ESA, "0.000") & "+" _
'   & Format(POR, "0.000") & "+" & Format(Val(ThisWorkbook.Sheets(ListName).Cells(7, 2)), "0.000") & "+" & Format(PAR, "0.0") & "+" & Format(PET, "0.000") & "+" & Format(NW, "0.0")
'
'    .Cells(1, 1) = "丈1.1梱包目:" & テスラ指定サイズ.CoilH1 & "本"
'    .Cells(2, 1) = "ﾋﾟﾝ"
'    .Cells(3, 1) = "ﾁｯﾌﾟﾎﾞｰﾙ"
'    .Cells(4, 1) = "ｽﾄﾚｯﾁﾌｨﾙﾑ"
'    .Cells(5, 1) = "ｴｻﾌｫｰﾑ"
'    .Cells(6, 1) = "ﾎﾟﾘｼｰﾄ"
'    .Cells(7, 1) = "ﾊｰﾄﾞﾎﾞｰﾄﾞ"
'    .Cells(8, 1) = "樹脂ﾊﾟﾚｯﾄ"
'    .Cells(9, 1) = "ﾍﾟｯﾄﾊﾞﾝﾄﾞ"
'    .Cells(10, 1) = "NW"
'    .Cells(11, 1) = "風袋"
'    .Cells(12, 1) = "GW"
'
'    .Cells(2, 2) = Format(PI, "0.000")
'    .Cells(2, 2).NumberFormatLocal = "0.000_ "
'    .Cells(3, 2) = Format(TIP, "0.00")
'    .Cells(3, 2).NumberFormatLocal = "0.00_ "
'    .Cells(4, 2) = Format(SUT, "0.000")
'    .Cells(4, 2).NumberFormatLocal = "0.000_ "
'    .Cells(5, 2) = Format(ESA, "0.000")
'    .Cells(5, 2).NumberFormatLocal = "0.000_ "
'    .Cells(6, 2) = Format(POR, "0.000")
'    .Cells(6, 2).NumberFormatLocal = "0.000_ "
'    .Cells(8, 2) = Format(PAR, "0.0")
'    .Cells(8, 2).NumberFormatLocal = "0.0_ "
'    .Cells(9, 2) = Format(PET, "0.000")
'    .Cells(9, 2).NumberFormatLocal = "0.000_ "
'    .Cells(10, 2) = Format(NW, "0.0")
'    .Cells(10, 2).NumberFormatLocal = "0.0_ "
'    .Cells(11, 2) = Format(Application.WorksheetFunction.Round(HU, 0), "0.0")
'    .Cells(11, 2).NumberFormatLocal = "0.0_ "
'    .Cells(12, 2) = Format(Application.WorksheetFunction.Round(GW, 0), "0.0") 'Format(GW, "0.0") '
'    .Cells(12, 2).NumberFormatLocal = "0.0_ "
'  End If
'  If テスラ指定サイズ.CoilH2 <> "" Then '************************************************2
'  .Cells(14, 3) = 4 & "*" & W1(1) & "*" & K1(1)
'
'  .Cells(15, 3) = "(" & テスラ指定サイズ.CoilH2 & "+" & 1 & ")" & "*" & W1(2) & "*" & K1(2)
'
'  .Cells(16, 3) = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & 5 & "*" & W1(3) & "*" & K1(3) & "[約5周計算]"
'
'    If TA > 0.82 Then
'     .Cells(17, 3) = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & 2 & "*" & W1(4) & "*" & K1(4) & "[2周計算]"
'    Else
'     .Cells(17, 3) = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & W1(4) & "*" & K1(4) & "[1周計算]"
'    End If
'
'  .Cells(18, 3) = "[" & "(" & 550 & "^ 2 * 3.141592" & "*" & 2 & ")" & "]" & "+" & "[" & "(" & "(" & TA2 & "+" & TA & "/ 2)" & "*" & 1100 & "*" & 3.141592 & "/ 10 ^ 3" & ")" & ")" & _
'   "/ 10 ^ 6" & "]" & "*" & W1(6) & "*" & K1(6) & "  " & "包み分は高さの半分とする"
'
'  .Cells(20, 3) = 19 & "*" & 2
'
'  .Cells(21, 3) = "[" & "(" & 1100 & "*" & 3.141592 & "/ 10 ^ 3" & "+" & 0.2 & ")" & "*" & 4 & ")" & "]" & "+" & "[" & "(" & "(" _
'   & TA2 & "+" & 0.2 & ")" & "*" & 4 & ")" & "*" & 2 & ")" & "+" & "(" & "(" _
'   & 1100 & "*" & 4 & ")" & "*" & 2 & ") / 10 ^ 3" & "]" & "*" & W1(5) & "*" & K1(5) & "  " & "締める時の余分を200mmとした"
'
'    If テスラ指定サイズ.Take1.Visible = True Then
'      .Cells(22, 3) = テスラ指定サイズ.CoilH2 & "*" & T1
'    End If
'
'  .Cells(23, 3) = Format(PI2, "0.000") & "+" & Format(TIP2, "0.00") & "+" & Format(SUT2, "0.000") & "+" & Format(ESA2, "0.000") & "+" _
'   & Format(POR2, "0.000") & "+" & Format(Val(ThisWorkbook.Sheets(ListName).Cells(19, 2)), "0.000") & "+" & Format(PAR2, "0.0") & "+" & Format(PET2, "0.000")
'
'  .Cells(24, 3) = Format(PI2, "0.000") & "+" & Format(TIP2, "0.00") & "+" & Format(SUT2, "0.000") & "+" & Format(ESA2, "0.000") & "+" _
'   & Format(POR2, "0.000") & "+" & Format(Val(ThisWorkbook.Sheets(ListName).Cells(19, 2)), "0.000") & "+" & Format(PAR2, "0.0") & "+" & Format(PET2, "0.000") & "+" & Format(NW2, "0.0")
'
'    .Cells(13, 1) = "丈1.2梱包目:" & テスラ指定サイズ.CoilH2 & "本"
'    .Cells(14, 1) = "ﾋﾟﾝ"
'    .Cells(15, 1) = "ﾁｯﾌﾟﾎﾞｰﾙ"
'    .Cells(16, 1) = "ｽﾄﾚｯﾁﾌｨﾙﾑ"
'    .Cells(17, 1) = "ｴｻﾌｫｰﾑ"
'    .Cells(18, 1) = "ﾎﾟﾘｼｰﾄ"
'    .Cells(19, 1) = "ﾊｰﾄﾞﾎﾞｰﾄﾞ"
'    .Cells(20, 1) = "樹脂ﾊﾟﾚｯﾄ"
'    .Cells(21, 1) = "ﾍﾟｯﾄﾊﾞﾝﾄﾞ"
'    .Cells(22, 1) = "NW"
'    .Cells(23, 1) = "風袋"
'    .Cells(24, 1) = "GW"
'
'    .Cells(14, 2) = Format(PI2, "0.000")
'    .Cells(14, 2).NumberFormatLocal = "0.000_ "
'    .Cells(15, 2) = Format(TIP2, "0.00")
'    .Cells(15, 2).NumberFormatLocal = "0.00_ "
'    .Cells(16, 2) = Format(SUT2, "0.000")
'    .Cells(16, 2).NumberFormatLocal = "0.000_ "
'    .Cells(17, 2) = Format(ESA2, "0.000")
'    .Cells(17, 2).NumberFormatLocal = "0.000_ "
'    .Cells(18, 2) = Format(POR2, "0.000")
'    .Cells(18, 2).NumberFormatLocal = "0.000_ "
'    .Cells(20, 2) = Format(PAR2, "0.0")
'    .Cells(20, 2).NumberFormatLocal = "0.0_ "
'    .Cells(21, 2) = Format(PET2, "0.000")
'    .Cells(21, 2).NumberFormatLocal = "0.000_ "
'    .Cells(22, 2) = Format(NW2, "0.0")
'    .Cells(22, 2).NumberFormatLocal = "0.0_ "
'    .Cells(23, 2) = Format(Application.WorksheetFunction.Round(HU2, 0), "0.0")
'    .Cells(23, 2).NumberFormatLocal = "0.0_ "
'    .Cells(24, 2) = Format(Application.WorksheetFunction.Round(GW2, 0), "0.0")
'    .Cells(24, 2).NumberFormatLocal = "0.0_ "
'  End If
'  If テスラ指定サイズ.Take1.Visible = False Then '丈1非表示のときは1の部分をクリア
'   ThisWorkbook.Sheets(ListName).Range("A1:C24").ClearContents
'  End If
'  If テスラ指定サイズ.Take2.Visible = False Then '丈２非表示のときは2の部分をクリア
'    ThisWorkbook.Sheets(ListName).Range("A25:C48").ClearContents
'  End If
' End With
'    ThisWorkbook.Sheets(ListName).Activate
'    Cells.Select
'    With Selection.Font
'        .name = "メイリオ"
'        .size = 10.5
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With
'    ThisWorkbook.Sheets(ListName).Cells.EntireColumn.AutoFit
'
' Call 印刷設定
'
' Application.ScreenUpdating = True
'End Sub
'''Sub HB計算1()
'''Dim col As Double 'コイル巾
'''Dim TA As Single, TA2 As Single 'コイル高さ
'''Dim HB As Double, HB2 As Double 'ハードボード
'''Dim ListName As String, SheetsName As String
'''Dim ws As Worksheet, Flag As Boolean
'''Dim i As Long
'''Dim n As String
'''n = ThisWorkbook.Sheets("設定").Cells(Rows.count, 1).End(xlUp).Row
'''ListName = "風袋計算"
'''SheetsName = ListName
'''
'''   表記用.HB = "" '表記xﾌｵｰﾑｸﾘｱ
''' '*************************係数読み込み
'''  For i = 3 To n
'''      W1(i - 2) = ThisWorkbook.Sheets("設定").Cells(i, 2)
'''      K1(i - 2) = ThisWorkbook.Sheets("設定").Cells(i, 3)
'''  Next i
''' '*************************コイル直径を1100とする(パレットサイズと同等)
'''    Dim GAI As Double, GAI2 As Double
'''    GAI = Val(1100 * 3.141592 / 10 ^ 3) 'M単位
'''    GAI2 = Val(1100 * 3.141592 / 10 ^ 3) 'M単位
''' '*************************コイル高さ
'''    With テスラ指定サイズ
'''       If .CheckBox1 = True Then
'''         col = 40
'''       ElseIf .CheckBox2 = True Then
'''         col = 33
'''       ElseIf .CheckBox3 = True Then
'''         col = 73
'''       ElseIf .CheckBox4 = True Then
'''         col = 82.5
'''       ElseIf .CheckBox5 = True Then
'''         col = 53.5
'''       ElseIf .CheckBox6 = True Then
'''         col = 63
'''       End If
'''
'''       If .OptionButton7 = True Or .OptionButton8 = True Then
'''         col = 40
'''       ElseIf .OptionButton9 = True Or .OptionButton10 = True Then
'''         col = 33
'''       ElseIf .OptionButton1 = True Or .OptionButton2 = True Then
'''         col = 73
'''       ElseIf .OptionButton3 = True Or .OptionButton4 = True Then
'''         col = 82.5
'''       ElseIf .OptionButton5 = True Or .OptionButton6 = True Then
'''         col = 53.5
'''       ElseIf .OptionButton11 = True Or .OptionButton12 = True Then
'''         col = 63
'''       End If
'''      If .CoilH1 = "" And .CoilH2 = "" Then
'''      Exit Sub
'''      Else
'''
'''        If IsNumeric(.CoilH1) = True Then
'''        Dim TIPT As Double 'ﾁｯﾌﾟﾎﾞｰﾙの高さ　1枚0.7mm
'''        TIPT = Val((.CoilH1 + 1) * 0.7 / 10 ^ 3)
'''         TA = Val(.CoilH1 * col / 10 ^ 3) + TIPT 'ﾃｷｽﾄﾎﾞｯｸｽの数字×ﾁｪｯｸの入っているｺｲﾙ 'M単位
'''        Else
'''        End If
'''
'''        If IsNumeric(.CoilH2) = True Then
'''        Dim TIPT2 As Double 'ﾁｯﾌﾟﾎﾞｰﾙの高さ　1枚0.7mm
'''        TIPT2 = Val((.CoilH2 + 1) * 0.7 / 10 ^ 3)
'''         TA2 = Val(.CoilH2 * col / 10 ^ 3) + TIPT2 'ﾃｷｽﾄﾎﾞｯｸｽの数字×ﾁｪｯｸの入っているｺｲﾙ 'M単位
'''        Else
'''        End If
'''
'''      End If
'''    End With
''' '*************************HB
'''   '固定物2種とそれ以外を分ける　500*2000 5700*2000
'''   '固定物八角HB
'''    With テスラ指定サイズ
'''       If .CheckBox1 = True Then
'''         Select Case .CoilH1
'''         Case 14
'''          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case Else '外周よりやや長いHB
'''          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         End Select
'''         Select Case .CoilH2
'''         Case 14
'''          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case Else '外周よりやや長いHB
'''          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         End Select
'''       ElseIf .CheckBox2 = True Then
'''         Select Case .CoilH1
'''         Case 15
'''          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case Else
'''          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         End Select
'''         Select Case .CoilH2
'''         Case 15
'''          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case Else
'''          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         End Select
'''       ElseIf .CheckBox3 = True Then
'''         Select Case .CoilH1
'''         Case 8
'''          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case 7
'''          HB = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case Else
'''          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         End Select
'''         Select Case .CoilH2
'''         Case 8
'''          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case 7
'''          HB2 = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case Else
'''          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         End Select
'''       ElseIf .CheckBox4 = True Then
'''         Select Case .CoilH1
'''         Case 7
'''          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case 6
'''          HB = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case Else
'''          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         End Select
'''         Select Case .CoilH2
'''         Case 7
'''          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case 6
'''          HB2 = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case Else
'''          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         End Select
'''       ElseIf .CheckBox5 = True Then '新サイズ53巾　ﾃﾞﾌｫ積み数は10本
'''         Select Case .CoilH1
'''         Case 10
'''          HB = 2 * W1(17) * K1(17) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(17) & "*" & K1(17) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(17) & "*" & K1(17) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case Else
'''          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         End Select
'''         Select Case .CoilH2
'''         Case 10
'''          HB2 = 2 * W1(17) * K1(17) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(17) & "*" & K1(17) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case Else
'''          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         End Select
'''       ElseIf .CheckBox6 = True Then
'''         Select Case .CoilH1
'''         Case 9
'''          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case 8
'''          HB = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case Else
'''          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         End Select
'''         Select Case .CoilH2
'''         Case 9
'''          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case 8
'''          HB2 = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case Else
'''          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         End Select
'''       End If
'''     '*******************************************************************************
'''       If .OptionButton7 = True Or .OptionButton8 = True Then
'''         Select Case .CoilH1
'''         Case 14
'''          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case Else
'''          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         End Select
'''         Select Case .CoilH2
'''         Case 14
'''          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case Else
'''          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         End Select
'''       ElseIf .OptionButton9 = True Or .OptionButton10 = True Then
'''         Select Case .CoilH1
'''         Case 15
'''          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case Else
'''          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         End Select
'''         Select Case .CoilH2
'''         Case 15
'''          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case Else
'''          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         End Select
'''       ElseIf .OptionButton1 = True Or .OptionButton2 = True Then
'''         Select Case .CoilH1
'''         Case 8
'''          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case 7
'''          HB = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case Else
'''          HB = (GAI + 0.4) * TA * W1(7) * K1(7)
'''          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         End Select
'''         Select Case .CoilH2
'''         Case 8
'''          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case 7
'''          HB2 = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case Else
'''          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         End Select
'''       ElseIf .OptionButton3 = True Or .OptionButton4 = True Then
'''         Select Case .CoilH1
'''         Case 7
'''          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case 6
'''          HB = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case Else
'''          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         End Select
'''         Select Case .CoilH2
'''         Case 7
'''          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case 6
'''          HB2 = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case Else
'''          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         End Select
'''       ElseIf .OptionButton5 = True Or .OptionButton6 = True Then '新サイズ53巾　ﾃﾞﾌｫ積み数10本
'''         Select Case .CoilH1
'''         Case 10
'''          HB = 2 * W1(17) * K1(17) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(17) & "*" & K1(17) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(17) & "*" & K1(17) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case Else
'''          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         End Select
'''         Select Case .CoilH2
'''         Case 10
'''          HB2 = 2 * W1(17) * K1(17) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(17) & "*" & K1(17) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case Else
'''          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         End Select
'''       ElseIf .OptionButton11 = True Or .OptionButton12 = True Then
'''         Select Case .CoilH1
'''         Case 9
'''          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case 8
'''          HB = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'''          表記用.HBsiki = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         Case Else
'''          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          If テスラ指定サイズ.CoilH1 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(7, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(7, 2) = Format(HB, "0.000")
'''          End If
'''         End Select
'''         Select Case .CoilH2
'''         Case 9
'''          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case 8
'''          HB2 = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         Case Else
'''          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'''          If テスラ指定サイズ.CoilH2 <> "" Then
'''          ThisWorkbook.Sheets(ListName).Cells(19, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'''          ThisWorkbook.Sheets(ListName).Cells(19, 2) = Format(HB2, "0.000")
'''          End If
'''         End Select
'''       End If
'''    End With
'''
'''   表記用.HB = Format(HB, "0.000")
'''
'''End Sub
''Sub 計算2() '丈2
'''NW　GW　高さの計算
'''サイズ変更があった場合巾修正を忘れると高さ計算が狂う為注意
''Dim col As Double 'コイル巾
''Dim TA As Single, TA2 As Single 'コイル高さ
''Dim KTA As Single, KTA2 As Single '梱包高さ
''Dim PI As Double, PI2 As Double 'ﾋﾟﾝ
''Dim TIP As Double, TIP2 As Double 'ﾁｯﾌﾟﾎﾞｰﾙ
''Dim SUT As Double, SUT2 As Double 'ｽﾄﾚｯﾁﾌｨﾙﾑ
''Dim ESA As Double, ESA2 As Double 'ｴｻﾌｫ-ﾑ
''Dim POR As Double, POR2 As Double 'ポリシート
''Dim HB As Double, HB2 As Double 'ハードボード
''Dim PET As Double, PET2 As Double 'バンド
''Dim PAR As Double, PAR2 As Double 'パレット
''Dim NW As Double, NW2 As Double 'NW
''Dim HU As Double, HU2 As Double '風袋
''Dim GW As Double, GW2 As Double 'GW
''Dim i As Long
''Dim n As String
''n = ThisWorkbook.Sheets("設定").Cells(Rows.count, 1).End(xlUp).Row
'' '*************************風袋シートの作成
''  Dim ListName As String, SheetsName As String
''  Dim ws As Worksheet, Flag As Boolean
''
''    ListName = "風袋計算"
''    SheetsName = ListName
''
''    For Each ws In ThisWorkbook.Worksheets
''        If ws.name = SheetsName Then Flag = True
''    Next ws
'' Application.ScreenUpdating = False
''    If Flag = True Then    '同じ名前のシートあればそのまま
''    Else 'なければ新たにシートを追加する
''      ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(Sheets.count)
''      ActiveSheet.name = ListName
''      ThisWorkbook.Sheets(ListName).Visible = False
''    End If
'' '*************************係数読み込み
''  For i = 3 To n 'n=15
''      W1(i - 2) = ThisWorkbook.Sheets("設定").Cells(i, 2)
''      K1(i - 2) = ThisWorkbook.Sheets("設定").Cells(i, 3)
''  Next i
'' '*************************コイル高さ　新サイズの為修正OptionButton5、6　2019.04.26
''    With テスラ指定サイズ
''       If .CheckBox1 = True Then
''         col = 40
''       ElseIf .CheckBox2 = True Then
''         col = 33
''       ElseIf .CheckBox3 = True Then
''         col = 73
''       ElseIf .CheckBox4 = True Then
''         col = 82.5
''       ElseIf .CheckBox5 = True Then
''         col = 53.5
''       ElseIf .CheckBox6 = True Then
''         col = 63
''       End If
''
''       If .OptionButton7 = True Or .OptionButton8 = True Then
''         col = 40
''       ElseIf .OptionButton9 = True Or .OptionButton10 = True Then
''         col = 33
''       ElseIf .OptionButton1 = True Or .OptionButton2 = True Then
''         col = 73
''       ElseIf .OptionButton3 = True Or .OptionButton4 = True Then
''         col = 82.5
''       ElseIf .OptionButton5 = True Or .OptionButton6 = True Then
''         col = 53.5
''       ElseIf .OptionButton11 = True Or .OptionButton12 = True Then
''         col = 63
''       End If
''
''      If .CoilH3 = "" And .CoilH4 = "" Then
''        .NW3 = "" '本数が空白の時NWGWは空白にする
''        .GW3 = ""
''        .TA3 = ""
''        .NW4 = "" '本数が空白の時NWGWは空白にする
''        .GW4 = ""
''        .TA4 = ""
''      Exit Sub
''      Else
''      '2019.05.08 .CoilH3が.CoilH1になっていた修正
''      If IsNumeric(.CoilH3) = True Then
''        Dim TIPT As Double 'ﾁｯﾌﾟﾎﾞｰﾙの高さ　1枚0.7mm
''        TIPT = Val((.CoilH3 + 1) * 0.7 / 10 ^ 3)
''        'TAはコイルとチップの高さのみとしないと別の計算に影響する
''        'HB上下の高さが入っていなかった1枚2.35mm　2019.04.26
''         TA = Val(.CoilH3 * col / 10 ^ 3) + TIPT   'ﾃｷｽﾄﾎﾞｯｸｽの数字×ﾁｪｯｸの入っているｺｲﾙ 'M単位
''         KTA = Val((TA * 1000) + 150 + 150) + (2.35 * 2) '樹脂パレット高さ+150*2
''         .TA3 = Format(Application.WorksheetFunction.RoundUp(KTA, -1), "0.0") ' & Chr(10) & (TA * 1000) + 150 + 150 '　1桁繰り上げ
''        Else
''        End If
''
''        If IsNumeric(.CoilH4) = True Then
''        Dim TIPT2 As Double 'ﾁｯﾌﾟﾎﾞｰﾙの高さ　1枚0.7mm
''        TIPT2 = Val((.CoilH4 + 1) * 0.7 / 10 ^ 3)
''        'HB上下の高さが入っていなかった1枚2.35mm　2019.04.26
''         TA2 = Val(.CoilH4 * col / 10 ^ 3) + TIPT2  'ﾃｷｽﾄﾎﾞｯｸｽの数字×ﾁｪｯｸの入っているｺｲﾙ 'M単位
''         KTA2 = Val((TA2 * 1000) + 150 + 150) + (2.35 * 2) '樹脂パレット高さ+150*2
''         .TA4 = Format(Application.WorksheetFunction.RoundUp(KTA2, -1), "0.0") ' & Chr(10) & (TA2 * 1000) + 150 + 150 '　1桁繰り上げ
''        Else
''        End If
''      End If
''    End With
'' '*************************ﾋﾟﾝの本数は固定で4本
''    PI = Val(4 * W1(1) * K1(1))
''    PI2 = Val(4 * W1(1) * K1(1))
'' '*************************ﾁｯﾌﾟﾎﾞｰﾙ数は積み本数+1
''  With テスラ指定サイズ
''   If .CoilH3 <> "" Then
''    TIP = Val((.CoilH3 + 1) * W1(2) * K1(2))
''   End If
''
''   If .CoilH4 <> "" Then
''    TIP2 = Val((.CoilH4 + 1) * W1(2) * K1(2))
''   End If
''  End With
'' '*************************ｽﾄﾚｯﾁﾌｨﾙﾑ
''   'コイル直径を1100とする(パレットサイズと同等)
''    Dim GAI As Double, GAI2 As Double
''    GAI = Val(1100 * 3.141592 / 10 ^ 3) 'M単位
''    GAI2 = Val(1100 * 3.141592 / 10 ^ 3) 'M単位
''    SUT = Val(GAI * 5 * W1(3) * K1(3)) '5周は巻く
''    SUT2 = Val(GAI2 * 5 * W1(3) * K1(3)) '5周は巻く
'' '*************************ｴｻﾌｫｰﾑ
''    If TA > 0.82 Then 'エサフォームの高さが550　1.5倍ほどの高さになった時2周とする
''     ESA = Val(GAI * 2 * W1(4) * K1(4)) '2周
''    Else
''     ESA = Val(GAI * W1(4) * K1(4)) '1周
''    End If
''
''    If TA2 > 0.82 Then 'エサフォームの高さが550　1.5倍ほどの高さになった時2周とする
''     ESA2 = Val(GAI2 * 2 * W1(4) * K1(4)) '2周
''    Else
''     ESA2 = Val(GAI2 * W1(4) * K1(4)) '1周
''    End If
'' '*************************ﾎﾟﾘｼｰﾄ
''   'コイル直径を1100とする(パレットサイズと同等)
''    Dim AREA As Double, AREA1 As Double '円部分
''    Dim AREA2 As Double, AREA3 As Double '同部分
''    AREA = Val(550 ^ 2 * 3.141592)
''    AREA2 = Val((TA + TA / 2) * GAI) '高さﾌﾟﾗｽ分で包む　ﾌﾟﾗｽ分は高さの半分とする
''    POR = Val((AREA * 2 + AREA2) / 10 ^ 6 * W1(6) * K1(6))
''
''    AREA1 = Val(550 ^ 2 * 3.141592)
''    AREA3 = Val((TA2 + TA2 / 2) * GAI2) '高さﾌﾟﾗｽ分で包む　ﾌﾟﾗｽ分は高さの半分とする
''    POR2 = Val((AREA1 * 2 + AREA3) / 10 ^ 6 * W1(6) * K1(6))
'' '*************************HB
''   Call HB計算2
'' '*************************ペットバンド
'' '胴巻き4本
''  Dim DOU As Double, DOU2 As Double
''   DOU = Val((GAI + 0.2) * 4) '締める時の余分を200mmとした
''   DOU2 = Val((GAI2 + 0.2) * 4) '締める時の余分を200mmとした
'' '縦バンド4本
''  Dim TATE As Double, TATE2 As Double
''   TATE = Val((((TA + 0.2) * 4) * 2) + ((1100 * 4) * 2) / 10 ^ 3) '締める時の余分を200mmとした
''   PET = (DOU + TATE) * W1(5) * K1(5)
''   TATE2 = Val((((TA2 + 0.2) * 4) * 2) + ((1100 * 4) * 2) / 10 ^ 3) '締める時の余分を200mmとした
''   PET2 = (DOU2 + TATE2) * W1(5) * K1(5)
'' '*************************パレット
''  '樹脂パレット固定
''  PAR = Val(38) '2台分
''  PAR2 = Val(38) '2台分
'' '*************************丈1と丈2の重量からKgを抜く 変数とﾌｫｰﾑﾗﾍﾞﾙに注意
'' Dim T1 As Double, T2 As Double
''  With テスラ指定サイズ
''  If .lblWeight1 <> "" Then
''   T1 = Replace(.lblWeight1, "kg", "")
''  End If
''  If .lblWeight2 <> "" Then
''   T2 = Replace(.lblWeight2, "kg", "")
''  End If
'' '*************************NWの計算1
''   If .CoilH3 = "" Or .CoilH3 = " " Then '丈１　１つめ
''    .NW3 = "" '本数が空白の時NWGWTAは空白にする
''    .GW3 = ""
''    .TA3 = ""
''   Else
''     If .Take2.Visible = True Then 'フレーム
''        NW = Format((Val(.CoilH3 * T2)), "0.0")
''       .NW3 = Format(NW, "0.0")
''      If .NW3 = "0.0" Then
''       .NW3 = ""
''      End If
''     End If
''  '*************************風袋の計算1
''       HU = Format(Val((PI + TIP + SUT + ESA + POR + Val(ThisWorkbook.Sheets(ListName).Cells(31, 2)) + PAR + PET)), "0.0")
''       HU = Format(Application.WorksheetFunction.Round((HU), 0), "0.0") '四捨五入繰り上げ
''  '*************************合計する GW1
''     If .NW3 <> "" Then
''       GW = Format(Val(PI + TIP + SUT + ESA + POR + Val(ThisWorkbook.Sheets(ListName).Cells(31, 2)) + PAR + PET + NW), "0.0")
''      .GW3 = Format(Application.WorksheetFunction.Round(GW, 0), "0.0") '四捨五入繰り上げ
''     Else
''     .GW3 = ""
''     End If
''   End If
''
'' '*******************************************************NWの計算2
''
''   If .CoilH4 = "" Or .CoilH4 = " " Then  '丈１　2つめ
''    .NW4 = "" '本数が空白の時NWGWTAは空白にする
''    .GW4 = ""
''    .TA4 = ""
''   Else
''     If .Take2.Visible = True Then 'フレーム
''        NW2 = Format((Val(.CoilH4 * T2)), "0.0")
''       .NW4 = Format(NW2, "0.0")
''      If .NW4 = "0.0" Then
''       .NW4 = ""
''      End If
''     End If
''  '*************************風袋の計算2
''       HU2 = Format(Val(PI2 + TIP2 + SUT2 + ESA2 + POR2 + Val(ThisWorkbook.Sheets(ListName).Cells(43, 2)) + PAR2 + PET2), "0.0") '四捨五入繰り上げ
''       HU2 = Format(Application.WorksheetFunction.Round((HU2), 0), "0.0") '四捨五入繰り上げ
''  '*************************合計する GW2
''     If .NW4 <> "" Then
''       GW2 = Format(Val(PI2 + TIP2 + SUT2 + ESA2 + POR2 + Val(ThisWorkbook.Sheets(ListName).Cells(43, 2)) + PAR2 + PET2 + NW2), "0.0")
''      .GW4 = Format(Application.WorksheetFunction.Round(GW2, 0), "0.0") '四捨五入繰り上げ
''     Else
''      .GW4 = ""
''     End If
''   End If
''
''  End With
'' '*************************検算用をシートへ移す
'' With ThisWorkbook.Sheets(ListName)
''  If テスラ指定サイズ.CoilH3 <> "" Then '************************************************3
''  .Cells(26, 3) = 4 & "*" & W1(1) & "*" & K1(1)
''
''  .Cells(27, 3) = "(" & テスラ指定サイズ.CoilH1 & "+" & 1 & ")" & "*" & W1(2) & "*" & K1(2)
''
''  .Cells(28, 3) = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & 5 & "*" & W1(3) & "*" & K1(3) & "[約5周計算]"
''
''    If TA > 0.82 Then
''     .Cells(29, 3) = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & 2 & "*" & W1(4) & "*" & K1(4) & "[2周計算]"
''    Else
''     .Cells(29, 3) = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & W1(4) & "*" & K1(4) & "[1周計算]"
''    End If
''
''  .Cells(30, 3) = "[" & "(" & 550 & "^ 2 * 3.141592" & "*" & 2 & ")" & "]" & "+" & "[" & "(" & "(" & TA & "+" & TA & "/ 2)" & "*" & 1100 & "*" & 3.141592 & "/ 10 ^ 3" & ")" & ")" & _
''   "/ 10 ^ 6" & "]" & "*" & W1(6) & "*" & K1(6) & "  " & "包み分は高さの半分とする"
''
''  .Cells(32, 3) = 19 & "*" & 2
''
''  .Cells(33, 3) = "[" & "(" & 1100 & "*" & 3.141592 & "/ 10 ^ 3" & "+" & 0.2 & ")" & "*" & 4 & ")" & "]" & "+" & "[" & "(" & "(" _
''   & TA & "+" & 0.2 & ")" & "*" & 4 & ")" & "*" & 2 & ")" & "+" & "(" & "(" _
''   & 1100 & "*" & 4 & ")" & "*" & 2 & ") / 10 ^ 3" & "]" & "*" & W1(5) & "*" & K1(5) & "  " & "締める時の余分を200mmとした"
''
''    If テスラ指定サイズ.Take2.Visible = True Then
''      .Cells(34, 3) = テスラ指定サイズ.CoilH3 & "*" & T2
''    End If
''
''  .Cells(35, 3) = Format(PI, "0.00") & "+" & Format(TIP, "0.00") & "+" & Format(SUT, "0.00") & "+" & Format(ESA, "0.00") & "+" _
''   & Format(POR, "0.00") & "+" & Format(Val(ThisWorkbook.Sheets(ListName).Cells(31, 2)), "0.00") & "+" & Format(PAR, "0.0") & "+" & Format(PET, "0.00")
''
''  .Cells(36, 3) = Format(PI, "0.00") & "+" & Format(TIP, "0.00") & "+" & Format(SUT, "0.00") & "+" & Format(ESA, "0.00") & "+" _
''   & Format(POR, "0.00") & "+" & Format(Val(ThisWorkbook.Sheets(ListName).Cells(31, 2)), "0.00") & "+" & Format(PAR, "0.0") & "+" & Format(PET, "0.00") & "+" & Format(NW, "0.0")
''
''    .Cells(25, 1) = "丈2.1梱包目:" & テスラ指定サイズ.CoilH3 & "本"
''    .Cells(26, 1) = "ﾋﾟﾝ"
''    .Cells(27, 1) = "ﾁｯﾌﾟﾎﾞｰﾙ"
''    .Cells(28, 1) = "ｽﾄﾚｯﾁﾌｨﾙﾑ"
''    .Cells(29, 1) = "ｴｻﾌｫｰﾑ"
''    .Cells(30, 1) = "ﾎﾟﾘｼｰﾄ"
''    .Cells(31, 1) = "ﾊｰﾄﾞﾎﾞｰﾄﾞ"
''    .Cells(32, 1) = "樹脂ﾊﾟﾚｯﾄ"
''    .Cells(33, 1) = "ﾍﾟｯﾄﾊﾞﾝﾄﾞ"
''    .Cells(34, 1) = "NW"
''    .Cells(35, 1) = "風袋"
''    .Cells(36, 1) = "GW"
''
''    .Cells(26, 2) = Format(PI, "0.000")
''    .Cells(26, 2).NumberFormatLocal = "0.000_ "
''    .Cells(27, 2) = Format(TIP, "0.00")
''    .Cells(27, 2).NumberFormatLocal = "0.00_ "
''    .Cells(28, 2) = Format(SUT, "0.000")
''    .Cells(28, 2).NumberFormatLocal = "0.000_ "
''    .Cells(29, 2) = Format(ESA, "0.000")
''    .Cells(29, 2).NumberFormatLocal = "0.000_ "
''    .Cells(30, 2) = Format(POR, "0.000")
''    .Cells(30, 2).NumberFormatLocal = "0.000_ "
''    .Cells(32, 2) = Format(PAR, "0.0")
''    .Cells(32, 2).NumberFormatLocal = "0.0_ "
''    .Cells(33, 2) = Format(PET, "0.000")
''    .Cells(33, 2).NumberFormatLocal = "0.000_ "
''    .Cells(34, 2) = Format(NW, "0.0")
''    .Cells(34, 2).NumberFormatLocal = "0.0_ "
''    .Cells(35, 2) = Format(Application.WorksheetFunction.Round(HU, 0), "0.0")
''    .Cells(35, 2).NumberFormatLocal = "0.0_ "
''    .Cells(36, 2) = Format(Application.WorksheetFunction.Round(GW, 0), "0.0")
''    .Cells(36, 2).NumberFormatLocal = "0.0_ "
''  End If
''  If テスラ指定サイズ.CoilH4 <> "" Then '************************************************4
''  .Cells(38, 3) = 4 & "*" & W1(1) & "*" & K1(1)
''
''  .Cells(39, 3) = "(" & テスラ指定サイズ.CoilH2 & "+" & 1 & ")" & "*" & W1(2) & "*" & K1(2)
''
''  .Cells(40, 3) = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & 5 & "*" & W1(3) & "*" & K1(3) & "[約5周計算]"
''
''    If TA > 0.82 Then
''     .Cells(41, 3) = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & 2 & "*" & W1(4) & "*" & K1(4) & "[2周計算]"
''    Else
''     .Cells(41, 3) = 1100 & "*" & 3.141592 & "/ 10 ^ 3)" & "*" & W1(4) & "*" & K1(4) & "[1周計算]"
''    End If
''
''  .Cells(42, 3) = "[" & "(" & 550 & "^ 2 * 3.141592" & "*" & 2 & ")" & "]" & "+" & "[" & "(" & "(" & TA2 & "+" & TA & "/ 2)" & "*" & 1100 & "*" & 3.141592 & "/ 10 ^ 3" & ")" & ")" & _
''   "/ 10 ^ 6" & "]" & "*" & W1(6) & "*" & K1(6) & "  " & "包み分は高さの半分とする"
''
''  .Cells(44, 3) = 19 & "*" & 2
''
''  .Cells(45, 3) = "[" & "(" & 1100 & "*" & 3.141592 & "/ 10 ^ 3" & "+" & 0.2 & ")" & "*" & 4 & ")" & "]" & "+" & "[" & "(" & "(" _
''   & TA2 & "+" & 0.2 & ")" & "*" & 4 & ")" & "*" & 2 & ")" & "+" & "(" & "(" _
''   & 1100 & "*" & 4 & ")" & "*" & 2 & ") / 10 ^ 3" & "]" & "*" & W1(5) & "*" & K1(5) & "  " & "締める時の余分を200mmとした"
''
''    If テスラ指定サイズ.Take2.Visible = True Then
''      .Cells(46, 3) = テスラ指定サイズ.CoilH4 & "*" & T2
''    End If
''
''  .Cells(47, 3) = Format(PI2, "0.00") & "+" & Format(TIP2, "0.00") & "+" & Format(SUT2, "0.00") & "+" & Format(ESA2, "0.00") & "+" _
''   & Format(POR2, "0.00") & "+" & Format(Val(ThisWorkbook.Sheets(ListName).Cells(43, 2)), "0.00") & "+" & Format(PAR2, "0.0") & "+" & Format(PET2, "0.00")
''
''  .Cells(48, 3) = Format(PI2, "0.00") & "+" & Format(TIP2, "0.00") & "+" & Format(SUT2, "0.00") & "+" & Format(ESA2, "0.00") & "+" _
''   & Format(POR2, "0.00") & "+" & Format(Val(ThisWorkbook.Sheets(ListName).Cells(43, 2)), "0.00") & "+" & Format(PAR2, "0.0") & "+" & Format(PET2, "0.00") & "+" & Format(NW2, "0.0")
''
''    .Cells(37, 1) = "丈2.2梱包目:" & テスラ指定サイズ.CoilH4 & "本"
''    .Cells(38, 1) = "ﾋﾟﾝ"
''    .Cells(39, 1) = "ﾁｯﾌﾟﾎﾞｰﾙ"
''    .Cells(40, 1) = "ｽﾄﾚｯﾁﾌｨﾙﾑ"
''    .Cells(41, 1) = "ｴｻﾌｫｰﾑ"
''    .Cells(42, 1) = "ﾎﾟﾘｼｰﾄ"
''    .Cells(43, 1) = "ﾊｰﾄﾞﾎﾞｰﾄﾞ"
''    .Cells(44, 1) = "樹脂ﾊﾟﾚｯﾄ"
''    .Cells(45, 1) = "ﾍﾟｯﾄﾊﾞﾝﾄﾞ"
''    .Cells(46, 1) = "NW"
''    .Cells(47, 1) = "風袋"
''    .Cells(48, 1) = "GW"
''
''    .Cells(38, 2) = Format(PI2, "0.000")
''    .Cells(38, 2).NumberFormatLocal = "0.000_ "
''    .Cells(39, 2) = Format(TIP2, "0.00")
''    .Cells(39, 2).NumberFormatLocal = "0.00_ "
''    .Cells(40, 2) = Format(SUT2, "0.000")
''    .Cells(40, 2).NumberFormatLocal = "0.000_ "
''    .Cells(41, 2) = Format(ESA2, "0.000")
''    .Cells(41, 2).NumberFormatLocal = "0.000_ "
''    .Cells(42, 2) = Format(POR2, "0.000")
''    .Cells(42, 2).NumberFormatLocal = "0.000_ "
''    .Cells(44, 2) = Format(PAR2, "0.0")
''    .Cells(44, 2).NumberFormatLocal = "0.0_ "
''    .Cells(45, 2) = Format(PET2, "0.000")
''    .Cells(45, 2).NumberFormatLocal = "0.000_ "
''    .Cells(46, 2) = Format(NW2, "0.0")
''    .Cells(46, 2).NumberFormatLocal = "0.0_ "
''    .Cells(47, 2) = Format(Application.WorksheetFunction.Round(HU2, 0), "0.0")
''    .Cells(47, 2).NumberFormatLocal = "0.0_ "
''    .Cells(48, 2) = Format(Application.WorksheetFunction.Round(GW2, 0), "0.0")
''    .Cells(48, 2).NumberFormatLocal = "0.0_ "
''  End If
''  If テスラ指定サイズ.Take1.Visible = False Then '丈1非表示のときは1の部分をクリア
''   ThisWorkbook.Sheets(ListName).Range("A1:C24").ClearContents
''  End If
''  If テスラ指定サイズ.Take2.Visible = False Then '丈２非表示のときは2の部分をクリア
''    ThisWorkbook.Sheets(ListName).Range("A25:C48").ClearContents
''  End If
'' End With
''    ThisWorkbook.Sheets(ListName).Activate
''    Cells.Select
''    With Selection.Font
''        .name = "メイリオ"
''        .size = 10.5
''        .Strikethrough = False
''        .Superscript = False
''        .Subscript = False
''        .OutlineFont = False
''        .Shadow = False
''        .Underline = xlUnderlineStyleNone
''        .ThemeColor = xlThemeColorLight1
''        .TintAndShade = 0
''        .ThemeFont = xlThemeFontNone
''    End With
''    ThisWorkbook.Sheets(ListName).Cells.EntireColumn.AutoFit
''
'' Call 印刷設定
''
'' Application.ScreenUpdating = True
''End Sub
'
'Sub HB計算2()
'Dim col As Double 'コイル巾
'Dim TA As Single, TA2 As Single 'コイル高さ
'Dim HB As Double, HB2 As Double 'ハードボード
'Dim ListName As String, SheetsName As String
'Dim ws As Worksheet, Flag As Boolean
'Dim i As Long
'Dim n As String
'n = ThisWorkbook.Sheets("設定").Cells(Rows.count, 1).End(xlUp).Row
'ListName = "風袋計算"
'SheetsName = ListName
' '*************************係数読み込み
'  For i = 3 To n 'n=15
'      W1(i - 2) = ThisWorkbook.Sheets("設定").Cells(i, 2)
'      K1(i - 2) = ThisWorkbook.Sheets("設定").Cells(i, 3)
'  Next i
' '*************************コイル直径を1100とする(パレットサイズと同等)
'    Dim GAI As Double, GAI2 As Double
'    GAI = Val(1100 * 3.141592 / 10 ^ 3) 'M単位
'    GAI2 = Val(1100 * 3.141592 / 10 ^ 3) 'M単位
' '*************************コイル高さ
'    With テスラ指定サイズ
'       If .CheckBox1 = True Then
'         col = 40
'       ElseIf .CheckBox2 = True Then
'         col = 33
'       ElseIf .CheckBox3 = True Then
'         col = 73
'       ElseIf .CheckBox4 = True Then
'         col = 82.5
'       ElseIf .CheckBox5 = True Then
'         col = 53.5
'       ElseIf .CheckBox6 = True Then
'         col = 63
'       End If
'
'       If .OptionButton7 = True Or .OptionButton8 = True Then
'         col = 40
'       ElseIf .OptionButton9 = True Or .OptionButton10 = True Then
'         col = 33
'       ElseIf .OptionButton1 = True Or .OptionButton2 = True Then
'         col = 73
'       ElseIf .OptionButton3 = True Or .OptionButton4 = True Then
'         col = 82.5
'       ElseIf .OptionButton5 = True Or .OptionButton6 = True Then
'         col = 53.5
'       ElseIf .OptionButton11 = True Or .OptionButton12 = True Then
'         col = 63
'       End If
'
'        If IsNumeric(.CoilH3) = True Then
'        Dim TIPT As Double 'ﾁｯﾌﾟﾎﾞｰﾙの高さ　1枚0.7mm
'        TIPT = Val((.CoilH3 + 1) * 0.7 / 10 ^ 3)
'         TA = Val(.CoilH3 * col / 10 ^ 3) + TIPT 'ﾃｷｽﾄﾎﾞｯｸｽの数字×ﾁｪｯｸの入っているｺｲﾙ 'M単位
'        Else
'        End If
'
'        If IsNumeric(.CoilH4) = True Then
'        Dim TIPT2 As Double 'ﾁｯﾌﾟﾎﾞｰﾙの高さ　1枚0.7mm
'        TIPT2 = Val((.CoilH4 + 1) * 0.7 / 10 ^ 3)
'         TA2 = Val(.CoilH4 * col / 10 ^ 3) + TIPT2 'ﾃｷｽﾄﾎﾞｯｸｽの数字×ﾁｪｯｸの入っているｺｲﾙ 'M単位
'        Else
'        End If
'
'    End With
' '*************************HB
'   '固定物2種とそれ以外を分ける　500*2000 5700*2000
'   '固定物八角HB
'    With テスラ指定サイズ
'       If .CheckBox1 = True Then
'         Select Case .CoilH3
'         Case 14
'          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case Else '外周よりやや長いHB
'          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         End Select
'         Select Case .CoilH4
'         Case 14
'          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case Else '外周よりやや長いHB
'          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         End Select
'       ElseIf .CheckBox2 = True Then
'         Select Case .CoilH3
'         Case 15
'          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case Else
'          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         End Select
'         Select Case .CoilH4
'         Case 15
'          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case Else
'          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         End Select
'       ElseIf .CheckBox3 = True Then
'         Select Case .CoilH3
'         Case 8
'          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case 7
'          HB = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case Else
'          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         End Select
'         Select Case .CoilH4
'         Case 8
'          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case 7
'          HB2 = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case Else
'          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         End Select
'       ElseIf .CheckBox4 = True Then
'         Select Case .CoilH3
'         Case 7
'          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case 6
'          HB = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case Else
'          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         End Select
'         Select Case .CoilH4
'         Case 7
'          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case 6
'          HB2 = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case Else
'          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         End Select
'       ElseIf .CheckBox5 = True Then '2019.06.10 53巾 ﾃﾞﾌｫ積み数は10本積み
'         Select Case .CoilH3
'         Case 10
'          HB = 2 * W1(17) * K1(17) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(17) & "*" & K1(17) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(17) & "*" & K1(17) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case Else
'          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         End Select
'         Select Case .CoilH4
'         Case 10
'          HB2 = 2 * W1(17) * K1(17) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(17) & "*" & K1(17) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case Else
'          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         End Select
'       ElseIf .CheckBox6 = True Then
'         Select Case .CoilH3
'         Case 9
'          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case 8
'          HB = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case Else
'          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         End Select
'         Select Case .CoilH4
'         Case 9
'          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case 8
'          HB2 = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case Else
'          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         End Select
'       End If
'       '********************************************************************
'       If .OptionButton7 = True Or .OptionButton8 = True Then
'         Select Case .CoilH3
'         Case 14
'          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case Else
'          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         End Select
'         Select Case .CoilH4
'         Case 14
'          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case Else
'          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         End Select
'       ElseIf .OptionButton9 = True Or .OptionButton10 = True Then
'         Select Case .CoilH3
'         Case 15
'          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case Else
'          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         End Select
'         Select Case .CoilH4
'         Case 15
'          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case Else
'          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         End Select
'       ElseIf .OptionButton1 = True Or .OptionButton2 = True Then
'         Select Case .CoilH3
'         Case 8
'          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case 7
'          HB = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case Else
'          HB = (GAI + 0.4) * TA * W1(7) * K1(7)
'          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         End Select
'         Select Case .CoilH4
'         Case 8
'          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case 7
'          HB2 = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case Else
'          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         End Select
'       ElseIf .OptionButton3 = True Or .OptionButton4 = True Then
'         Select Case .CoilH3
'         Case 7
'          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case 6
'          HB = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case Else
'          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         End Select
'         Select Case .CoilH4
'         Case 7
'          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case 6
'          HB2 = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case Else
'          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         End Select
'       ElseIf .OptionButton5 = True Or .OptionButton6 = True Then '2019.06.10 53巾 ﾃﾞﾌｫ積み数は10本積み
'         Select Case .CoilH3
'         Case 10
'          HB = 2 * W1(17) * K1(17) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(17) & "*" & K1(17) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(17) & "*" & K1(17) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case Else
'          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         End Select
'         Select Case .CoilH4
'         Case 10
'          HB2 = 2 * W1(17) * K1(17) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(17) & "*" & K1(17) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case Else
'          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         End Select
'       ElseIf .OptionButton11 = True Or .OptionButton12 = True Then
'         Select Case .CoilH3
'         Case 9
'          HB = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case 8
'          HB = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'          表記用.HBsiki = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         Case Else
'          HB = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          表記用.HBsiki = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          If テスラ指定サイズ.CoilH3 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(31, 3) = "(" & GAI & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(31, 2) = Format(HB, "0.000")
'          End If
'         End Select
'         Select Case .CoilH4
'         Case 9
'          HB2 = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(9) & "*" & K1(9) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case 8
'          HB2 = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = 2 & "*" & W1(8) & "*" & K1(8) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[上下2枚胴巻き2枚]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         Case Else
'          HB2 = (GAI2 + 0.4) * TA2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'          If テスラ指定サイズ.CoilH4 <> "" Then
'          ThisWorkbook.Sheets(ListName).Cells(43, 3) = "(" & GAI2 & "+" & 0.4 & ")" & "*" & TA & "*" & W1(7) & "*" & K1(7) & "+" & 2 & "*" & W1(10) & "*" & K1(10) & "[外周よりやや長いHB]"
'          ThisWorkbook.Sheets(ListName).Cells(43, 2) = Format(HB2, "0.000")
'          End If
'         End Select
'       End If
'    End With
'
'
'End Sub

'
'Sub 羅列計算()
'Dim col As Double 'コイル巾
'Dim TA As Single 'コイル高さ
'Dim PI As Double 'ﾋﾟﾝ
'Dim TIP As Double 'ﾁｯﾌﾟﾎﾞｰﾙ
'Dim SUT As Double 'ｽﾄﾚｯﾁﾌｨﾙﾑ
'Dim ESA As Double 'ｴｻﾌｫ-ﾑ
'Dim POR As Double 'ポリシート
'Dim HB As Double 'ハードボード
'Dim PET As Double 'バンド
'Dim PAR As Double 'パレット
'Dim NW As Double 'NW
'Dim HU As Double '風袋
'Dim GW As Double 'GW
'Dim i As Long, Q As Long
'Dim n As String
'n = ThisWorkbook.Sheets("設定").Cells(Rows.count, 1).End(xlUp).Row
'
' '*************************風袋シートの作成
'  Dim ListName As String, SheetsName As String
'  Dim ws As Worksheet, Flag As Boolean
'
'    ListName = "50まで"
'    SheetsName = ListName
'
'    For Each ws In ThisWorkbook.Worksheets
'        If ws.name = SheetsName Then Flag = True
'    Next ws
' Application.ScreenUpdating = False
'    If Flag = True Then    '同じ名前のシートあれば消す
'        Application.DisplayAlerts = False
'        ThisWorkbook.Sheets(SheetsName).Delete
'        Application.DisplayAlerts = True
'    End If
'
'    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(Sheets.count) '新たにシートを追加する
'    ActiveSheet.name = ListName
'    ThisWorkbook.Sheets(ListName).Visible = False
' '*************************ラベル設定
' With ThisWorkbook.Sheets(ListName)
'   .Cells(1, 2) = "ﾋﾟﾝ" & Chr(10) & "(kg)"
'   .Cells(1, 3) = "ﾁｯﾌﾟﾎﾞｰﾙ" & Chr(10) & "(kg)"
'   .Cells(1, 4) = "ｽﾄﾚｯﾁﾌｨﾙﾑ" & Chr(10) & "(kg)"
'   .Cells(1, 5) = "ｴｻﾌｫｰﾑ" & Chr(10) & "(kg)"
'   .Cells(1, 6) = "ﾎﾟﾘｼｰﾄ" & Chr(10) & "(kg)"
'   .Cells(1, 7) = "ﾊｰﾄﾞﾎﾞｰﾄﾞ" & Chr(10) & "(kg)"
'   .Cells(1, 8) = "樹脂ﾊﾟﾚｯﾄ" & Chr(10) & "(kg)"
'   .Cells(1, 9) = "ﾍﾟｯﾄﾊﾞﾝﾄﾞ" & Chr(10) & "(kg)"
'   .Cells(1, 10) = "NW" & Chr(10) & "(kg)"
'   .Cells(1, 11) = "風袋" & Chr(10) & "(kg)"
'   .Cells(1, 12) = "GW" & Chr(10) & "(kg)"
'   .Cells(1, 13) = "ｺｲﾙ高さ" & Chr(10) & "(mm)"
'   .Cells(1, 14) = "梱包高さ" & Chr(10) & "(mm)"
'   .Cells(1, 15) = "ﾁｯﾌﾟﾎﾞｰﾙ高さ" & Chr(10) & "(mm)"
'   .Cells(1, 1) = "本数"
' End With
' '*************************ｽﾀｰﾄ
' For Q = 1 To 50
'    '~~~~~~~~~~~~~~~~~~~~~~~~設定値変数~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'     For i = 3 To n 'n=15
'         W1(i - 2) = ThisWorkbook.Sheets("設定").Cells(i, 2)
'         K1(i - 2) = ThisWorkbook.Sheets("設定").Cells(i, 3)
'     Next i
'
'    With テスラ指定サイズ
'       If .CheckBox1 = True Then
'         col = 40
'       ElseIf .CheckBox2 = True Then
'         col = 33
'       ElseIf .CheckBox3 = True Then
'         col = 73
'       ElseIf .CheckBox4 = True Then
'         col = 82.5
'       ElseIf .CheckBox5 = True Then
'         col = 53.5
'       ElseIf .CheckBox6 = True Then
'         col = 63
'       End If
'
'       If .OptionButton7 = True Or .OptionButton8 = True Then
'         col = 40
'       ElseIf .OptionButton9 = True Or .OptionButton10 = True Then
'         col = 33
'       ElseIf .OptionButton1 = True Or .OptionButton2 = True Then
'         col = 73
'       ElseIf .OptionButton3 = True Or .OptionButton4 = True Then
'         col = 82.5
'       ElseIf .OptionButton5 = True Or .OptionButton6 = True Then
'         col = 53.5
'       ElseIf .OptionButton11 = True Or .OptionButton12 = True Then
'         col = 63
'       End If
'       '高さ
'       Dim TIPT As Double 'ﾁｯﾌﾟﾎﾞｰﾙの高さ　1枚0.7mm
'       TIPT = Val((Q + 1) * 0.7 / 10 ^ 3)
'       TA = Val(Q * col / 10 ^ 3) + TIPT 'ﾃｷｽﾄﾎﾞｯｸｽの数字×ﾁｪｯｸの入っているｺｲﾙ 'M単位
'    End With
'  '*****************************************************************************計算処理
' With ThisWorkbook.Sheets(ListName)
'      .Cells(Q + 1, 1) = Q
'      .Cells(Q + 1, 2) = Val(4 * W1(1) * K1(1))
'      .Cells(Q + 1, 3) = Val((Q + 1) * W1(2) * K1(2))
'    Dim GAI As Double
'    GAI = Val(1100 * 3.141592 / 10 ^ 3) 'M単位
'      .Cells(Q + 1, 4) = Val(GAI * 5 * W1(3) * K1(3))  '5周は巻く
'      If TA > 0.82 Then 'エサフォームの高さが550　1.5倍ほどの高さになった時2周とする
'       .Cells(Q + 1, 5) = Val(GAI * 2 * W1(4) * K1(4)) '2周
'      Else
'       .Cells(Q + 1, 5) = Val(GAI * W1(4) * K1(4)) '1周
'      End If
'    Dim AREA As Double '円部分
'    Dim AREA2 As Double '同部分
'      AREA = Val(550 ^ 2 * 3.141592)
'      AREA2 = Val((TA + TA / 2) * GAI) '高さﾌﾟﾗｽ分で包む　ﾌﾟﾗｽ分は高さの半分とする
'      .Cells(Q + 1, 6) = Val((AREA * 2 + AREA2) / 10 ^ 6 * W1(6) * K1(6))
' End With
'
'    With テスラ指定サイズ
'       If .CheckBox1 = True Then
'         Select Case Q
'         Case 14
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'         Case Else
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'         End Select
'       ElseIf .CheckBox2 = True Then
'         Select Case Q
'         Case 15
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10)  '2枚
'         Case Else
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'         End Select
'       ElseIf .CheckBox3 = True Then
'         Select Case Q
'         Case 8
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'         Case 7
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'         Case Else
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'         End Select
'       ElseIf .CheckBox4 = True Then
'         Select Case Q
'         Case 7
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'         Case 6
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'         Case Else
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'         End Select
'       ElseIf .CheckBox5 = True Then
'         Select Case Q
'         Case 10
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(17) * K1(17) + 2 * W1(10) * K1(10) '2枚
'         Case Else
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'         End Select
'       ElseIf .CheckBox6 = True Then
'         Select Case Q
'         Case 9
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'         Case 8
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'         Case Else
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = (GAI + 0.4) * TA * 2 * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'         End Select
'       End If
'
'       If .OptionButton7 = True Or .OptionButton8 = True Then
'         Select Case Q
'         Case 14
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'         Case Else
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'         End Select
'       ElseIf .OptionButton9 = True Or .OptionButton10 = True Then
'         Select Case Q
'         Case 15
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'         Case Else
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'         End Select
'       ElseIf .OptionButton1 = True Or .OptionButton2 = True Then
'         Select Case Q
'         Case 8
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'         Case 7
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'         Case Else
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = (GAI + 0.4) * TA * W1(7) * K1(7)
'         End Select
'       ElseIf .OptionButton3 = True Or .OptionButton4 = True Then
'         Select Case Q
'         Case 7
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'         Case 6
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'         Case Else
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'         End Select
'       ElseIf .OptionButton5 = True Or .OptionButton6 = True Then
'         Select Case Q
'         Case 10
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(17) * K1(17) + 2 * W1(10) * K1(10) '2枚
'         Case Else
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'         End Select
'       ElseIf .OptionButton11 = True Or .OptionButton12 = True Then
'         Select Case Q
'         Case 9
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(9) * K1(9) + 2 * W1(10) * K1(10) '2枚
'         Case 8
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = 2 * W1(8) * K1(8) + 2 * W1(10) * K1(10) '2枚
'         Case Else
'          ThisWorkbook.Sheets(ListName).Cells(Q + 1, 7) = (GAI + 0.4) * TA * W1(7) * K1(7) + 2 * W1(10) * K1(10)
'         End Select
'       End If
'    End With
'    With ThisWorkbook.Sheets(ListName)
'       .Cells(Q + 1, 8) = Val(38)  '2台分
'      '胴巻き4本
'      Dim DOU As Double
'      DOU = Val((GAI + 0.2) * 4) '締める時の余分を200mmとした
'      '縦バンド4本
'      Dim TATE As Double
'      TATE = Val((((TA + 0.2) * 4) * 2) + ((1100 * 4) * 2) / 10 ^ 3) '締める時の余分を200mmとした
'        .Cells(Q + 1, 9) = (DOU + TATE) * W1(5) * K1(5)
'      Dim T1 As Double
'      Dim T2 As Double
'      '入力本数は関係なく1~50まで表示させる
'        If テスラ指定サイズ.lblWeight1 <> "" And テスラ指定サイズ.lblWeight2 <> "" Then
'         T1 = Replace(テスラ指定サイズ.lblWeight1, "kg", "")
'         T2 = Replace(テスラ指定サイズ.lblWeight2, "kg", "")
'        ElseIf テスラ指定サイズ.lblWeight1 <> "" Then
'         T1 = Replace(テスラ指定サイズ.lblWeight1, "kg", "")
'        ElseIf テスラ指定サイズ.lblWeight2 <> "" Then
'         T2 = Replace(テスラ指定サイズ.lblWeight2, "kg", "")
'        End If
'      'NW
'       If テスラ指定サイズ.Take1.Visible = True And テスラ指定サイズ.Take2.Visible = True Then
'        .Cells(Q + 1, 10) = Format((Val(Q * T1)), "0.0")
'       ElseIf テスラ指定サイズ.Take1.Visible = True Then
'        .Cells(Q + 1, 10) = Format((Val(Q * T1)), "0.0")
'       ElseIf テスラ指定サイズ.Take2.Visible = True Then
'        .Cells(Q + 1, 10) = Format((Val(Q * T2)), "0.0")
'       End If
'      '*************************************************************風袋.GW
'        .Cells(Q + 1, 11) = Format(Application.WorksheetFunction.Round((.Cells(Q + 1, 2) + .Cells(Q + 1, 3) + .Cells(Q + 1, 4) + .Cells(Q + 1, 5) + .Cells(Q + 1, 6) + .Cells(Q + 1, 7) + .Cells(Q + 1, 8) + .Cells(Q + 1, 9)), 0), "0.0") '四捨五入繰り上げ
'        .Cells(Q + 1, 12) = Format(Application.WorksheetFunction.Round((.Cells(Q + 1, 2) + .Cells(Q + 1, 3) + .Cells(Q + 1, 4) + .Cells(Q + 1, 5) + .Cells(Q + 1, 6) + .Cells(Q + 1, 7) + .Cells(Q + 1, 8) + .Cells(Q + 1, 9) + .Cells(Q + 1, 10)), 0), "0.0") '四捨五入繰り上げ
'      '*************************************************************ｺｲﾙ高さ
'         .Cells(Q + 1, 13) = Format((Val(TA * 1000)))
'      '*************************************************************梱包高さ　1桁繰り上げ
'         .Cells(Q + 1, 14) = Format(Application.WorksheetFunction.RoundUp((Val((TA * 1000) + 150 + 150)), -1), "0.0") '(Val((TA * 1000) + 150 + 150))
'      '*************************************************************ﾁｯﾌﾟ高さ
'         .Cells(Q + 1, 15) = Format(Val(TIPT * 1000))
'
'       .Range("D2:G53,I2:I53").NumberFormatLocal = "#,##0.000_ "
'       .Range("H2:H53,J2:O53").NumberFormatLocal = "0.0_ "
'   End With
'  Next Q
'    ThisWorkbook.Sheets(ListName).Activate
'    Cells.Select
'    With Selection.Font
'        .name = "メイリオ"
'        .size = 10.5
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ThemeColor = xlThemeColorLight1
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With
'    ThisWorkbook.Sheets(ListName).Cells.EntireColumn.AutoFit
'    '*****************************************************罫線
'    Range("A1:O51").Select
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    With Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .weight = xlThin
'    End With
'    With Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .weight = xlThin
'    End With
'    With Selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .weight = xlThin
'    End With
'    With Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .weight = xlThin
'    End With
'    With Selection.Borders(xlInsideVertical)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .weight = xlThin
'    End With
'    With Selection.Borders(xlInsideHorizontal)
'        .LineStyle = xlContinuous
'        .ColorIndex = 0
'        .TintAndShade = 0
'        .weight = xlThin
'    End With
'    '*****************************************************印刷設定
'    Application.PrintCommunication = False
'    With ActiveSheet.PageSetup
'        .PrintTitleRows = ""
'        .PrintTitleColumns = ""
'    End With
'    Application.PrintCommunication = True
'    ActiveSheet.PageSetup.PrintArea = "A1:O51"
'    Application.PrintCommunication = False
'    With ActiveSheet.PageSetup
'        .LeftHeader = ""
'        .CenterHeader = ""
'        .RightHeader = ""
'        If テスラ指定サイズ.lblSize1 <> "" And テスラ指定サイズ.lblSize2 <> "" Then
'         .LeftFooter = "＊検査番号：" & テスラ指定サイズ.lblKensaNo & "＊" & "ｺｲﾙｻｲｽﾞ：" & テスラ指定サイズ.lblSize1 & " " & テスラ指定サイズ.lblSize2
'        ElseIf テスラ指定サイズ.lblSize1 <> "" And テスラ指定サイズ.lblSize2 = "" Then
'         .LeftFooter = "＊検査番号：" & テスラ指定サイズ.lblKensaNo & "＊" & "ｺｲﾙｻｲｽﾞ：" & テスラ指定サイズ.lblSize1
'        ElseIf テスラ指定サイズ.lblSize1 = "" And テスラ指定サイズ.lblSize2 <> "" Then
'         .LeftFooter = "＊検査番号：" & テスラ指定サイズ.lblKensaNo & "＊" & "ｺｲﾙｻｲｽﾞ：" & テスラ指定サイズ.lblSize2
'        End If
'        .CenterFooter = ""
'        .RightFooter = Format$(Date, "yyyy.mm.dd") & "印刷"
'        .leftMargin = Application.InchesToPoints(0)
'        .RightMargin = Application.InchesToPoints(0)
'        .TopMargin = Application.InchesToPoints(0)
'        .BottomMargin = Application.InchesToPoints(0)
'        .HeaderMargin = Application.InchesToPoints(0)
'        .FooterMargin = Application.InchesToPoints(0)
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        .CenterHorizontally = True
'        .CenterVertically = False
'        .PaperSize = xlPaperA4
'        .FirstPageNumber = xlAutomatic
'
'        .orientation = xlPortrait '縦　xlPortrait　横 xllandscape
'        .FitToPagesWide = 1 '1ページで納める
'        .FitToPagesTall = 1
'
'    End With
'    Application.PrintCommunication = True
'    ActiveWindow.View = xlPageBreakPreview
'
' Application.ScreenUpdating = True
'End Sub

