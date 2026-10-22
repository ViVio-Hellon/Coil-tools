Attribute VB_Name = "テスラ補助"
Option Explicit
'Windows API宣言
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000
#If Win64 Then
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
#Else
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
#End If

'--------------------------------------------------------------
' 定数: ラベル内オフセット (ベース行からの相対位置)
'--------------------------------------------------------------
Public Const OFS_BC_KATA  As Long = -1   ' バーコード(型番)
Public Const OFS_COIL     As Long = 0    ' コイル番号
Public Const OFS_KENSA    As Long = 0    ' 検査番号
Public Const OFS_KATABAN  As Long = 1    ' 型番
Public Const OFS_WEIGHT   As Long = 2    ' 重量
Public Const OFS_BC_KEN   As Long = 3    ' バーコード(検番+コイル+重量)

'--------------------------------------------------------------
' 定数: 左面・右面の列番号
'--------------------------------------------------------------
Public Const COL_COIL_L    As Long = 7:   Public Const COL_COIL_R    As Long = 39 ' コイル番号
Public Const COL_KEN_L     As Long = 4:   Public Const COL_KEN_R     As Long = 36 '検査番号
Public Const COL_KATA_L    As Long = 4:   Public Const COL_KATA_R    As Long = 36 '型番
Public Const COL_WT_L      As Long = 4:   Public Const COL_WT_R      As Long = 36 '重量
Public Const COL_BC_KATA_L As Long = 9:   Public Const COL_BC_KATA_R As Long = 41 'バーコード(型番)
Public Const COL_BC_KEN_L  As Long = 2:   Public Const COL_BC_KEN_R  As Long = 34 'バーコード(検番等)
'--------------------------------------------------------------
' 定数: ヘッダー部の座標
'--------------------------------------------------------------
Public Const HDR_KEN_ROW   As Long = 3:   Public Const HDR_KEN_COL   As Long = 22
Public Const HDR_KEN2_ROW  As Long = 92:  Public Const HDR_KEN2_COL  As Long = 22
Public Const HDR_WT1_ROW   As Long = 3:   Public Const HDR_WT1_COL   As Long = 37
Public Const HDR_WT2_ROW   As Long = 92:  Public Const HDR_WT2_COL   As Long = 37


'例: 1.2mm×50.0mm を OB15/16、CB ck_12_50 として追加する場合
'modTeslaConfig
'1～12OptionButton1 ～ OptionButton12コントロール名の番号そのまま
'13 lblSpec_08_53_112の次 (仮想的なOBインデックス)
'名前付きコントロールは Controls("OptionButton" & i) のループで取れないので、DetectSelectedOB で個別判定した後に「内部管理用の番号」として13/14を返し、GetSizeConfig に渡す仕組み
'' ① GetSizeConfig に Case 追加
'        Case 13, 14: GetSizeConfig = Array("0.8mm×53.5mm", "BJB7610400QR", "g", 3, 92)
'        Case 15, 16: GetSizeConfig = Array("1.2mm×50.0mm", "BJBxxxxxxQR", "h", 3, 0)    ' ★追加
'        Case Else:   GetSizeConfig = Array("", "", "", 0, 0)
'
'' ② GetCBtoPair に Case 追加 (CBモードが必要な場合)
'        Case 7:  ob1 = 15: ob2 = 16   ' 1.2mm×50.0mm  ★追加
'
'' ③ DetectSelectedOB に名前付きOB追加
'    If frm.lblSpec_12_50_1.Value Then DetectSelectedOB = 15: Exit Function
'    If frm.lblSpec_12_50_2.Value Then DetectSelectedOB = 16: Exit Function
'テスラ指定サイズ UserForm
'' ④ OB Click イベント追加
'Private Sub lblSpec_12_50_1_Click()
'    SelectSize "1.2mm×50.0mm 　丈1", 3, 1
'    ' ClearCoilInfo が必要なら追加
'End Sub
'
'Private Sub lblSpec_12_50_2_Click()
'    SelectSize "1.2mm×50.0mm 　丈2", 3, 2
'End Sub
'
'' ⑤ CB Click イベント追加
'Private Sub ck_12_50_Click()
'    If Me.ck_12_50.Value Then SelectCheckBox 0, isNamed:=True
'End Sub

'' ⑥ SelectCheckBox 内の名前付きCB分岐を拡張
'    If isNamed Then
'        If Me.ck_08_53.Value Then ob1 = 13: ob2 = 14
'        If Me.ck_12_50.Value Then ob1 = 15: ob2 = 16   ' ★追加
'    Else
'        GetCBtoPair cbIdx, ob1, ob2
'    End If
'
'' ⑦ 同様に チェック入力、チェック印刷、InpCheck、IsAnySizeSelected 内の
''    名前付きコントロール判定箇所に追加
'' ⑧ CommandButton3/4 のループ上限
'    For obIdx = 1 To 16 Step 2   ' 14 → 16
'' ⑨ CB側・OB側の Select Case に追加
'      Case .ck_12_50:  mat.col = 50#
'      Case .lblSpec_12_50_1 Or .lblSpec_12_50_2: mat.col = 50#
'

Public Sub Auto_Open()
ThisWorkbook.Sheets("起動").Select
Call フォーム1起動
End Sub
Public Sub フォーム1起動()
テスラ指定サイズ.Show vbModeless
End Sub
Public Sub フォーム4起動()
UserForm4.Show vbModeless
End Sub

'--------------------------------------------------------------
' ラベル16枚のベース行 (検査番号行 = オフセット0)
'--------------------------------------------------------------
Public Function GetLabelBaseRows() As Variant
    GetLabelBaseRows = Array( _
        8, 18, 28, 38, 48, 57, 66, 75, _
        97, 106, 115, 124, 133, 142, 151, 160)
End Function

'--------------------------------------------------------------
' シート名組立
'--------------------------------------------------------------
Public Function BuildSheetName(ByVal obIdx As Long) As String
    Dim cfg As Variant: cfg = GetSizeConfig(obIdx)
    BuildSheetName = cfg(0) & " 　丈" & GetTakeType(obIdx)
    DebugLog "[BuildSheetName] obIdx=" & obIdx & " → """ & BuildSheetName & """"
End Function

'==============================================================================
' modTeslaWrite  ― コイル番号自動生成版
' メイン: シートへ全座標一括転記
'書込先シート,検査番号,重量,型番,コイル番号の先頭2桁に使用
'==============================================================================
Public Sub WriteAllToSheet(ByVal ws As Worksheet, _
                           ByVal ken As String, _
                           ByVal wt As String, _
                           ByVal kataban As String, _
                           ByVal takeType As Long)
    Dim baseRows As Variant
    baseRows = GetLabelBaseRows()
    
'    ' WriteAllToSheet 冒頭に追加
'    DebugLog "[WriteAllToSheet] ken=" & ken & " wt=" & wt & " kataban=" & kataban & " takeType=" & takeType


    Dim bcKata As String
    bcKata = "*" & kataban & "*"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo Cleanup

    '--- ヘッダー部 ---
    ws.Cells(HDR_KEN_ROW, HDR_KEN_COL) = ken
'    DebugLog "[WriteAllToSheet] Header(3,22) OK"
    
    ws.Cells(HDR_KEN2_ROW, HDR_KEN2_COL) = ken
'    DebugLog "[WriteAllToSheet] Header(92,22) OK"
    
    ws.Cells(HDR_WT1_ROW, HDR_WT1_COL) = wt
'    DebugLog "[WriteAllToSheet] Header(3,37) OK"
    
    ws.Cells(HDR_WT2_ROW, HDR_WT2_COL) = wt
'    DebugLog "[WriteAllToSheet] Header(9,37) OK"

    '--- ラベル 16枚 × 左右2面 ---
    Dim i As Long, r As Long
    Dim seqL As Long, seqR As Long
    Dim coilL As String, coilR As String

    For i = 0 To UBound(baseRows)
        r = baseRows(i)
        DebugLog "[WriteAllToSheet] Loop i=" & i & " r=" & r
        '--- コイル番号連番 ---
        If i < 8 Then
            seqL = i + 1:  seqR = i + 9
        Else
            seqL = i + 9:  seqR = i + 17
        End If

        coilL = "-" & Format(takeType, "00") & Format(seqL, "00")
        coilR = "-" & Format(takeType, "00") & Format(seqR, "00")

        ' コイル番号 (base+0, col 7/39)
        ws.Cells(r + OFS_COIL, COL_COIL_L) = coilL
        ws.Cells(r + OFS_COIL, COL_COIL_R) = coilR

        ' 検査番号 (base+0, col 4/36)
        ws.Cells(r + OFS_COIL, COL_COIL_L).NumberFormat = "@"
        ws.Cells(r + OFS_COIL, COL_COIL_L).Value = coilL
        ws.Cells(r + OFS_COIL, COL_COIL_R).NumberFormat = "@"
        ws.Cells(r + OFS_COIL, COL_COIL_R).Value = coilR
        
        ' 型番 (base+1, col 4/36)
        ws.Cells(r + OFS_KATABAN, COL_KATA_L) = kataban
        ws.Cells(r + OFS_KATABAN, COL_KATA_R) = kataban

        ' 重量 (base+2, col 4/36)
        ws.Cells(r + OFS_WEIGHT, COL_WT_L) = wt
        ws.Cells(r + OFS_WEIGHT, COL_WT_L).NumberFormatLocal = "0.0_ "
        ws.Cells(r + OFS_WEIGHT, COL_WT_R) = wt
        ws.Cells(r + OFS_WEIGHT, COL_WT_R).NumberFormatLocal = "0.0_ "
        
        ' 重量単位 (base+2, col 7/39)
        '7列目/39列目はコイル番号列（COL_COIL_L/COL_COIL_R）と同じ列番号だが、行が違う（コイル番号はbase+0）ので干渉しない
        ws.Cells(r + OFS_WEIGHT, COL_COIL_L) = "kg"
        ws.Cells(r + OFS_WEIGHT, COL_COIL_R) = "kg"

        ' バーコード(型番) (base-1, col 9/41)
        ws.Cells(r + OFS_BC_KATA, COL_BC_KATA_L) = bcKata
        ws.Cells(r + OFS_BC_KATA, COL_BC_KATA_R) = bcKata

        ' バーコード(検番+コイル+重量) (base+3, col 2/34)
        ws.Cells(r + OFS_BC_KEN, COL_BC_KEN_L) = "*" & ken & coilL & " " & wt & "*"
        ws.Cells(r + OFS_BC_KEN, COL_BC_KEN_R) = "*" & ken & coilR & " " & wt & "*"
    Next i

Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        DebugLog "[WriteAllToSheet] Error " & Err.Number & ": " & Err.Description
'        MsgBox "WriteAllToSheet でエラー: " & Err.Description, vbExclamation   ' ★一時的に追加
        Err.Clear
    End If
End Sub

'--------------------------------------------------------------
' 選択中の OptionButton を検出し、設定Index を返す
' 見つからなければ 0 を返す
'--------------------------------------------------------------
Public Function DetectSelectedOB(ByVal frm As MSForms.UserForm) As Long
    Dim i As Long

    For i = 1 To 12
        If frm.Controls("OptionButton" & i).Value Then
            DetectSelectedOB = i
            DebugLog "[DetectSelectedOB] OptionButton" & i
            Exit Function
        End If
    Next i

    If frm.lblSpec_08_53_1.Value Then
        DetectSelectedOB = 13
        DebugLog "[DetectSelectedOB] lblSpec_08_53_1 → 13"
        Exit Function
    End If
    If frm.lblSpec_08_53_2.Value Then
        DetectSelectedOB = 14
        DebugLog "[DetectSelectedOB] lblSpec_08_53_2 → 14"
        Exit Function
    End If

    DetectSelectedOB = 0
    DebugLog "[DetectSelectedOB] 未選択 → 0"
End Function

'--------------------------------------------------------------
' CheckBox → OBペア (ByRef で返す)
'--------------------------------------------------------------
Public Sub GetCBtoPair(ByVal cbIdx As Long, _
                       ByRef ob1 As Long, _
                       ByRef ob2 As Long)
    Select Case cbIdx
        Case 1:  ob1 = 7:  ob2 = 8    ' 1.0mm×40.0mm
        Case 2:  ob1 = 9:  ob2 = 10   ' 1.0mm×33.0mm
        Case 3:  ob1 = 1:  ob2 = 2    ' 1.0mm×73mm
        Case 4:  ob1 = 3:  ob2 = 4    ' 0.6mm×82.5mm
        Case 5:  ob1 = 5:  ob2 = 6    ' 1.0mm×53.5mm
        Case 6:  ob1 = 11: ob2 = 12   ' 1.0mm×63.0mm
        Case Else: ob1 = 0: ob2 = 0
    End Select
End Sub

' --- マスタ補足 (未使用OB用/将来拡張) ---
' 0.8×53.5×Coil  g  BJB7610400QR
'--------------------------------------------------------------
' 丈タイプ判定 (奇数=丈1, 偶数=丈2)
'--------------------------------------------------------------
Public Function GetTakeType(ByVal obIdx As Long) As Long
    GetTakeType = IIf(obIdx Mod 2 = 1, 1, 2)
End Function

'2019.04.18
'0.6mm×43.5mm 　丈1 0.6mm×43.5mm 　丈2 削除
'1.0mm×53.5mm 　丈1 1.0mm×53.5mm 　丈2 追加(入れ替え)
Sub ボタン1_Click()
テスラ指定サイズ.Show vbModeless
End Sub

Public Function InpCheck(a As Long) As Boolean
    Dim i As Long
    With テスラ指定サイズ

        Select Case a
        Case 1  '丈1
            For i = 1 To 11 Step 2
                If .Controls("OptionButton" & i) Then
                    If .txtWeight1 = "" Then InpCheck = True
                    Exit For
                End If
            Next i
            '--- 名前付き OB (丈1) ---
            If .lblSpec_08_53_1.Value Then
                If .txtWeight1 = "" Then InpCheck = True
            End If

        Case 2  '丈2
            For i = 2 To 12 Step 2
                If .Controls("OptionButton" & i) Then
                    If .txtWeight2 = "" Then InpCheck = True
                    Exit For
                End If
            Next i
            '--- 名前付き OB (丈2) ---
            If .lblSpec_08_53_2.Value Then
                If .txtWeight2 = "" Then InpCheck = True
            End If

        Case 3  '丈1,2
            For i = 1 To 6
                If .Controls("CheckBox" & i) Then
                    If .txtWeight1 = "" Or .txtWeight2 = "" Then InpCheck = True
                    Exit For
                End If
            Next i
            '--- 名前付き CB ---
            If .ck_08_53.Value Then
                If .txtWeight1 = "" Or .txtWeight2 = "" Then InpCheck = True
            End If

        End Select
    End With
End Function

'==============================================================================
' チェック入力  ― WriteAllToSheet 呼び出し追加版
'==============================================================================
Public Sub チェック入力()
    Dim ken As String
    ken = StrConv(UCase(テスラ指定サイズ.txtKensaNo), vbNarrow)

    Dim cbIdx As Long, Found As Boolean
    Dim ob1 As Long, ob2 As Long

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

    If Not Found Then Exit Sub

    DebugLog "[チェック入力] ob1=" & ob1 & " ob2=" & ob2

    '--- 丈1 ---
    Dim cfg1 As Variant:    cfg1 = GetSizeConfig(ob1)
    Dim dRow1 As Long:      dRow1 = CLng(cfg1(3))
    Dim wt1 As String:      wt1 = テスラ指定サイズ.txtWeight1
    Dim kataban1 As String: kataban1 = CStr(cfg1(1))
    Dim coilNo1 As String:  coilNo1 = CStr(cfg1(2))

    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Sheets(BuildSheetName(ob1))
    ws1.Activate
    Call WriteAllToSheet(ws1, ken, wt1, kataban1, 1)

    テスラ指定サイズ.lblKensaNo = ws1.Cells(dRow1, HDR_KEN_COL)
    テスラ指定サイズ.lblWeight1 = ws1.Cells(dRow1, HDR_WT1_COL) & "kg"
    テスラ指定サイズ.lblSize1 = ws1.Cells(dRow1, 2)

    '--- 丈2 ---
    Dim cfg2 As Variant:    cfg2 = GetSizeConfig(ob2)
    Dim dRow2 As Long:      dRow2 = CLng(cfg2(3))
    Dim wt2 As String:      wt2 = テスラ指定サイズ.txtWeight2
    Dim kataban2 As String: kataban2 = CStr(cfg2(1))
    Dim coilNo2 As String:  coilNo2 = CStr(cfg2(2))

    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets(BuildSheetName(ob2))
    ws2.Activate
    Call WriteAllToSheet(ws2, ken, wt2, kataban2, 2)

    テスラ指定サイズ.lblWeight2 = ws2.Cells(dRow2, HDR_WT1_COL) & "kg"
    テスラ指定サイズ.lblSize2 = ws2.Cells(dRow2, 2)

    ThisWorkbook.Sheets("起動").Activate
End Sub

'==============================================================================
' チェック印刷  ―  リファクタリング版
'==============================================================================
Public Sub チェック印刷()
    Dim cbIdx As Long, Found As Boolean
    Dim ob1 As Long, ob2 As Long

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

    If Not Found Then Exit Sub

    PrintSheet BuildSheetName(ob1)
    PrintSheet BuildSheetName(ob2)
    ThisWorkbook.Sheets("起動").Activate
End Sub

'--------------------------------------------------------------
' 1シート印刷 (表示→印刷→非表示)
'--------------------------------------------------------------
Private Sub PrintSheet(ByVal shtName As String)
    With ThisWorkbook.Sheets(shtName)
        .Visible = True
        .Activate
        .PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        .Visible = False
    End With
End Sub

'--------------------------------------------------------------
' サイズ構成マスタ
'--------------------------------------------------------------
' 戻り値: Array(シート名ベース, 型番, コイル記号, データ行, 重量追加行)
'cfg(0)"1.0mm×73mm"  cfg(1)"BJB7604000QR"  cfg(2)"c"  cfg(3)3 データ行（ラベル表示用の読取行）cfg(4)0重量追加行（0=なし、92=92行目にも書込）
'--------------------------------------------------------------
Public Function GetSizeConfig(ByVal obIdx As Long) As Variant
    Select Case obIdx
        'シート名ベース,型番 ,記号,行,追加行
        Case 1, 2:   GetSizeConfig = Array("1.0mm×73mm", "BJB7604000QR", "c", 3, 0)
        Case 3, 4:   GetSizeConfig = Array("0.6mm×82.5mm", "BJB7604200QR", "e", 3, 0)
        Case 5, 6:   GetSizeConfig = Array("1.0mm×53.5mm", "BJB7606500QR", "d", 3, 92)
        Case 7, 8:   GetSizeConfig = Array("1.0mm×40.0mm", "BJB7604300QR", "b", 3, 92)
        Case 9, 10:  GetSizeConfig = Array("1.0mm×33.0mm", "BJB7603900QR", "a", 3, 92)
        Case 11, 12: GetSizeConfig = Array("1.0mm×63.0mm", "BJB7605900QR", "f", 92, 0)
        Case 13, 14: GetSizeConfig = Array("0.8mm×53.5mm", "BJB7610400QR", "g", 3, 92)
        Case Else:   GetSizeConfig = Array("", "", "", 0, 0)
    End Select
End Function

Public Function GetSizeDisplayName(ByVal obIdx As Long) As String
    Dim cfg As Variant: cfg = GetSizeConfig(obIdx)
    GetSizeDisplayName = CStr(cfg(0))   ' "1.0mm×73mm" 等
End Function

'--------------------------------------------------------------
' OBモード: 選択中OBの実シート名を返す (印刷・転記用)
' 未選択なら "" を返す
'--------------------------------------------------------------
Public Function GetSelectedOBSheetName() As String
    Dim obIdx As Long
    obIdx = DetectSelectedOB(テスラ指定サイズ) '引数フォーム
    If obIdx > 0 Then
        GetSelectedOBSheetName = BuildSheetName(obIdx)
    Else
        GetSelectedOBSheetName = ""
    End If
End Function

'--------------------------------------------------------------
' CBモード: 選択中CBの表示名を返す (MsgBox確認用)
' 未選択なら "" を返す
'--------------------------------------------------------------
Public Function GetSelectedCBDisplayName() As String
    Dim cbIdx As Long
    Dim ob1 As Long, ob2 As Long

    For cbIdx = 1 To 6
        If テスラ指定サイズ.Controls("CheckBox" & cbIdx).Value Then
            GetCBtoPair cbIdx, ob1, ob2
            GetSelectedCBDisplayName = GetSizeDisplayName(ob1) & " 　丈1,2"
            Exit Function
        End If
    Next cbIdx

    If テスラ指定サイズ.ck_08_53.Value Then
        GetSelectedCBDisplayName = "0.8mm×53.5mm 　丈1,2"
        Exit Function
    End If

    GetSelectedCBDisplayName = ""
End Function

Public Sub シート追加1()
Dim ListName As String
Dim SheetsName As String
ListName = (テスラ全サイズ.ComboBox1.text & "1") 'コンボボックスから選んだものをシート名にする
SheetsName = ListName

Dim ws As Worksheet, Flag As Boolean

For Each ws In Worksheets
    If ws.name = SheetsName Then Flag = True
Next ws

If Flag = True Then    '同じ名前のシートあれば消す
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SheetsName).Delete
    Application.DisplayAlerts = True
End If

ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(Sheets.count) '新たにシートを追加する
ActiveSheet.name = ListName 'コンボボックスから選んだものをシート名にする
End Sub
Public Sub シート追加2()
Dim ListName2 As String
Dim SheetsName As String
ListName2 = (テスラ全サイズ.ComboBox1.text & "2") 'コンボボックスから選んだものをシート名にする
SheetsName = ListName2

Dim ws As Worksheet, Flag As Boolean

For Each ws In Worksheets
    If ws.name = SheetsName Then Flag = True
Next ws

If Flag = True Then    '同じ名前のシートあれば消す
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SheetsName).Delete
    Application.DisplayAlerts = True
End If

ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(Sheets.count) '新たにシートを追加する
ActiveSheet.name = ListName2 'コンボボックスから選んだものをシート名にする
End Sub
Public Sub コピー1()
ThisWorkbook.Sheets("原本").Visible = True
ThisWorkbook.Sheets("原本").Activate
Cells.Select
Selection.Copy
Dim ListName As String
ListName = (テスラ全サイズ.ComboBox1.text & "1")
Worksheets(ListName).Select
'Worksheets(Worksheets.Count).Select
Cells.Select
Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
    , SkipBlanks:=False, Transpose:=False
    
Call 行列
    
ThisWorkbook.Sheets("原本").Visible = False
Application.CutCopyMode = False
End Sub
Public Sub コピー2()
ThisWorkbook.Sheets("原本").Visible = True
ThisWorkbook.Sheets("原本").Activate
Cells.Select
Selection.Copy
Dim ListName2 As String
ListName2 = (テスラ全サイズ.ComboBox1.text & "2")
Worksheets(ListName2).Select
'Worksheets(Worksheets.Count).Select
Cells.Select
Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
  , SkipBlanks:=False, Transpose:=False
  
Call 行列2
  
ThisWorkbook.Sheets("原本").Visible = False
Application.CutCopyMode = False
End Sub

Public Sub 入力() '2018.05.18 ラベル２つ追加
  
Dim ListName As String
Dim ListName2 As String
ListName = (テスラ全サイズ.ComboBox1.text & "1")
ListName2 = (テスラ全サイズ.ComboBox1.text & "2")
    
テスラ全サイズ.Label19 = ""
テスラ全サイズ.Label21 = ""
テスラ全サイズ.Label16 = ""
テスラ全サイズ.Label18 = ""
テスラ全サイズ.Label17 = ""
テスラ全サイズ.Label22 = ""
テスラ全サイズ.Label26 = ""
テスラ全サイズ.Label27 = ""
テスラ全サイズ.Label28 = ""
テスラ全サイズ.Label29 = ""


Dim a As Boolean
Dim b As Boolean
Dim c As Boolean
Dim D As Boolean

If テスラ全サイズ.TextBox2 <> "" Then
a = True
ElseIf テスラ全サイズ.TextBox3 <> "" Then
b = True
ElseIf テスラ全サイズ.TextBox4 <> "" Then
c = True
ElseIf テスラ全サイズ.TextBox5 <> "" Then
D = True
End If

If テスラ全サイズ.CheckBox1 = True Then '2ケタ
With ThisWorkbook.Sheets(ListName)
  .Activate
  
  .Cells(3, 7) = "'-0101" '副番
  .Cells(13, 7) = "'-0102"
  .Cells(23, 7) = "'-0103"
  .Cells(33, 7) = "'-0104"
  .Cells(43, 7) = "'-0105"
  .Cells(53, 7) = "'-0106"
  .Cells(63, 7) = "'-0107"
  .Cells(73, 7) = "'-0108"
  .Cells(83, 7) = "'-0109" '2018.05.18追加
  
  .Cells(3, 39) = "'-0110"
  .Cells(13, 39) = "'-0111"
  .Cells(23, 39) = "'-0112"
  .Cells(33, 39) = "'-0113"
  .Cells(43, 39) = "'-0114"
  .Cells(53, 39) = "'-0115"
  .Cells(63, 39) = "'-0116"
  .Cells(73, 39) = "'-0117"
  .Cells(83, 39) = "'-0118" '2018.05.18追加
  
  .Cells(96, 7) = "'-0119" '副番
  .Cells(106, 7) = "'-0120"
  .Cells(116, 7) = "'-0121"
  .Cells(126, 7) = "'-0122"
  .Cells(136, 7) = "'-0123"
  .Cells(146, 7) = "'-0124"
  .Cells(156, 7) = "'-0125"
  .Cells(166, 7) = "'-0126"
  .Cells(176, 7) = "'-0127" '2018.05.18追加
  
  .Cells(96, 39) = "'-0128"
  .Cells(106, 39) = "'-0129"
  .Cells(116, 39) = "'-0130"
  .Cells(126, 39) = "'-0131"
  .Cells(136, 39) = "'-0132"
  .Cells(146, 39) = "'-0133"
  .Cells(156, 39) = "'-0134"
  .Cells(166, 39) = "'-0135"
  .Cells(176, 39) = "'-0136" '2018.05.18追加
  
  .Cells(92, 2) = ListName 'サイズ
  .Cells(185, 2) = ListName 'サイズ
  .Cells(92, 22) = テスラ全サイズ.TextBox1 '検査番号
  .Cells(185, 22) = テスラ全サイズ.TextBox1 '検査番号
  .Cells(92, 37) = テスラ全サイズ.TextBox2 '重量1
  .Cells(185, 37) = テスラ全サイズ.TextBox2 '重量1
テスラ全サイズ.Label16 = "「Coil1」" & Chr(10) & .Cells(92, 22) '検査番号
テスラ全サイズ.Label17 = .Cells(92, 37) & "kg"  '重量1
テスラ全サイズ.Label19 = "「Sheet1」" & Chr(10) & .Cells(92, 2) 'サイズ1

End With
With ThisWorkbook.Sheets(ListName2)
  .Activate
  
  .Cells(3, 7) = "'-0201" '副番
  .Cells(13, 7) = "'-0202"
  .Cells(23, 7) = "'-0203"
  .Cells(33, 7) = "'-0204"
  .Cells(43, 7) = "'-0205"
  .Cells(53, 7) = "'-0206"
  .Cells(63, 7) = "'-0207"
  .Cells(73, 7) = "'-0208"
  .Cells(83, 7) = "'-0209" '2018.05.18追加
  
  .Cells(3, 39) = "'-0210"
  .Cells(13, 39) = "'-0211"
  .Cells(23, 39) = "'-0212"
  .Cells(33, 39) = "'-0213"
  .Cells(43, 39) = "'-0214"
  .Cells(53, 39) = "'-0215"
  .Cells(63, 39) = "'-0216"
  .Cells(73, 39) = "'-0217"
  .Cells(83, 39) = "'-0218" '2018.05.18追加
  
  .Cells(96, 7) = "'-0219" '副番
  .Cells(106, 7) = "'-0220"
  .Cells(116, 7) = "'-0221"
  .Cells(126, 7) = "'-0222"
  .Cells(136, 7) = "'-0223"
  .Cells(146, 7) = "'-0224"
  .Cells(156, 7) = "'-0225"
  .Cells(166, 7) = "'-0226"
  .Cells(176, 7) = "'-0227" '2018.05.18追加
  
  .Cells(96, 39) = "'-0228"
  .Cells(106, 39) = "'-0229"
  .Cells(116, 39) = "'-0230"
  .Cells(126, 39) = "'-0231"
  .Cells(136, 39) = "'-0232"
  .Cells(146, 39) = "'-0233"
  .Cells(156, 39) = "'-0234"
  .Cells(166, 39) = "'-0235"
  .Cells(176, 39) = "'-0236" '2018.05.18追加
  
  .Cells(92, 2) = ListName2 'サイズ
  .Cells(185, 2) = ListName2 'サイズ
  .Cells(92, 22) = テスラ全サイズ.TextBox1 '検査番号
  .Cells(185, 22) = テスラ全サイズ.TextBox1 '検査番号
  .Cells(92, 37) = テスラ全サイズ.TextBox3 '重量1
  .Cells(185, 37) = テスラ全サイズ.TextBox3 '重量1
テスラ全サイズ.Label18 = "「Coil2」" & Chr(10) & .Cells(92, 22) '検査番号
テスラ全サイズ.Label22 = .Cells(92, 37) & "kg"  '重量2
テスラ全サイズ.Label21 = "「Sheet2」" & Chr(10) & .Cells(92, 2) 'サイズ2
End With

ElseIf テスラ全サイズ.CheckBox2 = True Then '副番3ケタ

With ThisWorkbook.Sheets(ListName)
  .Activate
  
  .Cells(3, 7) = "'-1101" '副番
  .Cells(13, 7) = "'-1102"
  .Cells(23, 7) = "'-1103"
  .Cells(33, 7) = "'-1104"
  .Cells(43, 7) = "'-1105"
  .Cells(53, 7) = "'-1106"
  .Cells(63, 7) = "'-1107"
  .Cells(73, 7) = "'-1108"
  .Cells(83, 7) = "'-1109" '2018.05.18追加
  
  .Cells(3, 39) = "'-1110"
  .Cells(13, 39) = "'-1111"
  .Cells(23, 39) = "'-1112"
  .Cells(33, 39) = "'-1113"
  .Cells(43, 39) = "'-1114"
  .Cells(53, 39) = "'-1115"
  .Cells(63, 39) = "'-1116"
  .Cells(73, 39) = "'-1117"
  .Cells(83, 39) = "'-1118" '2018.05.18追加
  
  .Cells(96, 7) = "'-1201" '副番
  .Cells(106, 7) = "'-1202"
  .Cells(116, 7) = "'-1203"
  .Cells(126, 7) = "'-1204"
  .Cells(136, 7) = "'-1205"
  .Cells(146, 7) = "'-1206"
  .Cells(156, 7) = "'-1207"
  .Cells(166, 7) = "'-1208"
  .Cells(176, 7) = "'-1209" '2018.05.18追加
  
  .Cells(96, 39) = "'-1210"
  .Cells(106, 39) = "'-1211"
  .Cells(116, 39) = "'-1212"
  .Cells(126, 39) = "'-1213"
  .Cells(136, 39) = "'-1214"
  .Cells(146, 39) = "'-1215"
  .Cells(156, 39) = "'-1216"
  .Cells(166, 39) = "'-1217"
  .Cells(176, 39) = "'-1218" '2018.05.18追加
  
  .Cells(92, 22) = テスラ全サイズ.TextBox1 '検査番号
  .Cells(92, 2) = ListName & "-1" 'サイズ1-1-?
  .Cells(92, 37) = テスラ全サイズ.TextBox2 '重量1-1-?
  .Cells(185, 22) = テスラ全サイズ.TextBox1 '検査番号
  .Cells(185, 2) = ListName & "-2" 'サイズ1-2-?
  .Cells(185, 37) = テスラ全サイズ.TextBox4 '重量1-2-?


テスラ全サイズ.Label16 = "「Coil1」" & Chr(10) & .Cells(92, 22) '検査番号
テスラ全サイズ.Label17 = .Cells(92, 37) & "kg"  '重量1-1-?
テスラ全サイズ.Label26 = .Cells(185, 37) & "kg"  '重量1-2-?
テスラ全サイズ.Label19 = "「Sheet1」" & Chr(10) & .Cells(92, 2) 'サイズ1-1-?
テスラ全サイズ.Label28 = .Cells(185, 2) 'サイズ1-2-?
End With

With ThisWorkbook.Sheets(ListName2)
  .Activate
  
  .Cells(3, 7) = "'-2101" '副番
  .Cells(13, 7) = "'-2102"
  .Cells(23, 7) = "'-2103"
  .Cells(33, 7) = "'-2104"
  .Cells(43, 7) = "'-2105"
  .Cells(53, 7) = "'-2106"
  .Cells(63, 7) = "'-2107"
  .Cells(73, 7) = "'-2108"
  .Cells(83, 7) = "'-2109" '2018.05.18追加
  
  .Cells(3, 39) = "'-2110"
  .Cells(13, 39) = "'-2111"
  .Cells(23, 39) = "'-2112"
  .Cells(33, 39) = "'-2113"
  .Cells(43, 39) = "'-2114"
  .Cells(53, 39) = "'-2115"
  .Cells(63, 39) = "'-2116"
  .Cells(73, 39) = "'-2117"
  .Cells(83, 39) = "'-2118" '2018.05.18追加
  
  .Cells(96, 7) = "'-2201" '副番
  .Cells(106, 7) = "'-2202"
  .Cells(116, 7) = "'-2203"
  .Cells(126, 7) = "'-2204"
  .Cells(136, 7) = "'-2205"
  .Cells(146, 7) = "'-2206"
  .Cells(156, 7) = "'-2207"
  .Cells(166, 7) = "'-2208"
  .Cells(176, 7) = "'-2209" '2018.05.18追加
  
  .Cells(96, 39) = "'-2210"
  .Cells(106, 39) = "'-2211"
  .Cells(116, 39) = "'-2212"
  .Cells(126, 39) = "'-2213"
  .Cells(136, 39) = "'-2214"
  .Cells(146, 39) = "'-2215"
  .Cells(156, 39) = "'-2216"
  .Cells(166, 39) = "'-2217"
  .Cells(176, 39) = "'-2218" '2018.05.18追加
  
  .Cells(92, 22) = テスラ全サイズ.TextBox1 '検査番号
  .Cells(92, 2) = ListName2 & "-1" 'サイズ2-1-?
  .Cells(92, 37) = テスラ全サイズ.TextBox3 '重量2-1-?
  .Cells(185, 22) = テスラ全サイズ.TextBox1 '検査番号
  .Cells(185, 2) = ListName2 & "-2" 'サイズ2-2-?
  .Cells(185, 37) = テスラ全サイズ.TextBox5 '重量2-2-?
  
テスラ全サイズ.Label18 = "「Coil2」" & Chr(10) & .Cells(92, 22) '検査番号
テスラ全サイズ.Label22 = .Cells(92, 37) & "kg"  '重量2-1-?
テスラ全サイズ.Label27 = .Cells(185, 37) & "kg"  '重量2-2-?
テスラ全サイズ.Label21 = "「Sheet2」" & Chr(10) & .Cells(92, 2) 'サイズ2-1-?
テスラ全サイズ.Label29 = .Cells(185, 2) 'サイズ2-2-?

End With

If a = False And c = False And b = False And D = False Then
    With ThisWorkbook.Sheets(ListName)
     .Cells(92, 22) = ""  '検査番号
     .Cells(92, 2) = ""   'サイズ1-1-?
     .Cells(92, 37) = "" '重量1-1-?
     
     .Cells(185, 22) = "" '検査番号
     .Cells(185, 2) = "" 'サイズ1-2-?
     .Cells(185, 37) = "" '重量1-2-?
    End With
    With ThisWorkbook.Sheets(ListName2)
      .Cells(92, 22) = "" '検査番号
      .Cells(92, 2) = "" 'サイズ2-1-?
      .Cells(92, 37) = "" '重量2-1-?
      
      .Cells(185, 22) = "" '検査番号
      .Cells(185, 2) = "" 'サイズ2-2-?
      .Cells(185, 37) = "" '重量2-2-?
   End With
End If

End If
'重量空白時はサイズ検番を消す
With ThisWorkbook.Sheets(ListName)
  If .Cells(92, 37) = "" Then
     .Cells(92, 22) = ""  '検査番号
     .Cells(92, 2) = ""   'サイズ1-1-?
  End If
End With
With ThisWorkbook.Sheets(ListName)
  If .Cells(185, 37) = "" Then
     .Cells(185, 22) = "" '検査番号
     .Cells(185, 2) = "" 'サイズ1-2-?
  End If
End With
  
With ThisWorkbook.Sheets(ListName2)
  If .Cells(92, 37) = "" Then
     .Cells(92, 22) = ""  '検査番号
     .Cells(92, 2) = ""   'サイズ1-1-?
  End If
End With
  
With ThisWorkbook.Sheets(ListName2)
  If .Cells(185, 37) = "" Then
     .Cells(185, 22) = "" '検査番号
     .Cells(185, 2) = "" 'サイズ1-2-?
  End If
End With

End Sub
Public Sub 型番入力()

'1.0×33×Coil a BJB7603900QR
'1.0×40×Coil b BJB7604300QR
'1.0×73×Coil c BJB7604000QR
'1.0×53.5×Coil d BJB7606500QR
'0.6×82.5×Coil e BJB7604200QR
'1.0×63.0×Coil f BJB7605900QR

Dim b As Integer
Dim i As Long
Dim a As Long
Dim D As Long 'リスト数

Dim ListName As String
Dim ListName2 As String

ListName = (テスラ全サイズ.ComboBox1.text & "1")
ListName2 = (テスラ全サイズ.ComboBox1.text & "2")

b = テスラ全サイズ.ComboBox1.ListIndex
D = テスラ全サイズ.ComboBox1.ListCount

テスラ全サイズ.Label23 = "登録数" & D & "件"

With ThisWorkbook.Sheets(ListName)
  .Activate
  .Cells(4, 4) = Worksheets("型番").Cells(b + 2, 3)
  .Cells(97, 4) = Worksheets("型番").Cells(b + 2, 3) '幅が広い場合不要
End With
With ThisWorkbook.Sheets(ListName2)
  .Activate
  .Cells(4, 4) = Worksheets("型番").Cells(b + 2, 3)
  .Cells(97, 4) = Worksheets("型番").Cells(b + 2, 3) '幅が広い場合不要
End With

End Sub

Sub 印刷範囲1() '1.0×63.0追加　2018.05.18
'2018.05.25　３ケタ時の範囲修正
  Dim ListName As String
  Dim ListName2 As String
  
    ListName = (テスラ全サイズ.ComboBox1.text & "1")
    ListName2 = (テスラ全サイズ.ComboBox1.text & "2")
    

If テスラ全サイズ.CheckBox2 = True Then
    With ThisWorkbook.Sheets(ListName)
        If .Cells(92, 37) <> "" And .Cells(185, 37) <> "" Then
        ActiveWindow.View = xlPageBreakPreview
        ActiveSheet.PageSetup.PrintArea = "$A$1:$BJ$186"
        ActiveWindow.View = xlNormalView
        ElseIf .Cells(185, 37) <> "" And .Cells(93, 37) = "" Then
        ActiveWindow.View = xlPageBreakPreview
        ActiveSheet.PageSetup.PrintArea = "$A$94:$BJ$186"
        ActiveWindow.View = xlNormalView
        ElseIf .Cells(185, 37) = "" And .Cells(93, 37) <> "" Then
        ActiveWindow.View = xlPageBreakPreview
        ActiveSheet.PageSetup.PrintArea = "$A$1:$BJ$93"
        ActiveWindow.View = xlNormalView
        Else
        End If
    End With
ElseIf テスラ全サイズ.ComboBox1.text = "1.0×63.0×Coil" Or テスラ全サイズ.ComboBox1.text = "1.0×73×Coil" Or テスラ全サイズ.ComboBox1.text = "0.6×82.5×Coil" Then '条が少ないもの
    ActiveSheet.PageSetup.PrintArea = "$A$1:$BJ$93"
        ActiveWindow.View = xlNormalView
Else
    ActiveWindow.View = xlPageBreakPreview
    ActiveSheet.PageSetup.PrintArea = "$A$1:$BJ$186"
        ActiveWindow.View = xlNormalView
End If


End Sub
Sub 印刷範囲2() '1.0×63.0追加　2018.05.18
'2018.05.25　３ケタ時の範囲修正
  Dim ListName As String
  Dim ListName2 As String
  
    ListName = (テスラ全サイズ.ComboBox1.text & "1")
    ListName2 = (テスラ全サイズ.ComboBox1.text & "2")
    

If テスラ全サイズ.CheckBox2 = True Then
    
    With ThisWorkbook.Sheets(ListName2)
        If .Cells(92, 37) <> "" And .Cells(185, 37) <> "" Then
        ActiveWindow.View = xlPageBreakPreview
        ActiveSheet.PageSetup.PrintArea = "$A$1:$BJ$186"
        ActiveWindow.View = xlNormalView
        ElseIf .Cells(185, 37) <> "" And .Cells(93, 37) = "" Then
        ActiveWindow.View = xlPageBreakPreview
        ActiveSheet.PageSetup.PrintArea = "$A$94:$BJ$186"
        ActiveWindow.View = xlNormalView
        ElseIf .Cells(185, 37) = "" And .Cells(93, 37) <> "" Then
        ActiveWindow.View = xlPageBreakPreview
        ActiveSheet.PageSetup.PrintArea = "$A$1:$BJ$93"
        ActiveWindow.View = xlNormalView
        Else
        End If
    End With
ElseIf テスラ全サイズ.ComboBox1.text = "1.0×63.0×Coil" Or テスラ全サイズ.ComboBox1.text = "1.0×73×Coil" Or テスラ全サイズ.ComboBox1.text = "0.6×82.5×Coil" Then '条が少ないもの
    ActiveSheet.PageSetup.PrintArea = "$A$1:$BJ$93"
        ActiveWindow.View = xlNormalView
Else
    ActiveWindow.View = xlPageBreakPreview
    ActiveSheet.PageSetup.PrintArea = "$A$1:$BJ$186"
        ActiveWindow.View = xlNormalView
End If


End Sub
Sub 型番表示()

  Dim b As Integer
  Dim i As Long
  Dim c As String
  Dim a As Long
  Dim j As Long '列
  Dim D As Long 'リスト数
  
  Dim ListName As String
  Dim ListName2 As String
  
    ListName = (テスラ全サイズ.ComboBox1.text & "1")
    ListName2 = (テスラ全サイズ.ComboBox1.text & "2")
  
    b = テスラ全サイズ.ComboBox1.ListIndex
    'c = テスラ全サイズ.ComboBox1.List(b, 0)
    D = テスラ全サイズ.ComboBox1.ListCount
    
'      UserForm3.Label1 = .Cells(b + 1, 2)
'      UserForm3.Label2 = .Cells(b + 1, 3)
'      UserForm3.Label3 = .Cells(b + 1, 4)Worksheets("型番").Cells(b + 2, 1)
'      UserForm3.Label4 = .Cells(b + 1, 5)
      
テスラ全サイズ.Label30 = "型番" & "　" & Worksheets("型番").Cells(b + 2, 3)

End Sub

Sub 印刷実行()

    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.text = ""
        .EvenPage.CenterHeader.text = ""
        .EvenPage.RightHeader.text = ""
        .EvenPage.LeftFooter.text = ""
        .EvenPage.CenterFooter.text = ""
        .EvenPage.RightFooter.text = ""
        .FirstPage.LeftHeader.text = ""
        .FirstPage.CenterHeader.text = ""
        .FirstPage.RightHeader.text = ""
        .FirstPage.LeftFooter.text = ""
        .FirstPage.CenterFooter.text = ""
        .FirstPage.RightFooter.text = ""
    End With
    Application.PrintCommunication = True
    
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False '印刷

End Sub

Sub 行列()

Dim myBaseRng As Range, myNewRng As Range
Dim myRow As Long, myClm As Long

    Dim ListName As String
    ListName = (テスラ全サイズ.ComboBox1.text & "1")

Set myBaseRng = ThisWorkbook.Sheets("原本").Range("A1:BJ189")
Set myNewRng = ThisWorkbook.Sheets(ListName).Range("A1")

  With myBaseRng
    
    For myClm = 1 To .Columns.count
    
        myNewRng.Columns(myClm).ColumnWidth = .Columns(myClm).ColumnWidth
    
    Next
    
    
    For myRow = 1 To .Rows.count
    
        myNewRng.Rows(myRow).RowHeight = .Rows(myRow).RowHeight
    
    Next
  
  End With

End Sub

Sub 行列2()

Dim myBaseRng As Range, myNewRng As Range
Dim myRow As Long, myClm As Long

    Dim ListName2 As String
    ListName2 = (テスラ全サイズ.ComboBox1.text & "2")

Set myBaseRng = ThisWorkbook.Sheets("原本").Range("A1:BJ189")
Set myNewRng = ThisWorkbook.Sheets(ListName2).Range("A1")

  With myBaseRng
    
    For myClm = 1 To .Columns.count
    
        myNewRng.Columns(myClm).ColumnWidth = .Columns(myClm).ColumnWidth
    
    Next
    
    
    For myRow = 1 To .Rows.count
    
        myNewRng.Rows(myRow).RowHeight = .Rows(myRow).RowHeight
    
    Next
  
  End With

End Sub

' フォームをリサイズ可能にするための設定
Public Sub FormSetting()
Dim result As Long
Dim hwnd As Long
Dim Wnd_STYLE As Long

hwnd = GetActiveWindow()
Wnd_STYLE = GetWindowLong(hwnd, GWL_STYLE)
Wnd_STYLE = Wnd_STYLE Or WS_THICKFRAME Or &H30000

result = SetWindowLong(hwnd, GWL_STYLE, Wnd_STYLE)
result = DrawMenuBar(hwnd)
End Sub
'ﾌｫﾙﾀﾞ・ﾌｧｲﾙの存在確認'※ﾈｯﾄﾜｰｸ上のﾌｧｲﾙ･ﾌｫﾙﾀﾞをﾁｪｯｸする際、Dir関数では存在しない場合ｴﾗｰが発生する。
Public Function PathCheck(path As String) As Boolean
'Dim fso As New FileSystemObject '(参照設定)Microsoft Scripting Runtime
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
Select Case GetFileSystemType(path) 'ﾁｪｯｸｶﾃｺﾞﾘ(1:ﾌｧｲﾙ, 2:ﾌｫﾙﾀﾞ)
  Case 1: If FileLen(path) >= 0 Then PathCheck = True 'ある(ﾌｧｲﾙｻｲｽﾞ追加)
  Case 2: PathCheck = True 'ある
  Case 3: PathCheck = False 'ない
End Select
Set fso = Nothing
End Function
'ﾌｫﾙﾀﾞ・ﾌｧｲﾙを見分ける
Public Function GetFileSystemType(path As String) As Long
'Dim fso As New FileSystemObject '(参照設定)Microsoft Scripting Runtime
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
Select Case True
  Case fso.fileExists(path):   GetFileSystemType = 1
  Case fso.FolderExists(path): GetFileSystemType = 2
  Case Else: GetFileSystemType = 3
End Select
Set fso = Nothing
End Function
'指定のﾌﾞｯｸが開かれているか確認する
Public Function IsBookOpen(filePath As String) As Boolean
On Error Resume Next
Open filePath For Append As #1
Close #1
If Err.Number > 0 Then
 IsBookOpen = True
Else
 IsBookOpen = False
End If
End Function
'任意のﾜｰｸｼｰﾄの存在ﾁｪｯｸ
Public Function WsCheck(ByRef wsName As String) As Boolean
Dim ws As Worksheet 'ﾜｰｸｼｰﾄ
WsCheck = False
For Each ws In ThisWorkbook.Worksheets
  If ws.name = wsName Then
    WsCheck = True
    Exit For
  End If
Next ws
Set ws = Nothing
End Function
'指定の名前のワークシートを追加
Public Sub AddWS(ByRef wsName As String)
Application.DisplayAlerts = False
Dim ws As Worksheet 'ﾜｰｸｼｰﾄ
For Each ws In ThisWorkbook.Worksheets
  If ws.name = wsName Then
    ThisWorkbook.Sheets(wsName).Visible = xlSheetVisible '_BeforeUpdate直後にUserForm_QueryCloseが実行されるとエラー_2021.10.26
    ThisWorkbook.Worksheets(wsName).Delete
    Exit For
  End If
Next ws
ThisWorkbook.Worksheets.Add(before:=Worksheets(1)).name = wsName 'before:=Worksheets(1)左端に追加　'Sheets(Sheets.Count)'右端に追加
Set ws = Nothing 'ｵﾌﾞｼﾞｪｸﾄの解放
Application.DisplayAlerts = True
End Sub
Public Function ShJuge(temp As String) As Worksheet 'なければ作り非表示
If WsCheck(temp) = False Then
 Call AddWS(temp)
End If
'ThisWorkbook.Sheets(Temp).Visible = False
Set ShJuge = ThisWorkbook.Sheets(temp)
End Function
'選択した範囲は空かどうか。データが無い場合は中断
Public Sub EmptyRangeJudge(Optional tgtRange As Range, Optional msg As String = "空のWorkSheet")
If Application.WorksheetFunction.CountA(tgtRange) = 0 Then
'  MsgBox msg & "を参照しています。中断します。", vbExclamation
'  End
End If
End Sub
'最終行の取得
Public Function UR(Optional SH As Worksheet, Optional n As Long = 0) As Long
If SH Is Nothing Then Set SH = ActiveSheet
Dim tgtRange As Range
Call EmptyRangeJudge(SH.UsedRange)
Dim lDummy As Long '仕様上のバグ対策(xlLastCellがﾘｾｯﾄされない)
Select Case n
  Case 0 '(含む：書式設定、非表示行)
  
    Application.EnableEvents = False 'ｲﾍﾞﾝﾄSelectChange抑制
    lDummy = SH.UsedRange.Row 'UsedRangeを無駄に使用して初期化
    SH.Unprotect
    UR = SH.Cells(1, 1).SpecialCells(xlLastCell).Row
    Application.EnableEvents = True
  Case Else '(含まない：書式設定、非表示行)
    UR = SH.Cells(Rows.count, n).End(xlUp).Row
    Set tgtRange = Range(SH.Cells(1, n), SH.Cells(UR, n))
    Call EmptyRangeJudge(tgtRange, "空の列")
End Select
End Function
'最終列の取得
Public Function UC(Optional SH As Worksheet, Optional n As Long = 0) As Long
If SH Is Nothing Then Set SH = ActiveSheet
Call EmptyRangeJudge(SH.UsedRange)
Dim lDummy As Long '仕様上のバグ対策(xlLastCellがﾘｾｯﾄされない)
Dim tgtRange As Range
Select Case n
  Case 0 '(含む：書式設定、非表示列)
    Application.EnableEvents = False 'ｲﾍﾞﾝﾄSelectChange抑制
    lDummy = SH.UsedRange.Column 'UsedRangeを無駄に使用して初期化
    UC = SH.Cells(1, 1).SpecialCells(xlLastCell).Column
    Application.EnableEvents = True
  Case Else '(含まない：書式設定、非表示列)
    UC = SH.Cells(n, Columns.count).End(xlToLeft).Column
    Set tgtRange = Range(SH.Cells(1, 1), SH.Cells(n, UC))
    Call EmptyRangeJudge(tgtRange, "空の行")
End Select
End Function
'ﾌｧｲﾙ配列変換（1次元配列）
Public Function GetOneDimData(fileName As String) As Variant
Dim FreeNum As Long: FreeNum = FreeFile
Dim buf() As Byte 'Binary読込用全ﾃﾞｰﾀ
If Not PathCheck(fileName) Then Exit Function 'ﾌｧｲﾙｻｲｽﾞ確認
Open fileName For Binary As #FreeNum 'ﾊﾞｲﾅﾘ形式でﾌｧｲﾙを一括読み込み
ReDim buf(LOF(1))
Get #FreeNum, , buf
Close #FreeNum
'Unicodeに変換し改行ｺｰﾄﾞ毎に区切って配列化（ﾃﾞｰﾀを1行毎に区切る）'ﾃﾞｰﾀを配列で返す。
GetOneDimData = Split(StrConv(buf(), vbUnicode), vbCrLf, , vbBinaryCompare)
End Function
'ファイルを開く
Public Function xlsFileOpenWb(filePath As String, Optional mode As Boolean = True) As Object '"ReadOnly"
Dim fileName As String: fileName = GetFileName(filePath)
If IsBookOpen(fileName) Then Exit Function
workbooks.Open filePath, ReadOnly:=mode
Set xlsFileOpenWb = ActiveWorkbook
End Function
'ファイルを閉じる
Public Sub xlsFileClose(tgtWB As Workbook, Optional saveflag As Boolean = False) '"ReadOnly"
If Not IsBookOpen(tgtWB.name) Then Exit Sub
Application.DisplayAlerts = False
If saveflag Then tgtWB.Save
tgtWB.Close
Application.DisplayAlerts = True
End Sub
'ﾌﾙﾊﾟｽからﾌｧｲﾙ名を抜き出す
Public Function GetFileName(PathName As String) As String
Dim temp As Variant: temp = Split(PathName, "\")
GetFileName = temp(UBound(temp))
End Function
'指定WSデータの2次元配列化
Public Function shDataGet(shName As String, Optional k As Long = 0) As Variant
Dim SH As Worksheet: Set SH = Worksheets(shName)
shDataGet = Range(SH.Cells(1, 1), SH.Cells(UR(SH), UC(SH, k)))
End Function
'指定WSデータの2次元配列化
Public Function UseWs(SH As Worksheet, Optional k As Long = 0) As Variant
UseWs = Range(SH.Cells(1, 1), SH.Cells(UR(SH), UC(SH, k)))
End Function
Public Sub CheckBookOpened()
Dim filePath
Dim ret As Boolean
'filepath = OfficeList.Terget & OfficeList.Pass
ret = IsBookOpened(filePath) ' 開いているかチェック
On Error Resume Next
  If (ret = True) Then
    Debug.Print "開いている"
    Exit Sub
  Else
    Debug.Print "開いていない"
  End If
End Sub
' 配列の次元数を求める
' 引数1 : vlArray / 対象の変数' 戻り値 : [Integer] / -1:配列ではない, 0:初期化されていないか要素無し,1 以上: 配列の次元数
Function Get_ArrayDim(ByVal vlArray As Variant) As Long
Dim i As Long
Dim varTmp As Variant
i = 0
On Error GoTo ErrLine
If Not IsArray(vlArray) Then GoTo ErrLine
  Do
    i = i + 1
    varTmp = UBound(vlArray, i)
    If varTmp = -1 Then GoTo ErrLine
  Loop
ErrLine:
If Err.Number <> 0 Then Err.Clear
Get_ArrayDim = i - 1
End Function
'二次元配列に並べ替え
Public Sub merge_sort2(ByRef arr As Variant, ByVal col As Long) 'うまく書かれたクイックソートの1.5倍程度処理時間がかかる、メモリ使用量はクイックソートの3倍程度
  Dim irekae As Variant
  Dim indexer As Variant
  Dim tmp1() As Variant
  Dim tmp2() As Variant
  Dim i As Long
  ReDim irekae(LBound(arr, 1) To UBound(arr, 1))
  ReDim indexer(LBound(arr, 1) To UBound(arr, 1))
  ReDim tmp1(LBound(arr, 1) To UBound(arr, 1))
  ReDim tmp2(LBound(arr, 1) To UBound(arr, 1))
  For i = LBound(arr, 1) To UBound(arr, 1) Step 2
      If i + 1 > UBound(arr, 1) Then
          irekae(i) = arr(i, col)
          indexer(i) = i
          Exit For
      End If
      If arr(i + 1, col) < arr(i, col) Then
          irekae(i) = arr(i + 1, col)
          irekae(i + 1) = arr(i, col)
          indexer(i) = i + 1
          indexer(i + 1) = i
      Else
          irekae(i) = arr(i, col)
          irekae(i + 1) = arr(i + 1, col)
          indexer(i) = i
          indexer(i + 1) = i + 1
      End If
  Next
  Dim ST1 As Long
  Dim en1 As Long
  Dim st2 As Long
  Dim en2 As Long
  Dim n As Long
  i = 1
  Do While i * 2 <= UBound(arr, 1)
      i = i * 2
      n = 0
      Do While en2 + i - 1 < UBound(arr, 1)
          n = n + 1
          ST1 = i * 2 * (n - 1) + LBound(arr, 1)
          en1 = i * 2 * (n - 1) + i - 1 + LBound(arr, 1)
          st2 = en1 + 1
          en2 = IIf(st2 + i - 1 >= UBound(arr, 1), UBound(arr, 1), st2 + i - 1)
          Call merge2(irekae, indexer, tmp1, tmp2, ST1, en1, st2, en2)
      Loop
      en2 = 0
  Loop
  Dim ret As Variant
  ReDim ret(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
  For i = LBound(arr, 1) To UBound(arr, 1)
      For n = LBound(arr, 2) To UBound(arr, 2)
          If IsObject(arr(indexer(i), n)) Then
              Set ret(i, n) = arr(indexer(i), n)
          Else
              ret(i, n) = arr(indexer(i), n)
          End If
      Next
  Next
  arr = ret
End Sub

Public Sub merge2(ByRef irekae As Variant, _
ByRef indexer As Variant, _
ByRef tmpArr() As Variant, _
ByRef tmpIndexer() As Variant, _
ByVal ST1 As Long, _
ByVal en1 As Long, _
ByVal st2 As Long, _
ByVal en2 As Long)
    Dim j As Long
    Dim n As Long
    Dim i As Long
    For i = ST1 To en2
        tmpArr(i) = irekae(i)
        tmpIndexer(i) = indexer(i)
    Next
    j = ST1
    n = st2
    Do While (j < en1 + 1 Or n < en2 + 1)
        If n >= en2 + 1 Then
            irekae(j + n - st2) = tmpArr(j)
            indexer(j + n - st2) = tmpIndexer(j)
            j = j + 1
        ElseIf j < en1 + 1 And tmpArr(j) <= tmpArr(n) Then
            irekae(j + n - st2) = tmpArr(j)
            indexer(j + n - st2) = tmpIndexer(j)
            j = j + 1
        Else
            irekae(j + n - st2) = tmpArr(n)
            indexer(j + n - st2) = tmpIndexer(n)
            n = n + 1
        End If
    Loop
End Sub
Public Sub 半角へ変換() '数値のみを半角へ変換
Dim i As Long, n As Long, LR As Long, LC As Long
Dim SaveSh As Worksheet: Set SaveSh = ThisWorkbook.Sheets("Sheet1")
LR = UR(SaveSh, 1)
LC = UC(SaveSh, 1)
Dim SaveR As Range:      Set SaveR = Range(SaveSh.Cells(2, 1), SaveSh.Cells(LR, 2))
Dim SaveTemp():          SaveTemp = SaveR 'シートを配列へ

For n = LBound(SaveTemp, 1) To UBound(SaveTemp, 1)
  For i = 1 To Len(SaveTemp(n, 1)) '数値のみを半角へ変換
    If Mid(SaveTemp(n, 1), i, 1) Like "[０-９]" Then
      Mid(SaveTemp(n, 1), i, 1) = StrConv(Mid(SaveTemp(n, 1), i, 1), vbNarrow)
    End If
  Next i
  SaveTemp(n, 1) = numExtract(SaveTemp(n, 1)) '数字だけ抽出
  If SaveTemp(n, 1) <> "" Then SaveSh.Cells(n + 1, 3) = val(SaveTemp(n, 1))
  SaveSh.Cells(n + 1, 3).NumberFormatLocal = "0.0_ "
  SaveTemp(n, 2) = numExtract(SaveTemp(n, 2))
  If SaveTemp(n, 2) <> "" Then SaveSh.Cells(n + 1, 4) = val(SaveTemp(n, 2))
  SaveSh.Cells(n + 1, 4).NumberFormatLocal = "0.0_ "
Next n
'Stop
End Sub
Public Function ConvertToHalf(temp As Variant) As String '数値のみを半角へ変換
Dim i As Long
For i = 1 To Len(temp)
  If Mid(temp, i, 1) Like "[０-９]" Then
    Mid(temp, i, 1) = StrConv(Mid(temp, i, 1), vbNarrow)
  End If
Next i
ConvertToHalf = numExtract(temp) '数字だけ抽出
'Stop
End Function

'文字列から数字だけ抽出
Public Function numExtract(StringValue As Variant) As String
 Dim i As Integer
 Dim numText As String
 Dim temp
 If InStr(StringValue, "/") > 0 Then '1.0/2M スラッシュ前の数字を反映
  temp = Split(StringValue, "/")
  StringValue = temp(LBound(temp))
 End If
' Stop
 For i = 1 To Len(StringValue) '1 to 文字列（最終文字）まで For Nextで影響させる
    numText = Mid(StringValue, i, 1) 'Midで文字列を左から順にnumTextに格納
    If numText Like "[0-9.]" Then: numExtract = numExtract & numText 'numText[0-9]に該当する場合 変数numExtractに格納していく**小数点含む
 Next i
End Function

