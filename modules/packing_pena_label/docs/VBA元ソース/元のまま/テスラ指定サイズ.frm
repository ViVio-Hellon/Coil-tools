VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} テスラ指定サイズ 
   Caption         =   "2026.03.26"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17730
   OleObjectBlob   =   "テスラ指定サイズ.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "テスラ指定サイズ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========== プログレス表示 ==========
Private progressForm As Object
Private mSuppressEvent As Boolean

'==============================================================================
' テスラ指定サイズ UserForm - リファクタリング版
' 共通処理: OptionButton選択時
'--------------------------------------------------------------
Private Sub SelectSize(ByVal shtName As String, _
                       ByVal dataRow As Long, _
                       ByVal takeType As Long)
    If mSuppressEvent Then Exit Sub   ' ★ 追加
    mSuppressEvent = True
    Dim i As Long

    '--- CheckBox 全解除 ---
    For i = 1 To 6
        Me.Controls("CheckBox" & i) = False
    Next i
    Me.Controls("ck_08_53") = False

    '--- Take / TextBox 表示切替 ---
    If takeType = 1 Then
        Me.Take1.Visible = True
        Me.Take2.Visible = False
        Me.txtWeight1.Visible = True
        Me.txtWeight2.Visible = False
        Me.txtWeight2 = ""
    Else
        Me.Take1.Visible = False
        Me.Take2.Visible = True
        Me.txtWeight2.Visible = True
        Me.txtWeight1.Visible = False
        Me.txtWeight1 = ""
    End If

    '--- ラベル初期化 ---
    Me.lblKensaNo = ""
    Me.lblSize1 = ""
    Me.lblWeight1 = ""
    Me.lblSize2 = ""
    Me.lblWeight2 = ""

    '--- シートからデータ取得 ---
    With ThisWorkbook.Sheets(shtName)
        .Activate
        Me.lblKensaNo = .Cells(dataRow, 22)            ' 検査番号

        If takeType = 1 Then
            Me.lblWeight1 = .Cells(dataRow, 37) & "kg" ' 重量(丈1)
            Me.lblSize1 = .Cells(dataRow, 2)            ' サイズ(丈1)
        Else
            Me.lblWeight2 = .Cells(dataRow, 37) & "kg" ' 重量(丈2)
            Me.lblSize2 = .Cells(dataRow, 2)            ' サイズ(丈2)
        End If
    End With

    ThisWorkbook.Sheets("起動").Activate
     mSuppressEvent = False
End Sub

'--------------------------------------------------------------
' 特殊クリア: OB5/OB6 用 CoilH, NW, GW, TA リセット
'--------------------------------------------------------------
Private Sub ClearCoilInfo(ByVal s1 As String, ByVal s2 As String)
    Me.Controls("CoilH" & s1) = ""
    Me.Controls("CoilH" & s2) = ""
    Me.Controls("NW" & s1).Caption = ""
    Me.Controls("NW" & s2).Caption = ""
    Me.Controls("GW" & s1).Caption = ""
    Me.Controls("GW" & s2).Caption = ""
    Me.Controls("TA" & s1).Caption = ""
    Me.Controls("TA" & s2).Caption = ""
End Sub

'==============================================================================
' 各 OptionButton_Click （1行呼び出し）
'==============================================================================
Private Sub lblSpec_08_53_1_Click()
  SelectSize "0.8mm×53.5mm 　丈1", 3, 1
  ClearCoilInfo "1", "2"
End Sub

Private Sub lblSpec_08_53_2_Click()
  SelectSize "0.8mm×53.5mm 　丈2", 3, 2
  ClearCoilInfo "3", "4"
End Sub

Private Sub OptionButton1_Click()
    SelectSize "1.0mm×73mm 　丈1", 3, 1
End Sub

Private Sub OptionButton2_Click()
    SelectSize "1.0mm×73mm 　丈2", 3, 2
End Sub

Private Sub OptionButton3_Click()
    SelectSize "0.6mm×82.5mm 　丈1", 3, 1
End Sub

Private Sub OptionButton4_Click()
    SelectSize "0.6mm×82.5mm 　丈2", 3, 2
End Sub

Private Sub OptionButton5_Click()
    SelectSize "1.0mm×53.5mm 　丈1", 3, 1
    ClearCoilInfo "1", "2"
End Sub

Private Sub OptionButton6_Click()
    SelectSize "1.0mm×53.5mm 　丈2", 3, 2
    ClearCoilInfo "3", "4"
End Sub

Private Sub OptionButton7_Click()
    SelectSize "1.0mm×40.0mm 　丈1", 3, 1
End Sub

Private Sub OptionButton8_Click()
    SelectSize "1.0mm×40.0mm 　丈2", 3, 2
End Sub

Private Sub OptionButton9_Click()
    SelectSize "1.0mm×33.0mm 　丈1", 3, 1
End Sub

Private Sub OptionButton10_Click()
    SelectSize "1.0mm×33.0mm 　丈2", 3, 2
End Sub

Private Sub OptionButton11_Click()
    SelectSize "1.0mm×63.0mm 　丈1", 92, 1
End Sub

Private Sub OptionButton12_Click()
    SelectSize "1.0mm×63.0mm 　丈2", 92, 2
End Sub

'==============================================================================
' テスラ指定サイズ UserForm  ―  CheckBox_Click 共通処理 + 各イベント
'==============================================================================

'--------------------------------------------------------------
' 共通処理: CheckBox 選択時（丈1,2 同時モード）
'==============================================================================
' SelectCheckBox  ― pair を Long 2変数に変更
'==============================================================================
Private Sub SelectCheckBox(ByVal cbIdx As Long, _
                           Optional ByVal isNamed As Boolean = False)
    If mSuppressEvent Then Exit Sub
    mSuppressEvent = True

    Dim i As Long

    '--- 他の CheckBox を全解除し、自分だけ ON ---
    For i = 1 To 6
        Me.Controls("CheckBox" & i) = False
    Next i
    Me.ck_08_53.Value = False

    If isNamed Then
        Me.ck_08_53.Value = True
    Else
        Me.Controls("CheckBox" & cbIdx) = True
    End If

    Me.Take1.Visible = True
    Me.Take2.Visible = True
    Me.txtWeight1.Visible = True
    Me.txtWeight2.Visible = True

    For i = 1 To 12
        Me.Controls("OptionButton" & i) = False
    Next i
    Me.lblSpec_08_53_1.Value = False
    Me.lblSpec_08_53_2.Value = False

    Me.lblKensaNo = ""
    Me.lblSize1 = ""
    Me.lblWeight1 = ""
    Me.lblSize2 = ""
    Me.lblWeight2 = ""

    '--- OBペア取得 (Long 変数で受ける) ---
    Dim ob1 As Long, ob2 As Long
    If isNamed Then
        ob1 = 13: ob2 = 14
    Else
        GetCBtoPair cbIdx, ob1, ob2
    End If

    DebugLog "[SelectCheckBox] ob1=" & ob1 & " ob2=" & ob2

    Dim cfg1 As Variant: cfg1 = GetSizeConfig(ob1)
    Dim cfg2 As Variant: cfg2 = GetSizeConfig(ob2)
    Dim row1 As Long:    row1 = CLng(cfg1(3))
    Dim row2 As Long:    row2 = CLng(cfg2(3))

    With ThisWorkbook.Sheets(BuildSheetName(ob1))
        Me.lblKensaNo = .Cells(row1, HDR_KEN_COL)
        Me.lblWeight1 = .Cells(row1, HDR_WT1_COL) & "kg"
        Me.lblSize1 = .Cells(row1, 2)
    End With

    With ThisWorkbook.Sheets(BuildSheetName(ob2))
        Me.lblWeight2 = .Cells(row2, HDR_WT1_COL) & "kg"
        Me.lblSize2 = .Cells(row2, 2)
    End With

    mSuppressEvent = False
End Sub
'==============================================================================
' 各 CheckBox_Click （1行呼び出し）
'==============================================================================

Private Sub CheckBox1_Click()
    If Me.CheckBox1.Value Then SelectCheckBox 1
End Sub

Private Sub CheckBox2_Click()
    If Me.CheckBox2.Value Then SelectCheckBox 2
End Sub

Private Sub CheckBox3_Click()
    If Me.CheckBox3.Value Then SelectCheckBox 3
End Sub

Private Sub CheckBox4_Click()
    If Me.CheckBox4.Value Then SelectCheckBox 4
End Sub

Private Sub CheckBox5_Click()
    If Me.CheckBox5.Value Then SelectCheckBox 5
End Sub

Private Sub CheckBox6_Click()
    If Me.CheckBox6.Value Then SelectCheckBox 6
End Sub

Private Sub ck_08_53_Click()
    If Me.ck_08_53.Value Then SelectCheckBox 0, isNamed:=True
End Sub

Private Sub CoilH1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select
End Sub

Private Sub CoilH1_Change()
 With テスラ指定サイズ
  If Not IsNumeric(.CoilH1.text) Then .CoilH1.text = ""
 End With
End Sub
Private Sub CoilH2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select
End Sub

Private Sub CoilH2_Change()
 With テスラ指定サイズ
  If Not IsNumeric(.CoilH2.text) Then .CoilH2.text = ""
 End With
End Sub
Private Sub CoilH3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select
End Sub

Private Sub CoilH3_Change()
 With テスラ指定サイズ
  If Not IsNumeric(.CoilH3.text) Then .CoilH3.text = ""
 End With
End Sub
Private Sub CoilH4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select
End Sub

Private Sub CoilH4_Change()
 With テスラ指定サイズ
  If Not IsNumeric(.CoilH4.text) Then .CoilH4.text = ""
 End With
End Sub
'Private Sub CheckBox7_AfterUpdate()
'With テスラ指定サイズ
' If .TIP1000 = True Then
'    If .Take1.Visible = False And .Take2.Visible = False Then
'    Exit Sub
'    End If
'    If .CheckBox1 = False And .CheckBox2 = False And .CheckBox3 = False And .CheckBox4 = _
'    False And .CheckBox5 = False And .CheckBox6 = False And .OptionButton1 = False And .OptionButton2 = _
'    False And .OptionButton3 = False And .OptionButton4 = False And .OptionButton5 = False And .OptionButton6 = _
'    False And .OptionButton7 = False And .OptionButton8 = False And .OptionButton9 = False And .OptionButton10 = _
'    False And .OptionButton11 = False And .OptionButton12 = False Then
'     MsgBox "コイルサイズが選択されていません"
'     Exit Sub
'    Else
'     Call 計算1
'     Call 計算2
'
'     Call ラベル転記CK
'     Call ラベル転記OP
'    End If
' Else
'  MsgBox "チップボールチェックがありません"
' End If
'End With
'End Sub

'==============================================================================
' CommandButton1_Click リファクタリング版
'==============================================================================
' ■ コントロール名変更
'   TextBox1 → txtKensaNo  (検査番号入力)
'   TextBox2 → txtWeight1  (重量入力・丈1用)
'   TextBox3 → txtWeight2  (重量入力・丈2用)
'   CommandButton1 → そのまま or btnSave 等、任意で変更可
'==============================================================================
Private Sub CommandButton1_Click()

    Dim ken As String
    ken = StrConv(UCase(Me.txtKensaNo), vbNarrow)

    '--- 検査番号バリデーション ---
    If Not ValidateKensaNo(ken) Then Exit Sub

    '--- 重量入力チェック ---
    If InpCheck(1) Then MsgBox "重量入力がありません　丈1": Exit Sub
    If InpCheck(2) Then MsgBox "重量入力がありません　丈2": Exit Sub
    If InpCheck(3) Then MsgBox "重量入力がありません　丈1　or　丈2": Exit Sub

    Call チェック入力

    '--- 選択中 OB を検出して転記 ---
    Dim obIdx As Long
    obIdx = DetectSelectedOB(Me)
    If obIdx > 0 Then
        ExecuteWrite ken, obIdx
    End If

End Sub

'--------------------------------------------------------------
' 転記実行
' ExecuteWrite  ― 引数変更 (coilNo → takeType)
'==============================================================================
Private Sub ExecuteWrite(ByVal ken As String, ByVal obIdx As Long)

    Dim cfg      As Variant:  cfg = GetSizeConfig(obIdx)
    Dim shtName  As String:   shtName = BuildSheetName(obIdx)
    Dim takeType As Long:     takeType = GetTakeType(obIdx)
    Dim kataban  As String:   kataban = CStr(cfg(1))

    Dim wt As String
    If takeType = 1 Then
        wt = Me.txtWeight1
    Else
        wt = Me.txtWeight2
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(shtName)
    ws.Activate

    Call WriteAllToSheet(ws, ken, wt, kataban, takeType)  ' ★ takeType を渡す

    Dim dataRow As Long: dataRow = CLng(cfg(3))
    Me.lblKensaNo = ws.Cells(dataRow, HDR_KEN_COL)

    If takeType = 1 Then
        Me.lblWeight1 = ws.Cells(dataRow, HDR_WT1_COL) & "kg"
        Me.lblSize1 = ws.Cells(dataRow, 2)
    Else
        Me.lblWeight2 = ws.Cells(dataRow, HDR_WT1_COL) & "kg"
        Me.lblSize2 = ws.Cells(dataRow, 2)
    End If

    ThisWorkbook.Sheets("起動").Activate
End Sub

'--------------------------------------------------------------
' 検査番号バリデーション
'--------------------------------------------------------------
Private Function ValidateKensaNo(ByVal ken As String) As Boolean
    ValidateKensaNo = False
    Dim msg As String
    msg = "検査NOが7桁入力されていない" & Chr(10) & "もしくは"

    If Not IsNumeric(Mid(ken, 2, 1)) Then
        MsgBox msg & "6桁目にアルファベットが入力されています"
    ElseIf Not IsNumeric(Mid(ken, 7, 1)) Then
        MsgBox msg & "1桁目にアルファベットが入力されています"
    ElseIf Not IsNumeric(Mid(ken, 4, 4)) Then
        MsgBox msg & "1～4桁目にアルファベットが入力されています"
    ElseIf IsNumeric(Mid(ken, 1, 1)) Then
        MsgBox msg & "7桁目に数字が入力されています"
    Else
        ValidateKensaNo = True
    End If
End Function

Private Sub CommandButton10_Click()
If MsgBox(sheetName & "  " & "印刷する？", vbYesNo) = vbYes Then
  If MsgBox(sheetName & "  " & "ﾗﾍﾞﾙ用紙から普通紙に変えてください", vbYesNo) = vbYes Then
'    Call 羅列計算
'    ThisWorkbook.Sheets("50まで").Activate
'    ThisWorkbook.Sheets("50まで").Visible = True 'シート表示
'    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
'    IgnorePrintAreas:=False
'    ThisWorkbook.Sheets("50まで").Visible = False 'シート非表示
    
    ThisWorkbook.Sheets("風袋計算").Activate
    Call 風袋罫線設定
    ThisWorkbook.Sheets("風袋計算").Visible = True 'シート表示
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False
    ThisWorkbook.Sheets("風袋計算").Visible = False 'シート非表示
  Else
  End If
Else
End If
End Sub

Private Sub CommandButton11_Click()
    ListName = "50まで"
    SheetsName = ListName
    
    For Each ws In Worksheets
        If ws.name = SheetsName Then Flag = True
    Next ws
 Application.ScreenUpdating = False
    If Flag = True Then
    ThisWorkbook.Sheets("50まで").Activate
    ThisWorkbook.Sheets("50まで").Visible = True 'シート表示
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False
    End If
 Application.ScreenUpdating = True
End Sub

Private Sub CommandButton12_Click()
    ListName = "風袋計算"
    SheetsName = ListName
 Application.ScreenUpdating = False
    For Each ws In Worksheets
        If ws.name = SheetsName Then Flag = True
    Next ws
    If Flag = True Then
    ThisWorkbook.Sheets("風袋計算").Activate
    ThisWorkbook.Sheets("風袋計算").Visible = True 'シート表示
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False
    End If
 Application.ScreenUpdating = True
End Sub

Private Sub CommandButton13_Click()
Dim col As Double 'コイル巾
    With テスラ指定サイズ
       If .CheckBox1 = True Then
         col = 40
       ElseIf .CheckBox2 = True Then
         col = 33
       ElseIf .CheckBox3 = True Then
         col = 73
       ElseIf .CheckBox4 = True Then
         col = 82.5
       ElseIf .CheckBox5 = True Then
         col = 43.5
       ElseIf .CheckBox6 = True Then
         col = 63
       End If
       
       If .OptionButton7 = True Or .OptionButton8 = True Then
         col = 40
       ElseIf .OptionButton9 = True Or .OptionButton10 = True Then
         col = 33
       ElseIf .OptionButton1 = True Or .OptionButton2 = True Then
         col = 73
       ElseIf .OptionButton3 = True Or .OptionButton4 = True Then
         col = 82.5
       ElseIf .OptionButton5 = True Or .OptionButton6 = True Then
         col = 43.5
       ElseIf .OptionButton11 = True Or .OptionButton12 = True Then
         col = 63
       End If
    End With
Dim x As Long
Dim TA As Double
Dim TIPT As Double 'ﾁｯﾌﾟﾎﾞｰﾙの高さ　1枚0.7mm

         x = Application.InputBox("コイル積み数", Type:=2)
        TIPT = val((x + 1) * 0.7 / 10 ^ 3)
         TA = val(1 * col / 10 ^ 3) + TIPT 'ﾃｷｽﾄﾎﾞｯｸｽの数字×ﾁｪｯｸの入っているｺｲﾙ 'M単位
         KTA = val((TA * 1000) + 150 + 150)  '樹脂パレット高さ+150*2
     MsgBox "ﾁｯﾌﾟ:" & TIPT * 1000 & Chr(10) & "1条分:" & TA * 1000 & Chr(10) & "ﾊﾟﾚｯﾄ+分(繰上値):" & Application.WorksheetFunction.RoundUp((KTA), -1) & " : " & "そのまま" & KTA
         
End Sub

Private Sub CommandButton14_Click()
Call ListOen
End Sub
'
'Private Sub CommandButton15_Click()
'If Me.CoilH1 = "" And Me.CoilH2 = "" And Me.CoilH3 = "" And Me.CoilH4 = "" Then Exit Sub
'
'Call クリア '入力内容クリア
'
' '*************************風袋シートの作成
'Dim SheetsName As String: SheetsName = "風袋計算"
'Dim ws As Worksheet, Flag As Boolean
'For Each ws In ThisWorkbook.Worksheets
'  If ws.name = SheetsName Then Flag = True
'Next ws
'If Flag Then     '同じ名前のシートあれば消す
'  Application.DisplayAlerts = False
'  ThisWorkbook.Sheets(SheetsName).Delete
'  Application.DisplayAlerts = True
'End If
'ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(Sheets.count) '新たにシートを追加する
'ActiveSheet.name = SheetsName
'ThisWorkbook.Sheets(SheetsName).Visible = False
'
'With テスラ指定サイズ
'  If .TIP1000 = True Or .TIP950 = True Or .TIP820 = True Then
'     If .Take1.Visible = False And .Take2.Visible = False Then Exit Sub
'
'      If .CheckBox1 = False And .CheckBox2 = False And .CheckBox3 = False And .CheckBox4 = _
'      False And .CheckBox5 = False And .CheckBox6 = False And .OptionButton1 = False And .OptionButton2 = _
'      False And .OptionButton3 = False And .OptionButton4 = False And .OptionButton5 = False And .OptionButton6 = _
'      False And .OptionButton7 = False And .OptionButton8 = False And .OptionButton9 = False And .OptionButton10 = _
'      False And .OptionButton11 = False And .OptionButton12 = False Then
'      MsgBox "コイルサイズが選択されていません"
'        .Kei = False '初期化
'        Exit Sub
'      Else
'        UFProgress.Show vbModeless
'        Call ProgressBar_count(0, 455)
'        If Me.CoilH4 <> "" Then Call All計算(Me.CoilH4, 4) '4から実行する(ﾌｫｰﾑ表記は丈1_No1)
'        If Me.CoilH3 <> "" Then Call All計算(Me.CoilH3, 3)
'        If Me.CoilH2 <> "" Then Call All計算(Me.CoilH2, 2)
'        If Me.CoilH1 <> "" Then Call All計算(Me.CoilH1, 1)
'
'        Call ProgressBar_count(455, 888)
'
'        Call ラベル転記CK
'        Call ラベル転記OP
'
'        Call ProgressBar_count(888, 1000)
'        Unload UFProgress
'     End If
'  Else
'    MsgBox "チップボールチェックがありません"
'    Exit Sub
'  End If
'End With
'ThisWorkbook.Activate
'ThisWorkbook.Sheets("起動").Select
'End Sub

Private Sub CommandButton16_Click()
If IsAllCoilEmpty() Then Exit Sub

Call クリア
Application.ScreenUpdating = False

'====== 風袋シートの作成 ======
RecreateSheet "風袋計算"

With テスラ指定サイズ
  '====== バリデーション ======
  If Not IsValidSettings() Then
    Application.ScreenUpdating = True
    Exit Sub
  End If
  
  '====== プログレスバー初期化 ======
  ShowProgressForm
  
  '====== DB_Materials実行 ======
  Dim coilCount As Long: coilCount = CountFilledCoils()
  Dim currentIndex As Long: currentIndex = 0
  
  If .CoilH4 <> "" Then
    UpdateProgress currentIndex, coilCount, "梱包4を処理中..."
    Call DB_Materials(.CoilH4, 4)
    currentIndex = currentIndex + 1
  End If
  
  If .CoilH3 <> "" Then
    UpdateProgress currentIndex, coilCount, "梱包3を処理中..."
    Call DB_Materials(.CoilH3, 3)
    currentIndex = currentIndex + 1
  End If
  
  If .CoilH2 <> "" Then
    UpdateProgress currentIndex, coilCount, "梱包2を処理中..."
    Call DB_Materials(.CoilH2, 2)
    currentIndex = currentIndex + 1
  End If
  
  If .CoilH1 <> "" Then
    UpdateProgress currentIndex, coilCount, "梱包1を処理中..."
    Call DB_Materials(.CoilH1, 1)
    currentIndex = currentIndex + 1
  End If
  
  UpdateProgress coilCount, coilCount, "ラベル転記中..."
  Call ラベル転記CK
  Call ラベル転記OP
  
  '====== プログレスバー終了 ======
  HideProgressForm
End With

ThisWorkbook.Activate
ThisWorkbook.Sheets("起動").Select
Application.ScreenUpdating = True
End Sub

'========== バリデーション関数 ==========
Private Function IsAllCoilEmpty() As Boolean
  With テスラ指定サイズ
    IsAllCoilEmpty = (.CoilH1 = "" And .CoilH2 = "" And .CoilH3 = "" And .CoilH4 = "")
  End With
End Function

Private Function IsValidSettings() As Boolean
  With テスラ指定サイズ
    'チップボール選択確認
    If Not (.TIP1000 Or .TIP950 Or .TIP820) Then
      MsgBox "チップボールチェックがありません"
      .Kei = False
      IsValidSettings = False
      Exit Function
    End If
    
    '丈フレーム表示確認
    If Not (.Take1.Visible Or .Take2.Visible) Then
      IsValidSettings = False
      Exit Function
    End If
    
    'コイルサイズ選択確認
    If Not IsAnySizeSelected() Then
      MsgBox "コイルサイズが選択されていません"
      .Kei = False
      IsValidSettings = False
      Exit Function
    End If
    
    IsValidSettings = True
  End With
End Function

Private Function IsAnySizeSelected() As Boolean
  With テスラ指定サイズ
    IsAnySizeSelected = (.CheckBox1 Or .CheckBox2 Or .CheckBox3 Or .CheckBox4 Or .CheckBox5 Or .CheckBox6 Or _
                         .ck_08_53 Or _
                         .OptionButton1 Or .OptionButton2 Or .OptionButton3 Or .OptionButton4 Or .OptionButton5 Or _
                         .OptionButton6 Or .OptionButton7 Or .OptionButton8 Or .OptionButton9 Or .OptionButton10 Or _
                         .OptionButton11 Or .OptionButton12 Or _
                         .lblSpec_08_53_1 Or .lblSpec_08_53_2)
  End With
End Function

Private Function CountFilledCoils() As Long
  Dim count As Long: count = 0
  With テスラ指定サイズ
    If .CoilH1 <> "" Then count = count + 1
    If .CoilH2 <> "" Then count = count + 1
    If .CoilH3 <> "" Then count = count + 1
    If .CoilH4 <> "" Then count = count + 1
  End With
  CountFilledCoils = count
End Function

'========== シート操作 ==========
Private Sub RecreateSheet(sheetName As String)
  Dim ws As Worksheet, Found As Boolean: Found = False
  
  For Each ws In ThisWorkbook.Worksheets
    If ws.name = sheetName Then
      Found = True
      Exit For
    End If
  Next ws
  
  If Found Then
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
  End If
  
  ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(Sheets.count)
  ActiveSheet.name = sheetName
  ThisWorkbook.Sheets(sheetName).Visible = False
End Sub

Private Sub ShowProgressForm()
  Set progressForm = UserForms.Add("UFProgress")
  progressForm.Show False
End Sub

Private Sub UpdateProgress(current As Long, total As Long, message As String)
  If Not progressForm Is Nothing Then
    With progressForm
      .LabelProgress.Caption = message
      .ProgressBar1.Value = Int((current / total) * 100)
    End With
    DoEvents
  End If
End Sub

Private Sub HideProgressForm()
  If Not progressForm Is Nothing Then
    Unload progressForm
    Set progressForm = Nothing
  End If
End Sub
Private Sub CommandButton2_Click()
Application.ScreenUpdating = False

Dim obSheet As String: obSheet = GetSelectedOBSheetName()

If obSheet <> "" Then
    '--- OBモード: 単体印刷 ---
    If MsgBox(obSheet & " を印刷しますか？", vbYesNo) = vbYes Then
        Call PrintSheet(obSheet)
    End If
Else
    '--- CBモード: 丈1,2 同時印刷 ---
    Dim cbName As String: cbName = GetSelectedCBDisplayName()
    If cbName <> "" Then
        If MsgBox(cbName & " を印刷しますか？", vbYesNo) = vbYes Then
            Call チェック印刷
        End If
    End If
End If

Call クリア
ThisWorkbook.Sheets("起動").Select
Application.ScreenUpdating = True
End Sub
'========== シート印刷 ==========
Private Sub PrintSheet(sheetName As String)
If sheetName = "" Then Exit Sub

Dim ws As Worksheet
On Error Resume Next
Set ws = ThisWorkbook.Sheets(sheetName)
On Error GoTo 0

If ws Is Nothing Then
    MsgBox "指定されたシート「" & sheetName & "」は存在しません。", vbExclamation
    Exit Sub
End If

ws.Visible = True
ws.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
ws.Visible = False
End Sub

Private Sub CommandButton3_Click()
    Application.ScreenUpdating = False
    Dim obIdx As Long
    For obIdx = 1 To 14 Step 2
        ThisWorkbook.Sheets(BuildSheetName(obIdx)).Visible = True
        ThisWorkbook.Sheets(BuildSheetName(obIdx + 1)).Visible = True
    Next obIdx
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton4_Click()
    Application.ScreenUpdating = False
    Dim obIdx As Long
    For obIdx = 1 To 14 Step 2
        ThisWorkbook.Sheets(BuildSheetName(obIdx)).Visible = False
        ThisWorkbook.Sheets(BuildSheetName(obIdx + 1)).Visible = False
    Next obIdx
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton5_Click()
テスラ全サイズ.Show vbModeless
Unload テスラ指定サイズ
Unload UserForm4
End Sub

Private Sub CommandButton6_Click()
With テスラ指定サイズ
 If .TIP1000 = True Then
    If .Take1.Visible = False And .Take2.Visible = False Then
    Exit Sub
    End If
    If .CheckBox1 = False And .CheckBox2 = False And .CheckBox3 = False And .CheckBox4 = _
    False And .CheckBox5 = False And .CheckBox6 = False And .OptionButton1 = False And .OptionButton2 = _
    False And .OptionButton3 = False And .OptionButton4 = False And .OptionButton5 = False And .OptionButton6 = _
    False And .OptionButton7 = False And .OptionButton8 = False And .OptionButton9 = False And .OptionButton10 = _
    False And .OptionButton11 = False And .OptionButton12 = False Then
     MsgBox "コイルサイズが選択されていません"
     Exit Sub
    Else
'     Call 計算1
'     Call 計算2
     Call DB_Materials(.CoilH4, 4)
     Call DB_Materials(.CoilH3, 3)
     Call DB_Materials(.CoilH2, 2)
     Call DB_Materials(.CoilH4, 1)
     Call ラベル転記CK
     Call ラベル転記OP
    End If
 Else
  MsgBox "チップボールチェックがありません"
 End If
End With

  Dim ListName As String, SheetsName As String
  Dim ws As Worksheet, Flag As Boolean
  
    ListName = "風袋計算"
    SheetsName = ListName
 Application.ScreenUpdating = False
    For Each ws In Worksheets
        If ws.name = SheetsName Then Flag = True
    Next ws
    If Flag = True Then
      ThisWorkbook.Sheets(ListName).Visible = True
      ThisWorkbook.Sheets(ListName).Activate
      ActiveWindow.Zoom = 70
    End If
 Application.ScreenUpdating = True

End Sub

Private Sub CommandButton7_Click()
表記用.Show vbModeless
End Sub

Private Sub CommandButton8_Click()
Dim i As Long, tmp As Variant, a As Long
Dim EndRow As Long '最終行判断

EndRow = Cells(Rows.count, 1).End(xlUp).Row
 Application.ScreenUpdating = False
 '全角　半角変換
 '*****************************************************************
        For i = 2 To EndRow
          For a = 2 To 13
            Cells(i, a) = Replace(Cells(i, a), "Ｘ", "Ｘ")
            Cells(i, a) = Replace(Cells(i, a), "ｘ", "Ｘ")
            Cells(i, a) = Replace(Cells(i, a), "X", "Ｘ")
            Cells(i, a) = Replace(Cells(i, a), "x", "Ｘ")
            Cells(i, a) = Replace(Cells(i, a), "×", "Ｘ")
            
            Cells(i, a) = Replace(Cells(i, a), "（", "(")
            Cells(i, a) = Replace(Cells(i, a), "）", ")")
            Cells(i, a) = Replace(Cells(i, a), "０", "0")
            Cells(i, a) = Replace(Cells(i, a), "５", "5")
            Cells(i, a) = Replace(Cells(i, a), "８", "8")
            Cells(i, a) = Replace(Cells(i, a), "４", "4")
            Cells(i, a) = Replace(Cells(i, a), "１", "1")
            Cells(i, a) = Replace(Cells(i, a), "２", "2")
            Cells(i, a) = Replace(Cells(i, a), "３", "3")
            Cells(i, a) = Replace(Cells(i, a), "６", "6")
            Cells(i, a) = Replace(Cells(i, a), "７", "7")
            Cells(i, a) = Replace(Cells(i, a), "９", "9")
          Next a
        Next i
 Application.ScreenUpdating = True
End Sub

Private Sub CommandButton9_Click()
Call 羅列計算
  Dim ListName As String, SheetsName As String
  Dim ws As Worksheet, Flag As Boolean
  
    ListName = "50まで"
    SheetsName = ListName
    
    For Each ws In Worksheets
        If ws.name = SheetsName Then Flag = True
    Next ws
 Application.ScreenUpdating = False
    If Flag = True Then
      ThisWorkbook.Sheets(ListName).Visible = True
      ThisWorkbook.Sheets(ListName).Activate
    End If
 Application.ScreenUpdating = True
End Sub
'
'Private Sub Kei_AfterUpdate()
'Call クリア '入力内容クリア
'With テスラ指定サイズ
'  If .Kei = True Then
'    If .TIP1000 = True Then
'       If .Take1.Visible = False And .Take2.Visible = False Then
'        .Kei = False '初期化
'        Exit Sub
'       End If
'       If .CheckBox1 = False And .CheckBox2 = False And .CheckBox3 = False And .CheckBox4 = _
'       False And .CheckBox5 = False And .CheckBox6 = False And .OptionButton1 = False And .OptionButton2 = _
'       False And .OptionButton3 = False And .OptionButton4 = False And .OptionButton5 = False And .OptionButton6 = _
'       False And .OptionButton7 = False And .OptionButton8 = False And .OptionButton9 = False And .OptionButton10 = _
'       False And .OptionButton11 = False And .OptionButton12 = False Then
'        MsgBox "コイルサイズが選択されていません"
'        .Kei = False '初期化
'        Exit Sub
'       Else
'
'        Call 計算1
'        Call 計算2
'
'        Call ラベル転記CK
'        Call ラベル転記OP
'
'       End If
'    Else
'     MsgBox "チップボールチェックがありません"
'      .Kei = False '初期化
'      Exit Sub
'    End If
'  End If
'End With
'End Sub

'Private Sub Frame3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'警告.Show vbModeless
'End Sub
Private Sub txtKensaNo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限
Select Case KeyAscii.Value
Case 48 To 57 '0~9
Case 65 To 90 'A~Z
Case 97 To 122 'a~z
KeyAscii.Value = KeyAscii.Value - 32 '大文字強制変換
Case Else
KeyAscii.Value = 0
End Select
End Sub

Private Sub txtWeight1_Change()
If テスラ指定サイズ.txtWeight1 = "-" Then Exit Sub
If Not IsNumeric(テスラ指定サイズ.txtWeight1.text) Then テスラ指定サイズ.txtWeight1.text = ""
End Sub

Private Sub txtWeight2_Change()
If テスラ指定サイズ.txtWeight2 = "-" Then Exit Sub
If Not IsNumeric(テスラ指定サイズ.txtWeight2.text) Then テスラ指定サイズ.txtWeight2.text = ""
End Sub

Private Sub TIP1000_Click()
'警告.Show vbModeless
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

 'If MsgBox("フォームを閉じますか？", vbYesNo) = vbYes Then
 ThisWorkbook.Activate
 ThisWorkbook.Sheets("起動").Select
 If UserForm4.Visible = True Then Unload UserForm4
 If 警告.Visible = True Then Unload 警告
 'End If
 
End Sub

