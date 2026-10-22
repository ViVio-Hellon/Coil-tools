VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} テスラ全サイズ 
   Caption         =   "UserForm2"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12135
   OleObjectBlob   =   "テスラ全サイズ.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "テスラ全サイズ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()
  With テスラ全サイズ
      If .Controls("CheckBox" & 1) = True Then
         .Controls("CheckBox" & 2) = False
         .Controls("CheckBox" & 1) = True
      End If
  End With
  
   Me.Label24.Visible = False
   Me.Label25.Visible = False
   
   Me.TextBox4 = ""
   Me.TextBox5 = ""
   
   Me.TextBox4.Visible = False
   Me.TextBox5.Visible = False
   
   Me.Label10 = "重量　丈１"
   Me.Label9 = "重量　丈2"

   Me.Label28 = ""
   Me.Label29 = ""
   
   Me.Label26 = ""
   Me.Label27 = ""

   Me.Frame1.Caption = "丈1"
   Me.Frame2.Caption = "丈2"

End Sub

Private Sub CheckBox2_Click()
  With テスラ全サイズ
      If .Controls("CheckBox" & 2) = True Then
         .Controls("CheckBox" & 1) = False
         .Controls("CheckBox" & 2) = True
      End If
  End With
  
   Me.Label24.Visible = True
   Me.Label25.Visible = True
   
   Me.TextBox4.Visible = True
   Me.TextBox5.Visible = True
   
   Me.Label10 = "重量　１-1-？"
   Me.Label9 = "重量　2-1-？"
   
   Me.Frame1.Caption = "元コイル巾位置　1"
   Me.Frame2.Caption = "元コイル巾位置　2"
  
End Sub

Private Sub ComboBox1_Change()
'style　直接入力禁止
テスラ全サイズ.Label19 = ""
テスラ全サイズ.Label21 = ""
テスラ全サイズ.Label16 = ""
テスラ全サイズ.Label18 = ""
テスラ全サイズ.Label17 = ""
テスラ全サイズ.Label22 = ""
テスラ全サイズ.Label26 = ""
テスラ全サイズ.Label27 = ""
    Application.ScreenUpdating = False
    
        Call シート追加1
        Call コピー1
        Call シート追加2
        Call コピー2
        
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton1_Click()
'サイズ追加時　追加項目
'印刷範囲指定、型番サイズ記載
  Dim ListName As String
  Dim ListName2 As String
    ListName = (テスラ全サイズ.ComboBox1.text & "1")
    ListName2 = (テスラ全サイズ.ComboBox1.text & "2")
    
    
  Dim ws As Worksheet, Flag As Boolean
  Dim flag2 As Boolean
    
    For Each ws In Worksheets
        If ws.name = ListName Then Flag = True
    Next ws
    
    For Each ws In Worksheets
        If ws.name = ListName2 Then flag2 = True
    Next ws
    
  If Flag = False Then  'シートがないときはなにもしない
      MsgBox "おちつけシートがねぇぞ"
  ElseIf flag2 = False Then
      MsgBox "おちつけシートがねぇぞ"
  Else
        
    If Me.CheckBox1 = False And Me.CheckBox2 = False Then
      MsgBox "2ケタ、3ケタどちらにもチェックが入っていません"
    ElseIf Me.CheckBox1 = True Then
        If IsNumeric(Mid(テスラ全サイズ.TextBox1, 2, 1)) = False Then
              MsgBox "検査NOが7桁入力されていない" & Chr(10) & "もしくは6桁目にアルファベットが入力されています"
            ElseIf IsNumeric(Mid(テスラ全サイズ.TextBox1, 7, 1)) = False Then
              MsgBox "検査NOが7桁入力されていない" & Chr(10) & "もしくは1桁目にアルファベットが入力されています"
            ElseIf IsNumeric(Mid(テスラ全サイズ.TextBox1, 4, 4)) = False Then
              MsgBox "検査NOが7桁入力されていない" & Chr(10) & "もしくは1～4桁目にアルファベットが入力されています"
            ElseIf IsNumeric(Mid(テスラ全サイズ.TextBox1, 1, 1)) = True Then
              MsgBox "検査NOが7桁入力されていない" & Chr(10) & "もしくは7桁目に数字が入力されています"
    '        ElseIf Me.TextBox2 = "" Or Me.TextBox3 = "" Then
    '          MsgBox "重量が入力されていません"
        Else
        Call 入力
        Call 型番入力
        Call 型番表示
        End If
    ElseIf Me.CheckBox2 = True Then
        If IsNumeric(Mid(テスラ全サイズ.TextBox1, 2, 1)) = False Then
              MsgBox "検査NOが7桁入力されていない" & Chr(10) & "もしくは6桁目にアルファベットが入力されています"
            ElseIf IsNumeric(Mid(テスラ全サイズ.TextBox1, 7, 1)) = False Then
              MsgBox "検査NOが7桁入力されていない" & Chr(10) & "もしくは1桁目にアルファベットが入力されています"
            ElseIf IsNumeric(Mid(テスラ全サイズ.TextBox1, 4, 4)) = False Then
              MsgBox "検査NOが7桁入力されていない" & Chr(10) & "もしくは1～4桁目にアルファベットが入力されています"
            ElseIf IsNumeric(Mid(テスラ全サイズ.TextBox1, 1, 1)) = True Then
              MsgBox "検査NOが7桁入力されていない" & Chr(10) & "もしくは7桁目に数字が入力されています"
    '        ElseIf Me.TextBox2 = "" Or Me.TextBox3 = "" Then
    '          MsgBox "重量が入力されていません"
    '        ElseIf Me.TextBox4 = "" Or Me.TextBox5 = "" Then
    '          MsgBox "重量が入力されていません"
        Else
        Call 入力
        Call 型番入力
        Call 型番表示
        End If
    Else
    End If
'フォーム表示更新
        If テスラ全サイズ.CheckBox2 = True Then
        
            テスラ全サイズ.TextBox2 = ""
            テスラ全サイズ.TextBox3 = ""
            テスラ全サイズ.TextBox4 = ""
            テスラ全サイズ.TextBox5 = ""
        
            With ThisWorkbook.Sheets(ListName)
              テスラ全サイズ.Label16 = "「Coil1」" & Chr(10) & .Cells(92, 22) '検査番号
              テスラ全サイズ.Label17 = .Cells(92, 37) & "kg"  '重量1-1-?
              If テスラ全サイズ.Label17 = "kg" Then
                 テスラ全サイズ.Label17 = ""
              End If
              テスラ全サイズ.Label26 = .Cells(185, 37) & "kg"  '重量1-2-?
              If テスラ全サイズ.Label26 = "kg" Then
                 テスラ全サイズ.Label26 = ""
              End If
              テスラ全サイズ.Label19 = "「Sheet1」" & Chr(10) & .Cells(92, 2) 'サイズ1-1-?
              テスラ全サイズ.Label28 = .Cells(185, 2) 'サイズ1-2-?
             End With
             With ThisWorkbook.Sheets(ListName2)
              テスラ全サイズ.Label18 = "「Coil2」" & Chr(10) & .Cells(92, 22) '検査番号
              テスラ全サイズ.Label22 = .Cells(92, 37) & "kg"  '重量2-1-?
              If テスラ全サイズ.Label22 = "kg" Then
                 テスラ全サイズ.Label22 = ""
              End If
              テスラ全サイズ.Label27 = .Cells(185, 37) & "kg"  '重量2-2-?
              If テスラ全サイズ.Label27 = "kg" Then
                 テスラ全サイズ.Label27 = ""
              End If
              テスラ全サイズ.Label21 = "「Sheet2」" & Chr(10) & .Cells(92, 2) 'サイズ2-1-?
              テスラ全サイズ.Label29 = .Cells(185, 2) 'サイズ2-2-?
             End With
             
        ElseIf テスラ全サイズ.CheckBox1 = True Then
        
            テスラ全サイズ.TextBox2 = ""
            テスラ全サイズ.TextBox3 = ""
            テスラ全サイズ.TextBox4 = ""
            テスラ全サイズ.TextBox5 = ""
        
            With ThisWorkbook.Sheets(ListName)
              テスラ全サイズ.Label16 = "「Coil1」" & Chr(10) & .Cells(92, 22) '検査番号
              テスラ全サイズ.Label17 = .Cells(92, 37) & "kg"  '重量1-1-?
              If テスラ全サイズ.Label17 = "kg" Then
                 テスラ全サイズ.Label17 = ""
              End If
              テスラ全サイズ.Label26 = ""  '重量1-2-?
              テスラ全サイズ.Label19 = "「Sheet1」" & Chr(10) & .Cells(92, 2) 'サイズ1-1-?
              テスラ全サイズ.Label28 = "" 'サイズ1-2-?
             End With
             With ThisWorkbook.Sheets(ListName2)
              テスラ全サイズ.Label18 = "「Coil2」" & Chr(10) & .Cells(92, 22) '検査番号
              テスラ全サイズ.Label22 = .Cells(92, 37) & "kg"  '重量2-1-?
              If テスラ全サイズ.Label22 = "kg" Then
                 テスラ全サイズ.Label22 = ""
              End If
              テスラ全サイズ.Label27 = ""   '重量2-2-?
              テスラ全サイズ.Label21 = "「Sheet2」" & Chr(10) & .Cells(92, 2) 'サイズ2-1-?
              テスラ全サイズ.Label29 = "" 'サイズ2-2-?
             End With
             
         End If
         
  End If
End Sub

Private Sub CommandButton2_Click()
Dim ListName As String: ListName = (テスラ全サイズ.ComboBox1.text & "1")
Dim ListName2 As String: ListName2 = (テスラ全サイズ.ComboBox1.text & "2")
If MsgBox(sheetName & "  " & "印刷する？", vbYesNo) = vbYes Then
  If ThisWorkbook.Sheets(ListName).Cells(92, 37) = "" And ThisWorkbook.Sheets(ListName).Cells(185, 37) = "" And ThisWorkbook.Sheets(ListName2).Cells(92, 37) = "" And ThisWorkbook.Sheets(ListName2).Cells(185, 37) = "" Then
      MsgBox "重量インプットがありません"
      Exit Sub
  ElseIf ThisWorkbook.Sheets(ListName).Cells(92, 37) <> "" And ThisWorkbook.Sheets(ListName).Cells(185, 37) <> "" And ThisWorkbook.Sheets(ListName2).Cells(92, 37) <> "" And ThisWorkbook.Sheets(ListName2).Cells(185, 37) <> "" Then
      ThisWorkbook.Sheets(ListName).Activate
      Call 印刷範囲1
      Call 印刷実行
      ThisWorkbook.Sheets(ListName2).Activate
      Call 印刷範囲2
      Call 印刷実行
  ElseIf ThisWorkbook.Sheets(ListName).Cells(92, 37) <> "" And ThisWorkbook.Sheets(ListName2).Cells(92, 37) <> "" Then
      ThisWorkbook.Sheets(ListName).Activate
      Call 印刷範囲1
      Call 印刷実行
      ThisWorkbook.Sheets(ListName2).Activate
      Call 印刷範囲2
      Call 印刷実行
  ElseIf ThisWorkbook.Sheets(ListName).Cells(92, 37) <> "" And ThisWorkbook.Sheets(ListName2).Cells(185, 37) <> "" Then
      ThisWorkbook.Sheets(ListName).Activate
      Call 印刷範囲1
      Call 印刷実行
      ThisWorkbook.Sheets(ListName2).Activate
      Call 印刷範囲2
      Call 印刷実行
  ElseIf ThisWorkbook.Sheets(ListName).Cells(185, 37) <> "" And ThisWorkbook.Sheets(ListName2).Cells(185, 37) <> "" Then
      ThisWorkbook.Sheets(ListName).Activate
      Call 印刷範囲1
      Call 印刷実行
      ThisWorkbook.Sheets(ListName2).Activate
      Call 印刷範囲2
      Call 印刷実行
  ElseIf ThisWorkbook.Sheets(ListName).Cells(185, 37) <> "" And ThisWorkbook.Sheets(ListName2).Cells(185, 37) <> "" Then
      ThisWorkbook.Sheets(ListName).Activate
      Call 印刷範囲1
      Call 印刷実行
      ThisWorkbook.Sheets(ListName2).Activate
      Call 印刷範囲2
      Call 印刷実行
  ElseIf ThisWorkbook.Sheets(ListName).Cells(92, 37) <> "" And ThisWorkbook.Sheets(ListName).Cells(185, 37) <> "" And ThisWorkbook.Sheets(ListName2).Cells(92, 37) <> "" Then
      ThisWorkbook.Sheets(ListName).Activate
      Call 印刷範囲1
      Call 印刷実行
      ThisWorkbook.Sheets(ListName2).Activate
      Call 印刷範囲2
      Call 印刷実行
  ElseIf ThisWorkbook.Sheets(ListName).Cells(92, 37) <> "" And ThisWorkbook.Sheets(ListName).Cells(185, 37) <> "" And ThisWorkbook.Sheets(ListName2).Cells(185, 37) <> "" Then
      ThisWorkbook.Sheets(ListName).Activate
      Call 印刷範囲1
      Call 印刷実行
      ThisWorkbook.Sheets(ListName2).Activate
      Call 印刷範囲2
      Call 印刷実行
  ElseIf ThisWorkbook.Sheets(ListName).Cells(92, 37) <> "" And ThisWorkbook.Sheets(ListName2).Cells(92, 37) = "" And ThisWorkbook.Sheets(ListName2).Cells(185, 37) = "" Then
      ThisWorkbook.Sheets(ListName).Activate
      Call 印刷範囲1
      Call 印刷実行
  ElseIf ThisWorkbook.Sheets(ListName).Cells(185, 37) <> "" And ThisWorkbook.Sheets(ListName2).Cells(92, 37) = "" And ThisWorkbook.Sheets(ListName2).Cells(185, 37) = "" Then
      ThisWorkbook.Sheets(ListName).Activate
      Call 印刷範囲1
      Call 印刷実行
  ElseIf ThisWorkbook.Sheets(ListName).Cells(92, 37) <> "" And ThisWorkbook.Sheets(ListName).Cells(185, 37) <> "" And ThisWorkbook.Sheets(ListName2).Cells(92, 37) = "" And ThisWorkbook.Sheets(ListName2).Cells(185, 37) = "" Then
      ThisWorkbook.Sheets(ListName).Activate
      Call 印刷範囲1
      Call 印刷実行
  ElseIf ThisWorkbook.Sheets(ListName2).Cells(92, 37) <> "" And ThisWorkbook.Sheets(ListName2).Cells(185, 37) <> "" And ThisWorkbook.Sheets(ListName).Cells(92, 37) = "" And ThisWorkbook.Sheets(ListName).Cells(185, 37) = "" Then
      ThisWorkbook.Sheets(ListName2).Activate
      Call 印刷範囲2
      Call 印刷実行
  ElseIf ThisWorkbook.Sheets(ListName2).Cells(92, 37) <> "" And ThisWorkbook.Sheets(ListName).Cells(92, 37) = "" And ThisWorkbook.Sheets(ListName).Cells(185, 37) = "" Then
      ThisWorkbook.Sheets(ListName2).Activate
      Call 印刷範囲2
      Call 印刷実行
  ElseIf ThisWorkbook.Sheets(ListName2).Cells(185, 37) <> "" And ThisWorkbook.Sheets(ListName).Cells(92, 37) = "" And ThisWorkbook.Sheets(ListName).Cells(185, 37) = "" Then
      ThisWorkbook.Sheets(ListName2).Activate
      Call 印刷範囲2
      Call 印刷実行
  End If
End If
End Sub

Private Sub CommandButton3_Click()
Worksheets(Worksheets.count).Select
  If MsgBox("削除する？", vbYesNo) = vbYes Then
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
  End If
End Sub

Private Sub CommandButton4_Click()

Unload テスラ全サイズ

テスラ全サイズ.Show vbModeless

End Sub

Private Sub CommandButton5_Click()
  Dim ListName As String
  Dim ListName2 As String
    ListName = (テスラ全サイズ.ComboBox1.text & "1")
    ListName2 = (テスラ全サイズ.ComboBox1.text & "2")
If テスラ全サイズ.ComboBox1.text = "" Then
Else
Worksheets(ListName).Select
   MsgBox "sheet" & ListName & "を開きました"
End If
End Sub

Private Sub CommandButton6_Click()
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

Private Sub CommandButton7_Click()
Call 印刷範囲1
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub TextBox2_Change()
If テスラ全サイズ.TextBox2 = "-" Then Exit Sub
If Not IsNumeric(テスラ全サイズ.TextBox2.text) Then テスラ全サイズ.TextBox2.text = ""
End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub TextBox3_Change()
If テスラ全サイズ.TextBox3 = "-" Then Exit Sub
If Not IsNumeric(テスラ全サイズ.TextBox3.text) Then テスラ全サイズ.TextBox3.text = ""
End Sub
Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub TextBox4_Change()
If テスラ全サイズ.TextBox4 = "-" Then Exit Sub
If Not IsNumeric(テスラ全サイズ.TextBox4.text) Then テスラ全サイズ.TextBox4.text = ""
End Sub
Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub TextBox5_Change()
If テスラ全サイズ.TextBox5 = "-" Then Exit Sub
If Not IsNumeric(テスラ全サイズ.TextBox5.text) Then テスラ全サイズ.TextBox5.text = ""
End Sub

Private Sub UserForm_Initialize()

  Dim i As Integer
  
   Me.ComboBox1.Clear
   
  With テスラ全サイズ
     With Worksheets("型番")
        For i = 2 To .Cells(Rows.count, 1).End(xlUp).Row
         Me.ComboBox1.AddItem .Cells(i, 1)
        Next i
     End With
  End With
  
  
End Sub
