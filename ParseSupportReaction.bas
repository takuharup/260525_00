Attribute VB_Name = "ParseSupportReaction"
Option Explicit

Public Sub ParseSupportReaction()

    Dim filePath As String
    Dim fileNum As Integer
    Dim ws As Worksheet
    Dim sheetName As String
    Dim line As String
    Dim trimmedLine As String
    Dim parts() As String

    Dim loadNumber As Long
    Dim loadName As String
    Dim rowIdx As Long
    Dim lineNum As Long

    Dim nodeStr As String
    Dim rx As Double, ry As Double, rz As Double
    Dim rmx As Double, rmy As Double, rmz As Double

    ' --- ファイル選択 ---
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "支点反力テキストファイルを選択してください"
    fd.Filters.Clear
    fd.Filters.Add "テキストファイル", "*.txt"
    fd.AllowMultiSelect = False

    If fd.Show <> True Then
        Exit Sub
    End If
    filePath = fd.SelectedItems(1)

    ' --- 新規シート作成 ---
    sheetName = "支点反力_" & Format(Now, "YYYYMMDD_HHMMSS")
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName

    ' --- ヘッダー行 ---
    ws.Cells(1, 1).Value = "荷重番号"
    ws.Cells(1, 2).Value = "荷重名称"
    ws.Cells(1, 3).Value = "節点番号"
    ws.Cells(1, 4).Value = "RX"
    ws.Cells(1, 5).Value = "RY"
    ws.Cells(1, 6).Value = "RZ"
    ws.Cells(1, 7).Value = "RMX"
    ws.Cells(1, 8).Value = "RMY"
    ws.Cells(1, 9).Value = "RMZ"

    rowIdx = 2
    lineNum = 0
    loadNumber = 0
    loadName = ""

    ' --- ファイル読み込み ---
    fileNum = FreeFile
    On Error GoTo FileOpenError
    Open filePath For Input As #fileNum
    On Error GoTo 0

    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        lineNum = lineNum + 1
        trimmedLine = Trim(line)

        ' 空行スキップ
        If Len(trimmedLine) = 0 Then GoTo NextLine

        ' ===== 区切り線スキップ
        If Left(trimmedLine, 5) = "=====" Then GoTo NextLine

        ' 荷重番号・荷重名称行
        If InStr(line, "荷重番号") > 0 Then
            loadNumber = ExtractLoadNumber(line)
            loadName = ExtractLoadName(line)
            GoTo NextLine
        End If

        ' 節点番号ヘッダー行スキップ
        If InStr(trimmedLine, "節点番号") > 0 Then GoTo NextLine

        ' 合計行
        If Left(trimmedLine, 2) = "合計" Then
            parts = SplitNormalized(trimmedLine)
            If UBound(parts) >= 6 Then
                If Not ValidateNumericParts(parts, 1, 6, lineNum) Then GoTo NextLine
                rx  = CDbl(parts(1))
                ry  = CDbl(parts(2))
                rz  = CDbl(parts(3))
                rmx = CDbl(parts(4))
                rmy = CDbl(parts(5))
                rmz = CDbl(parts(6))
                WriteRow ws, rowIdx, loadNumber, loadName, "合計", rx, ry, rz, rmx, rmy, rmz
                rowIdx = rowIdx + 1
            Else
                Debug.Print "Warning: 合計行のフィールド数不足 (line " & lineNum & "): " & line
            End If
            GoTo NextLine
        End If

        ' 節点データ行（先頭が数値）
        If IsNumeric(Left(trimmedLine, InStr(trimmedLine & " ", " ") - 1)) Then
            parts = SplitNormalized(trimmedLine)
            If UBound(parts) >= 6 Then
                If Not IsNumeric(parts(0)) Then GoTo NextLine
                If Not ValidateNumericParts(parts, 1, 6, lineNum) Then GoTo NextLine
                nodeStr = parts(0)
                rx  = CDbl(parts(1))
                ry  = CDbl(parts(2))
                rz  = CDbl(parts(3))
                rmx = CDbl(parts(4))
                rmy = CDbl(parts(5))
                rmz = CDbl(parts(6))
                WriteRow ws, rowIdx, loadNumber, loadName, nodeStr, rx, ry, rz, rmx, rmy, rmz
                rowIdx = rowIdx + 1
            Else
                Debug.Print "Warning: 節点行のフィールド数不足 (line " & lineNum & "): " & line
            End If
        End If

NextLine:
    Loop

    Close #fileNum

    Dim dataRows As Long
    dataRows = rowIdx - 2
    MsgBox "完了しました。" & vbCrLf & _
           "出力シート名: " & sheetName & vbCrLf & _
           "データ行数: " & dataRows & " 行", _
           vbInformation, "ParseSupportReaction"
    Exit Sub

FileOpenError:
    MsgBox "ファイルを開けませんでした。" & vbCrLf & Err.Description, vbCritical, "ParseSupportReaction"

End Sub

' -------------------------------------------------------
' 連続スペースを正規化してSplit
' -------------------------------------------------------
Private Function SplitNormalized(ByVal s As String) As String()
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    SplitNormalized = Split(Trim(s), " ")
End Function

' -------------------------------------------------------
' 荷重番号を抽出: "荷重番号 =  100" の数値部分
' -------------------------------------------------------
Private Function ExtractLoadNumber(ByVal line As String) As Long
    Dim pos As Long
    Dim token As String
    pos = InStr(line, "荷重番号")
    If pos = 0 Then
        ExtractLoadNumber = 0
        Exit Function
    End If
    ' "荷重番号 = " 以降から荷重名称 手前まで取り出す
    ' 同行に "荷重名称" があるため、その手前で切り取る
    Dim sub1 As String
    Dim posName As Long
    posName = InStr(line, "荷重名称")
    If posName > 0 Then
        sub1 = Mid(line, pos, posName - pos)
    Else
        sub1 = Mid(line, pos)
    End If
    ' "=" 以降をTrim
    Dim eqPos As Long
    eqPos = InStr(sub1, "=")
    If eqPos = 0 Then
        ExtractLoadNumber = 0
        Exit Function
    End If
    token = Trim(Mid(sub1, eqPos + 1))
    ' 先頭の数値トークンを取り出す
    token = SplitNormalized(token)(0)
    If IsNumeric(token) Then
        ExtractLoadNumber = CLng(token)
    Else
        ExtractLoadNumber = 0
    End If
End Function

' -------------------------------------------------------
' 荷重名称を抽出: "荷重名称 =  （文字列）" の文字列部分
' -------------------------------------------------------
Private Function ExtractLoadName(ByVal line As String) As String
    Dim pos As Long
    pos = InStr(line, "荷重名称")
    If pos = 0 Then
        ExtractLoadName = ""
        Exit Function
    End If
    Dim sub1 As String
    sub1 = Mid(line, pos)
    Dim eqPos As Long
    eqPos = InStr(sub1, "=")
    If eqPos = 0 Then
        ExtractLoadName = ""
        Exit Function
    End If
    ExtractLoadName = Trim(Mid(sub1, eqPos + 1))
End Function

' -------------------------------------------------------
' 指定範囲のparts要素が全てNumericか検証
' -------------------------------------------------------
Private Function ValidateNumericParts(ByRef parts() As String, _
                                       ByVal fromIdx As Integer, _
                                       ByVal toIdx As Integer, _
                                       ByVal lineNum As Long) As Boolean
    Dim i As Integer
    For i = fromIdx To toIdx
        If Not IsNumeric(parts(i)) Then
            Debug.Print "Warning: 数値変換失敗 parts(" & i & ")='" & parts(i) & "' (line " & lineNum & ")"
            ValidateNumericParts = False
            Exit Function
        End If
    Next i
    ValidateNumericParts = True
End Function

' -------------------------------------------------------
' 節点番号フィルタ抽出
' -------------------------------------------------------
Public Sub FilterByNode()

    Dim srcSheetName As String
    Dim srcWs As Worksheet
    Dim dstWs As Worksheet
    Dim dstSheetName As String

    Dim nodeInput As String
    Dim nodeTokens() As String
    Dim i As Integer

    Dim lastRow As Long
    Dim dstRow As Long
    Dim cellVal As String

    ' --- ステップ① 元テーブルシート選択 ---
    srcSheetName = InputBox("抽出元のシート名を入力してください" & vbCrLf & _
                            "（例：支点反力_20260525_143022）", "FilterByNode")
    If StrPtr(srcSheetName) = 0 Then Exit Sub       ' キャンセル
    If Trim(srcSheetName) = "" Then Exit Sub        ' 空入力もサイレント終了

    On Error Resume Next
    Set srcWs = ThisWorkbook.Worksheets(srcSheetName)
    On Error GoTo 0
    If srcWs Is Nothing Then
        MsgBox "シート「" & srcSheetName & "」が見つかりません。", vbCritical, "FilterByNode"
        Exit Sub
    End If

    ' --- ステップ② 節点番号リスト入力 ---
    nodeInput = InputBox("抽出する節点番号をカンマ区切りで入力してください" & vbCrLf & _
                         "（例： 1,3,合計　または　合計）", "FilterByNode")
    If StrPtr(nodeInput) = 0 Then Exit Sub          ' キャンセル
    If Trim(nodeInput) = "" Then
        MsgBox "節点番号が入力されていません。", vbExclamation, "FilterByNode"
        Exit Sub
    End If

    ' カンマSplitして各要素をTrim
    nodeTokens = Split(nodeInput, ",")
    For i = 0 To UBound(nodeTokens)
        nodeTokens(i) = Trim(nodeTokens(i))
    Next i

    ' --- 新規シート作成 ---
    dstSheetName = "抽出_" & Format(Now, "YYYYMMDD_HHMMSS")
    Set dstWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    dstWs.Name = dstSheetName

    ' --- ヘッダーコピー ---
    dstWs.Rows(1).Value = srcWs.Rows(1).Value
    dstRow = 2

    ' --- データ行走査 ---
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row

    Dim extractCount As Long
    extractCount = 0

    For i = 2 To lastRow
        cellVal = Trim(CStr(srcWs.Cells(i, 3).Value))
        If IsInList(cellVal, nodeTokens) Then
            dstWs.Rows(dstRow).Value = srcWs.Rows(i).Value
            dstRow = dstRow + 1
            extractCount = extractCount + 1
        End If
    Next i

    ' --- 完了メッセージ ---
    If extractCount = 0 Then
        MsgBox "該当する節点番号の行がありませんでした。" & vbCrLf & _
               "出力シート名: " & dstSheetName, vbExclamation, "FilterByNode"
    Else
        MsgBox "完了しました。" & vbCrLf & _
               "出力シート名: " & dstSheetName & vbCrLf & _
               "抽出行数: " & extractCount & " 行", vbInformation, "FilterByNode"
    End If

End Sub

' -------------------------------------------------------
' 節点番号リスト一致判定
' -------------------------------------------------------
Private Function IsInList(ByVal target As String, ByRef list() As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(list)
        If Trim(list(i)) = Trim(target) Then
            IsInList = True
            Exit Function
        End If
    Next i
    IsInList = False
End Function

' -------------------------------------------------------
' 1行分をシートへ書き込む
' -------------------------------------------------------
Private Sub WriteRow(ByVal ws As Worksheet, ByVal rowIdx As Long, _
                     ByVal loadNum As Long, ByVal loadNm As String, _
                     ByVal nodeVal As String, _
                     ByVal rx As Double, ByVal ry As Double, ByVal rz As Double, _
                     ByVal rmx As Double, ByVal rmy As Double, ByVal rmz As Double)
    ws.Cells(rowIdx, 1).Value = loadNum
    ws.Cells(rowIdx, 2).Value = loadNm
    ' 節点番号が数値文字列なら Long、"合計"ならそのまま文字列
    If IsNumeric(nodeVal) Then
        ws.Cells(rowIdx, 3).Value = CLng(nodeVal)
    Else
        ws.Cells(rowIdx, 3).Value = nodeVal
    End If
    ws.Cells(rowIdx, 4).Value = rx
    ws.Cells(rowIdx, 5).Value = ry
    ws.Cells(rowIdx, 6).Value = rz
    ws.Cells(rowIdx, 7).Value = rmx
    ws.Cells(rowIdx, 8).Value = rmy
    ws.Cells(rowIdx, 9).Value = rmz
End Sub
