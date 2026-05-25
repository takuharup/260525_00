Attribute VB_Name = "Module1"
Option Explicit

Private Type ReactionRow
    LoadNo   As Long
    LoadName As String
    NodeNo   As String
    RX       As Double
    RY       As Double
    RZ       As Double
    RMX      As Double
    RMY      As Double
    RMZ      As Double
End Type

' -------------------------------------------------------
' UserForm1をVBAプロジェクトへ恒久作成（import後に1回だけ実行）
' -------------------------------------------------------
Public Sub CreateUserForm()

    Dim vbp As Object
    Set vbp = ThisWorkbook.VBProject

    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents("UserForm1")
    On Error GoTo 0

    Dim vbc As Object
    On Error Resume Next
    Set vbc = vbp.VBComponents.Add(3)   ' vbext_ct_MSForm
    If Err.Number <> 0 Then
        MsgBox "VBAプロジェクトへのアクセスが許可されていません。" & vbCrLf & _
               "「ファイル」→「オプション」→「トラストセンター」→" & vbCrLf & _
               "「トラストセンターの設定」→「マクロの設定」で" & vbCrLf & _
               "「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」" & vbCrLf & _
               "を有効にしてから再実行してください。", vbCritical, "CreateUserForm"
        Exit Sub
    End If
    On Error GoTo 0

    vbc.Name = "UserForm1"
    vbc.Properties("Caption")         = "節点番号選択"
    vbc.Properties("Width")           = 282
    vbc.Properties("Height")          = 372
    vbc.Properties("StartUpPosition") = 1

    Dim lbl As Object
    Set lbl = vbc.Designer.Controls.Add("Forms.Label.1")
    lbl.Caption = "転記する節点番号を選択してください（複数選択可）"
    lbl.Left = 6 : lbl.Top = 6 : lbl.Width = 264 : lbl.Height = 18

    Dim lst As Object
    Set lst = vbc.Designer.Controls.Add("Forms.ListBox.1")
    lst.Name = "lstNodes"
    lst.Left = 6 : lst.Top = 30 : lst.Width = 264 : lst.Height = 246
    lst.MultiSelect = 1

    Dim btnOK As Object
    Set btnOK = vbc.Designer.Controls.Add("Forms.CommandButton.1")
    btnOK.Name = "btnOK"
    btnOK.Caption = "OK"
    btnOK.Left = 60 : btnOK.Top = 288 : btnOK.Width = 72 : btnOK.Height = 24

    Dim btnCancel As Object
    Set btnCancel = vbc.Designer.Controls.Add("Forms.CommandButton.1")
    btnCancel.Name = "btnCancel"
    btnCancel.Caption = "キャンセル"
    btnCancel.Left = 156 : btnCancel.Top = 288 : btnCancel.Width = 90 : btnCancel.Height = 24

    Dim fc As String
    fc = "Private Sub btnOK_Click()" & vbCrLf & _
         "    Me.Tag = ""OK"" : Me.Hide" & vbCrLf & _
         "End Sub" & vbCrLf & _
         "Private Sub btnCancel_Click()" & vbCrLf & _
         "    Me.Tag = ""Cancel"" : Me.Hide" & vbCrLf & _
         "End Sub" & vbCrLf & _
         "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf & _
         "    If CloseMode = vbFormControlMenu Then" & vbCrLf & _
         "        Me.Tag = ""Cancel"" : Me.Hide : Cancel = True" & vbCrLf & _
         "    End If" & vbCrLf & _
         "End Sub"
    vbc.CodeModule.AddFromString fc

    ThisWorkbook.Save
    MsgBox "UserForm1 を作成しました。次回から ParseAndSelectNodes をそのまま実行できます。", _
           vbInformation, "CreateUserForm"

End Sub

' -------------------------------------------------------
' TXTを解析して新規シートへ出力
' -------------------------------------------------------
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

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "支点反力テキストファイルを選択してください"
    fd.Filters.Clear
    fd.Filters.Add "テキストファイル", "*.txt"
    fd.AllowMultiSelect = False

    If fd.Show <> True Then Exit Sub
    filePath = fd.SelectedItems(1)

    sheetName = "支点反力_" & Format(Now, "YYYYMMDD_HHMMSS")
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName

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

    fileNum = FreeFile
    On Error GoTo FileOpenError
    Open filePath For Input As #fileNum
    On Error GoTo 0

    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        lineNum = lineNum + 1
        trimmedLine = Trim(line)

        If Len(trimmedLine) = 0 Then GoTo NextLine
        If Left(trimmedLine, 5) = "=====" Then GoTo NextLine

        If InStr(line, "荷重番号") > 0 Then
            loadNumber = ExtractLoadNumber(line)
            loadName = ExtractLoadName(line)
            GoTo NextLine
        End If

        If InStr(trimmedLine, "節点番号") > 0 Then GoTo NextLine

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

    MsgBox "完了しました。" & vbCrLf & _
           "出力シート名: " & sheetName & vbCrLf & _
           "データ行数: " & (rowIdx - 2) & " 行", _
           vbInformation, "ParseSupportReaction"
    Exit Sub

FileOpenError:
    MsgBox "ファイルを開けませんでした。" & vbCrLf & Err.Description, vbCritical, "ParseSupportReaction"

End Sub

' -------------------------------------------------------
' 既存テーブルを節点番号でフィルタして新規シートへ出力
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

    srcSheetName = InputBox("抽出元のシート名を入力してください" & vbCrLf & _
                            "（例：支点反力_20260525_143022）", "FilterByNode")
    If StrPtr(srcSheetName) = 0 Then Exit Sub
    If Trim(srcSheetName) = "" Then Exit Sub

    On Error Resume Next
    Set srcWs = ThisWorkbook.Worksheets(srcSheetName)
    On Error GoTo 0
    If srcWs Is Nothing Then
        MsgBox "シート「" & srcSheetName & "」が見つかりません。", vbCritical, "FilterByNode"
        Exit Sub
    End If

    nodeInput = InputBox("抽出する節点番号をカンマ区切りで入力してください" & vbCrLf & _
                         "（例： 1,3,合計　または　合計）", "FilterByNode")
    If StrPtr(nodeInput) = 0 Then Exit Sub
    If Trim(nodeInput) = "" Then
        MsgBox "節点番号が入力されていません。", vbExclamation, "FilterByNode"
        Exit Sub
    End If

    nodeTokens = Split(nodeInput, ",")
    For i = 0 To UBound(nodeTokens)
        nodeTokens(i) = Trim(nodeTokens(i))
    Next i

    dstSheetName = "抽出_" & Format(Now, "YYYYMMDD_HHMMSS")
    Set dstWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    dstWs.Name = dstSheetName

    dstWs.Rows(1).Value = srcWs.Rows(1).Value
    dstRow = 2

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
' TXTを1パースしてUserForm1で節点選択 → 新規シートへ転記
' -------------------------------------------------------
Public Sub ParseAndSelectNodes()

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "支点反力テキストファイルを選択してください"
    fd.Filters.Clear
    fd.Filters.Add "テキストファイル", "*.txt"
    fd.AllowMultiSelect = False
    If fd.Show <> True Then Exit Sub
    Dim filePath As String
    filePath = fd.SelectedItems(1)

    Dim rows()    As ReactionRow
    Dim rowCount  As Long
    Dim nodeList() As String
    Dim nodeCount  As Long
    rowCount = 0
    nodeCount = 0

    Dim fileNum As Integer
    fileNum = FreeFile
    On Error GoTo FileOpenError2
    Open filePath For Input As #fileNum
    On Error GoTo 0

    Dim line As String
    Dim trimmedLine As String
    Dim parts() As String
    Dim loadNumber As Long
    Dim loadName As String
    Dim lineNum As Long
    Dim nodeStr As String
    Dim rx As Double, ry As Double, rz As Double
    Dim rmx As Double, rmy As Double, rmz As Double
    lineNum = 0
    loadNumber = 0
    loadName = ""

    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        lineNum = lineNum + 1
        trimmedLine = Trim(line)

        If Len(trimmedLine) = 0 Then GoTo Skip1
        If Left(trimmedLine, 5) = "=====" Then GoTo Skip1
        If InStr(line, "荷重番号") > 0 Then
            loadNumber = ExtractLoadNumber(line)
            loadName = ExtractLoadName(line)
            GoTo Skip1
        End If
        If InStr(trimmedLine, "節点番号") > 0 Then GoTo Skip1

        If Left(trimmedLine, 2) = "合計" Then
            parts = SplitNormalized(trimmedLine)
            If UBound(parts) < 6 Then
                Debug.Print "Warning: 合計行フィールド不足 (line " & lineNum & ")"
                GoTo Skip1
            End If
            If Not ValidateNumericParts(parts, 1, 6, lineNum) Then GoTo Skip1
            nodeStr = "合計"
        ElseIf IsNumeric(Left(trimmedLine, InStr(trimmedLine & " ", " ") - 1)) Then
            parts = SplitNormalized(trimmedLine)
            If UBound(parts) < 6 Then
                Debug.Print "Warning: 節点行フィールド不足 (line " & lineNum & ")"
                GoTo Skip1
            End If
            If Not IsNumeric(parts(0)) Then GoTo Skip1
            If Not ValidateNumericParts(parts, 1, 6, lineNum) Then GoTo Skip1
            nodeStr = parts(0)
        Else
            GoTo Skip1
        End If

        rx  = CDbl(parts(1)) : ry  = CDbl(parts(2)) : rz  = CDbl(parts(3))
        rmx = CDbl(parts(4)) : rmy = CDbl(parts(5)) : rmz = CDbl(parts(6))

        ReDim Preserve rows(rowCount)
        rows(rowCount).LoadNo   = loadNumber
        rows(rowCount).LoadName = loadName
        rows(rowCount).NodeNo   = nodeStr
        rows(rowCount).RX  = rx  : rows(rowCount).RY  = ry  : rows(rowCount).RZ  = rz
        rows(rowCount).RMX = rmx : rows(rowCount).RMY = rmy : rows(rowCount).RMZ = rmz
        rowCount = rowCount + 1

        Dim alreadyExists As Boolean
        alreadyExists = False
        If nodeCount > 0 Then alreadyExists = IsInList(nodeStr, nodeList)
        If Not alreadyExists Then
            ReDim Preserve nodeList(nodeCount)
            nodeList(nodeCount) = nodeStr
            nodeCount = nodeCount + 1
        End If
Skip1:
    Loop
    Close #fileNum

    If rowCount = 0 Then
        MsgBox "データ行が見つかりませんでした。", vbExclamation, "ParseAndSelectNodes"
        Exit Sub
    End If

    Dim selectedNodes() As String
    Dim selCount As Integer
    If Not ShowNodeSelectForm(nodeList, nodeCount, selectedNodes, selCount) Then
        Exit Sub
    End If
    If selCount = 0 Then
        MsgBox "節点番号が選択されていません。", vbExclamation, "ParseAndSelectNodes"
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "支点反力_" & Format(Now, "YYYYMMDD_HHMMSS")
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName

    Dim headers As Variant
    headers = Array("荷重番号", "荷重名称", "節点番号", "RX", "RY", "RZ", "RMX", "RMY", "RMZ")
    Dim c As Integer
    For c = 0 To 8
        ws.Cells(1, c + 1).Value = headers(c)
    Next c

    Dim rowIdx As Long
    rowIdx = 2
    Dim k As Long
    For k = 0 To rowCount - 1
        If IsInList(rows(k).NodeNo, selectedNodes) Then
            ws.Cells(rowIdx, 1).Value = rows(k).LoadNo
            ws.Cells(rowIdx, 2).Value = rows(k).LoadName
            If IsNumeric(rows(k).NodeNo) Then
                ws.Cells(rowIdx, 3).Value = CLng(rows(k).NodeNo)
            Else
                ws.Cells(rowIdx, 3).Value = rows(k).NodeNo
            End If
            ws.Cells(rowIdx, 4).Value = rows(k).RX
            ws.Cells(rowIdx, 5).Value = rows(k).RY
            ws.Cells(rowIdx, 6).Value = rows(k).RZ
            ws.Cells(rowIdx, 7).Value = rows(k).RMX
            ws.Cells(rowIdx, 8).Value = rows(k).RMY
            ws.Cells(rowIdx, 9).Value = rows(k).RMZ
            rowIdx = rowIdx + 1
        End If
    Next k

    MsgBox "完了しました。" & vbCrLf & _
           "出力シート名: " & sheetName & vbCrLf & _
           "転記行数: " & (rowIdx - 2) & " 行", vbInformation, "ParseAndSelectNodes"
    Exit Sub

FileOpenError2:
    MsgBox "ファイルを開けませんでした。" & vbCrLf & Err.Description, vbCritical, "ParseAndSelectNodes"

End Sub

' -------------------------------------------------------
' UserForm1を使って節点番号を選択させる
' 戻り値: True=OK, False=キャンセル
' -------------------------------------------------------
Private Function ShowNodeSelectForm( _
    ByRef nodeList() As String, _
    ByVal nodeCount As Long, _
    ByRef selectedNodes() As String, _
    ByRef selCount As Integer) As Boolean

    ShowNodeSelectForm = False
    selCount = 0

    Dim frm As Object
    On Error Resume Next
    Set frm = VBA.UserForms.Add("UserForm1")
    If Err.Number <> 0 Then
        MsgBox "UserForm1 が見つかりません。" & vbCrLf & _
               "先に Module1.CreateUserForm() を実行してください。", _
               vbCritical, "ParseAndSelectNodes"
        Exit Function
    End If
    On Error GoTo 0

    Dim i As Long
    For i = 0 To nodeCount - 1
        frm.Controls("lstNodes").AddItem nodeList(i)
    Next i

    frm.Show   ' モーダル：OK/Cancel までここでブロック

    If frm.Tag = "OK" Then
        For i = 0 To frm.Controls("lstNodes").ListCount - 1
            If frm.Controls("lstNodes").Selected(i) Then
                ReDim Preserve selectedNodes(selCount)
                selectedNodes(selCount) = frm.Controls("lstNodes").List(i)
                selCount = selCount + 1
            End If
        Next i
        ShowNodeSelectForm = True
    End If

    Unload frm

End Function

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
    Dim sub1 As String
    Dim posName As Long
    posName = InStr(line, "荷重名称")
    If posName > 0 Then
        sub1 = Mid(line, pos, posName - pos)
    Else
        sub1 = Mid(line, pos)
    End If
    Dim eqPos As Long
    eqPos = InStr(sub1, "=")
    If eqPos = 0 Then
        ExtractLoadNumber = 0
        Exit Function
    End If
    token = Trim(Mid(sub1, eqPos + 1))
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
