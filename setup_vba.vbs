Option Explicit

'=================================================================
' setup_vba.vbs
' SupportReaction.xlsm を新規作成し、
' ParseSupportReaction モジュールと FormNodeSelect フォームを
' 自動セットアップする。
'
' 前提条件：
'   Excel「ファイル」→「オプション」→「トラストセンター」→
'   「トラストセンターの設定」→「マクロの設定」で
'   「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」
'   にチェックを入れてから実行してください。
'=================================================================

Dim fso, xl, wb, savePath

Set fso = CreateObject("Scripting.FileSystemObject")
savePath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "SupportReaction.xlsm")

' ---- Excel 起動 ----
Set xl = CreateObject("Excel.Application")
xl.Visible       = False
xl.DisplayAlerts = False

' ---- 新規ブック作成 ----
Set wb = xl.Workbooks.Add()

' ---- .xlsm として保存（52 = xlOpenXMLWorkbookMacroEnabled）----
On Error Resume Next
wb.SaveAs savePath, 52
If Err.Number <> 0 Then
    MsgBox "保存に失敗しました。" & vbCrLf & Err.Description, 16, "エラー"
    wb.Close False : xl.Quit : WScript.Quit
End If
On Error GoTo 0

' ---- VBProject アクセス確認 ----
On Error Resume Next
Dim chk : chk = wb.VBProject.VBComponents.Count
If Err.Number <> 0 Then
    MsgBox "VBAプロジェクトにアクセスできません。" & vbCrLf & vbCrLf & _
           "Excel の「ファイル」→「オプション」→「トラストセンター」→" & vbCrLf & _
           "「トラストセンターの設定」→「マクロの設定」で" & vbCrLf & _
           "「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」" & vbCrLf & _
           "を有効にしてから再実行してください。", 16, "エラー"
    wb.Close False : xl.Quit : WScript.Quit
End If
On Error GoTo 0

' ---- 標準モジュール ParseSupportReaction 追加 ----
On Error Resume Next
wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents("ParseSupportReaction")
On Error GoTo 0

Dim oMod
Set oMod = wb.VBProject.VBComponents.Add(1)   ' vbext_ct_StdModule
oMod.Name = "ParseSupportReaction"
With oMod.CodeModule
    If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
    .AddFromString BuildModuleCode()
End With

' ---- UserForm FormNodeSelect 追加 ----
On Error Resume Next
wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents("FormNodeSelect")
On Error GoTo 0

Dim oForm, oCtrl
Set oForm = wb.VBProject.VBComponents.Add(3)  ' vbext_ct_MSForm
oForm.Name = "FormNodeSelect"
oForm.Properties("Caption")         = "節点番号選択"
oForm.Properties("Width")           = 282
oForm.Properties("Height")          = 372
oForm.Properties("StartUpPosition") = 1       ' 1 = CenterOwner

' ラベル
Set oCtrl = oForm.Designer.Controls.Add("Forms.Label.1")
oCtrl.Name    = "lblInstruction"
oCtrl.Caption = "転記する節点番号を選択してください（複数選択可）"
oCtrl.Left = 6 : oCtrl.Top = 6 : oCtrl.Width = 264 : oCtrl.Height = 18

' リストボックス（MultiSelect = 1 : fmMultiSelectMulti）
Set oCtrl = oForm.Designer.Controls.Add("Forms.ListBox.1")
oCtrl.Name        = "lstNodes"
oCtrl.Left        = 6 : oCtrl.Top = 30 : oCtrl.Width = 264 : oCtrl.Height = 246
oCtrl.MultiSelect = 1

' OK ボタン
Set oCtrl = oForm.Designer.Controls.Add("Forms.CommandButton.1")
oCtrl.Name    = "btnOK"
oCtrl.Caption = "OK"
oCtrl.Left    = 60 : oCtrl.Top = 288 : oCtrl.Width = 72 : oCtrl.Height = 24

' キャンセルボタン
Set oCtrl = oForm.Designer.Controls.Add("Forms.CommandButton.1")
oCtrl.Name    = "btnCancel"
oCtrl.Caption = "キャンセル"
oCtrl.Left    = 156 : oCtrl.Top = 288 : oCtrl.Width = 90 : oCtrl.Height = 24

' フォームイベントコード
With oForm.CodeModule
    If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
    .AddFromString BuildFormCode()
End With

' ---- 保存して完了 ----
wb.Save
xl.Visible = True

MsgBox "セットアップ完了しました。" & vbCrLf & vbCrLf & _
       savePath & vbCrLf & vbCrLf & _
       "Alt+F8 から以下のマクロを実行できます：" & vbCrLf & _
       "  ParseSupportReaction  … txt を全行変換してシートへ出力" & vbCrLf & _
       "  FilterByNode          … 既存テーブルから節点番号で絞り込み" & vbCrLf & _
       "  ParseAndSelectNodes   … txt 読込 → フォームで節点選択 → 転記", _
       64, "SupportReaction セットアップ"

'=================================================================
' ヘルパー：改行付き行追加
'=================================================================
Function L(s)
    L = s & vbCrLf
End Function

'=================================================================
' VBA モジュールコード本体
' コード中の | は最後に Chr(34) へ置換（ダブルクォート代替）
'=================================================================
Function BuildModuleCode()
    Dim c : c = ""

    '--- Option Explicit & Type 宣言 ---
    c = c & L("Option Explicit")
    c = c & L("")
    c = c & L("Private Type ReactionRow")
    c = c & L("    LoadNo   As Long")
    c = c & L("    LoadName As String")
    c = c & L("    NodeNo   As String")
    c = c & L("    RX       As Double")
    c = c & L("    RY       As Double")
    c = c & L("    RZ       As Double")
    c = c & L("    RMX      As Double")
    c = c & L("    RMY      As Double")
    c = c & L("    RMZ      As Double")
    c = c & L("End Type")

    '--- ParseSupportReaction ---
    c = c & L("")
    c = c & L("Public Sub ParseSupportReaction()")
    c = c & L("")
    c = c & L("    Dim filePath As String")
    c = c & L("    Dim fileNum As Integer")
    c = c & L("    Dim ws As Worksheet")
    c = c & L("    Dim sheetName As String")
    c = c & L("    Dim line As String")
    c = c & L("    Dim trimmedLine As String")
    c = c & L("    Dim parts() As String")
    c = c & L("    Dim loadNumber As Long")
    c = c & L("    Dim loadName As String")
    c = c & L("    Dim rowIdx As Long")
    c = c & L("    Dim lineNum As Long")
    c = c & L("    Dim nodeStr As String")
    c = c & L("    Dim rx As Double, ry As Double, rz As Double")
    c = c & L("    Dim rmx As Double, rmy As Double, rmz As Double")
    c = c & L("")
    c = c & L("    Dim fd As FileDialog")
    c = c & L("    Set fd = Application.FileDialog(msoFileDialogFilePicker)")
    c = c & L("    fd.Title = |支点反力テキストファイルを選択してください|")
    c = c & L("    fd.Filters.Clear")
    c = c & L("    fd.Filters.Add |テキストファイル|, |*.txt|")
    c = c & L("    fd.AllowMultiSelect = False")
    c = c & L("    If fd.Show <> True Then Exit Sub")
    c = c & L("    filePath = fd.SelectedItems(1)")
    c = c & L("")
    c = c & L("    sheetName = |支点反力_| & Format(Now, |YYYYMMDD_HHMMSS|)")
    c = c & L("    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))")
    c = c & L("    ws.Name = sheetName")
    c = c & L("")
    c = c & L("    ws.Cells(1, 1).Value = |荷重番号|")
    c = c & L("    ws.Cells(1, 2).Value = |荷重名称|")
    c = c & L("    ws.Cells(1, 3).Value = |節点番号|")
    c = c & L("    ws.Cells(1, 4).Value = |RX|")
    c = c & L("    ws.Cells(1, 5).Value = |RY|")
    c = c & L("    ws.Cells(1, 6).Value = |RZ|")
    c = c & L("    ws.Cells(1, 7).Value = |RMX|")
    c = c & L("    ws.Cells(1, 8).Value = |RMY|")
    c = c & L("    ws.Cells(1, 9).Value = |RMZ|")
    c = c & L("")
    c = c & L("    rowIdx = 2")
    c = c & L("    lineNum = 0")
    c = c & L("    loadNumber = 0")
    c = c & L("    loadName = ||")
    c = c & L("")
    c = c & L("    fileNum = FreeFile")
    c = c & L("    On Error GoTo FileOpenError")
    c = c & L("    Open filePath For Input As #fileNum")
    c = c & L("    On Error GoTo 0")
    c = c & L("")
    c = c & L("    Do While Not EOF(fileNum)")
    c = c & L("        Line Input #fileNum, line")
    c = c & L("        lineNum = lineNum + 1")
    c = c & L("        trimmedLine = Trim(line)")
    c = c & L("        If Len(trimmedLine) = 0 Then GoTo NextLine")
    c = c & L("        If Left(trimmedLine, 5) = |=====| Then GoTo NextLine")
    c = c & L("        If InStr(line, |荷重番号|) > 0 Then")
    c = c & L("            loadNumber = ExtractLoadNumber(line)")
    c = c & L("            loadName = ExtractLoadName(line)")
    c = c & L("            GoTo NextLine")
    c = c & L("        End If")
    c = c & L("        If InStr(trimmedLine, |節点番号|) > 0 Then GoTo NextLine")
    c = c & L("        If Left(trimmedLine, 2) = |合計| Then")
    c = c & L("            parts = SplitNormalized(trimmedLine)")
    c = c & L("            If UBound(parts) >= 6 Then")
    c = c & L("                If Not ValidateNumericParts(parts, 1, 6, lineNum) Then GoTo NextLine")
    c = c & L("                rx  = CDbl(parts(1))")
    c = c & L("                ry  = CDbl(parts(2))")
    c = c & L("                rz  = CDbl(parts(3))")
    c = c & L("                rmx = CDbl(parts(4))")
    c = c & L("                rmy = CDbl(parts(5))")
    c = c & L("                rmz = CDbl(parts(6))")
    c = c & L("                WriteRow ws, rowIdx, loadNumber, loadName, |合計|, rx, ry, rz, rmx, rmy, rmz")
    c = c & L("                rowIdx = rowIdx + 1")
    c = c & L("            Else")
    c = c & L("                Debug.Print |Warning: 合計行のフィールド数不足 (line | & lineNum & |): | & line")
    c = c & L("            End If")
    c = c & L("            GoTo NextLine")
    c = c & L("        End If")
    c = c & L("        If IsNumeric(Left(trimmedLine, InStr(trimmedLine & | |, | |) - 1)) Then")
    c = c & L("            parts = SplitNormalized(trimmedLine)")
    c = c & L("            If UBound(parts) >= 6 Then")
    c = c & L("                If Not IsNumeric(parts(0)) Then GoTo NextLine")
    c = c & L("                If Not ValidateNumericParts(parts, 1, 6, lineNum) Then GoTo NextLine")
    c = c & L("                nodeStr = parts(0)")
    c = c & L("                rx  = CDbl(parts(1))")
    c = c & L("                ry  = CDbl(parts(2))")
    c = c & L("                rz  = CDbl(parts(3))")
    c = c & L("                rmx = CDbl(parts(4))")
    c = c & L("                rmy = CDbl(parts(5))")
    c = c & L("                rmz = CDbl(parts(6))")
    c = c & L("                WriteRow ws, rowIdx, loadNumber, loadName, nodeStr, rx, ry, rz, rmx, rmy, rmz")
    c = c & L("                rowIdx = rowIdx + 1")
    c = c & L("            Else")
    c = c & L("                Debug.Print |Warning: 節点行のフィールド数不足 (line | & lineNum & |): | & line")
    c = c & L("            End If")
    c = c & L("        End If")
    c = c & L("NextLine:")
    c = c & L("    Loop")
    c = c & L("    Close #fileNum")
    c = c & L("")
    c = c & L("    Dim dataRows As Long")
    c = c & L("    dataRows = rowIdx - 2")
    c = c & L("    MsgBox |完了しました。| & vbCrLf & _")
    c = c & L("           |出力シート名: | & sheetName & vbCrLf & _")
    c = c & L("           |データ行数: | & dataRows & | 行|, _")
    c = c & L("           vbInformation, |ParseSupportReaction|")
    c = c & L("    Exit Sub")
    c = c & L("")
    c = c & L("FileOpenError:")
    c = c & L("    MsgBox |ファイルを開けませんでした。| & vbCrLf & Err.Description, vbCritical, |ParseSupportReaction|")
    c = c & L("")
    c = c & L("End Sub")

    '--- SplitNormalized ---
    c = c & L("")
    c = c & L("Private Function SplitNormalized(ByVal s As String) As String()")
    c = c & L("    Do While InStr(s, |  |) > 0")
    c = c & L("        s = Replace(s, |  |, | |)")
    c = c & L("    Loop")
    c = c & L("    SplitNormalized = Split(Trim(s), | |)")
    c = c & L("End Function")

    '--- ExtractLoadNumber ---
    c = c & L("")
    c = c & L("Private Function ExtractLoadNumber(ByVal line As String) As Long")
    c = c & L("    Dim pos As Long")
    c = c & L("    Dim token As String")
    c = c & L("    pos = InStr(line, |荷重番号|)")
    c = c & L("    If pos = 0 Then")
    c = c & L("        ExtractLoadNumber = 0")
    c = c & L("        Exit Function")
    c = c & L("    End If")
    c = c & L("    Dim sub1 As String")
    c = c & L("    Dim posName As Long")
    c = c & L("    posName = InStr(line, |荷重名称|)")
    c = c & L("    If posName > 0 Then")
    c = c & L("        sub1 = Mid(line, pos, posName - pos)")
    c = c & L("    Else")
    c = c & L("        sub1 = Mid(line, pos)")
    c = c & L("    End If")
    c = c & L("    Dim eqPos As Long")
    c = c & L("    eqPos = InStr(sub1, |=|)")
    c = c & L("    If eqPos = 0 Then")
    c = c & L("        ExtractLoadNumber = 0")
    c = c & L("        Exit Function")
    c = c & L("    End If")
    c = c & L("    token = Trim(Mid(sub1, eqPos + 1))")
    c = c & L("    token = SplitNormalized(token)(0)")
    c = c & L("    If IsNumeric(token) Then")
    c = c & L("        ExtractLoadNumber = CLng(token)")
    c = c & L("    Else")
    c = c & L("        ExtractLoadNumber = 0")
    c = c & L("    End If")
    c = c & L("End Function")

    '--- ExtractLoadName ---
    c = c & L("")
    c = c & L("Private Function ExtractLoadName(ByVal line As String) As String")
    c = c & L("    Dim pos As Long")
    c = c & L("    pos = InStr(line, |荷重名称|)")
    c = c & L("    If pos = 0 Then")
    c = c & L("        ExtractLoadName = ||")
    c = c & L("        Exit Function")
    c = c & L("    End If")
    c = c & L("    Dim sub1 As String")
    c = c & L("    sub1 = Mid(line, pos)")
    c = c & L("    Dim eqPos As Long")
    c = c & L("    eqPos = InStr(sub1, |=|)")
    c = c & L("    If eqPos = 0 Then")
    c = c & L("        ExtractLoadName = ||")
    c = c & L("        Exit Function")
    c = c & L("    End If")
    c = c & L("    ExtractLoadName = Trim(Mid(sub1, eqPos + 1))")
    c = c & L("End Function")

    '--- ValidateNumericParts ---
    c = c & L("")
    c = c & L("Private Function ValidateNumericParts(ByRef parts() As String, _")
    c = c & L("                                       ByVal fromIdx As Integer, _")
    c = c & L("                                       ByVal toIdx As Integer, _")
    c = c & L("                                       ByVal lineNum As Long) As Boolean")
    c = c & L("    Dim i As Integer")
    c = c & L("    For i = fromIdx To toIdx")
    c = c & L("        If Not IsNumeric(parts(i)) Then")
    c = c & L("            Debug.Print |Warning: 数値変換失敗 parts(| & i & |)='| & parts(i) & |' (line | & lineNum & |)|")
    c = c & L("            ValidateNumericParts = False")
    c = c & L("            Exit Function")
    c = c & L("        End If")
    c = c & L("    Next i")
    c = c & L("    ValidateNumericParts = True")
    c = c & L("End Function")

    '--- FilterByNode ---
    c = c & L("")
    c = c & L("Public Sub FilterByNode()")
    c = c & L("")
    c = c & L("    Dim srcSheetName As String")
    c = c & L("    Dim srcWs As Worksheet")
    c = c & L("    Dim dstWs As Worksheet")
    c = c & L("    Dim dstSheetName As String")
    c = c & L("    Dim nodeInput As String")
    c = c & L("    Dim nodeTokens() As String")
    c = c & L("    Dim i As Integer")
    c = c & L("    Dim lastRow As Long")
    c = c & L("    Dim dstRow As Long")
    c = c & L("    Dim cellVal As String")
    c = c & L("")
    c = c & L("    srcSheetName = InputBox(|抽出元のシート名を入力してください| & vbCrLf & _")
    c = c & L("                            |（例：支点反力_20260525_143022）|, |FilterByNode|)")
    c = c & L("    If StrPtr(srcSheetName) = 0 Then Exit Sub")
    c = c & L("    If Trim(srcSheetName) = || Then Exit Sub")
    c = c & L("")
    c = c & L("    On Error Resume Next")
    c = c & L("    Set srcWs = ThisWorkbook.Worksheets(srcSheetName)")
    c = c & L("    On Error GoTo 0")
    c = c & L("    If srcWs Is Nothing Then")
    c = c & L("        MsgBox |シート「| & srcSheetName & |」が見つかりません。|, vbCritical, |FilterByNode|")
    c = c & L("        Exit Sub")
    c = c & L("    End If")
    c = c & L("")
    c = c & L("    nodeInput = InputBox(|抽出する節点番号をカンマ区切りで入力してください| & vbCrLf & _")
    c = c & L("                         |（例： 1,3,合計　または　合計）|, |FilterByNode|)")
    c = c & L("    If StrPtr(nodeInput) = 0 Then Exit Sub")
    c = c & L("    If Trim(nodeInput) = || Then")
    c = c & L("        MsgBox |節点番号が入力されていません。|, vbExclamation, |FilterByNode|")
    c = c & L("        Exit Sub")
    c = c & L("    End If")
    c = c & L("")
    c = c & L("    nodeTokens = Split(nodeInput, |,|)")
    c = c & L("    For i = 0 To UBound(nodeTokens)")
    c = c & L("        nodeTokens(i) = Trim(nodeTokens(i))")
    c = c & L("    Next i")
    c = c & L("")
    c = c & L("    dstSheetName = |抽出_| & Format(Now, |YYYYMMDD_HHMMSS|)")
    c = c & L("    Set dstWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))")
    c = c & L("    dstWs.Name = dstSheetName")
    c = c & L("    dstWs.Rows(1).Value = srcWs.Rows(1).Value")
    c = c & L("    dstRow = 2")
    c = c & L("")
    c = c & L("    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row")
    c = c & L("    Dim extractCount As Long")
    c = c & L("    extractCount = 0")
    c = c & L("")
    c = c & L("    For i = 2 To lastRow")
    c = c & L("        cellVal = Trim(CStr(srcWs.Cells(i, 3).Value))")
    c = c & L("        If IsInList(cellVal, nodeTokens) Then")
    c = c & L("            dstWs.Rows(dstRow).Value = srcWs.Rows(i).Value")
    c = c & L("            dstRow = dstRow + 1")
    c = c & L("            extractCount = extractCount + 1")
    c = c & L("        End If")
    c = c & L("    Next i")
    c = c & L("")
    c = c & L("    If extractCount = 0 Then")
    c = c & L("        MsgBox |該当する節点番号の行がありませんでした。| & vbCrLf & _")
    c = c & L("               |出力シート名: | & dstSheetName, vbExclamation, |FilterByNode|")
    c = c & L("    Else")
    c = c & L("        MsgBox |完了しました。| & vbCrLf & _")
    c = c & L("               |出力シート名: | & dstSheetName & vbCrLf & _")
    c = c & L("               |抽出行数: | & extractCount & | 行|, vbInformation, |FilterByNode|")
    c = c & L("    End If")
    c = c & L("")
    c = c & L("End Sub")

    '--- IsInList ---
    c = c & L("")
    c = c & L("Private Function IsInList(ByVal target As String, ByRef list() As String) As Boolean")
    c = c & L("    Dim i As Integer")
    c = c & L("    For i = 0 To UBound(list)")
    c = c & L("        If Trim(list(i)) = Trim(target) Then")
    c = c & L("            IsInList = True")
    c = c & L("            Exit Function")
    c = c & L("        End If")
    c = c & L("    Next i")
    c = c & L("    IsInList = False")
    c = c & L("End Function")

    '--- ParseAndSelectNodes ---
    c = c & L("")
    c = c & L("Public Sub ParseAndSelectNodes()")
    c = c & L("")
    c = c & L("    Dim fd As FileDialog")
    c = c & L("    Set fd = Application.FileDialog(msoFileDialogFilePicker)")
    c = c & L("    fd.Title = |支点反力テキストファイルを選択してください|")
    c = c & L("    fd.Filters.Clear")
    c = c & L("    fd.Filters.Add |テキストファイル|, |*.txt|")
    c = c & L("    fd.AllowMultiSelect = False")
    c = c & L("    If fd.Show <> True Then Exit Sub")
    c = c & L("    Dim filePath As String")
    c = c & L("    filePath = fd.SelectedItems(1)")
    c = c & L("")
    c = c & L("    Dim rows()    As ReactionRow")
    c = c & L("    Dim rowCount  As Long")
    c = c & L("    Dim nodeList() As String")
    c = c & L("    Dim nodeCount  As Long")
    c = c & L("    rowCount = 0")
    c = c & L("    nodeCount = 0")
    c = c & L("")
    c = c & L("    Dim fileNum As Integer")
    c = c & L("    fileNum = FreeFile")
    c = c & L("    On Error GoTo FileOpenError2")
    c = c & L("    Open filePath For Input As #fileNum")
    c = c & L("    On Error GoTo 0")
    c = c & L("")
    c = c & L("    Dim line As String")
    c = c & L("    Dim trimmedLine As String")
    c = c & L("    Dim parts() As String")
    c = c & L("    Dim loadNumber As Long")
    c = c & L("    Dim loadName As String")
    c = c & L("    Dim lineNum As Long")
    c = c & L("    Dim nodeStr As String")
    c = c & L("    Dim rx As Double, ry As Double, rz As Double")
    c = c & L("    Dim rmx As Double, rmy As Double, rmz As Double")
    c = c & L("    lineNum = 0 : loadNumber = 0 : loadName = ||")
    c = c & L("")
    c = c & L("    Do While Not EOF(fileNum)")
    c = c & L("        Line Input #fileNum, line")
    c = c & L("        lineNum = lineNum + 1")
    c = c & L("        trimmedLine = Trim(line)")
    c = c & L("        If Len(trimmedLine) = 0 Then GoTo Skip1")
    c = c & L("        If Left(trimmedLine, 5) = |=====| Then GoTo Skip1")
    c = c & L("        If InStr(line, |荷重番号|) > 0 Then")
    c = c & L("            loadNumber = ExtractLoadNumber(line)")
    c = c & L("            loadName = ExtractLoadName(line)")
    c = c & L("            GoTo Skip1")
    c = c & L("        End If")
    c = c & L("        If InStr(trimmedLine, |節点番号|) > 0 Then GoTo Skip1")
    c = c & L("        If Left(trimmedLine, 2) = |合計| Then")
    c = c & L("            parts = SplitNormalized(trimmedLine)")
    c = c & L("            If UBound(parts) < 6 Then")
    c = c & L("                Debug.Print |Warning: 合計行フィールド不足 (line | & lineNum & |)|")
    c = c & L("                GoTo Skip1")
    c = c & L("            End If")
    c = c & L("            If Not ValidateNumericParts(parts, 1, 6, lineNum) Then GoTo Skip1")
    c = c & L("            nodeStr = |合計|")
    c = c & L("        ElseIf IsNumeric(Left(trimmedLine, InStr(trimmedLine & | |, | |) - 1)) Then")
    c = c & L("            parts = SplitNormalized(trimmedLine)")
    c = c & L("            If UBound(parts) < 6 Then")
    c = c & L("                Debug.Print |Warning: 節点行フィールド不足 (line | & lineNum & |)|")
    c = c & L("                GoTo Skip1")
    c = c & L("            End If")
    c = c & L("            If Not IsNumeric(parts(0)) Then GoTo Skip1")
    c = c & L("            If Not ValidateNumericParts(parts, 1, 6, lineNum) Then GoTo Skip1")
    c = c & L("            nodeStr = parts(0)")
    c = c & L("        Else")
    c = c & L("            GoTo Skip1")
    c = c & L("        End If")
    c = c & L("        rx  = CDbl(parts(1)) : ry  = CDbl(parts(2)) : rz  = CDbl(parts(3))")
    c = c & L("        rmx = CDbl(parts(4)) : rmy = CDbl(parts(5)) : rmz = CDbl(parts(6))")
    c = c & L("        ReDim Preserve rows(rowCount)")
    c = c & L("        rows(rowCount).LoadNo   = loadNumber")
    c = c & L("        rows(rowCount).LoadName = loadName")
    c = c & L("        rows(rowCount).NodeNo   = nodeStr")
    c = c & L("        rows(rowCount).RX  = rx  : rows(rowCount).RY  = ry  : rows(rowCount).RZ  = rz")
    c = c & L("        rows(rowCount).RMX = rmx : rows(rowCount).RMY = rmy : rows(rowCount).RMZ = rmz")
    c = c & L("        rowCount = rowCount + 1")
    c = c & L("        Dim alreadyExists As Boolean")
    c = c & L("        alreadyExists = False")
    c = c & L("        If nodeCount > 0 Then alreadyExists = IsInList(nodeStr, nodeList)")
    c = c & L("        If Not alreadyExists Then")
    c = c & L("            ReDim Preserve nodeList(nodeCount)")
    c = c & L("            nodeList(nodeCount) = nodeStr")
    c = c & L("            nodeCount = nodeCount + 1")
    c = c & L("        End If")
    c = c & L("Skip1:")
    c = c & L("    Loop")
    c = c & L("    Close #fileNum")
    c = c & L("")
    c = c & L("    If rowCount = 0 Then")
    c = c & L("        MsgBox |データ行が見つかりませんでした。|, vbExclamation, |ParseAndSelectNodes|")
    c = c & L("        Exit Sub")
    c = c & L("    End If")
    c = c & L("")
    c = c & L("    Dim frm As FormNodeSelect")
    c = c & L("    Set frm = New FormNodeSelect")
    c = c & L("    Dim j As Integer")
    c = c & L("    For j = 0 To nodeCount - 1")
    c = c & L("        frm.lstNodes.AddItem nodeList(j)")
    c = c & L("    Next j")
    c = c & L("    frm.Show")
    c = c & L("")
    c = c & L("    If frm.Tag <> |OK| Then")
    c = c & L("        Unload frm")
    c = c & L("        Exit Sub")
    c = c & L("    End If")
    c = c & L("")
    c = c & L("    Dim selectedNodes() As String")
    c = c & L("    Dim selCount As Integer")
    c = c & L("    selCount = 0")
    c = c & L("    For j = 0 To frm.lstNodes.ListCount - 1")
    c = c & L("        If frm.lstNodes.Selected(j) Then")
    c = c & L("            ReDim Preserve selectedNodes(selCount)")
    c = c & L("            selectedNodes(selCount) = frm.lstNodes.List(j)")
    c = c & L("            selCount = selCount + 1")
    c = c & L("        End If")
    c = c & L("    Next j")
    c = c & L("    Unload frm")
    c = c & L("")
    c = c & L("    If selCount = 0 Then")
    c = c & L("        MsgBox |節点番号が選択されていません。|, vbExclamation, |ParseAndSelectNodes|")
    c = c & L("        Exit Sub")
    c = c & L("    End If")
    c = c & L("")
    c = c & L("    Dim ws As Worksheet")
    c = c & L("    Dim sheetName As String")
    c = c & L("    sheetName = |支点反力_| & Format(Now, |YYYYMMDD_HHMMSS|)")
    c = c & L("    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))")
    c = c & L("    ws.Name = sheetName")
    c = c & L("")
    c = c & L("    Dim headers As Variant")
    c = c & L("    headers = Array(|荷重番号|, |荷重名称|, |節点番号|, |RX|, |RY|, |RZ|, |RMX|, |RMY|, |RMZ|)")
    c = c & L("    Dim ci As Integer")
    c = c & L("    For ci = 0 To 8")
    c = c & L("        ws.Cells(1, ci + 1).Value = headers(ci)")
    c = c & L("    Next ci")
    c = c & L("")
    c = c & L("    Dim rowIdx As Long")
    c = c & L("    rowIdx = 2")
    c = c & L("    Dim k As Long")
    c = c & L("    For k = 0 To rowCount - 1")
    c = c & L("        If IsInList(rows(k).NodeNo, selectedNodes) Then")
    c = c & L("            ws.Cells(rowIdx, 1).Value = rows(k).LoadNo")
    c = c & L("            ws.Cells(rowIdx, 2).Value = rows(k).LoadName")
    c = c & L("            If IsNumeric(rows(k).NodeNo) Then")
    c = c & L("                ws.Cells(rowIdx, 3).Value = CLng(rows(k).NodeNo)")
    c = c & L("            Else")
    c = c & L("                ws.Cells(rowIdx, 3).Value = rows(k).NodeNo")
    c = c & L("            End If")
    c = c & L("            ws.Cells(rowIdx, 4).Value = rows(k).RX")
    c = c & L("            ws.Cells(rowIdx, 5).Value = rows(k).RY")
    c = c & L("            ws.Cells(rowIdx, 6).Value = rows(k).RZ")
    c = c & L("            ws.Cells(rowIdx, 7).Value = rows(k).RMX")
    c = c & L("            ws.Cells(rowIdx, 8).Value = rows(k).RMY")
    c = c & L("            ws.Cells(rowIdx, 9).Value = rows(k).RMZ")
    c = c & L("            rowIdx = rowIdx + 1")
    c = c & L("        End If")
    c = c & L("    Next k")
    c = c & L("")
    c = c & L("    MsgBox |完了しました。| & vbCrLf & _")
    c = c & L("           |出力シート名: | & sheetName & vbCrLf & _")
    c = c & L("           |転記行数: | & (rowIdx - 2) & | 行|, vbInformation, |ParseAndSelectNodes|")
    c = c & L("    Exit Sub")
    c = c & L("")
    c = c & L("FileOpenError2:")
    c = c & L("    MsgBox |ファイルを開けませんでした。| & vbCrLf & Err.Description, vbCritical, |ParseAndSelectNodes|")
    c = c & L("")
    c = c & L("End Sub")

    '--- WriteRow ---
    c = c & L("")
    c = c & L("Private Sub WriteRow(ByVal ws As Worksheet, ByVal rowIdx As Long, _")
    c = c & L("                     ByVal loadNum As Long, ByVal loadNm As String, _")
    c = c & L("                     ByVal nodeVal As String, _")
    c = c & L("                     ByVal rx As Double, ByVal ry As Double, ByVal rz As Double, _")
    c = c & L("                     ByVal rmx As Double, ByVal rmy As Double, ByVal rmz As Double)")
    c = c & L("    ws.Cells(rowIdx, 1).Value = loadNum")
    c = c & L("    ws.Cells(rowIdx, 2).Value = loadNm")
    c = c & L("    If IsNumeric(nodeVal) Then")
    c = c & L("        ws.Cells(rowIdx, 3).Value = CLng(nodeVal)")
    c = c & L("    Else")
    c = c & L("        ws.Cells(rowIdx, 3).Value = nodeVal")
    c = c & L("    End If")
    c = c & L("    ws.Cells(rowIdx, 4).Value = rx")
    c = c & L("    ws.Cells(rowIdx, 5).Value = ry")
    c = c & L("    ws.Cells(rowIdx, 6).Value = rz")
    c = c & L("    ws.Cells(rowIdx, 7).Value = rmx")
    c = c & L("    ws.Cells(rowIdx, 8).Value = rmy")
    c = c & L("    ws.Cells(rowIdx, 9).Value = rmz")
    c = c & L("End Sub")

    ' | → Chr(34) に置換してダブルクォートを復元
    BuildModuleCode = Replace(c, "|", Chr(34))
End Function

'=================================================================
' FormNodeSelect イベントコード
'=================================================================
Function BuildFormCode()
    Dim c : c = ""
    c = c & L("Option Explicit")
    c = c & L("")
    c = c & L("Private Sub btnOK_Click()")
    c = c & L("    Me.Tag = " & Chr(34) & "OK" & Chr(34))
    c = c & L("    Me.Hide")
    c = c & L("End Sub")
    c = c & L("")
    c = c & L("Private Sub btnCancel_Click()")
    c = c & L("    Me.Tag = " & Chr(34) & "Cancel" & Chr(34))
    c = c & L("    Me.Hide")
    c = c & L("End Sub")
    c = c & L("")
    c = c & L("Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)")
    c = c & L("    If CloseMode = vbFormControlMenu Then")
    c = c & L("        Me.Tag = " & Chr(34) & "Cancel" & Chr(34))
    c = c & L("        Me.Hide")
    c = c & L("        Cancel = True")
    c = c & L("    End If")
    c = c & L("End Sub")
    BuildFormCode = c
End Function
