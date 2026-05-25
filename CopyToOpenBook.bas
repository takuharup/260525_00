Attribute VB_Name = "CopyToOpenBook"
Option Explicit

' -------------------------------------------------------
' ThisWorkbook内の指定シートを別の開いているブックへコピー
' -------------------------------------------------------
Public Sub CopyToOpenBook()

    ' --- コピー元シートの選択 ---
    Dim srcSheetName As String
    srcSheetName = InputBox("コピー元のシート名を入力してください" & vbCrLf & _
                            "（例：支点反力_20260525_143022）", "CopyToOpenBook")
    If StrPtr(srcSheetName) = 0 Then Exit Sub
    If Trim(srcSheetName) = "" Then Exit Sub

    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = ThisWorkbook.Worksheets(srcSheetName)
    On Error GoTo 0
    If srcWs Is Nothing Then
        MsgBox "シート「" & srcSheetName & "」が見つかりません。", vbCritical, "CopyToOpenBook"
        Exit Sub
    End If

    ' --- コピー先ブックの選択 ---
    Dim bookList As String
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            bookList = bookList & wb.Name & vbCrLf
        End If
    Next wb

    If Len(bookList) = 0 Then
        MsgBox "コピー先となる別のブックが開かれていません。" & vbCrLf & _
               "コピー先のブックを開いてから再実行してください。", vbExclamation, "CopyToOpenBook"
        Exit Sub
    End If

    Dim dstBookName As String
    dstBookName = InputBox("コピー先のブック名を入力してください：" & vbCrLf & bookList, "CopyToOpenBook")
    If StrPtr(dstBookName) = 0 Then Exit Sub
    If Trim(dstBookName) = "" Then Exit Sub

    Dim dstWb As Workbook
    On Error Resume Next
    Set dstWb = Workbooks(dstBookName)
    On Error GoTo 0
    If dstWb Is Nothing Then
        MsgBox "ブック「" & dstBookName & "」が見つかりません。", vbCritical, "CopyToOpenBook"
        Exit Sub
    End If

    ' --- シートをコピー ---
    srcWs.Copy After:=dstWb.Sheets(dstWb.Sheets.Count)

    MsgBox "「" & srcSheetName & "」を「" & dstBookName & "」にコピーしました。", _
           vbInformation, "CopyToOpenBook"

End Sub
