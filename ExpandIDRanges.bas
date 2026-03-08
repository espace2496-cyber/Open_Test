Attribute VB_Name = "ExpandIDRanges"
Option Explicit

' ============================================================
' ID範囲展開マクロ
' "A21PF51 TO A21PF61" のような範囲指定を個別IDに展開する
' ============================================================

Sub ExpandIDRanges()
    Dim inputRange As Range
    Dim cell As Range
    Dim outputSheet As Worksheet
    Dim outputRow As Long
    Dim cellValue As String

    ' 選択範囲を取得
    On Error Resume Next
    Set inputRange = Application.InputBox( _
        prompt:="展開したいID範囲が入っているセルを選択してください。" & vbCrLf & _
               "例: A21PF51 TO A21PF61", _
        Title:="ID範囲展開", _
        Type:=8)
    On Error GoTo 0

    If inputRange Is Nothing Then Exit Sub

    ' 出力先シートを作成
    Dim sheetName As String
    sheetName = "展開結果"

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set outputSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    outputSheet.Name = sheetName

    ' ヘッダー
    outputSheet.Cells(1, 1).Value = "元のID範囲"
    outputSheet.Cells(1, 2).Value = "展開後ID"
    outputSheet.Range("A1:B1").Font.Bold = True
    outputRow = 2

    ' 各セルを処理
    For Each cell In inputRange
        cellValue = Trim(cell.Value)
        If Len(cellValue) = 0 Then GoTo NextCell

        If InStr(UCase(cellValue), " TO ") > 0 Then
            outputRow = ExpandOneRange(cellValue, outputSheet, outputRow)
        Else
            ' TO がない場合はそのまま出力
            outputSheet.Cells(outputRow, 1).Value = cellValue
            outputSheet.Cells(outputRow, 2).Value = cellValue
            outputRow = outputRow + 1
        End If
NextCell:
    Next cell

    ' 列幅自動調整
    outputSheet.Columns("A:B").AutoFit
    outputSheet.Activate

    MsgBox "展開完了しました。" & vbCrLf & _
           "展開結果: " & (outputRow - 2) & " 件", vbInformation, "完了"
End Sub

' 1つの範囲文字列を展開して出力シートに書き込む
' 戻り値: 次の出力行番号
Private Function ExpandOneRange(ByVal rangeStr As String, _
                                 ByRef ws As Worksheet, _
                                 ByVal startRow As Long) As Long
    Dim parts() As String
    Dim idStart As String
    Dim idEnd As String
    Dim prefix As String
    Dim numStart As Long
    Dim numEnd As Long
    Dim numWidth As Long
    Dim i As Long
    Dim row As Long

    ' "TO" で分割
    parts = Split(UCase(Trim(rangeStr)), " TO ")
    If UBound(parts) < 1 Then
        ' TOが見つからない場合はそのまま出力
        ws.Cells(startRow, 1).Value = rangeStr
        ws.Cells(startRow, 2).Value = rangeStr
        ExpandOneRange = startRow + 1
        Exit Function
    End If

    ' 元の文字列からケースを保持して取得
    Dim originalParts() As String
    Dim toPos As Long
    toPos = InStr(UCase(Trim(rangeStr)), " TO ")
    idStart = Trim(Left(Trim(rangeStr), toPos - 1))
    idEnd = Trim(Mid(Trim(rangeStr), toPos + 4))

    ' 末尾の数字部分を抽出
    prefix = ExtractPrefix(idStart)
    numStart = ExtractNumber(idStart)
    numEnd = ExtractNumber(idEnd)

    ' 数字の桁数を保持（ゼロパディング用）
    numWidth = Len(idStart) - Len(prefix)

    ' 範囲チェック
    If numStart > numEnd Then
        Dim tmp As Long
        tmp = numStart
        numStart = numEnd
        numEnd = tmp
    End If

    row = startRow
    For i = numStart To numEnd
        ws.Cells(row, 1).Value = rangeStr
        ws.Cells(row, 2).Value = prefix & Format(i, String(numWidth, "0"))
        row = row + 1
    Next i

    ExpandOneRange = row
End Function

' 文字列から末尾の数字部分を除いたプレフィックスを取得
Private Function ExtractPrefix(ByVal s As String) As String
    Dim i As Long
    For i = Len(s) To 1 Step -1
        If Not IsNumeric(Mid(s, i, 1)) Then
            ExtractPrefix = Left(s, i)
            Exit Function
        End If
    Next i
    ExtractPrefix = ""
End Function

' 文字列から末尾の数字部分を取得
Private Function ExtractNumber(ByVal s As String) As Long
    Dim i As Long
    Dim numStr As String
    For i = Len(s) To 1 Step -1
        If Not IsNumeric(Mid(s, i, 1)) Then
            numStr = Mid(s, i + 1)
            If Len(numStr) > 0 Then
                ExtractNumber = CLng(numStr)
            Else
                ExtractNumber = 0
            End If
            Exit Function
        End If
    Next i
    ' 全部数字の場合
    ExtractNumber = CLng(s)
End Function
