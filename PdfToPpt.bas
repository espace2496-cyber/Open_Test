'==============================================================================
' PDF to PowerPoint Converter - Excel VBA Macro
'==============================================================================
' 概要: PDFファイルを画像に変換し、全ページをPowerPointに貼り付けます
'
' 前提条件:
'   - Windows 10 以降（標準のPDF描画APIを使用）
'   - Microsoft PowerPoint がインストールされていること
'   - 外部ソフトのインストールは不要です
'
' VBAの参照設定で以下を追加すること:
'   - Microsoft PowerPoint XX.X Object Library
'==============================================================================

Option Explicit

' --- 設定値 ---
Private Const DPI As Long = 200  ' 画像の解像度（DPI）

'==============================================================================
' メイン処理: PDFを選択してPowerPointに変換する
'==============================================================================
Public Sub ConvertPdfToPpt()
    Dim pdfPath As String
    Dim outputFolder As String
    Dim pageCount As Long
    Dim pptPath As String

    ' PDFファイルを選択
    pdfPath = SelectPdfFile()
    If pdfPath = "" Then
        MsgBox "PDFファイルが選択されませんでした。", vbInformation
        Exit Sub
    End If

    ' 一時フォルダを作成
    outputFolder = CreateTempFolder(pdfPath)

    Application.StatusBar = "PDF変換中... ページ数を取得しています"

    ' PowerShell + Windows.Data.Pdf でPDFを画像に変換
    pageCount = ConvertPdfToImages(pdfPath, outputFolder)
    If pageCount = 0 Then
        MsgBox "PDFから画像への変換に失敗しました。" & vbCrLf & vbCrLf & _
               "以下を確認してください:" & vbCrLf & _
               "- Windows 10 以降であること" & vbCrLf & _
               "- PowerShellが使用可能であること" & vbCrLf & _
               "- PDFファイルが破損していないこと", vbExclamation
        CleanupTempFolder outputFolder
        Application.StatusBar = False
        Exit Sub
    End If

    Application.StatusBar = "PDF変換完了（" & pageCount & "ページ）。PowerPoint作成中..."

    ' PowerPointに画像を貼り付け
    pptPath = CreatePptFromImages(outputFolder, pageCount, pdfPath)

    ' 一時ファイルを削除
    CleanupTempFolder outputFolder

    Application.StatusBar = False

    If pptPath <> "" Then
        MsgBox "PowerPointファイルを作成しました:" & vbCrLf & pptPath, vbInformation
    End If
End Sub

'==============================================================================
' PDFファイル選択ダイアログ
'==============================================================================
Private Function SelectPdfFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "PDFファイルを選択してください"
        .Filters.Clear
        .Filters.Add "PDFファイル", "*.pdf"
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectPdfFile = .SelectedItems(1)
        Else
            SelectPdfFile = ""
        End If
    End With
End Function

'==============================================================================
' 一時フォルダを作成
'==============================================================================
Private Function CreateTempFolder(ByVal pdfPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempFolder As String
    tempFolder = fso.GetParentFolderName(pdfPath) & "\" & _
                 fso.GetBaseName(pdfPath) & "_temp_images"

    If Not fso.FolderExists(tempFolder) Then
        fso.CreateFolder tempFolder
    End If

    CreateTempFolder = tempFolder
End Function

'==============================================================================
' PowerShell + Windows.Data.Pdf API でPDFを画像に変換
' Windows 10/11 標準搭載のAPIを使用（外部ソフト不要）
' 戻り値: 変換したページ数（失敗時は0）
'==============================================================================
Private Function ConvertPdfToImages(ByVal pdfPath As String, _
                                     ByVal outputFolder As String) As Long
    ' PowerShellスクリプトを一時ファイルに書き出して実行
    Dim psScriptPath As String
    psScriptPath = Environ("TEMP") & "\pdf_to_images.ps1"

    Dim resultFile As String
    resultFile = Environ("TEMP") & "\pdf_convert_result.txt"

    ' PowerShellスクリプトを生成
    Dim psCode As String
    psCode = BuildPowerShellScript(pdfPath, outputFolder, resultFile)

    ' スクリプトファイルに書き出し
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ts As Object
    Set ts = fso.CreateTextFile(psScriptPath, True, False)
    ts.Write psCode
    ts.Close

    ' PowerShellを実行
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & psScriptPath & """"

    Dim exitCode As Long
    exitCode = wsh.Run(cmd, 0, True)

    ' 結果ファイルからページ数を読み取る
    Dim pageCount As Long
    pageCount = 0

    If fso.FileExists(resultFile) Then
        Set ts = fso.OpenTextFile(resultFile, 1)
        Dim resultText As String
        resultText = Trim(ts.ReadAll)
        ts.Close
        fso.DeleteFile resultFile

        If IsNumeric(resultText) Then
            pageCount = CLng(resultText)
        End If
    End If

    ' スクリプトファイルを削除
    If fso.FileExists(psScriptPath) Then
        fso.DeleteFile psScriptPath
    End If

    ConvertPdfToImages = pageCount
End Function

'==============================================================================
' PowerShellスクリプトを生成
' Windows.Data.Pdf (WinRT API) を使用してPDFを画像に変換
'==============================================================================
Private Function BuildPowerShellScript(ByVal pdfPath As String, _
                                        ByVal outputFolder As String, _
                                        ByVal resultFile As String) As String
    Dim s As String

    ' エスケープ処理
    Dim escapedPdf As String
    Dim escapedOut As String
    Dim escapedResult As String
    escapedPdf = Replace(pdfPath, "'", "''")
    escapedOut = Replace(outputFolder, "'", "''")
    escapedResult = Replace(resultFile, "'", "''")

    s = ""
    s = s & "# PDF to Images using Windows.Data.Pdf (Windows 10+ built-in API)" & vbCrLf
    s = s & "try {" & vbCrLf
    s = s & "    # Load required WinRT assemblies" & vbCrLf
    s = s & "    Add-Type -AssemblyName System.Runtime.WindowsRuntime" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    # Helper function to await WinRT async operations" & vbCrLf
    s = s & "    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() |" & vbCrLf
    s = s & "        Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and" & vbCrLf
    s = s & "        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation``1' })[0]" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Function Await($WinRtTask, $ResultType) {" & vbCrLf
    s = s & "        $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)" & vbCrLf
    s = s & "        $netTask = $asTask.Invoke($null, @($WinRtTask))" & vbCrLf
    s = s & "        $netTask.Wait(-1) | Out-Null" & vbCrLf
    s = s & "        $netTask.Result" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    # Also need AsTask for IAsyncAction (no result)" & vbCrLf
    s = s & "    $asTaskAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() |" & vbCrLf
    s = s & "        Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and" & vbCrLf
    s = s & "        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Function AwaitAction($WinRtTask) {" & vbCrLf
    s = s & "        $netTask = $asTaskAction.Invoke($null, @($WinRtTask))" & vbCrLf
    s = s & "        $netTask.Wait(-1) | Out-Null" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    # Load Windows.Data.Pdf WinRT type" & vbCrLf
    s = s & "    [Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime] | Out-Null" & vbCrLf
    s = s & "    [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime] | Out-Null" & vbCrLf
    s = s & "    [Windows.Storage.Streams.RandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime] | Out-Null" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    # Open PDF file" & vbCrLf
    s = s & "    $pdfPath = '" & escapedPdf & "'" & vbCrLf
    s = s & "    $outputFolder = '" & escapedOut & "'" & vbCrLf
    s = s & "    $dpi = " & DPI & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    $storageFile = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdfPath)) ([Windows.Storage.StorageFile])" & vbCrLf
    s = s & "    $pdfDoc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)) ([Windows.Data.Pdf.PdfDocument])" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    $pageCount = $pdfDoc.PageCount" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    # Render each page to PNG" & vbCrLf
    s = s & "    for ($i = 0; $i -lt $pageCount; $i++) {" & vbCrLf
    s = s & "        $page = $pdfDoc.GetPage($i)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        # Set render options for desired DPI" & vbCrLf
    s = s & "        $renderOptions = New-Object Windows.Data.Pdf.PdfPageRenderOptions" & vbCrLf
    s = s & "        $scale = $dpi / 96.0" & vbCrLf
    s = s & "        $renderOptions.DestinationWidth  = [uint32]($page.Size.Width * $scale)" & vbCrLf
    s = s & "        $renderOptions.DestinationHeight = [uint32]($page.Size.Height * $scale)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        # Create output stream" & vbCrLf
    s = s & "        $pageNum = ($i + 1).ToString('000')" & vbCrLf
    s = s & "        $outputPath = Join-Path $outputFolder ""page_$pageNum.png""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream" & vbCrLf
    s = s & "        AwaitAction ($page.RenderToStreamAsync($stream, $renderOptions))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        # Save stream to file" & vbCrLf
    s = s & "        $fileStream = [System.IO.File]::Create($outputPath)" & vbCrLf
    s = s & "        $stream.Seek(0)" & vbCrLf
    s = s & "        $dotNetStream = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($stream)" & vbCrLf
    s = s & "        $dotNetStream.CopyTo($fileStream)" & vbCrLf
    s = s & "        $fileStream.Close()" & vbCrLf
    s = s & "        $dotNetStream.Close()" & vbCrLf
    s = s & "        $stream.Dispose()" & vbCrLf
    s = s & "        $page.Dispose()" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    # Write page count to result file" & vbCrLf
    s = s & "    $pageCount | Out-File -FilePath '" & escapedResult & "' -NoNewline" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "catch {" & vbCrLf
    s = s & "    '0' | Out-File -FilePath '" & escapedResult & "' -NoNewline" & vbCrLf
    s = s & "    Write-Error $_.Exception.Message" & vbCrLf
    s = s & "}" & vbCrLf

    BuildPowerShellScript = s
End Function

'==============================================================================
' 画像からPowerPointを作成
'==============================================================================
Private Function CreatePptFromImages(ByVal imageFolder As String, _
                                      ByVal pageCount As Long, _
                                      ByVal pdfPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' PowerPointアプリケーションを起動
    Dim pptApp As Object
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    ' 新しいプレゼンテーションを作成
    Dim pres As Object
    Set pres = pptApp.Presentations.Add

    ' スライドサイズを取得
    Dim slideWidth As Single
    Dim slideHeight As Single
    slideWidth = pres.PageSetup.slideWidth
    slideHeight = pres.PageSetup.slideHeight

    Dim i As Long
    For i = 1 To pageCount
        Dim imagePath As String
        imagePath = imageFolder & "\page_" & Format(i, "000") & ".png"

        If Not fso.FileExists(imagePath) Then
            GoTo NextPage
        End If

        Application.StatusBar = "PowerPoint作成中... " & i & "/" & pageCount & " ページ"

        ' 空白スライドを追加（12 = ppLayoutBlank）
        Dim slide As Object
        Set slide = pres.Slides.Add(i, 12)

        ' 画像を挿入
        Dim pic As Object
        Set pic = slide.Shapes.AddPicture( _
            FileName:=imagePath, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=0, _
            Top:=0)

        ' 画像をスライドに合わせてリサイズ（アスペクト比を維持）
        Dim scaleW As Single
        Dim scaleH As Single
        scaleW = slideWidth / pic.Width
        scaleH = slideHeight / pic.Height

        Dim scaleFactor As Single
        If scaleW < scaleH Then
            scaleFactor = scaleW
        Else
            scaleFactor = scaleH
        End If

        pic.Width = pic.Width * scaleFactor
        pic.Height = pic.Height * scaleFactor

        ' 画像を中央に配置
        pic.Left = (slideWidth - pic.Width) / 2
        pic.Top = (slideHeight - pic.Height) / 2

NextPage:
    Next i

    ' PowerPointファイルを保存
    Dim pptPath As String
    pptPath = fso.GetParentFolderName(pdfPath) & "\" & _
              fso.GetBaseName(pdfPath) & ".pptx"

    pres.SaveAs pptPath

    CreatePptFromImages = pptPath
End Function

'==============================================================================
' 一時フォルダとファイルを削除
'==============================================================================
Private Sub CleanupTempFolder(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then
        fso.DeleteFolder folderPath, True
    End If
End Sub
