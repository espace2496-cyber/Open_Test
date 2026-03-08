'==============================================================================
' PDF to PowerPoint Converter - Excel VBA Macro
'==============================================================================
' 概要: PDFファイルを画像に変換し、全ページをPowerPointに貼り付けます
'
' 前提条件:
'   1. Ghostscript がインストールされていること
'      ダウンロード: https://www.ghostscript.com/releases/gsdnld.html
'   2. VBAの参照設定で以下を追加すること:
'      - Microsoft PowerPoint XX.X Object Library
'      - Microsoft Scripting Runtime
'==============================================================================

Option Explicit

' --- 設定値 ---
Private Const GS_PATH As String = "C:\Program Files\gs\gs10.02.1\bin\gswin64c.exe"
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

    ' PDFのページ数を取得
    pageCount = GetPdfPageCount(pdfPath)
    If pageCount = 0 Then
        MsgBox "PDFのページ数を取得できませんでした。" & vbCrLf & _
               "Ghostscriptのパスを確認してください。" & vbCrLf & _
               "現在の設定: " & GS_PATH, vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "PDF変換中... 全" & pageCount & "ページ"

    ' PDFを画像に変換
    If Not ConvertPdfToImages(pdfPath, outputFolder, pageCount) Then
        MsgBox "PDFから画像への変換に失敗しました。", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If

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
' GhostscriptでPDFのページ数を取得
'==============================================================================
Private Function GetPdfPageCount(ByVal pdfPath As String) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ghostscriptが存在するか確認
    If Not fso.FileExists(GS_PATH) Then
        ' よく使われるパスを自動検索
        Dim altPath As String
        altPath = FindGhostscript()
        If altPath = "" Then
            GetPdfPageCount = 0
            Exit Function
        End If
    End If

    Dim gsExe As String
    If fso.FileExists(GS_PATH) Then
        gsExe = GS_PATH
    Else
        gsExe = FindGhostscript()
    End If

    ' ページ数取得用のPostScriptコマンド
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\pdf_pagecount.txt"

    Dim cmd As String
    cmd = """" & gsExe & """ -q -dNODISPLAY -dNOSAFER --permit-file-read=""" & pdfPath & """ -c """ & _
          "(" & Replace(pdfPath, "\", "/") & ") (r) file runpdfbegin pdfpagecount = quit"""

    ' WScript.Shell を使用してコマンドを実行し、出力を取得
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    Dim exec As Object
    Set exec = wsh.exec("cmd /c " & cmd)

    ' 出力を待つ
    Do While exec.Status = 0
        DoEvents
    Loop

    Dim output As String
    output = Trim(exec.StdOut.ReadAll)

    If IsNumeric(output) Then
        GetPdfPageCount = CLng(output)
    Else
        ' 別の方法でページ数を取得
        GetPdfPageCount = GetPdfPageCountAlt(pdfPath, gsExe)
    End If
End Function

'==============================================================================
' 代替方法でPDFのページ数を取得
'==============================================================================
Private Function GetPdfPageCountAlt(ByVal pdfPath As String, ByVal gsExe As String) As Long
    ' 1ページずつ変換を試みてページ数を数える方法
    Dim tempOut As String
    tempOut = Environ("TEMP") & "\gs_test_page.png"

    Dim page As Long
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    For page = 1 To 500  ' 最大500ページまで
        Dim cmd As String
        cmd = """" & gsExe & """ -q -dNOPAUSE -dBATCH -dSAFER " & _
              "-dFirstPage=" & page & " -dLastPage=" & page & " " & _
              "-sDEVICE=pngalpha -r72 -o """ & tempOut & """ """ & pdfPath & """"

        Dim exitCode As Long
        exitCode = wsh.Run("cmd /c " & cmd, 0, True)

        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")

        If Not fso.FileExists(tempOut) Then
            GetPdfPageCountAlt = page - 1
            Exit Function
        End If

        ' ファイルサイズが0なら終了
        If fso.GetFile(tempOut).Size = 0 Then
            fso.DeleteFile tempOut
            GetPdfPageCountAlt = page - 1
            Exit Function
        End If

        fso.DeleteFile tempOut
    Next page

    GetPdfPageCountAlt = page - 1
End Function

'==============================================================================
' Ghostscriptの自動検索
'==============================================================================
Private Function FindGhostscript() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' よく使われるインストールパスを検索
    Dim basePaths As Variant
    basePaths = Array( _
        "C:\Program Files\gs", _
        "C:\Program Files (x86)\gs" _
    )

    Dim basePath As Variant
    For Each basePath In basePaths
        If fso.FolderExists(CStr(basePath)) Then
            Dim folder As Object
            Set folder = fso.GetFolder(CStr(basePath))

            Dim subFolder As Object
            For Each subFolder In folder.SubFolders
                Dim gsExe As String
                gsExe = subFolder.Path & "\bin\gswin64c.exe"
                If fso.FileExists(gsExe) Then
                    FindGhostscript = gsExe
                    Exit Function
                End If

                gsExe = subFolder.Path & "\bin\gswin32c.exe"
                If fso.FileExists(gsExe) Then
                    FindGhostscript = gsExe
                    Exit Function
                End If
            Next subFolder
        End If
    Next basePath

    FindGhostscript = ""
End Function

'==============================================================================
' PDFを画像（PNG）に変換
'==============================================================================
Private Function ConvertPdfToImages(ByVal pdfPath As String, _
                                     ByVal outputFolder As String, _
                                     ByVal pageCount As Long) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim gsExe As String
    If fso.FileExists(GS_PATH) Then
        gsExe = GS_PATH
    Else
        gsExe = FindGhostscript()
    End If

    If gsExe = "" Then
        ConvertPdfToImages = False
        Exit Function
    End If

    ' 全ページを一括で画像に変換
    Dim outputPattern As String
    outputPattern = outputFolder & "\page_%03d.png"

    Dim cmd As String
    cmd = """" & gsExe & """ -dNOPAUSE -dBATCH -dSAFER " & _
          "-sDEVICE=pngalpha -r" & DPI & " " & _
          "-o """ & outputPattern & """ """ & pdfPath & """"

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    Dim exitCode As Long
    exitCode = wsh.Run("cmd /c " & cmd, 0, True)

    ' 変換結果を確認（最低1ページ分の画像が存在するか）
    ConvertPdfToImages = fso.FileExists(outputFolder & "\page_001.png")
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

    ' スライドサイズをA4横に設定（必要に応じて変更可能）
    ' 標準(4:3): 幅=720pt, 高さ=540pt
    ' ワイド(16:9): 幅=960pt, 高さ=540pt
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

        ' 空白スライドを追加（レイアウト7 = ppLayoutBlank）
        Dim slide As Object
        Set slide = pres.Slides.Add(i, 12)  ' 12 = ppLayoutBlank

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
