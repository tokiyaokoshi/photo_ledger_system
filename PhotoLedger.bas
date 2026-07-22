Option Explicit

' ==========================================
' 写真台帳自動作成マクロ（最終ページの空白削除対応版）
' ==========================================
Sub CreatePhotoLedger()
    Dim fd As FileDialog
    Dim targetFolder As String
    
    ' 1. フォルダ選択ダイアログを表示
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "写真が入っているフォルダを選択してください"
    If fd.Show = -1 Then
        targetFolder = fd.SelectedItems(1)
    Else
        MsgBox "処理がキャンセルされました。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False ' 画面のチラつきを防止して高速化
    
    ' 2. ページ余白を「狭い（上下左右 1.27cm）」に設定
    Dim sec As Section
    For Each sec In ActiveDocument.Sections
        sec.PageSetup.TopMargin = CentimetersToPoints(1.27)
        sec.PageSetup.bottomMargin = CentimetersToPoints(1.27)
        sec.PageSetup.leftMargin = CentimetersToPoints(1.27)
        sec.PageSetup.rightMargin = CentimetersToPoints(1.27)
    Next sec
    
    ' 3. ファイル操作オブジェクトの作成（Late Binding）
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 文書の末尾にカーソルを移動
    Selection.EndKey Unit:=wdStory
    
    ' 4. フォルダ処理の開始
    Call ProcessFolder(fso, fso.GetFolder(targetFolder), 0, targetFolder)
    
    ' ==========================================
    ' ▼追加：最後の余分な改ページ（空白ページ）を削除
    ' ==========================================
    Selection.EndKey Unit:=wdStory
    Selection.TypeBackspace
    Selection.TypeBackspace
    ' ==========================================
    
    Application.ScreenUpdating = True
    MsgBox "写真台帳の作成が完了しました！", vbInformation
End Sub

' ------------------------------------------
' フォルダを再帰的に処理する関数
' ------------------------------------------
Sub ProcessFolder(fso As Object, folder As Object, depth As Integer, rootFolder As String)
    Dim file As Object
    Dim subFolder As Object
    Dim ext As String
    
    ' ▼1. 見出しの作成（ルートフォルダ以外）
    If folder.Path <> rootFolder Then
        Dim headingLevel As Integer
        headingLevel = depth
        If headingLevel < 1 Then headingLevel = 1
        If headingLevel > 9 Then headingLevel = 9
        
        ' 見出し名から「00_」などの数字プレフィックスを除去
        Dim displayName As String
        Dim underscorePos As Integer
        
        displayName = folder.Name
        underscorePos = InStr(displayName, "_")
        
        ' 最初の「_」が存在し、かつ「_」より前が数字の場合のみ除去
        If underscorePos > 1 Then
            If IsNumeric(Left(displayName, underscorePos - 1)) Then
                displayName = Mid(displayName, underscorePos + 1)
            End If
        End If
        
        Selection.EndKey Unit:=wdStory
        ' Wordの標準組み込みスタイル（wdStyleHeading1 = -2）を計算して適用
        Selection.Style = ActiveDocument.Styles(-2 - (headingLevel - 1))
        Selection.TypeText Text:=displayName
        Selection.TypeParagraph
        Selection.Style = ActiveDocument.Styles("標準") ' スタイルを標準に戻す
    End If
    
    ' ▼2. フォルダ内の画像ファイルだけを配列に集める
    Dim arrFiles() As String
    Dim fileCount As Integer: fileCount = 0
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "gif" Or ext = "bmp" Then
            fileCount = fileCount + 1
            ReDim Preserve arrFiles(1 To fileCount)
            arrFiles(fileCount) = file.Path
        End If
    Next file
    
    ' ▼3. 画像がある場合は表を作成
    If fileCount > 0 Then
        ' ファイル名順にソート
        Dim x As Integer, y As Integer, tempStr As String
        For x = 1 To fileCount - 1
            For y = x + 1 To fileCount
                If arrFiles(x) > arrFiles(y) Then
                    tempStr = arrFiles(x)
                    arrFiles(x) = arrFiles(y)
                    arrFiles(y) = tempStr
                End If
            Next y
        Next x
        
        ' 3枚ずつ処理する
        Dim i As Integer, r As Integer
        Dim chunkSize As Integer: chunkSize = 3
        Dim currentChunkSize As Integer
        Dim tbl As Table
        Dim cellRange As Range
        Dim shp As InlineShape
        
        For i = 1 To fileCount Step chunkSize
            currentChunkSize = fileCount - i + 1
            If currentChunkSize > chunkSize Then currentChunkSize = chunkSize
            
            Selection.EndKey Unit:=wdStory
            
            ' 表の「上」に改行を追加
            Selection.TypeParagraph
            
            Set tbl = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=currentChunkSize + 1, NumColumns:=2)
            tbl.Borders.Enable = True ' 表に格子線を引く
            
            ' 表の横幅を余白いっぱいに広げる
            tbl.PreferredWidthType = wdPreferredWidthPercent
            tbl.PreferredWidth = 100 ' ページ幅100%
            
            ' 左の画像列を72%、右の備考列を28%の比率で分割
            tbl.Columns(1).PreferredWidthType = wdPreferredWidthPercent
            tbl.Columns(1).PreferredWidth = 72
            tbl.Columns(2).PreferredWidthType = wdPreferredWidthPercent
            tbl.Columns(2).PreferredWidth = 28
            
            ' ヘッダーの設定
            tbl.Cell(1, 1).Range.Text = "表示画面"
            tbl.Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            tbl.Cell(1, 2).Range.Text = "備 考"
            tbl.Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            ' データ行の設定
            For r = 1 To currentChunkSize
                ' 画像の挿入（左セル）
                Set cellRange = tbl.Cell(r + 1, 1).Range
                cellRange.Collapse wdCollapseStart
                cellRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
                
                Set shp = ActiveDocument.InlineShapes.AddPicture(FileName:=arrFiles(i + r - 1), LinkToFile:=False, SaveWithDocument:=True, Range:=cellRange)
                
                ' 画像サイズ調整（縦横比維持）
                shp.LockAspectRatio = msoTrue
                
                ' 1. まず1ページに3枚収まる安全な高さ（7.5cm）を指定
                shp.Height = CentimetersToPoints(7.5)
                
                ' 2. 表の幅が広がったのに合わせて、限界幅を13cmに拡張
                If shp.Width > CentimetersToPoints(13) Then
                    shp.Width = CentimetersToPoints(13)
                End If
                
                ' 備考（ファイル名）の挿入（右セル）
                tbl.Cell(r + 1, 2).Range.Text = fso.GetBaseName(arrFiles(i + r - 1))
            Next r
            
            Selection.EndKey Unit:=wdStory
            Selection.InsertBreak Type:=wdPageBreak
            
            ' カーソルを1つ左（改ページ記号の直前）に戻す
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            ' Backspaceキーを押して、表と改ページの間の隙間を削除する
            Selection.TypeBackspace
            
            ' 次の処理のためにカーソルを一番後ろに戻す
            Selection.EndKey Unit:=wdStory
        Next i
    End If
    
    ' ▼4. サブフォルダの処理（再帰）
    Dim arrFolders() As Object
    Dim fldCount As Integer: fldCount = folder.SubFolders.Count
    If fldCount > 0 Then
        ReDim arrFolders(1 To fldCount)
        Dim fldIdx As Integer: fldIdx = 1
        For Each subFolder In folder.SubFolders
            Set arrFolders(fldIdx) = subFolder
            fldIdx = fldIdx + 1
        Next subFolder
        
        ' フォルダ名順にソート
        Dim tempFld As Object
        For x = 1 To fldCount - 1
            For y = x + 1 To fldCount
                If arrFolders(x).Name > arrFolders(y).Name Then
                    Set tempFld = arrFolders(x)
                    Set arrFolders(x) = arrFolders(y)
                    Set arrFolders(y) = tempFld
                End If
            Next y
        Next x
        
        ' ソート順にサブフォルダを処理
        For x = 1 To fldCount
            Call ProcessFolder(fso, arrFolders(x), depth + 1, rootFolder)
        Next x
    End If
End Sub