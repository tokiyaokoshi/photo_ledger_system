Option Explicit

' ==========================================
' 写真台帳自動作成マクロ（3枚グループ化対応）
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
        
        Selection.EndKey Unit:=wdStory
        ' Wordの標準組み込みスタイル（wdStyleHeading1 = -2）を計算して適用
        Selection.Style = ActiveDocument.Styles(-2 - (headingLevel - 1))
        Selection.TypeText Text:=folder.Name
        Selection.TypeParagraph
        Selection.Style = ActiveDocument.Styles("標準") ' スタイルを標準に戻す
    End If
    
    ' ▼2. フォルダ内の画像ファイルだけを配列に集める（Pythonの valid_images に相当）
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
        ' ファイル名順にソート（Pythonの sorted に相当）
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
            Set tbl = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=currentChunkSize + 1, NumColumns:=2)
            tbl.Borders.Enable = True ' 表に格子線を引く
            
            ' 列幅の設定 (11.5cm と 4.5cm)
            tbl.Columns(1).PreferredWidthType = wdPreferredWidthPoints
            tbl.Columns(1).PreferredWidth = CentimetersToPoints(11.5)
            tbl.Columns(2).PreferredWidthType = wdPreferredWidthPoints
            tbl.Columns(2).PreferredWidth = CentimetersToPoints(4.5)
            
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
                shp.LockAspectRatio = msoTrue
                shp.Width = CentimetersToPoints(11) ' 画像幅を11cmに縮小
                
                ' 備考（ファイル名）の挿入（右セル）
                tbl.Cell(r + 1, 2).Range.Text = fso.GetBaseName(arrFiles(i + r - 1))
            Next r
            
            ' 表の後に改ページを入れる
            Selection.EndKey Unit:=wdStory
            Selection.InsertBreak Type:=wdPageBreak
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
