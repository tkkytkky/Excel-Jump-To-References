Sub Jump()
    Dim rng As Range
    Dim cell As Range
    Dim formula As String
    Dim cellAddresses As Collection
    Dim refSet As Object
    
    ' JumpFormがすでに開いている場合は閉じる
    If Not JumpForm Is Nothing Then
        If JumpForm.Visible Then
            Unload JumpForm
        End If
    End If
    
    ' 現在選択されているセルを取得
    Set rng = Selection
    Set cell = rng.Cells(1, 1)
    formula = cell.formula
    If Left(formula, 1) = "=" Then
        formula = Mid(formula, 2) ' 先頭の=だけを除去
    End If
    
    ' 参照先セルを格納するコレクションを作成
    Set cellAddresses = New Collection
    Set refSet = CreateObject("Scripting.Dictionary")
    
    ' シート名付きセル参照を取得
    Call ExtractSheetReferences(formula, cellAddresses, refSet)
    
    ' シート名なしのセル参照を取得
    Call ExtractSimpleReferences(formula, cellAddresses, refSet)
    
    ' UserFormを表示
    Call ShowReferences(cellAddresses)
End Sub
Private Sub ExtractSheetReferences(ByVal formula As String, ByRef cellAddresses As Collection, ByRef refSet As Object)
    Dim sheetNames As Collection
    Dim sheet As Worksheet
    Dim sheetName As Variant
    Dim startPos As Long
    Dim endPos As Long
    Dim ref As String
    Dim delimiters As String
    Dim patterns As Variant
    
    ' シート名を数式内で検索
    Set sheetNames = New Collection
    For Each sheet In ActiveWorkbook.Sheets
        sheetNames.Add sheet.Name
    Next sheet
    
    'セル参照の終端となる文字
    delimiters = "+-*/&),"
    
    ' シート名を数式内で検索
    For Each sheetName In sheetNames
        patterns = Array("'" & sheetName & "'!", sheetName & "!")
        
        For Each Pattern In patterns
            startPos = 1
            Do While startPos > 0
                startPos = InStr(startPos, formula, Pattern)
                If startPos > 0 Then
                    endPos = startPos + Len(Pattern)
                    
                    ' セル参照の終わりを探す
                    Do While endPos <= Len(formula) And InStr(delimiters, Mid(formula, endPos, 1)) = 0
                        endPos = endPos + 1
                    Loop
                    
                    ref = Mid(formula, startPos, endPos - startPos)
                    If Not refSet.Exists(ref) Then
                        cellAddresses.Add ref
                        refSet.Add ref, True
                    End If
                    startPos = endPos
                End If
            Loop
        Next Pattern
    Next sheetName
End Sub
Private Sub ExtractSimpleReferences(ByVal formula As String, ByRef cellAddresses As Collection, ByRef refSet As Object)
    Dim regex As Object
    Dim matches As Object
    Dim ref As Object
    Dim foundInSheet As Boolean
    Dim item As Variant
    
    ' 正規表現のセットアップ
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "([A-Z]+\d+)"
    
    ' すべてのセル参照を抽出
    If regex.Test(formula) Then
        Set matches = regex.Execute(formula)
        For Each ref In matches
            foundInSheet = False
            
            For Each item In cellAddresses
                ' シート名がある場合に同じ参照を探す
                If InStr(item, ref.Value) > 0 Then
                    foundInSheet = True
                    Exit For
                End If
            Next item
            
            'シート名なしで重複しない場合のみ追加
            If Not foundInSheet And Not refSet.Exists(ref.Value) Then
                cellAddresses.Add ref.Value
                refSet.Add ref.Value, True
            End If
        Next ref
    End If
End Sub
Private Sub ShowReferences(ByVal cellAddresses As Collection)
    Dim form As New JumpForm
    Dim item As Variant
    Dim cell As Range
    Dim leftPos As Long
    Dim topPos As Long
    
    ' 現在選択されているセルを取得
    Set cell = Selection.Cells(1, 1)
    
    ' フォームのデフォルト位置を設定
    leftPos = cell.Left + cell.Width
    topPos = cell.Top
    
    form.StartUpPosition = 0 ' 位置を手動設定に
    form.Left = leftPos
    form.Top = topPos
    
    ' フォームの位置調整
    If form.Left < cell.Left + cell.Width And form.Left + form.Width > cell.Left Then
        form.Left = cell.Left - form.Width
    End If
    
    If form.Top < cell.Top + cell.Height And form.Top + form.Height > cell.Top Then
        form.Top = cell.Top + cell.Height
    End If
    
    ' 数式をラベルに表示
    form.FormulaLabel.Caption = cell.formula
    
    ' シート名をラベルに表示
    form.SheetNameLabel.Caption = ActiveSheet.Name
    
    For Each item In cellAddresses
        form.ListBox.AddItem item
    Next item
    
    form.Show
End Sub
