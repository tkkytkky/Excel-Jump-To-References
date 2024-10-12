Private Sub Jump_Click()
    If ListBox.ListIndex <> -1 Then
        Dim selectionRef As String
        selectedRef = ListBox.Value
    
        Dim splitRef() As String
        Dim targetSheet As Worksheet
        Dim sheetName As String
        Dim formulaSheetName As String
    
        If InStr(selectedRef, "!") > 0 Then
            splitRef = Split(selectedRef, "!")
            sheetName = Replace(splitRef(0), "'", "")
            On Error Resume Next
            Set targetSheet = ActiveWorkbook.Sheets(sheetName)
            On Error GoTo 0
            If Not targetSheet Is Nothing Then
                targetSheet.Activate
                targetSheet.Range(splitRef(1)).Select
            Else
                MsgBox "シートが見つかりません: " & sheetName
            End If
        Else
            ' 同じシート内の参照
            ' Set targetSheet = ActiveSheet
            sheetName = SheetNameLabel.Caption
            Set targetSheet = Worksheets(sheetName)
            targetSheet.Activate
            targetSheet.Range(selectedRef).Select
        End If
    
        ' Unload Meはせず、ポップアップは開いたままにする
        ListBox.SetFocus
    Else
        MsgBox "セルを選択してください。"
    End If
End Sub
Private Sub UserForm_Initialize()
    ' 初期化時にフォーカスを設定
    Me.ListBox.SetFocus
End Sub
Private Sub ListBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
