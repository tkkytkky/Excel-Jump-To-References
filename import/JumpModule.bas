Attribute VB_Name = "JumpModule"
Sub Jump()
Attribute Jump.VB_ProcData.VB_Invoke_Func = "J\n14"
    Dim rng As Range
    Dim cell As Range
    Dim formula As String
    Dim cellAddresses As Collection
    Dim refSet As Object
    
    ' JumpForm�����łɊJ���Ă���ꍇ�͕���
    If Not JumpForm Is Nothing Then
        If JumpForm.Visible Then
            Unload JumpForm
        End If
    End If
    
    ' ���ݑI������Ă���Z�����擾
    Set rng = Selection
    Set cell = rng.Cells(1, 1)
    formula = cell.formula
    If Left(formula, 1) = "=" Then
        formula = Mid(formula, 2) ' �擪��=����������
    End If
    
    ' �Q�Ɛ�Z�����i�[����R���N�V�������쐬
    Set cellAddresses = New Collection
    Set refSet = CreateObject("Scripting.Dictionary")
    
    ' �V�[�g���t���Z���Q�Ƃ��擾
    Call ExtractSheetReferences(formula, cellAddresses, refSet)
    
    ' �V�[�g���Ȃ��̃Z���Q�Ƃ��擾
    Call ExtractSimpleReferences(formula, cellAddresses, refSet)
    
    ' UserForm��\��
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
    
    ' �V�[�g���𐔎����Ō���
    Set sheetNames = New Collection
    For Each sheet In ActiveWorkbook.Sheets
        sheetNames.Add sheet.Name
    Next sheet
    
    '�Z���Q�Ƃ̏I�[�ƂȂ镶��
    delimiters = "+-*/&),"
    
    ' �V�[�g���𐔎����Ō���
    For Each sheetName In sheetNames
        patterns = Array("'" & sheetName & "'!", sheetName & "!")
        
        For Each Pattern In patterns
            startPos = 1
            Do While startPos > 0
                startPos = InStr(startPos, formula, Pattern)
                If startPos > 0 Then
                    endPos = startPos + Len(Pattern)
                    
                    ' �Z���Q�Ƃ̏I����T��
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
    
    ' ���K�\���̃Z�b�g�A�b�v
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "([A-Z]+\d+)"
    
    ' ���ׂẴZ���Q�Ƃ𒊏o
    If regex.Test(formula) Then
        Set matches = regex.Execute(formula)
        For Each ref In matches
            foundInSheet = False
            
            For Each item In cellAddresses
                ' �V�[�g��������ꍇ�ɓ����Q�Ƃ�T��
                If InStr(item, ref.Value) > 0 Then
                    foundInSheet = True
                    Exit For
                End If
            Next item
            
            '�V�[�g���Ȃ��ŏd�����Ȃ��ꍇ�̂ݒǉ�
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
    
    ' ���ݑI������Ă���Z�����擾
    Set cell = Selection.Cells(1, 1)
    
    ' �t�H�[���̃f�t�H���g�ʒu��ݒ�
    leftPos = cell.Left + cell.Width
    topPos = cell.Top
    
    form.StartUpPosition = 0 ' �ʒu���蓮�ݒ��
    form.Left = leftPos
    form.Top = topPos
    
    ' �t�H�[���̈ʒu����
    If form.Left < cell.Left + cell.Width And form.Left + form.Width > cell.Left Then
        form.Left = cell.Left - form.Width
    End If
    
    If form.Top < cell.Top + cell.Height And form.Top + form.Height > cell.Top Then
        form.Top = cell.Top + cell.Height
    End If
    
    ' ���������x���ɕ\��
    form.FormulaLabel.Caption = cell.formula
    
    ' �V�[�g�������x���ɕ\��
    form.SheetNameLabel.Caption = ActiveSheet.Name
    
    For Each item In cellAddresses
        form.ListBox.AddItem item
    Next item
    
    form.Show
End Sub
