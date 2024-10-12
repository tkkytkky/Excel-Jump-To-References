VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JumpForm 
   Caption         =   "JumpForm"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7995
   OleObjectBlob   =   "JumpForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "JumpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
                MsgBox "�V�[�g��������܂���: " & sheetName
            End If
        Else
            ' �����V�[�g���̎Q��
            ' Set targetSheet = ActiveSheet
            sheetName = SheetNameLabel.Caption
            Set targetSheet = Worksheets(sheetName)
            targetSheet.Activate
            targetSheet.Range(selectedRef).Select
        End If
    
        ' Unload Me�͂����A�|�b�v�A�b�v�͊J�����܂܂ɂ���
        ListBox.SetFocus
    Else
        MsgBox "�Z����I�����Ă��������B"
    End If
End Sub
Private Sub UserForm_Initialize()
    ' ���������Ƀt�H�[�J�X��ݒ�
    Me.ListBox.SetFocus
End Sub
Private Sub ListBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
