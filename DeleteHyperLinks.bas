Attribute VB_Name = "DeleteHyperLinks"
'�I��͈͂̃n�C�p�[�����N���폜����
'
'
Public Sub DeleteHyperLinks()

    Dim hyperlinksObj As Hyperlinks
    Dim tmpBk As Workbook
    Dim tmpR As Range
    Dim nowSht As Worksheet
    Dim nowAddress As String

    Dim cautionMessage As String: cautionMessage = "����Sub�v���V�[�W���́A" & vbLf & _
                                                   "���݂̑I��͈͂ɑ΂��ĕύX���s���܂��B" & vbLf & vbLf & _
                                                   "���s���܂���?"
    
    '���s�m�F
    retVal = MsgBox(cautionMessage, vbOKCancel + vbExclamation)
    If retVal <> vbOK Then
        Exit Sub
        
    End If
    
    '�V�[�g�I����ԃ`�F�b�N
    If ActiveWindow.SelectedSheets.count > 1 Then
        MsgBox "�����V�[�g���I������Ă��܂�" & vbLf & _
               "�s�v�ȃV�[�g�I�����������Ă�������"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '�I��͈͂̕ۑ�
    Set nowSht = ActiveSheet
    nowAddress = Selection.Address
    
    '�n�C�p�[�����N�̍폜
    For Each c In Selection
    
        Set hyperlinksObj = c.Hyperlinks
        numOfHyperlink = hyperlinksObj.count
        
        If (c.Address = c.MergeArea.Cells(1, 1).Address) Then '�ΏۃZ���������Z���̍���łȂ��ꍇ�́A�X�L�b�v
        
            Set c = c.MergeArea
            
            If numOfHyperlink > 0 Then '�n�C�p�[�����N�����݂���ꍇ
                
                'tmpBook���Ȃ���΍쐬����
                If tmpBk Is Nothing Then
                    Set tmpBk = Workbooks.Add
                    
                End If
                
                Set tmpR = tmpBk.Sheets(1).Range(c.Address)
                
                '������tmpBook�̃Z����backup����
                c.Copy
                tmpR.PasteSpecial _
                    Paste:=xlPasteFormats, _
                    Operation:=xlNone, _
                    SkipBlanks:=False, _
                    Transpose:=False
                
                For counter = 1 To numOfHyperlink
    
                    hyperlinksObj(counter).Delete
    
                Next counter
                
                'buckup����������\��t����
                tmpR.Copy
                c.PasteSpecial _
                    Paste:=xlPasteFormats, _
                    Operation:=xlNone, _
                    SkipBlanks:=False, _
                    Transpose:=False
                
            End If
            
        End If
        
    Next c
    
    'tmpBook������Εۑ������ɍ폜����
    If Not (tmpBk Is Nothing) Then
        tmpBk.Close SaveChanges:=False
        
    End If
    
    '�I��͈͂̕���
    nowSht.Range(nowAddress).Select
    
    Application.ScreenUpdating = True
    
    MsgBox "Done!"
    
End Sub
