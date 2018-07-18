Attribute VB_Name = "DeleteHyperLinks"
'選択範囲のハイパーリンクを削除する
'
'
Public Sub DeleteHyperLinks()

    Dim hyperlinksObj As Hyperlinks
    Dim tmpBk As Workbook
    Dim tmpR As Range
    Dim nowSht As Worksheet
    Dim nowAddress As String

    Dim cautionMessage As String: cautionMessage = "このSubプロシージャは、" & vbLf & _
                                                   "現在の選択範囲に対して変更を行います。" & vbLf & vbLf & _
                                                   "実行しますか?"
    
    '実行確認
    retVal = MsgBox(cautionMessage, vbOKCancel + vbExclamation)
    If retVal <> vbOK Then
        Exit Sub
        
    End If
    
    'シート選択状態チェック
    If ActiveWindow.SelectedSheets.count > 1 Then
        MsgBox "複数シートが選択されています" & vbLf & _
               "不要なシート選択を解除してください"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '選択範囲の保存
    Set nowSht = ActiveSheet
    nowAddress = Selection.Address
    
    'ハイパーリンクの削除
    For Each c In Selection
    
        Set hyperlinksObj = c.Hyperlinks
        numOfHyperlink = hyperlinksObj.count
        
        If (c.Address = c.MergeArea.Cells(1, 1).Address) Then '対象セルが結合セルの左上でない場合は、スキップ
        
            Set c = c.MergeArea
            
            If numOfHyperlink > 0 Then 'ハイパーリンクが存在する場合
                
                'tmpBookがなければ作成する
                If tmpBk Is Nothing Then
                    Set tmpBk = Workbooks.Add
                    
                End If
                
                Set tmpR = tmpBk.Sheets(1).Range(c.Address)
                
                '書式をtmpBookのセルにbackupする
                c.Copy
                tmpR.PasteSpecial _
                    Paste:=xlPasteFormats, _
                    Operation:=xlNone, _
                    SkipBlanks:=False, _
                    Transpose:=False
                
                For counter = 1 To numOfHyperlink
    
                    hyperlinksObj(counter).Delete
    
                Next counter
                
                'buckupした書式を貼り付ける
                tmpR.Copy
                c.PasteSpecial _
                    Paste:=xlPasteFormats, _
                    Operation:=xlNone, _
                    SkipBlanks:=False, _
                    Transpose:=False
                
            End If
            
        End If
        
    Next c
    
    'tmpBookがあれば保存せずに削除する
    If Not (tmpBk Is Nothing) Then
        tmpBk.Close SaveChanges:=False
        
    End If
    
    '選択範囲の復活
    nowSht.Range(nowAddress).Select
    
    Application.ScreenUpdating = True
    
    MsgBox "Done!"
    
End Sub
