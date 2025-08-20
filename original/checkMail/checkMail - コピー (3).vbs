'******説明******************************************************************************
'　保守用のRJ共有メールアドレス宛のメールを、PVC保守用の共有メールアドレスへ転送します
'****************************************************************************************
Option Explicit
On Error Resume Next

Dim objShell ' shell オブジェクト
Dim objOlApp    ' outlook オブジェクト
Dim objOlNs     ' NameSpace オブジェクト
Dim objOlFolder ' フォルダ情報
Dim objOlItem   ' アイテム情報
Dim lngMsgCnt   ' メッセージ数
Dim objOlFolderFilterResult ' フィルタ後のアイテム情報
Dim startDate   ' 検索開始日付

Dim tmpFolderPath
Const TMP_FOLDER_PATH = "_tmp"

ReDim targetEmailAddresses(3) ' 転送するメールアドレスの配列
Const targetEmailAddress1 = "zjc_apc-info@jp.ricoh.com" ' リコーアプリサポートセンター
Const targetEmailAddress2 = "zjc_kintone_support@jp.ricoh.com"  ' kintone保守
Const targetEmailAddress3 = "zjc_docuware_support@jp.ricoh.com"  ' docuware保守
Const targetEmailAddress4 = "kazunari_nakasone@jp.ricoh.com"  ' 仲宗根リコーアドレス

targetEmailAddresses(0) = targetEmailAddress1
targetEmailAddresses(1) = targetEmailAddress2
targetEmailAddresses(2) = targetEmailAddress3
targetEmailAddresses(3) = targetEmailAddress4

ReDim senderEmailAddresses(3) '転送先のメールアドレスの配列
Const senderEmailAddress1 = "rits_apc@pvcjp.com" ' リコーアプリサポートセンター
Const senderEmailAddress2 = "ricoh_kintone_support@pvcjp.com"  ' kintone保守
Const senderEmailAddress3 = "ricoh_docuware_support@pvcjp.com"  ' docuware保守
Const senderEmailAddress4 = "kazunari_nakasone@pvcjp.com"  '仲宗根PVCアドレス
'Const senderEmailAddress4 = "moritaka_nishihira@pvcjp.com"  'テスト用

senderEmailAddresses(0) = senderEmailAddress1
senderEmailAddresses(1) = senderEmailAddress2
senderEmailAddresses(2) = senderEmailAddress3
senderEmailAddresses(3) = senderEmailAddress4

Set objOlApp = CreateObject("Outlook.Application")
Set objShell = CreateObject("Wscript.Shell")


Call Proc()

' 説明：メイン処理
Private Function Proc()
    If Err.Number = 0 Then
        tmpFolderPath = objShell.CurrentDirectory & "\" & TMP_FOLDER_PATH
        
        ' 過去7日間のメールをチェックする。
        startDate = DateAdd("d", - 7, Date)
        Set objOlNs = objOlApp.GetNameSpace("MAPI")
        lngMsgCnt = 0
        

        Dim mailAddress
        For Each mailAddress In targetEmailAddresses
            Call DeleteAllTmpFiles(tmpFolderPath)
            
            If mailAddress = targetEmailAddress4 Then
                Set objOlFolder = objOlNs.GetDefaultFolder(6)
            Else
                Dim rec
                Set rec = objOlNs.CreateRecipient(mailAddress)
                Set objOlFolder = objOlNs.GetSharedDefaultFolder(rec,6)
            End If
            
            Set objOlFolderFilterResult = objOlFolder.Items.Restrict("[ReceivedTime] >= '" & startDate & "'")
            For Each objOlItem In objOlFolderFilterResult
                Call DeleteAllTmpFiles(tmpFolderPath)
                If objOlItem.UnRead = True Then
                    'メールを送信する
                    If SendMail(objOlApp, objOlItem, mailAddress) Then
                        objOlItem.UnRead = False
                        lngMsgCnt = lngMsgCnt + 1
                    Else
                        'イベントログに出力
                        objShell.LogEvent 1, "メールの送信に失敗しました"
                        Set objOlApp = Nothing
                        WScript.Quit
                    End If
                End If
            Next
        Next
    Else
        'イベントログに出力
        objShell.LogEvent 1, Err.Description
    End If
    
    Set objOlApp = Nothing
    Set objShell = Nothing
    
End Function

' 説明：メールを送信します
Private Function SendMail(objOlApp, objOlItem, mailAddress)
    Dim mailItem
    Set mailItem = CreateNewMailItem(objOlApp, objOlItem, mailAddress)
    mailItem.send
    'mailItem.display
    
    SendMail = (Err.Number = 0)
End Function

' 説明：メールアイテムを作成します
Private Function CreateNewMailItem(objOlApp, objOlItem, mailAddress)
    Dim newMailItem
    Set newMailItem = objOlApp.CreateItem(0)
    
    Dim toMailAddress
    Dim subject
    Dim jushinBi
    Dim body
    Dim attachmentFiles
    
    Call SaveTmpFiles(objOlItem, tmpFolderPath)
    
    body = "保守メールを転送します。内容の確認と対応をお願いします。" & vbCrLf
    body = body & "受信日：" & objOlItem.ReceivedTime & vbCrLf
    body = body & "--------------------------------------------------------" & vbCrLf
    body = body & objOlItem.Body & vbCrLf
    
    Select Case mailAddress
        Case targetEmailAddress1
        toMailAddress = senderEmailAddress1
        subject = "【APC保守 問合せ】" & objOlItem.Subject
        
        Case targetEmailAddress2
        toMailAddress = senderEmailAddress2
        subject = "【kintone保守 問合せ】" & objOlItem.Subject
        
        Case targetEmailAddress3
        toMailAddress = senderEmailAddress3
        subject = "【DocuWare保守 問合せ】" & objOlItem.Subject
        
        Case targetEmailAddress4
        toMailAddress = senderEmailAddress4
        subject = "【自動転送】" & objOlItem.Subject
        body = "このメッセージは自動転送されました。" & vbCrLf
        body = body & "受信日：" & objOlItem.ReceivedTime & vbCrLf
        body = body & "----------------------------------------------" & vbCrLf
        body = body & objOlItem.Body & vbCrLf
        
    End Select
    
    newMailItem.To = toMailAddress
    newMailItem.Subject = subject
    newMailItem.Body = body
    

    Set newMailItem = AttachedFileForMailItem(newMailItem, tmpFolderPath)
    
    Set CreateNewMailItem = newMailItem
End Function

'説明：一時フォルダにあるファイルを全削除します
Private Sub DeleteAllTmpFiles(path)
    Dim fObj
    Set fObj = CreateObject("Scripting.FileSystemObject")
    
    If fObj.FolderExists(path) Then
        Dim tmpFolder
        Set tmpFolder = fObj.GetFolder(path)
        
        Dim file
        For Each file In tmpFolder.files
            Call fObj.DeleteFile(path & "\" & file.Name, True)
        Next
    End If
End Sub

'説明：一時フォルダに添付ファイルをコピーします
Private Sub SaveTmpFiles(mailItem, path)
    '一時フォルダが無ければ作る
    Dim fObj
    Set fObj = CreateObject("Scripting.FileSystemObject")
    
    If fObj.FolderExists(path) = False Then
        Call fObj.CreateFolder(path)
    End If
    
    Dim attachmentFile
    For Each attachmentFile In mailItem.Attachments

        Call attachmentFile.SaveAsFile(path & "\" & attachmentFile)

    Next
End Sub

'説明：メールにファイルを添付します
Private Function AttachedFileForMailItem(newMailItem, path)
    Dim fObj
    Set fObj = CreateObject("Scripting.FileSystemObject")
    
    Dim tmpFolder
    Set tmpFolder = fObj.GetFolder(path)
    
    Dim file
    For Each file In tmpFolder.files
        Call newMailItem.Attachments.Add(path & "\" & file.name)
    Next
    
    Set AttachedFileForMailItem = newMailItem
End Function