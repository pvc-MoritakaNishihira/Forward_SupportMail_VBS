'******����******************************************************************************
'�@�ێ�p��RJ���L���[���A�h���X���̃��[�����APVC�ێ�p�̋��L���[���A�h���X�֓]�����܂�
'****************************************************************************************
Option Explicit
On Error Resume Next

Dim objShell ' shell �I�u�W�F�N�g
Dim objOlApp    ' outlook �I�u�W�F�N�g
Dim objOlNs     ' NameSpace �I�u�W�F�N�g
Dim objOlFolder ' �t�H���_���
Dim objOlItem   ' �A�C�e�����
Dim lngMsgCnt   ' ���b�Z�[�W��
Dim objOlFolderFilterResult ' �t�B���^��̃A�C�e�����
Dim startDate   ' �����J�n���t

Dim tmpFolderPath
Const TMP_FOLDER_PATH = "_tmp"

ReDim targetEmailAddresses(3) ' �]�����郁�[���A�h���X�̔z��
Const targetEmailAddress1 = "zjc_apc-info@jp.ricoh.com" ' ���R�[�A�v���T�|�[�g�Z���^�[
Const targetEmailAddress2 = "zjc_kintone_support@jp.ricoh.com"  ' kintone�ێ�
Const targetEmailAddress3 = "zjc_docuware_support@jp.ricoh.com"  ' docuware�ێ�
Const targetEmailAddress4 = "kazunari_nakasone@jp.ricoh.com"  ' ���@�����R�[�A�h���X

targetEmailAddresses(0) = targetEmailAddress1
targetEmailAddresses(1) = targetEmailAddress2
targetEmailAddresses(2) = targetEmailAddress3
targetEmailAddresses(3) = targetEmailAddress4

ReDim senderEmailAddresses(3) '�]����̃��[���A�h���X�̔z��
Const senderEmailAddress1 = "rits_apc@pvcjp.com" ' ���R�[�A�v���T�|�[�g�Z���^�[
Const senderEmailAddress2 = "ricoh_kintone_support@pvcjp.com"  ' kintone�ێ�
Const senderEmailAddress3 = "ricoh_docuware_support@pvcjp.com"  ' docuware�ێ�
Const senderEmailAddress4 = "kazunari_nakasone@pvcjp.com"  '���@��PVC�A�h���X
'Const senderEmailAddress4 = "moritaka_nishihira@pvcjp.com"  '�e�X�g�p

senderEmailAddresses(0) = senderEmailAddress1
senderEmailAddresses(1) = senderEmailAddress2
senderEmailAddresses(2) = senderEmailAddress3
senderEmailAddresses(3) = senderEmailAddress4

Set objOlApp = CreateObject("Outlook.Application")
Set objShell = CreateObject("Wscript.Shell")


Call Proc()

' �����F���C������
Private Function Proc()
    If Err.Number = 0 Then
        tmpFolderPath = objShell.CurrentDirectory & "\" & TMP_FOLDER_PATH
        
        ' �ߋ�7���Ԃ̃��[�����`�F�b�N����B
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
                    '���[���𑗐M����
                    If SendMail(objOlApp, objOlItem, mailAddress) Then
                        objOlItem.UnRead = False
                        lngMsgCnt = lngMsgCnt + 1
                    Else
                        '�C�x���g���O�ɏo��
                        objShell.LogEvent 1, "���[���̑��M�Ɏ��s���܂���"
                        Set objOlApp = Nothing
                        WScript.Quit
                    End If
                End If
            Next
        Next
    Else
        '�C�x���g���O�ɏo��
        objShell.LogEvent 1, Err.Description
    End If
    
    Set objOlApp = Nothing
    Set objShell = Nothing
    
End Function

' �����F���[���𑗐M���܂�
Private Function SendMail(objOlApp, objOlItem, mailAddress)
    Dim mailItem
    Set mailItem = CreateNewMailItem(objOlApp, objOlItem, mailAddress)
    mailItem.send
    'mailItem.display
    
    SendMail = (Err.Number = 0)
End Function

' �����F���[���A�C�e�����쐬���܂�
Private Function CreateNewMailItem(objOlApp, objOlItem, mailAddress)
    Dim newMailItem
    Set newMailItem = objOlApp.CreateItem(0)
    
    Dim toMailAddress
    Dim subject
    Dim jushinBi
    Dim body
    Dim attachmentFiles
    
    Call SaveTmpFiles(objOlItem, tmpFolderPath)
    
    body = "�ێ烁�[����]�����܂��B���e�̊m�F�ƑΉ������肢���܂��B" & vbCrLf
    body = body & "��M���F" & objOlItem.ReceivedTime & vbCrLf
    body = body & "--------------------------------------------------------" & vbCrLf
    body = body & objOlItem.Body & vbCrLf
    
    Select Case mailAddress
        Case targetEmailAddress1
        toMailAddress = senderEmailAddress1
        subject = "�yAPC�ێ� �⍇���z" & objOlItem.Subject
        
        Case targetEmailAddress2
        toMailAddress = senderEmailAddress2
        subject = "�ykintone�ێ� �⍇���z" & objOlItem.Subject
        
        Case targetEmailAddress3
        toMailAddress = senderEmailAddress3
        subject = "�yDocuWare�ێ� �⍇���z" & objOlItem.Subject
        
        Case targetEmailAddress4
        toMailAddress = senderEmailAddress4
        subject = "�y�����]���z" & objOlItem.Subject
        body = "���̃��b�Z�[�W�͎����]������܂����B" & vbCrLf
        body = body & "��M���F" & objOlItem.ReceivedTime & vbCrLf
        body = body & "----------------------------------------------" & vbCrLf
        body = body & objOlItem.Body & vbCrLf
        
    End Select
    
    newMailItem.To = toMailAddress
    newMailItem.Subject = subject
    newMailItem.Body = body
    

    Set newMailItem = AttachedFileForMailItem(newMailItem, tmpFolderPath)
    
    Set CreateNewMailItem = newMailItem
End Function

'�����F�ꎞ�t�H���_�ɂ���t�@�C����S�폜���܂�
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

'�����F�ꎞ�t�H���_�ɓY�t�t�@�C�����R�s�[���܂�
Private Sub SaveTmpFiles(mailItem, path)
    '�ꎞ�t�H���_��������΍��
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

'�����F���[���Ƀt�@�C����Y�t���܂�
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