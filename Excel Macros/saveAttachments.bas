Attribute VB_Name = "saveAttachments"
Public Sub saveAttachments()
'Update 20191101
    Dim objOL As Object 'Outlook.Application
    Dim objMsg As Outlook.MailItem
    Dim objAttachments As Outlook.Attachments
    Dim objSelection As Outlook.Selection
    Dim i As Long
    Dim lngCount As Long
    Dim strFile As String
    Dim strFolderpath As String
    'Dim strDeletedFiles As String
    strFolderpath = CreateObject("WScript.Shell").SpecialFolders(16)
    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    strFolderpath = strFolderpath & "\Attachments\"
    Dim fileNamePrepend As String
    
    For Each objMsg In objSelection
        fileNamePrepend = Replace(Replace(Replace(objMsg.Subject, "Change Assignment: ", ""), "/", "_"), ":", "") & " - "
        Set objAttachments = objMsg.Attachments
        lngCount = objAttachments.Count
        'strDeletedFiles = ""
        If lngCount > 0 Then
            For i = lngCount To 1 Step -1
                strFile = strFolderpath & fileNamePrepend & objAttachments.Item(i).FileName
                Debug.Print strFile
                objAttachments.Item(i).SaveAsFile strFile
                'objAttachments.Item(i).Delete()
                'If objMsg.BodyFormat <> olFormatHTML Then
                '    strDeletedFiles = strDeletedFiles & vbCrLf & "<Error! Hyperlink reference not valid.>"
                'Else
                '    strDeletedFiles = strDeletedFiles & "<br>" & "<a href='file://" & _
                '    strFile & "'>" & strFile & "</a>"
                'End If
            Next i
            'If objMsg.BodyFormat <> olFormatHTML Then
            '    objMsg.Body = vbCrLf & "The file(s) were saved to " & strDeletedFiles & vbCrLf & objMsg.Body
            'Else
            '    objMsg.HTMLBody = "<p>" & "The file(s) were saved to " & strDeletedFiles & "</p>" & objMsg.HTMLBody
            'End If
            'objMsg.Save
        End If
    Next
ExitSub:
        Set objAttachments = Nothing
        Set objMsg = Nothing
        Set objSelection = Nothing
        Set objOL = Nothing
End Sub

