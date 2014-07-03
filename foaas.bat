Private Sub myOlItems_ItemAdd(ByVal Item As Object)

End Sub

Sub AutoResponse(objmsg As Outlook.MailItem)

    ' define my reply message
    Dim objReply As MailItem
    ' let's get ourselves the inbox!
    Dim inbox As MAPIFolder
    Set inbox = Application.GetNamespace("MAPI"). _
    GetDefaultFolder(olFolderInbox)

    ' Let's get this reply going!
    Set objReply = objmsg.Reply
    ' Subject Re: their subject. Standard
    objReply.Subject = "Re: " & objReply.Subject
    ' Body - you define this, use the variable for the unread count in inbox
    
    Dim teststring As String
    teststring = Replace(objmsg.SenderName, " ", "%20")
    objReply.Body = "I am currently unavailable but you may find the information you need here www.foaas.com/off/" & teststring & "<MYNAME>"

    ' Send this thing!
    objReply.Send
    ' Reset
    Set objReply = Nothing

End Sub
