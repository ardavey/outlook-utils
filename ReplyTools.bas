Sub ReplyAll()
    Dim rpl As Outlook.MailItem
    Dim itm As Object
     
    Set itm = Common.GetCurrentItem
    Set rpl = itm.ReplyAll
    rpl.Display
     
    Set rpl = Nothing
    Set itm = Nothing
End Sub

Sub ThankYou()
    Dim itm As Object
    Dim rpl As Outlook.MailItem
    
    Set itm = Common.GetCurrentItem
    Set rpl = itm.Reply
    
    With rpl
        .HTMLBody = "<p class=MsoNormal>That's great, thanks!</p>" + rpl.HTMLBody
        .Display
    End With
    
    Set itm = Nothing
    Set rpl = Nothing
End Sub
