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
    Dim replyStrings()
    Dim idx As Integer
    
    Set itm = Common.GetCurrentItem
    Set rpl = itm.Reply
    ' Extend this with your own selection of thankful replies - one will be selected at random
    replyStrings = Array( _
        "That's great, thank you!", _
        "Nice one, thanks!", _
        "Excellent, thanks for that!", _
        "Thanks for that, much appreciated!" _
    )
    
    Randomize
    idx = Int(Rnd() * (UBound(replyStrings) - LBound(replyStrings) + 1))
    
    ' Squirt our stock reply in where the cursor would appear in the editor
    Set objDoc = rpl.GetInspector.WordEditor
    objDoc.Characters(1).InsertBefore replyStrings(idx)
    
    ' Oddly, this currently results in an empty email when calling rpl.Send ...
    ' rpl.Display will open the reply window correctly and the user can hit the Send button
    rpl.Display
    
    Set itm = Nothing
    Set rpl = Nothing
    
End Sub
