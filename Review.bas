Sub FlagForReview()
    Dim itm As Object
    Set itm = Common.GetCurrentItem
    
    With itm
        .MarkAsTask olMarkTomorrow
        .FlagRequest = "Review"
        ' Beware: This will *replace* any existing categories on the selected MailItem
        .Categories = "Review"
        .Save
    End With
    
    Set itm = Nothing
End Sub

Sub Done()
    Dim itm As Object
    Set itm = Common.GetCurrentItem
    
    With itm
        .FlagStatus = olFlagComplete
        .Save
    End With
    
    Set itm = Nothing
End Sub
