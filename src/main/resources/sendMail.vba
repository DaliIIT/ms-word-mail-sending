Sub SendMail()
 Options.SendMailAttach = True
    ActiveWindow.EnvelopeVisible = True
    Set oItem = ActiveDocument.MailEnvelope.Item
    With oItem
        .Subject = "Testing 123"
        .Recipients.Add "medalijallouli@gmail.com"
        .Send
    End With
    Set oItem = Nothing
    ActiveWindow.EnvelopeVisible = False
End Sub