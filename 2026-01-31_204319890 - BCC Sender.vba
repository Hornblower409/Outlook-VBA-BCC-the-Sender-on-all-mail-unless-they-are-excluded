' =====================================================================
' BCC the Sender on all mail unless they are excluded.
' Version: 2026-01-31_200432473
' RE: https://www.reddit.com/r/Outlook/comments/1qsa71g/mettre_automatiquement_en_bcc_cci_la_boite_mail/
' Must be in ThisOutlookSession because of the ItemSend event hook.
' General Help with Outlook VBA: https://www.slipstick.com/developer/how-to-use-outlooks-vba-editor/
'
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

    '   If Not a MailItem - done
    '
    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub

    '   Cast the Item object to a MailItem
    '
    Dim Mail As Outlook.MailItem
    Set Mail = Item
    
    '   Add the Sender as a BCC
    '
    Dim BCCRecipient As Outlook.Recipient
    Set BCCRecipient = Mail.Recipients.Add(Mail.SenderEmailAddress)
    BCCRecipient.Type = Outlook.OlMailRecipientType.olBCC

    '   Make sure the Sender address resolves
    '
    BCCRecipient.Resolve
    If Not BCCRecipient.Resolved Then
        MsgBox "BCCRecipient did not Resolve." & vbCrLf & vbCrLf & BCCRecipient.Address
        Mail.Recipients.Remove BCCRecipient.Index
        Cancel = True
        Exit Sub
    End If
    
    '   Get the Sender's SMTP Address
    '
    Dim AdrEntry As Outlook.AddressEntry
    Set AdrEntry = BCCRecipient.AddressEntry
    
    Dim SMTPAdr As String
    SMTPAdr = AdrEntry.Address
    
    Dim ExchUser As Outlook.ExchangeUser
    If AdrEntry.Type = "EX" Then
        Set ExchUser = AdrEntry.GetExchangeUser
        SMTPAdr = ExchUser.PrimarySmtpAddress
    End If
    
    SMTPAdr = LCase(SMTPAdr)
    
    '   If we do not want to BCC this SMTP - Remove it
    '
    Select Case SMTPAdr
    
        '   SMTP email addresses to be excluded must be all lower case
        '
        Case "smtpexclude01@domain.com", "smtpexclude02@domain.com"
            Mail.Recipients.Remove BCCRecipient.Index
            
        Case Else
            '   Continue
            
    End Select
    
End Sub
' =====================================================================
