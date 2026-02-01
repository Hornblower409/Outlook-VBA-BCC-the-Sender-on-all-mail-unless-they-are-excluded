' =====================================================================
' BCC Sender with exclusions
' Version: 2026-02-01_080436918
' https://github.com/Hornblower409/Outlook-VBA-BCC-Sender-with-exclusions
' RE: https://www.reddit.com/r/Outlook/comments/1qsa71g/mettre_automatiquement_en_bcc_cci_la_boite_mail/
' Must be in ThisOutlookSession because of the ItemSend event hook.
' General Help with Outlook VBA: https://www.slipstick.com/developer/how-to-use-outlooks-vba-editor/
'
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

    '   MAPI Property PR_SMTP_ADDRESS (PidTagSmtpAddress)
    '
    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

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

    Dim SenderSMTP As String
    SenderSMTP = AdrEntry.Address

    Dim ExchUser As Outlook.ExchangeUser
    If AdrEntry.Type = "EX" Then
        Set ExchUser = AdrEntry.GetExchangeUser
        SenderSMTP = ExchUser.PrimarySmtpAddress
    End If

    '   If Sender is Excluded - Remove the BCC & Done
    '
    Select Case LCase(SenderSMTP)

        '   SMTP Sender email addresses to exclude
        '   (Must be all lower case)
        '
        Case "smtpexclude01@domain.com", "smtpexclude02@domain.com"
            Mail.Recipients.Remove BCCRecipient.Index
            Exit Sub

        Case Else
            '   Continue

    End Select
    
    '   If Any Recipient is Excluded - Remove the BCC & Done
    '
    Dim Recipients As Outlook.Recipients
    Set Recipients = Mail.Recipients
    Dim Recipient As Outlook.Recipient
    Dim RecipientSMTP As String
    Dim PropAccess As Outlook.PropertyAccessor
    
    For Each Recipient In Recipients: Do
    
        '   Get the Recipient SMTP
        '   If it doesn't have one (e.g. Distribution List) - Skip it
        '
        Set PropAccess = Recipient.PropertyAccessor
        Dim ErrorNumber As Long
        On Error Resume Next
            RecipientSMTP = PropAccess.GetProperty(PR_SMTP_ADDRESS)
            ErrorNumber = Err.Number
        On Error GoTo 0
        If ErrorNumber <> 0 Then Exit Do ' Continue with Next Recipient
        
        Select Case LCase(RecipientSMTP)
        
            '   SMTP Recipient email addresses to exclude
            '   (Must be all lower case)
            '
            Case "recipientexclude01@domain.com", "recipientexclude02@domain.com"
                Mail.Recipients.Remove BCCRecipient.Index
                Exit Sub

            Case Else
                '   Continue
        
        End Select
    
    Loop While False: Next Recipient

End Sub
' =====================================================================
