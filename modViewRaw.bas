Option Explicit

Public Sub ViewRawHeaders()
  
  Dim sel As Selection, msg As MailItem, newMessage As MailItem ', i As Long
  
  Set sel = Outlook.Application.ActiveExplorer.Selection
  
  If sel.Count = 0 Then
    'MsgBox "You must select at least 1 message first!", vbExclamation
    Exit Sub
  End If
  
  For Each msg In sel
    Set newMessage = Outlook.Application.CreateItem(olMailItem)
    newMessage.body = GetEmailHeaders(msg)
    'If msg.Attachments.Count Then
      'For i = 1 To msg.Attachments.Count
        'newMessage.Attachments.Add msg.Attachments.Item(i).FileName, olByValue, 1, msg.Attachments.Item(i).DisplayName
      'Next
    'End If
    If LenB(newMessage.body) Then newMessage.Display
  Next
  
  Set sel = Nothing
  
End Sub

Public Sub ViewRawMessage()
  
  Dim sel As Selection, msg As MailItem, newMessage As MailItem ', i As Long
  
  Set sel = Outlook.Application.ActiveExplorer.Selection
  
  If sel.Count = 0 Then
    'MsgBox "You must select at least 1 message first!", vbExclamation
    Exit Sub
  End If
  
  For Each msg In sel
    Set newMessage = Outlook.Application.CreateItem(olMailItem)
    newMessage.body = GetEmailHeaders(msg) & vbCrLf & vbCrLf & msg.HTMLBody
    'If msg.Attachments.Count Then
      'For i = 1 To msg.Attachments.Count
        'newMessage.Attachments.Add msg.Attachments.Item(i).FileName, olByValue, 1, msg.Attachments.Item(i).DisplayName
      'Next
    'End If
    If LenB(newMessage.body) Then newMessage.Display
  Next
  
  Set sel = Nothing
  
End Sub

Public Sub ForwardSpamToAuthorities()
  
  Dim sel As Selection, msg As MailItem, newMessage As Object, i As Long, subject As String
  Dim ns As Outlook.Namespace, junkFolder As Folder

  Set sel = Outlook.Application.ActiveExplorer.Selection
  Set ns = Application.GetNamespace("MAPI")
  Set junkFolder = ns.GetDefaultFolder(olFolderJunk)
  Set ns = Nothing

  If sel.Count = 0 Then
    'MsgBox "You must select at least 1 message first!", vbExclamation
    Exit Sub
  End If
  
  For Each msg In sel
    
    Dim sender As String, authority As String
        
    sender = GetSenderFromMail(msg)
    
    If LenB(sender) Then
      authority = LookupAppropriateAuthority(sender)
    End If
    
    Set newMessage = Outlook.Application.CreateItem(olMailItem)
        
    newMessage.subject = "FW: " & msg.subject
    newMessage.Attachments.Add msg
    newMessage.body = vbNullString
    newMessage.HTMLBody = vbNullString
            
    Set newMessage.SendUsingAccount = msg.SendUsingAccount
    
    If LenB(authority) And Left$(authority, 1) <> "<" Then
      newMessage.Recipients.Add authority
      newMessage.Recipients.ResolveAll
      newMessage.Send
    Else
      If LenB(sender) Then
        newMessage.Recipients.Add "Spammer: " & sender & IIf(Left$(authority, 1) = "<", GetString(authority, "<", ">"), vbNullString)
      End If
      newMessage.Display
    End If
                
    msg.UnRead = False
    
    If msg.Parent <> junkFolder Then
      msg.Move junkFolder
    End If
    
  Next
  
  Set sel = Nothing

End Sub

Private Function GetEmailHeaders(ByVal item As MailItem) As String

  Dim oPA As Outlook.PropertyAccessor
  
  Set oPA = item.PropertyAccessor
  'PR_TRANSPORT_MESSAGE_HEADERS_W
  GetEmailHeaders = oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")

End Function

Private Function GetAuthorities() As Collection

  Dim col As New Collection
  
  col.Add "com,org,net,edu,gov,mil spam@knujon.net"
  col.Add "com.au,org.au,net.au,edu.au,gov.au,mil.au report@submit.spam.acma.gov.au"
  col.Add "dk int@spamklage.dk"
  
  col.Add "fr <Please create an account at http://signalspam.net then submit the spam through the web form>"
  
  Set GetAuthorities = col

End Function

Private Function LookupAppropriateAuthority(ByVal address As String) As String

  Dim pos As Long, result As String
  Dim authorities As Collection
  Dim i As Long, j As Long
  
  pos = InStr(address, "@")
  
  If pos Then
    address = Mid$(address, pos + 1)
  End If
  
  Set authorities = GetAuthorities
  
  For i = 1 To authorities.Count
  
    Dim arr() As String, arr2() As String, authority As String, domains As String
    
    arr = Split(authorities(i), " ")
    domains = arr(0)
    authority = arr(1)
    
    If InStr(domains, ",") Then
      arr2 = Split(domains, ",")
    Else
      ReDim arr2(0) As String
      arr2(0) = domains
    End If
    
    For j = 0 To UBound(arr2)
      If LCase$(Right$(address, Len(arr2(j)))) = LCase$(arr2(j)) Then
        result = authority
        Exit For
      End If
    Next
  
    If LenB(result) Then
      Exit For
    End If
    
  Next
  
  LookupAppropriateAuthority = result
    
End Function

Private Function GetSenderFromMail(ByVal item As MailItem) As String

  Dim oPA As Outlook.PropertyAccessor
  Set oPA = item.PropertyAccessor
  
  'PidTagOriginalMessageId_W
  GetSenderFromMail = oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1046001F")

End Function

Private Function GetString(ByVal strData As String, ByVal Search1 As String, ByVal Search2 As String, Optional ByVal boolDontCheckCase As Boolean = True) As String
  
  Dim pos(1) As Long
  
  If boolDontCheckCase Then
    pos(0) = InStr(strData, Search1)
    pos(1) = InStr(pos(0) + Len(Search1), strData, Search2)
  Else
    pos(0) = InStr(LCase$(strData), Search1)
    pos(1) = InStr(pos(0) + Len(Search1), LCase$(strData), Search2)
  End If
  
  If pos(0) > 0 And pos(1) > pos(0) + Len(Search1) Then GetString = Mid$(strData, pos(0) + Len(Search1), (pos(1) - pos(0)) - Len(Search1))
  
End Function
