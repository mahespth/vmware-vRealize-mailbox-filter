
'
' Scan mailbox/folder for a particular message and move
' them to another folder if you find a another message
' that indicates the first message is closed.
' Written to handle vrealize messages but works with anything has
' a unique id in the message body.
'
' __author__     = "Steve Maher"
' __copyright__  = ""
' __credits__    = ["Steve Maher"]
' __license__    = ""
' __version__    = "1.0.1"
' __maintainer__ = "Steve Maher"
' __email__      = ""
' __status__     = ""
'
' Ref: https://docs.microsoft.com/en-us/office/vba/outlook/how-to/items-folders-and-stores/working-with-entryids-and-storeids
' Ref: https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem.entryid
'

Private Sub Application_NewMail()
    Call consolidate_vRealize_messages
End Sub


Sub consolidate_vRealize_messages()

Dim myOlApp As New Outlook.Application
Dim myNameSpace As Outlook.NameSpace
Dim myInbox As Outlook.Folder
Dim myitems As Outlook.Items
Dim myitem As Object

Dim countedSubject As Integer
Dim countedCleanup As Integer

Dim StoreID As String

Dim subjectRequired As String
Dim subjectClosing As String

Dim messageIdentifier As String
Dim messageIdentifierFieldId As String

closedMailBox = "vRealize Closed"
liveMailBox = "vRealize Messages"

Dim cancelledMessages
Set cancelledMessages = CreateObject("Scripting.Dictionary")

Dim newAlertMessages
Set newAlertMessages = CreateObject("Scripting.Dictionary")

Set myNameSpace = myOlApp.GetNamespace("MAPI")
Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)

StoreID = myInbox.StoreID

Set myDestFolder = myInbox.Folders(closedMailBox)

' ------------ use this to scan a named source folder ----------------
If Len(liveMailBox) > 1 Then

    Set srcFolder = myInbox.Folders(liveMailBox)
    Set myitems = srcFolder.Items
Else
' --------------- or this to scan all folders -----------------------
    Set myitems = myInbox.Items
End If

' --------------- subject line "contains" this text -----------------
subjectRequired = "[vRealize Operations Manager]"
subjectClosing = "[vRealize Operations Manager] cancelled alert"

' --------------- look for this in the body -------------------------
messageIdentifier = "Alert ID :"

' -- use this field as the key from the lines matched as per above --
messageIdentifierFieldId = 3

' --------------- Identify all closed messages ----------------------
For Each myitem In myitems
    If myitem.Class = olMail Then
        If InStr(1, myitem.Subject, subjectRequired, vbTextCompare) > 0 Then
            
            countedSubject = countedSubject + 1
            
            sText = myitem.Body
            vText = Split(sText, Chr(13))
            
            For i = UBound(vText) To 0 Step -1
                
                If InStr(1, vText(i), messageIdentifier) > 0 Then
                    alertID = Split(vText(i))
                    
                    If Not newAlertMessages.Exists(alertID(messageIdentifierFieldId)) Then
                        newAlertMessages.Add alertID(messageIdentifierFieldId), myitem.EntryID
                    Else
                        existingEntryIDs = newAlertMessages(alertID(messageIdentifierFieldId))
                        newAlertMessages.Remove (alertID(messageIdentifierFieldId))
                        
                        newAlertMessages.Add alertID(messageIdentifierFieldId), existingEntryIDs & " " & myitem.EntryID
                    End If
                    
                    If InStr(1, myitem.Subject, subjectClosing, vbTextCompare) > 0 Then
                        If Not cancelledMessages.Exists(alertID(messageIdentifierFieldId)) Then
                            countedCleanup = countedCleanup + 1
                            cancelledMessages.Add alertID(messageIdentifierFieldId), myitem.EntryID
                        End If
                    End If
                    
                    Exit For
                End If
            Next
        End If
    End If
Next myitem
' ------------------------------------------------------------------


' ------------- this is the code that moves the messages -----------
For Each EntryID In cancelledMessages
    Set Item = myNameSpace.GetItemFromID(cancelledMessages(EntryID), StoreID)
    Item.Move (myDestFolder)
    
    If newAlertMessages.Exists(EntryID) Then
        On Error Resume Next
        For Each mailId In Split(newAlertMessages(EntryID))
            Set Item = myNameSpace.GetItemFromID(mailId, StoreID)
            Item.Move (myDestFolder)
        Next mailId
    End If
Next EntryID
' ------------------------------------------------------------------

' --------------------- allow the object to be free'd --------------
Set myOlApp = Nothing

End Sub
