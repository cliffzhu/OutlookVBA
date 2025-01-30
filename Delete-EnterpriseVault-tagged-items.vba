' If you imported the PST files from EnterpriseVault and they are labeled as Archived and attachment is a link, you want to clean it up.
Dim totalCount As Long
Sub FindIPMNoteEmailsInArchive()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olArchive As Outlook.MAPIFolder
    Dim subFolder As Outlook.MAPIFolder
    
    ' Initialize Outlook application and namespace
    Set olApp = New Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    
    ' Access the Online Archive (change "YourEmail@example.com" to your actual email)
    Set olArchive = olNS.Folders("Online Archive - YourEmail@example.com").Folders("Sent Items")
    '.Folders("Deleted Items")
    '.Folders("Deleted Items")
        
    ' Access the Drafts folder
    Set olDrafts = olNS.GetDefaultFolder(olFolderDrafts)
    ' set start time
    Dim startTime As Date
    startTime = Now
    ' define total count

    totalCount = 0

   ' Search through all folders in the Online Archive
    
    Call SearchFolder(olArchive, olDrafts)
    ' get end time
    Dim endTime As Date
    ' Calculate the time taken to process the emails
    endTime = Now
    Debug.Print "Time taken: " & Format(endTime - startTime, "hh:mm:ss")
    Debug.Print "Total emails processed: " & totalCount
    ' Clean up objects
    Set olArchive = Nothing
    Set olNS = Nothing
    Set olApp = Nothing
End Sub

Sub SearchFolder(ByVal folder As Outlook.MAPIFolder, ByVal draftsFolder As Outlook.MAPIFolder)
    Dim olItems As Outlook.Items
    Dim olMail As Object ' Use Object to accommodate different item types
    Dim subFolder As Outlook.MAPIFolder
    Dim i As Long
        
    Debug.Print "Processing: " & folder.Name ' Print the name of the subfolder being processed
    DoEvents ' Allow other events to process
    
    ' Get items from the current folder
    Set olItems = folder.Items
    Dim folderitemcount As Long
    Debug.Print "Item count: " & olItems.Count
    ' Loop through items in the current folder
    For i = olItems.Count To 1 Step -1
        Set olMail = olItems(i)
        
        ' Check if the item is a MailItem and has a message class that starts with "IPM.Note"
        If TypeOf olMail Is Outlook.MailItem Then
            If InStr(1, olMail.MessageClass, "IPM.Note.EnterpriseVault") > 0 Then
                Debug.Print "Subject: " & olMail.Subject
                Debug.Print "Sender: " & olMail.SenderName
                Debug.Print "Received Time: " & olMail.ReceivedTime
                'Debug.Print "-----------------------------------"
                'DoEvents ' Allow other events to process
                'olMail.MessageClass = "IPM.Note"
                'olMail.Save ' Save changes to the item
                'Debug.Print "Changed Message Class to: IPM.Note"
                'Debug.Print "-----------------------------------"


                                ' Make a copy of the email and move it to Drafts folder
                'Dim copiedMail As Outlook.MailItem
                'Set copiedMail = olMail.Copy
                
                ' Move the copied email to Drafts folder
                'copiedMail.Move draftsFolder
                
                'Debug.Print "Copied email moved to Drafts."
                Debug.Print "-----------------------------------"
                DoEvents ' Allow other events to process
                'delete the original email
                olMail.Delete
                totalCount = totalCount + 1

            End If
        End If

        If (i Mod 300) = 0 Then
        Debug.Print "Processed " & i & " emails"
        DoEvents
        End If
    Next i
    
    ' Loop through each subfolder in the current folder and search recursively
    For Each subFolder In folder.Folders
       
        Call SearchFolder(subFolder, draftsFolder) ' Recursively search the subfolder
    Next subFolder
    
End Sub


