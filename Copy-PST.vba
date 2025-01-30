' Description: This script copies all items from the source PST files in a folder to the destination PST files. 
' why? The Enterprise Vault exports PST and has all item marked as archived and the attachments are converted as a link. 
' This script to make the item as a regular item again so they can be imported to Exchange online.
Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private itemsProcessed As Long
Private totalItems As Long




Sub CopyPSTsInFolder()
    Dim SourcePSTFolderPath As String
    Dim DestinationPSTFolderPath As String
    Dim FileSystem As Object
    Dim PSTFile As Object
    Dim olNamespace As Outlook.NameSpace
    Dim SourceDisplayName As String
    Dim SourceStore As Outlook.Folder
    Dim DestinationPSTPath As String
    Dim DestinationStore As Outlook.Folder

    ' Define source and destination paths
    SourcePSTFolderPath = "H:\Restored" ' Update with your source PST folder path
    DestinationPSTFolderPath = "E:\PST" ' Update with your destination PST folder path

    ' Initialize
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Set olNamespace = Application.GetNamespace("MAPI")

    ' Loop through PST files in the source folder
    For Each PSTFile In FileSystem.GetFolder(SourcePSTFolderPath).Files
        If LCase(FileSystem.GetExtensionName(PSTFile.Name)) = "pst" Then
            Debug.Print "Processing PST: " & PSTFile.Path

            ' Open the source PST in Outlook
            On Error Resume Next
            olNamespace.AddStore PSTFile.Path
            On Error GoTo 0

            ' Retrieve the display name of the PST
            SourceDisplayName = GetPSTDisplayNameByFilePath(olNamespace, PSTFile.Path)
            If SourceDisplayName = "" Then
                MsgBox "Could not open source PST: " & PSTFile.Path, vbExclamation
                GoTo ClosePST
            End If

            Debug.Print "Source Display Name: " & SourceDisplayName

            ' Create the destination PST path (file name matches source PST file)
            DestinationPSTPath = DestinationPSTFolderPath & "\" & SourceDisplayName & "-" & PSTFile.Name

            ' Create and open the destination PST
            On Error Resume Next
            olNamespace.AddStore DestinationPSTPath
            On Error GoTo 0

            ' Fix the display name of the newly created PST
            Set DestinationStore = GetPSTByFilePath(olNamespace, DestinationPSTPath)
            If DestinationStore Is Nothing Then
                MsgBox "Could not detect destination PST: " & DestinationPSTPath, vbExclamation
                GoTo ClosePST
            End If

            ' Rename the root folder of the destination PST
            On Error Resume Next
            DestinationStore.Name = SourceDisplayName & " - copy"
            On Error GoTo 0

            Debug.Print "Destination PST Display Name Set: " & DestinationStore.Name

            ' Get the root folder of the source PST
            Set SourceStore = olNamespace.Folders(SourceDisplayName)
            If SourceStore Is Nothing Then
                MsgBox "Source PST root folder not found: " & SourceDisplayName, vbExclamation
                GoTo ClosePST
            End If

            ' Initialize counters
            itemsProcessed = 0
            totalItems = 0


            ' Copy folders and items from source to destination
            CreateFoldersRecursive SourceStore, DestinationStore
            CopyFoldersOverwrite SourceStore, DestinationStore, totalItems

            Debug.Print "Copying complete! Total items copied: " & totalItems
            Debug.Print "Complete! Items copied from " & SourceDisplayName & ": " & totalItems & ": " & PSTFile.Path
            DoEvents ' Allow Outlook to process pending events

ClosePST:
            ' Close both PST files
            On Error Resume Next
            olNamespace.RemoveStore SourceStore
            olNamespace.RemoveStore DestinationStore
            On Error GoTo 0
        End If
    Next PSTFile

    Debug.Print "All PSTs processed!"
End Sub

Private Function GetPSTDisplayNameByFilePath(ByVal olNamespace As Outlook.NameSpace, ByVal PSTFilePath As String) As String
    Dim Store As Outlook.Store
    PSTFilePath = LCase(PSTFilePath) ' Normalize for comparison

    For Each Store In olNamespace.Stores
        On Error Resume Next
        If LCase(Store.FilePath) = PSTFilePath Then
            GetPSTDisplayNameByFilePath = Store.DisplayName
            Exit Function
        End If
        On Error GoTo 0
    Next Store

    GetPSTDisplayNameByFilePath = ""
End Function

Private Function GetPSTByFilePath(ByVal olNamespace As Outlook.NameSpace, ByVal PSTFilePath As String) As Outlook.Folder
    Dim Store As Outlook.Store
    PSTFilePath = LCase(PSTFilePath) ' Normalize for comparison

    For Each Store In olNamespace.Stores
        On Error Resume Next
        If LCase(Store.FilePath) = PSTFilePath Then
            Set GetPSTByFilePath = Store.GetRootFolder
            Exit Function
        End If
        On Error GoTo 0
    Next Store

    Set GetPSTByFilePath = Nothing
End Function

Private Sub CountTotalItems(ByVal Folder As Outlook.Folder)
    Dim SubFolder As Outlook.Folder
    totalItems = totalItems + Folder.Items.Count
    For Each SubFolder In Folder.Folders
        CountTotalItems SubFolder
    Next
End Sub
Private Sub CreateFoldersRecursive(ByVal SourceFolder As Outlook.Folder, ByVal DestinationFolder As Outlook.Folder)
    Dim SubFolder As Outlook.Folder
    Dim NewFolder As Outlook.Folder
    Dim FolderName As String
    
    ' Loop through each subfolder in the source folder
    For Each SubFolder In SourceFolder.Folders
        FolderName = SubFolder.Name
        On Error Resume Next
        
        ' Check if the folder already exists in the destination PST
        Set NewFolder = Nothing
        Set NewFolder = DestinationFolder.Folders(FolderName)
        
        ' If the folder does not exist, create it
        If NewFolder Is Nothing Then
            Set NewFolder = DestinationFolder.Folders.Add(FolderName)
            If Not NewFolder Is Nothing Then
                Debug.Print "Created folder: " & FolderName & " in " & DestinationFolder.Name
            Else
                Debug.Print "Failed to create folder: " & FolderName & " in " & DestinationFolder.Name
            End If
        Else
            Debug.Print "Folder exists: " & FolderName & " in " & DestinationFolder.Name
        End If
        On Error GoTo 0
        
        ' Ensure the folder is valid before recursion
        If Not NewFolder Is Nothing Then
            CreateFoldersRecursive SubFolder, NewFolder
        Else
            Debug.Print "Skipping recursion for folder: " & FolderName
        End If
    Next SubFolder
End Sub

Private Sub CopyFoldersOverwrite(ByVal SourceFolder As Outlook.Folder, ByVal DestinationFolder As Outlook.Folder, ByRef TotalItemsCopied As Long)
    Dim SubFolder As Outlook.Folder
    Dim NewFolder As Outlook.Folder
    Dim Item As Object
    Dim CopiedItem As Object
    Dim i As Long

    ' Debug output for the current folder with the total items copied so far
    Debug.Print "Processing folder: " & SourceFolder.Name & " | Total items copied so far: " & TotalItemsCopied
    DoEvents ' Allow Outlook to process other events and stay responsive

    ' Copy items in the current folder
    For i = 1 To SourceFolder.Items.Count
        On Error Resume Next
        Set Item = SourceFolder.Items(i)
        If Not Item Is Nothing Then
            Set CopiedItem = Item.Copy
            If Not CopiedItem Is Nothing Then
                CopiedItem.Move DestinationFolder
                TotalItemsCopied = TotalItemsCopied + 1
            End If
        End If
        'display progress on every 100 items
        If i Mod 100 = 0 Then
            Debug.Print "Items copied for current folder: " & i
            DoEvents ' Allow Outlook to process other events and stay responsive
        End If
        On Error GoTo 0
    Next i

    ' Process subfolders recursively
    For Each SubFolder In SourceFolder.Folders
        On Error Resume Next
        Set NewFolder = DestinationFolder.Folders(SubFolder.Name)
        If NewFolder Is Nothing Then
            Set NewFolder = DestinationFolder.Folders.Add(SubFolder.Name)
            Debug.Print "Created folder: " & SubFolder.Name
            DoEvents ' Allow Outlook to process other events and stay responsive
        End If
        On Error GoTo 0

        If Not NewFolder Is Nothing Then
            CopyFoldersOverwrite SubFolder, NewFolder, TotalItemsCopied
        End If
    Next SubFolder
End Sub

