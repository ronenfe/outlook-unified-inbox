# Outlook Unified Inbox (VBA)

A VBA macro for Outlook Classic that syncs emails from all account inboxes into a single unified folder in real time.

## How It Works

* On startup, watches every account's inbox using `WithEvents` — new mail is copied to the unified folder instantly
* A periodic timer sweeps all inboxes every 5 minutes to catch anything missed
* On first run, prompts you to pick the target folder; saves the selection to the Windows registry
* Checks the 100 most recent emails per inbox on each sweep
* Snapshots inbox EntryIDs before processing — never mutates a live collection mid-loop
* Writes a timestamped log file to `%USERPROFILE%\\UnifiedInbox.log` for troubleshooting

## Files

|File|Type|Purpose|
|-|-|-|
|`ThisOutlookSession`|Built-in|Startup/quit hooks|
|`modUnifiedSync`|Module|Timer, sync logic, folder management, logging|
|`clsInboxEvents`|Class Module|Per-inbox `WithEvents` watcher|

## Installation

1. Open Outlook, press **Alt+F11** to open the VBA Editor
2. In the Project Explorer (left panel), expand **Project (VbaProject.OTM)**
3. **Add the module:** Insert → Module → rename to `modUnifiedSync` → paste contents below
4. **Add the class:** Insert → Class Module → press F4, set `Name` to `clsInboxEvents` → paste contents below
5. **Edit ThisOutlookSession:** double-click it → replace all contents with the code below
6. Save (**Ctrl+S**), close and reopen Outlook
7. On first launch you will be prompted to pick your unified folder — select it once and it is remembered

## Code

### `ThisOutlookSession`

```vb
Option Explicit

Private Sub Application\_Startup()
    LogMsg "=== Outlook started ==="
    LogMsg "Session stores count: " \& Application.Session.Stores.Count
    InitWatchers
    StartTimer
    RunUnifiedSync
    LogMsg "Startup complete"
End Sub

Private Sub Application\_Quit()
    LogMsg "=== Outlook quitting ==="
    StopTimer
End Sub
```

### `clsInboxEvents`

```vb
Option Explicit

Public WithEvents Items As Outlook.Items

Private Sub Items\_ItemAdd(ByVal Item As Object)
    If TypeOf Item Is Outlook.MailItem Then
        Dim unified As Outlook.Folder
        Set unified = GetUnifiedFolder()
        If unified Is Nothing Then Exit Sub

        On Error Resume Next
        SyncNewItem Item, unified
        On Error GoTo 0
    End If
End Sub
```

### `modUnifiedSync`

```vb
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private timerID As LongPtr
#Else
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Private timerID As Long
#End If

Private watchers As New Collection
Public isSyncRunning As Boolean
Public lastSync As Date

' ---------------------------------------------------------------------------
' Logging
' ---------------------------------------------------------------------------

Public Sub LogMsg(msg As String)
    Dim logPath As String
    logPath = Environ("USERPROFILE") \& "\\UnifiedInbox.log"

    Dim fileNum As Integer
    fileNum = FreeFile
    Open logPath For Append As #fileNum
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") \& " | " \& msg
    Close #fileNum
End Sub

' ---------------------------------------------------------------------------
' Watchers
' ---------------------------------------------------------------------------

Public Sub InitWatchers()
    Dim ns As Outlook.NameSpace
    Dim store As Outlook.Store
    Dim inbox As Outlook.Folder
    Dim w As clsInboxEvents

    Set ns = Application.Session
    LogMsg "InitWatchers: " \& ns.Stores.Count \& " store(s) found"

    For Each store In ns.Stores
        Set inbox = Nothing
        On Error Resume Next
        Set inbox = store.GetDefaultFolder(olFolderInbox)
        On Error GoTo 0

        If Not inbox Is Nothing Then
            LogMsg "  Watching inbox: " \& store.DisplayName
            Set w = New clsInboxEvents
            Set w.Items = inbox.Items
            watchers.Add w
        Else
            LogMsg "  Skipped (no inbox): " \& store.DisplayName
        End If
    Next store

    LogMsg "InitWatchers complete, watchers: " \& watchers.Count
End Sub

' ---------------------------------------------------------------------------
' Timer
' ---------------------------------------------------------------------------

Public Sub StartTimer()
    StopTimer
    timerID = SetTimer(0, 0, 300000, AddressOf TimerCallback) ' 5 minutes
End Sub

Public Sub StopTimer()
    If timerID <> 0 Then
        KillTimer 0, timerID
        timerID = 0
    End If
End Sub

#If VBA7 Then
Public Sub TimerCallback(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Public Sub TimerCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    RunUnifiedSync
End Sub

' ---------------------------------------------------------------------------
' Sync orchestration
' ---------------------------------------------------------------------------

Public Sub RunUnifiedSync()
    If isSyncRunning Then
        LogMsg "RunUnifiedSync: already running, skipped"
        Exit Sub
    End If
    If lastSync <> 0 And Now - lastSync < TimeValue("00:01:00") Then
        LogMsg "RunUnifiedSync: too soon since last sync, skipped"
        Exit Sub
    End If

    LogMsg "RunUnifiedSync: starting"
    isSyncRunning = True
    lastSync = Now

    On Error GoTo Cleanup
    SyncAllInboxesToUnified
    LogMsg "RunUnifiedSync: finished"

Cleanup:
    If Err.Number <> 0 Then LogMsg "RunUnifiedSync ERROR: " \& Err.Description
    isSyncRunning = False
End Sub

Public Sub SyncAllInboxesToUnified()
    Dim ns As Outlook.NameSpace
    Dim store As Outlook.Store
    Dim inbox As Outlook.Folder

    Set ns = Application.Session

    For Each store In ns.Stores
        Set inbox = Nothing
        On Error Resume Next
        Set inbox = store.GetDefaultFolder(olFolderInbox)
        On Error GoTo 0

        If Not inbox Is Nothing Then
            SyncInbox inbox
        End If
    Next store
End Sub

Public Sub SyncInbox(inbox As Outlook.Folder)
    LogMsg "SyncInbox: " \& inbox.Parent.Name \& " (" \& inbox.Items.Count \& " items)"

    Dim unified As Outlook.Folder
    Set unified = GetUnifiedFolder()
    If unified Is Nothing Then
        LogMsg "SyncInbox: unified folder not found, aborting"
        Exit Sub
    End If

    Dim ns As Outlook.NameSpace
    Set ns = Application.Session

    ' Snapshot unified folder keys (capped at 200)
    Dim copied As Object
    Set copied = CreateObject("Scripting.Dictionary")

    Dim uItems As Outlook.Items
    Set uItems = unified.Items
    uItems.Sort "\[ReceivedTime]", True

    Dim uLimit As Long
    uLimit = uItems.Count
    If uLimit > 200 Then uLimit = 200

    Dim j As Long
    For j = 1 To uLimit
        Dim uItm As Object
        Set uItm = uItems(j)
        If TypeOf uItm Is Outlook.MailItem Then
            Dim uMail As Outlook.MailItem
            Set uMail = uItm
            copied(uMail.Subject \& "|" \& Format(uMail.ReceivedTime, "yyyymmddhhnnss")) = True
        End If
    Next j

    ' Snapshot inbox EntryIDs — never mutate a live collection while iterating
    Dim inboxItems As Outlook.Items
    Set inboxItems = inbox.Items
    inboxItems.Sort "\[ReceivedTime]", True

    Dim entryIDs(99) As String
    Dim candidateCount As Long

    Dim i As Long
    For i = 1 To inboxItems.Count
        Dim itm As Object
        Set itm = inboxItems(i)
        If TypeOf itm Is Outlook.MailItem Then
            entryIDs(candidateCount) = itm.EntryID
            candidateCount = candidateCount + 1
            If candidateCount >= 100 Then Exit For
        End If
    Next i

    ' Process candidates outside the live loop
    Dim k As Long
    For k = 0 To candidateCount - 1
        DoEvents
        Dim mail As Outlook.MailItem
        On Error Resume Next
        Set mail = ns.GetItemFromID(entryIDs(k))
        On Error GoTo 0

        If Not mail Is Nothing Then
            Dim mailKey As String
            mailKey = mail.Subject \& "|" \& Format(mail.ReceivedTime, "yyyymmddhhnnss")
            If copied.Exists(mailKey) Then
                LogMsg "  Hit already-synced item at k=" \& k \& ", stopping"
                Exit For
            End If
            SyncNewItem mail, unified
        End If
    Next k

    LogMsg "SyncInbox done: processed " \& k \& " candidate(s)"
End Sub

Public Sub SyncNewItem(mail As Outlook.MailItem, unified As Outlook.Folder)
    Dim copyItem As Outlook.MailItem
    Set copyItem = mail.Copy
    copyItem.Save
    copyItem.Move unified
    Set copyItem = Nothing
End Sub

' ---------------------------------------------------------------------------
' Folder management
' ---------------------------------------------------------------------------

Public Function GetUnifiedFolder() As Outlook.Folder
    Static cachedFolder As Outlook.Folder
    Dim ns As Outlook.NameSpace
    Set ns = Application.Session

    If Not cachedFolder Is Nothing Then
        Set GetUnifiedFolder = cachedFolder
        Exit Function
    End If

    Dim folderID As String, storeID As String
    folderID = GetSetting("UnifiedInbox", "Config", "FolderID", "")
    storeID = GetSetting("UnifiedInbox", "Config", "StoreID", "")

    If folderID <> "" And storeID <> "" Then
        On Error Resume Next
        Set cachedFolder = ns.GetFolderFromID(folderID, storeID)
        On Error GoTo 0
        If Not cachedFolder Is Nothing Then
            LogMsg "GetUnifiedFolder: loaded from registry -> " \& cachedFolder.Name
            Set GetUnifiedFolder = cachedFolder
            Exit Function
        End If
        LogMsg "GetUnifiedFolder: registry entry found but folder lookup failed, prompting user"
    Else
        LogMsg "GetUnifiedFolder: no registry entry, prompting user"
    End If

    MsgBox "Please select the folder to use as your Unified Inbox.", vbInformation
    Set cachedFolder = ns.PickFolder

    If cachedFolder Is Nothing Then
        LogMsg "GetUnifiedFolder: user cancelled folder selection"
        MsgBox "No folder selected. Unified Inbox sync disabled.", vbExclamation
        Exit Function
    End If

    SaveSetting "UnifiedInbox", "Config", "FolderID", cachedFolder.EntryID
    SaveSetting "UnifiedInbox", "Config", "StoreID", cachedFolder.StoreID
    LogMsg "GetUnifiedFolder: saved new folder -> " \& cachedFolder.Name

    Set GetUnifiedFolder = cachedFolder
End Function
```

## Resetting the Target Folder

To pick a different unified folder, run this once in the VBA Immediate Window (**Ctrl+G**):

```vb
DeleteSetting "UnifiedInbox", "Config"
```

Then restart Outlook — it will prompt you to pick again.

## Troubleshooting

### Log file

Every run writes to `%USERPROFILE%\\UnifiedInbox.log` (e.g. `C:\\Users\\YourName\\UnifiedInbox.log`). Open it in Notepad after a problematic startup.

A healthy startup looks like this:

```
2024-03-15 08:01:02 | === Outlook started ===
2024-03-15 08:01:02 | Session stores count: 3
2024-03-15 08:01:02 | InitWatchers: 3 store(s) found
2024-03-15 08:01:02 |   Watching inbox: user@gmail.com
2024-03-15 08:01:02 |   Watching inbox: user@outlook.com
2024-03-15 08:01:02 |   Watching inbox: Personal Folders
2024-03-15 08:01:02 | InitWatchers complete, watchers: 3
2024-03-15 08:01:03 | GetUnifiedFolder: loaded from registry -> Unified Inbox
2024-03-15 08:01:03 | RunUnifiedSync: starting
2024-03-15 08:01:04 | SyncInbox: user@gmail.com (47 items)
2024-03-15 08:01:04 |   Hit already-synced item at k=2, stopping
2024-03-15 08:01:04 | SyncInbox done: processed 2 candidate(s)
...
2024-03-15 08:01:05 | RunUnifiedSync: finished
2024-03-15 08:01:05 | Startup complete
```

### Common problems

|Symptom in log|Likely cause|Fix|
|-|-|-|
|`stores count: 0` or `stores count: 1`|Outlook not fully loaded when macro fired|Outlook takes a few seconds to connect accounts on startup; the 5-minute timer sweep will catch up automatically|
|`watchers: 0`|No inbox folders found across any store|Check that accounts are fully configured and connected|
|`GetUnifiedFolder: registry entry found but folder lookup failed`|Unified folder was deleted or moved|Run `DeleteSetting "UnifiedInbox", "Config"` in the Immediate Window and restart|
|`RunUnifiedSync: already running, skipped`|A previous sync is still in progress|Normal if the inbox is large; subsequent timer runs will catch up|
|`RunUnifiedSync ERROR: ...`|Exception during sync|Share the full error description for diagnosis|
|No log file created at all|Macros not enabled|File → Options → Trust Center → Macro Settings → Enable all macros|

### Sending logs to support

1. Open `%USERPROFILE%\\UnifiedInbox.log` in Notepad
2. Copy the lines from the startup session where the problem occurred (starting from `=== Outlook started ===`)
3. Share those lines — do not share the entire file if it contains subject lines you want to keep private

## Notes

* Emails are **copied** (not moved) from each inbox into the unified folder
* The unified folder can be in any account or a local PST
* Works with any number of accounts (Exchange, IMAP, POP3)
* The periodic sweep stops scanning an inbox as soon as it hits an already-synced email — if you manually delete emails from the unified folder they will not be re-synced on the next sweep (only future new arrivals are caught via `ItemAdd`)
* Requires macros to be enabled: File → Options → Trust Center → Macro Settings → Enable all macros

