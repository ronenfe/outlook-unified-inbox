# Outlook Unified Inbox (VBA)

A VBA macro for Outlook Classic that syncs emails from all account inboxes into a single unified folder in real time.

## How It Works

- On startup, watches every account's inbox using `WithEvents` — new mail is copied to the unified folder instantly
- A periodic timer sweeps all inboxes every 2 minutes to catch anything missed
- On first run, prompts you to pick the target folder; saves the selection to the Windows registry
- Checks the 100 most recent emails per inbox on each sweep

## Files

| File | Type | Purpose |
|------|------|---------|
| `ThisOutlookSession` | Built-in | Startup/quit hooks |
| `modUnifiedSync` | Module | Timer, sync logic, folder management |
| `clsInboxEvents` | Class Module | Per-inbox `WithEvents` watcher |

## Installation

1. Open Outlook, press **Alt+F11** to open the VBA Editor
2. In the Project Explorer (left panel), expand **Project (VbaProject.OTM)**
3. **Add the module:** Insert → Module → rename it to `modUnifiedSync` → paste contents of `modUnifiedSync.bas`
4. **Add the class:** Insert → Class Module → press F4, set `Name` to `clsInboxEvents` → paste contents of `clsInboxEvents.cls`
5. **Edit ThisOutlookSession:** double-click it → replace contents with `ThisOutlookSession.bas`
6. Save (**Ctrl+S**), close and reopen Outlook
7. On first launch you will be prompted to pick your unified folder — select it once and it is remembered

## Code

### `ThisOutlookSession`
```vb
Option Explicit

Private Sub Application_Startup()
    InitWatchers
    StartTimer
    RunUnifiedSync
End Sub

Private Sub Application_Quit()
    StopTimer
End Sub
```

### `clsInboxEvents`
```vb
Option Explicit

Public WithEvents Items As Outlook.Items

Private Sub Items_ItemAdd(ByVal Item As Object)
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

Public Sub InitWatchers()
    Dim ns As Outlook.NameSpace
    Dim store As Outlook.Store
    Dim inbox As Outlook.Folder
    Dim w As clsInboxEvents

    Set ns = Application.Session

    For Each store In ns.Stores
        Set inbox = Nothing
        On Error Resume Next
        Set inbox = store.GetDefaultFolder(olFolderInbox)
        On Error GoTo 0

        If Not inbox Is Nothing Then
            Set w = New clsInboxEvents
            Set w.Items = inbox.Items
            watchers.Add w
        End If
    Next store
End Sub

Public Sub StartTimer()
    StopTimer
    timerID = SetTimer(0, 0, 120000, AddressOf TimerCallback)
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

Public Sub RunUnifiedSync()
    If isSyncRunning Then Exit Sub
    If lastSync <> 0 And Now - lastSync < TimeValue("00:00:30") Then Exit Sub

    isSyncRunning = True
    lastSync = Now

    On Error GoTo Cleanup
    SyncAllInboxesToUnified

Cleanup:
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
    Dim unified As Outlook.Folder
    Set unified = GetUnifiedFolder()
    If unified Is Nothing Then Exit Sub

    Dim copied As Object
    Set copied = CreateObject("Scripting.Dictionary")

    Dim itm As Object
    For Each itm In unified.Items
        If TypeOf itm Is Outlook.MailItem Then
            Dim uMail As Outlook.MailItem
            Set uMail = itm
            copied(uMail.Subject & "|" & Format(uMail.ReceivedTime, "yyyymmddhhnnss")) = True
        End If
    Next itm

    Dim inboxItems As Outlook.Items
    Set inboxItems = inbox.Items
    inboxItems.Sort "[ReceivedTime]", True

    Dim checked As Long
    Dim i As Long
    For i = 1 To inboxItems.Count
        DoEvents

        Dim mailItem As Object
        Set mailItem = inboxItems(i)

        If TypeOf mailItem Is Outlook.MailItem Then
            Dim mail As Outlook.MailItem
            Set mail = mailItem

            checked = checked + 1
            If checked > 100 Then Exit For

            Dim mailKey As String
            mailKey = mail.Subject & "|" & Format(mail.ReceivedTime, "yyyymmddhhnnss")

            If copied.Exists(mailKey) Then Exit For

            SyncNewItem mail, unified
        End If
    Next i
End Sub

Public Sub SyncNewItem(mail As Outlook.MailItem, unified As Outlook.Folder)
    Dim copyItem As Outlook.MailItem
    Set copyItem = mail.Copy
    copyItem.Move unified
End Sub

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
            Set GetUnifiedFolder = cachedFolder
            Exit Function
        End If
    End If

    MsgBox "Please select the folder to use as your Unified Inbox.", vbInformation
    Set cachedFolder = ns.PickFolder

    If cachedFolder Is Nothing Then
        MsgBox "No folder selected. Unified Inbox sync disabled.", vbExclamation
        Exit Function
    End If

    SaveSetting "UnifiedInbox", "Config", "FolderID", cachedFolder.EntryID
    SaveSetting "UnifiedInbox", "Config", "StoreID", cachedFolder.StoreID

    Set GetUnifiedFolder = cachedFolder
End Function
```

## Resetting the Target Folder

To pick a different unified folder, run this once in the VBA Immediate Window (**Ctrl+G**):

```vb
DeleteSetting "UnifiedInbox", "Config"
```

Then restart Outlook — it will prompt you to pick again.

## Notes

- Emails are **copied** (not moved) from each inbox into the unified folder
- The unified folder can be in any account or a local PST
- Works with any number of accounts (Exchange, IMAP, POP3)
- Requires macros to be enabled: File → Options → Trust Center → Macro Settings → Enable all macros
