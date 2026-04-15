# Outlook Classic — Unified Inbox (VBA Macro)

A lightweight VBA macro for Outlook Classic that automatically copies incoming emails from all your accounts into a single unified inbox folder — in real time.

---

## The Problem

Outlook Classic doesn't have a true unified inbox that works as an actual folder. The built-in "All Inboxes" favorites view is read-only and can't receive copied items. If you have multiple email accounts (Gmail, Exchange, etc.), you're stuck switching between inboxes manually.

## The Solution

This macro hooks into every account's inbox at startup and listens for new mail. When an email arrives in any inbox, it's automatically copied into a folder of your choice — giving you one place to read everything.

---

## Features

- **Works with all account types** — Gmail (IMAP), Exchange, Outlook.com, POP3
- **Language-proof** — works with Hebrew, Arabic, or any RTL/non-Latin Outlook installation
- **Folder selection is persistent** — pick your unified folder once, it's remembered
- **Automatically hooks all inboxes** on Outlook startup
- **Batch-safe** — scans for all recently arrived emails on each trigger, so multiple emails arriving simultaneously are all copied
- **Deduplication** based on sender, subject, and received time — no duplicates in the unified folder
- **Handles IMAP download delays** with automatic retry

---

## Setup

### Step 1 — Open the VBA Editor

Press `Alt + F11` in Outlook.

---

### Step 2 — Create the Class Module

1. In the left panel, right-click your project → **Insert → Class Module**
2. Press `F4` → rename it to `clsInboxEvents`
3. Paste the following code:

```vb
Option Explicit
Public WithEvents Items As Outlook.Items

Private Sub Items_ItemAdd(ByVal Item As Object)
    Debug.Print Now & " — New mail detected in " & Items.Parent.FolderPath
    SyncRecentItems Items.Parent
End Sub
```

---

### Step 3 — Create the Standard Module

1. Right-click your project → **Insert → Module**
2. Press `F4` → rename it to `modUnified`
3. Paste the following code:

```vb
Option Explicit
Public InboxEventHandlers As Collection
Public Sub ManualStartup()
    Dim ns As Outlook.NameSpace
    Dim store As Outlook.Store
    Dim inbox As Outlook.Folder
    Dim handler As clsInboxEvents

    Set ns = Application.Session
    Set InboxEventHandlers = New Collection

    For Each store In ns.Stores
        On Error Resume Next
        Set inbox = store.GetDefaultFolder(olFolderInbox)
        On Error GoTo 0
        If Not inbox Is Nothing Then
            Set handler = New clsInboxEvents
            Set handler.Items = inbox.Items
            InboxEventHandlers.Add handler
            Debug.Print "Hooked: " & inbox.FolderPath
        End If
    Next store

    Debug.Print Now & " — Total hooks: " & InboxEventHandlers.Count
End Sub

Public Sub SyncRecentItems(inbox As Outlook.Folder)
    Dim unified As Outlook.Folder
    Set unified = GetUnifiedFolder()
    If unified Is Nothing Then Exit Sub

    Dim cutoff As Date
    cutoff = Now - (5 / 1440)  ' last 5 minutes

    Dim i As Integer
    For i = inbox.Items.Count To 1 Step -1
        Dim item As Object
        Set item = inbox.Items(i)

        If Not TypeOf item Is Outlook.MailItem Then GoTo NextItem

        Dim mail As Outlook.MailItem
        Set mail = item

        If mail.ReceivedTime < cutoff Then Exit For

        If Not IsAlreadyCopied(mail, unified) Then
            SyncNewItem mail, unified
        End If

NextItem:
    Next i
End Sub

Private Function IsAlreadyCopied(mail As Outlook.MailItem, unified As Outlook.Folder) As Boolean
    Dim i As Integer
    For i = 1 To unified.Items.Count
        If Not TypeOf unified.Items(i) Is Outlook.MailItem Then GoTo Next_i
        Dim existing As Outlook.MailItem
        Set existing = unified.Items(i)
        If existing.Subject = mail.Subject And _
           existing.SenderEmailAddress = mail.SenderEmailAddress And _
           Abs(existing.ReceivedTime - mail.ReceivedTime) < (1 / 1440) Then
            IsAlreadyCopied = True
            Exit Function
        End If
Next_i:
    Next i
    IsAlreadyCopied = False
End Function

Public Sub SyncNewItem(item As Outlook.MailItem, unified As Outlook.Folder)
    Dim newItem As Outlook.MailItem

    Dim i As Integer
    For i = 1 To 10
        On Error Resume Next
        Set newItem = item.Copy
        On Error GoTo 0
        If Not newItem Is Nothing Then Exit For
        Wait 1000
    Next i

    If newItem Is Nothing Then
        Debug.Print Now & " — Copy failed: " & item.Subject
        Exit Sub
    End If

    On Error Resume Next
    newItem.Move unified
    If Err.Number <> 0 Then
        Debug.Print Now & " — Move failed: " & Err.Description
    Else
        Debug.Print Now & " — Copied: " & item.Subject
    End If
    On Error GoTo 0
End Sub

Public Function GetUnifiedFolder() As Outlook.Folder
    Dim ns As Outlook.NameSpace
    Dim f As Outlook.Folder
    Dim folderPath As String

    Set ns = Application.Session
    folderPath = GetSetting("UnifiedSync", "Folders", "UnifiedPath", "")

    If folderPath = "" Then
        MsgBox "Select the unified folder you want new mail copied into.", vbInformation
        Set f = ns.PickFolder
        If f Is Nothing Then Exit Function
        SaveSetting "UnifiedSync", "Folders", "UnifiedPath", f.FolderPath
        Set GetUnifiedFolder = f
        Exit Function
    End If

    On Error Resume Next
    Dim parts() As String
    Dim start As Long
    parts = Split(folderPath, "\")
    start = 0
    Do While parts(start) = "" And start < UBound(parts)
        start = start + 1
    Loop
    Set f = ns.Folders(parts(start))
    Dim j As Long
    For j = start + 1 To UBound(parts)
        Set f = f.Folders(parts(j))
    Next j
    On Error GoTo 0

    If f Is Nothing Then
        MsgBox "Unified folder not found. Please select it again.", vbExclamation
        Set f = ns.PickFolder
        If f Is Nothing Then Exit Function
        SaveSetting "UnifiedSync", "Folders", "UnifiedPath", f.FolderPath
    End If

    Set GetUnifiedFolder = f
End Function

Public Sub ResetUnifiedFolder()
    On Error Resume Next
    DeleteSetting "UnifiedSync", "Folders", "UnifiedPath"
    On Error GoTo 0
    Debug.Print "Folder reset — run TestUnified to pick a new one"
End Sub

Public Sub TestUnified()
    Dim f As Outlook.Folder
    Set f = GetUnifiedFolder()
    If f Is Nothing Then
        Debug.Print "No folder selected"
    Else
        Debug.Print "Unified folder: " & f.FolderPath
    End If
End Sub

Public Sub Wait(ms As Long)
    Dim t As Single
    t = Timer
    Do While Timer < t + (ms / 1000)
        DoEvents
    Loop
End Sub
```

---

### Step 4 — Edit ThisOutlookSession

1. In the left panel under **Microsoft Outlook Objects**, double-click `ThisOutlookSession`
2. Paste the following:

```vb
Option Explicit
Public InboxEventHandlers As Collection

Private Sub Application_Startup()
    ManualStartup
End Sub
```

---

### Step 5 — Enable Macros

In Outlook: **File → Options → Trust Center → Trust Center Settings → Macro Settings**

Set to `Notifications for all macros` or `Enable all macros`.

---

### Step 6 — First Run

1. Create a folder in Outlook for your unified inbox (e.g. `All Mail`)
2. Restart Outlook — the macro will start automatically
3. On the first email arrival, a dialog will ask you to pick your unified folder
4. Pick your folder — it's saved permanently from that point

> Or run `TestUnified` manually from the VBA editor (`Alt+F11`) to pick the folder right away.

---

## Troubleshooting

**Macro doesn't run on startup**
- Check macro security settings (Step 5)
- Make sure `Application_Startup` is in `ThisOutlookSession`, not a regular module

**Folder picker doesn't appear**
- Run `ResetUnifiedFolder` in the VBA editor, then run `TestUnified`

**Emails not copying**
- Run `ManualStartup` manually to re-establish hooks
- Check the Immediate window (`Ctrl+G`) for debug output

**Multiple emails from the same sender not all appearing**
- This was a known issue with batch IMAP sync — when several emails arrive at once, Outlook may only fire `ItemAdd` once for the batch. The macro now handles this by scanning the inbox for all items received in the last 5 minutes on each trigger, so all emails in a batch are captured. If you were on an older version, make sure you've updated both `clsInboxEvents` and `modUnified` with the latest code above.

**Hebrew / RTL folder names not working**
- The macro uses folder path traversal which is language-safe
- If issues persist, run `ResetUnifiedFolder` and re-pick the folder

---

## Notes

- Emails that arrive while Outlook is closed will be copied when Outlook reopens and syncs
- The original email stays in its source inbox — only a copy is moved to the unified folder
- This macro does not sync read/unread status between the copy and original
- Deduplication is based on sender address, subject, and received time — not `EntryID`, which changes when an item is copied

---

## License

MIT
