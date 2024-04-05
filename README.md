# OutlookAddin_ItemSend_SmartAlerts

### Description

SmartAlerts manifest doesn't work in Mac New Outlook.

### Hardware & Software

MacBook Air M1 2020
OS: macOs 14.2.1

Outlook: 
Microsoft® Outlook for Mac
Version 16.83.1 (24031813)


### Steps to reproduce

1. Open Mac New Outlook app.
2. Go to "Get Add-ins" -> "My Add-ins"
3. Install "manifest_smartAlerts.xml" from file.
4. Close window.
5. Select any message in inbox.
6. Try to find addin icon on toolbar.
7. Start composing new message.
8. Try to find addin icon on toolbar in composing mode.
9. Send message to any address.

Expected result:
1. On step 6, addin icon should be present on toolbar.
2. On step 8, addin icon should be present on toolbar.
3. On step 9, addin should block message.
