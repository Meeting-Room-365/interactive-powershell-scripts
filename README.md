```
::::    ::::  :::::::::: :::::::::: ::::::::::: ::::::::::: ::::    :::  ::::::::   :::::::::   ::::::::   ::::::::  ::::    ::::    ::::::::   ::::::::  :::::::::: 
+:+:+: :+:+:+ :+:        :+:            :+:         :+:     :+:+:   :+: :+:    :+:  :+:    :+: :+:    :+: :+:    :+: +:+:+: :+:+:+  :+:    :+: :+:    :+: :+:    :+: 
+:+ +:+:+ +:+ +:+        +:+            +:+         +:+     :+:+:+  +:+ +:+         +:+    +:+ +:+    +:+ +:+    +:+ +:+ +:+:+ +:+         +:+ +:+        +:+        
+#+  +:+  +#+ +#++:++#   +#++:++#       +#+         +#+     +#+ +:+ +#+ :#:         +#++:++#:  +#+    +:+ +#+    +:+ +#+  +:+  +#+      +#++:  +#++:++#+  +#++:++#+  
+#+       +#+ +#+        +#+            +#+         +#+     +#+  +#+#+# +#+   +#+#  +#+    +#+ +#+    +#+ +#+    +#+ +#+       +#+         +#+ +#+    +#+        +#+ 
#+#       #+# #+#        #+#            #+#         #+#     #+#   #+#+# #+#    #+#  #+#    #+# #+#    #+# #+#    #+# #+#       #+#  #+#    #+# #+#    #+# #+#    #+# 
###       ### ########## ##########     ###     ########### ###    ####  ########   ###    ###  ########   ########  ###       ###   ########   ########   ########  
```
# Interactive PowerShell Scripts
Interactive Microsoft Office 365 PowerShell Management scripts.

Copyright © 2024 Meeting Room 365 llc. All Rights Reserved.

Visit www.meetingroom365.com for more details.

-----
## Room List Manager
![Room List Manager](./RoomListManager.jpg)
Used to create meeting rooms and workspaces, create room lists (for room finder), and organize rooms and workspaces inside room lists.

Can also run a script to fix subjects for newly-created Meeting Room resource mailboxes.

[View PowerShell Script](./RoomListManager.ps1)

-----

# Instructions for Running PowerShell Scripts

## Windows
Install PowerShell 7.4 (or later) for Windows 10/11

https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4

#### Prepare Powershell
1. Open the Start Menu.
2. Search for PowerShell, right-click the top result, and select the Run as administrator option.
3. Type the following command to allow scripts to run and press Enter: `Set-ExecutionPolicy RemoteSigned`
4. Type "A" and press Enter (if applicable).

#### Run the script
Right click on RoomListManager.ps1 and click "Run as Administrator"

## MacOS

https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-macos?view=powershell-7.4

```shell
brew install powershell/tap/powershell

# Start PowerShell
pwsh

./RoomListManager.ps1
```

## Linux

https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-linux?view=powershell-7.4

```shell
# Install PowerShell
sudo apt install powershell

# Start PowerShell
pwsh

./RoomListManager.ps1
```