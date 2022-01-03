# How this service were created
This service Install TopShelf in NuGetPackage, and this is using Timer that mean every 1 second would be write timestamp date and time to certain file.

# Enabled Event Log on windows Machine
- Log on to the computer as an administrator.
- Click Start, click Run, type Regedit in the Open box, and then click OK. - The Registry Editor window appears.
- Locate the following registry subkey
```
Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\EventLog
```
- Right-click Eventlog, and then click Permissions. The Permissions for Eventlog dialog box appears.

![Permission Event Log](./Images/EventLogPermission1.jpg)

- Click Add, add the user account or group that you want and set the following permissions: `Full Control`.

![Permission Event Log](./Images/EventLogPermission2.jpg)

- Locate the following registry subkey
```
Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\EventLog\Security
```

![Permission Event Log](./Images/EventLogPermission3.jpg)

- Click Add, add the user account or group that you want and set the following permissions: `Full Control`.

![Permission Event Log](./Images/EventLogPermission4.jpg)

# How deploy this service
- Copy all files in /bin/debug/ to somewhere else folder
- Open CMD as Administrator write for example 
```
$ cd C:\Temp\Released Application\ExCoreService
$ dir
$ ExCoreService.exe install start
$ ExCoreService.exe uninstall
```