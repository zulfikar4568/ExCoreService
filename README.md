# How this service were created
This service Install TopShelf in  NuGetPackage, and this is using Timer that mean every 1 second would be write timestamp date and time to certain file.

# How deploy this service
1. Copy all files in /bin/debug/ to somewhere else folder
2. Open CMD as Administrator write for example 
```
$ cd C:\Temp\Released Application\ExCoreService
$ dir
$ ExCoreService.exe install start
$ ExCoreService.exe uninstall
```