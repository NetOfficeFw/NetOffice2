:: C:\Windows\SysWOW64\regsvr32.exe /s PowerPointAddin.comhost.dll
C:\Windows\System32\regsvr32.exe /s PowerPointAddin.comhost.dll

reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\NetOfficeSamples.PowerPointAddin" /f /v LoadBehavior /t REG_DWORD /d 3
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\NetOfficeSamples.PowerPointAddin" /f /v FriendlyName /t REG_SZ /d "PowerPoint Addin (.NET Core 3.1 / NetOffice 2)"
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\NetOfficeSamples.PowerPointAddin" /f /v Description /t REG_SZ /d "Sample addin running in .NET Core 3.1 using NetOffice 2 alpha"


reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\NetOfficeSamples.PowerPointAddin" /f /v LoadBehavior /t REG_DWORD /d 3
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\NetOfficeSamples.PowerPointAddin" /f /v FriendlyName /t REG_SZ /d "PowerPoint Addin (.NET Core 3.1 / NetOffice 2)"
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\NetOfficeSamples.PowerPointAddin" /f /v Description /t REG_SZ /d "Sample addin running in .NET Core 3.1 using NetOffice 2 alpha"
