:: C:\Windows\SysWOW64\regsvr32.exe /s WordAddin.comhost.dll
C:\Windows\System32\regsvr32.exe /s WordAddin.comhost.dll

reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\NetOfficeSamples.WordAddin" /f /v LoadBehavior /t REG_DWORD /d 3
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\NetOfficeSamples.WordAddin" /f /v FriendlyName /t REG_SZ /d "Word Addin (.NET 6 / NetOffice 2)"
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\NetOfficeSamples.WordAddin" /f /v Description /t REG_SZ /d "Sample addin running in .NET 6 using NetOffice 2 alpha"
