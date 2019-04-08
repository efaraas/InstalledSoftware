REM *******   vbscript to list installed programs  *******
%SystemRoot%\system32\cscript.exe //NoLogo ListInstalledPrograms_HKLM_64bit.vbs > InstalledPrograms_HKLM_64bit.csv
%SystemRoot%\system32\cscript.exe //NoLogo ListInstalledPrograms_HKLM_32bit.vbs > InstalledPrograms_HKLM_32bit.csv