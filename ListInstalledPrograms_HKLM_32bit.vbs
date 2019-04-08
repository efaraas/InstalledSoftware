Option Explicit 

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE 
Dim strComputer, strKey

strComputer = "." 
strKey = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" 

Dim objReg, strSubkey, arrSubkeys 

Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv") 
objReg.EnumKey HKLM, strKey, arrSubkeys 

'Loop registry key.
Dim DisplayName,DisplayVersion, InstallDate, EstimatedSize, UninstallString

WScript.echo  "DisplayName;DisplayVersion;InstallDate;EstimatedSize;UninstallString"
For Each strSubkey In arrSubkeys 
	objReg.GetStringValue HKLM, strKey & strSubkey, "DisplayName" , DisplayName
	If DisplayName <> "" Then 
		objReg.GetStringValue HKLM, strKey & strSubkey, "DisplayVersion" , DisplayVersion
		objReg.GetStringValue HKLM, strKey & strSubkey, "InstallDate", InstallDate
		objReg.GetDWORDValue HKLM, strKey & strSubkey, "EstimatedSize" , EstimatedSize
		
		If  EstimatedSize <> "" Then 
			WScript.echo DisplayName & ";" & DisplayVersion& ";" & InstallDate & ";" & Round(EstimatedSize/1024, 3) 
		Else 
			WScript.echo DisplayName & ";" & DisplayVersion& ";" & InstallDate & ";" 
		End If 
		
	End If 
Next
