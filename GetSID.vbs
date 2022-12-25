'Description: Zeigt alle Benutzerkonten und SID
'Created on : Apr 5, 2010
'Prerequisite: Windows

Set WshShell = CreateObject("Wscript.Shell")
Set fso = Wscript.CreateObject("Scripting.FilesystemObject")
fName = WshShell.SpecialFolders("Desktop") & "\GetSID.txt"
Set b = fso.CreateTextFile(fName, true)
b.writeblanklines 1
b.writeline string(61,"*")
b.writeline "Benutzerkonten mit SIDs und Profile-Pfade."
b.WriteLine "GetSID.vbs - 2010 by Don Metteo"
b.WriteLine "https://think.unblog.ch"
b.writeline string(61,"*")
b.writeblanklines 1

strProfileBranch = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\"
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colAccounts = objWMIService.ExecQuery _
 ("Select * From Win32_UserAccount")

For Each objAccount in colAccounts
 If objAccount.Name = "HelpAssistant" or objAccount.Name = "SUPPORT_388945a0" then
 else
 b.writeline "Username : " & objAccount.Name
 b.writeline "SID : " & objAccount.SID
 b.writeline "Profile Path : " & GetHomePath(objAccount.SID)
 b.writeblanklines 1
 end if
Next

Function GetHomePath(strSID)
 On Error Resume Next
 GetHomePath = WshShell.ExpandEnvironmentStrings(Trim(WshShell.RegRead (strProfileBranch & strSID & "\ProfileImagePath")))
 On Error Goto 0
End Function

b.writeline string(61,"*")
b.close
WshShell.Run "notepad.exe " & fName

Set fso = Nothing
set Wshshell = Nothing
