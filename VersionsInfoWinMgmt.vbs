Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\7-Zip"
strValueName = "Path"

oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

MsgBox "Die Version von 7-Zip ist: " & GetVer(strValue & "\7z.exe")

function GetVer(pfad)
	on error resume next
    set fs = CreateObject("Scripting.FileSystemObject")
	GetVer = CStr(fs.getFileVersion(pfad))

	if not err.Number=0 then
		GetVer = "??"
		err.clear
	end if
end function