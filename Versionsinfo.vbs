set wshshell = CreateObject("WScript.Shell")

zipPath = wshshell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\7-Zip\Path")

MsgBox "Die Version von 7-Zip ist: " & GetVer(zipPath & "\7z.exe")

function GetVer(pfad)
	on error resume next
    set fs = CreateObject("Scripting.FileSystemObject")
	GetVer = CStr(fs.getFileVersion(pfad))

	if not err.Number=0 then
		GetVer = "??"
		err.clear
	end if
end function