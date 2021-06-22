Const HKEY_CURRENT_USER = &H80000001
strComputer = "."
strKeyPath = "Control Panel\Desktop\WindowMetrics"
strEntryName ="AppliedDPI"
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\default:StdRegProv")
objReg.GetDwordValue HKEY_CURRENT_USER, strKeyPath, strEntryName, strValue
Wscript.Echo strValue, "DPI"
