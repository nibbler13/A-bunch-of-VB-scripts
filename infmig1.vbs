Option Explicit
dim WshShell,reg_key

set WshShell = CreateObject("wscript.shell")

reg_key = WshShell.RegRead("HKEY_CURRENT_USER\SOFTWARE\Mich\Infoscreen\TerminalID")
WshShell.RegWrite"HKEY_CURRENT_USER\SOFTWARE\Mich\Infoscreen\TerminalIDLarge",reg_key,"REG_SZ"
