Set objShell = CreateObject("Wscript.Shell")
objShell.Run "pwsh.exe -NoProfile -ExecutionPolicy Bypass -File ""JsonLookup.ps1""", 0, False