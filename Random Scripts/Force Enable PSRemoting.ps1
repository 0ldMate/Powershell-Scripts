$cn = 'localhost'
([wmiclass]"\\$cn\root\cimv2:win32_process").Create('powershell Enable-PSRemoting -Force')