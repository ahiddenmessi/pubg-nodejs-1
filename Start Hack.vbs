ip = Inputbox("Enter Local IP Address","Input Required")

Set pust = CreateObject("Shell.Application") 
pust.ShellExecute "node", "index.js " & ip, "", "runas", 1 