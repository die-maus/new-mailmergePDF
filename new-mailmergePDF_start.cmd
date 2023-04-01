powershell -command "& { Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser }"
powershell -command "& { . .\new-mailmergePDF.ps1 }"
powershell -command "& { Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope CurrentUser }"

pause