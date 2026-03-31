#$dest = "$env:USERPROFILE\Desktop\Analyse-Files.ps1"
$dest = "$([Environment]::GetFolderPath('Desktop'))\Analyse-Files.ps1"
Invoke-RestMethod "https://raw.githubusercontent.com/Don-Paterson/File-Analysis/main/Analyse-Files.ps1" |
    Out-File -FilePath $dest -Encoding UTF8 -NoNewline
pwsh -ExecutionPolicy Bypass -File $dest
