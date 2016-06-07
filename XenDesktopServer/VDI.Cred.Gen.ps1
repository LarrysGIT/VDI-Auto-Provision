
#$VICred = Get-Credential -Message "For VI server."
#$ControllerCred = Get-Credential

Write-Host "Controller username: " -NoNewline -ForegroundColor Yellow
$u = Read-Host

Write-Host "Controller password: " -NoNewline -ForegroundColor Yellow
$p = Read-Host -AsSecureString
if($u -and $p)
{
    $Hash1 = ConvertTo-SecureString -String $u -AsPlainText -Force | ConvertFrom-SecureString
    $Hash2 = $p | ConvertFrom-SecureString

    New-Item -Path 'HKCU:\Software\company\Sharepoint' -ErrorAction:SilentlyContinue
    New-ItemProperty -Path 'HKCU:\Software\company\Sharepoint' -Name 'Hash1' -Value $Hash1 -Force
    New-ItemProperty -Path 'HKCU:\Software\company\Sharepoint' -Name 'Hash2' -Value $Hash2 -Force
}

Write-Host "VDI Service Account username: " -NoNewline -ForegroundColor Yellow
$u = Read-Host

Write-Host "VDI Service Account password: " -NoNewline -ForegroundColor Yellow
$p = Read-Host -AsSecureString
if($u -and $p)
{
    $Hash1 = ConvertTo-SecureString -String $u -AsPlainText -Force | ConvertFrom-SecureString
    $Hash2 = $p | ConvertFrom-SecureString

    New-Item -Path 'HKCU:\Software\company\Windows7' -ErrorAction:SilentlyContinue
    New-ItemProperty -Path 'HKCU:\Software\company\Windows7' -Name 'Hash1' -Value $Hash1 -Force
    New-ItemProperty -Path 'HKCU:\Software\company\Windows7' -Name 'Hash2' -Value $Hash2 -Force
}

Write-Host 'VDI Local Account username: ' -NoNewline -ForegroundColor Yellow
$u = Read-Host

Write-Host 'VDI Local Account password: ' -NoNewline -ForegroundColor Yellow
$p = Read-Host -AsSecureString
if($u -and $p)
{
    $Hash1 = ConvertTo-SecureString -String $u -AsPlainText -Force | ConvertFrom-SecureString
    $Hash2 = $p | ConvertFrom-SecureString

    New-Item -Path 'HKCU:\Software\company\Windows7' -ErrorAction:SilentlyContinue
    New-ItemProperty -Path 'HKCU:\Software\company\Windows7' -Name 'Hash5' -Value $Hash1 -Force
    New-ItemProperty -Path 'HKCU:\Software\company\Windows7' -Name 'Hash6' -Value $Hash2 -Force
}
