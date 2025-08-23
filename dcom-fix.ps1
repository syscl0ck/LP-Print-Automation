# PowerShell script to fix COM Desktop path issue
$paths = @(
  "$env:windir\System32\config\systemprofile\Desktop",
  "$env:windir\SysWOW64\config\systemprofile\Desktop"
)

foreach ($path in $paths) {
    if (-Not (Test-Path $path)) {
        New-Item -Path $path -ItemType Directory | Out-Null
        Write-Host "Created: $path"
    } else {
        Write-Host "Already exists: $path"
    }
}