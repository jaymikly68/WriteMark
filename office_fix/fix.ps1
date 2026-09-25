$ErrorActionPreference = 'Stop'
param([switch]$NoPause)
$keys = @(
    @{ p = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Privacy'; n = 'DisconnectedState';                   v = 2 },
    @{ p = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Privacy'; n = 'usercontentdisabled';                 v = 2 },
    @{ p = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Privacy'; n = 'downloadcontentdisabled';             v = 2 },
    @{ p = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Privacy'; n = 'controllerconnectedservicesenabled';  v = 2 },
    @{ p = 'HKCU:\Software\Policies\Microsoft\office\Common\ClientTelemetry'; n = 'sendtelemetry';                    v = 3 }
)
Write-Host '=== Word exit lag fix (lightweight) ===' -ForegroundColor Cyan
foreach ($k in $keys) {
    if (-not (Test-Path -LiteralPath $k.p)) {
        New-Item -Path $k.p -Force | Out-Null
    }
    New-ItemProperty -LiteralPath $k.p -Name $k.n -Value $k.v -PropertyType DWord -Force | Out-Null
    Write-Host ("  OK  {0}  {1} = {2}" -f $k.p, $k.n, $k.v)
}
Write-Host ''
Write-Host 'Done. Please fully close Word and open it again.' -ForegroundColor Green
Write-Host 'Not good? Run revert.bat to undo everything.' -ForegroundColor Yellow
Write-Host ''
if ($NoPause) { Start-Sleep -Seconds 2 } else { Read-Host 'Press Enter to close' }
