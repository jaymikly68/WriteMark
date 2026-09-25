$ErrorActionPreference = 'Stop'
$keys = @(
    @{ p = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Privacy'; n = 'DisconnectedState' },
    @{ p = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Privacy'; n = 'usercontentdisabled' },
    @{ p = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Privacy'; n = 'downloadcontentdisabled' },
    @{ p = 'HKCU:\Software\Policies\Microsoft\Office\16.0\Common\Privacy'; n = 'controllerconnectedservicesenabled' },
    @{ p = 'HKCU:\Software\Policies\Microsoft\office\Common\ClientTelemetry'; n = 'sendtelemetry' }
)
Write-Host '=== Undo Word exit lag fix ===' -ForegroundColor Cyan
foreach ($k in $keys) {
    if (Test-Path -LiteralPath $k.p) {
        Remove-ItemProperty -LiteralPath $k.p -Name $k.n -Force -ErrorAction SilentlyContinue
    }
    Write-Host ("  removed  {0}  {1}" -f $k.p, $k.n)
}
Write-Host ''
Write-Host 'Office is back to its default state. Restart Word.' -ForegroundColor Green
Write-Host ''
Read-Host 'Press Enter to close'
