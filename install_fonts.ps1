$Source = "T:\PathTo\Top\Folder" 

$fonts = Get-ChildItem -Path $Source -Recurse -File |
         Where-Object { $_.Extension -match '\.(ttf|otf)$' }

if (-not $fonts) {
    Write-Host "No .ttf or .otf fonts found under $Source" -ForegroundColor Red
    return
}

Write-Host "Found $($fonts.Count) font(s). Beginning install..." -ForegroundColor Cyan

$shell = New-Object -ComObject Shell.Application
$fontsFolder = $shell.Namespace(0x14)

foreach ($font in $fonts) {
    Write-Host "Installing: $($font.FullName)" -ForegroundColor Yellow
    try {
        $fontsFolder.CopyHere($font.FullName, 0x10)
        Write-Host "  -> Installed" -ForegroundColor Green
    }
    catch {
        Write-Host "  -> FAILED: $_" -ForegroundColor Red
    }
}

Write-Host "Done."
