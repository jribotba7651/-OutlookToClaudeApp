# Build script for OutlookToClaudeApp
Write-Host "=== OutlookToClaudeApp Build Script ===" -ForegroundColor Cyan

# Try to find dotnet
$dotnetPath = $null
$possiblePaths = @(
    "C:\Program Files\dotnet\dotnet.exe",
    "C:\Program Files (x86)\dotnet\dotnet.exe",
    "$env:ProgramFiles\dotnet\dotnet.exe",
    "$env:USERPROFILE\.dotnet\tools\dotnet.exe"
)

foreach ($path in $possiblePaths) {
    if (Test-Path $path) {
        $dotnetPath = $path
        Write-Host "Found dotnet at: $dotnetPath" -ForegroundColor Green
        break
    }
}

# Try to find MSBuild
if (-not $dotnetPath) {
    Write-Host "dotnet not found, searching for MSBuild..." -ForegroundColor Yellow

    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"

    if (Test-Path $vswhere) {
        $msbuildPath = & $vswhere -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe | Select-Object -First 1

        if ($msbuildPath) {
            Write-Host "Found MSBuild at: $msbuildPath" -ForegroundColor Green
            Write-Host "Building with MSBuild..." -ForegroundColor Cyan
            & $msbuildPath OutlookToClaudeApp.csproj /p:Configuration=Release /p:Platform=x64 /p:RuntimeIdentifier=win-x64 /p:SelfContained=true /p:PublishSingleFile=true /t:Publish

            if ($LASTEXITCODE -eq 0) {
                Write-Host "`nBuild succeeded!" -ForegroundColor Green
                Write-Host "Executable location: bin\Release\net8.0-windows\win-x64\publish\" -ForegroundColor Green
            } else {
                Write-Host "`nBuild failed!" -ForegroundColor Red
            }
            exit $LASTEXITCODE
        }
    }
}

if ($dotnetPath) {
    Write-Host "Building with dotnet..." -ForegroundColor Cyan
    & $dotnetPath publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true

    if ($LASTEXITCODE -eq 0) {
        Write-Host "`nBuild succeeded!" -ForegroundColor Green
        Write-Host "Executable location: bin\Release\net8.0-windows\win-x64\publish\" -ForegroundColor Green
    } else {
        Write-Host "`nBuild failed!" -ForegroundColor Red
    }
} else {
    Write-Host "`nERROR: Neither dotnet nor MSBuild found!" -ForegroundColor Red
    Write-Host "`nPlease install .NET 8.0 SDK from:" -ForegroundColor Yellow
    Write-Host "https://dotnet.microsoft.com/download/dotnet/8.0" -ForegroundColor Cyan
    Write-Host "`nOr install Visual Studio 2022 with .NET desktop development workload." -ForegroundColor Yellow
    exit 1
}
